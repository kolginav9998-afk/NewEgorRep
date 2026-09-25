' WmsIssue — выдача по ЕИ (MASTER SPEC v0.3 §4, §5, §9, §10, §12–§16, §18; D-011…D-014, D-017, D-033).
' ЕИ → данные товара → количество → получатель → «Провести»; «Исправить» и «Удалить» — составные операции (сторно
' старого движения и новое движение одной строкой журнала). Всё, что меняет учёт, идёт одной операцией через
' WmsCore.ApplyOperation; обработчик изменения листа только показывает предпросмотр и ничего не проводит.
Option Explicit

' recipients directory (sheet «Получатели»: A abbreviation, B full name), read into memory on first use and after
' every change of that sheet — the sheet is never scanned per keystroke (spec §31). Global: module-level Private
' variables of LibreOffice Basic are reset at every new macro call, the cache must live for the session.
Global gRcLoaded As Boolean
Global gRcN As Long
Global gRcKey() As String
Global gRcName() As String
Global gRcNameKey() As String
Global gRcDupKey() As Boolean
Global gRcDup As String

' UnpostedCount stopped at UNPOSTED_MAX_BLOCKS: its result is a lower bound
Global gUnpostedMore As Boolean

' change handler guard (re-entry: the handler writes into the sheet it watches)
Private mInHandler As Boolean

' validated candidate of the row being posted or corrected (filled by CheckInputs)
Private mCn As Long
Private mCcanon As String
Private mCdata As Variant
Private mCstock As Double
Private mCqty As Double
Private mCwho As String
Private mCwhoNote As String
Private mCdate As Double

' ================================================================ sheets

Function IssuesSheet() As Object
    IssuesSheet = gDoc.Sheets.getByName(SH_ISSUES)
End Function

Function StockSheet() As Object
    StockSheet = gDoc.Sheets.getByName(SH_STOCK)
End Function

' "" when the sheets of the issue scenario exist (a renamed or missing sheet blocks the start)
Function SheetsProblem() As String
    Dim names As Variant, i As Integer
    names = Array(SH_ISSUES, SH_STOCK, SH_RCPT)
    For i = 0 To UBound(names)
        If Not gDoc.Sheets.hasByName(names(i)) Then
            SheetsProblem = "нет листа «" & names(i) & "»"
            Exit Function
        End If
    Next i
    If IssuesSheet().getCellByPosition(IC_EI, 0).getString() <> "Внутренний код" Or IssuesSheet().getCellByPosition(IC_CTL, 0).getString() <> "Контроль" Then
        SheetsProblem = "лист «" & SH_ISSUES & "»: заголовок не совпадает с A:R спецификации"
        Exit Function
    End If
    If StockSheet().getCellByPosition(SC_EI, 0).getString() <> "ЕИ" Or StockSheet().getCellByPosition(SC_QTY, 0).getString() <> "Остаток" Then
        SheetsProblem = "лист «" & SH_STOCK & "»: заголовок не совпадает с реестром ЕИ"
        Exit Function
    End If
    SheetsProblem = ""
End Function

' ================================================================ EI: normalization and the registry (spec §4, §5)

' Accepts "123", "00000123", "ЕИ-123", "ЕИ-00000123", "EI-123", any case, Cyrillic or Latin letters mixed, a dash or a
' space after the prefix, or a number. Canonical form: ЕИ-00000123.
Function NormalizeEI(v As Variant, ByRef n As Long, ByRef canon As String, ByRef msg As String) As Boolean
    Dim t As String, c1 As String, c2 As String, x As Double
    n = 0
    canon = ""
    msg = ""
    Select Case VarType(v)
    Case 2, 3, 4, 5, 6
        x = CDbl(v)
        If x < 1 Or x > 99999999 Or x <> Int(x) Then
            msg = "«" & NumStr(x) & "» не похоже на номер ЕИ"
            Exit Function
        End If
        n = CLng(x)
    Case Else
        t = UCase(Trim(CStr(v)))
        If t = "" Then
            msg = "не указан ЕИ"
            Exit Function
        End If
        c1 = Left(t, 1)
        c2 = Mid(t, 2, 1)
        If (c1 = "Е" Or c1 = "E") And (c2 = "И" Or c2 = "I") Then
            t = Mid(t, 3)
            If Left(t, 1) = "-" Or Left(t, 1) = " " Or Left(t, 1) = Chr(8211) Then t = Mid(t, 2)
        End If
        If t = "" Or Len(t) > EI_DIGITS Or Not IsDigits(t) Then
            msg = "«" & Trim(CStr(v)) & "» не похоже на номер ЕИ (пример: ЕИ-00000123 или 123)"
            Exit Function
        End If
        n = CLng(t)
        If n < 1 Then
            msg = "номер ЕИ должен быть больше 0"
            Exit Function
        End If
    End Select
    canon = EiCanon(n)
    NormalizeEI = True
End Function

Function EiCanon(n As Long) As String
    EiCanon = EI_PREFIX & Right(String(EI_DIGITS, "0") & n, EI_DIGITS)
End Function

' dense addressing: registry row index = EI number, the key in that row is checked (spec §14: the row is only a hint)
Function StockRowProblem(n As Long, canon As String) As String
    Dim s As String
    StockRowProblem = ""
    If n > MAX_SHEET_ROW Then
        StockRowProblem = canon & " не найден в «" & SH_STOCK & "»"
        Exit Function
    End If
    s = StockSheet().getCellByPosition(SC_EI, n).getString()
    If s = canon Then Exit Function
    If s = "" Then
        StockRowProblem = canon & " не найден в «" & SH_STOCK & "»"
    Else
        StockRowProblem = "реестр «" & SH_STOCK & "» повреждён: в строке " & (n + 1) & " записан «" & s & "», ожидался " & canon
    End If
End Function

' A..I of registry row n
Function StockData(n As Long) As Variant
    StockData = StockSheet().getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
End Function

' the balance of registry row n; False when the cell does not hold a number ≥ 0
Function StockQty(n As Long, ByRef q As Double) As Boolean
    Dim c As Object
    c = StockSheet().getCellByPosition(SC_QTY, n)
    q = 0
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then
        StockQty = True
    ElseIf c.getType() = com.sun.star.table.CellContentType.VALUE Then
        q = c.getValue()
        StockQty = (q >= 0)
    End If
End Function

Function Round3(x As Double) As Double
    If x >= 0 Then
        Round3 = Int(x * 1000 + 0.5) / 1000
    Else
        Round3 = -Int(-x * 1000 + 0.5) / 1000
    End If
End Function

' quantity as the user writes it: comma, at most 3 decimals, no trailing zeros
Function QtyText(x As Double) As String
    Dim m As Double, ip As Double, fr As Long, s As String
    m = Int(x * 1000 + 0.5)
    ip = Int(m / 1000)
    fr = CLng(m - ip * 1000)
    s = Format(ip, "0")
    If fr > 0 Then
        s = s & "," & Right("000" & fr, 3)
        Do While Right(s, 1) = "0"
            s = Left(s, Len(s) - 1)
        Loop
    End If
    QtyText = s
End Function

' ================================================================ recipients (spec §10, D-012)

Sub RecipientsInvalidate()
    gRcLoaded = False
End Sub

' «Получатели» sheet event: the directory is read again at the next use
Sub OnRecipientsChange(Optional oTarget As Variant)
    gRcLoaded = False
End Sub

Private Sub RecipientsLoad()
    Dim sh As Object, cur As Object, last As Long, d As Variant, i As Long, j As Long, n As Long, k As String
    gRcN = 0
    gRcDup = ""
    ReDim gRcKey(0)
    ReDim gRcName(0)
    ReDim gRcNameKey(0)
    ReDim gRcDupKey(0)
    gRcLoaded = True
    If Not gDoc.Sheets.hasByName(SH_RCPT) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_RCPT)
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    If last < 1 Then Exit Sub
    d = sh.getCellRangeByPosition(0, 1, 1, last).getDataArray()
    ReDim gRcKey(last)
    ReDim gRcName(last)
    ReDim gRcNameKey(last)
    ReDim gRcDupKey(last)
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(1))) <> "" Then
            gRcKey(n) = LCase(Trim(CStr(d(i)(0))))
            gRcName(n) = Trim(CStr(d(i)(1)))
            gRcNameKey(n) = LCase(gRcName(n))
            n = n + 1
        End If
    Next i
    gRcN = n
    ' an abbreviation must belong to one recipient (D-012); a repeated one is ambiguous and never expanded
    For i = 0 To n - 1
        k = gRcKey(i)
        If k <> "" And Not gRcDupKey(i) Then
            For j = i + 1 To n - 1
                If gRcKey(j) = k Then
                    gRcDupKey(i) = True
                    gRcDupKey(j) = True
                End If
            Next j
            If gRcDupKey(i) Then gRcDup = gRcDup & IIf(gRcDup <> "", ", ", "") & "«" & k & "»"
        End If
    Next i
End Sub

' repeated abbreviations of the directory ("" when none) — for the self-check
Function RecipientsDuplicates() As String
    If Not gRcLoaded Then RecipientsLoad()
    RecipientsDuplicates = gRcDup
End Function

Function RecipientsCount() As Long
    If Not gRcLoaded Then RecipientsLoad()
    RecipientsCount = gRcN
End Function

' 0 known full name, 1 abbreviation expanded, 2 not in the directory (warning), 3 ambiguous abbreviation, 4 empty
Function ResolveRecipient(sIn As String, ByRef sFull As String, ByRef note As String) As Integer
    Dim t As String, k As String, i As Long
    If Not gRcLoaded Then RecipientsLoad()
    t = Trim(sIn)
    sFull = t
    note = ""
    If t = "" Then
        note = "не указан получатель («Кому выдано»)"
        ResolveRecipient = 4
        Exit Function
    End If
    k = LCase(t)
    For i = 0 To gRcN - 1
        If gRcKey(i) = k Then
            If gRcDupKey(i) Then
                note = "сокращение «" & t & "» в справочнике «" & SH_RCPT & "» встречается несколько раз — введите полное имя"
                ResolveRecipient = 3
            Else
                sFull = gRcName(i)
                ResolveRecipient = 1
            End If
            Exit Function
        End If
    Next i
    For i = 0 To gRcN - 1
        If gRcNameKey(i) = k Then
            sFull = gRcName(i)
            ResolveRecipient = 0
            Exit Function
        End If
    Next i
    note = "получатель «" & t & "» не найден в справочнике «" & SH_RCPT & "» — будет записан как введён"
    ResolveRecipient = 2
End Function

' ================================================================ date (H) — required, no automatic business rule

' a Calc date (01.01.2000–31.12.2099), or a text dd.mm.yyyy; False with msg otherwise
Function ParseIssueDate(oCell As Object, ByRef d As Double, ByRef msg As String) As Boolean
    Dim ft As Long, x As Double
    d = 0
    msg = ""
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        msg = "не указана дата выдачи"
    Case com.sun.star.table.CellContentType.TEXT
        ParseIssueDate = ParseDateText(oCell.getString(), d, msg)
    Case Else
        x = oCell.getValue()
        ft = gDoc.getNumberFormats().getByKey(oCell.NumberFormat).Type
        If (ft And com.sun.star.util.NumberFormat.DATE) = 0 Then
            msg = "дата выдачи должна быть датой, например 25.09.2026"
        Else
            ParseIssueDate = DateInRange(Int(x), d, msg)
        End If
    End Select
End Function

Function ParseDateText(ByVal s As String, ByRef d As Double, ByRef msg As String) As Boolean
    Dim a As Variant, dd As Integer, mm As Integer, yy As Integer
    s = Trim(s)
    a = Split(s, ".")
    msg = "«" & s & "» не является датой (формат дд.мм.гггг)"
    If UBound(a) <> 2 Then Exit Function
    If Not IsDigits(CStr(a(0))) Or Not IsDigits(CStr(a(1))) Or Not IsDigits(CStr(a(2))) Then Exit Function
    If Len(a(0)) > 2 Or Len(a(1)) > 2 Or Len(a(2)) <> 4 Then Exit Function
    dd = CInt(a(0))
    mm = CInt(a(1))
    yy = CInt(a(2))
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function
    If yy < 2000 Or yy > 2099 Then
        msg = "дата выдачи вне допустимого диапазона 2000–2099"
        Exit Function
    End If
    ' a day the month does not have (31.02): DateSerial would raise a Basic runtime error, so check it first
    If dd > 30 And (mm = 4 Or mm = 6 Or mm = 9 Or mm = 11) Then Exit Function
    If mm = 2 Then
        If dd > 29 Then Exit Function
        If dd = 29 And Not ((yy Mod 4 = 0 And yy Mod 100 <> 0) Or yy Mod 400 = 0) Then Exit Function
    End If
    ParseDateText = DateInRange(CDbl(DateSerial(yy, mm, dd)), d, msg)
End Function

Private Function DateInRange(x As Double, ByRef d As Double, ByRef msg As String) As Boolean
    If x < CDbl(DateSerial(2000, 1, 1)) Or x > CDbl(DateSerial(2099, 12, 31)) Then
        msg = "дата выдачи вне допустимого диапазона 2000–2099"
        Exit Function
    End If
    d = x
    msg = ""
    DateInRange = True
End Function

' ================================================================ row state

' kind of row r: "EMPTY", "OPEN" (not posted), "POSTED", "DELETED", "COPY", "FOREIGN" (a key without a WMS status)
Function RowKind(r As Long) As String
    Dim sh As Object, ctl As String, key As Object
    sh = IssuesSheet()
    ctl = sh.getCellByPosition(IC_CTL, r).getString()
    If Left(ctl, 5) = "КОПИЯ" Then
        RowKind = "COPY"
        Exit Function
    End If
    key = sh.getCellByPosition(IC_NO, r)
    If key.getType() = com.sun.star.table.CellContentType.EMPTY Then
        RowKind = "OPEN"
    ElseIf Left(ctl, Len(ST_POSTED)) = ST_POSTED Then
        RowKind = "POSTED"
    ElseIf ctl = ST_DELETED Then
        RowKind = "DELETED"
    Else
        RowKind = "FOREIGN"
    End If
End Function

' the key of a posted row (0 when it is not a whole positive number)
Private Function RowKey(r As Long) As Long
    Dim c As Object, x As Double
    c = IssuesSheet().getCellByPosition(IC_NO, r)
    If c.getType() <> com.sun.star.table.CellContentType.VALUE Then Exit Function
    x = c.getValue()
    If x >= 1 And x = Int(x) And x <= 2000000000 Then RowKey = CLng(x)
End Function

' writes a status into «Контроль» of an unposted row (never of a posted, cancelled or copy row)
Private Sub SetStatus(r As Long, s As String)
    Dim c As Object
    c = IssuesSheet().getCellByPosition(IC_CTL, r)
    If c.getString() = s Then Exit Sub
    If s = "" Then
        ClearCell(c)
    Else
        c.setString(s)
    End If
End Sub

' ================================================================ checks shared by preview, posting and correction

' Validates EI, quantity, date and recipient given as values (not cells); fills the candidate mC*.
' qtyAvailExtra: quantity this row already holds of the same EI (a correction may reuse it). "" or the reason.
Private Function CheckInputs(vEI As Variant, oQty As Object, vQty As Variant, oDate As Object, vDate As Variant, sWho As String, qtyAvailExtra As Double, extraEI As Long) As String
    Dim msg As String, st As Integer, p As String, s As Double, x As Double
    If Not NormalizeEI(vEI, mCn, mCcanon, msg) Then
        CheckInputs = msg
        Exit Function
    End If
    p = StockRowProblem(mCn, mCcanon)
    If p <> "" Then
        CheckInputs = p
        Exit Function
    End If
    mCdata = StockData(mCn)
    If Not StockQty(mCn, s) Then
        CheckInputs = "остаток " & mCcanon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    mCstock = s
    If mCn = extraEI Then s = s + qtyAvailExtra
    If s <= 0 Then
        CheckInputs = "остаток " & mCcanon & " равен 0 — выдача невозможна"
        Exit Function
    End If
    If Not (oQty Is Nothing) Then
        st = ParseQtyCell(oQty, mCqty, msg)
    ElseIf VarType(vQty) = 8 Then
        st = ParseQtyText(CStr(vQty), mCqty, msg)
    Else
        ' a number (dialog or test): the same contract as the parser — positive, at most 3 decimals
        x = CDbl(vQty)
        st = 0
        If x <= 0 Then
            msg = "количество должно быть больше 0"
            st = 5
        ElseIf Abs(x * 1000 - Int(x * 1000 + 0.5)) > 0.000001 Then
            msg = "не более " & QTY_MAX_DECIMALS & " знаков после запятой"
            st = 4
        Else
            mCqty = Int(x * 1000 + 0.5) / 1000
        End If
    End If
    If st <> 0 Then
        CheckInputs = msg
        Exit Function
    End If
    If mCqty > SysNum(SK_MAX_QTY) Then
        CheckInputs = "количество больше порога MAX_QTY (" & SysStr(SK_MAX_QTY) & ")"
        Exit Function
    End If
    If mCqty > s + 0.0000001 Then
        CheckInputs = "недостаточно: остаток " & mCcanon & " — " & QtyText(s) & " " & mCdata(SC_UNIT) & ", запрошено " & QtyText(mCqty)
        Exit Function
    End If
    If oDate Is Nothing Then
        If VarType(vDate) = 8 Then
            If Not ParseDateText(CStr(vDate), mCdate, msg) Then
                CheckInputs = msg
                Exit Function
            End If
        ElseIf Not DateInRange(Int(CDbl(vDate)), mCdate, msg) Then
            CheckInputs = msg
            Exit Function
        End If
    ElseIf Not ParseIssueDate(oDate, mCdate, msg) Then
        CheckInputs = msg
        Exit Function
    End If
    st = ResolveRecipient(sWho, mCwho, mCwhoNote)
    If st = 3 Or st = 4 Then
        CheckInputs = mCwhoNote
        Exit Function
    End If
    CheckInputs = ""
End Function

' ================================================================ light change handler (spec §13, D-017)

' «Выдачи» sheet event "Content changed": previews for the changed rows. Leaves at once for other columns, never scans
' the sheet, is guarded against re-entry and ignores changes made while an operation runs.
Sub OnIssuesChange(Optional oTarget As Variant)
    Dim addrs As Variant, i As Long
    If mInHandler Or gBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If oTarget.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        addrs = oTarget.getRangeAddresses()
    Else
        addrs = Array(oTarget.getRangeAddress())
    End If
    For i = 0 To UBound(addrs)
        HandleRange(addrs(i))
    Next i
    Exit Sub
EH:
    mInHandler = False
End Sub

Private Sub HandleRange(ra As Variant)
    Dim r As Long, r0 As Long, r1 As Long, n As Long, um As Object, sh As Object, cur As Object, last As Long, inCtx As Boolean
    If ra.EndRow < 1 Then Exit Sub
    ' only the inputs matter (A: a key that appeared by paste/fill, E F H I L: the preview and the row status)
    If Not (Touches(ra, IC_NO) Or Touches(ra, IC_QTY) Or Touches(ra, IC_PCT) Or Touches(ra, IC_DATE) Or Touches(ra, IC_WHO) Or Touches(ra, IC_EI)) Then Exit Sub
    mInHandler = True
    WmsInit()
    sh = IssuesSheet()
    r0 = ra.StartRow
    If r0 < 1 Then r0 = 1
    r1 = ra.EndRow
    If r1 - r0 + 1 > PREVIEW_MAX_ROWS Then
        ' a column cleared or filled far below the data: nothing to show beyond the used area
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        last = cur.getRangeAddress().EndRow
        If r1 > last Then r1 = last
    End If
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: подстановка")
    inCtx = True
    n = 0
    For r = r0 To r1
        n = n + 1
        If n <= PREVIEW_MAX_ROWS Then
            PreviewRow(r, Touches(ra, IC_NO), Touches(ra, IC_WHO))
        Else
            MassChangedRow(r)
        End If
    Next r
    um.leaveUndoContext()
    mInHandler = False
    Exit Sub
EH:
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    mInHandler = False
End Sub

Private Function Touches(ra As Variant, c As Integer) As Boolean
    Touches = (ra.StartColumn <= c And ra.EndColumn >= c)
End Function

' rows beyond PREVIEW_MAX_ROWS of one change: the previews are removed so that stale names cannot mislead
Private Sub MassChangedRow(r As Long)
    If RowKind(r) <> "OPEN" Then Exit Sub
    ClearPreview(r)
    SetStatus(r, "Проверьте строку: изменено слишком много строк сразу, предпросмотр не выполнен")
End Sub

Private Sub ClearPreview(r As Long)
    Dim sh As Object, cols As Variant, i As Integer
    sh = IssuesSheet()
    cols = Array(IC_NAME, IC_ART, IC_UNIT, IC_PLACE, IC_CAT, IC_BEFORE, IC_AFTER)
    For i = 0 To UBound(cols)
        If sh.getCellByPosition(cols(i), r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(cols(i), r))
    Next i
End Sub

' preview of one row: product data from the EI (C D G J K, balance in P, canonical L), expanded recipient (I) and the
' first problem among the filled inputs in «Контроль» (R). Posted, cancelled and copy rows are never touched.
Sub PreviewRow(r As Long, bKeyTouched As Boolean, bWhoTouched As Boolean)
    Dim sh As Object, kind As String, n As Long, canon As String, msg As String, p As String, d As Variant, s As Double
    Dim status As String, warn As String, who As String, full As String, note As String, st As Integer, q As Double, dt As Double
    Dim cEI As Object, cQty As Object, cDate As Object, known As Boolean
    sh = IssuesSheet()
    kind = RowKind(r)
    If kind = "FOREIGN" Or (bKeyTouched And (kind = "POSTED" Or kind = "DELETED")) Then
        ' a key that appeared through paste, fill or drag is a copy when WMS never issued it or it repeats (spec §15)
        CopyCheck(r)
        Exit Sub
    End If
    If kind <> "OPEN" Then Exit Sub
    ' EI → preview
    cEI = sh.getCellByPosition(IC_EI, r)
    If cEI.getType() = com.sun.star.table.CellContentType.EMPTY Then
        ClearPreview(r)
    ElseIf Not NormalizeEI(cEI.getString(), n, canon, msg) Then
        ClearPreview(r)
        status = msg
    Else
        p = StockRowProblem(n, canon)
        If p <> "" Then
            ClearPreview(r)
            status = p
        Else
            known = True
            d = StockData(n)
            If Not StockQty(n, s) Then s = 0
            SetIfDiff(sh.getCellByPosition(IC_EI, r), canon)
            SetIfDiff(sh.getCellByPosition(IC_NAME, r), d(SC_NAME))
            SetIfDiff(sh.getCellByPosition(IC_ART, r), d(SC_ART))
            SetIfDiff(sh.getCellByPosition(IC_UNIT, r), d(SC_UNIT))
            SetIfDiff(sh.getCellByPosition(IC_PLACE, r), d(SC_PLACE))
            SetIfDiff(sh.getCellByPosition(IC_CAT, r), d(SC_CAT))
            SetIfDiff(sh.getCellByPosition(IC_BEFORE, r), s)
            If sh.getCellByPosition(IC_AFTER, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(IC_AFTER, r))
            If s <= 0 Then status = "остаток " & canon & " равен 0 — выдача невозможна"
        End If
    End If
    ' quantity
    cQty = sh.getCellByPosition(IC_QTY, r)
    If status = "" And cQty.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        st = ParseQtyCell(cQty, q, msg)
        If st <> 0 Then
            status = msg
        ElseIf known And q > s + 0.0000001 Then
            status = "недостаточно: остаток " & canon & " — " & QtyText(s) & " " & d(SC_UNIT) & ", запрошено " & QtyText(q)
        End If
    End If
    If status = "" And sh.getCellByPosition(IC_PCT, r).getString() <> "" Then
        status = "«Количество в %» пока не используется — укажите количество в «Количество»"
    End If
    ' date
    cDate = sh.getCellByPosition(IC_DATE, r)
    If status = "" And cDate.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not ParseIssueDate(cDate, dt, msg) Then status = msg
    End If
    ' recipient: an abbreviation becomes the full name (spec §10)
    who = sh.getCellByPosition(IC_WHO, r).getString()
    If Trim(who) <> "" Then
        st = ResolveRecipient(who, full, note)
        If st = 1 And bWhoTouched Then SetIfDiff(sh.getCellByPosition(IC_WHO, r), full)
        If st = 3 And status = "" Then status = note
        If st = 2 Then warn = note
    End If
    If status <> "" Then
        SetStatus(r, "Ошибка: " & status)
    ElseIf warn <> "" Then
        SetStatus(r, "Внимание: " & warn)
    Else
        SetStatus(r, "")
    End If
End Sub

Private Sub SetIfDiff(c As Object, v As Variant)
    Select Case VarType(v)
    Case 2, 3, 4, 5, 6
        If c.getType() = com.sun.star.table.CellContentType.VALUE Then
            If Abs(c.getValue() - CDbl(v)) < 0.0000001 Then Exit Sub
        End If
        c.setValue(CDbl(v))
    Case Else
        If CStr(v) = "" Then
            If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(c)
        ElseIf c.getString() <> CStr(v) Or c.getType() <> com.sun.star.table.CellContentType.TEXT Then
            c.setString(CStr(v))
        End If
    End Select
End Sub

' a key in A that did not come from WMS: КОПИЯ when WMS never issued it or it appears more than once (spec §15)
Private Sub CopyCheck(r As Long)
    Dim sh As Object, k As Long, fa As Object, cnt As Double, why As String
    sh = IssuesSheet()
    k = RowKey(r)
    If k = 0 Then
        why = "ключ «" & sh.getCellByPosition(IC_NO, r).getString() & "» не выдавался WMS"
    ElseIf k >= SysNum(SK_NEXT_NO) Then
        why = "ключ " & k & " не выдавался WMS"
    Else
        fa = CreateUnoService("com.sun.star.sheet.FunctionAccess")
        cnt = fa.callFunction("COUNTIF", Array(sh.getCellRangeByPosition(IC_NO, 1, IC_NO, MAX_SHEET_ROW), k))
        If cnt > 1 Then why = "ключ " & k & " повторяется"
    End If
    If why = "" And RowKind(r) = "FOREIGN" Then why = "№ " & k & " без отметки о проведении"
    If why <> "" Then sh.getCellByPosition(IC_CTL, r).setString("КОПИЯ: " & why & " — не проводится, очистите строку кнопкой «Очистить»")
End Sub

' ================================================================ «Провести»

' Posts row r (0-based) of «Выдачи». Returns OK:<seq>, SKIP:<why> (nothing done), ERR:<why> (nothing done, the reason
' is in «Контроль»), BLOCKED:<why> or an ApplyOperation error (ERR-RB, ERR-CRITICAL).
Function IssuePostRow(r As Long) As String
    Dim why As String, sh As Object, kind As String, k As Long, res As String, s2 As Double
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        IssuePostRow = "ERR:выберите строку выдачи (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = IssuesSheet()
    kind = RowKind(r)
    Select Case kind
    Case "POSTED"
        IssuePostRow = "SKIP:уже проведено (№ " & RowKey(r) & ")"
        Exit Function
    Case "DELETED"
        IssuePostRow = "SKIP:выдача № " & RowKey(r) & " удалена (сторно)"
        Exit Function
    Case "COPY"
        IssuePostRow = "SKIP:строка помечена как КОПИЯ — очистите её кнопкой «Очистить»"
        Exit Function
    Case "FOREIGN"
        mInHandler = True
        CopyCheck(r)
        mInHandler = False
        IssuePostRow = "SKIP:в строке № без отметки о проведении — строка не проводится"
        Exit Function
    End Select
    why = PostingBlockReason()
    If why <> "" Then
        If Left(why, 8) = "BLOCKED:" Then SetStatus(r, "Не проведено: WMS заблокирована — " & Mid(why, 9) & ". См. лист «" & SH_MAIN & "»")
        IssuePostRow = why
        Exit Function
    End If
    If sh.getCellByPosition(IC_PCT, r).getString() <> "" Then
        why = "«Количество в %» пока не используется — укажите количество в «Количество»"
    Else
        why = CheckInputs(sh.getCellByPosition(IC_EI, r).getString(), sh.getCellByPosition(IC_QTY, r), Empty, sh.getCellByPosition(IC_DATE, r), Empty, _
            sh.getCellByPosition(IC_WHO, r).getString(), 0, 0)
    End If
    If why <> "" Then
        SetStatus(r, "Не проведено: " & why)
        IssuePostRow = "ERR:" & why
        Exit Function
    End If
    ' phase 2: one operation — the row, the registry balance and the counter change together
    k = CLng(SysNum(SK_NEXT_NO))
    s2 = Round3(mCstock - mCqty)
    PlanBegin("ISSUE", CStr(k))
    PlanField("NO", k)
    PlanField("EI", mCcanon)
    PlanField("QTY", mCqty)
    PlanField("UNIT", mCdata(SC_UNIT))
    PlanField("WHO", mCwho)
    PlanField("DATE", Format(mCdate, "YYYY-MM-DD"))
    PlanField("DOC", sh.getCellByPosition(IC_DOC, r).getString())
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", mCstock)
    PlanField("BAL_AFTER", s2)
    PlanRowValues(r, k, mCstock, s2, ST_POSTED)
    PlanLockBits(SH_ISSUES, r, 0, IC_LAST, ISSUE_LOCKS_POSTED)
    PlanSetValue(SH_STOCK, mCn, SC_QTY, s2, False)
    PlanSetValue(SYS_SHEET, SK_NEXT_NO, 1, k + 1, False)
    ' «Заказы».X of the receipt that created this EI shows its balance (Phase 3)
    WmsOrders.PlanStockMirror(mCn, s2)
    res = ApplyOperation(0, 0)
    If Left(res, 6) = "ERR-RB" Then SetStatus(r, "Не проведено: " & Mid(res, 8))
    IssuePostRow = res
    Exit Function
EH:
    ' only the checks can get here (ApplyOperation handles its own errors): nothing was written
    IssuePostRow = "ERR:внутренняя ошибка проверки строки: " & Error$ & " (код " & Err & ")"
    On Error Resume Next
    mInHandler = False
    SetStatus(r, "Не проведено: " & Mid(IssuePostRow, 5))
End Function

' the values of a posted row from the candidate: A key, product data, E F? H I L, balances P Q, status R
Private Sub PlanRowValues(r As Long, k As Long, before As Double, after As Double, status As String)
    PlanSetValue(SH_ISSUES, r, IC_NO, k, False)
    PlanSetValue(SH_ISSUES, r, IC_NAME, mCdata(SC_NAME), True)
    PlanSetValue(SH_ISSUES, r, IC_ART, mCdata(SC_ART), True)
    PlanSetValue(SH_ISSUES, r, IC_QTY, mCqty, True)
    PlanSetValue(SH_ISSUES, r, IC_UNIT, mCdata(SC_UNIT), True)
    PlanSetValue(SH_ISSUES, r, IC_DATE, mCdate, True)
    PlanSetValue(SH_ISSUES, r, IC_WHO, mCwho, True)
    PlanSetValue(SH_ISSUES, r, IC_PLACE, mCdata(SC_PLACE), True)
    PlanSetValue(SH_ISSUES, r, IC_CAT, mCdata(SC_CAT), True)
    PlanSetValue(SH_ISSUES, r, IC_EI, mCcanon, True)
    PlanSetValue(SH_ISSUES, r, IC_BEFORE, before, True)
    PlanSetValue(SH_ISSUES, r, IC_AFTER, after, True)
    PlanSetValue(SH_ISSUES, r, IC_CTL, status, True)
End Sub

' ================================================================ «Исправить» — one composite operation (spec §25)
' The old movement is reversed and the new one applied in the same journal line: a crash leaves either the old state
' (before the line) or the finished correction (replayed from the line) — never a state in between.

Function IssueFixRow(r As Long, vEI As Variant, vQty As Variant, vWho As Variant, vDate As Variant) As String
    Dim why As String, sh As Object, k As Long, n1 As Long, canon1 As String, msg As String, q1 As Double, s1 As Double
    Dim who1 As String, d1 As Double, res As String, pNew As Double, qNew As Double, oldCtl As String, p As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        IssueFixRow = "ERR:выберите строку выдачи (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = IssuesSheet()
    why = PostedRowProblem(r, "Исправить")
    If why <> "" Then
        IssueFixRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        IssueFixRow = why
        Exit Function
    End If
    k = RowKey(r)
    ' the old movement as the row holds it (locked cells, written by WMS)
    If Not NormalizeEI(sh.getCellByPosition(IC_EI, r).getString(), n1, canon1, msg) Then
        IssueFixRow = "ERR:строка № " & k & ": " & msg
        Exit Function
    End If
    p = StockRowProblem(n1, canon1)
    If p <> "" Then
        IssueFixRow = "ERR:" & p
        Exit Function
    End If
    q1 = sh.getCellByPosition(IC_QTY, r).getValue()
    who1 = sh.getCellByPosition(IC_WHO, r).getString()
    d1 = sh.getCellByPosition(IC_DATE, r).getValue()
    If Not StockQty(n1, s1) Then
        IssueFixRow = "ERR:остаток " & canon1 & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    ' the new movement: the old quantity of the same EI is available to it
    why = CheckInputs(vEI, Nothing, vQty, Nothing, vDate, CStr(vWho), q1, n1)
    If why <> "" Then
        IssueFixRow = "ERR:" & why
        Exit Function
    End If
    If mCn = n1 And Abs(mCqty - q1) < 0.0000001 And mCwho = who1 And mCdate = d1 Then
        IssueFixRow = "SKIP:ничего не изменилось"
        Exit Function
    End If
    If mCn = n1 Then
        pNew = Round3(s1 + q1)
    Else
        pNew = mCstock
    End If
    qNew = Round3(pNew - mCqty)
    oldCtl = sh.getCellByPosition(IC_CTL, r).getString()
    PlanBegin("ISSUE_FIX", CStr(k))
    PlanField("NO", k)
    PlanField("OLD_EI", canon1)
    PlanField("OLD_QTY", q1)
    PlanField("OLD_WHO", who1)
    PlanField("OLD_DATE", Format(d1, "YYYY-MM-DD"))
    PlanField("EI", mCcanon)
    PlanField("QTY", mCqty)
    PlanField("UNIT", mCdata(SC_UNIT))
    PlanField("WHO", mCwho)
    PlanField("DATE", Format(mCdate, "YYYY-MM-DD"))
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", pNew)
    PlanField("BAL_AFTER", qNew)
    PlanRowValues(r, k, pNew, qNew, ST_FIXED)
    PlanLockBits(SH_ISSUES, r, 0, IC_LAST, ISSUE_LOCKS_POSTED)
    If mCn = n1 Then
        PlanSetValue(SH_STOCK, n1, SC_QTY, qNew, False)
        WmsOrders.PlanStockMirror(n1, qNew)
    Else
        PlanSetValue(SH_STOCK, n1, SC_QTY, Round3(s1 + q1), False)
        PlanSetValue(SH_STOCK, mCn, SC_QTY, qNew, False)
        WmsOrders.PlanStockMirror(n1, Round3(s1 + q1))
        WmsOrders.PlanStockMirror(mCn, qNew)
    End If
    res = ApplyOperation(0, 0)
    IssueFixRow = res
    Exit Function
EH:
    IssueFixRow = "ERR:внутренняя ошибка проверки исправления: " & Error$ & " (код " & Err & ")"
End Function

' "" when row r is a posted (or corrected) issue that an action may change
Private Function PostedRowProblem(r As Long, sAction As String) As String
    Dim kind As String
    kind = RowKind(r)
    Select Case kind
    Case "POSTED"
        If RowKey(r) = 0 Then
            PostedRowProblem = "в строке нет правильного № выдачи"
        Else
            PostedRowProblem = ""
        End If
    Case "DELETED"
        PostedRowProblem = "выдача № " & RowKey(r) & " уже удалена (сторно)"
    Case "COPY", "FOREIGN"
        PostedRowProblem = "строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
    Case "OPEN"
        PostedRowProblem = "строка не проведена — «" & sAction & "» нужно только для проведённых выдач"
    Case Else
        PostedRowProblem = "строка пуста"
    End Select
End Function

' ================================================================ «Удалить» — storno of a posted issue
' The balance is returned, the journal gets ISSUE_DEL, the row keeps its № (never reused) and its data for the history
' and gets the unambiguous status «Удалено (сторно)». The row is not hidden.

Function IssueDeleteRow(r As Long) As String
    Dim why As String, sh As Object, k As Long, n1 As Long, canon1 As String, msg As String, q1 As Double, s1 As Double, p As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        IssueDeleteRow = "ERR:выберите строку выдачи (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = IssuesSheet()
    If RowKind(r) = "DELETED" Then
        IssueDeleteRow = "SKIP:выдача № " & RowKey(r) & " уже удалена (сторно)"
        Exit Function
    End If
    why = PostedRowProblem(r, "Удалить")
    If why <> "" Then
        IssueDeleteRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        IssueDeleteRow = why
        Exit Function
    End If
    k = RowKey(r)
    If Not NormalizeEI(sh.getCellByPosition(IC_EI, r).getString(), n1, canon1, msg) Then
        IssueDeleteRow = "ERR:строка № " & k & ": " & msg
        Exit Function
    End If
    p = StockRowProblem(n1, canon1)
    If p <> "" Then
        IssueDeleteRow = "ERR:" & p
        Exit Function
    End If
    If Not StockQty(n1, s1) Then
        IssueDeleteRow = "ERR:остаток " & canon1 & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    q1 = sh.getCellByPosition(IC_QTY, r).getValue()
    PlanBegin("ISSUE_DEL", CStr(k))
    PlanField("NO", k)
    PlanField("EI", canon1)
    PlanField("QTY", q1)
    PlanField("WHO", sh.getCellByPosition(IC_WHO, r).getString())
    PlanField("DATE", Format(sh.getCellByPosition(IC_DATE, r).getValue(), "YYYY-MM-DD"))
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", s1)
    PlanField("BAL_AFTER", Round3(s1 + q1))
    PlanSetValue(SH_STOCK, n1, SC_QTY, Round3(s1 + q1), False)
    PlanSetValue(SH_ISSUES, r, IC_CTL, ST_DELETED, False)
    PlanLockBits(SH_ISSUES, r, 0, IC_LAST, ISSUE_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(n1, Round3(s1 + q1))
    IssueDeleteRow = ApplyOperation(0, 0)
    Exit Function
EH:
    IssueDeleteRow = "ERR:внутренняя ошибка проверки удаления: " & Error$ & " (код " & Err & ")"
End Function

' ================================================================ «Очистить» — unposted row or copy (no accounting change)

Function IssueClearRow(r As Long) As String
    Dim sh As Object, kind As String, wasProt As Boolean, i As Integer, pr As Object
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        IssueClearRow = "ERR:выберите строку выдачи (не заголовок)"
        Exit Function
    End If
    If gBusy Then
        IssueClearRow = "BUSY:операция уже выполняется"
        Exit Function
    End If
    kind = RowKind(r)
    If kind = "POSTED" Then
        IssueClearRow = "ERR:строка проведена (№ " & RowKey(r) & ") — для отмены выдачи используйте «Удалить»"
        Exit Function
    End If
    If kind = "DELETED" Then
        IssueClearRow = "ERR:выдача № " & RowKey(r) & " удалена (сторно) — строка хранит историю и не очищается"
        Exit Function
    End If
    sh = IssuesSheet()
    mInHandler = True
    On Error GoTo EH
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    sh.getCellRangeByPosition(0, r, IC_LAST, r).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME _
        + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA)
    ' a row inserted between posted rows or a copy carries the protection of its origin: back to the input layout
    For i = 0 To IC_LAST
        pr = sh.getCellByPosition(i, r).CellProtection
        pr.IsLocked = (Mid(ISSUE_LOCKS_OPEN, i + 1, 1) = "1")
        sh.getCellByPosition(i, r).CellProtection = pr
    Next i
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
    IssueClearRow = "OK:" & IIf(kind = "OPEN", "строка очищена", "копия очищена")
    Exit Function
EH:
    IssueClearRow = "ERR:не удалось очистить строку: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
End Function

' ================================================================ startup note

' rows with an EI but no № (not posted), spec §23. Two Calc queries instead of a row loop: the empty cells of A within
' the used rows (blocks of consecutive rows), then the filled cells of L in each block. FunctionAccess COUNTIFS cannot be
' used: it copies every range argument into one temporary sheet, two whole columns overflow it. -1 when it fails;
' after UNPOSTED_MAX_BLOCKS scattered blocks the count so far is returned as a lower bound (gUnpostedMore = True).
Function UnpostedCount() As Long
    Dim sh As Object, cur As Object, last As Long, blocks As Variant, got As Variant, i As Long, j As Long, n As Long
    Dim flags As Long
    On Error GoTo EH
    gUnpostedMore = False
    sh = IssuesSheet()
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    If last < 1 Then Exit Function
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    blocks = sh.getCellRangeByPosition(IC_NO, 1, IC_NO, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= UNPOSTED_MAX_BLOCKS Then
            gUnpostedMore = True
            Exit For
        End If
        got = sh.getCellRangeByPosition(IC_EI, blocks(i).StartRow, IC_EI, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            n = n + got(j).EndRow - got(j).StartRow + 1
        Next j
    Next i
    UnpostedCount = n
    Exit Function
EH:
    UnpostedCount = -1
End Function
