' WmsAdjust — корректировки остатка: перемещение ЕИ (MOVE), списание (WRITE_OFF), инвентаризационная корректировка
' (INV_ADJ) — задание «FINAL WMS MARATHON», §2 «FINAL CORE». Лёгкий лист «Корректировки»: одна строка — одна операция
' WmsCore.ApplyOperation; исправление (*_FIX) и сторно (*_DEL) — тоже одна операция каждая.
'
'  - MOVE: ЕИ → новое место → дата → причина. Количество не меняется; старое и новое место — в журнале и в строке (M, J).
'    Место ЕИ меняется только операцией (тихого изменения нет): «Наличие».F защищён, каждое изменение — в журнале.
'  - WRITE_OFF: ЕИ → количество → дата → причина. Остаток уменьшается, не ниже 0. Это не выдача человеку.
'  - INV_ADJ: ЕИ → учётный остаток → фактический остаток → разница → дата → основание. Записывается корректировка
'    (разница), история не переписывается. Учётный остаток (I) — остаток, с которым сравнивали пересчёт: если он указан
'    (пересчёт по снимку, пакет инвентаризации) и за это время остаток ЕИ изменился — отказ, нужен пересчёт.
'
' Служебные данные: _ADJ — плотно, строка = № корректировки: вид, ЕИ, изменение остатка, состояние LIVE / STORNO,
' подсказка строки листа, места до/после, учётный и фактический остаток, дата. Номер строки листа — только подсказка:
' строка проверяется ключом (№), при несовпадении ищется формулой MATCH движка Calc (_IDX).
Option Explicit

' details of the last AdjustRowKind: the № of the row, the row hint was outdated, why FOREIGN
Global gAjN As Long
Global gAjMoved As Boolean
Global gAjWhy As String

Private mInHandler As Boolean

' the row being checked (CheckRowInputs): kind code, EI, registry row, balance, place, inputs
Private mKind As String
Private mEi As Long
Private mCanon As String
Private mSdata As Variant
Private mSbal As Double
Private mSplace As String
Private mQty As Double
Private mFact As Double
Private mBook As Double
Private mDiff As Double
Private mPlace As String
Private mDate As Double
Private mReason As String
Private mBatch As String
' 1.0: the write-off was refused as more than the balance (QtyAsk may read it as parts); the text CheckRowInputs reads
' instead of G when the storekeeper chose the parts
Private mShort As Boolean
Private mAltQty As String
' 1.0 (М1): the new EI of a split being posted
Private mNewEi As Long
Private mNewCanon As String
' 1.0 (З1): a return to the supplier — the supplier of the EI and the document of the return (J)
Private mSupplier As String
Private mDoc As String

' ================================================================ sheets and kinds

Function AdjustSheet() As Object
    AdjustSheet = gDoc.Sheets.getByName(SH_ADJUST)
End Function

Function AdjSheet() As Object
    AdjSheet = gDoc.Sheets.getByName(SH_ADJ)
End Function

' "" when the sheets of the corrections exist and have the expected layout (checked at every start, fail-closed)
Function AdjustSheetsProblem() As String
    Dim sh As Object
    If Not gDoc.Sheets.hasByName(SH_ADJUST) Then
        AdjustSheetsProblem = "нет листа «" & SH_ADJUST & "»"
        Exit Function
    End If
    If Not gDoc.Sheets.hasByName(SH_ADJ) Then
        AdjustSheetsProblem = "нет листа «" & SH_ADJ & "»"
        Exit Function
    End If
    sh = AdjustSheet()
    If sh.getCellByPosition(AC_NO, 0).getString() <> "№" Or sh.getCellByPosition(AC_KIND, 0).getString() <> "Вид" _
        Or sh.getCellByPosition(AC_CTL, 0).getString() <> "Контроль" Or sh.getCellByPosition(AC_NOTE, 0).getString() <> "Комментарий" Then
        AdjustSheetsProblem = "лист «" & SH_ADJUST & "»: заголовок не совпадает с A:S"
        Exit Function
    End If
    If AdjSheet().getCellByPosition(AJ_NO, 0).getString() <> "№" Or AdjSheet().getCellByPosition(AJ_DATE, 0).getString() <> "Дата" Then
        AdjustSheetsProblem = "служебный лист " & SH_ADJ & " повреждён (заголовок)"
        Exit Function
    End If
    AdjustSheetsProblem = ""
End Function

' the kind code (MOVE / WRITE_OFF / INV_ADJ) of a name of column B («Перемещение» …), "" when unknown
Function KindCode(ByVal s As String) As String
    Dim names As Variant, codes As Variant, i As Integer
    s = LCase(Trim(s))
    names = Split(ADJ_KIND_NAMES, "|")
    codes = Split(ADJ_KIND_CODES, "|")
    For i = 0 To UBound(names)
        If s = LCase(names(i)) Then
            KindCode = codes(i)
            Exit Function
        End If
    Next i
    KindCode = ""
End Function

Function KindName(code As String) As String
    Dim names As Variant, codes As Variant, i As Integer
    names = Split(ADJ_KIND_NAMES, "|")
    codes = Split(ADJ_KIND_CODES, "|")
    For i = 0 To UBound(codes)
        If code = codes(i) Then
            KindName = names(i)
            Exit Function
        End If
    Next i
    KindName = code
End Function

' ================================================================ service rows (dense addressing, the key is checked)

Function AdjRow(n As Long) As Variant
    AdjRow = AdjSheet().getCellRangeByPosition(0, n, AJ_LAST, n).getDataArray()(0)
End Function

Function AdjRegistered(n As Long, d As Variant) As Boolean
    If n < 1 Or n > MAX_SHEET_ROW Then Exit Function
    d = AdjRow(n)
    If VarType(d(AJ_NO)) <> 5 Then Exit Function
    AdjRegistered = (d(AJ_NO) = n)
End Function

' ================================================================ row model of «Корректировки»

' Kind of row r: "EMPTY", "OPEN" (not posted), "POSTED", "DELETED", "COPY" (marked), "FOREIGN" (a № that WMS did not
' register at this row). gAjN, gAjMoved, gAjWhy.
Function AdjustRowKind(r As Long) As String
    Dim sh As Object, d As Variant, x As Double, n As Long, rd As Variant, hint As Long, c As Object, i As Integer
    sh = AdjustSheet()
    gAjN = 0
    gAjMoved = False
    gAjWhy = ""
    d = sh.getCellRangeByPosition(0, r, AC_LAST, r).getDataArray()(0)
    If Left(CStr(d(AC_CTL)), 5) = "КОПИЯ" Then
        AdjustRowKind = "COPY"
        Exit Function
    End If
    If CStr(d(AC_NO)) = "" Then
        AdjustRowKind = "EMPTY"
        For i = AC_KIND To AC_REASON
            If i = AC_KIND Or i = AC_EI Or i = AC_QTY Or i = AC_FACT Or i = AC_PLACE Or i = AC_DATE Or i = AC_REASON Then
                If CStr(d(i)) <> "" Then
                    AdjustRowKind = "OPEN"
                    Exit For
                End If
            End If
        Next i
        Exit Function
    End If
    If VarType(d(AC_NO)) <> 5 Then
        gAjWhy = "№ «" & d(AC_NO) & "» не выдавался WMS"
        AdjustRowKind = "FOREIGN"
        Exit Function
    End If
    x = d(AC_NO)
    If x < 1 Or x <> Int(x) Or x > MAX_SHEET_ROW Then
        gAjWhy = "неверный № «" & d(AC_NO) & "»"
        AdjustRowKind = "FOREIGN"
        Exit Function
    End If
    n = CLng(x)
    If n >= SysNum(SK_NEXT_ADJ) Then
        gAjWhy = "№ корректировки " & n & " WMS не выдавала"
        AdjustRowKind = "FOREIGN"
        Exit Function
    End If
    If Not AdjRegistered(n, rd) Then
        gAjWhy = "корректировка № " & n & " не зарегистрирована WMS"
        AdjustRowKind = "FOREIGN"
        Exit Function
    End If
    If VarType(rd(AJ_ROW)) = 5 Then hint = CLng(rd(AJ_ROW)) Else hint = -1
    If hint <> r Then
        If hint >= 1 And hint <= MAX_SHEET_ROW Then
            c = sh.getCellByPosition(AC_NO, hint)
            If c.getType() = com.sun.star.table.CellContentType.VALUE Then
                If c.getValue() = n Then
                    gAjWhy = "корректировка № " & n & " повторяет строку " & (hint + 1)
                    AdjustRowKind = "FOREIGN"
                    Exit Function
                End If
            End If
        End If
        If WmsOrders.CountAdjustNo(n) > 1 Then
            If FirstAdjustRow(n) <> r Then
                gAjWhy = "корректировка № " & n & " повторяется"
                AdjustRowKind = "FOREIGN"
                Exit Function
            End If
        End If
        gAjMoved = True
    End If
    gAjN = n
    If CStr(rd(AJ_STATE)) = RV_STORNO Then AdjustRowKind = "DELETED" Else AdjustRowKind = "POSTED"
End Function

' the first row holding correction № n in A that is not marked КОПИЯ; -1 when none
Function FirstAdjustRow(n As Long) As Long
    Dim r As Long, start As Long, i As Integer, sh As Object
    sh = AdjustSheet()
    FirstAdjustRow = -1
    For i = 1 To 20
        r = WmsOrders.FindAdjustNoRow(n, start)
        If r < 1 Then Exit Function
        If Left(sh.getCellByPosition(AC_CTL, r).getString(), 5) <> "КОПИЯ" Then
            FirstAdjustRow = r
            Exit Function
        End If
        start = r
    Next i
End Function

' 0-based «Корректировки» row of correction № n: the hint of _ADJ checked by the key, otherwise found by MATCH; -1 when none
Function AdjustRowOf(n As Long, hint As Variant) As Long
    Dim sh As Object, c As Object, h As Long
    sh = AdjustSheet()
    If VarType(hint) = 5 Then h = CLng(hint) Else h = -1
    If h >= 1 And h <= MAX_SHEET_ROW Then
        c = sh.getCellByPosition(AC_NO, h)
        If c.getType() = com.sun.star.table.CellContentType.VALUE Then
            If c.getValue() = n Then
                AdjustRowOf = h
                Exit Function
            End If
        End If
    End If
    AdjustRowOf = FirstAdjustRow(n)
End Function

Sub CopyCheckAdjust(r As Long)
    If AdjustRowKind(r) <> "FOREIGN" Then Exit Sub
    AdjustSheet().getCellByPosition(AC_CTL, r).setString("КОПИЯ: " & gAjWhy & " — не проводится, очистите строку кнопкой «Очистить»")
End Sub

Private Sub SetCtl(r As Long, s As String)
    WmsOrders.SetIfDiff(AdjustSheet().getCellByPosition(AC_CTL, r), s)
End Sub

' ================================================================ the EI of a row

' the registry data of EI n (mEi, mCanon, mSdata, mSbal, mSplace): "" or the reason
Private Function LoadEI(v As Variant) As String
    Dim msg As String, p As String
    If Not WmsIssue.NormalizeEI(v, mEi, mCanon, msg) Then
        LoadEI = msg
        Exit Function
    End If
    p = WmsIssue.StockRowProblem(mEi, mCanon)
    If p <> "" Then
        LoadEI = p
        Exit Function
    End If
    mSdata = WmsIssue.StockData(mEi)
    If Not WmsIssue.StockQty(mEi, mSbal) Then
        LoadEI = "остаток " & mCanon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If CStr(mSdata(SC_STATE)) = EI_ST_STORNO Then
        LoadEI = mCanon & ": приход удалён (сторно) — корректировать нечего"
        Exit Function
    End If
    mSplace = CStr(mSdata(SC_PLACE))
    LoadEI = ""
End Function

' a fact (actual balance) of an inventory count: 0 is allowed (nothing was found)
Private Function ParseFact(c As Object, ByRef q As Double, ByRef msg As String) As Boolean
    Dim s As String, ft As Long, info As String
    If c.getType() = com.sun.star.table.CellContentType.TEXT Then
        s = Trim(c.getString())
        If s = "0" Or s = "0,0" Or s = "0.0" Or s = "0,00" Or s = "0.00" Or s = "0,000" Or s = "0.000" Then
            q = 0
            ParseFact = True
            Exit Function
        End If
    ElseIf c.getType() = com.sun.star.table.CellContentType.VALUE Then
        ft = gDoc.getNumberFormats().getByKey(c.NumberFormat).Type
        If c.getValue() = 0 And (ft And (2 Or 4 Or 32 Or 64 Or 128 Or 1024)) = 0 Then
            q = 0
            ParseFact = True
            Exit Function
        End If
    End If
    ' PRIME 0.8: a count in the parts of the packing of the EI («3 кан») is a balance in the unit of its document
    ParseFact = WmsPack.QtyAsEI(c, Empty, "фактический остаток (H)", mEi, mSbal, CStr(mSdata(SC_UNIT)), "FACT", 0, q, info, msg)
End Function

' the same for a value of a window or a test: a text or a number, 0 allowed
Private Function ParseFactValue(v As Variant, ByRef q As Double, ByRef msg As String) As Boolean
    Dim s As String, info As String
    If VarType(v) = 8 Then
        s = Trim(CStr(v))
        If s = "0" Or s = "0,0" Or s = "0.0" Or s = "0,00" Or s = "0.00" Or s = "0,000" Or s = "0.000" Then
            q = 0
            ParseFactValue = True
            Exit Function
        End If
    ElseIf CDbl(v) = 0 Then
        q = 0
        ParseFactValue = True
        Exit Function
    End If
    ParseFactValue = WmsPack.QtyAsEI(Nothing, v, "фактический остаток", mEi, mSbal, CStr(mSdata(SC_UNIT)), "FACT", 0, q, info, msg)
End Function

Private Function IsBlank(c As Object) As Boolean
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then
        IsBlank = True
    ElseIf c.getType() = com.sun.star.table.CellContentType.TEXT Then
        IsBlank = (Trim(c.getString()) = "")
    End If
End Function

' ================================================================ checks of a row before posting (shared with the preview)

' Validates the inputs of row r: "" or the reason; fills m*. bNeedAll: posting needs every field of the kind (the
' preview checks only what is filled).
Private Function CheckRowInputs(r As Long, bNeedAll As Boolean) As String
    Dim sh As Object, why As String, c As Object, msg As String, x As Double, info As String, sTo As String, sWhat As String, sVerb As String
    sh = AdjustSheet()
    mShort = False
    mKind = KindCode(sh.getCellByPosition(AC_KIND, r).getString())
    If mKind = "" Then
        If IsBlank(sh.getCellByPosition(AC_KIND, r)) Then
            CheckRowInputs = "не указан вид (B): Перемещение, Списание, Инвентаризация или Разделение"
        Else
            CheckRowInputs = "вид (B) «" & sh.getCellByPosition(AC_KIND, r).getString() & "» неизвестен: Перемещение, Списание, Инвентаризация или Разделение"
        End If
        Exit Function
    End If
    If IsBlank(sh.getCellByPosition(AC_EI, r)) Then
        CheckRowInputs = "не указан ЕИ (C)"
        Exit Function
    End If
    why = LoadEI(sh.getCellByPosition(AC_EI, r).getString())
    If why <> "" Then
        CheckRowInputs = why
        Exit Function
    End If
    ' only the field of the kind: a value in a field of another kind is an error, not silently ignored
    Select Case mKind
    Case "MOVE"
        If Not IsBlank(sh.getCellByPosition(AC_QTY, r)) Or Not IsBlank(sh.getCellByPosition(AC_FACT, r)) Or Not IsBlank(sh.getCellByPosition(AC_BOOK, r)) Then
            CheckRowInputs = "перемещение не меняет количество — очистите G «Списать», H «Фактический остаток» и I «Учётный остаток»"
            Exit Function
        End If
        mPlace = Trim(sh.getCellByPosition(AC_PLACE, r).getString())
        ' 1.0 (М10): «+1-2-5» / «−1-2-3» — one place more or less in the list of places of the EI
        If PlaceChange(mPlace) Then
            why = WmsView.PlaceEdit(mSplace, mPlace, sTo)
            If why <> "" Then
                CheckRowInputs = "новое место (J): " & why
                Exit Function
            End If
            mPlace = sTo
        End If
        If mPlace = "" Then
            If bNeedAll Then
                CheckRowInputs = "не указано новое место (J)"
                Exit Function
            End If
        ElseIf LCase(mPlace) = LCase(Trim(mSplace)) Then
            CheckRowInputs = "новое место «" & mPlace & "» совпадает с текущим местом " & mCanon & " — перемещать некуда"
            Exit Function
        End If
    Case "WRITE_OFF", "SUP_RET"
        If mKind = "SUP_RET" Then
            ' 1.0 (З1): G — how much goes back to the supplier (in parts too), J — the document of the return (when there is one)
            If Not IsBlank(sh.getCellByPosition(AC_FACT, r)) Or Not IsBlank(sh.getCellByPosition(AC_BOOK, r)) Then
                CheckRowInputs = "для возврата поставщику заполняются G «сколько вернуть» и, если есть, J — номер документа возврата; очистите H «Фактический остаток» " _
                    & "и I «Учётный остаток»"
                Exit Function
            End If
            mDoc = Trim(sh.getCellByPosition(AC_PLACE, r).getString())
            mSupplier = WmsSupply.EiSupplier(mEi)
            sWhat = "сколько вернуть (G)"
            sVerb = "вернуть "
        ElseIf Not IsBlank(sh.getCellByPosition(AC_FACT, r)) Or Not IsBlank(sh.getCellByPosition(AC_PLACE, r)) Or Not IsBlank(sh.getCellByPosition(AC_BOOK, r)) Then
            CheckRowInputs = "для списания заполняется только G «Списать» — очистите H «Фактический остаток», I «Учётный остаток» и J «Новое место»"
            Exit Function
        Else
            sWhat = "количество списания (G)"
            sVerb = "списать "
        End If
        c = sh.getCellByPosition(AC_QTY, r)
        If mAltQty <> "" Then
            If Not WmsPack.QtyAsEI(Nothing, mAltQty, sWhat, mEi, mSbal, CStr(mSdata(SC_UNIT)), "OUT", 0, mQty, info, msg) Then
                CheckRowInputs = msg
                Exit Function
            End If
        ElseIf Not IsBlank(c) Or bNeedAll Then
            If Not WmsPack.QtyAsEI(c, Empty, sWhat, mEi, mSbal, CStr(mSdata(SC_UNIT)), "OUT", 0, mQty, info, msg) Then
                CheckRowInputs = msg
                Exit Function
            End If
            If mQty > mSbal + 0.0000001 Then
                mShort = True
                CheckRowInputs = sVerb & WmsIssue.QtyText(mQty) & " нельзя: остаток " & mCanon & " — " & WmsPack.QtyWithParts(mEi, mSbal, CStr(mSdata(SC_UNIT)))
                Exit Function
            End If
        End If
    Case "SPLIT"
        ' 1.0 (М1): G — how much goes to the new EI (in parts too: «2 кан», «4,5 м»), J — its place (empty: the same place)
        If Not IsBlank(sh.getCellByPosition(AC_FACT, r)) Or Not IsBlank(sh.getCellByPosition(AC_BOOK, r)) Then
            CheckRowInputs = "для разделения заполняются G «сколько отделить» и, если нужно, J «Новое место» — очистите H «Фактический остаток» и I «Учётный остаток»"
            Exit Function
        End If
        c = sh.getCellByPosition(AC_QTY, r)
        If mAltQty <> "" Then
            If Not WmsPack.QtyAsEI(Nothing, mAltQty, "сколько отделить (G)", mEi, mSbal, CStr(mSdata(SC_UNIT)), "OUT", 0, mQty, info, msg) Then
                CheckRowInputs = msg
                Exit Function
            End If
        ElseIf Not IsBlank(c) Or bNeedAll Then
            If Not WmsPack.QtyAsEI(c, Empty, "сколько отделить (G)", mEi, mSbal, CStr(mSdata(SC_UNIT)), "OUT", 0, mQty, info, msg) Then
                CheckRowInputs = msg
                Exit Function
            End If
            If mQty > mSbal + 0.0000001 Then
                mShort = True
                CheckRowInputs = "отделить " & WmsIssue.QtyText(mQty) & " нельзя: остаток " & mCanon & " — " & WmsPack.QtyWithParts(mEi, mSbal, CStr(mSdata(SC_UNIT)))
                Exit Function
            End If
            If mQty >= mSbal - 0.0000001 Then
                CheckRowInputs = "отделить весь остаток " & mCanon & " нельзя — у ЕИ ничего не останется; чтобы переложить его целиком, выберите вид «Перемещение»"
                Exit Function
            End If
        End If
        mPlace = Trim(sh.getCellByPosition(AC_PLACE, r).getString())
        If mPlace = "" Then mPlace = mSplace
    Case "INV_ADJ"
        If Not IsBlank(sh.getCellByPosition(AC_QTY, r)) Or Not IsBlank(sh.getCellByPosition(AC_PLACE, r)) Then
            CheckRowInputs = "для инвентаризации заполняется H «Фактический остаток» — очистите G «Списать» и J «Новое место»"
            Exit Function
        End If
        mBook = mSbal
        c = sh.getCellByPosition(AC_BOOK, r)
        If Not IsBlank(c) Then
            If c.getType() <> com.sun.star.table.CellContentType.VALUE Then
                If Not ParseFact(c, x, msg) Then
                    CheckRowInputs = "учётный остаток (I): " & msg
                    Exit Function
                End If
            Else
                x = c.getValue()
            End If
            If Abs(x - mSbal) > 0.0000001 Then
                CheckRowInputs = "учётный остаток изменился: при пересчёте " & WmsIssue.QtyText(x) & ", сейчас " & WmsIssue.QtyText(mSbal) _
                    & " — после пересчёта были движения ЕИ; пересчитайте заново или очистите I «Учётный остаток»"
                Exit Function
            End If
        End If
        c = sh.getCellByPosition(AC_FACT, r)
        If Not IsBlank(c) Or bNeedAll Then
            If Not ParseFact(c, mFact, msg) Then
                CheckRowInputs = msg
                Exit Function
            End If
            ' 1.0: a recount that matches the book is recorded too (difference 0): the date of the last count of the EI is kept
            mDiff = WmsIssue.Round3(mFact - mBook)
            If Abs(mDiff) < 0.0000001 Then mDiff = 0
        End If
    End Select
    c = sh.getCellByPosition(AC_DATE, r)
    If Not IsBlank(c) Or bNeedAll Then
        If Not WmsOrders.ParseDateCellAs(c, "дата (K)", mDate, msg) Then
            CheckRowInputs = msg
            Exit Function
        End If
    End If
    mReason = Trim(sh.getCellByPosition(AC_REASON, r).getString())
    If mReason = "" And bNeedAll Then
        CheckRowInputs = "не указана " & IIf(mKind = "INV_ADJ", "основание (L)", "причина (L)")
        Exit Function
    End If
    mBatch = Trim(sh.getCellByPosition(AC_BATCH, r).getString())
    CheckRowInputs = ""
End Function

' 1.0 (М10): the new place is a change of the list of places («+1-2-5», «−1-2-3», «-1-2-3»): no place begins with a sign
Private Function PlaceChange(s As String) As Boolean
    Dim c As String
    c = Left(s, 1)
    PlaceChange = (c = "+" Or c = "-" Or c = "−" Or c = "–" Or c = "—")
End Function

' what the row will do (after CheckRowInputs): for «Контроль»
Private Function Describe() As String
    Dim u As String
    u = IIf(CStr(mSdata(SC_UNIT)) <> "", " " & mSdata(SC_UNIT), "")
    Select Case mKind
    Case "MOVE"
        Describe = mCanon & ": место «" & mSplace & "» → «" & mPlace & "», количество " & WmsIssue.QtyText(mSbal) & u & " не меняется"
    Case "WRITE_OFF"
        ' 1.0: an EI with a packing — the quantities in its parts too («списать 0,25 шт (1 кан), остаток 1 шт (4 кан) → …»)
        Describe = mCanon & ": списать " & WmsPack.QtyWithParts(mEi, mQty, CStr(mSdata(SC_UNIT))) & ", остаток " _
            & WmsPack.QtyWithParts(mEi, mSbal, CStr(mSdata(SC_UNIT))) & " → " & WmsPack.QtyWithParts(mEi, WmsIssue.Round3(mSbal - mQty), CStr(mSdata(SC_UNIT)))
    Case "SUP_RET"
        Describe = mCanon & ": вернуть поставщику" & IIf(mSupplier <> "", " «" & mSupplier & "»", "") & " " & WmsPack.QtyWithParts(mEi, mQty, CStr(mSdata(SC_UNIT))) _
            & ", остаток " & WmsPack.QtyWithParts(mEi, mSbal, CStr(mSdata(SC_UNIT))) & " → " & WmsPack.QtyWithParts(mEi, WmsIssue.Round3(mSbal - mQty), CStr(mSdata(SC_UNIT))) _
            & IIf(mDoc <> "", ", документ «" & mDoc & "»", "")
    Case "SPLIT"
        Describe = mCanon & ": отделить " & WmsPack.QtyWithParts(mEi, mQty, CStr(mSdata(SC_UNIT))) & " в новый ЕИ (место «" & mPlace & "»), у " & mCanon _
            & " останется " & WmsPack.QtyWithParts(mEi, WmsIssue.Round3(mSbal - mQty), CStr(mSdata(SC_UNIT)))
    Case "INV_ADJ"
        If mDiff = 0 Then
            Describe = mCanon & ": пересчёт совпал с учётом (" & WmsPack.QtyWithParts(mEi, mBook, CStr(mSdata(SC_UNIT))) _
                & ") — будет записан с разницей 0, остаток не меняется"
        Else
            Describe = mCanon & ": учётный " & WmsPack.QtyWithParts(mEi, mBook, CStr(mSdata(SC_UNIT))) & ", фактический " _
                & WmsPack.QtyWithParts(mEi, mFact, CStr(mSdata(SC_UNIT))) & ", разница " & IIf(mDiff > 0, "+", "") & WmsIssue.QtyText(mDiff) & u
        End If
    End Select
End Function

' "" when correction № n may be issued now: the dense _ADJ is behind the counter and no row of the sheet holds n
Private Function NewAdjProblem(n As Long) As String
    If n < 1 Or n >= MAX_SHEET_ROW Then
        NewAdjProblem = "номера корректировок исчерпаны (NEXT_ADJ " & n & ")"
    ElseIf WmsOrders.LastRow(AdjSheet()) >= n Then
        NewAdjProblem = "служебная таблица корректировок " & SH_ADJ & " не согласована со счётчиком NEXT_ADJ " & n & " — нужна самопроверка"
    ElseIf WmsOrders.CountAdjustNo(n) > 0 Then
        NewAdjProblem = "№ корректировки " & n & " уже записан на листе «" & SH_ADJUST & "» — новый номер не выдаётся. Нужна проверка ответственным"
    Else
        NewAdjProblem = ""
    End If
End Function

' ================================================================ plan helpers

Private Sub PlanAdjRow(n As Long, sState As String, r As Long, delta As Double, sFrom As String, sTo As String)
    PlanSetValue(SH_ADJ, n, AJ_NO, n, False)
    PlanSetValue(SH_ADJ, n, AJ_KIND, mKind, False)
    PlanSetValue(SH_ADJ, n, AJ_EI, mCanon, False)
    PlanSetValue(SH_ADJ, n, AJ_QTY, delta, False)
    PlanSetValue(SH_ADJ, n, AJ_STATE, sState, False)
    PlanSetValue(SH_ADJ, n, AJ_ROW, r, False)
    If sFrom <> "" Then PlanSetValue(SH_ADJ, n, AJ_FROM, sFrom, False)
    If sTo <> "" Then PlanSetValue(SH_ADJ, n, AJ_TO, sTo, False)
    If mKind = "INV_ADJ" Then
        PlanSetValue(SH_ADJ, n, AJ_BOOK, mBook, False)
        PlanSetValue(SH_ADJ, n, AJ_FACT, mFact, False)
    End If
    PlanSetValue(SH_ADJ, n, AJ_DATE, mDate, False)
    If mKind = "SPLIT" Then PlanSetValue(SH_ADJ, n, AJ_NEW, mNewCanon, False)
End Sub

' 1.0 (М1): the registry row of the new EI of a split — the card of the source (name, article, unit, category, the kind of
' its source), its own place and balance; the packing of the source and the formula «Остаток частями»
Private Sub PlanSplitStock(m As Long, q As Double, nAdj As Long)
    Dim pu As String, pn As Double, pk As Object
    PlanSetValue(SH_STOCK, m, SC_EI, mNewCanon, False)
    PlanSetValue(SH_STOCK, m, SC_NAME, mSdata(SC_NAME), False)
    If CStr(mSdata(SC_ART)) <> "" Then PlanSetValue(SH_STOCK, m, SC_ART, mSdata(SC_ART), False)
    PlanSetValue(SH_STOCK, m, SC_UNIT, mSdata(SC_UNIT), False)
    PlanSetValue(SH_STOCK, m, SC_QTY, q, False)
    PlanSetValue(SH_STOCK, m, SC_PLACE, mPlace, False)
    If CStr(mSdata(SC_CAT)) <> "" Then PlanSetValue(SH_STOCK, m, SC_CAT, mSdata(SC_CAT), False)
    PlanSetValue(SH_STOCK, m, SC_STATE, EI_ST_ACTIVE, False)
    PlanSetValue(SH_STOCK, m, SC_SRC, "Разделение " & mCanon & " (корректировка № " & nAdj & ")", False)
    If CStr(mSdata(SC_STYPE)) <> "" Then PlanSetValue(SH_STOCK, m, SC_STYPE, mSdata(SC_STYPE), False)
    If gDoc.Sheets.hasByName(SH_PACK) Then
        If WmsPack.PackOf(mEi, pu, pn) Then
            pk = gDoc.Sheets.getByName(SH_PACK)
            PlanSetValue(SH_PACK, m, PK_EI, mNewCanon, False)
            PlanSetValue(SH_PACK, m, PK_UNIT, pu, False)
            PlanSetValue(SH_PACK, m, PK_N, pn, False)
            PlanSetValue(SH_PACK, m, PK_NOTE, "как у " & mCanon & " (разделение)", False)
            PlanFormula(SH_STOCK, m, 10, WmsPack.PartsFormula(m))
        End If
    End If
End Sub

' 1.0: the questions after CheckRowInputs (see WmsIssue.AskBeforePost): the readings of a quantity WMS never guesses
' (WmsPack.QtyAsk) and an unusual date (WmsAsk.DateAsk). why — the refusal of CheckRowInputs. "" to post, SKIP:/ERR:…
Private Function AskAdjust(why As String, r As Long) As String
    Dim st As Integer, msg As String, info As String, u As String, nm As String
    If why <> "" And Not mShort Then
        AskAdjust = "ERR:" & why
        Exit Function
    End If
    u = CStr(mSdata(SC_UNIT))
    nm = CStr(mSdata(SC_NAME))
    Select Case mKind
    Case "MOVE"
        ' 1.0: the new place checked against the map and «Места»
        msg = WmsMap.PlaceAsk(mPlace, "перемещение не проведено")
        If msg <> "" Then
            AskAdjust = msg
            Exit Function
        End If
    Case "WRITE_OFF", "SUP_RET"
        If why <> "" Then
            msg = why
            st = WmsPack.QtyAsk(mEi, nm, u, IIf(mKind = "SUP_RET", "Вернуть", "Списать"), "OUT", mSbal, 0, False, mQty, info, msg)
            If st = 0 Then
                ' the parts chosen: the row is checked again with them (the date and the reason were not read yet)
                mAltQty = WmsPack.PartsText(gPkParts, gPkUnit)
                msg = CheckRowInputs(r, True)
                mAltQty = ""
                If msg <> "" Then st = 2
            End If
        Else
            st = WmsPack.QtyAsk(mEi, nm, u, IIf(mKind = "SUP_RET", "Вернуть", "Списать"), "OUT", mSbal, 0, True, mQty, info, msg)
        End If
    Case "SPLIT"
        If why <> "" Then
            msg = why
            st = WmsPack.QtyAsk(mEi, nm, u, "Отделить", "OUT", mSbal, 0, False, mQty, info, msg)
            If st = 0 Then
                mAltQty = WmsPack.PartsText(gPkParts, gPkUnit)
                msg = CheckRowInputs(r, True)
                mAltQty = ""
                If msg <> "" Then st = 2
            End If
        Else
            st = WmsPack.QtyAsk(mEi, nm, u, "Отделить", "OUT", mSbal, 0, True, mQty, info, msg)
        End If
        If st = 0 And LCase(mPlace) <> LCase(Trim(mSplace)) Then
            msg = WmsMap.PlaceAsk(mPlace, "разделение не проведено")
            If msg <> "" Then
                AskAdjust = msg
                Exit Function
            End If
        End If
    Case "INV_ADJ"
        st = WmsPack.QtyAsk(mEi, nm, u, "Факт", "FACT", mBook, 0, True, mFact, info, msg)
        If st = 0 Then
            mDiff = WmsIssue.Round3(mFact - mBook)
            If Abs(mDiff) < 0.0000001 Then mDiff = 0
        End If
    End Select
    If st = 1 Then
        AskAdjust = "SKIP:" & msg
    ElseIf st = 2 Then
        AskAdjust = "ERR:" & msg
    ElseIf WmsAsk.DateAsk(mDate, "корректировки", "Провести", msg) <> 0 Then
        AskAdjust = "SKIP:" & msg
    End If
End Function

' the common journal fields of a correction
Private Sub PlanCommonFields(n As Long, r As Long)
    PlanField("ADJ", n)
    PlanField("KIND", mKind)
    PlanField("EI", mCanon)
    PlanField("DATE", Format(mDate, "YYYY-MM-DD"))
    PlanField("REASON", mReason)
    PlanField("ROW", r + 1)
    If mBatch <> "" Then PlanField("BATCH", mBatch)
End Sub

' ================================================================ «Провести» — MOVE / WRITE_OFF / INV_ADJ

' Posts row r (0-based). OK:<seq>, SKIP:<why>, ERR:<why> (a problem of the row, the reason in «Контроль»),
' ERR-SYS:<why> (a problem of WMS itself), BLOCKED:<why> or an ApplyOperation error.
Function AdjustPostRow(r As Long) As String
    Dim why As String, kind As String, n As Long, res As String, b2 As Double, sh As Object, delta As Double
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        AdjustPostRow = "ERR:выберите строку корректировки (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = AdjustSheet()
    kind = AdjustRowKind(r)
    Select Case kind
    Case "POSTED"
        AdjustPostRow = "SKIP:уже проведено (корректировка № " & gAjN & ")"
        Exit Function
    Case "DELETED"
        AdjustPostRow = "SKIP:корректировка № " & gAjN & " удалена (сторно)"
        Exit Function
    Case "COPY"
        AdjustPostRow = "SKIP:строка помечена как КОПИЯ — очистите её кнопкой «Очистить»"
        Exit Function
    Case "FOREIGN"
        mInHandler = True
        CopyCheckAdjust(r)
        mInHandler = False
        AdjustPostRow = "SKIP:в строке № корректировки, которую WMS здесь не проводила — строка не проводится"
        Exit Function
    Case "EMPTY"
        AdjustPostRow = "ERR:строка пуста"
        Exit Function
    End Select
    why = PostingBlockReason()
    If why <> "" Then
        If Left(why, 8) = "BLOCKED:" Then SetCtl(r, "Не проведено: WMS заблокирована — " & Mid(why, 9) & ". См. лист «" & SH_MAIN & "»")
        AdjustPostRow = why
        Exit Function
    End If
    why = CheckRowInputs(r, True)
    ' 1.0: «3» of an EI with a packing (3 шт or 3 кан), a percent below 1, a date in the future or a year back — asked
    why = AskAdjust(why, r)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & Mid(why, InStr(why, ":") + 1))
        If Left(why, 5) = "SKIP:" Then AdjustPostRow = "SKIP:корректировка не проведена — " & Mid(why, 6) Else AdjustPostRow = why
        Exit Function
    End If
    n = CLng(SysNum(SK_NEXT_ADJ))
    why = NewAdjProblem(n)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        AdjustPostRow = "ERR-SYS:" & why
        Exit Function
    End If
    ' 1.0 (М1): a split issues a new EI — the counter must be safe, as for a receipt (D-059)
    If mKind = "SPLIT" Then
        mNewEi = CLng(SysNum(SK_NEXT_EI))
        why = WmsOrders.NewEIProblem(mNewEi)
        If why <> "" Then
            SetCtl(r, "Не проведено: " & why)
            AdjustPostRow = "ERR-SYS:" & why
            Exit Function
        End If
        mNewCanon = WmsIssue.EiCanon(mNewEi)
    End If
    b2 = mSbal
    delta = 0
    If mKind = "WRITE_OFF" Or mKind = "SPLIT" Or mKind = "SUP_RET" Then delta = -mQty
    If mKind = "INV_ADJ" Then delta = mDiff
    b2 = WmsIssue.Round3(mSbal + delta)
    PlanBegin(mKind, CStr(n))
    PlanCommonFields(n, r)
    Select Case mKind
    Case "SPLIT"
        PlanField("QTY", mQty)
        PlanField("NEW_EI", mNewCanon)
        PlanField("FROM", mSplace)
        PlanField("TO", mPlace)
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        WmsPack.PlanQtyFields()
    Case "MOVE"
        PlanField("FROM", mSplace)
        PlanField("TO", mPlace)
        PlanField("BAL", mSbal)
    Case "WRITE_OFF", "SUP_RET"
        PlanField("QTY", mQty)
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        WmsPack.PlanQtyFields()
        If mKind = "SUP_RET" Then
            PlanField("SUPPLIER", mSupplier)
            If mDoc <> "" Then PlanField("DOC", mDoc)
        End If
    Case "INV_ADJ"
        PlanField("BOOK", mBook)
        PlanField("FACT", mFact)
        PlanField("DIFF", mDiff)
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        WmsPack.PlanQtyFields()
    End Select
    ' one operation, the writes grouped by sheet: _SYS → «Наличие» → _ADJ → «Корректировки» → «Заказы».X
    PlanSetValue(SYS_SHEET, SK_NEXT_ADJ, 1, n + 1, False)
    If mKind = "SPLIT" Then PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, mNewEi + 1, False)
    If mKind = "MOVE" Then
        PlanSetValue(SH_STOCK, mEi, SC_PLACE, mPlace, False)
        PlanAdjRow(n, RV_LIVE, r, 0, mSplace, mPlace)
    ElseIf mKind = "SPLIT" Then
        PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
        PlanSplitStock(mNewEi, mQty, n)
        PlanAdjRow(n, RV_LIVE, r, delta, mSplace, mPlace)
    Else
        PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
        PlanAdjRow(n, RV_LIVE, r, delta, mSplace, "")
    End If
    PlanSetValue(SH_ADJUST, r, AC_NO, n, False)
    PlanSetValue(SH_ADJUST, r, AC_KIND, KindName(mKind), True)
    PlanSetValue(SH_ADJUST, r, AC_EI, mCanon, True)
    PlanSetValue(SH_ADJUST, r, AC_NAME, mSdata(SC_NAME), True)
    PlanSetValue(SH_ADJUST, r, AC_ART, mSdata(SC_ART), True)
    PlanSetValue(SH_ADJUST, r, AC_UNIT, mSdata(SC_UNIT), True)
    Select Case mKind
    Case "MOVE"
        PlanSetValue(SH_ADJUST, r, AC_PLACE, mPlace, True)
    Case "WRITE_OFF", "SUP_RET"
        PlanSetValue(SH_ADJUST, r, AC_QTY, mQty, True)
        PlanSetValue(SH_ADJUST, r, AC_DIFF, delta, True)
        If mKind = "SUP_RET" And mDoc <> "" Then PlanSetValue(SH_ADJUST, r, AC_PLACE, mDoc, True)
    Case "SPLIT"
        PlanSetValue(SH_ADJUST, r, AC_QTY, mQty, True)
        PlanSetValue(SH_ADJUST, r, AC_PLACE, mPlace, True)
        PlanSetValue(SH_ADJUST, r, AC_DIFF, delta, True)
    Case "INV_ADJ"
        PlanSetValue(SH_ADJUST, r, AC_FACT, mFact, True)
        PlanSetValue(SH_ADJUST, r, AC_BOOK, mBook, True)
        PlanSetValue(SH_ADJUST, r, AC_DIFF, mDiff, True)
    End Select
    PlanSetValue(SH_ADJUST, r, AC_DATE, mDate, True)
    PlanSetValue(SH_ADJUST, r, AC_REASON, mReason, True)
    PlanSetValue(SH_ADJUST, r, AC_FROM, mSplace, True)
    PlanSetValue(SH_ADJUST, r, AC_BEFORE, mSbal, True)
    PlanSetValue(SH_ADJUST, r, AC_AFTER, b2, True)
    ' 1.0 (М1): the row of a split names the new EI (the start «Проведено» is what every check reads)
    ' 1.0 (З1): a return names its supplier
    If mKind = "SUP_RET" Then
        PlanSetValue(SH_ADJUST, r, AC_CTL, ST_POSTED & IIf(mSupplier <> "", " · поставщику " & mSupplier, ""), True)
    Else
        PlanSetValue(SH_ADJUST, r, AC_CTL, ST_POSTED & IIf(mKind = "SPLIT", " · новый " & mNewCanon, ""), True)
    End If
    ' the batch the row came from (a tool's file) is restored with the row after a crash, like the other inputs
    If mBatch <> "" Then PlanSetValue(SH_ADJUST, r, AC_BATCH, mBatch, True)
    PlanLockBits(SH_ADJUST, r, 0, AC_LAST, ADJUST_LOCKS_POSTED)
    If mKind <> "MOVE" Then WmsOrders.PlanStockMirror(mEi, b2)
    res = ApplyOperation(0, 0)
    If Left(res, 6) = "ERR-RB" Then SetCtl(r, "Не проведено: " & Mid(res, 8))
    AdjustPostRow = res
    Exit Function
EH:
    AdjustPostRow = "ERR-SYS:внутренняя ошибка проверки строки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    On Error Resume Next
    mInHandler = False
    SetCtl(r, "Не проведено: " & Mid(AdjustPostRow, 9))
End Function

' «Провести» for a block of rows (D-059): every visible open row with a kind (B) is posted by its own operation; a
' problem of one row does not stop the others; a system error of WMS stops the whole block.
Function AdjustPostRange(r0 As Long, r1 As Long) As String
    Dim r As Long, res As String, nOk As Long, nErr As Long, nSkip As Long, sh As Object, firstErr As String, sysErr As String
    Dim stopRow As Long
    On Error GoTo EH
    sh = AdjustSheet()
    For r = r0 To r1
        If sh.getRows().getByIndex(r).IsVisible Then
            If IsBlank(sh.getCellByPosition(AC_KIND, r)) Then
                nSkip = nSkip + 1
            ElseIf AdjustRowKind(r) = "OPEN" Then
                res = AdjustPostRow(r)
                If WmsReceipt.IsSystemResult(res) Then
                    If Left(res, 3) = "OK:" Then nOk = nOk + 1
                    sysErr = "строка " & (r + 1) & ": " & res
                    stopRow = r + 1
                    Exit For
                ElseIf Left(res, 3) = "OK:" Then
                    nOk = nOk + 1
                ElseIf Left(res, 5) = "SKIP:" Then
                    nSkip = nSkip + 1
                Else
                    nErr = nErr + 1
                    If firstErr = "" Then firstErr = "строка " & (r + 1) & ": " & Mid(res, InStr(res, ":") + 1)
                End If
            Else
                nSkip = nSkip + 1
            End If
        End If
    Next r
DONE:
    AdjustPostRange = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";STOP=" & stopRow & Chr(10) & firstErr & Chr(10) & sysErr
    Exit Function
EH:
    sysErr = "строка " & (r + 1) & ": ERR-SYS:внутренняя ошибка проведения блока: " & Error$ & " (код " & Err & ")"
    stopRow = r + 1
    Resume DONE
End Function

' "" when a row of this kind is a posted correction that an action may change
Private Function PostedRowProblem(kind As String, sAction As String) As String
    Select Case kind
    Case "POSTED"
        PostedRowProblem = ""
    Case "DELETED"
        PostedRowProblem = "корректировка № " & gAjN & " уже удалена (сторно)"
    Case "COPY", "FOREIGN"
        PostedRowProblem = "строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
    Case "OPEN"
        PostedRowProblem = "строка не проведена — «" & sAction & "» нужно только для проведённых корректировок"
    Case Else
        PostedRowProblem = "строка пуста"
    End Select
End Function

' the posted correction of row r (kind "POSTED" or "DELETED" known): mKind, mEi, mCanon, registry data; "" or the reason
Private Function LoadPosted(r As Long, rd As Variant) As String
    Dim msg As String, p As String
    If Not AdjRegistered(gAjN, rd) Then
        LoadPosted = "ERR-SYS:корректировка № " & gAjN & " не найдена в " & SH_ADJ
        Exit Function
    End If
    mKind = CStr(rd(AJ_KIND))
    If mKind <> "MOVE" And mKind <> "WRITE_OFF" And mKind <> "INV_ADJ" And mKind <> "SPLIT" And mKind <> "SUP_RET" Then
        LoadPosted = "ERR-SYS:в " & SH_ADJ & " у корректировки № " & gAjN & " неизвестный вид «" & mKind & "»"
        Exit Function
    End If
    mCanon = CStr(rd(AJ_EI))
    If Not WmsOrders.StrictEI(mCanon, mEi) Then
        LoadPosted = "ERR-SYS:в " & SH_ADJ & " у корректировки № " & gAjN & " неверный ЕИ «" & mCanon & "»"
        Exit Function
    End If
    p = WmsIssue.StockRowProblem(mEi, mCanon)
    If p <> "" Then
        LoadPosted = "ERR:" & p
        Exit Function
    End If
    mSdata = WmsIssue.StockData(mEi)
    If Not WmsIssue.StockQty(mEi, mSbal) Then
        LoadPosted = "ERR:остаток " & mCanon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    mSplace = CStr(mSdata(SC_PLACE))
    mBatch = ""
    LoadPosted = ""
End Function

' ================================================================ «Исправить» — MOVE_FIX / WRITE_OFF_FIX / INV_ADJ_FIX
' One composite operation: the old correction is reversed and the new one applied in one journal line. The EI and the
' kind stay (another EI or kind: storno and a new correction). vMain: MOVE — the new place, WRITE_OFF — the quantity,
' INV_ADJ — the actual balance (the book balance of the count stays). A move is corrected only while the EI is still
' where it moved it; a correction never makes the balance negative.

Function AdjustFixRow(r As Long, vMain As Variant, vDate As Variant, vReason As Variant) As String
    Dim why As String, kind As String, n As Long, rd As Variant, msg As String, sh As Object, res As String
    Dim q1 As Double, d1 As Double, reason1 As String, to1 As String, from1 As String, delta1 As Double, delta2 As Double
    Dim before As Double, b2 As Double, info As String, st As Integer, to2 As String, sVerb As String, sWhat As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        AdjustFixRow = "ERR:выберите строку корректировки (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = AdjustSheet()
    kind = AdjustRowKind(r)
    why = PostedRowProblem(kind, "Исправить")
    If why <> "" Then
        AdjustFixRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        AdjustFixRow = why
        Exit Function
    End If
    n = gAjN
    why = LoadPosted(r, rd)
    If why <> "" Then
        AdjustFixRow = why
        Exit Function
    End If
    ' 1.0 (М1): a split is not corrected — its new EI has its own card and label; cancel it and split again
    If mKind = "SPLIT" Then
        AdjustFixRow = "ERR:разделение № " & n & " не исправляется — удалите его (сторно; можно, пока у нового ЕИ " & rd(AJ_NEW) _
            & " не было движений) и проведите заново"
        Exit Function
    End If
    delta1 = rd(AJ_QTY)
    d1 = rd(AJ_DATE)
    reason1 = sh.getCellByPosition(AC_REASON, r).getString()
    from1 = CStr(rd(AJ_FROM))
    to1 = CStr(rd(AJ_TO))
    ' the date and the reason: always; the main value by the kind
    If VarType(vDate) = 8 Then
        If Not WmsOrders.ParseDateTextAs(CStr(vDate), "дата", mDate, msg) Then
            AdjustFixRow = "ERR:" & msg
            Exit Function
        End If
    Else
        mDate = Int(CDbl(vDate))
    End If
    mReason = Trim(CStr(vReason))
    If mReason = "" Then
        AdjustFixRow = "ERR:не указана " & IIf(mKind = "INV_ADJ", "основание", "причина")
        Exit Function
    End If
    Select Case mKind
    Case "MOVE"
        mPlace = Trim(CStr(vMain))
        ' 1.0 (М10): «+1-2-5» / «−1-2-3» — the places the move wrote (the window shows them) with one place more or less
        If PlaceChange(mPlace) Then
            msg = WmsView.PlaceEdit(to1, mPlace, to2)
            If msg <> "" Then
                AdjustFixRow = "ERR:новое место: " & msg
                Exit Function
            End If
            mPlace = to2
        End If
        If mPlace = "" Then
            AdjustFixRow = "ERR:не указано новое место"
            Exit Function
        End If
        If LCase(Trim(mSplace)) <> LCase(Trim(to1)) Then
            AdjustFixRow = "ERR:после перемещения № " & n & " место " & mCanon & " менялось (сейчас «" & mSplace & "», перемещено в «" & to1 _
                & "») — исправить нельзя; оформите новое перемещение"
            Exit Function
        End If
        If LCase(mPlace) = LCase(Trim(from1)) Then
            AdjustFixRow = "ERR:новое место совпадает с местом до перемещения («" & from1 & "») — чтобы отменить перемещение, нажмите «Удалить»"
            Exit Function
        End If
        If mPlace = to1 And mDate = Int(d1) And mReason = reason1 Then
            AdjustFixRow = "SKIP:ничего не изменилось"
            Exit Function
        End If
        PlanBegin("MOVE_FIX", CStr(n))
        PlanCommonFields(n, r)
        PlanField("FROM", from1)
        PlanField("OLD_TO", to1)
        PlanField("TO", mPlace)
        PlanField("OLD_DATE", Format(d1, "YYYY-MM-DD"))
        PlanField("OLD_REASON", reason1)
        If mPlace <> mSplace Then PlanSetValue(SH_STOCK, mEi, SC_PLACE, mPlace, False)
        PlanSetValue(SH_ADJ, n, AJ_TO, mPlace, False)
        PlanSetValue(SH_ADJUST, r, AC_PLACE, mPlace, False)
    Case "WRITE_OFF", "SUP_RET"
        q1 = -delta1
        If mKind = "SUP_RET" Then
            sVerb = "вернуть "
            sWhat = "сколько вернуть"
        Else
            sVerb = "списать "
            sWhat = "количество списания"
        End If
        If Not WmsPack.QtyAsEI(Nothing, vMain, sWhat, mEi, WmsIssue.Round3(mSbal + q1), CStr(mSdata(SC_UNIT)), "OUT", 0, mQty, info, msg) Then
            AdjustFixRow = "ERR:" & msg
            Exit Function
        End If
        ' 1.0: a new quantity «1» of an EI with a packing or a percent below 1 — asked as at the posting
        If Abs(mQty - q1) > 0.0000001 Then
            msg = sVerb & WmsIssue.QtyText(mQty) & " нельзя: остаток " & mCanon & " без этой корректировки — " _
                & WmsPack.QtyWithParts(mEi, WmsIssue.Round3(mSbal + q1), CStr(mSdata(SC_UNIT)))
            st = WmsPack.QtyAsk(mEi, CStr(mSdata(SC_NAME)), CStr(mSdata(SC_UNIT)), IIf(mKind = "SUP_RET", "Вернуть", "Списать"), "OUT", WmsIssue.Round3(mSbal + q1), 0, _
                mQty <= WmsIssue.Round3(mSbal + q1) + 0.0000001, mQty, info, msg)
            If st = 1 Then
                AdjustFixRow = "SKIP:" & IIf(mKind = "SUP_RET", "возврат поставщику", "списание") & " не исправлено — " & msg
                Exit Function
            ElseIf st = 2 Then
                AdjustFixRow = "ERR:" & msg
                Exit Function
            End If
        End If
        If mQty > WmsIssue.Round3(mSbal + q1) + 0.0000001 Then
            AdjustFixRow = "ERR:" & sVerb & WmsIssue.QtyText(mQty) & " нельзя: остаток " & mCanon & " без этой корректировки — " _
                & WmsIssue.QtyText(WmsIssue.Round3(mSbal + q1))
            Exit Function
        End If
        If Abs(mQty - q1) < 0.0000001 And mDate = Int(d1) And mReason = reason1 Then
            AdjustFixRow = "SKIP:ничего не изменилось"
            Exit Function
        End If
        delta2 = -mQty
        before = WmsIssue.Round3(mSbal - delta1)
        b2 = WmsIssue.Round3(before + delta2)
        PlanBegin(mKind & "_FIX", CStr(n))
        PlanCommonFields(n, r)
        PlanField("OLD_QTY", q1)
        PlanField("QTY", mQty)
        PlanField("OLD_DATE", Format(d1, "YYYY-MM-DD"))
        PlanField("OLD_REASON", reason1)
        WmsPack.PlanQtyFields()
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        If Abs(b2 - mSbal) > 0.0000001 Then PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
        PlanSetValue(SH_ADJ, n, AJ_QTY, delta2, False)
        PlanSetValue(SH_ADJUST, r, AC_QTY, mQty, False)
        PlanSetValue(SH_ADJUST, r, AC_DIFF, delta2, False)
        PlanSetValue(SH_ADJUST, r, AC_BEFORE, before, False)
        PlanSetValue(SH_ADJUST, r, AC_AFTER, b2, False)
    Case "INV_ADJ"
        mBook = rd(AJ_BOOK)
        If Not ParseFactValue(vMain, mFact, msg) Then
            AdjustFixRow = "ERR:" & msg
            Exit Function
        End If
        If VarType(rd(AJ_FACT)) = 5 Then
            If Abs(mFact - rd(AJ_FACT)) > 0.0000001 Then
                st = WmsPack.QtyAsk(mEi, CStr(mSdata(SC_NAME)), CStr(mSdata(SC_UNIT)), "Факт", "FACT", mBook, 0, True, mFact, info, msg)
                If st = 1 Then
                    AdjustFixRow = "SKIP:инвентаризация не исправлена — " & msg
                    Exit Function
                ElseIf st = 2 Then
                    AdjustFixRow = "ERR:" & msg
                    Exit Function
                End If
            End If
        End If
        ' 1.0: corrected to the book balance — the recount matched (difference 0); «Удалить» cancels the count itself
        delta2 = WmsIssue.Round3(mFact - mBook)
        If Abs(delta2) < 0.0000001 Then delta2 = 0
        before = WmsIssue.Round3(mSbal - delta1)
        b2 = WmsIssue.Round3(before + delta2)
        If b2 < -0.0000001 Then
            AdjustFixRow = "ERR:остаток " & mCanon & " — " & WmsIssue.QtyText(mSbal) & ": после исправления он стал бы отрицательным (" _
                & WmsIssue.QtyText(b2) & "); сначала удалите (сторно) последующие выдачи"
            Exit Function
        End If
        If Abs(delta2 - delta1) < 0.0000001 And mDate = Int(d1) And mReason = reason1 Then
            AdjustFixRow = "SKIP:ничего не изменилось"
            Exit Function
        End If
        mDiff = delta2
        PlanBegin("INV_ADJ_FIX", CStr(n))
        PlanCommonFields(n, r)
        PlanField("BOOK", mBook)
        PlanField("OLD_FACT", rd(AJ_FACT))
        PlanField("FACT", mFact)
        PlanField("OLD_DIFF", delta1)
        PlanField("DIFF", delta2)
        PlanField("OLD_DATE", Format(d1, "YYYY-MM-DD"))
        PlanField("OLD_REASON", reason1)
        WmsPack.PlanQtyFields()
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        If Abs(b2 - mSbal) > 0.0000001 Then PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
        PlanSetValue(SH_ADJ, n, AJ_QTY, delta2, False)
        PlanSetValue(SH_ADJ, n, AJ_FACT, mFact, False)
        PlanSetValue(SH_ADJUST, r, AC_FACT, mFact, False)
        PlanSetValue(SH_ADJUST, r, AC_DIFF, delta2, False)
        PlanSetValue(SH_ADJUST, r, AC_BEFORE, before, False)
        PlanSetValue(SH_ADJUST, r, AC_AFTER, b2, False)
    End Select
    If CLng(rd(AJ_ROW)) <> r Then PlanSetValue(SH_ADJ, n, AJ_ROW, r, False)
    If mDate <> Int(d1) Then PlanSetValue(SH_ADJ, n, AJ_DATE, mDate, False)
    PlanSetValue(SH_ADJUST, r, AC_DATE, mDate, False)
    PlanSetValue(SH_ADJUST, r, AC_REASON, mReason, False)
    PlanSetValue(SH_ADJUST, r, AC_CTL, ST_FIXED, False)
    PlanLockBits(SH_ADJUST, r, 0, AC_LAST, ADJUST_LOCKS_POSTED)
    If mKind <> "MOVE" Then WmsOrders.PlanStockMirror(mEi, b2)
    res = ApplyOperation(0, 0)
    AdjustFixRow = res
    Exit Function
EH:
    AdjustFixRow = "ERR-SYS:внутренняя ошибка проверки исправления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Удалить» — MOVE_DEL / WRITE_OFF_DEL / INV_ADJ_DEL (storno)
' The correction is reversed: a move puts the EI back to its place before the move (only while the EI is still where
' the move put it), a write-off returns its quantity, an inventory correction takes its difference back (never below 0).
' The row stays as the history with «Удалено (сторно)», its № is never reused.

Function AdjustDeleteRow(r As Long) As String
    Dim why As String, kind As String, n As Long, rd As Variant, sh As Object, delta1 As Double, b2 As Double
    Dim from1 As String, to1 As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        AdjustDeleteRow = "ERR:выберите строку корректировки (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = AdjustSheet()
    kind = AdjustRowKind(r)
    If kind = "DELETED" Then
        AdjustDeleteRow = "SKIP:корректировка № " & gAjN & " уже удалена (сторно)"
        Exit Function
    End If
    why = PostedRowProblem(kind, "Удалить")
    If why <> "" Then
        AdjustDeleteRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        AdjustDeleteRow = why
        Exit Function
    End If
    n = gAjN
    why = DeleteProblem(r, rd)
    If why <> "" Then
        AdjustDeleteRow = why
        Exit Function
    End If
    delta1 = rd(AJ_QTY)
    from1 = CStr(rd(AJ_FROM))
    to1 = CStr(rd(AJ_TO))
    mDate = rd(AJ_DATE)
    mReason = sh.getCellByPosition(AC_REASON, r).getString()
    mBatch = sh.getCellByPosition(AC_BATCH, r).getString()
    b2 = WmsIssue.Round3(mSbal - delta1)
    PlanBegin(mKind & "_DEL", CStr(n))
    PlanCommonFields(n, r)
    Select Case mKind
    Case "MOVE"
        PlanField("FROM", from1)
        PlanField("TO", to1)
        PlanSetValue(SH_STOCK, mEi, SC_PLACE, from1, False)
    Case "WRITE_OFF", "SUP_RET"
        PlanField("QTY", -delta1)
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
    Case "INV_ADJ"
        PlanField("BOOK", rd(AJ_BOOK))
        PlanField("FACT", rd(AJ_FACT))
        PlanField("DIFF", delta1)
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
    Case "SPLIT"
        ' the quantity back to the source; the new EI closed (balance 0, «Приход удалён (сторно)» — its number is not reused)
        PlanField("QTY", -delta1)
        PlanField("NEW_EI", mNewCanon)
        PlanField("BAL_BEFORE", mSbal)
        PlanField("BAL_AFTER", b2)
        PlanSetValue(SH_STOCK, mEi, SC_QTY, b2, False)
        PlanSetValue(SH_STOCK, mNewEi, SC_QTY, 0, False)
        PlanSetValue(SH_STOCK, mNewEi, SC_STATE, EI_ST_STORNO, False)
    End Select
    PlanSetValue(SH_ADJ, n, AJ_STATE, RV_STORNO, False)
    If CLng(rd(AJ_ROW)) <> r Then PlanSetValue(SH_ADJ, n, AJ_ROW, r, False)
    PlanSetValue(SH_ADJUST, r, AC_CTL, ST_DELETED, False)
    PlanLockBits(SH_ADJUST, r, 0, AC_LAST, ADJUST_LOCKS_POSTED)
    If mKind <> "MOVE" Then WmsOrders.PlanStockMirror(mEi, b2)
    AdjustDeleteRow = ApplyOperation(0, 0)
    Exit Function
EH:
    AdjustDeleteRow = "ERR-SYS:внутренняя ошибка проверки удаления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' "" when the posted correction of row r (gAjN) may be cancelled now; otherwise ERR:/ERR-SYS: and the reason. The UI asks
' this before its confirmation (a refusal is the first window).
Function DeleteProblem(r As Long, rd As Variant) As String
    Dim why As String, delta1 As Double
    why = LoadPosted(r, rd)
    If why <> "" Then
        DeleteProblem = why
        Exit Function
    End If
    If CStr(mSdata(SC_STATE)) = EI_ST_STORNO Then
        DeleteProblem = "ERR-SYS:" & mCanon & ": приход удалён (сторно), а корректировка № " & gAjN & " действует — нужна самопроверка"
        Exit Function
    End If
    delta1 = rd(AJ_QTY)
    Select Case mKind
    Case "SPLIT"
        DeleteProblem = SplitDeleteProblem(rd)
        Exit Function
    Case "MOVE"
        If LCase(Trim(mSplace)) <> LCase(Trim(CStr(rd(AJ_TO)))) Then
            DeleteProblem = "ERR:после перемещения № " & gAjN & " место " & mCanon & " менялось (сейчас «" & mSplace & "», перемещено в «" _
                & rd(AJ_TO) & "») — сторно вернуло бы ЕИ не туда; оформите новое перемещение"
            Exit Function
        End If
    Case Else
        If WmsIssue.Round3(mSbal - delta1) < -0.0000001 Then
            DeleteProblem = "ERR:остаток " & mCanon & " — " & WmsIssue.QtyText(mSbal) & ": сторно корректировки (" & IIf(delta1 > 0, "+", "") _
                & WmsIssue.QtyText(delta1) & ") сделало бы его отрицательным; сначала удалите (сторно) последующие выдачи"
            Exit Function
        End If
    End Select
    DeleteProblem = ""
End Function

' 1.0 (М1): a split is cancelled only while its new EI is as the split made it — the whole quantity, no movement, no line
' that names it; then the quantity goes back to the source EI and the new EI is closed (its number is never reused)
Private Function SplitDeleteProblem(rd As Variant) As String
    Dim m As Long, canonM As String, p As String, bm As Double, d As Variant, nIss As Long, nRet As Long, nAdj As Long, nLive As Long
    canonM = CStr(rd(AJ_NEW))
    If Not WmsOrders.StrictEI(canonM, m) Then
        SplitDeleteProblem = "ERR-SYS:в " & SH_ADJ & " у разделения № " & gAjN & " неверный новый ЕИ «" & canonM & "»"
        Exit Function
    End If
    p = WmsIssue.StockRowProblem(m, canonM)
    If p <> "" Then
        SplitDeleteProblem = "ERR:" & p
        Exit Function
    End If
    d = WmsIssue.StockData(m)
    If Not WmsIssue.StockQty(m, bm) Then
        SplitDeleteProblem = "ERR:остаток " & canonM & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If Abs(bm + rd(AJ_QTY)) > 0.0000005 Then
        SplitDeleteProblem = "ERR:у нового ЕИ " & canonM & " остаток уже " & WmsIssue.QtyText(bm) & " (при разделении " & WmsIssue.QtyText(-rd(AJ_QTY)) _
            & ") — по нему были движения; разделение не удаляется. Верните остаток обычными операциями"
        Exit Function
    End If
    If Not WmsReturn.LiveDependents(canonM, nIss, nRet) Or Not LiveAdjustments(canonM, nAdj) Or Not WmsSpecial.LiveLinesOfEI(canonM, nLive) Then
        SplitDeleteProblem = "ERR-SYS:не удалось проверить движения нового ЕИ " & canonM & " (формула движка Calc) — сторно не выполнено"
        Exit Function
    End If
    If nIss + nRet + nAdj + nLive > 0 Then
        SplitDeleteProblem = "ERR:у нового ЕИ " & canonM & " есть действующие движения (выдач " & nIss & ", возвратов " & nRet & ", корректировок " & nAdj _
            & ", строк иного прихода " & nLive & ") — сначала их сторно"
        Exit Function
    End If
    If WmsIssue.Round3(mSbal - rd(AJ_QTY)) < -0.0000001 Then
        SplitDeleteProblem = "ERR-SYS:сторно разделения № " & gAjN & " дало бы отрицательный остаток " & mCanon
        Exit Function
    End If
    mNewEi = m
    mNewCanon = canonM
    SplitDeleteProblem = ""
End Function

' ================================================================ «Проверить»

Function AdjustCheckRow(r As Long) As String
    Dim kind As String, why As String, rd As Variant, sh As Object
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        AdjustCheckRow = "ERR:выберите строку корректировки (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = AdjustSheet()
    kind = AdjustRowKind(r)
    Select Case kind
    Case "OPEN"
        mInHandler = True
        why = CheckRowInputs(r, True)
        If why <> "" Then
            SetCtl(r, "Ошибка: " & why)
            AdjustCheckRow = "ERR:" & why
        Else
            SetCtl(r, "Можно провести: " & Describe())
            AdjustCheckRow = "OK:можно провести: " & Describe()
        End If
        mInHandler = False
    Case "POSTED", "DELETED"
        why = LoadPosted(r, rd)
        If why <> "" Then
            AdjustCheckRow = "ERR:корректировка № " & gAjN & ": " & Mid(why, InStr(why, ":") + 1)
            Exit Function
        End If
        AdjustCheckRow = "OK:корректировка № " & gAjN & " (" & KindName(mKind) & ", " & IIf(kind = "DELETED", "удалена (сторно)", "действует") & "): " _
            & mCanon & IIf(mKind = "MOVE", ", место «" & rd(AJ_FROM) & "» → «" & rd(AJ_TO) & "»", ", изменение остатка " _
            & IIf(rd(AJ_QTY) > 0, "+", "") & WmsIssue.QtyText(rd(AJ_QTY))) & "; сейчас: место «" & mSplace & "», остаток " & WmsIssue.QtyText(mSbal)
    Case "COPY", "FOREIGN"
        AdjustCheckRow = "ERR:строка — КОПИЯ" & IIf(gAjWhy <> "", " (" & gAjWhy & ")", "") & ", её можно только очистить кнопкой «Очистить»"
    Case Else
        AdjustCheckRow = "ERR:строка пуста — укажите вид (B) и ЕИ (C)"
    End Select
    Exit Function
EH:
    mInHandler = False
    AdjustCheckRow = "ERR:внутренняя ошибка проверки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ dependents of an EI (D-069 for the corrections)

' The live corrections of EI canon (_ADJ: C = canon, E = LIVE): counted by one formula of the Calc engine (COUNTIFS; no
' loop over a sheet). False when the formula could not be used: the caller refuses.
Function LiveAdjustments(canon As String, ByRef nAdj As Long) As Boolean
    Dim la As Long, q As String, s As String, ok As Boolean
    nAdj = 0
    If Not gDoc.Sheets.hasByName(SH_ADJ) Then
        LiveAdjustments = True
        Exit Function
    End If
    la = WmsOrders.LastRow(AdjSheet())
    If la < 1 Then
        LiveAdjustments = True
        Exit Function
    End If
    q = """" & Replace(canon, """", """""") & """"
    s = WmsOrders.ScratchEval("COUNTIFS($'" & SH_ADJ & "'.$C$2:$C$" & (la + 1) & ";" & q & ";$'" & SH_ADJ & "'.$E$2:$E$" & (la + 1) _
        & ";""" & RV_LIVE & """)", ok)
    If Not ok Then Exit Function
    If Not IsDigits(s) Or s = "" Then Exit Function
    nAdj = CLng(s)
    LiveAdjustments = True
End Function

' ================================================================ light change handler (spec §13, D-008)

Sub OnAdjustChange(Optional oTarget As Variant)
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
    Dim r As Long, r1 As Long, n As Long, um As Object, inCtx As Boolean, last As Long
    If ra.EndRow < 1 Then Exit Sub
    ' the inputs A (a key that appeared by paste) B C G H I J K L matter; D E F M N O P Q are written by WMS, R S are free
    If ra.StartColumn > AC_REASON Then Exit Sub
    If ra.StartColumn >= AC_NAME And ra.EndColumn <= AC_UNIT Then Exit Sub
    mInHandler = True
    WmsInit()
    r1 = ra.EndRow
    If r1 - ra.StartRow + 1 > PREVIEW_MAX_ROWS Then
        last = WmsOrders.LastRow(AdjustSheet())
        If r1 > last Then r1 = last
    End If
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: проверка строки корректировки")
    inCtx = True
    For r = IIf(ra.StartRow < 1, 1, ra.StartRow) To r1
        n = n + 1
        If n > PREVIEW_MAX_ROWS Then Exit For
        PreviewAdjustRow(r)
    Next r
    um.leaveUndoContext()
    mInHandler = False
    Exit Sub
EH:
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    mInHandler = False
End Sub

' The preview of an open row: the EI in C → D E F from the registry, M (current place), N (current balance); for an
' inventory I (the book balance) when empty and P (the difference); O the balance after. «Контроль» Q: the first
' problem of the filled inputs, otherwise what the row will do. A value the user typed is never replaced.
Sub PreviewAdjustRow(r As Long)
    Dim sh As Object, kind As String, why As String
    sh = AdjustSheet()
    kind = AdjustRowKind(r)
    Select Case kind
    Case "FOREIGN"
        CopyCheckAdjust(r)
        Exit Sub
    Case "EMPTY"
        ClearDerived(r, True)
        Exit Sub
    Case "OPEN"
    Case Else
        Exit Sub
    End Select
    If IsBlank(sh.getCellByPosition(AC_EI, r)) Then
        ClearDerived(r, False)
        SetCtl(r, IIf(KindCode(sh.getCellByPosition(AC_KIND, r).getString()) = "" And Not IsBlank(sh.getCellByPosition(AC_KIND, r)), _
            "Ошибка: вид (B) — Перемещение, Списание, Инвентаризация или Разделение", "Укажите вид (B) и ЕИ (C)"))
        Exit Sub
    End If
    why = LoadEI(sh.getCellByPosition(AC_EI, r).getString())
    If why <> "" Then
        ClearDerived(r, False)
        SetCtl(r, "Ошибка: " & why)
        Exit Sub
    End If
    WmsOrders.SetIfDiff(sh.getCellByPosition(AC_EI, r), mCanon)
    WmsOrders.SetIfDiff(sh.getCellByPosition(AC_NAME, r), mSdata(SC_NAME))
    WmsOrders.SetIfDiff(sh.getCellByPosition(AC_ART, r), mSdata(SC_ART))
    WmsOrders.SetIfDiff(sh.getCellByPosition(AC_UNIT, r), mSdata(SC_UNIT))
    WmsOrders.SetIfDiff(sh.getCellByPosition(AC_FROM, r), mSplace)
    WmsOrders.SetIfDiff(sh.getCellByPosition(AC_BEFORE, r), mSbal)
    ' I «Учётный остаток» belongs to an inventory: shown for it when empty, removed for another kind
    If KindCode(sh.getCellByPosition(AC_KIND, r).getString()) = "INV_ADJ" Then
        If IsBlank(sh.getCellByPosition(AC_BOOK, r)) Then sh.getCellByPosition(AC_BOOK, r).setValue(mSbal)
    ElseIf sh.getCellByPosition(AC_BOOK, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        ClearCell(sh.getCellByPosition(AC_BOOK, r))
    End If
    why = CheckRowInputs(r, False)
    If why <> "" Then
        If sh.getCellByPosition(AC_AFTER, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(AC_AFTER, r))
        If sh.getCellByPosition(AC_DIFF, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(AC_DIFF, r))
        SetCtl(r, "Ошибка: " & why)
        Exit Sub
    End If
    Select Case mKind
    Case "WRITE_OFF", "SUP_RET"
        If IsBlank(sh.getCellByPosition(AC_QTY, r)) Then
            SetCtl(r, IIf(mKind = "SUP_RET", "Укажите, сколько вернуть поставщику (G)", "Укажите количество списания (G)") & "; остаток " _
                & WmsPack.QtyWithParts(mEi, mSbal, CStr(mSdata(SC_UNIT))))
            Exit Sub
        End If
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_DIFF, r), -mQty)
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_AFTER, r), WmsIssue.Round3(mSbal - mQty))
    Case "SPLIT"
        If IsBlank(sh.getCellByPosition(AC_QTY, r)) Then
            SetCtl(r, "Укажите, сколько отделить в новый ЕИ (G); остаток " & WmsPack.QtyWithParts(mEi, mSbal, CStr(mSdata(SC_UNIT))))
            Exit Sub
        End If
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_DIFF, r), -mQty)
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_AFTER, r), WmsIssue.Round3(mSbal - mQty))
    Case "INV_ADJ"
        If IsBlank(sh.getCellByPosition(AC_FACT, r)) Then
            SetCtl(r, "Укажите фактический остаток (H); учётный " & WmsIssue.QtyText(mSbal))
            Exit Sub
        End If
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_DIFF, r), mDiff)
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_AFTER, r), mFact)
    Case "MOVE"
        If mPlace = "" Then
            SetCtl(r, "Укажите новое место (J); сейчас «" & mSplace & "»")
            Exit Sub
        End If
        WmsOrders.SetIfDiff(sh.getCellByPosition(AC_AFTER, r), mSbal)
    End Select
    If IsBlank(sh.getCellByPosition(AC_DATE, r)) Or IsBlank(sh.getCellByPosition(AC_REASON, r)) Then
        SetCtl(r, Describe() & " · укажите " & IIf(IsBlank(sh.getCellByPosition(AC_DATE, r)), "дату (K)" & IIf(IsBlank(sh.getCellByPosition(AC_REASON, r)), " и ", ""), "") _
            & IIf(IsBlank(sh.getCellByPosition(AC_REASON, r)), IIf(mKind = "INV_ADJ", "основание (L)", "причину (L)"), ""))
    Else
        SetCtl(r, "Можно провести: " & Describe())
    End If
End Sub

' the WMS-filled cells of an open row (D E F M N O P, and Q when bAll); I only when WMS put the book balance there
Private Sub ClearDerived(r As Long, bAll As Boolean)
    Dim sh As Object, cols As Variant, i As Integer
    sh = AdjustSheet()
    cols = Array(AC_NAME, AC_ART, AC_UNIT, AC_FROM, AC_BEFORE, AC_AFTER, AC_DIFF)
    For i = 0 To UBound(cols)
        If sh.getCellByPosition(cols(i), r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(cols(i), r))
    Next i
    If bAll Then
        If sh.getCellByPosition(AC_CTL, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(AC_CTL, r))
    End If
End Sub

Sub HandlerOff(b As Boolean)
    mInHandler = b
End Sub

' ================================================================ «Очистить» — an unposted row or a copy (no accounting change)

Function AdjustClearRow(r As Long) As String
    Dim sh As Object, kind As String, wasProt As Boolean, i As Integer, pr As Object, x As Double, d As Variant
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        AdjustClearRow = "ERR:выберите строку корректировки (не заголовок)"
        Exit Function
    End If
    If gBusy Then
        AdjustClearRow = "BUSY:операция уже выполняется"
        Exit Function
    End If
    sh = AdjustSheet()
    kind = AdjustRowKind(r)
    Select Case kind
    Case "POSTED"
        AdjustClearRow = "ERR:строка — проведённая корректировка № " & gAjN & " — для отмены используйте «Удалить»"
        Exit Function
    Case "DELETED"
        AdjustClearRow = "ERR:корректировка № " & gAjN & " удалена (сторно) — строка хранит историю и не очищается"
        Exit Function
    Case "COPY"
        If sh.getCellByPosition(AC_NO, r).getType() = com.sun.star.table.CellContentType.VALUE Then
            x = sh.getCellByPosition(AC_NO, r).getValue()
            If x >= 1 And x = Int(x) And x <= MAX_SHEET_ROW Then
                If AdjRegistered(CLng(x), d) Then
                    If VarType(d(AJ_ROW)) = 5 Then
                        If CLng(d(AJ_ROW)) = r Then
                            AdjustClearRow = "ERR:строка — зарегистрированная корректировка № " & CLng(x) & "; копией является другая строка с этим №"
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End Select
    mInHandler = True
    On Error GoTo EH
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    sh.getCellRangeByPosition(0, r, AC_LAST, r).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME _
        + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA)
    For i = 0 To AC_LAST
        pr = sh.getCellByPosition(i, r).CellProtection
        pr.IsLocked = (Mid(ADJUST_LOCKS_OPEN, i + 1, 1) = "1")
        sh.getCellByPosition(i, r).CellProtection = pr
    Next i
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
    AdjustClearRow = "OK:" & IIf(kind = "COPY" Or kind = "FOREIGN", "копия очищена", "строка очищена")
    Exit Function
EH:
    AdjustClearRow = "ERR:не удалось очистить строку: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
End Function

' ================================================================ «Главная»: corrections not posted yet (spec §23)

' rows with a kind (B) but no № (A): two Calc queries per block instead of a row loop. -1 when it fails.
' the live positive inventory corrections of every EI (INV_ADJ with a difference > 0), index = EI number: a surplus found
' by an inventory may raise a balance above what was received — the self-checks of the receipts allow exactly this much
Function PlusByEI(nextEI As Long) As Variant
    Dim out() As Double, last As Long, d As Variant, i As Long, n As Long
    ReDim out(nextEI)
    last = CLng(SysNum(SK_NEXT_ADJ)) - 1
    If last >= 1 Then
        d = AdjSheet().getCellRangeByPosition(0, 1, AJ_LAST, last).getDataArray()
        For i = 0 To UBound(d)
            If CStr(d(i)(AJ_KIND)) = "INV_ADJ" And CStr(d(i)(AJ_STATE)) = RV_LIVE And VarType(d(i)(AJ_QTY)) = 5 Then
                If d(i)(AJ_QTY) > 0 And WmsOrders.StrictEI(CStr(d(i)(AJ_EI)), n) Then
                    If n >= 1 And n <= nextEI Then out(n) = out(n) + d(i)(AJ_QTY)
                End If
            End If
        Next i
    End If
    PlusByEI = out
End Function

Function UnpostedAdjustments() As Long
    Dim sh As Object, last As Long, blocks As Variant, got As Variant, i As Long, j As Long, n As Long, flags As Long
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_ADJUST) Then Exit Function
    sh = AdjustSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    blocks = sh.getCellRangeByPosition(AC_NO, 1, AC_NO, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= UNPOSTED_MAX_BLOCKS Then Exit For
        got = sh.getCellRangeByPosition(AC_KIND, blocks(i).StartRow, AC_KIND, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            n = n + got(j).EndRow - got(j).StartRow + 1
        Next j
    Next i
    UnpostedAdjustments = n
    Exit Function
EH:
    UnpostedAdjustments = -1
End Function

' ================================================================ self-check (a part of «WMS-Доктор»)
' Every correction of _ADJ (№, kind, EI, change, state), NEXT_ADJ above every №, the row hints; the EIs named exist in
' the registry. Chunks, only the needed columns.

Function AdjustCheck() As String
    Dim ad As Object, sh As Object, last As Long, r0 As Long, r1 As Long, i As Long, d As Variant, t0 As Long, nextAdj As Long
    Dim n As Long, nBad As Long, first As String, nStale As Long, nLive As Long, nStorno As Long, maxN As Long, x As Double, lastR As Long
    Dim colA As Variant, colA0 As Long, k As String, st As String, ei As Long, nKind(4) As Long
    On Error GoTo EH
    t0 = GetSystemTicks()
    ad = AdjSheet()
    sh = AdjustSheet()
    nextAdj = CLng(SysNum(SK_NEXT_ADJ))
    lastR = WmsOrders.LastRow(sh)
    colA0 = -1
    last = WmsOrders.LastRow(ad)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = ad.getCellRangeByPosition(0, r0, AJ_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(d(i)(AJ_NO)) <> "" Then
                If VarType(d(i)(AJ_NO)) <> 5 Then
                    Bad(nBad, first, SH_ADJ & " строка " & (r0 + i + 1) & ": № не число")
                ElseIf d(i)(AJ_NO) <> r0 + i Then
                    Bad(nBad, first, SH_ADJ & " строка " & (r0 + i + 1) & ": записан № " & d(i)(AJ_NO) & " (плотная адресация нарушена)")
                Else
                    n = n + 1
                    If r0 + i > maxN Then maxN = r0 + i
                    k = CStr(d(i)(AJ_KIND))
                    st = CStr(d(i)(AJ_STATE))
                    If k = "MOVE" Then
                        nKind(0) = nKind(0) + 1
                    ElseIf k = "WRITE_OFF" Then
                        nKind(1) = nKind(1) + 1
                    ElseIf k = "INV_ADJ" Then
                        nKind(2) = nKind(2) + 1
                    ElseIf k = "SUP_RET" Then
                        nKind(4) = nKind(4) + 1
                    ElseIf k = "SPLIT" Then
                        nKind(3) = nKind(3) + 1
                        If Not WmsOrders.StrictEI(CStr(d(i)(AJ_NEW)), ei) Then
                            Bad(nBad, first, "разделение № " & (r0 + i) & ": неверный новый ЕИ «" & d(i)(AJ_NEW) & "»")
                        ElseIf ei >= SysNum(SK_NEXT_EI) Then
                            Bad(nBad, first, "разделение № " & (r0 + i) & ": новый ЕИ " & d(i)(AJ_NEW) & " не выдавался")
                        End If
                    Else
                        Bad(nBad, first, "корректировка № " & (r0 + i) & ": неизвестный вид «" & k & "»")
                    End If
                    If Not WmsOrders.StrictEI(CStr(d(i)(AJ_EI)), ei) Then
                        Bad(nBad, first, "корректировка № " & (r0 + i) & ": неверный ЕИ «" & d(i)(AJ_EI) & "»")
                    ElseIf ei >= SysNum(SK_NEXT_EI) Then
                        Bad(nBad, first, "корректировка № " & (r0 + i) & ": ЕИ " & d(i)(AJ_EI) & " не выдавался")
                    End If
                    If VarType(d(i)(AJ_QTY)) <> 5 Then
                        Bad(nBad, first, "корректировка № " & (r0 + i) & ": изменение остатка не число")
                    ElseIf k = "MOVE" And d(i)(AJ_QTY) <> 0 Then
                        Bad(nBad, first, "перемещение № " & (r0 + i) & ": изменение остатка " & d(i)(AJ_QTY) & " (должно быть 0)")
                    ElseIf (k = "WRITE_OFF" Or k = "SPLIT" Or k = "SUP_RET") And d(i)(AJ_QTY) >= 0 Then
                        Bad(nBad, first, KindName(k) & " № " & (r0 + i) & ": изменение остатка " & d(i)(AJ_QTY) & " (должно быть меньше 0)")
                    End If
                    If st = RV_LIVE Then
                        nLive = nLive + 1
                    ElseIf st = RV_STORNO Then
                        nStorno = nStorno + 1
                    Else
                        Bad(nBad, first, "корректировка № " & (r0 + i) & ": неизвестное состояние «" & st & "»")
                    End If
                    x = -1
                    If VarType(d(i)(AJ_ROW)) = 5 Then x = d(i)(AJ_ROW)
                    If x < 1 Or x > lastR Then
                        nStale = nStale + 1
                    Else
                        If CLng(x) < colA0 Or CLng(x) >= colA0 + CHECK_CHUNK_ROWS Or colA0 < 0 Then
                            colA0 = CLng(x)
                            colA = sh.getCellRangeByPosition(AC_NO, colA0, AC_NO, IIf(colA0 + CHECK_CHUNK_ROWS - 1 > lastR, lastR, colA0 + CHECK_CHUNK_ROWS - 1)).getDataArray()
                        End If
                        If VarType(colA(CLng(x) - colA0)(0)) <> 5 Then
                            nStale = nStale + 1
                        ElseIf colA(CLng(x) - colA0)(0) <> r0 + i Then
                            nStale = nStale + 1
                        End If
                    End If
                End If
            End If
        Next i
    Next r0
    If maxN >= nextAdj Then Bad(nBad, first, "NEXT_ADJ " & nextAdj & " не больше максимального № корректировки " & maxN)
    If nBad > 0 Then
        AdjustCheck = "ОШИБКА: расхождений " & nBad & " (" & first & ")"
    Else
        AdjustCheck = "корректировок " & n & " (перемещений " & nKind(0) & ", списаний " & nKind(1) & ", инвентаризаций " & nKind(2) & IIf(nKind(3) > 0, ", разделений " & nKind(3), "") & IIf(nKind(4) > 0, ", возвратов поставщику " & nKind(4), "") & "; из них сторно " _
            & nStorno & "), расхождений нет, NEXT_ADJ " & nextAdj & " выше всех №" _
            & IIf(nStale > 0, ", устаревших подсказок строк " & nStale & " (исправятся при следующей операции)", "") & ", " & (GetSystemTicks() - t0) & " мс"
    End If
    Exit Function
EH:
    AdjustCheck = "ОШИБКА: проверка корректировок не выполнена: " & Error$ & " (строка " & Erl & ")"
End Function

Private Sub Bad(ByRef nBad As Long, ByRef first As String, s As String)
    nBad = nBad + 1
    If first = "" Then first = s
End Sub
