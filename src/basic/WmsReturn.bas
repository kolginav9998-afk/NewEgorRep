' WmsReturn — возврат выданного товара (MASTER SPEC v0.3 §11, §14–§16, §18, §25; D-004, D-011, D-033; задание Core Phase 4).
' Возврат всегда ссылается на конкретную проведённую выдачу (№ выдачи Phase 2 — стабильный ключ) и возвращает количество
' в остаток того же ЕИ. Каждый возврат, его исправление (RETURN_FIX) и сторно (RETURN_DEL) — одна операция
' WmsCore.ApplyOperation: номер возврата, остаток и место ЕИ в «Наличие», запись _RET, итоги выдачи _ISS, строка «Возврат»
' и «Заказы».X меняются вместе или не меняются вовсе.
'
' Служебные данные (лист «Возврат» — только пользовательская таблица A:O):
'  - _RET — плотно, строка = № возврата: № выдачи, ЕИ, количество, состояние LIVE / STORNO, подсказка строки «Возврат»;
'  - _ISS — плотно, строка = № выдачи: подсказка строки «Выдачи», возвращено (сумма действующих возвратов), их число.
' Номер строки — только подсказка: перед использованием строка проверяется ключом (№), при несовпадении ищется одной
' формулой MATCH движка Calc (_IDX), и подсказку исправляет та же операция.
' «Выдачи».M/N (Вернул, Вернул в %) — legacy/информационные (D-004): возвраты туда не пишутся и из них не считаются.
Option Explicit

' details of the last ReturnRowKind: return № of the row, the row hint was outdated (rows were inserted above), why FOREIGN
Global gRtN As Long
Global gRtMoved As Boolean
Global gRtWhy As String
' the EI and the recipient of the issue the last IssueForDialog described (for the window of «Найти выдачу»)
Global gRtIssueEI As String
Global gRtIssueWho As String

Private mInHandler As Boolean

' the issue a return refers to (CheckIssue)
Private mIk As Long
Private mIrow As Long
Private mImoved As Boolean
Private mIqty As Double
Private mIunit As String
Private mIdate As Double
Private mIwho As String
Private mIei As Long
Private mIcanon As String
Private mIret As Double
Private mIcnt As Long
' the registry row of its EI
Private mSdata As Variant
Private mSbal As Double
' the return being posted (CheckRowInputs)
Private mRqty As Double
Private mRdate As Double
Private mRplace As String

' ================================================================ sheets

Function ReturnsSheet() As Object
    ReturnsSheet = gDoc.Sheets.getByName(SH_RETURNS)
End Function

Function RetSheet() As Object
    RetSheet = gDoc.Sheets.getByName(SH_RET)
End Function

Function IssSheet() As Object
    IssSheet = gDoc.Sheets.getByName(SH_ISS)
End Function

' "" when the sheets of the returns exist and have the expected layout (checked at every start, fail-closed)
Function SheetsProblem() As String
    Dim names As Variant, i As Integer, sh As Object
    names = Array(SH_RETURNS, SH_RET, SH_ISS)
    For i = 0 To UBound(names)
        If Not gDoc.Sheets.hasByName(names(i)) Then
            SheetsProblem = "нет листа «" & names(i) & "»"
            Exit Function
        End If
    Next i
    sh = ReturnsSheet()
    If sh.getCellByPosition(RC_NO, 0).getString() <> "№ возврата" Or sh.getCellByPosition(RC_ISSUE, 0).getString() <> "№ выдачи" _
        Or sh.getCellByPosition(RC_CTL, 0).getString() <> "Контроль" Or sh.getCellByPosition(RC_NOTE, 0).getString() <> "Комментарий" Then
        SheetsProblem = "лист «" & SH_RETURNS & "»: заголовок не совпадает с A:O"
        Exit Function
    End If
    If RetSheet().getCellByPosition(RT_NO, 0).getString() <> "№ возврата" Or RetSheet().getCellByPosition(RT_ROW, 0).getString() <> "Строка (индекс)" Then
        SheetsProblem = "служебный лист " & SH_RET & " повреждён (заголовок)"
        Exit Function
    End If
    If IssSheet().getCellByPosition(IS_NO, 0).getString() <> "№ выдачи" Or IssSheet().getCellByPosition(IS_CNT, 0).getString() <> "Возвратов" Then
        SheetsProblem = "служебный лист " & SH_ISS & " повреждён (заголовок)"
        Exit Function
    End If
    SheetsProblem = ""
End Function

' ================================================================ service rows (dense addressing, the key is checked)

' A..F of _RET row n (the return № n)
Function RetRow(n As Long) As Variant
    RetRow = RetSheet().getCellRangeByPosition(0, n, RT_LAST, n).getDataArray()(0)
End Function

' True when return № n is registered in _RET (its row holds n)
Function RetRegistered(n As Long, d As Variant) As Boolean
    If n < 1 Or n > MAX_SHEET_ROW Then Exit Function
    d = RetRow(n)
    If VarType(d(RT_NO)) <> 5 Then Exit Function
    RetRegistered = (d(RT_NO) = n)
End Function

' A..D of _ISS row k (the returns of issue № k)
Function IssRow(k As Long) As Variant
    IssRow = IssSheet().getCellRangeByPosition(0, k, IS_LAST, k).getDataArray()(0)
End Function

Function IssRegistered(k As Long, d As Variant) As Boolean
    If k < 1 Or k > MAX_SHEET_ROW Then Exit Function
    d = IssRow(k)
    If VarType(d(IS_NO)) <> 5 Then Exit Function
    IssRegistered = (d(IS_NO) = k)
End Function

' the live returns of issue № k: their sum and number; True when there is at least one (for ISSUE_FIX / ISSUE_DEL)
Function IssueReturns(k As Long, ByRef q As Double, ByRef cnt As Long) As Boolean
    Dim d As Variant
    q = 0
    cnt = 0
    If Not gDoc.Sheets.hasByName(SH_ISS) Then Exit Function
    If Not IssRegistered(k, d) Then Exit Function
    If VarType(d(IS_RET)) = 5 Then q = d(IS_RET)
    If VarType(d(IS_CNT)) = 5 Then cnt = CLng(d(IS_CNT))
    IssueReturns = (cnt > 0)
End Function

' ================================================================ the issue a return refers to

' 0-based «Выдачи» row of issue № k: the hint of _ISS checked by the key, otherwise the first row holding k in A that is
' not marked КОПИЯ (MATCH of the Calc engine); moved = a hint existed and was outdated; -1 when not found
Function FindIssueRow(k As Long, ByRef moved As Boolean) As Long
    Dim sh As Object, d As Variant, hint As Long, c As Object, r As Long, start As Long, i As Integer
    sh = WmsIssue.IssuesSheet()
    moved = False
    FindIssueRow = -1
    hint = -1
    If IssRegistered(k, d) Then
        If VarType(d(IS_ROW)) = 5 Then hint = CLng(d(IS_ROW))
    End If
    If hint >= 1 And hint <= MAX_SHEET_ROW Then
        c = sh.getCellByPosition(IC_NO, hint)
        If c.getType() = com.sun.star.table.CellContentType.VALUE Then
            If c.getValue() = k And Left(sh.getCellByPosition(IC_CTL, hint).getString(), 5) <> "КОПИЯ" Then
                FindIssueRow = hint
                Exit Function
            End If
        End If
    End If
    start = 0
    For i = 1 To 20
        r = WmsOrders.FindIssueNoRow(k, start)
        If r < 1 Then Exit Function
        If Left(sh.getCellByPosition(IC_CTL, r).getString(), 5) <> "КОПИЯ" Then
            moved = (hint >= 0)
            FindIssueRow = r
            Exit Function
        End If
        start = r
    Next i
End Function

' "" when issue № k is a posted issue whose EI is in the registry; fills mI* (issue), mS* (registry row of its EI)
Private Function CheckIssue(k As Long) As String
    Dim sh As Object, kind As String, d As Variant, di As Variant, n As Long, p As String
    CheckIssue = ""
    mIk = k
    If k < 1 Or k >= SysNum(SK_NEXT_NO) Then
        CheckIssue = "выдачи № " & k & " нет — WMS её не проводила"
        Exit Function
    End If
    mIrow = FindIssueRow(k, mImoved)
    If mIrow < 1 Then
        CheckIssue = "выдача № " & k & " не найдена на листе «" & SH_ISSUES & "»"
        Exit Function
    End If
    kind = WmsIssue.RowKind(mIrow)
    If kind = "DELETED" Then
        CheckIssue = "выдача № " & k & " удалена (сторно) — возврат по ней невозможен"
        Exit Function
    End If
    If kind <> "POSTED" Then
        CheckIssue = "строка выдачи № " & k & " не является проведённой выдачей"
        Exit Function
    End If
    sh = WmsIssue.IssuesSheet()
    d = sh.getCellRangeByPosition(0, mIrow, IC_LAST, mIrow).getDataArray()(0)
    If Not WmsOrders.StrictEI(CStr(d(IC_EI)), n) Then
        CheckIssue = "в строке выдачи № " & k & " нет правильного ЕИ"
        Exit Function
    End If
    mIei = n
    mIcanon = CStr(d(IC_EI))
    If VarType(d(IC_QTY)) <> 5 Or VarType(d(IC_DATE)) <> 5 Then
        CheckIssue = "в строке выдачи № " & k & " количество или дата не число — строка повреждена"
        Exit Function
    End If
    mIqty = d(IC_QTY)
    mIdate = Int(d(IC_DATE))
    mIunit = CStr(d(IC_UNIT))
    mIwho = CStr(d(IC_WHO))
    mIret = 0
    mIcnt = 0
    If IssRegistered(k, di) Then
        If VarType(di(IS_RET)) = 5 Then mIret = di(IS_RET)
        If VarType(di(IS_CNT)) = 5 Then mIcnt = CLng(di(IS_CNT))
    End If
    p = WmsIssue.StockRowProblem(mIei, mIcanon)
    If p <> "" Then
        CheckIssue = p
        Exit Function
    End If
    mSdata = WmsIssue.StockData(mIei)
    If Not WmsIssue.StockQty(mIei, mSbal) Then
        CheckIssue = "остаток " & mIcanon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If CStr(mSdata(SC_UNIT)) <> mIunit Then
        CheckIssue = "единица " & mIcanon & " в «" & SH_STOCK & "» («" & mSdata(SC_UNIT) & "») не совпадает с единицей выдачи («" & mIunit _
            & "») — пересчёт единиц не выполняется"
        Exit Function
    End If
End Function

' what can still be returned by issue mIk (after CheckIssue): issued − live returns (exceptQty: a return being corrected)
Private Function Available(exceptQty As Double) As Double
    Available = WmsIssue.Round3(mIqty - mIret + exceptQty)
End Function

' «Выдача № 12 от 20.09.2026, Иванов: выдано 10 шт, возвращено 3, можно вернуть 7» (after CheckIssue)
Private Function IssueInfo() As String
    Dim u As String
    u = IIf(mIunit <> "", " " & mIunit, "")
    IssueInfo = "выдача № " & mIk & " от " & Format(mIdate, "DD.MM.YYYY") & ", «" & mIwho & "»: выдано " & WmsIssue.QtyText(mIqty) & u _
        & ", возвращено " & WmsIssue.QtyText(mIret) & ", можно вернуть " & WmsIssue.QtyText(Available(0))
End Function

' after CheckRowInputs: the place change the return will make, for «Контроль» ("" when the place stays) — a place left
' in J from another issue must never move an EI silently
Private Function PlaceNote() As String
    If mRplace <> CStr(mSdata(SC_PLACE)) Then PlaceNote = "; место ЕИ изменится: «" & mSdata(SC_PLACE) & "» → «" & mRplace & "»"
End Function

' issue № typed into B: a whole number (number or text of digits); 0 when empty, -1 when not a number
Function IssueNoOf(c As Object) As Long
    Dim x As Double, s As String
    Select Case c.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        IssueNoOf = 0
    Case com.sun.star.table.CellContentType.VALUE
        x = c.getValue()
        If x >= 1 And x = Int(x) And x <= 2000000000 Then IssueNoOf = CLng(x) Else IssueNoOf = -1
    Case Else
        s = Trim(c.getString())
        If Left(s, 1) = "№" Then s = Trim(Mid(s, 2))
        If s <> "" And IsDigits(s) And Len(s) <= 9 Then IssueNoOf = CLng(s) Else IssueNoOf = -1
    End Select
End Function

' ================================================================ row model of «Возврат»

' Kind of row r: "EMPTY", "OPEN" (not posted), "POSTED", "DELETED", "COPY" (marked), "FOREIGN" (a № that WMS did not
' register at this row: never issued, not a return of this WMS, or a copy of a registered row). gRtN, gRtMoved, gRtWhy.
Function ReturnRowKind(r As Long) As String
    Dim sh As Object, d As Variant, x As Double, n As Long, rd As Variant, hint As Long, c As Object, i As Integer
    sh = ReturnsSheet()
    gRtN = 0
    gRtMoved = False
    gRtWhy = ""
    d = sh.getCellRangeByPosition(0, r, RC_LAST, r).getDataArray()(0)
    If Left(CStr(d(RC_CTL)), 5) = "КОПИЯ" Then
        ReturnRowKind = "COPY"
        Exit Function
    End If
    If CStr(d(RC_NO)) = "" Then
        ReturnRowKind = "EMPTY"
        For i = RC_ISSUE To RC_PLACE
            If i = RC_ISSUE Or i = RC_EI Or i = RC_QTY Or i = RC_DATE Or i = RC_WHO Or i = RC_PLACE Then
                If CStr(d(i)) <> "" Then
                    ReturnRowKind = "OPEN"
                    Exit For
                End If
            End If
        Next i
        Exit Function
    End If
    If VarType(d(RC_NO)) <> 5 Then
        gRtWhy = "№ «" & d(RC_NO) & "» не выдавался WMS"
        ReturnRowKind = "FOREIGN"
        Exit Function
    End If
    x = d(RC_NO)
    If x < 1 Or x <> Int(x) Or x > MAX_SHEET_ROW Then
        gRtWhy = "неверный № «" & d(RC_NO) & "»"
        ReturnRowKind = "FOREIGN"
        Exit Function
    End If
    n = CLng(x)
    If n >= SysNum(SK_NEXT_RET) Then
        gRtWhy = "№ возврата " & n & " WMS не выдавала"
        ReturnRowKind = "FOREIGN"
        Exit Function
    End If
    If Not RetRegistered(n, rd) Then
        gRtWhy = "возврат № " & n & " не зарегистрирован WMS"
        ReturnRowKind = "FOREIGN"
        Exit Function
    End If
    If VarType(rd(RT_ROW)) = 5 Then hint = CLng(rd(RT_ROW)) Else hint = -1
    If hint <> r Then
        ' the registered row is elsewhere: when it still holds the key this row is a copy, otherwise rows moved
        If hint >= 1 And hint <= MAX_SHEET_ROW Then
            c = sh.getCellByPosition(RC_NO, hint)
            If c.getType() = com.sun.star.table.CellContentType.VALUE Then
                If c.getValue() = n Then
                    gRtWhy = "возврат № " & n & " повторяет строку " & (hint + 1)
                    ReturnRowKind = "FOREIGN"
                    Exit Function
                End If
            End If
        End If
        If WmsOrders.CountReturnNo(n) > 1 Then
            If FirstReturnRow(n) <> r Then
                gRtWhy = "возврат № " & n & " повторяется"
                ReturnRowKind = "FOREIGN"
                Exit Function
            End If
        End If
        gRtMoved = True
    End If
    gRtN = n
    If CStr(rd(RT_STATE)) = RV_STORNO Then ReturnRowKind = "DELETED" Else ReturnRowKind = "POSTED"
End Function

' the first row holding return № n in A that is not marked КОПИЯ; -1 when none
Function FirstReturnRow(n As Long) As Long
    Dim r As Long, start As Long, i As Integer, sh As Object
    sh = ReturnsSheet()
    FirstReturnRow = -1
    For i = 1 To 20
        r = WmsOrders.FindReturnNoRow(n, start)
        If r < 1 Then Exit Function
        If Left(sh.getCellByPosition(RC_CTL, r).getString(), 5) <> "КОПИЯ" Then
            FirstReturnRow = r
            Exit Function
        End If
        start = r
    Next i
End Function

' 0-based «Возврат» row of return № n: the hint of _RET checked by the key, otherwise found by MATCH; -1 when none
Function ReturnRowOf(n As Long, hint As Variant) As Long
    Dim sh As Object, c As Object, h As Long
    sh = ReturnsSheet()
    If VarType(hint) = 5 Then h = CLng(hint) Else h = -1
    If h >= 1 And h <= MAX_SHEET_ROW Then
        c = sh.getCellByPosition(RC_NO, h)
        If c.getType() = com.sun.star.table.CellContentType.VALUE Then
            If c.getValue() = n Then
                ReturnRowOf = h
                Exit Function
            End If
        End If
    End If
    ReturnRowOf = FirstReturnRow(n)
End Function

' a № in A that WMS did not register at this row marks the row КОПИЯ (spec §15)
Sub CopyCheckReturn(r As Long)
    If ReturnRowKind(r) <> "FOREIGN" Then Exit Sub
    ReturnsSheet().getCellByPosition(RC_CTL, r).setString("КОПИЯ: " & gRtWhy & " — не проводится, очистите строку кнопкой «Очистить»")
End Sub

' writes the text of «Контроль» (N) of an unposted row
Private Sub SetCtl(r As Long, s As String)
    WmsOrders.SetIfDiff(ReturnsSheet().getCellByPosition(RC_CTL, r), s)
End Sub

' ================================================================ checks of a row before posting (shared with the preview)

' Validates the inputs of row r against the issue in B: "" or the reason; fills mI*, mS*, mR*. bNeedAll: posting needs
' the quantity and the date (the preview checks only what is filled).
Private Function CheckRowInputs(r As Long, bNeedAll As Boolean) As String
    Dim sh As Object, k As Long, why As String, n As Long, canon As String, msg As String, c As Object, who As String, full As String
    Dim note As String, st As Integer, av As Double
    sh = ReturnsSheet()
    k = IssueNoOf(sh.getCellByPosition(RC_ISSUE, r))
    If k = 0 Then
        CheckRowInputs = "не указан № выдачи (B) — введите его или нажмите «Найти выдачу»"
        Exit Function
    End If
    If k < 0 Then
        CheckRowInputs = "№ выдачи (B) «" & sh.getCellByPosition(RC_ISSUE, r).getString() & "» — не номер"
        Exit Function
    End If
    why = CheckIssue(k)
    If why <> "" Then
        CheckRowInputs = why
        Exit Function
    End If
    c = sh.getCellByPosition(RC_EI, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not WmsIssue.NormalizeEI(c.getString(), n, canon, msg) Then
            CheckRowInputs = msg
            Exit Function
        End If
        If n <> mIei Then
            CheckRowInputs = "ЕИ " & canon & " не совпадает с ЕИ выдачи № " & k & " (" & mIcanon & ") — возврат относится к ЕИ исходной выдачи"
            Exit Function
        End If
    End If
    who = Trim(sh.getCellByPosition(RC_WHO, r).getString())
    If who <> "" Then
        st = WmsIssue.ResolveRecipient(who, full, note)
        If st = 3 Then
            CheckRowInputs = note
            Exit Function
        End If
        If LCase(Trim(full)) <> LCase(Trim(mIwho)) Then
            CheckRowInputs = "возвращает «" & full & "», а выдача № " & k & " была «" & mIwho & "» — возврат от другого получателя " _
                & "не поддерживается (это будущий сценарий передачи)"
            Exit Function
        End If
    End If
    av = Available(0)
    c = sh.getCellByPosition(RC_QTY, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Or bNeedAll Then
        If Not WmsOrders.ParseQtyAs(c, "количество возврата (F)", mRqty, msg) Then
            CheckRowInputs = msg
            Exit Function
        End If
        If av <= 0.0000001 Then
            CheckRowInputs = "по выдаче № " & k & " всё уже возвращено (выдано " & WmsIssue.QtyText(mIqty) & ", возвращено " & WmsIssue.QtyText(mIret) & ")"
            Exit Function
        End If
        If mRqty > av + 0.0000001 Then
            CheckRowInputs = "больше, чем можно вернуть: " & IssueInfo()
            Exit Function
        End If
    End If
    c = sh.getCellByPosition(RC_DATE, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Or bNeedAll Then
        If Not WmsOrders.ParseDateCellAs(c, "дата возврата (H)", mRdate, msg) Then
            CheckRowInputs = msg
            Exit Function
        End If
        If mRdate < mIdate Then
            CheckRowInputs = "дата возврата " & Format(mRdate, "DD.MM.YYYY") & " раньше даты выдачи № " & k & " (" & Format(mIdate, "DD.MM.YYYY") & ")"
            Exit Function
        End If
    End If
    mRplace = Trim(sh.getCellByPosition(RC_PLACE, r).getString())
    If mRplace = "" Then mRplace = CStr(mSdata(SC_PLACE))
    CheckRowInputs = ""
End Function

' "" when return № n may be issued now: the dense _RET is behind the counter and no row of «Возврат» holds n
Private Function NewRetProblem(n As Long) As String
    If n < 1 Or n >= MAX_SHEET_ROW Then
        NewRetProblem = "номера возвратов исчерпаны (NEXT_RET " & n & ")"
    ElseIf WmsOrders.LastRow(RetSheet()) >= n Then
        NewRetProblem = "служебная таблица возвратов " & SH_RET & " не согласована со счётчиком NEXT_RET " & n & " — нужна самопроверка"
    ElseIf WmsOrders.CountReturnNo(n) > 0 Then
        NewRetProblem = "№ возврата " & n & " уже записан на листе «" & SH_RETURNS & "» — новый номер не выдаётся. Нужна проверка ответственным"
    Else
        NewRetProblem = ""
    End If
End Function

' ================================================================ plan helpers (one operation writes all of these)

Private Sub PlanRetRow(n As Long, k As Long, canon As String, q As Double, sState As String, r As Long)
    PlanSetValue(SH_RET, n, RT_NO, n, False)
    PlanSetValue(SH_RET, n, RT_ISSUE, k, False)
    PlanSetValue(SH_RET, n, RT_EI, canon, False)
    PlanSetValue(SH_RET, n, RT_QTY, q, False)
    PlanSetValue(SH_RET, n, RT_STATE, sState, False)
    PlanSetValue(SH_RET, n, RT_ROW, r, False)
End Sub

Private Sub PlanIssRow(k As Long, iRow As Long, q As Double, cnt As Long)
    PlanSetValue(SH_ISS, k, IS_NO, k, False)
    PlanSetValue(SH_ISS, k, IS_ROW, iRow, False)
    PlanSetValue(SH_ISS, k, IS_RET, q, False)
    PlanSetValue(SH_ISS, k, IS_CNT, cnt, False)
End Sub

' the values of a posted «Возврат» row: A №, B C (the issue and its EI), D E G K from the registry, F H I J, L M, N status
Private Sub PlanRowValues(r As Long, n As Long, before As Double, after As Double, status As String)
    PlanSetValue(SH_RETURNS, r, RC_NO, n, False)
    PlanSetValue(SH_RETURNS, r, RC_ISSUE, mIk, True)
    PlanSetValue(SH_RETURNS, r, RC_EI, mIcanon, True)
    PlanSetValue(SH_RETURNS, r, RC_NAME, mSdata(SC_NAME), True)
    PlanSetValue(SH_RETURNS, r, RC_ART, mSdata(SC_ART), True)
    PlanSetValue(SH_RETURNS, r, RC_QTY, mRqty, True)
    PlanSetValue(SH_RETURNS, r, RC_UNIT, mIunit, True)
    PlanSetValue(SH_RETURNS, r, RC_DATE, mRdate, True)
    PlanSetValue(SH_RETURNS, r, RC_WHO, mIwho, True)
    PlanSetValue(SH_RETURNS, r, RC_PLACE, mRplace, True)
    PlanSetValue(SH_RETURNS, r, RC_CAT, mSdata(SC_CAT), True)
    PlanSetValue(SH_RETURNS, r, RC_BEFORE, before, True)
    PlanSetValue(SH_RETURNS, r, RC_AFTER, after, True)
    PlanSetValue(SH_RETURNS, r, RC_CTL, status, True)
End Sub

' ================================================================ «Провести» — RETURN

' Posts row r (0-based) of «Возврат». OK:<seq>, SKIP:<why> (nothing done), ERR:<why> (a problem of the row: nothing done,
' the reason is in «Контроль»), ERR-SYS:<why> (a problem of WMS itself: the return counter is not safe, an internal error),
' BLOCKED:<why> or an ApplyOperation error (ERR-RB, ERR-CRITICAL).
Function ReturnPostRow(r As Long) As String
    Dim why As String, kind As String, n As Long, res As String, b2 As Double
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReturnPostRow = "ERR:выберите строку возврата (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = ReturnRowKind(r)
    Select Case kind
    Case "POSTED"
        ReturnPostRow = "SKIP:уже проведено (возврат № " & gRtN & ")"
        Exit Function
    Case "DELETED"
        ReturnPostRow = "SKIP:возврат № " & gRtN & " удалён (сторно)"
        Exit Function
    Case "COPY"
        ReturnPostRow = "SKIP:строка помечена как КОПИЯ — очистите её кнопкой «Очистить»"
        Exit Function
    Case "FOREIGN"
        mInHandler = True
        CopyCheckReturn(r)
        mInHandler = False
        ReturnPostRow = "SKIP:в строке № возврата, который WMS здесь не проводила — строка не проводится"
        Exit Function
    Case "EMPTY"
        ReturnPostRow = "ERR:строка пуста"
        Exit Function
    End Select
    why = PostingBlockReason()
    If why <> "" Then
        If Left(why, 8) = "BLOCKED:" Then SetCtl(r, "Не проведено: WMS заблокирована — " & Mid(why, 9) & ". См. лист «" & SH_MAIN & "»")
        ReturnPostRow = why
        Exit Function
    End If
    why = CheckRowInputs(r, True)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        ReturnPostRow = "ERR:" & why
        Exit Function
    End If
    ' the new return № must be safe: a failure here is a problem of WMS, not of this row (D-059)
    n = CLng(SysNum(SK_NEXT_RET))
    why = NewRetProblem(n)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        ReturnPostRow = "ERR-SYS:" & why
        Exit Function
    End If
    b2 = WmsIssue.Round3(mSbal + mRqty)
    PlanBegin("RETURN", CStr(n))
    PlanField("RET", n)
    PlanField("ISSUE", mIk)
    PlanField("EI", mIcanon)
    PlanField("QTY", mRqty)
    PlanField("UNIT", mIunit)
    PlanField("WHO", mIwho)
    PlanField("DATE", Format(mRdate, "YYYY-MM-DD"))
    PlanField("PLACE", mRplace)
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", mSbal)
    PlanField("BAL_AFTER", b2)
    ' one operation, the writes grouped by sheet: _SYS → «Наличие» → _RET → _ISS → «Возврат» → «Заказы».X
    PlanSetValue(SYS_SHEET, SK_NEXT_RET, 1, n + 1, False)
    PlanSetValue(SH_STOCK, mIei, SC_QTY, b2, False)
    If mRplace <> CStr(mSdata(SC_PLACE)) Then PlanSetValue(SH_STOCK, mIei, SC_PLACE, mRplace, False)
    PlanRetRow(n, mIk, mIcanon, mRqty, RV_LIVE, r)
    PlanIssRow(mIk, mIrow, WmsIssue.Round3(mIret + mRqty), mIcnt + 1)
    PlanRowValues(r, n, mSbal, b2, ST_POSTED)
    PlanLockBits(SH_RETURNS, r, 0, RC_LAST, RETURN_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(mIei, b2)
    res = ApplyOperation(0, 0)
    If Left(res, 6) = "ERR-RB" Then SetCtl(r, "Не проведено: " & Mid(res, 8))
    ReturnPostRow = res
    Exit Function
EH:
    ' only the checks can get here (ApplyOperation handles its own errors): nothing was written
    ReturnPostRow = "ERR-SYS:внутренняя ошибка проверки строки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    On Error Resume Next
    mInHandler = False
    SetCtl(r, "Не проведено: " & Mid(ReturnPostRow, 9))
End Function

' «Провести» for a block of rows (the rule D-059): every visible row with a quantity F is posted by its own operation;
' a row without F is skipped, so is a row already posted, cancelled or a copy; hidden rows (a filter) are not touched.
' A problem of one row (the reason in its N) does not stop the others; a system error of WMS stops the whole block.
' Result as WmsReceipt.ReceiptPostRange: "OK=…;ERR=…;SKIP=…;STOP=<1-based row, 0 = none>", first rejection, system error.
Function ReturnPostRange(r0 As Long, r1 As Long) As String
    Dim r As Long, res As String, nOk As Long, nErr As Long, nSkip As Long, sh As Object, firstErr As String, sysErr As String
    Dim stopRow As Long
    On Error GoTo EH
    sh = ReturnsSheet()
    For r = r0 To r1
        If sh.getRows().getByIndex(r).IsVisible Then
            If sh.getCellByPosition(RC_QTY, r).getType() = com.sun.star.table.CellContentType.EMPTY Then
                nSkip = nSkip + 1
            ElseIf ReturnRowKind(r) = "OPEN" Then
                res = ReturnPostRow(r)
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
    ReturnPostRange = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";STOP=" & stopRow & Chr(10) & firstErr & Chr(10) & sysErr
    Exit Function
EH:
    sysErr = "строка " & (r + 1) & ": ERR-SYS:внутренняя ошибка проведения блока: " & Error$ & " (код " & Err & ")"
    stopRow = r + 1
    Resume DONE
End Function

' ================================================================ «Исправить» — RETURN_FIX (one composite operation)
' The old return is reversed and the new one applied in one journal line: quantity F, date H, place J. The issue and the
' EI stay (another issue: storno and a new return). After the correction: the live returns of the issue ≤ the issued
' quantity, and the balance of the EI does not become negative (what was issued out of the return stays issued).

Function ReturnFixRow(r As Long, vQty As Variant, vDate As Variant, vPlace As Variant) As String
    Dim why As String, kind As String, n As Long, rd As Variant, k As Long, q1 As Double, d1 As Double, p1 As String, msg As String
    Dim sh As Object, res As String, before As Double, b2 As Double, place0 As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReturnFixRow = "ERR:выберите строку возврата (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = ReturnsSheet()
    kind = ReturnRowKind(r)
    why = PostedRowProblem(kind, "Исправить")
    If why <> "" Then
        ReturnFixRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        ReturnFixRow = why
        Exit Function
    End If
    n = gRtN
    If Not RetRegistered(n, rd) Then
        ReturnFixRow = "ERR-SYS:возврат № " & n & " не найден в " & SH_RET
        Exit Function
    End If
    k = CLng(rd(RT_ISSUE))
    q1 = rd(RT_QTY)
    why = CheckIssue(k)
    If why <> "" Then
        ReturnFixRow = "ERR:" & why
        Exit Function
    End If
    If CStr(rd(RT_EI)) <> mIcanon Then
        ReturnFixRow = "ERR-SYS:возврат № " & n & " записан на " & rd(RT_EI) & ", а выдача № " & k & " — на " & mIcanon & " — нужна самопроверка"
        Exit Function
    End If
    d1 = sh.getCellByPosition(RC_DATE, r).getValue()
    p1 = sh.getCellByPosition(RC_PLACE, r).getString()
    If Not WmsOrders.ParseQtyValueAs(vQty, "количество возврата", mRqty, msg) Then
        ReturnFixRow = "ERR:" & msg
        Exit Function
    End If
    If VarType(vDate) = 8 Then
        If Not WmsOrders.ParseDateTextAs(CStr(vDate), "дата возврата", mRdate, msg) Then
            ReturnFixRow = "ERR:" & msg
            Exit Function
        End If
    Else
        mRdate = Int(CDbl(vDate))
    End If
    mRplace = Trim(CStr(vPlace))
    If mRplace = "" Then mRplace = CStr(mSdata(SC_PLACE))
    If mRqty > Available(q1) + 0.0000001 Then
        ReturnFixRow = "ERR:больше, чем можно вернуть по выдаче № " & k & ": выдано " & WmsIssue.QtyText(mIqty) & ", другими возвратами возвращено " _
            & WmsIssue.QtyText(WmsIssue.Round3(mIret - q1)) & ", можно не больше " & WmsIssue.QtyText(Available(q1))
        Exit Function
    End If
    If mRqty < q1 - 0.0000001 And mSbal < WmsIssue.Round3(q1 - mRqty) - 0.0000001 Then
        ReturnFixRow = "ERR:остаток " & mIcanon & " — " & WmsIssue.QtyText(mSbal) & ": из возвращённого уже выдано, уменьшение возврата на " _
            & WmsIssue.QtyText(WmsIssue.Round3(q1 - mRqty)) & " сделало бы остаток отрицательным; сначала удалите (сторно) последующие выдачи"
        Exit Function
    End If
    If mRdate < mIdate Then
        ReturnFixRow = "ERR:дата возврата " & Format(mRdate, "DD.MM.YYYY") & " раньше даты выдачи № " & k & " (" & Format(mIdate, "DD.MM.YYYY") & ")"
        Exit Function
    End If
    If Abs(mRqty - q1) < 0.0000001 And mRdate = Int(d1) And mRplace = p1 Then
        ReturnFixRow = "SKIP:ничего не изменилось"
        Exit Function
    End If
    place0 = CStr(mSdata(SC_PLACE))
    before = WmsIssue.Round3(mSbal - q1)
    b2 = WmsIssue.Round3(before + mRqty)
    PlanBegin("RETURN_FIX", CStr(n))
    PlanField("RET", n)
    PlanField("ISSUE", k)
    PlanField("EI", mIcanon)
    PlanField("OLD_QTY", q1)
    PlanField("QTY", mRqty)
    PlanField("OLD_DATE", Format(d1, "YYYY-MM-DD"))
    PlanField("DATE", Format(mRdate, "YYYY-MM-DD"))
    PlanField("OLD_PLACE", p1)
    PlanField("PLACE", mRplace)
    PlanField("WHO", mIwho)
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", mSbal)
    PlanField("BAL_AFTER", b2)
    If Abs(b2 - mSbal) > 0.0000001 Then PlanSetValue(SH_STOCK, mIei, SC_QTY, b2, False)
    If mRplace <> place0 Then PlanSetValue(SH_STOCK, mIei, SC_PLACE, mRplace, False)
    PlanSetValue(SH_RET, n, RT_QTY, mRqty, False)
    If CLng(rd(RT_ROW)) <> r Then PlanSetValue(SH_RET, n, RT_ROW, r, False)
    PlanIssRow(k, mIrow, WmsIssue.Round3(mIret - q1 + mRqty), mIcnt)
    PlanSetValue(SH_RETURNS, r, RC_QTY, mRqty, False)
    PlanSetValue(SH_RETURNS, r, RC_DATE, mRdate, False)
    PlanSetValue(SH_RETURNS, r, RC_PLACE, mRplace, False)
    PlanSetValue(SH_RETURNS, r, RC_BEFORE, before, False)
    PlanSetValue(SH_RETURNS, r, RC_AFTER, b2, False)
    PlanSetValue(SH_RETURNS, r, RC_CTL, ST_FIXED, False)
    PlanLockBits(SH_RETURNS, r, 0, RC_LAST, RETURN_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(mIei, b2)
    res = ApplyOperation(0, 0)
    ReturnFixRow = res
    Exit Function
EH:
    ReturnFixRow = "ERR-SYS:внутренняя ошибка проверки исправления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' "" when a row of this kind is a posted return that an action may change
Private Function PostedRowProblem(kind As String, sAction As String) As String
    Select Case kind
    Case "POSTED"
        PostedRowProblem = ""
    Case "DELETED"
        PostedRowProblem = "возврат № " & gRtN & " уже удалён (сторно)"
    Case "COPY", "FOREIGN"
        PostedRowProblem = "строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
    Case "OPEN"
        PostedRowProblem = "строка не проведена — «" & sAction & "» нужно только для проведённых возвратов"
    Case Else
        PostedRowProblem = "строка пуста"
    End Select
End Function

' ================================================================ «Удалить» — RETURN_DEL (storno)
' The returned quantity leaves the EI again, _RET gets STORNO, the issue gets it back as returnable; the row stays as the
' history with the status «Удалено (сторно)», its № is never reused. Refused when the balance of the EI is smaller than
' the return: part of it was issued again (the balance would become negative — storno those issues first).

Function ReturnDeleteRow(r As Long) As String
    Dim why As String, kind As String, n As Long, rd As Variant, k As Long, q1 As Double, di As Variant, iret As Double, icnt As Long
    Dim ei As Long, canon As String, p As String, sd As Variant, bal As Double, b2 As Double, iRow As Long, moved As Boolean, sh As Object
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReturnDeleteRow = "ERR:выберите строку возврата (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = ReturnsSheet()
    kind = ReturnRowKind(r)
    If kind = "DELETED" Then
        ReturnDeleteRow = "SKIP:возврат № " & gRtN & " уже удалён (сторно)"
        Exit Function
    End If
    why = PostedRowProblem(kind, "Удалить")
    If why <> "" Then
        ReturnDeleteRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        ReturnDeleteRow = why
        Exit Function
    End If
    n = gRtN
    If Not RetRegistered(n, rd) Then
        ReturnDeleteRow = "ERR-SYS:возврат № " & n & " не найден в " & SH_RET
        Exit Function
    End If
    k = CLng(rd(RT_ISSUE))
    q1 = rd(RT_QTY)
    canon = CStr(rd(RT_EI))
    If Not WmsOrders.StrictEI(canon, ei) Then
        ReturnDeleteRow = "ERR-SYS:в " & SH_RET & " у возврата № " & n & " неверный ЕИ «" & canon & "»"
        Exit Function
    End If
    If Not IssRegistered(k, di) Then
        ReturnDeleteRow = "ERR-SYS:итоги выдачи № " & k & " не найдены в " & SH_ISS & " — нужна самопроверка"
        Exit Function
    End If
    iret = di(IS_RET)
    icnt = CLng(di(IS_CNT))
    iRow = FindIssueRow(k, moved)
    If iRow < 1 Then iRow = CLng(di(IS_ROW))
    p = WmsIssue.StockRowProblem(ei, canon)
    If p <> "" Then
        ReturnDeleteRow = "ERR:" & p
        Exit Function
    End If
    sd = WmsIssue.StockData(ei)
    If Not WmsIssue.StockQty(ei, bal) Then
        ReturnDeleteRow = "ERR:остаток " & canon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If bal < q1 - 0.0000001 Then
        ReturnDeleteRow = "ERR:остаток " & canon & " — " & WmsIssue.QtyText(bal) & ", а возврат — " & WmsIssue.QtyText(q1) _
            & ": из возвращённого уже выдано, сторно сделало бы остаток отрицательным; сначала удалите (сторно) последующие выдачи"
        Exit Function
    End If
    b2 = WmsIssue.Round3(bal - q1)
    PlanBegin("RETURN_DEL", CStr(n))
    PlanField("RET", n)
    PlanField("ISSUE", k)
    PlanField("EI", canon)
    PlanField("QTY", q1)
    PlanField("WHO", sh.getCellByPosition(RC_WHO, r).getString())
    PlanField("DATE", Format(sh.getCellByPosition(RC_DATE, r).getValue(), "YYYY-MM-DD"))
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", bal)
    PlanField("BAL_AFTER", b2)
    PlanSetValue(SH_STOCK, ei, SC_QTY, b2, False)
    PlanSetValue(SH_RET, n, RT_STATE, RV_STORNO, False)
    If CLng(rd(RT_ROW)) <> r Then PlanSetValue(SH_RET, n, RT_ROW, r, False)
    PlanIssRow(k, iRow, WmsIssue.Round3(iret - q1), icnt - 1)
    PlanSetValue(SH_RETURNS, r, RC_CTL, ST_DELETED, False)
    PlanLockBits(SH_RETURNS, r, 0, RC_LAST, RETURN_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(ei, b2)
    ReturnDeleteRow = ApplyOperation(0, 0)
    Exit Function
EH:
    ReturnDeleteRow = "ERR-SYS:внутренняя ошибка проверки удаления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Проверить»

' An open row: all the checks of «Провести», the result in «Контроль» (N). A posted row: the return, its issue and the
' registry — OK:<text> or ERR:<text> for a message.
Function ReturnCheckRow(r As Long) As String
    Dim kind As String, why As String, n As Long, rd As Variant, k As Long, sh As Object, bal As Double
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReturnCheckRow = "ERR:выберите строку возврата (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = ReturnsSheet()
    kind = ReturnRowKind(r)
    Select Case kind
    Case "OPEN"
        mInHandler = True
        why = CheckRowInputs(r, True)
        If why <> "" Then
            SetCtl(r, "Ошибка: " & why)
            ReturnCheckRow = "ERR:" & why
        Else
            SetCtl(r, "Можно провести: " & IssueInfo() & PlaceNote())
            ReturnCheckRow = "OK:можно провести: " & IssueInfo() & PlaceNote()
        End If
        mInHandler = False
    Case "POSTED", "DELETED"
        n = gRtN
        If Not RetRegistered(n, rd) Then
            ReturnCheckRow = "ERR:возврат № " & n & " не найден в " & SH_RET
            Exit Function
        End If
        k = CLng(rd(RT_ISSUE))
        why = CheckIssue(k)
        If why <> "" Then
            ReturnCheckRow = "ERR:возврат № " & n & ": " & why
            Exit Function
        End If
        ReturnCheckRow = "OK:возврат № " & n & " (" & IIf(kind = "DELETED", "удалён (сторно)", "действует") & "): " & rd(RT_EI) & ", " _
            & WmsIssue.QtyText(rd(RT_QTY)) & " " & mIunit & "; " & IssueInfo() & "; остаток " & mIcanon & " — " & WmsIssue.QtyText(mSbal)
    Case "COPY", "FOREIGN"
        ReturnCheckRow = "ERR:строка — КОПИЯ" & IIf(gRtWhy <> "", " (" & gRtWhy & ")", "") & ", её можно только очистить кнопкой «Очистить»"
    Case Else
        ReturnCheckRow = "ERR:строка пуста — введите № выдачи или ЕИ"
    End Select
    Exit Function
EH:
    mInHandler = False
    ReturnCheckRow = "ERR:внутренняя ошибка проверки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Найти выдачу»

' The posted issues of EI canon (and of recipient who, when given) that can still be returned: "k1;k2;…" in the order
' of the sheet, at most ISSUE_LIST_MAX of the latest ones; found by one array formula of the Calc engine over the used
' rows of «Выдачи» (no loop over the sheet), then each candidate is checked by its key. "" when none. ok = False when
' the formula could not be used.
Function IssuesOfEI(canon As String, who As String, ByRef ok As Boolean) As String
    Dim sh As Object, last As Long, rngA As String, rngL As String, rngI As String, rngR As String, f As String, s As String
    Dim a As Variant, i As Long, out As String, k As Long, n As Long
    sh = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(sh)
    ok = True
    If last < 1 Then Exit Function
    rngA = "$'" & SH_ISSUES & "'.$A$2:$A$" & (last + 1)
    rngI = "$'" & SH_ISSUES & "'.$I$2:$I$" & (last + 1)
    rngL = "$'" & SH_ISSUES & "'.$L$2:$L$" & (last + 1)
    rngR = "$'" & SH_ISSUES & "'.$R$2:$R$" & (last + 1)
    f = "TEXTJOIN("";"";1;IF((" & rngL & "=""" & Replace(canon, """", """""") & """)*(LEFT(" & rngR & ";" & Len(ST_POSTED) & ")=""" & ST_POSTED & """)"
    If who <> "" Then f = f & "*(LOWER(TRIM(" & rngI & "))=""" & Replace(LCase(Trim(who)), """", """""") & """)"
    f = f & ";" & rngA & ";""""))"
    s = WmsOrders.ScratchEval(f, ok)
    If Not ok Or s = "" Then Exit Function
    a = Split(s, ";")
    For i = UBound(a) To 0 Step -1
        If IsDigits(CStr(a(i))) And Len(a(i)) <= 9 Then k = CLng(a(i)) Else k = 0
        ' an unmarked copy of an issue row (saved without WMS) repeats its №: each issue is listed once
        If k > 0 And InStr(";" & out & ";", ";" & k & ";") = 0 Then
            If CheckIssue(k) = "" Then
                If Available(0) > 0.0000001 Then
                    out = k & IIf(out <> "", ";", "") & out
                    n = n + 1
                    If n >= ISSUE_LIST_MAX Then Exit For
                End If
            End If
        End If
    Next i
    IssuesOfEI = out
End Function

' ================================================================ dependents of an EI (D-069: the storno of its receipt)

' The live movements that depend on EI canon: posted issues («Выдачи»: L = canon, R «Проведено» / «Проведено
' (исправлено)») and live returns (_RET: C = canon, E = LIVE) — counted by one formula of the Calc engine over the used
' rows of both sheets (COUNTIFS; no loop over a sheet). False when the formula could not be used: the caller refuses,
' nothing may be decided without it.
Function LiveDependents(canon As String, ByRef nIss As Long, ByRef nRet As Long) As Boolean
    Dim li As Long, lr As Long, q As String, l As String, rr As String, c As String, e As String, f As String, g As String
    Dim s As String, ok As Boolean, p As Long
    nIss = 0
    nRet = 0
    li = WmsOrders.LastRow(WmsIssue.IssuesSheet())
    lr = WmsOrders.LastRow(RetSheet())
    q = """" & Replace(canon, """", """""") & """"
    f = "0"
    If li >= 1 Then
        l = "$'" & SH_ISSUES & "'.$L$2:$L$" & (li + 1)
        rr = "$'" & SH_ISSUES & "'.$R$2:$R$" & (li + 1)
        f = "COUNTIFS(" & l & ";" & q & ";" & rr & ";""" & ST_POSTED & """)+COUNTIFS(" & l & ";" & q & ";" & rr & ";""" & ST_FIXED & """)"
    End If
    g = "0"
    If lr >= 1 Then
        c = "$'" & SH_RET & "'.$C$2:$C$" & (lr + 1)
        e = "$'" & SH_RET & "'.$E$2:$E$" & (lr + 1)
        g = "COUNTIFS(" & c & ";" & q & ";" & e & ";""" & RV_LIVE & """)"
    End If
    s = WmsOrders.ScratchEval("(" & f & ")&""|""&(" & g & ")", ok)
    If Not ok Then Exit Function
    p = InStr(s, "|")
    If p < 2 Then Exit Function
    If Not IsDigits(Left(s, p - 1)) Or Not IsDigits(Mid(s, p + 1)) Then Exit Function
    nIss = CLng(Left(s, p - 1))
    nRet = CLng(Mid(s, p + 1))
    LiveDependents = True
End Function

' a short description of issue № k for a list (after CheckIssue): «№ 12 от 20.09.2026 — Иванов: можно вернуть 7 шт»
Function IssueLine(k As Long) As String
    If CheckIssue(k) <> "" Then
        IssueLine = "№ " & k & " — недоступна"
        Exit Function
    End If
    IssueLine = "№ " & k & " от " & Format(mIdate, "DD.MM.YYYY") & " — «" & mIwho & "»: выдано " & WmsIssue.QtyText(mIqty) & ", возвращено " _
        & WmsIssue.QtyText(mIret) & ", можно вернуть " & WmsIssue.QtyText(Available(0)) & IIf(mIunit <> "", " " & mIunit, "")
End Function

' the data a dialog needs for issue № k: "" and the values, or the reason
Function IssueForDialog(k As Long, ByRef sInfo As String, ByRef sPlace As String, ByRef avail As Double) As String
    IssueForDialog = CheckIssue(k)
    If IssueForDialog <> "" Then Exit Function
    gRtIssueEI = mIcanon
    gRtIssueWho = mIwho
    sInfo = mIcanon & " " & mSdata(SC_NAME) & "; " & IssueInfo()
    sPlace = CStr(mSdata(SC_PLACE))
    avail = Available(0)
End Function

' ================================================================ light change handler (spec §13, D-008)

' «Возврат» sheet event "Content changed": the preview of open rows. Leaves at once for columns that do not matter, never
' scans the sheet, is guarded against re-entry and ignores changes made by operations.
Sub OnReturnsChange(Optional oTarget As Variant)
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
    ' the inputs A (a key that appeared by paste) B C F H I J matter; D E G K L M N are written by WMS, O is free text
    If ra.StartColumn > RC_PLACE Then Exit Sub
    If ra.StartColumn = RC_NAME And ra.EndColumn <= RC_ART Then Exit Sub
    mInHandler = True
    WmsInit()
    r1 = ra.EndRow
    If r1 - ra.StartRow + 1 > PREVIEW_MAX_ROWS Then
        last = WmsOrders.LastRow(ReturnsSheet())
        If r1 > last Then r1 = last
    End If
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: проверка строки возврата")
    inCtx = True
    For r = IIf(ra.StartRow < 1, 1, ra.StartRow) To r1
        n = n + 1
        If n > PREVIEW_MAX_ROWS Then Exit For
        PreviewReturnRow(r, (ra.StartColumn <= RC_NO And ra.EndColumn >= RC_NO))
    Next r
    um.leaveUndoContext()
    mInHandler = False
    Exit Sub
EH:
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    mInHandler = False
End Sub

' The preview of an open row: the issue in B → C (its EI), D E G K from the registry, I (its recipient) and J (the current
' place) when empty, L the current balance; «Контроль» N: the first problem of the filled inputs, otherwise what the
' issue allows. Without B, an EI in C shows the product and asks for the issue. A value the user typed is never replaced.
' Posted, cancelled and copy rows are never touched here; a № in A that WMS did not write marks the row КОПИЯ.
Sub PreviewReturnRow(r As Long, bKeyTouched As Boolean)
    Dim sh As Object, kind As String, k As Long, why As String, n As Long, canon As String, msg As String, p As String, sd As Variant, bal As Double
    sh = ReturnsSheet()
    kind = ReturnRowKind(r)
    Select Case kind
    Case "FOREIGN"
        CopyCheckReturn(r)
        Exit Sub
    Case "EMPTY"
        ClearDerived(r, True)
        Exit Sub
    Case "OPEN"
    Case Else
        Exit Sub
    End Select
    k = IssueNoOf(sh.getCellByPosition(RC_ISSUE, r))
    If k > 0 Then
        why = CheckIssue(k)
        If why = "" Then
            FillFromIssue(r)
            why = CheckRowInputs(r, False)
            If why = "" Then
                SetCtl(r, IIf(sh.getCellByPosition(RC_QTY, r).getType() <> com.sun.star.table.CellContentType.EMPTY _
                    And sh.getCellByPosition(RC_DATE, r).getType() <> com.sun.star.table.CellContentType.EMPTY, "Можно провести: ", "") & UCase(Left(IssueInfo(), 1)) & Mid(IssueInfo(), 2) _
                    & PlaceNote())
            Else
                SetCtl(r, "Ошибка: " & why)
            End If
        Else
            ClearDerived(r, False)
            SetCtl(r, "Ошибка: " & why)
        End If
        Exit Sub
    End If
    If k < 0 Then
        ClearDerived(r, False)
        SetCtl(r, "Ошибка: № выдачи (B) «" & sh.getCellByPosition(RC_ISSUE, r).getString() & "» — не номер")
        Exit Sub
    End If
    ' no issue yet: an EI shows its product, the issue is still needed
    If sh.getCellByPosition(RC_EI, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not WmsIssue.NormalizeEI(sh.getCellByPosition(RC_EI, r).getString(), n, canon, msg) Then
            ClearDerived(r, False)
            SetCtl(r, "Ошибка: " & msg)
            Exit Sub
        End If
        p = WmsIssue.StockRowProblem(n, canon)
        If p <> "" Then
            ClearDerived(r, False)
            SetCtl(r, "Ошибка: " & p)
            Exit Sub
        End If
        sd = WmsIssue.StockData(n)
        If Not WmsIssue.StockQty(n, bal) Then bal = 0
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_EI, r), canon)
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_NAME, r), sd(SC_NAME))
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_ART, r), sd(SC_ART))
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_UNIT, r), sd(SC_UNIT))
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_CAT, r), sd(SC_CAT))
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_BEFORE, r), bal)
        SetCtl(r, "Укажите № выдачи (B) или нажмите «Найти выдачу» — возврат засчитывается конкретной выдаче")
        Exit Sub
    End If
    ClearDerived(r, False)
    SetCtl(r, "Укажите № выдачи (B) или ЕИ (C)")
End Sub

' the data of the issue (after CheckIssue) shown in an open row; C I J only when empty (never replaces what the user typed)
Private Sub FillFromIssue(r As Long)
    Dim sh As Object
    sh = ReturnsSheet()
    If sh.getCellByPosition(RC_EI, r).getType() = com.sun.star.table.CellContentType.EMPTY Then sh.getCellByPosition(RC_EI, r).setString(mIcanon)
    WmsOrders.SetIfDiff(sh.getCellByPosition(RC_NAME, r), mSdata(SC_NAME))
    WmsOrders.SetIfDiff(sh.getCellByPosition(RC_ART, r), mSdata(SC_ART))
    WmsOrders.SetIfDiff(sh.getCellByPosition(RC_UNIT, r), mIunit)
    WmsOrders.SetIfDiff(sh.getCellByPosition(RC_CAT, r), mSdata(SC_CAT))
    WmsOrders.SetIfDiff(sh.getCellByPosition(RC_BEFORE, r), mSbal)
    If Trim(sh.getCellByPosition(RC_WHO, r).getString()) = "" Then sh.getCellByPosition(RC_WHO, r).setString(mIwho)
    If Trim(sh.getCellByPosition(RC_PLACE, r).getString()) = "" And CStr(mSdata(SC_PLACE)) <> "" Then sh.getCellByPosition(RC_PLACE, r).setString(mSdata(SC_PLACE))
    If sh.getCellByPosition(RC_AFTER, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(RC_AFTER, r))
End Sub

' the WMS-filled cells of an open row (D E G K L M, and N when bAll)
Private Sub ClearDerived(r As Long, bAll As Boolean)
    Dim sh As Object, cols As Variant, i As Integer
    sh = ReturnsSheet()
    cols = Array(RC_NAME, RC_ART, RC_UNIT, RC_CAT, RC_BEFORE, RC_AFTER)
    For i = 0 To UBound(cols)
        If sh.getCellByPosition(cols(i), r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(cols(i), r))
    Next i
    If bAll Then
        If sh.getCellByPosition(RC_CTL, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(RC_CTL, r))
    End If
End Sub

' handler guard for the UI module (its own writes into «Возврат» outside operations, e.g. «Найти выдачу»)
Sub HandlerOff(b As Boolean)
    mInHandler = b
End Sub

' ================================================================ «Очистить» — an unposted row or a copy (no accounting change)

Function ReturnClearRow(r As Long) As String
    Dim sh As Object, kind As String, wasProt As Boolean, i As Integer, pr As Object, x As Double, d As Variant
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReturnClearRow = "ERR:выберите строку возврата (не заголовок)"
        Exit Function
    End If
    If gBusy Then
        ReturnClearRow = "BUSY:операция уже выполняется"
        Exit Function
    End If
    sh = ReturnsSheet()
    kind = ReturnRowKind(r)
    Select Case kind
    Case "POSTED"
        ReturnClearRow = "ERR:строка — проведённый возврат № " & gRtN & " — для отмены используйте «Удалить»"
        Exit Function
    Case "DELETED"
        ReturnClearRow = "ERR:возврат № " & gRtN & " удалён (сторно) — строка хранит историю и не очищается"
        Exit Function
    Case "COPY"
        ' defence: never the registered row of a return, even when it carries a copy mark
        If sh.getCellByPosition(RC_NO, r).getType() = com.sun.star.table.CellContentType.VALUE Then
            x = sh.getCellByPosition(RC_NO, r).getValue()
            If x >= 1 And x = Int(x) And x <= MAX_SHEET_ROW Then
                If RetRegistered(CLng(x), d) Then
                    If VarType(d(RT_ROW)) = 5 Then
                        If CLng(d(RT_ROW)) = r Then
                            ReturnClearRow = "ERR:строка — зарегистрированный возврат № " & CLng(x) & "; копией является другая строка с этим №"
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
    sh.getCellRangeByPosition(0, r, RC_LAST, r).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME _
        + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA)
    ' a row inserted between posted rows or a copy carries the protection of its origin: back to the input layout
    For i = 0 To RC_LAST
        pr = sh.getCellByPosition(i, r).CellProtection
        pr.IsLocked = (Mid(RETURN_LOCKS_OPEN, i + 1, 1) = "1")
        sh.getCellByPosition(i, r).CellProtection = pr
    Next i
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
    ReturnClearRow = "OK:" & IIf(kind = "COPY" Or kind = "FOREIGN", "копия очищена", "строка очищена")
    Exit Function
EH:
    ReturnClearRow = "ERR:не удалось очистить строку: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
End Function

' ================================================================ «Главная»: returns not posted yet (spec §23)

' rows with a quantity (F) but no № (A): two Calc queries per block instead of a row loop. -1 when it fails.
Function UnpostedReturns() As Long
    Dim sh As Object, last As Long, blocks As Variant, got As Variant, i As Long, j As Long, n As Long, flags As Long
    On Error GoTo EH
    sh = ReturnsSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    blocks = sh.getCellRangeByPosition(RC_NO, 1, RC_NO, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= UNPOSTED_MAX_BLOCKS Then Exit For
        got = sh.getCellRangeByPosition(RC_QTY, blocks(i).StartRow, RC_QTY, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            n = n + got(j).EndRow - got(j).StartRow + 1
        Next j
    Next i
    UnpostedReturns = n
    Exit Function
EH:
    UnpostedReturns = -1
End Function

' ================================================================ self-check (spec §49, a part of «WMS-Доктор»)
' Every return of _RET (№, issue, quantity, state), the sums of the live returns of every issue against _ISS, the issue
' itself (posted, returned ≤ issued), NEXT_RET above every return №, the row hints. Chunks, only the needed columns.

Function ReturnsCheck() As String
    Dim rt As Object, isx As Object, iss As Object, last As Long, r0 As Long, r1 As Long, i As Long, d As Variant, t0 As Long
    Dim nextRet As Long, nextNo As Long, sumQ() As Double, cnt() As Long, k As Long, nRet As Long, nStorno As Long, nBad As Long
    Dim first As String, nIss As Long, nStale As Long, q As Double, dq As Variant, dr As Variant, dk As Variant, st As String
    Dim issQ() As Double, issState() As Integer, x As Double, retRows As Object, maxRet As Long, lastIss As Long, lastR As Long
    Dim colA As Variant, colA0 As Long
    On Error GoTo EH
    t0 = GetSystemTicks()
    rt = RetSheet()
    isx = IssSheet()
    iss = WmsIssue.IssuesSheet()
    retRows = ReturnsSheet()
    nextRet = CLng(SysNum(SK_NEXT_RET))
    nextNo = CLng(SysNum(SK_NEXT_NO))
    ReDim sumQ(nextNo)
    ReDim cnt(nextNo)
    lastR = WmsOrders.LastRow(retRows)
    colA0 = -1
    ' 1. _RET: every registered return
    last = WmsOrders.LastRow(rt)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = rt.getCellRangeByPosition(0, r0, RT_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(d(i)(RT_NO)) <> "" Then
                If VarType(d(i)(RT_NO)) <> 5 Then
                    Bad(nBad, first, SH_RET & " строка " & (r0 + i + 1) & ": № не число")
                ElseIf d(i)(RT_NO) <> r0 + i Then
                    Bad(nBad, first, SH_RET & " строка " & (r0 + i + 1) & ": записан № " & d(i)(RT_NO) & " (плотная адресация нарушена)")
                Else
                    nRet = nRet + 1
                    If r0 + i > maxRet Then maxRet = r0 + i
                    st = CStr(d(i)(RT_STATE))
                    If VarType(d(i)(RT_ISSUE)) <> 5 Or VarType(d(i)(RT_QTY)) <> 5 Then
                        Bad(nBad, first, "возврат № " & (r0 + i) & ": № выдачи или количество не число")
                    ElseIf d(i)(RT_ISSUE) < 1 Or d(i)(RT_ISSUE) >= nextNo Then
                        Bad(nBad, first, "возврат № " & (r0 + i) & ": выдачи № " & d(i)(RT_ISSUE) & " нет")
                    ElseIf d(i)(RT_QTY) <= 0 Then
                        Bad(nBad, first, "возврат № " & (r0 + i) & ": количество не больше 0")
                    ElseIf st = RV_LIVE Then
                        k = CLng(d(i)(RT_ISSUE))
                        sumQ(k) = sumQ(k) + d(i)(RT_QTY)
                        cnt(k) = cnt(k) + 1
                    ElseIf st = RV_STORNO Then
                        nStorno = nStorno + 1
                    Else
                        Bad(nBad, first, "возврат № " & (r0 + i) & ": неизвестное состояние «" & st & "»")
                    End If
                    ' the row hint: «Возврат».A of that row read in chunks (a stale hint is not an error: rows inserted above)
                    x = -1
                    If VarType(d(i)(RT_ROW)) = 5 Then x = d(i)(RT_ROW)
                    If x < 1 Or x > lastR Then
                        nStale = nStale + 1
                    Else
                        If CLng(x) < colA0 Or CLng(x) >= colA0 + CHECK_CHUNK_ROWS Or colA0 < 0 Then
                            colA0 = CLng(x)
                            colA = retRows.getCellRangeByPosition(RC_NO, colA0, RC_NO, IIf(colA0 + CHECK_CHUNK_ROWS - 1 > lastR, lastR, colA0 + CHECK_CHUNK_ROWS - 1)).getDataArray()
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
    If maxRet >= nextRet Then Bad(nBad, first, "NEXT_RET " & nextRet & " не больше максимального № возврата " & maxRet)
    ' 2. the issues: quantity and state by № (chunks of «Выдачи» A E R)
    ReDim issQ(nextNo)
    ReDim issState(nextNo)
    lastIss = WmsOrders.LastRow(iss)
    For r0 = 1 To lastIss Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > lastIss Then r1 = lastIss
        dk = iss.getCellRangeByPosition(IC_NO, r0, IC_NO, r1).getDataArray()
        dq = iss.getCellRangeByPosition(IC_QTY, r0, IC_QTY, r1).getDataArray()
        dr = iss.getCellRangeByPosition(IC_CTL, r0, IC_CTL, r1).getDataArray()
        For i = 0 To r1 - r0
            If VarType(dk(i)(0)) = 5 Then
                x = dk(i)(0)
                If x >= 1 And x < nextNo And x = Int(x) Then
                    k = CLng(x)
                    If Left(CStr(dr(i)(0)), Len(ST_POSTED)) = ST_POSTED And VarType(dq(i)(0)) = 5 Then
                        issQ(k) = dq(i)(0)
                        issState(k) = 1
                    ElseIf CStr(dr(i)(0)) = ST_DELETED Then
                        issState(k) = 2
                    End If
                End If
            End If
        Next i
    Next r0
    ' 3. _ISS against the sums of the live returns
    last = WmsOrders.LastRow(isx)
    If last > nextNo Then Bad(nBad, first, SH_ISS & ": строк больше, чем выдач (NEXT_NO " & nextNo & ")")
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        If r1 > nextNo - 1 Then r1 = nextNo - 1
        If r1 < r0 Then Exit For
        d = isx.getCellRangeByPosition(0, r0, IS_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            k = r0 + i
            If CStr(d(i)(IS_NO)) <> "" Then
                If VarType(d(i)(IS_NO)) <> 5 Then
                    Bad(nBad, first, SH_ISS & " строка " & (k + 1) & ": № не число")
                ElseIf d(i)(IS_NO) <> k Then
                    Bad(nBad, first, SH_ISS & " строка " & (k + 1) & ": записан № " & d(i)(IS_NO))
                Else
                    nIss = nIss + 1
                    If Abs(CDbl(d(i)(IS_RET)) - sumQ(k)) > 0.0005 Or CLng(d(i)(IS_CNT)) <> cnt(k) Then
                        Bad(nBad, first, "выдача № " & k & ": в " & SH_ISS & " возвращено " & d(i)(IS_RET) & " (" & d(i)(IS_CNT) & "), по возвратам " _
                            & WmsIssue.QtyText(sumQ(k)) & " (" & cnt(k) & ")")
                    End If
                End If
            ElseIf cnt(k) > 0 Then
                Bad(nBad, first, "выдача № " & k & ": есть действующие возвраты, но нет итогов в " & SH_ISS)
            End If
            If cnt(k) > 0 Then
                If issState(k) <> 1 Then
                    Bad(nBad, first, "выдача № " & k & " с действующими возвратами " & IIf(issState(k) = 2, "удалена (сторно)", "не найдена на листе «" & SH_ISSUES & "»"))
                ElseIf sumQ(k) > issQ(k) + 0.0005 Then
                    Bad(nBad, first, "выдача № " & k & ": возвращено " & WmsIssue.QtyText(sumQ(k)) & " больше выданного " & WmsIssue.QtyText(issQ(k)))
                End If
            End If
        Next i
    Next r0
    If nBad > 0 Then
        ReturnsCheck = "ОШИБКА: расхождений " & nBad & " (" & first & ")"
    Else
        ReturnsCheck = "возвратов " & nRet & " (из них сторно " & nStorno & "), выдач с возвратами " & nIss & ", расхождений нет, NEXT_RET " & nextRet _
            & " выше всех № возвратов" & IIf(nStale > 0, ", устаревших подсказок строк " & nStale & " (исправятся при следующей операции)", "") _
            & ", " & (GetSystemTicks() - t0) & " мс"
    End If
    Exit Function
EH:
    ReturnsCheck = "ОШИБКА: проверка возвратов не выполнена: " & Error$ & " (строка " & Erl & ")"
End Function

Private Sub Bad(ByRef nBad As Long, ByRef first As String, s As String)
    nBad = nBad + 1
    If first = "" Then first = s
End Sub
