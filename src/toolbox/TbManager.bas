' TbManager — WMS_MANAGER: отчёт руководителю и контроль смены (задание «FINAL WMS MARATHON», §6; M6 PRIME §5–§6).
' Показатели по снимку WMS за сегодня / неделю / месяц (или период B2–B3): складские операции, ЕИ в движении,
' инвентаризационные расхождения, проблемные заказы и документы, машины «Приход авто» (приехало, уехало, среднее время),
' и работа, не видная в движении товара, — «Журнал работы» кладовщика (разгрузки и погрузки, погрузчик, бухгалтерия,
' закупки, проблемные поставки, прочее). «Сформировать отчёт руководителю» — аккуратный документ (ODT и PDF) в
' ../WMS_Reports без служебных данных WMS; «В PDF» — лист «Отчёт». «Контроль дня» — одним запуском перед окончанием
' работы: машины на территории, непроведённые строки, неразобранный «Иной приход», проблемы диагностики (проверки
' WMS_DOCTOR), приходы без документов, просроченные заказы, резервная копия за сегодня, снимок; и «Что изменилось
' сегодня». Ничего не исправляет — только список и что делать.
Option Explicit

Private Const SH_REP = "Отчёт"
Private Const SH_LOG = "Журнал работы"
Private Const SH_CTL = "Контроль дня"
Public Const WORK_KINDS = "Разгрузка / погрузка|Работа погрузчиком|Запрос бухгалтерии|Закупки|Проблемная поставка|Прочие работы"
Private Const LONG_MIN = 120           ' a long visit of a vehicle, minutes (as on «Приход авто» of WMS)

Sub BtnMgrRefresh(Optional oEvent As Variant)
    TbMsg(MgrRefresh())
End Sub

Function MgrRefresh() As String
    Dim why As String, snap As String, t As Variant, i As Integer, s As String, nCar As Long
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        MgrRefresh = "ERR:" & why
        Exit Function
    End If
    t = Array("stock", "orders", "special", "issues", "returns", "adjustments")
    For i = 0 To UBound(t)
        s = s & IIf(s <> "", ", ", "") & t(i) & " " & TbLoadTable(snap, t(i) & ".csv", "_" & t(i), True)
    Next i
    nCar = TbLoadOptional(snap, "cars.csv", "_cars", True, TB_CAR_KEYS)
    ' a snapshot of WMS 0.6 has no vehicles: no sheet _cars (the report says «нет данных»)
    If nCar < 0 Then ThisComponent.Sheets.removeByName("_cars")
    s = s & ", cars " & IIf(nCar < 0, "нет в снимке (WMS до 0.7)", CStr(nCar))
    TbPutSetting(3, snap)
    MgrRefresh = "OK:" & s
    Exit Function
EH:
    MgrRefresh = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' the period of the report: B2 «День» / «Неделя» / «Месяц», B3 a date inside it → start (B4) and end (B5, exclusive)
Private Sub Period(ByRef a As Double, ByRef b As Double, ByRef title As String)
    Dim sh As Object, kind As String, d As Double, c As Object
    sh = ThisComponent.Sheets.getByName(SH_REP)
    kind = LCase(Trim(sh.getCellByPosition(1, 1).getString()))
    c = sh.getCellByPosition(1, 2)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then d = Int(c.getValue()) Else d = Int(CDbl(Now()))
    Select Case kind
    Case "день", "сегодня"
        a = d
        b = d + 1
        title = "за " & Format(CDate(d), "DD.MM.YYYY")
    Case "неделя"
        a = d - (WeekDay(CDate(d), 2) - 1)
        b = a + 7
        title = "за неделю " & Format(CDate(a), "DD.MM.YYYY") & "–" & Format(CDate(b - 1), "DD.MM.YYYY")
    Case Else
        a = TbMonthStart(Year(CDate(d)), Month(CDate(d)))
        b = TbMonthStart(Year(CDate(d)), Month(CDate(d)) + 1)
        title = "за " & Format(CDate(a), "MM.YYYY")
    End Select
End Sub

Sub BtnMgrReport(Optional oEvent As Variant)
    TbMsg(MgrReport())
End Sub

' «Сегодня» / «Неделя» / «Месяц»: the period with today inside, from the latest snapshot
Sub BtnMgrToday(Optional oEvent As Variant)
    TbMsg(MgrQuick("День"))
End Sub

Sub BtnMgrWeek(Optional oEvent As Variant)
    TbMsg(MgrQuick("Неделя"))
End Sub

Sub BtnMgrMonth(Optional oEvent As Variant)
    TbMsg(MgrQuick("Месяц"))
End Sub

Function MgrQuick(sKind As String) As String
    Dim sh As Object, res As String
    sh = ThisComponent.Sheets.getByName(SH_REP)
    sh.getCellByPosition(1, 1).setString(sKind)
    sh.getCellByPosition(1, 2).setString("")
    res = MgrRefresh()
    If Left(res, 3) <> "OK:" Then
        MgrQuick = res
        Exit Function
    End If
    MgrQuick = MgrReport()
    ThisComponent.getCurrentController().setActiveSheet(sh)
End Function

' the indicators of the period: formulas over the loaded tables and the work log
Function MgrReport() As String
    Dim sh As Object, a As Double, b As Double, title As String, r As Long, n(5) As Long, t As Variant, i As Integer, k As Variant, last As Long, nCar As Long
    Dim w As String
    On Error GoTo EH
    If Not ThisComponent.Sheets.hasByName("_issues") Then
        MgrReport = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    t = Array("_orders", "_special", "_issues", "_returns", "_adjustments", "_stock")
    For i = 0 To 5
        If ThisComponent.Sheets.hasByName(t(i)) Then n(i) = TbLastRow(ThisComponent.Sheets.getByName(t(i))) + 1 Else n(i) = 1
    Next i
    nCar = -1
    If ThisComponent.Sheets.hasByName("_cars") Then nCar = TbLastRow(ThisComponent.Sheets.getByName("_cars")) + 1
    If nCar >= 0 And nCar < 2 Then nCar = 2
    Period(a, b, title)
    sh = ThisComponent.Sheets.getByName(SH_REP)
    sh.getCellByPosition(1, 3).setValue(a)
    sh.getCellByPosition(1, 4).setValue(b - 1)
    last = TbLastRow(sh)
    If last >= 6 Then
        sh.getCellRangeByPosition(0, 6, 3, last).clearContents(1023)
        sh.getCellRangeByPosition(0, 6, 3, last).CharWeight = 100
    End If
    r = 6
    Head(sh, r, "ОТЧЁТ СКЛАДА " & UCase(title))
    sh.getCellByPosition(0, r).setString("Снимок WMS: " & SnapLabel())
    r = r + 2
    Head(sh, r, "СКЛАДСКИЕ ОПЕРАЦИИ")
    RowF(sh, r, "Приходов по заказам (ЕИ)", "=COUNTIFS($_orders.N2:N" & n(0) & ";"">=" & a & """;$_orders.N2:N" & n(0) & ";""<" & b & """;$_orders.V2:V" & n(0) & ";""<>"")")
    RowF(sh, r, "Иных приходов (строк)", "=COUNTIFS($_special.H2:H" & n(1) & ";"">=" & a & """;$_special.H2:H" & n(1) & ";""<" & b & """;$_special.A2:A" & n(1) & ";"">0"")")
    RowF(sh, r, "Выдач", "=COUNTIFS($_issues.H2:H" & n(2) & ";"">=" & a & """;$_issues.H2:H" & n(2) & ";""<" & b & """;$_issues.R2:R" & n(2) & ";""Проведено*"")")
    RowF(sh, r, "Выдано, количество", "=SUMIFS($_issues.E2:E" & n(2) & ";$_issues.H2:H" & n(2) & ";"">=" & a & """;$_issues.H2:H" & n(2) & ";""<" & b _
        & """;$_issues.R2:R" & n(2) & ";""Проведено*"")")
    RowF(sh, r, "Возвратов", "=COUNTIFS($_returns.H2:H" & n(3) & ";"">=" & a & """;$_returns.H2:H" & n(3) & ";""<" & b & """;$_returns.N2:N" & n(3) & ";""Проведено*"")")
    For Each k In Array("Перемещение", "Списание", "Инвентаризация")
        RowF(sh, r, k & " (операций)", "=COUNTIFS($_adjustments.K2:K" & n(4) & ";"">=" & a & """;$_adjustments.K2:K" & n(4) & ";""<" & b _
            & """;$_adjustments.B2:B" & n(4) & ";""" & k & """;$_adjustments.Q2:Q" & n(4) & ";""Проведено*"")")
    Next k
    w = "$_adjustments.K2:K" & n(4) & ";"">=" & a & """;$_adjustments.K2:K" & n(4) & ";""<" & b & """;$_adjustments.B2:B" & n(4) & ";""Инвентаризация"";$_adjustments.Q2:Q" _
        & n(4) & ";""Проведено*"""
    RowF(sh, r, "Инвентаризационных расхождений (строк с разницей)", "=COUNTIFS(" & w & ";$_adjustments.P2:P" & n(4) & ";""<>0"")")
    RowF(sh, r, "Недостача по инвентаризации, количество", "=-SUMIFS($_adjustments.P2:P" & n(4) & ";" & w & ";$_adjustments.P2:P" & n(4) & ";""<0"")")
    RowF(sh, r, "Излишек по инвентаризации, количество", "=SUMIFS($_adjustments.P2:P" & n(4) & ";" & w & ";$_adjustments.P2:P" & n(4) & ";"">0"")")
    RowV(sh, r, "ЕИ в движении (разных ЕИ с операциями)", EIsMoved(a, b))
    r = r + 1
    Head(sh, r, "ЗАКАЗЫ И ДОКУМЕНТЫ")
    RowF(sh, r, "Просроченных заказов (на момент снимка)", "=COUNTIF($_orders.W2:W" & n(0) & ";""Просрочено"")+COUNTIF($_orders.W2:W" & n(0) _
        & ";""Частично получено / просрочено"")")
    RowF(sh, r, "Получено без документов (на момент снимка)", "=COUNTIF($_orders.W2:W" & n(0) & ";""Получено без документов"")")
    RowF(sh, r, "Приходов за период без № или даты документа", "=SUMPRODUCT(($_orders.N2:N" & n(0) & ">=" & a & ")*($_orders.N2:N" & n(0) & "<" & b & ")*($_orders.V2:V" _
        & n(0) & "<>"""")*((($_orders.C2:C" & n(0) & "="""")+($_orders.O2:O" & n(0) & "=""""))>0))")
    RowF(sh, r, "Возможные дубли поступлений (на момент снимка)", "=COUNTIF($_orders.AB2:AB" & n(0) & ";""?*"")")
    r = r + 1
    Head(sh, r, "ТРАНСПОРТ («Приход авто»)")
    If nCar < 1 Then
        RowS(sh, r, "Данные транспорта", "нет в снимке (снимок WMS до версии 0.7)")
    Else
        RowF(sh, r, "Машин приехало", "=COUNTIFS($_cars.E2:E" & nCar & ";"">=" & a & """;$_cars.E2:E" & nCar & ";""<" & b & """;$_cars.H2:H" & nCar & ";""<>Отменён"")")
        RowF(sh, r, "Машин уехало", "=COUNTIFS($_cars.F2:F" & nCar & ";"">=" & a & """;$_cars.F2:F" & nCar & ";""<" & b & """;$_cars.H2:H" & nCar & ";""Уехал"")")
        RowF(sh, r, "Среднее время машины на территории, мин", "=IFERROR(ROUND(AVERAGEIFS($_cars.G2:G" & nCar & ";$_cars.F2:F" & nCar & ";"">=" & a & """;$_cars.F2:F" _
            & nCar & ";""<" & b & """;$_cars.H2:H" & nCar & ";""Уехал"");1);""нет уехавших машин"")")
        RowF(sh, r, "Дольше " & LONG_MIN & " мин", "=COUNTIFS($_cars.F2:F" & nCar & ";"">=" & a & """;$_cars.F2:F" & nCar & ";""<" & b & """;$_cars.H2:H" & nCar _
            & ";""Уехал"";$_cars.G2:G" & nCar & ";"">" & LONG_MIN & """)")
        RowF(sh, r, "На территории (на момент снимка)", "=COUNTIF($_cars.H2:H" & nCar & ";""На территории"")")
    End If
    r = r + 1
    Head(sh, r, "РАБОТА, НЕ ВИДНАЯ В ДВИЖЕНИИ ТОВАРА (журнал работы, часы)")
    For Each k In Split(WORK_KINDS, "|")
        RowF(sh, r, k, "=SUMIFS($'" & SH_LOG & "'.D$2:D$10000;$'" & SH_LOG & "'.B$2:B$10000;""" & k & """;$'" & SH_LOG & "'.A$2:A$10000;"">=" & a _
            & """;$'" & SH_LOG & "'.A$2:A$10000;""<" & b & """)")
        sh.getCellByPosition(2, r - 1).setFormula("=COUNTIFS($'" & SH_LOG & "'.B$2:B$10000;""" & k & """;$'" & SH_LOG & "'.A$2:A$10000;"">=" & a _
            & """;$'" & SH_LOG & "'.A$2:A$10000;""<" & b & """)")
    Next k
    RowF(sh, r, "Всего часов", "=SUMIFS($'" & SH_LOG & "'.D$2:D$10000;$'" & SH_LOG & "'.A$2:A$10000;"">=" & a & """;$'" & SH_LOG & "'.A$2:A$10000;""<" & b & """)")
    sh.getCellByPosition(2, r - 1).setFormula("=COUNTIFS($'" & SH_LOG & "'.A$2:A$10000;"">=" & a & """;$'" & SH_LOG & "'.A$2:A$10000;""<" & b & """)")
    ThisComponent.calculateAll()
    MgrReport = "OK:отчёт " & title
    Exit Function
EH:
    MgrReport = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function SnapLabel() As String
    Dim snap As String
    snap = TbSetting(3)
    If snap = "" Then Exit Function
    SnapLabel = Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1)
    If Right(SnapLabel, 1) = "/" Then SnapLabel = Left(SnapLabel, Len(SnapLabel) - 1)
    SnapLabel = SnapLabel & " (" & Replace(TbManifestValue(snap, "created"), "T", " ") & ")"
End Function

Private Sub Head(sh As Object, ByRef r As Long, s As String)
    sh.getCellByPosition(0, r).setString(s)
    sh.getCellByPosition(0, r).CharWeight = 150
    r = r + 1
End Sub

Private Sub RowF(sh As Object, ByRef r As Long, sLabel As String, f As String)
    sh.getCellByPosition(0, r).setString(sLabel)
    sh.getCellByPosition(1, r).setFormula(f)
    r = r + 1
End Sub

Private Sub RowS(sh As Object, ByRef r As Long, sLabel As String, s As String)
    sh.getCellByPosition(0, r).setString(sLabel)
    sh.getCellByPosition(1, r).setString(s)
    r = r + 1
End Sub

Private Sub RowV(sh As Object, ByRef r As Long, sLabel As String, v As Double)
    sh.getCellByPosition(0, r).setString(sLabel)
    sh.getCellByPosition(1, r).setValue(v)
    r = r + 1
End Sub

' the number of an EI («ЕИ-00000012» → 12; 0 — not an EI)
Private Function EINo(v As Variant) As Long
    Dim s As String
    s = CStr(v)
    If Left(s, 3) = "ЕИ-" And Len(s) = 11 And IsNumeric(Mid(s, 4)) Then EINo = CLng(Mid(s, 4))
End Function

Private Function Rows0(sName As String, nCols As Integer) As Variant
    Dim sh As Object, last As Long
    Rows0 = Array()
    If Not ThisComponent.Sheets.hasByName(sName) Then Exit Function
    sh = ThisComponent.Sheets.getByName(sName)
    last = TbLastRow(sh)
    If last >= 1 Then Rows0 = sh.getCellRangeByPosition(0, 1, nCols - 1, last).getDataArray()
End Function

Private Function InPeriod(v As Variant, a As Double, b As Double) As Boolean
    If VarType(v) = 5 Or VarType(v) = 7 Then InPeriod = (CDbl(v) >= a And CDbl(v) < b)
End Function

' the different EIs with an operation in [a, b): receipts (orders, «Иной приход»), issues, returns, corrections
Private Function EIsMoved(a As Double, b As Double) As Long
    Dim seen() As Boolean, top As Long, o As Variant, sp As Variant, iss As Variant, rt As Variant, ad As Variant, i As Long, e As Long, n As Long
    top = TbLastRow(ThisComponent.Sheets.getByName("_stock")) + 2
    ReDim seen(top)
    o = Rows0("_orders", 22)
    sp = Rows0("_special", 14)
    iss = Rows0("_issues", 18)
    rt = Rows0("_returns", 17)
    ad = Rows0("_adjustments", 17)
    For i = 0 To UBound(o)
        If InPeriod(o(i)(13), a, b) Then Mark(seen, EINo(o(i)(21)), n)
    Next i
    For i = 0 To UBound(sp)
        If VarType(sp(i)(0)) = 5 And InPeriod(sp(i)(7), a, b) Then Mark(seen, EINo(sp(i)(13)), n)
    Next i
    For i = 0 To UBound(iss)
        If VarType(iss(i)(0)) = 5 And InPeriod(iss(i)(7), a, b) And Left(CStr(iss(i)(17)), 9) = "Проведено" Then Mark(seen, EINo(iss(i)(11)), n)
    Next i
    For i = 0 To UBound(rt)
        If VarType(rt(i)(0)) = 5 And InPeriod(rt(i)(7), a, b) And CStr(rt(i)(13)) <> "Удалено (сторно)" Then Mark(seen, EINo(rt(i)(2)), n)
    Next i
    For i = 0 To UBound(ad)
        If VarType(ad(i)(0)) = 5 And InPeriod(ad(i)(10), a, b) And CStr(ad(i)(16)) <> "Удалено (сторно)" Then Mark(seen, EINo(ad(i)(2)), n)
    Next i
    EIsMoved = n
End Function

Private Sub Mark(seen() As Boolean, e As Long, ByRef n As Long)
    If e < 1 Or e > UBound(seen) Then Exit Sub
    If Not seen(e) Then
        seen(e) = True
        n = n + 1
    End If
End Sub

' ================================================================ «Сформировать отчёт руководителю»

Sub BtnMgrDoc(Optional oEvent As Variant)
    TbMsg(MgrDoc(True))
End Sub

' a tidy document of the report of the sheet «Отчёт» (its period) and the entries of «Журнал работы» of that period:
' ../WMS_Reports/Отчёт_руководителю_<с>_<по>.odt (and .pdf when bPdf) — only the figures of the report, no service data
Function MgrDoc(bPdf As Boolean) As String
    Dim sh As Object, last As Long, d As Variant, i As Long, doc As Object, txt As Object, cur As Object, args(0) As New com.sun.star.beans.PropertyValue
    Dim sfa As Object, sDir As String, url As String, a As Double, b As Double, tbl As Object, rows() As Variant, k As Long, lg As Object, ld As Variant, j As Long
    Dim hdr As String, nRows As Long
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_REP)
    last = TbLastRow(sh)
    If last < 7 Or sh.getCellByPosition(1, 3).getType() <> com.sun.star.table.CellContentType.VALUE Then
        MgrDoc = "ERR:сначала сформируйте отчёт за период («Сегодня», «Неделя», «Месяц» или «Отчёт за период»)"
        Exit Function
    End If
    a = sh.getCellByPosition(1, 3).getValue()
    b = sh.getCellByPosition(1, 4).getValue() + 1
    d = sh.getCellRangeByPosition(0, 6, 2, last).getDataArray()
    args(0).Name = "Hidden"
    args(0).Value = True
    doc = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, args())
    txt = doc.getText()
    cur = txt.createTextCursor()
    Para(txt, cur, CStr(d(0)(0)), 15, True)
    Para(txt, cur, CStr(d(1)(0)) & "; сформирован " & Format(Now(), "DD.MM.YYYY HH:MM"), 9, False)
    ' the sections of the sheet: a bold header, then «показатель — значение» rows up to the next empty row
    i = 2
    Do While i <= UBound(d)
        If CStr(d(i)(0)) <> "" And sh.getCellByPosition(0, 6 + i).CharWeight >= 150 Then
            hdr = CStr(d(i)(0))
            nRows = 0
            Do While i + 1 + nRows <= UBound(d)
                If CStr(d(i + 1 + nRows)(0)) = "" Then Exit Do
                nRows = nRows + 1
            Loop
            Para(txt, cur, "", 6, False)
            Para(txt, cur, hdr, 12, True)
            If nRows > 0 Then
                tbl = doc.createInstance("com.sun.star.text.TextTable")
                tbl.initialize(nRows, 2)
                txt.insertTextContent(txt.getEnd(), tbl, False)
                For j = 0 To nRows - 1
                    tbl.getCellByName("A" & (j + 1)).setString(CStr(d(i + 1 + j)(0)))
                    tbl.getCellByName("B" & (j + 1)).setString(Val2Text(d(i + 1 + j)(1)) & IIf(Left(hdr, 6) = "РАБОТА" And CStr(d(i + 1 + j)(2)) <> "", _
                        " ч (записей: " & d(i + 1 + j)(2) & ")", ""))
                Next j
            End If
            i = i + 1 + nRows
        Else
            i = i + 1
        End If
    Loop
    ' the entries of the work log of the period
    lg = ThisComponent.Sheets.getByName(SH_LOG)
    last = TbLastRow(lg)
    k = 0
    If last >= 1 Then
        ld = lg.getCellRangeByPosition(0, 1, 4, last).getDataArray()
        ReDim rows(UBound(ld))
        For i = 0 To UBound(ld)
            If InPeriod(ld(i)(0), a, b) Then
                rows(k) = Array(Format(CDate(ld(i)(0)), "DD.MM.YYYY"), CStr(ld(i)(1)), CStr(ld(i)(2)), Val2Text(ld(i)(3)), CStr(ld(i)(4)))
                k = k + 1
            End If
        Next i
    End If
    Para(txt, cur, "", 6, False)
    Para(txt, cur, "ЗАПИСИ ЖУРНАЛА РАБОТЫ ЗА ПЕРИОД", 12, True)
    If k = 0 Then
        Para(txt, cur, "записей нет", 10, False)
    Else
        tbl = doc.createInstance("com.sun.star.text.TextTable")
        tbl.initialize(k + 1, 5)
        txt.insertTextContent(txt.getEnd(), tbl, False)
        tbl.getCellRangeByName("A1:E1").setDataArray(Array(Array("Дата", "Вид работы", "Что сделано", "Часы", "Комментарий")))
        For i = 0 To k - 1
            For j = 0 To 4
                tbl.getCellByPosition(j, i + 1).setString(rows(i)(j))
            Next j
        Next i
    End If
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    sDir = TbResolve("../WMS_Reports")
    If Not sfa.exists(sDir) Then sfa.createFolder(sDir)
    url = sDir & "Отчёт_руководителю_" & Format(CDate(a), "YYYY-MM-DD") & "_" & Format(CDate(b - 1), "YYYY-MM-DD") & ".odt"
    args(0).Name = "FilterName"
    args(0).Value = "writer8"
    doc.storeToURL(url, args())
    If bPdf Then
        args(0).Value = "writer_pdf_Export"
        doc.storeToURL(Left(url, Len(url) - 4) & ".pdf", args())
    End If
    doc.close(True)
    MgrDoc = "OK:" & ConvertFromURL(url) & IIf(bPdf, " (и PDF)", "")
    Exit Function
EH:
    MgrDoc = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Sub Para(txt As Object, cur As Object, s As String, h As Double, bBold As Boolean)
    cur.gotoEnd(False)
    cur.CharHeight = h
    cur.CharWeight = IIf(bBold, 150, 100)
    txt.insertString(cur, s, False)
    txt.insertControlCharacter(cur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
End Sub

Private Function Val2Text(v As Variant) As String
    If VarType(v) = 5 Or VarType(v) = 3 Or VarType(v) = 2 Then
        Val2Text = Replace(CStr(v), ".", ",")
    Else
        Val2Text = CStr(v)
    End If
End Function

' «В PDF»: the sheet «Отчёт» into ../WMS_Reports/Report_<period>.pdf
Sub BtnMgrPdf(Optional oEvent As Variant)
    TbMsg(MgrPdf())
End Sub

Function MgrPdf() As String
    Dim d As String, sfa As Object, url As String, args(1) As New com.sun.star.beans.PropertyValue, fd(0) As New com.sun.star.beans.PropertyValue
    Dim sh As Object
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbResolve("../WMS_Reports")
    If Not sfa.exists(d) Then sfa.createFolder(d)
    sh = ThisComponent.Sheets.getByName(SH_REP)
    url = d & "Report_" & Format(sh.getCellByPosition(1, 3).getValue(), "YYYY-MM-DD") & "_" & Format(sh.getCellByPosition(1, 4).getValue(), "YYYY-MM-DD") & ".pdf"
    ' exactly the rows of the report (A:D down to the last one)
    fd(0).Name = "Selection"
    fd(0).Value = sh.getCellRangeByPosition(0, 0, 3, TbLastRow(sh))
    args(0).Name = "FilterName"
    args(0).Value = "calc_pdf_Export"
    args(1).Name = "FilterData"
    args(1).Value = fd()
    ThisComponent.storeToURL(url, args())
    MgrPdf = "OK:" & ConvertFromURL(url)
    Exit Function
EH:
    MgrPdf = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Контроль дня» and «Что изменилось сегодня»

Sub BtnMgrControl(Optional oEvent As Variant)
    TbMsg(MgrControl(0))
End Sub

' one run before the end of the shift: the latest snapshot loaded, the checks of the day and what changed today on the
' sheet «Контроль дня»; dToday — the day (0 — today; the tests fix it). Nothing is corrected
Function MgrControl(dToday As Double) As String
    Dim res As String, snap As String, sh As Object, last As Long, r As Long, out() As Variant, n As Long, nProb As Long, nWarn As Long, i As Long
    Dim ch As Variant, st As String
    On Error GoTo EH
    If dToday = 0 Then dToday = Int(CDbl(Now()))
    res = MgrRefresh()
    If Left(res, 3) <> "OK:" Then
        MgrControl = res
        Exit Function
    End If
    snap = TbSetting(3)
    ReDim out(20)
    CtlChecks(snap, dToday, out, n)
    sh = ThisComponent.Sheets.getByName(SH_CTL)
    last = TbLastRow(sh)
    If last >= 0 Then
        sh.getCellRangeByPosition(0, 0, 3, last + 1).clearContents(1023)
        sh.getCellRangeByPosition(0, 0, 3, last + 1).CellBackColor = -1
        sh.getCellRangeByPosition(0, 0, 3, last + 1).CharWeight = 100
    End If
    sh.getCellByPosition(0, 0).setString("КОНТРОЛЬ ДНЯ " & Format(CDate(dToday), "DD.MM.YYYY") & " — снимок " & SnapLabel())
    sh.getCellByPosition(0, 0).CharWeight = 150
    CtlHead(sh, 2, Array("Статус", "Проверка", "Подробности", "Что делать"))
    For i = 0 To n - 1
        sh.getCellRangeByPosition(0, 3 + i, 3, 3 + i).setDataArray(Array(out(i)))
        st = CStr(out(i)(0))
        If st = "ПРОБЛЕМА" Then
            nProb = nProb + 1
            sh.getCellByPosition(0, 3 + i).CellBackColor = RGB(244, 204, 204)
        ElseIf st = "ВНИМАНИЕ" Then
            nWarn = nWarn + 1
            sh.getCellByPosition(0, 3 + i).CellBackColor = RGB(255, 229, 153)
        ElseIf st = "OK" Then
            sh.getCellByPosition(0, 3 + i).CellBackColor = RGB(226, 239, 218)
        End If
    Next i
    r = 3 + n + 1
    sh.getCellByPosition(0, r).setString("ЧТО ИЗМЕНИЛОСЬ СЕГОДНЯ (" & Format(CDate(dToday), "DD.MM.YYYY") & ")")
    sh.getCellByPosition(0, r).CharWeight = 150
    r = r + 1
    CtlHead(sh, r, Array("Что", "Сегодня"))
    r = r + 1
    ch = Changes(dToday, nProb + nWarn)
    sh.getCellRangeByPosition(0, r, 1, r + UBound(ch)).setDataArray(ch)
    sh.getCellRangeByPosition(0, 3, 3, r + UBound(ch)).IsTextWrapped = True
    ThisComponent.getCurrentController().setActiveSheet(sh)
    MgrControl = IIf(nProb > 0, "ERR:", "OK:") & "контроль дня: проблем " & nProb & ", внимание " & nWarn & " — лист «" & SH_CTL & "»"
    Exit Function
EH:
    MgrControl = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Sub CtlHead(sh As Object, r As Long, heads As Variant)
    Dim rng As Object
    rng = sh.getCellRangeByPosition(0, r, UBound(heads), r)
    rng.setDataArray(Array(heads))
    rng.CharWeight = 150
    rng.CellBackColor = RGB(231, 230, 230)
End Sub

Private Sub AddCtl(out() As Variant, ByRef n As Long, st As String, what As String, detail As String, todo As String)
    If n > UBound(out) Then Exit Sub
    out(n) = Array(st, what, detail, todo)
    n = n + 1
End Sub

Private Function DT(v As Variant) As String
    If VarType(v) = 5 Or VarType(v) = 7 Then DT = Format(CDate(v), "DD.MM HH:MM") Else DT = CStr(v)
End Function

' the checks of «Контроль дня»
Private Sub CtlChecks(snap As String, dToday As Double, out() As Variant, ByRef n As Long)
    Dim cars As Variant, i As Long, s As String, nOpen As Long, nOld As Long, o As Variant, sp As Variant, iss As Variant, rt As Variant, ad As Variant
    Dim c(4) As Long, st As Variant, nRev As Long, oldest As Double, revList As String, dr As Variant, fails As String, warns As String, nF As Long, nW As Long
    Dim nNoDoc As Long, noDoc As String, nOver As Long, over As String, sfa As Object, bd As String, lst As Variant, best As Double, dtm As Variant, x As Double
    Dim created As String, rev() As Double, stock As Variant, e As Long
    ' 1. vehicles on the territory
    If Not ThisComponent.Sheets.hasByName("_cars") Then
        AddCtl(out, n, "ИНФО", "Машины на территории", "в снимке нет данных «Приход авто» (WMS до версии 0.7)", "")
    ElseIf TbLastRow(ThisComponent.Sheets.getByName("_cars")) < 1 Then
        If TbHasFile(snap, "cars.csv") Then
            AddCtl(out, n, "OK", "Машины на территории", "нет (визитов в журнале нет)", "")
        Else
            AddCtl(out, n, "ИНФО", "Машины на территории", "в снимке нет данных «Приход авто» (WMS до версии 0.7)", "")
        End If
    Else
        cars = Rows0("_cars", 8)
        For i = 0 To UBound(cars)
            If CStr(cars(i)(7)) = "На территории" Then
                nOpen = nOpen + 1
                If VarType(cars(i)(4)) = 5 Then
                    If Int(cars(i)(4)) < dToday Then nOld = nOld + 1
                End If
                If nOpen <= 8 Then s = s & IIf(s <> "", "; ", "") & "№ " & cars(i)(0) & " «" & cars(i)(2) & "» (" & cars(i)(3) & ") с " & DT(cars(i)(4))
            End If
        Next i
        If nOpen = 0 Then
            AddCtl(out, n, "OK", "Машины на территории", "нет: все приехавшие машины уехали", "")
        Else
            AddCtl(out, n, IIf(nOld > 0, "ПРОБЛЕМА", "ВНИМАНИЕ"), "Машины на территории", nOpen & ": " & s & IIf(nOld > 0, " — с прошлых дней: " & nOld, ""), _
                "Уехавшую машину выделите на листе «Приход авто» и нажмите УЕХАЛ (время поправит «Исправить»)")
        End If
    End If
    ' 2. rows typed but not posted
    o = Rows0("_orders", 25)
    sp = Rows0("_special", 18)
    iss = Rows0("_issues", 18)
    rt = Rows0("_returns", 14)
    ad = Rows0("_adjustments", 17)
    For i = 0 To UBound(o)
        If CStr(o(i)(21)) = "" And CStr(o(i)(5)) <> "" And Left(CStr(o(i)(24)), 5) <> "КОПИЯ" Then c(0) = c(0) + 1
    Next i
    For i = 0 To UBound(sp)
        If VarType(sp(i)(0)) <> 5 And (CStr(sp(i)(1)) <> "" Or CStr(sp(i)(5)) <> "") And Left(CStr(sp(i)(17)), 5) <> "КОПИЯ" Then c(1) = c(1) + 1
    Next i
    For i = 0 To UBound(iss)
        If VarType(iss(i)(0)) <> 5 And (CStr(iss(i)(11)) <> "" Or CStr(iss(i)(4)) <> "") And Left(CStr(iss(i)(17)), 5) <> "КОПИЯ" Then c(2) = c(2) + 1
    Next i
    For i = 0 To UBound(rt)
        If VarType(rt(i)(0)) <> 5 And (CStr(rt(i)(1)) <> "" Or CStr(rt(i)(5)) <> "") And Left(CStr(rt(i)(13)), 5) <> "КОПИЯ" Then c(3) = c(3) + 1
    Next i
    For i = 0 To UBound(ad)
        If VarType(ad(i)(0)) <> 5 And CStr(ad(i)(1)) <> "" And Left(CStr(ad(i)(16)), 5) <> "КОПИЯ" Then c(4) = c(4) + 1
    Next i
    s = ""
    If c(0) > 0 Then s = s & "«Заказы»: " & c(0) & "; "
    If c(1) > 0 Then s = s & "«Иной приход»: " & c(1) & "; "
    If c(2) > 0 Then s = s & "«Выдачи»: " & c(2) & "; "
    If c(3) > 0 Then s = s & "«Возврат»: " & c(3) & "; "
    If c(4) > 0 Then s = s & "«Корректировки»: " & c(4) & "; "
    If s = "" Then
        AddCtl(out, n, "OK", "Непроведённые строки", "нет", "")
    Else
        AddCtl(out, n, "ВНИМАНИЕ", "Непроведённые строки", Left(s, Len(s) - 2), "Проведите строки кнопками их листа или очистите лишние (в WMS «Главная» показывает, где)")
    End If
    ' 3. «Иной приход» that waits for identification («Требует разбора»), the oldest first
    stock = Rows0("_stock", 10)
    ReDim rev(UBound(stock) + 2)
    For i = 0 To UBound(sp)
        If VarType(sp(i)(0)) = 5 And VarType(sp(i)(7)) = 5 Then
            e = EINo(sp(i)(13))
            If e > 0 And e <= UBound(rev) Then
                If rev(e) = 0 Or sp(i)(7) < rev(e) Then rev(e) = sp(i)(7)
            End If
        End If
    Next i
    oldest = 0
    For i = 0 To UBound(stock)
        If CStr(stock(i)(7)) = "Требует разбора" Then
            nRev = nRev + 1
            e = EINo(stock(i)(0))
            If e > 0 And e <= UBound(rev) Then
                If rev(e) > 0 And (oldest = 0 Or rev(e) < oldest) Then oldest = rev(e)
            End If
            If nRev <= 8 Then revList = revList & IIf(revList <> "", ", ", "") & stock(i)(0) & " «" & stock(i)(1) & "»"
        End If
    Next i
    If nRev = 0 Then
        AddCtl(out, n, "OK", "«Иной приход»: требует разбора", "нет", "")
    Else
        AddCtl(out, n, IIf(oldest > 0 And dToday - oldest > 3, "ПРОБЛЕМА", "ВНИМАНИЕ"), "«Иной приход»: требует разбора", nRev & ": " & revList _
            & IIf(oldest > 0, " — самый давний с " & Format(CDate(oldest), "DD.MM.YYYY") & " (" & (dToday - oldest) & " дн.)", ""), _
            "Уточните, что это за позиции (документ, от кого), и поправьте их на листе «Иной приход»")
    End If
    ' 4. the problems of the diagnosis (the checks of WMS_DOCTOR on the same snapshot)
    dr = DoctorCollect(snap)
    For i = 0 To UBound(dr)
        If dr(i)(0) = "FAIL" Then
            nF = nF + 1
            If nF <= 4 Then fails = fails & IIf(fails <> "", "; ", "") & dr(i)(1) & ": " & Left(dr(i)(2), 120)
        ElseIf dr(i)(0) = "WARN" Then
            nW = nW + 1
            If nW <= 6 Then warns = warns & IIf(warns <> "", "; ", "") & dr(i)(1)
        End If
    Next i
    If nF > 0 Then
        AddCtl(out, n, "ПРОБЛЕМА", "Диагностика (проверки WMS_DOCTOR)", "ошибок " & nF & ": " & fails & IIf(nW > 0, "; предупреждений " & nW & " (" & warns & ")", ""), _
            "Откройте «Диагностика» (WMS_DOCTOR) → «Проверить»: там подробности и план по каждой проверке")
    ElseIf nW > 0 Then
        AddCtl(out, n, "ВНИМАНИЕ", "Диагностика (проверки WMS_DOCTOR)", "ошибок нет; предупреждений " & nW & ": " & warns, "Подробности и план — «Диагностика» (WMS_DOCTOR) → «Проверить»")
    Else
        AddCtl(out, n, "OK", "Диагностика (проверки WMS_DOCTOR)", "ошибок и предупреждений нет", "")
    End If
    ' 5. receipts without the documents they need
    For i = 0 To UBound(o)
        If CStr(o(i)(21)) <> "" And Left(CStr(o(i)(24)), 5) <> "КОПИЯ" Then
            If CStr(o(i)(22)) = "Получено без документов" Or CStr(o(i)(2)) = "" Or CStr(o(i)(14)) = "" Then
                nNoDoc = nNoDoc + 1
                If nNoDoc <= 8 Then noDoc = noDoc & IIf(noDoc <> "", ", ", "") & o(i)(21) & IIf(CStr(o(i)(0)) <> "", " (заказ " & o(i)(0) & ")", "")
            End If
        End If
    Next i
    If nNoDoc = 0 Then
        AddCtl(out, n, "OK", "Приходы без документов", "нет", "")
    Else
        AddCtl(out, n, "ВНИМАНИЕ", "Приходы без документов", nNoDoc & ": " & noDoc, "Получите документы и впишите № и дату документа в строку заказа («Исправить»)")
    End If
    ' 6. overdue orders
    For i = 0 To UBound(o)
        If CStr(o(i)(22)) = "Просрочено" Or CStr(o(i)(22)) = "Частично получено / просрочено" Then
            nOver = nOver + 1
            If nOver <= 8 Then over = over & IIf(over <> "", "; ", "") & o(i)(0) & " «" & o(i)(1) & "»" & IIf(CStr(o(i)(11)) <> "", " (" & o(i)(11) & ")", "")
        End If
    Next i
    If nOver = 0 Then
        AddCtl(out, n, "OK", "Просроченные заказы", "нет", "")
    Else
        AddCtl(out, n, "ВНИМАНИЕ", "Просроченные заказы", nOver & ": " & over, "Уточните сроки у поставщиков и закупок")
    End If
    ' 7. a backup made today (the folder WMS_Backups of the WMS folder)
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    bd = TbWmsDir() & "WMS_Backups/"
    If Not sfa.exists(bd) Then
        AddCtl(out, n, "ВНИМАНИЕ", "Резервная копия за сегодня", "папки WMS_Backups не видно (" & ConvertFromURL(bd) & ")", _
            "В WMS: «Главная» → «Резервная копия»; проверьте, что инструменты лежат рядом с рабочей WMS")
    Else
        lst = sfa.getFolderContents(bd, False)
        For i = 0 To UBound(lst)
            dtm = sfa.getDateTimeModified(lst(i))
            x = CDbl(DateSerial(dtm.Year, dtm.Month, dtm.Day)) + CDbl(TimeSerial(dtm.Hours, dtm.Minutes, dtm.Seconds))
            If x > best Then best = x
        Next i
        If best = 0 Then
            AddCtl(out, n, "ПРОБЛЕМА", "Резервная копия за сегодня", "копий нет", "В WMS: «Главная» → «Резервная копия»")
        ElseIf Int(best) < dToday Then
            AddCtl(out, n, "ВНИМАНИЕ", "Резервная копия за сегодня", "последняя копия " & Format(CDate(best), "DD.MM.YYYY HH:MM") & " — сегодня копии не было", _
                "В WMS: «Главная» → «Резервная копия» (и перенесите копию на носитель)")
        Else
            AddCtl(out, n, "OK", "Резервная копия за сегодня", "есть: " & Format(CDate(best), "DD.MM.YYYY HH:MM"), "")
        End If
    End If
    ' 8. the snapshot: intact (checked when loaded) and made today
    created = TbManifestValue(snap, "created")
    If Left(created, 10) = Format(CDate(dToday), "YYYY-MM-DD") Then
        AddCtl(out, n, "OK", "Снимок для инструментов", "сегодняшний (" & Replace(created, "T", " ") & "), целостность проверена (SHA-256)", "")
    Else
        AddCtl(out, n, "ВНИМАНИЕ", "Снимок для инструментов", "снимок от " & Replace(created, "T", " ") & " — не сегодняшний: контроль показывает состояние на тот момент", _
            "В WMS: «Главная» → «Экспорт для инструментов», затем «Контроль дня» ещё раз")
    End If
End Sub

' «Что изменилось сегодня»: the rows (what, today)
Private Function Changes(dToday As Double, nProblems As Long) As Variant
    Dim o As Variant, sp As Variant, iss As Variant, rt As Variant, ad As Variant, cars As Variant, i As Long, out(8) As Variant, e As Long, top As Long
    Dim nRc As Long, qRc As Double, nOrd As Long, nSp As Long, nIss As Long, qIss As Double, who As String, nWho As Long, nRt As Long, qRt As Double
    Dim nMv As Long, nWo As Long, qWo As Double, nInv As Long, qInv As Double, nArr As Long, nDep As Long, nOpen As Long, delta() As Double, eis() As Long
    Dim k As Long, j As Long, best As Long, s As String, stock As Variant, b As Double, nm As String
    b = dToday + 1
    o = Rows0("_orders", 25)
    sp = Rows0("_special", 18)
    iss = Rows0("_issues", 18)
    rt = Rows0("_returns", 14)
    ad = Rows0("_adjustments", 17)
    stock = Rows0("_stock", 2)
    top = UBound(stock) + 2
    ReDim delta(top)
    For i = 0 To UBound(o)
        If CStr(o(i)(21)) <> "" And InPeriod(o(i)(13), dToday, b) And Left(CStr(o(i)(24)), 5) <> "КОПИЯ" And InStr(CStr(o(i)(22)), "сторно") = 0 Then
            nOrd = nOrd + 1
            nRc = nRc + 1
            If VarType(o(i)(5)) = 5 Then
                qRc = qRc + o(i)(5)
                AddDelta(delta, EINo(o(i)(21)), o(i)(5))
            End If
        End If
    Next i
    For i = 0 To UBound(sp)
        If VarType(sp(i)(0)) = 5 And InPeriod(sp(i)(7), dToday, b) And CStr(sp(i)(16)) <> "Удалено (сторно)" Then
            nSp = nSp + 1
            nRc = nRc + 1
            If VarType(sp(i)(5)) = 5 Then
                qRc = qRc + sp(i)(5)
                AddDelta(delta, EINo(sp(i)(13)), sp(i)(5))
            End If
        End If
    Next i
    For i = 0 To UBound(iss)
        If VarType(iss(i)(0)) = 5 And InPeriod(iss(i)(7), dToday, b) And Left(CStr(iss(i)(17)), 9) = "Проведено" Then
            nIss = nIss + 1
            If VarType(iss(i)(4)) = 5 Then
                qIss = qIss + iss(i)(4)
                AddDelta(delta, EINo(iss(i)(11)), -iss(i)(4))
            End If
            If InStr(1, "|" & who & "|", "|" & iss(i)(8) & "|", 0) = 0 Then
                who = who & "|" & iss(i)(8)
                nWho = nWho + 1
            End If
        End If
    Next i
    For i = 0 To UBound(rt)
        If VarType(rt(i)(0)) = 5 And InPeriod(rt(i)(7), dToday, b) And CStr(rt(i)(13)) <> "Удалено (сторно)" Then
            nRt = nRt + 1
            If VarType(rt(i)(5)) = 5 Then
                qRt = qRt + rt(i)(5)
                AddDelta(delta, EINo(rt(i)(2)), rt(i)(5))
            End If
        End If
    Next i
    For i = 0 To UBound(ad)
        If VarType(ad(i)(0)) = 5 And InPeriod(ad(i)(10), dToday, b) And CStr(ad(i)(16)) <> "Удалено (сторно)" Then
            Select Case CStr(ad(i)(1))
            Case "Перемещение"
                nMv = nMv + 1
            Case "Списание"
                nWo = nWo + 1
                If VarType(ad(i)(15)) = 5 Then qWo = qWo - ad(i)(15)
            Case "Инвентаризация"
                nInv = nInv + 1
                If VarType(ad(i)(15)) = 5 Then qInv = qInv + ad(i)(15)
            End Select
            If VarType(ad(i)(15)) = 5 Then AddDelta(delta, EINo(ad(i)(2)), ad(i)(15))
        End If
    Next i
    If ThisComponent.Sheets.hasByName("_cars") Then
        cars = Rows0("_cars", 8)
        For i = 0 To UBound(cars)
            If CStr(cars(i)(7)) <> "Отменён" And VarType(cars(i)(0)) = 5 Then
                If InPeriod(cars(i)(4), dToday, b) Then nArr = nArr + 1
                If InPeriod(cars(i)(5), dToday, b) Then nDep = nDep + 1
                If CStr(cars(i)(7)) = "На территории" Then nOpen = nOpen + 1
            End If
        Next i
    End If
    ' the largest changes of the balances (up to five)
    For k = 1 To 5
        best = 0
        For e = 1 To top
            If Abs(delta(e)) > 0.0000001 Then
                If best = 0 Then
                    best = e
                ElseIf Abs(delta(e)) > Abs(delta(best)) Then
                    best = e
                End If
            End If
        Next e
        If best = 0 Then Exit For
        nm = ""
        If best - 1 <= UBound(stock) Then nm = " «" & stock(best - 1)(1) & "»"
        s = s & IIf(s <> "", "; ", "") & "ЕИ-" & Format(best, "00000000") & nm & " " & IIf(delta(best) > 0, "+", "") & Val2Text(delta(best))
        delta(best) = 0
    Next k
    out(0) = Array("Пришло", "поступлений " & nRc & " (по заказам " & nOrd & ", иной приход " & nSp & "), количество " & Val2Text(qRc))
    out(1) = Array("Выдано", "выдач " & nIss & ", количество " & Val2Text(qIss) & ", получателей " & nWho)
    out(2) = Array("Возвращено", "возвратов " & nRt & ", количество " & Val2Text(qRt))
    out(3) = Array("Перемещено", "перемещений " & nMv)
    out(4) = Array("Списано", "списаний " & nWo & ", количество " & Val2Text(qWo))
    out(5) = Array("Скорректировано (инвентаризация)", "корректировок " & nInv & ", итог разницы " & Val2Text(qInv))
    If ThisComponent.Sheets.hasByName("_cars") Then
        out(6) = Array("Машины", "приехало " & nArr & ", уехало " & nDep & ", на территории сейчас " & nOpen)
    Else
        out(6) = Array("Машины", "нет данных «Приход авто» в снимке")
    End If
    out(7) = Array("Главные изменения остатков", IIf(s = "", "нет", s))
    out(8) = Array("Проблемы", IIf(nProblems = 0, "нет", "пунктов «Контроля дня» с проблемой или вниманием: " & nProblems & " (выше)"))
    Changes = out
End Function

Private Sub AddDelta(delta() As Double, e As Long, q As Double)
    If e >= 1 And e <= UBound(delta) Then delta(e) = delta(e) + q
End Sub

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function MgrProbe() As String
    MgrProbe = "OK"
End Function
