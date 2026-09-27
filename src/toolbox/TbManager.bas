' TbManager — WMS_MANAGER: отчёт руководителю (задание «FINAL WMS MARATHON», §6). Автоматические показатели по снимку
' WMS за день / неделю / месяц и небольшой ручной «Журнал работы» кладовщика (разгрузки и погрузки, работа погрузчиком,
' запросы бухгалтерии, взаимодействие с закупками, проблемные поставки, прочие работы). «Сформировать отчёт» считает
' период, «В PDF» сохраняет отчёт рядом с инструментом (../WMS_Reports).
Option Explicit

Private Const SH_REP = "Отчёт"
Private Const SH_LOG = "Журнал работы"
Public Const WORK_KINDS = "Разгрузка / погрузка|Работа погрузчиком|Запрос бухгалтерии|Закупки|Проблемная поставка|Прочие работы"

Sub BtnMgrRefresh(Optional oEvent As Variant)
    TbMsg(MgrRefresh())
End Sub

Function MgrRefresh() As String
    Dim why As String, snap As String, t As Variant, i As Integer, s As String
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        MgrRefresh = "ERR:" & why
        Exit Function
    End If
    t = Array("orders", "special", "issues", "returns", "adjustments")
    For i = 0 To UBound(t)
        s = s & IIf(s <> "", ", ", "") & t(i) & " " & TbLoadTable(snap, t(i) & ".csv", "_" & t(i), True)
    Next i
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
    Case "день"
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

' the indicators of the period: formulas over the loaded tables and the work log
Function MgrReport() As String
    Dim sh As Object, a As Double, b As Double, title As String, r As Long, n(4) As Long, t As Variant, i As Integer, k As Variant, last As Long
    On Error GoTo EH
    If Not ThisComponent.Sheets.hasByName("_issues") Then
        MgrReport = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    t = Array("_orders", "_special", "_issues", "_returns", "_adjustments")
    For i = 0 To 4
        n(i) = TbLastRow(ThisComponent.Sheets.getByName(t(i))) + 1
    Next i
    Period(a, b, title)
    sh = ThisComponent.Sheets.getByName(SH_REP)
    sh.getCellByPosition(1, 3).setValue(a)
    sh.getCellByPosition(1, 4).setValue(b - 1)
    last = TbLastRow(sh)
    If last >= 6 Then sh.getCellRangeByPosition(0, 6, 3, last).clearContents(1023)
    r = 6
    sh.getCellByPosition(0, r).setString("ОТЧЁТ СКЛАДА " & UCase(title))
    sh.getCellByPosition(0, r).CharWeight = 150
    r = r + 1
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
    RowF(sh, r, "Просроченных заказов (на момент снимка)", "=COUNTIF($_orders.W2:W" & n(0) & ";""Просрочено"")+COUNTIF($_orders.W2:W" & n(0) _
        & ";""Частично получено / просрочено"")")
    r = r + 1
    sh.getCellByPosition(0, r).setString("ЖУРНАЛ РАБОТЫ (часы)")
    sh.getCellByPosition(0, r).CharWeight = 150
    r = r + 1
    For Each k In Split(WORK_KINDS, "|")
        RowF(sh, r, k, "=SUMIFS($'" & SH_LOG & "'.D$2:D$10000;$'" & SH_LOG & "'.B$2:B$10000;""" & k & """;$'" & SH_LOG & "'.A$2:A$10000;"">=" & a _
            & """;$'" & SH_LOG & "'.A$2:A$10000;""<" & b & """)")
    Next k
    RowF(sh, r, "Всего часов", "=SUMIFS($'" & SH_LOG & "'.D$2:D$10000;$'" & SH_LOG & "'.A$2:A$10000;"">=" & a & """;$'" & SH_LOG & "'.A$2:A$10000;""<" & b & """)")
    ThisComponent.calculateAll()
    MgrReport = "OK:отчёт " & title
    Exit Function
EH:
    MgrReport = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Sub RowF(sh As Object, ByRef r As Long, sLabel As String, f As String)
    sh.getCellByPosition(0, r).setString(sLabel)
    sh.getCellByPosition(1, r).setFormula(f)
    r = r + 1
End Sub

' «В PDF»: the sheet «Отчёт» into ../WMS_Reports/Отчёт_<period>.pdf
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
    fd(0).Name = "Selection"
    fd(0).Value = sh
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

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function MgrProbe() As String
    MgrProbe = "OK"
End Function
