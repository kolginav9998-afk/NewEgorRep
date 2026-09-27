' TbAnalytics — WMS_ANALYTICS: большая сводка по снимку WMS (задание «FINAL WMS MARATHON», §6). «Обновить из снимка»
' загружает таблицы снимка в скрытые листы и строит лист «Сводка»: остатки, стоимость (по ценам заказов), категории,
' источники, места, приходы / выдачи / возвраты / списания / инвентаризация / перемещения по месяцам, просроченные
' заказы и заказы без документов, оборачиваемость, частота операций, получатели, нулевые и малые остатки, проблемные
' позиции. Показатели — формулы движка Calc над загруженными таблицами (без цикла Basic по строкам); списки
' (категории, места, получатели) — сортировкой LibreOffice.
Option Explicit

Private Const SH_DASH = "Сводка"
Private mR As Long
Private mSh As Object

Sub BtnAnRefresh(Optional oEvent As Variant)
    TbMsg(AnRefresh(0))
End Sub

' dToday: the date the months count back from (0 — today; the tests fix it)
Function AnRefresh(dToday As Double) As String
    Dim why As String, snap As String, t As Variant, i As Integer, n(5) As Long, last As Long, d0 As Double, m As Integer, a As Double, b As Double
    Dim ms As Variant
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        AnRefresh = "ERR:" & why
        Exit Function
    End If
    t = Array("stock", "orders", "special", "issues", "returns", "adjustments")
    For i = 0 To UBound(t)
        n(i) = TbLoadTable(snap, t(i) & ".csv", "_" & t(i), True)
    Next i
    TbPutSetting(3, snap)
    If dToday = 0 Then dToday = Int(CDbl(Now()))
    mSh = ThisComponent.Sheets.getByName(SH_DASH)
    last = TbLastRow(mSh)
    If last >= 0 Then mSh.getCellRangeByPosition(0, 0, 8, last + 1).clearContents(1023)
    mR = 0
    Put0("Сводка WMS по снимку " & Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1, 60) & " (LAST_SEQ " & TbManifestValue(snap, "last_seq") _
        & ", " & TbManifestValue(snap, "created") & ")", "", "")
    mR = mR + 1
    ' 1. stock
    Put0("ОСТАТКИ", "", "")
    PutF("ЕИ в реестре", "=COUNTA($_stock.A2:A" & (n(0) + 1) & ")")
    PutF("ЕИ с остатком", "=COUNTIF($_stock.E2:E" & (n(0) + 1) & ";"">0"")")
    PutF("Нулевые остатки", "=COUNTIFS($_stock.A2:A" & (n(0) + 1) & ";""<>"";$_stock.E2:E" & (n(0) + 1) & ";0)")
    PutF("Малые остатки (≤ порога «Настройки»)", "=COUNTIFS($_stock.E2:E" & (n(0) + 1) & ";"">0"";$_stock.E2:E" & (n(0) + 1) & ";""<=""&$'Настройки'.B6)")
    PutF("Требует разбора", "=COUNTIF($_stock.H2:H" & (n(0) + 1) & ";""Требует разбора"")")
    PutF("Стоимость остатка по ценам заказов (ЕИ заказов)", "=SUMPRODUCT(IFERROR(VALUE($_orders.X2:X" & (n(1) + 1) & ");0);IFERROR(VALUE($_orders.J2:J" _
        & (n(1) + 1) & ");0))")
    mR = mR + 1
    ' 2. sources (fixed list) and categories, places (sorted lists)
    Put0("ИСТОЧНИКИ", "ЕИ", "Остаток")
    For Each ms In Array("Поставщик", "Офис", "Производство", "Детали", "Старый склад", "Иной приход")
        PutF2(ms, "=COUNTIF($_stock.J2:J" & (n(0) + 1) & ";""" & ms & """)", "=SUMIF($_stock.J2:J" & (n(0) + 1) & ";""" & ms & """;$_stock.E2:E" & (n(0) + 1) & ")")
    Next ms
    mR = mR + 1
    Put0("КАТЕГОРИИ", "ЕИ", "Остаток")
    PutGroups("_stock", 6, n(0), 4, 30, -1)
    mR = mR + 1
    Put0("МЕСТА (больше всего ЕИ)", "ЕИ", "Остаток")
    PutGroups("_stock", 5, n(0), 4, 20, -1)
    mR = mR + 1
    ' 3. movements by month (the last 12 months)
    Put0("ДВИЖЕНИЯ ПО МЕСЯЦАМ", "Приходы", "Выдачи")
    mSh.getCellRangeByPosition(3, mR - 1, 7, mR - 1).setDataArray(Array(Array("Возвраты", "Списания", "Инвентаризация", "Перемещения", "Выдано, кол-во")))
    For m = 11 To 0 Step -1
        a = TbMonthStart(Year(CDate(dToday)), Month(CDate(dToday)) - m)
        b = TbMonthStart(Year(CDate(dToday)), Month(CDate(dToday)) - m + 1)
        mSh.getCellByPosition(0, mR).setString(Format(CDate(a), "MM.YYYY"))
        mSh.getCellByPosition(1, mR).setFormula("=COUNTIFS($_orders.N2:N" & (n(1) + 1) & ";"">=" & a & """;$_orders.N2:N" & (n(1) + 1) & ";""<" & b _
            & """;$_orders.V2:V" & (n(1) + 1) & ";""<>"")+COUNTIFS($_special.H2:H" & (n(2) + 1) & ";"">=" & a & """;$_special.H2:H" & (n(2) + 1) _
            & ";""<" & b & """;$_special.A2:A" & (n(2) + 1) & ";"">0"")")
        mSh.getCellByPosition(2, mR).setFormula("=COUNTIFS($_issues.H2:H" & (n(3) + 1) & ";"">=" & a & """;$_issues.H2:H" & (n(3) + 1) & ";""<" & b _
            & """;$_issues.R2:R" & (n(3) + 1) & ";""Проведено*"")")
        mSh.getCellByPosition(3, mR).setFormula("=COUNTIFS($_returns.H2:H" & (n(4) + 1) & ";"">=" & a & """;$_returns.H2:H" & (n(4) + 1) & ";""<" & b _
            & """;$_returns.N2:N" & (n(4) + 1) & ";""Проведено*"")")
        mSh.getCellByPosition(4, mR).setFormula(AdjCount(n(5), a, b, "Списание"))
        mSh.getCellByPosition(5, mR).setFormula(AdjCount(n(5), a, b, "Инвентаризация"))
        mSh.getCellByPosition(6, mR).setFormula(AdjCount(n(5), a, b, "Перемещение"))
        mSh.getCellByPosition(7, mR).setFormula("=SUMIFS($_issues.E2:E" & (n(3) + 1) & ";$_issues.H2:H" & (n(3) + 1) & ";"">=" & a & """;$_issues.H2:H" _
            & (n(3) + 1) & ";""<" & b & """;$_issues.R2:R" & (n(3) + 1) & ";""Проведено*"")")
        mR = mR + 1
    Next m
    mR = mR + 1
    ' 4. orders
    Put0("ЗАКАЗЫ", "", "")
    PutF("Просрочено", "=COUNTIF($_orders.W2:W" & (n(1) + 1) & ";""Просрочено"")+COUNTIF($_orders.W2:W" & (n(1) + 1) & ";""Частично получено / просрочено"")")
    PutF("Ожидается", "=COUNTIF($_orders.W2:W" & (n(1) + 1) & ";""Ожидается"")")
    PutF("Получено без документов", "=COUNTIF($_orders.W2:W" & (n(1) + 1) & ";""Получено без документов"")")
    PutF("Возможные дубли поступлений", "=COUNTIF($_orders.AB2:AB" & (n(1) + 1) & ";""?*"")")
    mR = mR + 1
    ' 5. turnover and frequency
    a = dToday - 90
    Put0("ОБОРАЧИВАЕМОСТЬ И ЧАСТОТА", "", "")
    PutF("Выдано за 90 дней, кол-во", "=SUMIFS($_issues.E2:E" & (n(3) + 1) & ";$_issues.H2:H" & (n(3) + 1) & ";"">=" & a & """;$_issues.R2:R" & (n(3) + 1) _
        & ";""Проведено*"")")
    PutF("Остаток всего, кол-во", "=SUM($_stock.E2:E" & (n(0) + 1) & ")")
    PutF("Оборачиваемость за 90 дней (выдано / остаток)", "=IFERROR(B" & (mR - 1) & "/B" & mR & ";0)")
    PutF("Операций за 30 дней в день (приходы, выдачи, возвраты, корректировки)", "=(COUNTIFS($_issues.H2:H" & (n(3) + 1) & ";"">=" & (dToday - 30) _
        & """)+COUNTIFS($_returns.H2:H" & (n(4) + 1) & ";"">=" & (dToday - 30) & """)+COUNTIFS($_special.H2:H" & (n(2) + 1) & ";"">=" & (dToday - 30) _
        & """)+COUNTIFS($_orders.N2:N" & (n(1) + 1) & ";"">=" & (dToday - 30) & """)+COUNTIFS($_adjustments.K2:K" & (n(5) + 1) & ";"">=" & (dToday - 30) & """))/30")
    mR = mR + 1
    ' 6. recipients
    Put0("ПОЛУЧАТЕЛИ (больше всего выдач)", "Выдач", "Выдано, кол-во")
    PutGroups("_issues", 8, n(3), 4, 15, 17)
    mR = mR + 1
    ' 7. problem positions
    Put0("ПРОБЛЕМНЫЕ ПОЗИЦИИ", "", "")
    PutF("Отрицательные остатки (должно быть 0)", "=COUNTIF($_stock.E2:E" & (n(0) + 1) & ";""<0"")")
    PutF("Строки-копии (КОПИЯ)", "=COUNTIF($_issues.R2:R" & (n(3) + 1) & ";""КОПИЯ*"")+COUNTIF($_orders.Y2:Y" & (n(1) + 1) & ";""КОПИЯ*"")+COUNTIF($_special.R2:R" _
        & (n(2) + 1) & ";""КОПИЯ*"")+COUNTIF($_returns.N2:N" & (n(4) + 1) & ";""КОПИЯ*"")+COUNTIF($_adjustments.Q2:Q" & (n(5) + 1) & ";""КОПИЯ*"")")
    PutF("Непроведённые выдачи (ЕИ без №)", "=COUNTIFS($_issues.L2:L" & (n(3) + 1) & ";""<>"";$_issues.A2:A" & (n(3) + 1) & ";"""")")
    ThisComponent.calculateAll()
    AnRefresh = "OK:сводка построена: ЕИ " & n(0) & ", заказов " & n(1) & ", иного прихода " & n(2) & ", выдач " & n(3) & ", возвратов " & n(4) _
        & ", корректировок " & n(5)
    Exit Function
EH:
    AnRefresh = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function AdjCount(nAdj As Long, a As Double, b As Double, sKind As String) As String
    AdjCount = "=COUNTIFS($_adjustments.K2:K" & (nAdj + 1) & ";"">=" & a & """;$_adjustments.K2:K" & (nAdj + 1) & ";""<" & b & """;$_adjustments.B2:B" _
        & (nAdj + 1) & ";""" & sKind & """;$_adjustments.Q2:Q" & (nAdj + 1) & ";""Проведено*"")"
End Function

Private Sub Put0(a As String, b As String, c As String)
    mSh.getCellByPosition(0, mR).setString(a)
    If b <> "" Then mSh.getCellByPosition(1, mR).setString(b)
    If c <> "" Then mSh.getCellByPosition(2, mR).setString(c)
    mSh.getCellByPosition(0, mR).CharWeight = 150
    mR = mR + 1
End Sub

Private Sub PutF(a As String, f As String)
    mSh.getCellByPosition(0, mR).setString(a)
    mSh.getCellByPosition(1, mR).setFormula(f)
    mR = mR + 1
End Sub

Private Sub PutF2(a As String, f1 As String, f2 As String)
    mSh.getCellByPosition(0, mR).setString(a)
    mSh.getCellByPosition(1, mR).setFormula(f1)
    mSh.getCellByPosition(2, mR).setFormula(f2)
    mR = mR + 1
End Sub

' the groups of a column of a loaded table (the values sorted by LibreOffice on a scratch sheet, the runs counted), the
' largest by count first: value, rows, sum of the column qtyCol (0-based); at most nTop. statusCol ≥ 0: only the rows whose
' status starts with «Проведено» (a cancelled issue is not counted)
Private Sub PutGroups(sTable As String, keyCol As Integer, nRows As Long, qtyCol As Integer, nTop As Integer, statusCol As Integer)
    Dim src As Object, scr As Object, d As Variant, i As Long, g() As Variant, k As Long, cur As String, cnt As Long, sm As Double, j As Long
    Dim tmp As Variant, fld(0) As New com.sun.star.table.TableSortField, desc As Variant
    If nRows < 1 Then Exit Sub
    src = ThisComponent.Sheets.getByName(sTable)
    If ThisComponent.Sheets.hasByName("_scratch") Then ThisComponent.Sheets.removeByName("_scratch")
    ThisComponent.Sheets.insertNewByName("_scratch", ThisComponent.Sheets.getCount())
    scr = ThisComponent.Sheets.getByName("_scratch")
    scr.IsVisible = False
    scr.getCellRangeByPosition(0, 0, 0, nRows - 1).setDataArray(src.getCellRangeByPosition(keyCol, 1, keyCol, nRows).getDataArray())
    scr.getCellRangeByPosition(1, 0, 1, nRows - 1).setDataArray(src.getCellRangeByPosition(qtyCol, 1, qtyCol, nRows).getDataArray())
    If statusCol >= 0 Then scr.getCellRangeByPosition(2, 0, 2, nRows - 1).setDataArray(src.getCellRangeByPosition(statusCol, 1, statusCol, nRows).getDataArray())
    fld(0).Field = 0
    fld(0).IsAscending = True
    desc = scr.getCellRangeByPosition(0, 0, 2, nRows - 1).createSortDescriptor()
    For i = 0 To UBound(desc)
        If desc(i).Name = "SortFields" Then desc(i).Value = fld
    Next i
    scr.getCellRangeByPosition(0, 0, 2, nRows - 1).sort(desc)
    d = scr.getCellRangeByPosition(0, 0, 2, nRows - 1).getDataArray()
    ReDim g(nRows)
    cur = Chr(1)
    For i = 0 To nRows - 1
        If statusCol >= 0 Then
            If Left(CStr(d(i)(2)), 9) <> "Проведено" Then GoTo NEXT_ROW
        End If
        If CStr(d(i)(0)) <> cur Then
            If cur <> Chr(1) And cur <> "" Then
                g(k) = Array(cur, cnt, sm)
                k = k + 1
            End If
            cur = CStr(d(i)(0))
            cnt = 0
            sm = 0
        End If
        cnt = cnt + 1
        If VarType(d(i)(1)) = 5 Then sm = sm + d(i)(1)
NEXT_ROW:
    Next i
    If cur <> Chr(1) And cur <> "" Then
        g(k) = Array(cur, cnt, sm)
        k = k + 1
    End If
    ThisComponent.Sheets.removeByName("_scratch")
    ' the largest groups first (a short insertion sort over the groups)
    For i = 1 To k - 1
        tmp = g(i)
        j = i - 1
        Do While j >= 0
            If g(j)(1) >= tmp(1) Then Exit Do
            g(j + 1) = g(j)
            j = j - 1
        Loop
        g(j + 1) = tmp
    Next i
    For i = 0 To IIf(k > nTop, nTop, k) - 1
        mSh.getCellByPosition(0, mR).setString(g(i)(0))
        mSh.getCellByPosition(1, mR).setValue(g(i)(1))
        mSh.getCellByPosition(2, mR).setValue(g(i)(2))
        mR = mR + 1
    Next i
    If k = 0 Then
        mSh.getCellByPosition(0, mR).setString("(нет данных)")
        mR = mR + 1
    End If
End Sub

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function AnProbe() As String
    AnProbe = "OK"
End Function
