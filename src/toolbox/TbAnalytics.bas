' TbAnalytics — WMS_ANALYTICS: большая сводка по снимку WMS (задание «FINAL WMS MARATHON», §6; M6 PRIME §3–§4).
' «Обновить из снимка» загружает таблицы снимка в скрытые листы и строит лист «Сводка»: остатки, стоимость (по ценам
' заказов), источники, категории, места, приходы / выдачи / возвраты / списания / инвентаризация / перемещения по месяцам,
' наиболее выдаваемые позиции, просроченные заказы и заказы без документов, оборачиваемость, частота операций,
' получатели, проблемные позиции, транспорт. Лист «Транспорт» — машины «Приход авто»: по дням, неделям, месяцам, часам
' приезда, поставщикам (дольше всего — сверху), повторяющиеся машины, машины на территории; время стоянки называется
' только при достаточном числе визитов, иначе — «Недостаточно данных». Лист «Графики» — диаграммы LibreOffice по этим
' таблицам. Показатели «Сводки» — формулы движка Calc над загруженными таблицами; группы — сортировкой LibreOffice.
' «Найти закономерности» запускает движок insights/wms_insights.py (python3; читает только снимок) и показывает листы
' «Найденные закономерности» (Критично / Внимание / Наблюдение: объект, что обнаружено, основание, уверенность, что
' сделать) и «Расход» (темп расхода, дни запаса, дата исчерпания или «Недостаточно данных»). Снимок WMS 0.6 (без машин)
' читается: транспорт — «нет данных в снимке».
Option Explicit

Private Const SH_DASH = "Сводка"
Private Const SH_CARS = "Транспорт"
Private Const SH_CHARTS = "Графики"
Private Const SH_FOUND = "Найденные закономерности"
Private Const SH_CONS = "Расход"
Private Const CAR_MIN_CLOSED = 5       ' closed visits before the times of the vehicles are stated
Private Const CAR_MIN_PEAK = 20        ' visits before the peak hours are stated
Private mR As Long
Private mSh As Object
' the tables of the charts: the row of the column headers and the number of data rows below it
Private mMonths As Long
Private mCat(1) As Long
Private mSrc(1) As Long
Private mPlace(1) As Long
Private mTop(1) As Long
Private mRcp(1) As Long
Private mCarDays(1) As Long
Private mCarSup(1) As Long

Sub BtnAnRefresh(Optional oEvent As Variant)
    TbMsg(AnRefresh(0))
End Sub

' dToday: the date the months count back from (0 — today; the tests fix it)
Function AnRefresh(dToday As Double) As String
    Dim why As String, snap As String, t As Variant, i As Integer, n(5) As Long, last As Long, m As Integer, a As Double, b As Double
    Dim ms As Variant, nCar As Long, carMsg As String
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
    nCar = TbLoadOptional(snap, "cars.csv", "_cars", True, TB_CAR_KEYS)
    TbPutSetting(3, snap)
    If dToday = 0 Then dToday = Int(CDbl(Now()))
    mSh = ThisComponent.Sheets.getByName(SH_DASH)
    last = TbLastRow(mSh)
    If last >= 0 Then mSh.getCellRangeByPosition(0, 0, 8, last + 1).clearContents(1023)
    mR = 0
    Put0("Сводка WMS по снимку " & SnapName(snap) & " (LAST_SEQ " & TbManifestValue(snap, "last_seq") & ", " & TbManifestValue(snap, "created") & ")", "", "")
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
    mSrc(0) = mR - 1
    mSrc(1) = 6
    For Each ms In Array("Поставщик", "Офис", "Производство", "Детали", "Старый склад", "Иной приход")
        PutF2(ms, "=COUNTIF($_stock.J2:J" & (n(0) + 1) & ";""" & ms & """)", "=SUMIF($_stock.J2:J" & (n(0) + 1) & ";""" & ms & """;$_stock.E2:E" & (n(0) + 1) & ")")
    Next ms
    mR = mR + 1
    Put0("КАТЕГОРИИ", "ЕИ", "Остаток")
    mCat(0) = mR - 1
    mCat(1) = PutGroups("_stock", 6, n(0), 4, 30, -1, -1, -1, 0)
    mR = mR + 1
    Put0("МЕСТА (больше всего ЕИ)", "ЕИ", "Остаток")
    mPlace(0) = mR - 1
    mPlace(1) = PutGroups("_stock", 5, n(0), 4, 20, -1, -1, -1, 0)
    mR = mR + 1
    ' 3. movements by month (the last 12 months)
    Put0("ДВИЖЕНИЯ ПО МЕСЯЦАМ", "Приходы", "Выдачи")
    mMonths = mR - 1
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
    ' the most issued positions: posted issues of the last 90 days, the EI with its name
    Put0("НАИБОЛЕЕ ВЫДАВАЕМЫЕ ПОЗИЦИИ (90 дней)", "Выдач", "Выдано, кол-во")
    mTop(0) = mR - 1
    mTop(1) = PutGroups("_issues", 11, n(3), 4, 10, 17, 2, 7, dToday - 89)
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
    mRcp(0) = mR - 1
    mRcp(1) = PutGroups("_issues", 8, n(3), 4, 15, 17, -1, -1, 0)
    mR = mR + 1
    ' 7. problem positions
    Put0("ПРОБЛЕМНЫЕ ПОЗИЦИИ", "", "")
    PutF("Отрицательные остатки (должно быть 0)", "=COUNTIF($_stock.E2:E" & (n(0) + 1) & ";""<0"")")
    PutF("Строки-копии (КОПИЯ)", "=COUNTIF($_issues.R2:R" & (n(3) + 1) & ";""КОПИЯ*"")+COUNTIF($_orders.Y2:Y" & (n(1) + 1) & ";""КОПИЯ*"")+COUNTIF($_special.R2:R" _
        & (n(2) + 1) & ";""КОПИЯ*"")+COUNTIF($_returns.N2:N" & (n(4) + 1) & ";""КОПИЯ*"")+COUNTIF($_adjustments.Q2:Q" & (n(5) + 1) & ";""КОПИЯ*"")")
    PutF("Непроведённые выдачи (ЕИ без №)", "=COUNTIFS($_issues.L2:L" & (n(3) + 1) & ";""<>"";$_issues.A2:A" & (n(3) + 1) & ";"""")")
    mR = mR + 1
    ' 8. transport: a short block here, the details on «Транспорт»
    Put0("ТРАНСПОРТ («Приход авто»)", "", "")
    If nCar < 0 Then
        PutS("Данные транспорта", "нет в снимке (снимок WMS до версии 0.7)")
    Else
        PutF("Машин за 30 дней", "=COUNTIFS($_cars.B2:B" & (nCar + 1) & ";"">=" & (dToday - 29) & """;$_cars.H2:H" & (nCar + 1) & ";""<>Отменён"";$_cars.A2:A" _
            & (nCar + 1) & ";"">0"")")
        PutF("На территории (на момент снимка)", "=COUNTIF($_cars.H2:H" & (nCar + 1) & ";""На территории"")")
        PutF("Средняя стоянка за 30 дней, мин", "=IFERROR(ROUND(AVERAGEIFS($_cars.G2:G" & (nCar + 1) & ";$_cars.H2:H" & (nCar + 1) & ";""Уехал"";$_cars.B2:B" _
            & (nCar + 1) & ";"">=" & (dToday - 29) & """);1);""нет закрытых визитов"")")
        PutS("Подробно", "лист «" & SH_CARS & "»; диаграммы — лист «" & SH_CHARTS & "»")
    End If
    ThisComponent.calculateAll()
    carMsg = CarSheet(snap, nCar, dToday)
    Charts(snap, nCar >= 0)
    ThisComponent.calculateAll()
    AnRefresh = "OK:сводка построена: ЕИ " & EICount(n(0)) & ", заказов " & n(1) & ", иного прихода " & n(2) & ", выдач " & n(3) & ", возвратов " _
        & n(4) & ", корректировок " & n(5) & "; " & carMsg
    Exit Function
EH:
    AnRefresh = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' the EIs of the registry (its rows are dense by the number of the EI: a transferred or migrated registry has empty rows)
Private Function EICount(nRows As Long) As Long
    Dim a As Variant, i As Long
    If nRows < 1 Then Exit Function
    a = ThisComponent.Sheets.getByName("_stock").getCellRangeByPosition(0, 1, 0, nRows).queryContentCells(com.sun.star.sheet.CellFlags.STRING) _
        .getRangeAddresses()
    For i = 0 To UBound(a)
        EICount = EICount + a(i).EndRow - a(i).StartRow + 1
    Next i
End Function

Private Function SnapName(snap As String) As String
    SnapName = Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1, 60)
    If Right(SnapName, 1) = "/" Then SnapName = Left(SnapName, Len(SnapName) - 1)
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

Private Sub PutS(a As String, b As String)
    mSh.getCellByPosition(0, mR).setString(a)
    mSh.getCellByPosition(1, mR).setString(b)
    mR = mR + 1
End Sub

Private Sub PutF2(a As String, f1 As String, f2 As String)
    mSh.getCellByPosition(0, mR).setString(a)
    mSh.getCellByPosition(1, mR).setFormula(f1)
    mSh.getCellByPosition(2, mR).setFormula(f2)
    mR = mR + 1
End Sub

' a scratch sheet (hidden) with nCols columns of data array d sorted by LibreOffice by the first key column (and the second
' when k2 ≥ 0): the sorted data array; the sheet is removed again
Private Function SortedRows(d As Variant, nCols As Integer, k2 As Integer) As Variant
    Dim scr As Object, rng As Object, desc As Variant, i As Integer
    Dim f1(0) As New com.sun.star.table.TableSortField, f2(1) As New com.sun.star.table.TableSortField
    If ThisComponent.Sheets.hasByName("_scratch") Then ThisComponent.Sheets.removeByName("_scratch")
    ThisComponent.Sheets.insertNewByName("_scratch", ThisComponent.Sheets.getCount())
    scr = ThisComponent.Sheets.getByName("_scratch")
    scr.IsVisible = False
    rng = scr.getCellRangeByPosition(0, 0, nCols - 1, UBound(d))
    rng.setDataArray(d)
    f1(0).Field = 0
    f1(0).IsAscending = True
    f2(0).Field = 0
    f2(0).IsAscending = True
    f2(1).Field = IIf(k2 >= 0, k2, 0)
    f2(1).IsAscending = True
    desc = rng.createSortDescriptor()
    For i = 0 To UBound(desc)
        If desc(i).Name = "SortFields" Then
            If k2 >= 0 Then desc(i).Value = f2 Else desc(i).Value = f1
        End If
    Next i
    rng.sort(desc)
    SortedRows = rng.getDataArray()
    ThisComponent.Sheets.removeByName("_scratch")
End Function

' the groups of a column of a loaded table (the values sorted by LibreOffice, the runs counted), the largest by count
' first: value, rows, sum of the column qtyCol (0-based); at most nTop; the number of groups written. statusCol ≥ 0: only
' the rows whose status starts with «Проведено» (a cancelled issue is not counted); labelCol ≥ 0: the value is shown with
' the label of its first row (the EI with its name); dateCol ≥ 0: only the rows dated dFrom or later
Private Function PutGroups(sTable As String, keyCol As Integer, nRows As Long, qtyCol As Integer, nTop As Integer, statusCol As Integer, _
    labelCol As Integer, dateCol As Integer, dFrom As Double) As Long
    Dim src As Object, d As Variant, i As Long, g() As Variant, k As Long, cur As String, cnt As Long, sm As Double, j As Long, rows() As Variant
    Dim tmp As Variant, lbl As String, dk As Variant, dq As Variant, ds As Variant, dl As Variant, dt As Variant
    If nRows < 1 Then
        mSh.getCellByPosition(0, mR).setString("(нет данных)")
        mR = mR + 1
        Exit Function
    End If
    src = ThisComponent.Sheets.getByName(sTable)
    ' the columns (a column not used is read as the key column and ignored below); every row built at once
    dk = src.getCellRangeByPosition(keyCol, 1, keyCol, nRows).getDataArray()
    dq = src.getCellRangeByPosition(qtyCol, 1, qtyCol, nRows).getDataArray()
    ds = src.getCellRangeByPosition(IIf(statusCol >= 0, statusCol, keyCol), 1, IIf(statusCol >= 0, statusCol, keyCol), nRows).getDataArray()
    dl = src.getCellRangeByPosition(IIf(labelCol >= 0, labelCol, keyCol), 1, IIf(labelCol >= 0, labelCol, keyCol), nRows).getDataArray()
    dt = src.getCellRangeByPosition(IIf(dateCol >= 0, dateCol, keyCol), 1, IIf(dateCol >= 0, dateCol, keyCol), nRows).getDataArray()
    ReDim rows(nRows - 1)
    For i = 0 To nRows - 1
        rows(i) = Array(dk(i)(0), dq(i)(0), ds(i)(0), dl(i)(0), dt(i)(0))
    Next i
    d = SortedRows(rows, 5, -1)
    ReDim g(nRows)
    cur = Chr(1)
    For i = 0 To nRows - 1
        If statusCol >= 0 Then
            If Left(CStr(d(i)(2)), 9) <> "Проведено" Then GoTo NEXT_ROW
        End If
        If dateCol >= 0 Then
            If VarType(d(i)(4)) <> 5 Then GoTo NEXT_ROW
            If d(i)(4) < dFrom Then GoTo NEXT_ROW
        End If
        If CStr(d(i)(0)) <> cur Then
            If cur <> Chr(1) And cur <> "" Then
                g(k) = Array(lbl, cnt, sm)
                k = k + 1
            End If
            cur = CStr(d(i)(0))
            lbl = cur
            If labelCol >= 0 Then
                If CStr(d(i)(3)) <> "" Then lbl = cur & " «" & d(i)(3) & "»"
            End If
            cnt = 0
            sm = 0
        End If
        cnt = cnt + 1
        If VarType(d(i)(1)) = 5 Then sm = sm + d(i)(1)
NEXT_ROW:
    Next i
    If cur <> Chr(1) And cur <> "" Then
        g(k) = Array(lbl, cnt, sm)
        k = k + 1
    End If
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
    PutGroups = IIf(k > nTop, nTop, k)
    If k = 0 Then
        mSh.getCellByPosition(0, mR).setString("(нет данных)")
        mR = mR + 1
    End If
End Function

' ================================================================ «Транспорт»

' the threshold of a long visit (Настройки B7, minutes; 120 by default — as on «Приход авто» of WMS)
Private Function LongMin() As Long
    LongMin = Val(TbSetting(6))
    If LongMin <= 0 Then LongMin = 120
End Function

Private Sub CarHead(sh As Object, r As Long, heads As Variant)
    Dim rng As Object
    rng = sh.getCellRangeByPosition(0, r, UBound(heads), r)
    rng.setDataArray(Array(heads))
    rng.CharWeight = 150
    rng.CellBackColor = RGB(231, 230, 230)
End Sub

Private Sub CarTitle(sh As Object, r As Long, s As String)
    sh.getCellByPosition(0, r).setString(s)
    sh.getCellByPosition(0, r).CharWeight = 150
    sh.getCellByPosition(0, r).CharHeight = 12
End Sub

' the median of the sorted values v(a..b)
Private Function MedianOf(v As Variant, a As Long, b As Long) As Double
    Dim n As Long
    n = b - a + 1
    If n Mod 2 = 1 Then MedianOf = v(a + n \ 2) Else MedianOf = (v(a + n \ 2 - 1) + v(a + n \ 2)) / 2
End Function

Private Function R1(x As Double) As Double
    R1 = Int(x * 10 + 0.5) / 10
End Function

' «Транспорт»: the vehicles of «Приход авто» of the snapshot (a Basic pass over the visits — the snapshot is not the working
' book; the groups by supplier and by vehicle sorted by LibreOffice); the short result for the message
Private Function CarSheet(snap As String, nCar As Long, dToday As Double) As String
    Dim sh As Object, last As Long, d As Variant, i As Long, nV As Long, nOpen As Long, nCanc As Long, nCl As Long, r As Long, st As String
    Dim durs() As Double, sumD As Double, maxD As Double, maxNo As String, nLong As Long, lim As Long, cnt(29) As Long, sd(29) As Double, nd(29) As Long
    Dim wk(11) As Long, mo(11) As Long, hr(23) As Long, k As Long, dd As Double, mon0 As Double, m0 As Long, created As Double, cs As String
    Dim sup() As Variant, veh() As Variant, s As Variant, g As Long, j As Long, cur As String, a As Long, nOpenRows As Long, openRows() As Variant
    Dim grp() As Variant, ng As Long, tmp As Variant, best As Long, bestH As Integer, nKeys As Long, nVeh As Long, sups As String, vals() As Double
    Dim peak As String, sorted As Variant, cl As Variant, tv() As Variant
    sh = ThisComponent.Sheets.getByName(SH_CARS)
    last = TbLastRow(sh)
    If last >= 0 Then sh.getCellRangeByPosition(0, 0, 7, last + 1).clearContents(1023)
    If last >= 0 Then sh.getCellRangeByPosition(0, 0, 7, last + 1).CellBackColor = -1
    mCarDays(1) = 0
    mCarSup(1) = 0
    If nCar < 0 Then
        CarTitle(sh, 0, "Транспорт («Приход авто»): в снимке нет данных — снимок WMS до версии 0.7")
        CarSheet = "транспорт: нет данных в снимке (WMS до 0.7)"
        Exit Function
    End If
    lim = LongMin()
    cs = TbManifestValue(snap, "created")
    If Len(cs) >= 19 Then created = CDbl(DateSerial(Val(Left(cs, 4)), Val(Mid(cs, 6, 2)), Val(Mid(cs, 9, 2)))) _
        + CDbl(TimeSerial(Val(Mid(cs, 12, 2)), Val(Mid(cs, 15, 2)), Val(Mid(cs, 18, 2)))) Else created = CDbl(Now())
    If nCar > 0 Then d = ThisComponent.Sheets.getByName("_cars").getCellRangeByPosition(0, 1, 13, nCar).getDataArray() Else d = Array()
    ReDim durs(UBound(d) + 1)
    ReDim sup(UBound(d) + 1)
    ReDim veh(UBound(d) + 1)
    ReDim openRows(UBound(d) + 1)
    mon0 = dToday - (WeekDay(CDate(dToday), 2) - 1) - 77
    m0 = Year(CDate(dToday)) * 12 + Month(CDate(dToday)) - 11
    For i = 0 To UBound(d)
        If VarType(d(i)(0)) <> 5 Or VarType(d(i)(4)) <> 5 Then GoTo NEXT_VISIT
        st = CStr(d(i)(7))
        If st = "Отменён" Then
            nCanc = nCanc + 1
            GoTo NEXT_VISIT
        End If
        dd = Int(d(i)(4))
        ' the groups: supplier (key, duration of a closed visit, display), vehicle (key, arrival, text, supplier)
        cl = ""
        If st = "Уехал" And VarType(d(i)(6)) = 5 Then cl = d(i)(6)
        sup(nKeys) = Array(UCase(Trim(CStr(d(i)(3)))), cl, CStr(d(i)(3)))
        veh(nKeys) = Array(IIf(CStr(d(i)(13)) <> "", CStr(d(i)(13)), UCase(CStr(d(i)(2)))), d(i)(4), CStr(d(i)(2)), CStr(d(i)(3)))
        If st = "Уехал" And VarType(d(i)(6)) = 5 Then
            durs(nCl) = d(i)(6)
            nCl = nCl + 1
            sumD = sumD + d(i)(6)
            If d(i)(6) > maxD Then
                maxD = d(i)(6)
                maxNo = CStr(d(i)(0))
            End If
            If d(i)(6) > lim Then nLong = nLong + 1
            k = dd - (dToday - 29)
            If k >= 0 And k <= 29 Then
                sd(k) = sd(k) + d(i)(6)
                nd(k) = nd(k) + 1
            End If
        ElseIf st = "На территории" Then
            openRows(nOpen) = Array(d(i)(0), CStr(d(i)(2)), CStr(d(i)(3)), Format(CDate(d(i)(4)), "DD.MM.YYYY HH:MM"), R1((created - d(i)(4)) * 24), _
                IIf(dd < Int(created), "с прошлых дней — отметьте выезд или исправьте время", ""))
            nOpen = nOpen + 1
        End If
        nKeys = nKeys + 1
        nV = nV + 1
        k = dd - (dToday - 29)
        If k >= 0 And k <= 29 Then cnt(k) = cnt(k) + 1
        k = Int((dd - mon0) / 7)
        If k >= 0 And k <= 11 Then wk(k) = wk(k) + 1
        k = Year(CDate(dd)) * 12 + Month(CDate(dd)) - m0
        If k >= 0 And k <= 11 Then mo(k) = mo(k) + 1
        hr(Hour(CDate(d(i)(4)))) = hr(Hour(CDate(d(i)(4)))) + 1
NEXT_VISIT:
    Next i
    r = 0
    CarTitle(sh, r, "Транспорт («Приход авто») по снимку " & SnapName(snap) & "; дата отчёта " & Format(CDate(dToday), "DD.MM.YYYY") _
        & "; длительный визит — дольше " & lim & " мин (Настройки)")
    r = 2
    CarHead(sh, r, Array("СВОДКА", "Значение"))
    r = r + 1
    ' the sorted durations: the median, and «Недостаточно данных» below CAR_MIN_CLOSED closed visits
    If nCl > 0 Then
        ReDim vals(nCl - 1)
        ReDim tv(nCl - 1)
        For i = 0 To nCl - 1
            tv(i) = Array(durs(i))
        Next i
        sorted = SortedRows(tv, 1, -1)
        For i = 0 To nCl - 1
            vals(i) = sorted(i)(0)
        Next i
    End If
    sh.getCellRangeByPosition(0, r, 1, r + 3).setDataArray(Array(Array("Визитов (без отменённых)", nV), Array("Отменено визитов", nCanc), _
        Array("На территории (на момент снимка)", nOpen), Array("Уехало (закрытых визитов)", nCl)))
    r = r + 4
    If nCl >= CAR_MIN_CLOSED Then
        sh.getCellRangeByPosition(0, r, 1, r + 3).setDataArray(Array(Array("Средняя стоянка, мин", R1(sumD / nCl)), Array("Медиана стоянки, мин", R1(MedianOf(vals, 0, nCl - 1))), _
            Array("Самая долгая стоянка, мин (визит № " & maxNo & ")", R1(maxD)), Array("Доля стоянок дольше " & lim & " мин", Format(nLong / nCl, "0%"))))
    Else
        sh.getCellRangeByPosition(0, r, 1, r + 3).setDataArray(Array(Array("Средняя стоянка, мин", "Недостаточно данных (закрытых визитов " & nCl & ", нужно ≥ " & CAR_MIN_CLOSED & ")"), _
            Array("Медиана стоянки, мин", "Недостаточно данных"), Array("Самая долгая стоянка, мин", IIf(nCl > 0, R1(maxD), "нет закрытых визитов")), _
            Array("Доля стоянок дольше " & lim & " мин", "Недостаточно данных")))
    End If
    r = r + 4
    ' the peak two hours of the arrivals
    If nV >= CAR_MIN_PEAK Then
        best = -1
        For i = 0 To 22
            If hr(i) + hr(i + 1) > best Or (hr(i) + hr(i + 1) = best And hr(i) > hr(bestH)) Then
                best = hr(i) + hr(i + 1)
                bestH = i
            End If
        Next i
        peak = Format(bestH, "00") & ":00–" & Format(bestH + 2, "00") & ":00: " & best & " из " & nV & " (" & Format(best / nV, "0%") & ")"
    Else
        peak = "Недостаточно данных (визитов " & nV & ", нужно ≥ " & CAR_MIN_PEAK & ")"
    End If
    sh.getCellRangeByPosition(0, r, 1, r).setDataArray(Array(Array("Пик приездов (два часа)", peak)))
    r = r + 1
    ' the suppliers (by the mean of the closed visits, the longest first) and the repeated vehicles
    If nKeys > 0 Then
        ReDim Preserve sup(nKeys - 1)
        sorted = SortedRows(sup, 3, 1)
        ReDim grp(nKeys)
        i = 0
        Do While i < nKeys
            j = i
            Do While j + 1 < nKeys
                If sorted(j + 1)(0) <> sorted(i)(0) Then Exit Do
                j = j + 1
            Loop
            ' rows i..j: the closed durations come first, ascending (the empty ones are sorted last)
            a = i
            Do While a <= j
                If VarType(sorted(a)(1)) <> 5 Then Exit Do
                a = a + 1
            Loop
            ReDim vals(j - i)
            sumD = 0
            nLong = 0
            For g = i To a - 1
                vals(g - i) = sorted(g)(1)
                sumD = sumD + sorted(g)(1)
                If sorted(g)(1) > lim Then nLong = nLong + 1
            Next g
            If a > i Then
                grp(ng) = Array(sorted(i)(2), R1(sumD / (a - i)), R1(MedianOf(vals, 0, a - i - 1)), R1(vals(a - i - 1)), j - i + 1, Format(nLong / (a - i), "0%"), _
                    IIf(a - i < 3, "мало закрытых визитов (" & (a - i) & ")", ""))
            Else
                grp(ng) = Array(sorted(i)(2), "", "", "", j - i + 1, "", "нет закрытых визитов")
            End If
            ng = ng + 1
            i = j + 1
        Loop
        ' the longest mean first; the suppliers without closed visits last
        For i = 1 To ng - 1
            tmp = grp(i)
            j = i - 1
            Do While j >= 0
                If SupBefore(grp(j), tmp) Then Exit Do
                grp(j + 1) = grp(j)
                j = j - 1
            Loop
            grp(j + 1) = tmp
        Next i
        ReDim Preserve veh(nKeys - 1)
        sorted = SortedRows(veh, 4, 1)
        nVeh = 0
        i = 0
        Do While i < nKeys
            j = i
            Do While j + 1 < nKeys
                If sorted(j + 1)(0) <> sorted(i)(0) Then Exit Do
                j = j + 1
            Loop
            nVeh = nVeh + 1
            If j > i Then
                sups = ""
                For g = i To j
                    If InStr(1, "|" & sups & "|", "|" & sorted(g)(3) & "|", 0) = 0 Then sups = sups & IIf(sups <> "", "|", "") & sorted(g)(3)
                Next g
                sorted(i) = Array(sorted(j)(2), j - i + 1, Format(CDate(sorted(j)(1)), "DD.MM.YYYY"), Replace(sups, "|", ", "), "rep")
            End If
            i = j + 1
        Loop
    End If
    sh.getCellRangeByPosition(0, r, 1, r + 1).setDataArray(Array(Array("Разных машин", nVeh), Array("Разных поставщиков", ng)))
    r = r + 3
    ' on the territory
    CarTitle(sh, r, "НА ТЕРРИТОРИИ (на момент снимка " & Replace(cs, "T", " ") & ")")
    r = r + 1
    CarHead(sh, r, Array("№ визита", "Машина", "Поставщик", "Приезд", "Часов на территории", "Примечание"))
    r = r + 1
    If nOpen = 0 Then
        sh.getCellByPosition(0, r).setString("нет")
        r = r + 1
    Else
        ReDim Preserve openRows(nOpen - 1)
        sh.getCellRangeByPosition(0, r, 5, r + nOpen - 1).setDataArray(openRows)
        r = r + nOpen
    End If
    r = r + 1
    ' by day, week, month, hour
    CarTitle(sh, r, "ПО ДНЯМ (30 дней)")
    r = r + 1
    CarHead(sh, r, Array("День", "Машин", "Средняя стоянка, мин"))
    mCarDays(0) = r
    mCarDays(1) = 30
    r = r + 1
    For k = 0 To 29
        sh.getCellByPosition(0, r).setString(Format(CDate(dToday - 29 + k), "DD.MM"))
        sh.getCellByPosition(1, r).setValue(cnt(k))
        If nd(k) > 0 Then sh.getCellByPosition(2, r).setValue(R1(sd(k) / nd(k)))
        r = r + 1
    Next k
    r = r + 1
    CarTitle(sh, r, "ПО НЕДЕЛЯМ (12 недель) И МЕСЯЦАМ (12 месяцев)")
    r = r + 1
    CarHead(sh, r, Array("Неделя с", "Машин", "", "Месяц", "Машин"))
    r = r + 1
    For k = 0 To 11
        sh.getCellByPosition(0, r).setString(Format(CDate(mon0 + 7 * k), "DD.MM.YYYY"))
        sh.getCellByPosition(1, r).setValue(wk(k))
        g = m0 + k - 1
        sh.getCellByPosition(3, r).setString(Format((g Mod 12) + 1, "00") & "." & (g \ 12))
        sh.getCellByPosition(4, r).setValue(mo(k))
        r = r + 1
    Next k
    r = r + 1
    CarTitle(sh, r, "ЧАСЫ ПРИЕЗДА")
    r = r + 1
    CarHead(sh, r, Array("Час", "Приездов"))
    r = r + 1
    For k = 0 To 23
        sh.getCellByPosition(0, r).setString(Format(k, "00") & ":00")
        sh.getCellByPosition(1, r).setValue(hr(k))
        r = r + 1
    Next k
    r = r + 1
    CarTitle(sh, r, "ПОСТАВЩИКИ (дольше всего — сверху)")
    r = r + 1
    CarHead(sh, r, Array("Поставщик", "Средняя стоянка, мин", "Медиана, мин", "Самая долгая, мин", "Визитов", "Доля дольше " & lim & " мин", "Примечание"))
    mCarSup(0) = r
    r = r + 1
    For i = 0 To ng - 1
        sh.getCellRangeByPosition(0, r, 6, r).setDataArray(Array(grp(i)))
        r = r + 1
    Next i
    ' a chart shows the suppliers with closed visits only
    For i = 0 To ng - 1
        If VarType(grp(i)(1)) = 5 Then mCarSup(1) = mCarSup(1) + 1
    Next i
    r = r + 1
    CarTitle(sh, r, "ПОВТОРЯЮЩИЕСЯ МАШИНЫ (2 визита и больше)")
    r = r + 1
    CarHead(sh, r, Array("Машина", "Визитов", "Последний визит", "Поставщики"))
    r = r + 1
    g = 0
    If nKeys > 0 Then
        ' the repeated vehicles, the most visits first
        ReDim grp(nKeys)
        For i = 0 To nKeys - 1
            If UBound(sorted(i)) = 4 Then
                If sorted(i)(4) = "rep" Then
                    grp(g) = Array(sorted(i)(0), sorted(i)(1), sorted(i)(2), sorted(i)(3))
                    g = g + 1
                End If
            End If
        Next i
        For i = 1 To g - 1
            tmp = grp(i)
            j = i - 1
            Do While j >= 0
                If grp(j)(1) >= tmp(1) Then Exit Do
                grp(j + 1) = grp(j)
                j = j - 1
            Loop
            grp(j + 1) = tmp
        Next i
        For i = 0 To g - 1
            sh.getCellRangeByPosition(0, r, 3, r).setDataArray(Array(grp(i)))
            r = r + 1
        Next i
    End If
    If g = 0 Then sh.getCellByPosition(0, r).setString("нет")
    CarSheet = "транспорт: визитов " & nV & ", на территории " & nOpen & ", уехало " & nCl
End Function

' the order of the suppliers: a greater mean first, the ones without closed visits last
Private Function SupBefore(x As Variant, y As Variant) As Boolean
    If VarType(x(1)) <> 5 Then
        SupBefore = (VarType(y(1)) <> 5)
    ElseIf VarType(y(1)) <> 5 Then
        SupBefore = True
    Else
        SupBefore = (x(1) >= y(1))
    End If
End Function

' ================================================================ «Графики»

' the charts of LibreOffice over the tables of «Сводка» and «Транспорт» (rebuilt with them): the movements by month, the EIs
' by category, source and place, the most issued positions, the recipients, the vehicles by day and by supplier
Private Sub Charts(snap As String, bCars As Boolean)
    Dim g As Object, nm As Variant, i As Integer, cs As Object
    g = ThisComponent.Sheets.getByName(SH_CHARTS)
    nm = g.getCharts().getElementNames()
    For i = 0 To UBound(nm)
        g.getCharts().removeByName(nm(i))
    Next i
    g.getCellRangeByPosition(0, 0, 3, 0).clearContents(1023)
    g.getCellByPosition(0, 0).setString("Графики по снимку " & SnapName(snap) & " — таблицы на листах «" & SH_DASH & "» и «" & SH_CARS & "»")
    g.getCellByPosition(0, 0).CharWeight = 150
    cs = ThisComponent.Sheets.getByName(SH_CARS)
    AddChart(g, "months", "Движения по месяцам: приходы, выдачи, возвраты, списания", mSh, mMonths, 4, 12, False, True, 0)
    AddChart(g, "categories", "ЕИ по категориям", mSh, mCat(0), 1, mCat(1), True, False, 1)
    AddChart(g, "sources", "ЕИ по источникам", mSh, mSrc(0), 1, mSrc(1), True, False, 2)
    AddChart(g, "places", "ЕИ по местам (больше всего)", mSh, mPlace(0), 1, mPlace(1), True, False, 3)
    AddChart(g, "top_issued", "Наиболее выдаваемые позиции за 90 дней (выдач)", mSh, mTop(0), 1, mTop(1), True, False, 4)
    AddChart(g, "recipients", "Получатели (выдач)", mSh, mRcp(0), 1, mRcp(1), True, False, 5)
    If bCars Then
        AddChart(g, "cars_days", "Машины по дням (30 дней)", cs, mCarDays(0), 1, mCarDays(1), False, False, 6)
        AddChart(g, "cars_suppliers", "Средняя стоянка машин по поставщикам, мин", cs, mCarSup(0), 1, mCarSup(1), True, False, 7)
    End If
End Sub

' one chart in slot k (two in a row): the columns 0..nSer of the rows hdr..hdr+nRows of sheet src (the first row — the names
' of the series, the first column — the categories); bBars: horizontal bars (long names read along the axis)
Private Sub AddChart(g As Object, sName As String, sTitle As String, src As Object, hdr As Long, nSer As Integer, nRows As Long, bBars As Boolean, _
    bLegend As Boolean, k As Integer)
    Dim rect As New com.sun.star.awt.Rectangle, addr(0) As New com.sun.star.table.CellRangeAddress, ch As Object
    If nRows < 1 Then Exit Sub
    rect.X = 300 + (k Mod 2) * 16600
    rect.Y = 1000 + (k \ 2) * 9600
    rect.Width = 16300
    rect.Height = 9300
    addr(0) = src.getCellRangeByPosition(0, hdr, nSer, hdr + nRows).getRangeAddress()
    g.getCharts().addNewByName(sName, rect, addr(), True, True)
    ch = g.getCharts().getByName(sName).getEmbeddedObject()
    ch.lockControllers()
    ch.setDiagram(ch.createInstance("com.sun.star.chart.BarDiagram"))
    ch.getDiagram().Vertical = bBars
    ch.HasMainTitle = True
    ch.getTitle().String = sTitle
    ch.HasLegend = bLegend
    ch.unlockControllers()
End Sub

' ================================================================ «Найти закономерности»

Sub BtnAnInsights(Optional oEvent As Variant)
    TbMsg(AnInsights(""))
End Sub

' the pattern engine insights/wms_insights.py on the snapshot (python3; it only reads the snapshot): its files go to
' ../WMS_Reports/INSIGHTS_<time>/, the patterns to «Найденные закономерности», the consumption to «Расход». sToday
' «ГГГГ-ММ-ДД» or "" (the date of the snapshot)
Function AnInsights(sToday As String) As String
    Dim why As String, snap As String, sOut As String, sLog As String, res As String, sfa As Object, args As String, n As Long
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        AnInsights = "ERR:" & why
        Exit Function
    End If
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(TbResolve("../WMS_Reports")) Then sfa.createFolder(TbResolve("../WMS_Reports"))
    sOut = TbResolve("../WMS_Reports") & "INSIGHTS_" & Format(Now(), "YYYYMMDD-HHMMSS") & "/"
    sLog = TbFolder() & "insights_last.txt"
    args = "--snapshot '" & ConvertFromURL(snap) & "' --out '" & ConvertFromURL(sOut) & "' --wms '" & ConvertFromURL(TbWmsDir()) & "' --long-min " & LongMin()
    If sToday <> "" Then args = args & " --today " & sToday
    why = TbRunPython("insights/wms_insights.py", args, sLog)
    If why <> "" Then
        AnInsights = "ERR:" & why
        Exit Function
    End If
    res = Trim(Replace(TbReadText(sLog), Chr(10), " "))
    If Not sfa.exists(sOut & "insights.csv") Then
        AnInsights = "ERR:" & res
        Exit Function
    End If
    n = FoundSheet(snap, sOut)
    ConsSheet(sOut)
    AnInsights = "OK:" & res & "; отчёт " & ConvertFromURL(sOut) & "INSIGHTS_REPORT.md"
    Exit Function
EH:
    AnInsights = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' «Найденные закономерности»: the level (coloured), the topic, the object, what was found, why (the basis), the confidence,
' what to do; the number of patterns
Private Function FoundSheet(snap As String, sOut As String) As Long
    Dim sh As Object, src As Object, n As Long, d As Variant, last As Long, i As Long, nc(2) As Long, rows() As Variant, col As Long
    n = TbLoadCsv(sOut & "insights.csv", "_insights", True, "222222222")
    sh = ThisComponent.Sheets.getByName(SH_FOUND)
    last = TbLastRow(sh)
    If last >= 0 Then
        sh.getCellRangeByPosition(0, 0, 6, last + 1).clearContents(1023)
        sh.getCellRangeByPosition(0, 0, 6, last + 1).CellBackColor = -1
    End If
    CarTitle(sh, 0, "Найденные закономерности — снимок " & SnapName(snap) & ", " & Format(Now(), "DD.MM.YYYY HH:MM"))
    CarHead(sh, 3, Array("Уровень", "Тема", "Объект", "Что обнаружено", "Почему (основание)", "Уверенность", "Что сделать"))
    If n > 0 Then
        src = ThisComponent.Sheets.getByName("_insights")
        d = src.getCellRangeByPosition(0, 1, 6, n).getDataArray()
        sh.getCellRangeByPosition(0, 4, 6, 4 + n - 1).setDataArray(d)
        For i = 0 To n - 1
            Select Case CStr(d(i)(0))
            Case "Критично"
                col = RGB(244, 204, 204)
                nc(0) = nc(0) + 1
            Case "Внимание"
                col = RGB(255, 229, 153)
                nc(1) = nc(1) + 1
            Case Else
                col = RGB(221, 235, 247)
                nc(2) = nc(2) + 1
            End Select
            sh.getCellByPosition(0, 4 + i).CellBackColor = col
        Next i
        sh.getCellRangeByPosition(0, 4, 6, 4 + n - 1).IsTextWrapped = True
    Else
        sh.getCellByPosition(0, 4).setString("Закономерностей не найдено: истории мало или отклонений нет")
    End If
    sh.getCellByPosition(0, 1).setString("Критично: " & nc(0) & " · Внимание: " & nc(1) & " · Наблюдение: " & nc(2) _
        & " — расчёт детерминированный, по снимку; прогноз только при достаточной истории. Все файлы: " & ConvertFromURL(sOut))
    FoundSheet = n
End Function

' «Расход»: the rate of consumption of every EI, the days of stock, the date of depletion — or «Недостаточно данных»
Private Sub ConsSheet(sOut As String)
    Dim sh As Object, n As Long, last As Long
    n = TbLoadCsv(sOut & "consumption.csv", "_consumption", True, "22222222222222")
    sh = ThisComponent.Sheets.getByName(SH_CONS)
    last = TbLastRow(sh)
    If last >= 0 Then sh.getCellRangeByPosition(0, 0, 13, last + 1).clearContents(1023)
    CarTitle(sh, 0, "Расход по ЕИ: темп за 7 / 30 / 90 дней, дни запаса и дата исчерпания при продолжении наблюдаемого расхода")
    CarHead(sh, 2, Array("ЕИ", "Наименование", "Артикул", "Остаток", "Выдано 7 дн.", "Выдано 30 дн.", "Выдано 90 дн.", "Выдач 90 дн.", "Расход в день (30 дн.)", _
        "Расход в день (90 дн.)", "Дней запаса", "Исчерпание", "Уверенность", "Основание"))
    If n > 0 Then sh.getCellRangeByPosition(0, 3, 13, 3 + n - 1).setDataArray(ThisComponent.Sheets.getByName("_consumption").getCellRangeByPosition(0, 1, 13, n).getDataArray())
End Sub

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function AnProbe() As String
    AnProbe = "OK"
End Function
