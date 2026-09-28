' WmsCar — «Приход авто» (задание M6 PRIME, §1): журнал машин на территории склада. Это не складской приход: остатки
' он не меняет никогда. Кладовщик вводит в панели над таблицей только «Марка / госномер» и «Поставщик» и нажимает
' ПРИЕХАЛ; при выезде выделяет строку машины и нажимает УЕХАЛ. № визита, дату, время, длительность, статус, день недели,
' месяц, № за день и контроль пишет WMS. Каждое действие — одна операция WmsCore.ApplyOperation (журнал, откат,
' восстановление после сбоя, защита от Ctrl+Z):
'  - CAR_ARRIVE — новый визит № NEXT_CAR; отказ, если эта же машина (тот же ключ госномера) ещё на территории;
'  - CAR_DEPART — выезд открытого визита: время выезда, длительность фиксируется, статус «Уехал»;
'  - CAR_FIX — исправление госномера, поставщика или времени (ошибка ввода, забытое нажатие УЕХАЛ);
'  - CAR_CANCEL — отмена ошибочной записи (визит остаётся в истории со статусом «Отменён», в показателях не участвует).
' Адресация плотная: визит n — строка _CAR n и строка CAR_FIRST + n − 1 листа (строки пишет только WMS, они защищены,
' вставка строк на листе запрещена). Открытый визит машины и список машин на территории находятся формулами движка
' Calc (_IDX, TEXTJOIN), без просмотра истории циклом Basic.
Option Explicit

' test seam (TEST books only): the clock of the operations (0 — the real time)
Global gCarNowSet As Boolean
Global gCarNow As Double

' ================================================================ sheets, clock, texts

Function CarsSheet() As Object
    CarsSheet = gDoc.Sheets.getByName(SH_CARS)
End Function

Function CarSheet() As Object
    CarSheet = gDoc.Sheets.getByName(SH_CAR)
End Function

' "" when the sheets of the vehicle journal exist and have the expected layout (checked at every start, fail-closed)
Function CarSheetsProblem() As String
    Dim sh As Object, c As Object, k As Variant, x As Double
    If Not gDoc.Sheets.hasByName(SH_CARS) Then
        CarSheetsProblem = "нет листа «" & SH_CARS & "»"
        Exit Function
    End If
    If Not gDoc.Sheets.hasByName(SH_CAR) Then
        CarSheetsProblem = "нет листа «" & SH_CAR & "»"
        Exit Function
    End If
    sh = CarsSheet()
    If sh.getCellByPosition(CC_NO, CAR_HEAD_ROW).getString() <> "№ визита" Or sh.getCellByPosition(CC_PLATE, CAR_HEAD_ROW).getString() <> "Марка / госномер" _
        Or sh.getCellByPosition(CC_STATUS, CAR_HEAD_ROW).getString() <> "Статус" Or sh.getCellByPosition(CC_CTL, CAR_HEAD_ROW).getString() <> "Контроль" _
        Or sh.getCellByPosition(CP_PLATE_COL, CP_IN_ROW - 1).getString() <> "Марка / госномер" Or sh.getCellByPosition(CP_SUP_COL, CP_IN_ROW - 1).getString() <> "Поставщик" Then
        CarSheetsProblem = "лист «" & SH_CARS & "»: панель или заголовок таблицы не совпадают с ожидаемыми"
        Exit Function
    End If
    sh = CarSheet()
    If sh.getCellByPosition(CR_NO, 0).getString() <> "№ визита" Or sh.getCellByPosition(CR_SKEY, 0).getString() <> "Ключ поставщика" _
        Or sh.getCellByPosition(CD_PLATE, 0).getString() <> "Словарь: машина" Or sh.getCellByPosition(CD_SKEY, 0).getString() <> "Словарь: ключ поставщика" Then
        CarSheetsProblem = "служебный лист " & SH_CAR & " повреждён (заголовок)"
        Exit Function
    End If
    For Each k In Array(CD_NPLATE_COL, CD_NSUP_COL)
        c = sh.getCellByPosition(k, 0)
        If c.getType() <> com.sun.star.table.CellContentType.VALUE Then
            CarSheetsProblem = "служебный лист " & SH_CAR & " повреждён: нет размера словаря в строке 1"
            Exit Function
        End If
        x = c.getValue()
        If x < 0 Or x <> Int(x) Or x >= MAX_SHEET_ROW Then
            CarSheetsProblem = SH_CAR & ": размер словаря должен быть целым числом ≥ 0, найдено «" & c.getString() & "»"
            Exit Function
        End If
    Next k
    CarSheetsProblem = ""
End Function

' the time of the operations: the real clock, whole seconds (a TEST book may set its own)
Function CarNow() As Double
    Dim t As Double
    If gCarNowSet Then
        CarNow = gCarNow
    Else
        t = CDbl(Now())
        CarNow = Int(t * 86400 + 0.5) / 86400
    End If
End Function

' spaces trimmed and collapsed, tabs and line breaks become spaces
Function CleanText(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, sp As Boolean
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If Asc(ch) < 32 Or ch = Chr(160) Then ch = " "
        If ch = " " Then
            sp = True
        Else
            If sp And out <> "" Then out = out & " "
            sp = False
            out = out & ch
        End If
    Next i
    CleanText = out
End Function

Private Function IsPlateLetter(ch As String) As Boolean
    IsPlateLetter = (InStr(1, "АВЕКМНОРСТУХ", ch, 0) > 0)
End Function

Private Function IsDigit(ch As String) As Boolean
    IsDigit = (ch >= "0" And ch <= "9" And Len(ch) = 1)
End Function

' upper case, Latin letters that look like the Cyrillic letters of a plate (А В Е К М Н О Р С Т У Х) replaced by them, Ё → Е,
' only letters and digits kept
Function SearchKey(ByVal s As String) As String
    Dim i As Long, ch As String, p As Integer, out As String, a As Long
    s = UCase(s)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        p = InStr(1, "ABEKMHOPCTYX", ch, 0)
        If p > 0 Then ch = Mid("АВЕКМНОРСТУХ", p, 1)
        If ch = "Ё" Then ch = "Е"
        a = Asc(ch)          ' the Unicode value (AscW is VBA only)
        If (a >= 48 And a <= 57) Or (a >= 65 And a <= 90) Or (a >= 1040 And a <= 1071) Then out = out & ch
    Next i
    SearchKey = out
End Function

' The key of a vehicle: SearchKey of the text; when it holds a Russian plate (letter, three digits, two letters) — that
' plate (the last one in the text), without the region: «Газель А123ВС 77», «а 123 вс» and «A123BC» are one vehicle.
' Another text (a foreign plate, a trailer, a name) is its whole SearchKey. The text the user typed is kept as it is.
Function PlateKey(ByVal s As String) As String
    Dim t As String, i As Long
    t = SearchKey(s)
    For i = Len(t) - 5 To 1 Step -1
        If IsPlateLetter(Mid(t, i, 1)) And IsDigit(Mid(t, i + 1, 1)) And IsDigit(Mid(t, i + 2, 1)) And IsDigit(Mid(t, i + 3, 1)) _
            And IsPlateLetter(Mid(t, i + 4, 1)) And IsPlateLetter(Mid(t, i + 5, 1)) Then
            PlateKey = Mid(t, i, 6)
            Exit Function
        End If
    Next i
    PlateKey = t
End Function

' the key of a supplier: CleanText, upper case, Ё → Е, quotes dropped («Петрович», "Петрович" and петрович are one supplier)
Function SupplierKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase(CleanText(s))
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "Ё" Then ch = "Е"
        If InStr(1, "«»""'„“”", ch, 0) = 0 Then out = out & ch
    Next i
    SupplierKey = CleanText(out)
End Function

Function HhMm(t As Double) As String
    HhMm = Format(CDate(t), "HH:MM")
End Function

Function DayText(t As Double) As String
    DayText = Format(CDate(t), "DD.MM.YYYY")
End Function

' «1 ч 25 мин», «25 мин»
Function DurText(days As Double) As String
    Dim m As Long
    m = Int(days * 1440 + 0.5)
    If m < 0 Then m = 0
    If m >= 60 Then DurText = (m \ 60) & " ч " & (m Mod 60) & " мин" Else DurText = m & " мин"
End Function

Function WeekdayText(d As Double) As String
    WeekdayText = Array("Воскресенье", "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота")(WeekDay(CDate(d)) - 1)
End Function

' the threshold of a long stay (minutes): the setting of the panel when it holds a positive number, else CAR_LONG_MIN
Function LongMinutes() As Long
    Dim c As Object
    c = CarsSheet().getCellByPosition(CP_THR_COL, CP_VAL_ROW)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then
        If c.getValue() >= 1 And c.getValue() <= 100000 Then
            LongMinutes = CLng(c.getValue())
            Exit Function
        End If
    End If
    LongMinutes = CAR_LONG_MIN
End Function

' ================================================================ visits (dense addressing, the key is checked)

Function CarRow(n As Long) As Variant
    CarRow = CarSheet().getCellRangeByPosition(0, n, CR_LAST, n).getDataArray()(0)
End Function

Function CarRegistered(n As Long, d As Variant) As Boolean
    If n < 1 Or n >= SysNum(SK_NEXT_CAR) Or CAR_FIRST + n - 1 > MAX_SHEET_ROW Then Exit Function
    d = CarRow(n)
    If VarType(d(CR_NO)) <> 5 Then Exit Function
    CarRegistered = (d(CR_NO) = n)
End Function

' the row of visit n on the sheet: A..N
Function VisitCells(n As Long) As Variant
    VisitCells = CarsSheet().getCellRangeByPosition(0, CAR_FIRST + n - 1, CC_LAST, CAR_FIRST + n - 1).getDataArray()(0)
End Function

' the visit shown on row r of the sheet (0-based); 0 when the row is not a visit of this WMS
Function VisitAtRow(r As Long) As Long
    Dim n As Long, d As Variant, v As Variant
    n = r - CAR_FIRST + 1
    If Not CarRegistered(n, d) Then Exit Function
    v = CarsSheet().getCellByPosition(CC_NO, r)
    If v.getType() <> com.sun.star.table.CellContentType.VALUE Then Exit Function
    If v.getValue() <> n Then Exit Function
    VisitAtRow = n
End Function

' «А123ВС (Петрович, приехала 27.09 в 09:15)»
Function VisitText(n As Long) As String
    Dim v As Variant, s As String
    v = VisitCells(n)
    s = "«" & v(CC_PLATE) & "» (" & v(CC_SUP)
    If VarType(v(CC_ARR)) = 5 Then s = s & ", приезд " & Format(CDate(v(CC_ARR)), "DD.MM") & " в " & HhMm(CDbl(v(CC_ARR)))
    VisitText = s & ")"
End Function

' the open visits (ascending №): one TEXTJOIN of the Calc engine over _CAR; the MATCH of the open keys one by one when
' the formula could not be used
Function OpenVisits() As Variant
    Dim last As Long, s As String, ok As Boolean, a As Variant, out() As Long, i As Long, n As Long, k As Long, rng As String
    last = CLng(SysNum(SK_NEXT_CAR)) - 1
    If last < 1 Then
        OpenVisits = Array()
        Exit Function
    End If
    rng = "$'" & SH_CAR & "'.$C$2:$C$" & (last + 1)
    s = WmsOrders.ScratchEval("TEXTJOIN("";"";1;IF(" & rng & "=""" & CR_OPEN & """;ROW(" & rng & ")-1;""""))", ok)
    If ok Then
        If s = "" Then
            OpenVisits = Array()
            Exit Function
        End If
        a = Split(s, ";")
        ReDim out(UBound(a))
        For i = 0 To UBound(a)
            out(i) = CLng(a(i))
        Next i
        OpenVisits = out
        Exit Function
    End If
    ReDim out(15)
    n = 0
    Do
        k = WmsOrders.FindCarOpen("?*", n)
        If k < 1 Or k <= n Then Exit Do
        If k > last Then Exit Do
        If i > UBound(out) Then ReDim Preserve out(2 * i + 1)
        out(i) = k
        i = i + 1
        n = k
    Loop
    If i = 0 Then
        OpenVisits = Array()
    Else
        ReDim Preserve out(i - 1)
        OpenVisits = out
    End If
End Function

' the number of visits on the territory now (the COUNTIF of the Calc engine)
Function CarsOnSite() As Long
    CarsOnSite = WmsOrders.CountCarOpen("?*")
End Function

' the vehicles on the territory since an earlier day (УЕХАЛ forgotten?)
Function StaleOpen() As Long
    Dim o As Variant, i As Long, d As Variant, today As Double, k As Long
    o = OpenVisits()
    today = Int(CarNow())
    For i = 0 To UBound(o)
        d = CarRow(CLng(o(i)))
        If VarType(d(CR_ARR)) = 5 Then
            If Int(d(CR_ARR)) < today Then k = k + 1
        End If
    Next i
    StaleOpen = k
End Function

' the number of open order positions of the supplier on «Заказы» (the same rule as the hint of the panel); -1 unknown
Private Function OpenOrderPositions(sSup As String) As Long
    Dim q As String, f As String, s As String, ok As Boolean, l As String, w As String, st As Variant, i As Integer
    q = """=" & Replace(Replace(Replace(Replace(sSup, "~", "~~"), "*", "~*"), "?", "~?"), """", """""") & """"
    l = "$'" & SH_ORDERS & "'.$L$2:$L$1048576"
    w = "$'" & SH_ORDERS & "'.$W$2:$W$1048576"
    ' the open statuses (with the date statuses of an earlier core, until the refresh at the opening turns them)
    st = Array(OS_WAITING, OS_OVERDUE, OS_PARTIAL, OS_PARTIAL_OVERDUE)
    For i = 0 To UBound(st)
        f = f & IIf(f <> "", "+", "") & "COUNTIFS(" & l & ";" & q & ";" & w & ";""" & st(i) & """)"
    Next i
    s = WmsOrders.ScratchEval("(" & f & ")", ok)
    If ok And IsNumeric(s) Then OpenOrderPositions = CLng(s) Else OpenOrderPositions = -1
End Function

' "" when visit № n may be issued now: _CAR and the table are behind the counter
Private Function NewCarProblem(n As Long) As String
    Dim r As Long
    r = CAR_FIRST + n - 1
    If n < 1 Or r > MAX_SHEET_ROW Then
        NewCarProblem = "номера визитов исчерпаны (NEXT_CAR " & n & ")"
    ElseIf CarSheet().getCellByPosition(CR_NO, n).getType() <> com.sun.star.table.CellContentType.EMPTY _
        Or CarSheet().getCellByPosition(CR_KEY, n).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        NewCarProblem = "служебная таблица " & SH_CAR & " не согласована со счётчиком NEXT_CAR " & n & " — нужна самопроверка"
    ElseIf CarsSheet().getCellByPosition(CC_NO, r).getType() <> com.sun.star.table.CellContentType.EMPTY _
        Or CarsSheet().getCellByPosition(CC_PLATE, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        NewCarProblem = "строка " & (r + 1) & " листа «" & SH_CARS & "» уже занята — таблица не согласована со счётчиком NEXT_CAR " & n _
            & ". Нужна проверка ответственным"
    Else
        NewCarProblem = ""
    End If
End Function

' a new entry of a dictionary of _CAR when the key is not there yet (the lists of the input cells of the panel)
Private Sub PlanDictionary(key As String, sText As String, bPlate As Boolean)
    Dim k As Long, szCol As Integer, cText As Integer, cKey As Integer, sh As Object
    sh = CarSheet()
    If bPlate Then
        If WmsOrders.FindCarPlate(key) > 0 Then Exit Sub
        szCol = CD_NPLATE_COL
        cText = CD_PLATE
        cKey = CD_PKEY
    Else
        If WmsOrders.FindCarSupplier(key) > 0 Then Exit Sub
        szCol = CD_NSUP_COL
        cText = CD_SUP
        cKey = CD_SKEY
    End If
    k = CLng(sh.getCellByPosition(szCol, 0).getValue()) + 1
    If k < 1 Or k > MAX_SHEET_ROW Then Exit Sub
    ' a damaged dictionary is not extended (the self-check names it); the visit itself does not depend on it
    If sh.getCellByPosition(cKey, k).getType() <> com.sun.star.table.CellContentType.EMPTY Then Exit Sub
    PlanSetValue(SH_CAR, k, cText, sText, False)
    PlanSetValue(SH_CAR, k, cKey, key, False)
    PlanSetValue(SH_CAR, 0, szCol, k, False)
End Sub

' ================================================================ ПРИЕХАЛ — CAR_ARRIVE

' A new visit of the vehicle sPlate from the supplier sSup (texts as typed). The two input cells of the panel are emptied by
' the same operation. OK:<seq>|<visit №>|<row> (0-based), ERR:<why> (the input or the vehicle is already there),
' ERR-SYS:<why>, BLOCKED:… or an ApplyOperation error.
Function CarArrive(ByVal sPlate As String, ByVal sSup As String) As String
    Dim why As String, key As String, skey As String, n As Long, r As Long, t As Double, d As Double, nday As Long, o As Long
    Dim prev As Variant, nVis As Long, nOrd As Long, ctl As String, res As String, k As Long
    WmsInit()
    On Error GoTo EH
    sPlate = CleanText(sPlate)
    sSup = CleanText(sSup)
    If sPlate = "" Then
        CarArrive = "ERR:впишите марку и госномер машины в ячейку «Марка / госномер»"
        Exit Function
    End If
    If sSup = "" Then
        CarArrive = "ERR:впишите или выберите поставщика в ячейке «Поставщик»"
        Exit Function
    End If
    If Len(sPlate) > CAR_TEXT_MAX Or Len(sSup) > CAR_TEXT_MAX Then
        CarArrive = "ERR:слишком длинный текст (больше " & CAR_TEXT_MAX & " знаков) — впишите только марку и госномер, только поставщика"
        Exit Function
    End If
    key = PlateKey(sPlate)
    If key = "" Then
        CarArrive = "ERR:в «Марка / госномер» нет ни букв, ни цифр"
        Exit Function
    End If
    skey = SupplierKey(sSup)
    why = PostingBlockReason()
    If why <> "" Then
        CarArrive = why
        Exit Function
    End If
    o = WmsOrders.FindCarOpen(key, 0)
    If o > 0 Then
        CarArrive = "ERR:эта машина уже на территории — визит № " & o & " " & VisitText(o) & ". Сначала отметьте её выезд кнопкой УЕХАЛ"
        Exit Function
    End If
    n = CLng(SysNum(SK_NEXT_CAR))
    why = NewCarProblem(n)
    If why <> "" Then
        CarArrive = "ERR-SYS:" & why
        Exit Function
    End If
    r = CAR_FIRST + n - 1
    t = CarNow()
    d = Int(t)
    ' the № of the day follows the last visit written before (a № of an abandoned operation leaves an empty row)
    nday = 1
    For k = n - 1 To IIf(n > 50, n - 50, 1) Step -1
        prev = CarRow(k)
        If VarType(prev(CR_DAY)) = 5 And VarType(prev(CR_NDAY)) = 5 Then
            If prev(CR_DAY) = d Then nday = CLng(prev(CR_NDAY)) + 1
            Exit For
        End If
    Next k
    nVis = WmsOrders.CountCarVisits(key)
    nOrd = OpenOrderPositions(sSup)
    ctl = "Приезд записан · " & IIf(nVis > 0, "повторная машина: визитов раньше — " & nVis, "первый визит этой машины")
    If nOrd > 0 Then ctl = ctl & " · открытых позиций заказов поставщика: " & nOrd
    If nOrd = 0 Then ctl = ctl & " · открытых заказов поставщика нет"
    PlanBegin("CAR_ARRIVE", CStr(n))
    PlanField("CAR", n)
    PlanField("PLATE", sPlate)
    PlanField("KEY", key)
    PlanField("SUPPLIER", sSup)
    PlanField("TIME", Format(CDate(t), "YYYY-MM-DD") & " " & Format(CDate(t), "HH:MM:SS"))
    PlanField("ROW", r + 1)
    ' one operation, the writes grouped by sheet: _SYS → _CAR → «Приход авто»
    PlanSetValue(SYS_SHEET, SK_NEXT_CAR, 1, n + 1, False)
    PlanSetValue(SH_CAR, n, CR_NO, n, False)
    PlanSetValue(SH_CAR, n, CR_KEY, key, False)
    PlanSetValue(SH_CAR, n, CR_STATE, CR_OPEN, False)
    PlanSetValue(SH_CAR, n, CR_ROW, r, False)
    PlanSetValue(SH_CAR, n, CR_ARR, t, False)
    PlanSetValue(SH_CAR, n, CR_DAY, d, False)
    PlanSetValue(SH_CAR, n, CR_NDAY, nday, False)
    PlanSetValue(SH_CAR, n, CR_OKEY, key, False)
    PlanSetValue(SH_CAR, n, CR_OARR, t, False)
    PlanSetValue(SH_CAR, n, CR_SKEY, skey, False)
    PlanDictionary(key, sPlate, True)
    PlanDictionary(skey, sSup, False)
    PlanSetValue(SH_CARS, r, CC_NO, n, False)
    PlanSetValue(SH_CARS, r, CC_DATE, d, False)
    PlanSetValue(SH_CARS, r, CC_PLATE, sPlate, False)
    PlanSetValue(SH_CARS, r, CC_SUP, sSup, False)
    PlanSetValue(SH_CARS, r, CC_ARR, t, False)
    ' the stay of a vehicle on the territory runs while it is there (the value is fixed by УЕХАЛ)
    PlanFormula(SH_CARS, r, CC_DUR, "=NOW()-E" & (r + 1))
    PlanSetValue(SH_CARS, r, CC_STATUS, CS_OPEN, False)
    PlanSetValue(SH_CARS, r, CC_WDAY, WeekdayText(d), False)
    PlanSetValue(SH_CARS, r, CC_MONTH, Format(CDate(d), "YYYY-MM"), False)
    PlanSetValue(SH_CARS, r, CC_NDAY, nday, False)
    PlanSetValue(SH_CARS, r, CC_CTL, ctl, False)
    PlanSetValue(SH_CARS, r, CC_KEY, key, False)
    ' the input cells of the panel are emptied (derived: a replay after a crash finds them in any state)
    PlanDerived(SH_CARS, CP_IN_ROW, CP_PLATE_COL, "")
    PlanDerived(SH_CARS, CP_IN_ROW, CP_SUP_COL, "")
    res = ApplyOperation(0, 0)
    If Left(res, 3) = "OK:" Then res = res & "|" & n & "|" & r
    CarArrive = res
    Exit Function
EH:
    CarArrive = "ERR-SYS:внутренняя ошибка «ПРИЕХАЛ»: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ УЕХАЛ — CAR_DEPART

' «Контроль» of a visit with the part about its end («выезд …» / «исправлено …» / «отменён …») replaced
Private Function CtlWith(ByVal old As String, sPart As String) As String
    Dim p As Long
    p = InStr(old, " · выезд ")
    If p > 0 Then old = Left(old, p - 1)
    CtlWith = old & sPart
End Function

' "" when the visit may be changed now; otherwise the refusal (d — the _CAR row)
Private Function VisitProblem(n As Long, d As Variant, sAction As String) As String
    Dim v As Variant
    If Not CarRegistered(n, d) Then
        VisitProblem = "визит № " & n & " не найден — выделите строку машины в таблице"
        Exit Function
    End If
    v = VisitCells(n)
    If VarType(v(CC_NO)) <> 5 Then
        VisitProblem = "строка визита № " & n & " на листе повреждена — нужна проверка ответственным"
        Exit Function
    End If
    If v(CC_NO) <> n Or CStr(v(CC_KEY)) <> CStr(d(CR_KEY)) Then
        VisitProblem = "строка визита № " & n & " на листе не совпадает со служебной записью — нужна проверка ответственным"
        Exit Function
    End If
    Select Case CStr(d(CR_STATE))
    Case CR_OPEN
        VisitProblem = ""
    Case CR_CLOSED
        If sAction = "DEPART" Then
            VisitProblem = "машина " & VisitText(n) & " уже уехала: визит № " & n & ", выезд " & Format(CDate(d(CR_DEP)), "DD.MM") & " в " _
                & HhMm(CDbl(d(CR_DEP))) & ". Выделите строку машины, которая сейчас на территории"
        Else
            VisitProblem = ""
        End If
    Case CR_CANCELLED
        VisitProblem = "визит № " & n & " " & VisitText(n) & " отменён — его нельзя " & IIf(sAction = "DEPART", "закрыть", "изменить")
    Case Else
        VisitProblem = "состояние визита № " & n & " «" & d(CR_STATE) & "» неизвестно — нужна проверка ответственным"
    End Select
End Function

' The departure of visit n now. OK:<seq>|<n>|<row>, ERR:<why>, ERR-SYS:…, BLOCKED:…
Function CarDepart(n As Long) As String
    Dim d As Variant, v As Variant, why As String, t As Double, arr As Double, dur As Double, r As Long, res As String, lim As Long
    Dim sPart As String
    WmsInit()
    On Error GoTo EH
    why = VisitProblem(n, d, "DEPART")
    If why <> "" Then
        CarDepart = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        CarDepart = why
        Exit Function
    End If
    arr = CDbl(d(CR_ARR))
    t = CarNow()
    ' the times are whole seconds written as numbers: within a second they are the same moment
    If t < arr - 1 / 86400 Then
        CarDepart = "ERR:часы компьютера показывают " & DayText(t) & " " & HhMm(t) & " — раньше приезда этой машины (" & DayText(arr) & " " _
            & HhMm(arr) & "). Проверьте дату и время компьютера"
        Exit Function
    End If
    If t < arr Then t = arr
    dur = t - arr
    r = CAR_FIRST + n - 1
    v = VisitCells(n)
    lim = LongMinutes()
    sPart = " · выезд " & HhMm(t) & ", на территории " & DurText(dur)
    If dur * 1440 > lim Then sPart = sPart & " — дольше порога " & lim & " мин"
    PlanBegin("CAR_DEPART", CStr(n))
    PlanField("CAR", n)
    PlanField("PLATE", CStr(v(CC_PLATE)))
    PlanField("KEY", CStr(d(CR_KEY)))
    PlanField("TIME", Format(CDate(t), "YYYY-MM-DD") & " " & Format(CDate(t), "HH:MM:SS"))
    PlanField("MIN", Int(dur * 1440 + 0.5))
    PlanField("ROW", r + 1)
    PlanSetValue(SH_CAR, n, CR_STATE, CR_CLOSED, False)
    PlanSetValue(SH_CAR, n, CR_DEP, t, False)
    PlanSetValue(SH_CAR, n, CR_DUR, dur, False)
    PlanSetValue(SH_CAR, n, CR_OKEY, "", False)
    PlanSetValue(SH_CAR, n, CR_OARR, "", False)
    PlanSetValue(SH_CARS, r, CC_DEP, t, False)
    PlanDerived(SH_CARS, r, CC_DUR, dur)
    PlanSetValue(SH_CARS, r, CC_STATUS, CS_GONE, False)
    PlanSetValue(SH_CARS, r, CC_CTL, CtlWith(CStr(v(CC_CTL)), sPart), False)
    res = ApplyOperation(0, 0)
    If Left(res, 3) = "OK:" Then res = res & "|" & n & "|" & r
    CarDepart = res
    Exit Function
EH:
    CarDepart = "ERR-SYS:внутренняя ошибка «УЕХАЛ»: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Исправить» — CAR_FIX

' A time of the window of «Исправить»: "ЧЧ:ММ" (on the day base; bNext — the next day when it would be before tMin) or
' "ДД.ММ.ГГГГ ЧЧ:ММ". False with msg when the text is not a time. ("base" is a word of Basic: Option Base)
Private Function ParseClock(ByVal s As String, ByVal baseDay As Double, ByVal tMin As Double, ByVal bNext As Boolean, ByRef t As Double, ByRef msg As String) As Boolean
    Dim p As Variant, hh As Integer, mm As Integer, tm As String, dt As String, dd As Variant
    s = Trim(Replace(s, ".", ":", 1, -1, 0))
    ' a date part: "27:09:2026 10:40" after the replacement above — the first space splits it
    p = InStr(s, " ")
    If p > 0 Then
        dt = Left(s, p - 1)
        tm = Trim(Mid(s, p + 1))
    Else
        tm = s
    End If
    p = Split(tm, ":")
    If UBound(p) <> 1 Then GoTo BAD
    If Not IsDigits(CStr(p(0))) Or Not IsDigits(CStr(p(1))) Or Len(p(0)) > 2 Or Len(p(1)) <> 2 Then GoTo BAD
    hh = CInt(p(0))
    mm = CInt(p(1))
    If hh > 23 Or mm > 59 Then GoTo BAD
    If dt <> "" Then
        dd = Split(dt, ":")
        If UBound(dd) <> 2 Then GoTo BAD
        If Not IsDigits(CStr(dd(0))) Or Not IsDigits(CStr(dd(1))) Or Not IsDigits(CStr(dd(2))) Or Len(dd(2)) <> 4 Then GoTo BAD
        If CInt(dd(1)) < 1 Or CInt(dd(1)) > 12 Or CInt(dd(0)) < 1 Or CInt(dd(0)) > 31 Then GoTo BAD
        baseDay = CDbl(DateSerial(CInt(dd(2)), CInt(dd(1)), CInt(dd(0))))
        If Day(CDate(baseDay)) <> CInt(dd(0)) Then GoTo BAD
        bNext = False
    End If
    t = Int(baseDay) + (hh * 60 + mm) / 1440
    If bNext And t < tMin Then t = t + 1
    ParseClock = True
    Exit Function
BAD:
    msg = "время «" & s & "» не понято — впишите ЧЧ:ММ (например 09:15) или ДД.ММ.ГГГГ ЧЧ:ММ"
End Function

' Corrects visit n: the texts (sPlate, sSup; "" — unchanged) and the times (sArr: ЧЧ:ММ of the day of the arrival;
' sDep: ЧЧ:ММ or ДД.ММ.ГГГГ ЧЧ:ММ, only for a vehicle that left; "" — unchanged). OK:<seq>|<n>|<row>, ERR:<why>, …
Function CarFix(n As Long, ByVal sPlate As String, ByVal sSup As String, ByVal sArr As String, ByVal sDep As String) As String
    Dim d As Variant, v As Variant, why As String, key As String, skey As String, arr As Double, dep As Double, t As Double
    Dim bOpen As Boolean, o As Long, r As Long, res As String, changes As String, now_ As Double, msg As String, fx As Long
    Dim oldArr As Double, oldDep As Double, oldPlate As String, oldSup As String
    WmsInit()
    On Error GoTo EH
    why = VisitProblem(n, d, "FIX")
    If why <> "" Then
        CarFix = "ERR:" & why
        Exit Function
    End If
    v = VisitCells(n)
    bOpen = (d(CR_STATE) = CR_OPEN)
    oldPlate = CStr(v(CC_PLATE))
    oldSup = CStr(v(CC_SUP))
    oldArr = CDbl(d(CR_ARR))
    If Not bOpen Then oldDep = CDbl(d(CR_DEP))
    sPlate = CleanText(sPlate)
    sSup = CleanText(sSup)
    If sPlate = "" Then sPlate = oldPlate
    If sSup = "" Then sSup = oldSup
    If Len(sPlate) > CAR_TEXT_MAX Or Len(sSup) > CAR_TEXT_MAX Then
        CarFix = "ERR:слишком длинный текст (больше " & CAR_TEXT_MAX & " знаков)"
        Exit Function
    End If
    key = PlateKey(sPlate)
    If key = "" Then
        CarFix = "ERR:в «Марка / госномер» нет ни букв, ни цифр"
        Exit Function
    End If
    skey = SupplierKey(sSup)
    now_ = CarNow()
    arr = oldArr
    If Trim(sArr) <> "" Then
        If Not ParseClock(sArr, oldArr, 0, False, t, msg) Then
            CarFix = "ERR:приезд: " & msg
            Exit Function
        End If
        If Int(t) <> Int(oldArr) Then
            CarFix = "ERR:дату приезда изменить нельзя (визит записан за " & DayText(oldArr) & ") — отмените ошибочный визит кнопкой «Отменить визит»"
            Exit Function
        End If
        arr = t
    End If
    If arr > now_ + 1 / 1440 Then
        CarFix = "ERR:время приезда " & HhMm(arr) & " ещё не наступило"
        Exit Function
    End If
    dep = oldDep
    If Trim(sDep) <> "" Then
        If bOpen Then
            CarFix = "ERR:машина ещё на территории — время выезда записывает кнопка УЕХАЛ"
            Exit Function
        End If
        If Not ParseClock(sDep, arr, arr, True, t, msg) Then
            CarFix = "ERR:выезд: " & msg
            Exit Function
        End If
        dep = t
    End If
    If Not bOpen Then
        If dep < arr - 0.5 / 86400 Then
            CarFix = "ERR:выезд (" & DayText(dep) & " " & HhMm(dep) & ") раньше приезда (" & DayText(arr) & " " & HhMm(arr) & ")"
            Exit Function
        End If
        If dep < arr Then dep = arr
        If dep > now_ + 1 / 1440 Then
            CarFix = "ERR:время выезда " & DayText(dep) & " " & HhMm(dep) & " ещё не наступило"
            Exit Function
        End If
    End If
    If sPlate <> oldPlate Then changes = changes & ", машина «" & oldPlate & "» → «" & sPlate & "»"
    If sSup <> oldSup Then changes = changes & ", поставщик «" & oldSup & "» → «" & sSup & "»"
    If Abs(arr - oldArr) > 0.5 / 86400 Then changes = changes & ", приезд " & HhMm(oldArr) & " → " & HhMm(arr)
    If Not bOpen Then
        If Abs(dep - oldDep) > 0.5 / 86400 Then changes = changes & ", выезд " & HhMm(oldDep) & " → " & HhMm(dep)
    End If
    If changes = "" Then
        CarFix = "ERR:изменений нет — исправьте машину, поставщика или время"
        Exit Function
    End If
    changes = Mid(changes, 3)
    why = PostingBlockReason()
    If why <> "" Then
        CarFix = why
        Exit Function
    End If
    If bOpen And key <> CStr(d(CR_KEY)) Then
        o = WmsOrders.FindCarOpen(key, 0)
        If o > 0 And o <> n Then
            CarFix = "ERR:машина с этим госномером уже на территории — визит № " & o & " " & VisitText(o)
            Exit Function
        End If
    End If
    r = CAR_FIRST + n - 1
    If VarType(d(CR_FIXES)) = 5 Then fx = CLng(d(CR_FIXES))
    PlanBegin("CAR_FIX", CStr(n))
    PlanField("CAR", n)
    If sPlate <> oldPlate Then
        PlanField("PLATE_OLD", oldPlate)
        PlanField("PLATE", sPlate)
        PlanField("KEY", key)
    End If
    If sSup <> oldSup Then
        PlanField("SUPPLIER_OLD", oldSup)
        PlanField("SUPPLIER", sSup)
    End If
    PlanField("ARR", Format(CDate(arr), "YYYY-MM-DD") & " " & Format(CDate(arr), "HH:MM:SS"))
    If Not bOpen Then PlanField("DEP", Format(CDate(dep), "YYYY-MM-DD") & " " & Format(CDate(dep), "HH:MM:SS"))
    PlanField("ROW", r + 1)
    PlanSetValue(SH_CAR, n, CR_KEY, key, False)
    PlanSetValue(SH_CAR, n, CR_ARR, arr, False)
    If bOpen Then
        PlanSetValue(SH_CAR, n, CR_OKEY, key, False)
        PlanSetValue(SH_CAR, n, CR_OARR, arr, False)
    Else
        PlanSetValue(SH_CAR, n, CR_DEP, dep, False)
        PlanSetValue(SH_CAR, n, CR_DUR, dep - arr, False)
    End If
    PlanSetValue(SH_CAR, n, CR_FIXES, fx + 1, False)
    PlanSetValue(SH_CAR, n, CR_SKEY, skey, False)
    PlanDictionary(key, sPlate, True)
    PlanDictionary(skey, sSup, False)
    PlanSetValue(SH_CARS, r, CC_PLATE, sPlate, False)
    PlanSetValue(SH_CARS, r, CC_SUP, sSup, False)
    PlanSetValue(SH_CARS, r, CC_ARR, arr, False)
    If Not bOpen Then
        PlanSetValue(SH_CARS, r, CC_DEP, dep, False)
        PlanDerived(SH_CARS, r, CC_DUR, dep - arr)
    End If
    PlanSetValue(SH_CARS, r, CC_CTL, CStr(v(CC_CTL)) & " · исправлено " & Format(CDate(now_), "DD.MM HH:MM") & ": " & changes, False)
    PlanSetValue(SH_CARS, r, CC_KEY, key, False)
    res = ApplyOperation(0, 0)
    If Left(res, 3) = "OK:" Then res = res & "|" & n & "|" & r
    CarFix = res
    Exit Function
EH:
    CarFix = "ERR-SYS:внутренняя ошибка «Исправить»: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Отменить визит» — CAR_CANCEL

' Cancels visit n (an erroneous record): it stays in the history as «Отменён» and leaves every indicator.
Function CarCancel(n As Long) As String
    Dim d As Variant, v As Variant, why As String, r As Long, res As String, bOpen As Boolean
    WmsInit()
    On Error GoTo EH
    why = VisitProblem(n, d, "CANCEL")
    If why <> "" Then
        CarCancel = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        CarCancel = why
        Exit Function
    End If
    v = VisitCells(n)
    bOpen = (d(CR_STATE) = CR_OPEN)
    r = CAR_FIRST + n - 1
    PlanBegin("CAR_CANCEL", CStr(n))
    PlanField("CAR", n)
    PlanField("PLATE", CStr(v(CC_PLATE)))
    PlanField("KEY", CStr(d(CR_KEY)))
    PlanField("STATE_BEFORE", CStr(d(CR_STATE)))
    PlanField("ROW", r + 1)
    PlanSetValue(SH_CAR, n, CR_STATE, CR_CANCELLED, False)
    If bOpen Then
        PlanSetValue(SH_CAR, n, CR_OKEY, "", False)
        PlanSetValue(SH_CAR, n, CR_OARR, "", False)
        PlanDerived(SH_CARS, r, CC_DUR, "")
    End If
    PlanSetValue(SH_CARS, r, CC_STATUS, CS_CANCEL, False)
    PlanSetValue(SH_CARS, r, CC_CTL, CStr(v(CC_CTL)) & " · отменён " & Format(CDate(CarNow()), "DD.MM HH:MM") & " (ошибочная запись)", False)
    res = ApplyOperation(0, 0)
    If Left(res, 3) = "OK:" Then res = res & "|" & n & "|" & r
    CarCancel = res
    Exit Function
EH:
    CarCancel = "ERR-SYS:внутренняя ошибка «Отменить визит»: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ self-check (a subset of WMS_DOCTOR, spec §49)

Private Sub Bad(ByRef nBad As Long, ByRef first As String, s As String)
    nBad = nBad + 1
    If nBad <= 3 Then first = first & IIf(first <> "", "; ", "") & s
End Sub

' "$'sheet'.$C$r1:$C$r2"
Private Function Rg(sSheet As String, col As String, r1 As Long, r2 As Long) As String
    Rg = "$'" & sSheet & "'.$" & col & "$" & r1 & ":$" & col & "$" & r2
End Function

' The whole table checked by formulas of the Calc engine (one array formula, milliseconds on tens of thousands of visits):
' "problems|open|closed|cancelled|gaps", or "" when the formula could not be used. A problem is any visit that the
' detailed check below would name; a gap is a № handed out by an abandoned operation (no record in both places).
Private Function QuickCounts(last As Long) As String
    Dim a As String, b As String, c As String, e As String, f As String, i_ As String, j As String, sA As String, sC As String
    Dim sD As String, sE As String, sF As String, sH As String, sN As String, x As String, ok As Boolean, r As String, t As Long, d As String
    t = CAR_FIRST + last
    a = Rg(SH_CAR, "A", 2, last + 1)
    b = Rg(SH_CAR, "B", 2, last + 1)
    c = Rg(SH_CAR, "C", 2, last + 1)
    e = Rg(SH_CAR, "E", 2, last + 1)
    f = Rg(SH_CAR, "F", 2, last + 1)
    i_ = Rg(SH_CAR, "I", 2, last + 1)
    j = Rg(SH_CAR, "J", 2, last + 1)
    sA = Rg(SH_CARS, "A", CAR_FIRST + 1, t)
    sC = Rg(SH_CARS, "C", CAR_FIRST + 1, t)
    sD = Rg(SH_CARS, "D", CAR_FIRST + 1, t)
    sE = Rg(SH_CARS, "E", CAR_FIRST + 1, t)
    sF = Rg(SH_CARS, "F", CAR_FIRST + 1, t)
    sH = Rg(SH_CARS, "H", CAR_FIRST + 1, t)
    sN = Rg(SH_CARS, "N", CAR_FIRST + 1, t)
    d = "(" & a & "<>"""")"
    x = "SUMPRODUCT((" & a & "="""")<>(" & sA & "=""""))" _
        & "+SUMPRODUCT(" & d & "*(" & a & "<>ROW(" & a & ")-1))" _
        & "+SUMPRODUCT(" & d & "*(" & sA & "<>" & a & "))" _
        & "+SUMPRODUCT(" & d & "*((" & b & "="""")+(" & b & "<>" & sN & ")+(TRIM(" & sC & ")="""")+(TRIM(" & sD & ")="""")))" _
        & "+SUMPRODUCT(" & d & "*(NOT(ISNUMBER(" & e & "))+IFERROR(ABS(" & e & "-" & sE & ")>0.00001;1)))" _
        & "+COUNTIF(" & a & ";""<>"")-COUNTIF(" & c & ";""" & CR_OPEN & """)-COUNTIF(" & c & ";""" & CR_CLOSED & """)-COUNTIF(" & c & ";""" & CR_CANCELLED & """)" _
        & "+SUMPRODUCT((" & c & "=""" & CR_OPEN & """)*((" & sH & "<>""" & CS_OPEN & """)+(" & i_ & "<>" & b & ")+NOT(ISNUMBER(" & j & "))+(" & f & "<>"""")))" _
        & "+SUMPRODUCT((" & c & "=""" & CR_CLOSED & """)*((" & sH & "<>""" & CS_GONE & """)+(" & i_ & "<>"""")+(" & j & "<>"""")+NOT(ISNUMBER(" & f & "))" _
        & "+IFERROR(" & f & "<" & e & "-0.00001;1)+IFERROR(ABS(" & f & "-" & sF & ")>0.00001;1)))" _
        & "+SUMPRODUCT((" & c & "=""" & CR_CANCELLED & """)*((" & sH & "<>""" & CS_CANCEL & """)+(" & i_ & "<>"""")+(" & j & "<>"""")))"
    r = WmsOrders.ScratchEval("(" & x & ")&""|""&COUNTIF(" & c & ";""" & CR_OPEN & """)&""|""&COUNTIF(" & c & ";""" & CR_CLOSED & """)&""|""&COUNTIF(" _
        & c & ";""" & CR_CANCELLED & """)&""|""&COUNTBLANK(" & a & ")&""|""&TEXTJOIN("";"";1;" & i_ & ")", ok)
    If ok Then QuickCounts = r Else QuickCounts = ""
End Function

' Visit n on _CAR row n and on the table row CAR_FIRST + n − 1 for n < NEXT_CAR (a № of an abandoned operation: no record in
' both places), nothing after; the states and times agree; no vehicle has two open visits; the dictionaries hold their
' entries. The whole table is checked by formulas of the Calc engine; only when they find a problem (or cannot be used)
' the visits are read in chunks of CHECK_CHUNK_ROWS and the first problems are named.
Function CarCheck() As String
    Dim last As Long, n0 As Long, n1 As Long, i As Long, n As Long, dc As Variant, dv As Variant, c As Variant, v As Variant
    Dim nBad As Long, first As String, nOpen As Long, nGone As Long, nCanc As Long, nGap As Long, openKeys As String, t0 As Long, st As String
    Dim k As Integer, sz As Long, sh As Object, cars As Object, e As Variant, x As Double, q As String, qa As Variant, ks As Variant
    On Error GoTo EH
    t0 = GetSystemTicks()
    sh = CarSheet()
    cars = CarsSheet()
    last = CLng(SysNum(SK_NEXT_CAR)) - 1
    If last >= 1 Then q = QuickCounts(last)
    If last < 1 Then
        ' no visits yet
    ElseIf q <> "" Then
        qa = Split(q, "|")
        nOpen = CLng(qa(1))
        nGone = CLng(qa(2))
        nCanc = CLng(qa(3))
        nGap = CLng(qa(4))
        ' the open keys (few): none twice
        If UBound(qa) >= 5 Then
            If qa(5) <> "" Then
                ks = Split(qa(5), ";")
                openKeys = "|"
                For i = 0 To UBound(ks)
                    If InStr(1, openKeys, "|" & ks(i) & "|", 0) > 0 Then Bad(nBad, first, "машина " & ks(i) & " открыта дважды")
                    openKeys = openKeys & ks(i) & "|"
                Next i
            End If
        End If
    End If
    If last >= 1 And (q = "" Or Val(q) <> 0) Then
        ' the detailed check: names the first problems
        nOpen = 0
        nGone = 0
        nCanc = 0
        nGap = 0
        nBad = 0
        first = ""
        openKeys = "|"
        For n0 = 1 To last Step CHECK_CHUNK_ROWS
            n1 = n0 + CHECK_CHUNK_ROWS - 1
            If n1 > last Then n1 = last
            dc = sh.getCellRangeByPosition(0, n0, CR_LAST, n1).getDataArray()
            dv = cars.getCellRangeByPosition(0, CAR_FIRST + n0 - 1, CC_LAST, CAR_FIRST + n1 - 1).getDataArray()
            For i = 0 To n1 - n0
                n = n0 + i
                c = dc(i)
                v = dv(i)
                If CStr(c(CR_NO)) = "" And CStr(v(CC_NO)) = "" And CStr(c(CR_KEY)) = "" And CStr(v(CC_PLATE)) = "" Then
                    nGap = nGap + 1
                ElseIf VarType(c(CR_NO)) <> 5 Then
                    Bad(nBad, first, "_CAR: нет записи визита № " & n)
                ElseIf c(CR_NO) <> n Or VarType(v(CC_NO)) <> 5 Then
                    Bad(nBad, first, "визит № " & n & ": номер не совпадает со строкой")
                ElseIf v(CC_NO) <> n Then
                    Bad(nBad, first, "визит № " & n & ": в строке " & (CAR_FIRST + n) & " листа № «" & v(CC_NO) & "»")
                ElseIf CStr(c(CR_KEY)) = "" Or CStr(c(CR_KEY)) <> CStr(v(CC_KEY)) Or Trim(CStr(v(CC_PLATE))) = "" Or Trim(CStr(v(CC_SUP))) = "" Then
                    Bad(nBad, first, "визит № " & n & ": пустая машина, поставщик или ключ")
                ElseIf VarType(c(CR_ARR)) <> 5 Or VarType(v(CC_ARR)) <> 5 Then
                    Bad(nBad, first, "визит № " & n & ": нет времени приезда")
                ElseIf Abs(c(CR_ARR) - v(CC_ARR)) > 0.5 / 86400 Then
                    Bad(nBad, first, "визит № " & n & ": время приезда на листе и в " & SH_CAR & " различается")
                Else
                    st = CStr(c(CR_STATE))
                    If st = CR_OPEN Then
                        nOpen = nOpen + 1
                        If CStr(v(CC_STATUS)) <> CS_OPEN Or CStr(c(CR_OKEY)) <> CStr(c(CR_KEY)) Or VarType(c(CR_OARR)) <> 5 Or CStr(c(CR_DEP)) <> "" Then
                            Bad(nBad, first, "визит № " & n & ": открытый визит записан не полностью")
                        ElseIf InStr(1, openKeys, "|" & c(CR_KEY) & "|", 0) > 0 Then
                            Bad(nBad, first, "машина " & c(CR_KEY) & " открыта дважды (визит № " & n & ")")
                        Else
                            openKeys = openKeys & c(CR_KEY) & "|"
                        End If
                    ElseIf st = CR_CLOSED Then
                        nGone = nGone + 1
                        If CStr(v(CC_STATUS)) <> CS_GONE Or CStr(c(CR_OKEY)) <> "" Or CStr(c(CR_OARR)) <> "" Or VarType(c(CR_DEP)) <> 5 Then
                            Bad(nBad, first, "визит № " & n & ": закрытый визит записан не полностью")
                        ElseIf c(CR_DEP) < c(CR_ARR) - 0.5 / 86400 Then
                            Bad(nBad, first, "визит № " & n & ": выезд раньше приезда")
                        ElseIf VarType(v(CC_DEP)) <> 5 Then
                            Bad(nBad, first, "визит № " & n & ": на листе нет времени выезда")
                        ElseIf Abs(c(CR_DEP) - v(CC_DEP)) > 0.5 / 86400 Then
                            Bad(nBad, first, "визит № " & n & ": время выезда на листе и в " & SH_CAR & " различается")
                        End If
                    ElseIf st = CR_CANCELLED Then
                        nCanc = nCanc + 1
                        If CStr(v(CC_STATUS)) <> CS_CANCEL Or CStr(c(CR_OKEY)) <> "" Or CStr(c(CR_OARR)) <> "" Then
                            Bad(nBad, first, "визит № " & n & ": отменённый визит записан не полностью")
                        End If
                    Else
                        Bad(nBad, first, "визит № " & n & ": неизвестное состояние «" & st & "»")
                    End If
                End If
            Next i
        Next n0
    End If
    ' nothing after the last visit (the counter never lags behind the data)
    If last + 1 <= MAX_SHEET_ROW Then
        If sh.getCellByPosition(CR_NO, last + 1).getType() <> com.sun.star.table.CellContentType.EMPTY _
            Or sh.getCellByPosition(CR_KEY, last + 1).getType() <> com.sun.star.table.CellContentType.EMPTY Then
            Bad(nBad, first, SH_CAR & ": запись после последнего визита № " & last & " (счётчик NEXT_CAR отстаёт)")
        End If
    End If
    If CAR_FIRST + last <= MAX_SHEET_ROW Then
        If cars.getCellByPosition(CC_NO, CAR_FIRST + last).getType() <> com.sun.star.table.CellContentType.EMPTY _
            Or cars.getCellByPosition(CC_PLATE, CAR_FIRST + last).getType() <> com.sun.star.table.CellContentType.EMPTY Then
            Bad(nBad, first, "лист «" & SH_CARS & "»: строка " & (CAR_FIRST + last + 1) & " после последнего визита не пуста")
        End If
    End If
    ' the dictionaries: entries 1..size have a key, the entry after them is empty
    For Each e In Array(CD_NPLATE_COL, CD_NSUP_COL)
        x = sh.getCellByPosition(e, 0).getValue()
        sz = CLng(x)
        k = IIf(e = CD_NPLATE_COL, CD_PKEY, CD_SKEY)
        If sz > 0 Then
            dc = sh.getCellRangeByPosition(k, 1, k, sz + IIf(sz < MAX_SHEET_ROW, 1, 0)).getDataArray()
            For i = 0 To sz - 1
                If CStr(dc(i)(0)) = "" Then
                    Bad(nBad, first, SH_CAR & ": в словаре пустая запись " & (i + 1))
                    Exit For
                End If
            Next i
            If sz < MAX_SHEET_ROW Then
                If CStr(dc(sz)(0)) <> "" Then Bad(nBad, first, SH_CAR & ": запись словаря после его размера " & sz)
            End If
        ElseIf sh.getCellByPosition(k, 1).getType() <> com.sun.star.table.CellContentType.EMPTY Then
            Bad(nBad, first, SH_CAR & ": запись словаря при нулевом размере")
        End If
    Next e
    If nBad > 0 Then
        CarCheck = "ОШИБКА: проблем " & nBad & ", первые: " & first
    Else
        CarCheck = "визитов " & (last - nGap) & ": на территории " & nOpen & ", уехало " & nGone & ", отменено " & nCanc _
            & IIf(nGap > 0, "; номеров отложенных операций: " & nGap, "") & "; словари: машин " _
            & CLng(sh.getCellByPosition(CD_NPLATE_COL, 0).getValue()) & ", поставщиков " & CLng(sh.getCellByPosition(CD_NSUP_COL, 0).getValue()) _
            & ", " & (GetSystemTicks() - t0) & " мс"
    End If
    Exit Function
EH:
    CarCheck = "ОШИБКА проверки визитов: " & Error$ & " (строка " & Erl & ")"
End Function

' ================================================================ test seam (inert unless _SYS MODE = TEST)

' the clock of the vehicle operations: x > 0 — that moment (a date-time number), 0 — the real time again
Function TestCarNow(x As Double) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestCarNow = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gCarNowSet = (x > 0)
    gCarNow = x
    TestCarNow = "OK"
End Function
