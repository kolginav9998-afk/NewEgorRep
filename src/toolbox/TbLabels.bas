' TbLabels — WMS_LABELS: этикетки ЕИ по снимку (задание «FINAL WMS MARATHON», §6; M6 PRIME §10). Лист «Этикетки»: список
' ЕИ и число копий — вручную или «Подобрать ЕИ» (массовый выбор из снимка: место, категория, тип источника, даты
' получения, диапазон номеров, только с остатком); «Построить этикетки» — лист «Макет»: одна этикетка — одна страница
' размера этикетки (Настройки: ширина и высота, мм); на этикетке ЕИ крупно, наименование, артикул, место и штрихкод
' Code 128 (номер ЕИ, 8 цифр) — рисуется прямоугольниками, внешние шрифты не нужны. «Предпросмотр» — одна этикетка
' первого ЕИ списка так, как она напечатается. «Печать» — системный принтер LibreOffice / CUPS (обычное окно печати);
' «В PDF» — файл для проверки; каждое задание записывается в «История печати», «Повторная печать»
' возвращает список задания (Настройки: номер; пусто — последнее) и строит его заново; история — файл
' ../WMS_Labels/history.csv (книга инструмента не сохраняется). Драйвер конкретного принтера ядро
' WMS не касается; адаптер прямого протокола (TSPL) — отдельный скрипт WMS_TOOLBOX/labels/tspl_labels.py.
Option Explicit

Private Const SH_LIST = "Этикетки"
Private Const SH_LAYOUT = "Макет"
Private Const SH_HIST = "История печати"
Private Const ROWS_PER_LABEL = 5
Private Const PICK_COL = 4             ' the filters of «Подобрать ЕИ»: column E of «Этикетки», rows 2–10
Global gLblJob As Variant              ' the list (ЕИ, копий) of the layout built last — what «Печать» / «В PDF» record
                                       ' (Global: a Private module variable is reset at every run of a macro)

' the Code 128 patterns (bar, space, bar, space, bar, space widths in modules); 105 Start C, 106 Stop (7 elements)
Private Function Patterns() As Variant
    Patterns = Split("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 113222 123122 " _
        & "123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 131123 131321 " _
        & "112313 132113 132311 211313 231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 " _
        & "311321 331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 " _
        & "142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 412121 111143 " _
        & "111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 211232 2331112", " ")
End Function

' the bar widths (modules) of Code 128 C for an even string of digits: start C, the pairs, the checksum, stop
Function Code128C(digits As String) As String
    Dim p As Variant, s As String, i As Integer, v As Integer, sum As Long, w As Integer
    p = Patterns()
    s = p(105)
    sum = 105
    w = 1
    For i = 1 To Len(digits) Step 2
        v = CInt(Mid(digits, i, 2))
        s = s & p(v)
        sum = sum + w * v
        w = w + 1
    Next i
    s = s & p(sum Mod 103) & p(106)
    Code128C = s
End Function

' «Обновить из снимка»: the registry of the latest snapshot and the receipts (for «Подобрать ЕИ»); the EIs to print are listed on «Этикетки»: A ЕИ, B копий
Sub BtnLblRefresh(Optional oEvent As Variant)
    Dim why As String, snap As String, n As Long
    snap = TbPickSnapshot(why)
    If snap = "" Then
        TbMsg("ERR:" & why)
        Exit Sub
    End If
    n = TbLoadTable(snap, "stock.csv", "_stock", True)
    TbLoadTable(snap, "orders.csv", "_orders", True)
    TbLoadTable(snap, "special.csv", "_special", True)
    TbPutSetting(3, snap)
    TbMsg("OK:ЕИ в снимке: " & n)
End Sub

' ================================================================ «Подобрать ЕИ»

Private Function PickText(r As Integer) As String
    PickText = Trim(ThisComponent.Sheets.getByName(SH_LIST).getCellByPosition(PICK_COL, r).getString())
End Function

' a date of the filter (a date cell or «ДД.ММ.ГГГГ»); 0 — none
Private Function PickDate(r As Integer) As Double
    Dim c As Object, t As String
    c = ThisComponent.Sheets.getByName(SH_LIST).getCellByPosition(PICK_COL, r)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then
        PickDate = Int(c.getValue())
    Else
        t = Trim(c.getString())
        If Len(t) = 10 And Mid(t, 3, 1) = "." And Mid(t, 6, 1) = "." Then PickDate = CDbl(DateSerial(Val(Right(t, 4)), Val(Mid(t, 4, 2)), Val(Left(t, 2))))
    End If
End Function

Sub BtnLblPick(Optional oEvent As Variant)
    TbMsg(LblPick())
End Sub

' «Подобрать ЕИ»: the EIs of the snapshot matching every filter of E2:E10 (place starts with, category, source type,
' received from / to, EI from / to, only with a balance, copies each) replace the list A:B
Function LblPick() As String
    Dim lst As Object, st As Variant, i As Long, n As Long, e As Long, pl As String, cat As String, src As String, d1 As Double, d2 As Double
    Dim e1 As Long, e2 As Long, bBal As Boolean, cp As Long, rcv() As Double, o As Variant, sp As Variant, rows() As Variant, last As Long, why As String
    On Error GoTo EH
    If Not ThisComponent.Sheets.hasByName("_stock") Then
        LblPick = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    lst = ThisComponent.Sheets.getByName(SH_LIST)
    pl = LCase(PickText(1))
    cat = LCase(PickText(2))
    src = LCase(PickText(3))
    d1 = PickDate(4)
    d2 = PickDate(5)
    e1 = Val(PickText(6))
    e2 = Val(PickText(7))
    bBal = (LCase(Left(PickText(8), 1)) <> "н")
    cp = Val(PickText(9))
    If cp < 1 Then cp = 1
    st = ThisComponent.Sheets.getByName("_stock").getCellRangeByPosition(0, 1, 9, TbLastRow(ThisComponent.Sheets.getByName("_stock"))).getDataArray()
    ' the first receipt of every EI (orders, «Иной приход»)
    ReDim rcv(UBound(st) + 2)
    If d1 > 0 Or d2 > 0 Then
        ' the receipts are loaded by «Обновить из снимка» (And of Basic evaluates both sides: the checks are nested)
        If Not ThisComponent.Sheets.hasByName("_orders") Or Not ThisComponent.Sheets.hasByName("_special") Then
            LblPick = "ERR:для отбора по датам получения нажмите «Обновить из снимка» (загружаются приходы)"
            Exit Function
        End If
        If TbLastRow(ThisComponent.Sheets.getByName("_orders")) >= 1 Then
            o = ThisComponent.Sheets.getByName("_orders").getCellRangeByPosition(0, 1, 21, TbLastRow(ThisComponent.Sheets.getByName("_orders"))).getDataArray()
            For i = 0 To UBound(o)
                FirstRcv(rcv, NormEI(CStr(o(i)(21))), o(i)(13))
            Next i
        End If
        If TbLastRow(ThisComponent.Sheets.getByName("_special")) >= 1 Then
            sp = ThisComponent.Sheets.getByName("_special").getCellRangeByPosition(0, 1, 13, TbLastRow(ThisComponent.Sheets.getByName("_special"))).getDataArray()
            For i = 0 To UBound(sp)
                If VarType(sp(i)(0)) = 5 Then FirstRcv(rcv, NormEI(CStr(sp(i)(13))), sp(i)(7))
            Next i
        End If
    End If
    ReDim rows(UBound(st))
    For i = 0 To UBound(st)
        If CStr(st(i)(0)) = "" Or CStr(st(i)(7)) = "Приход удалён (сторно)" Then GoTo NEXT_EI
        e = i + 1
        If e1 > 0 And e < e1 Then GoTo NEXT_EI
        If e2 > 0 And e > e2 Then GoTo NEXT_EI
        If bBal Then
            If VarType(st(i)(4)) <> 5 Then GoTo NEXT_EI
            If st(i)(4) <= 0 Then GoTo NEXT_EI
        End If
        If pl <> "" And Left(LCase(CStr(st(i)(5))), Len(pl)) <> pl Then GoTo NEXT_EI
        If cat <> "" And LCase(CStr(st(i)(6))) <> cat Then GoTo NEXT_EI
        If src <> "" And LCase(CStr(st(i)(9))) <> src Then GoTo NEXT_EI
        If d1 > 0 And (rcv(e) = 0 Or rcv(e) < d1) Then GoTo NEXT_EI
        If d2 > 0 And (rcv(e) = 0 Or rcv(e) > d2) Then GoTo NEXT_EI
        rows(n) = Array(CStr(st(i)(0)), cp)
        n = n + 1
NEXT_EI:
    Next i
    last = TbLastRow(lst)
    If last >= 1 Then lst.getCellRangeByPosition(0, 1, 1, last).clearContents(1023)
    If n > 0 Then
        ReDim Preserve rows(n - 1)
        lst.getCellRangeByPosition(0, 1, 1, n).setDataArray(rows)
    End If
    If pl <> "" Then why = why & ", место «" & pl & "…»"
    If cat <> "" Then why = why & ", категория «" & cat & "»"
    If src <> "" Then why = why & ", источник «" & src & "»"
    If d1 > 0 Or d2 > 0 Then why = why & ", получены " & IIf(d1 > 0, "с " & Format(CDate(d1), "DD.MM.YYYY"), "") & IIf(d2 > 0, " по " & Format(CDate(d2), "DD.MM.YYYY"), "")
    If e1 > 0 Or e2 > 0 Then why = why & ", номера " & IIf(e1 > 0, "с " & e1, "") & IIf(e2 > 0, " по " & e2, "")
    LblPick = "OK:подобрано ЕИ " & n & " × " & cp & " копий" & IIf(bBal, " (только с остатком" & why & ")", IIf(why <> "", " (" & Mid(why, 3) & ")", "")) _
        & IIf(n * cp > 1000, " — этикеток больше 1000: проверьте фильтры", "")
    Exit Function
EH:
    LblPick = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Sub FirstRcv(rcv() As Double, e As String, d As Variant)
    Dim k As Long
    If e = "" Or (VarType(d) <> 5 And VarType(d) <> 7) Then Exit Sub
    k = CLng(Mid(e, 4))
    If k < 1 Or k > UBound(rcv) Then Exit Sub
    If rcv(k) = 0 Or CDbl(d) < rcv(k) Then rcv(k) = Int(CDbl(d))
End Sub

Private Function Setting(r As Integer, dflt As Double) As Double
    Dim s As String
    s = Trim(Replace(TbSetting(r), ",", "."))
    If s = "" Then Setting = dflt Else Setting = Val(s)
End Function

Private Function Flag(r As Integer) As Boolean
    Flag = (LCase(Left(Trim(TbSetting(r)), 1)) <> "н")
End Function

Sub BtnLblBuild(Optional oEvent As Variant)
    TbMsg(LblBuild())
End Sub

' «Построить этикетки»: the layout sheet from the list; «ERR:» when an EI is not in the snapshot
Function LblBuild() As String
    Dim lst As Object, last As Long
    If Not ThisComponent.Sheets.hasByName("_stock") Then
        LblBuild = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    lst = ThisComponent.Sheets.getByName(SH_LIST)
    last = ListLast(lst)
    If last < 1 Then
        LblBuild = "ERR:на листе «" & SH_LIST & "» нет ЕИ (колонка A, начиная со строки 2)"
        Exit Function
    End If
    LblBuild = BuildFrom(lst.getCellRangeByPosition(0, 1, 1, last).getDataArray())
End Function

' the last row of the list A:B (the filters of «Подобрать ЕИ» in D:E do not count)
Private Function ListLast(lst As Object) As Long
    Dim cur As Object, d As Variant, i As Long
    cur = lst.createCursor()
    cur.gotoEndOfUsedArea(False)
    d = lst.getCellRangeByPosition(0, 0, 0, cur.getRangeAddress().EndRow).getDataArray()
    For i = UBound(d) To 1 Step -1
        If CStr(d(i)(0)) <> "" Then
            ListLast = i
            Exit Function
        End If
    Next i
End Function

' the layout from the rows (ЕИ, копий)
Private Function BuildFrom(d As Variant) As String
    Dim st As Object, lay As Object, i As Long, k As Long, c As Long, n As Long, e As String, row As Variant, job() As Variant, nJob As Long
    Dim w As Double, h As Double, r As Long, bad As String, nBad As Long, dp As Object, ps As Object, stName As String, x As Long
    On Error GoTo EH
    st = ThisComponent.Sheets.getByName("_stock")
    ReDim job(UBound(d))
    w = Setting(5, 58)
    h = Setting(6, 40)
    If ThisComponent.Sheets.hasByName(SH_LAYOUT) Then ThisComponent.Sheets.removeByName(SH_LAYOUT)
    ThisComponent.Sheets.insertNewByName(SH_LAYOUT, ThisComponent.Sheets.getCount())
    lay = ThisComponent.Sheets.getByName(SH_LAYOUT)
    ' a page style of the label size: one label — one page
    stName = "WMS_Этикетка"
    If Not ThisComponent.StyleFamilies.getByName("PageStyles").hasByName(stName) Then
        ps = ThisComponent.createInstance("com.sun.star.style.PageStyle")
        ThisComponent.StyleFamilies.getByName("PageStyles").insertByName(stName, ps)
    End If
    ps = ThisComponent.StyleFamilies.getByName("PageStyles").getByName(stName)
    ps.Width = w * 100
    ps.Height = h * 100
    ps.LeftMargin = 200
    ps.RightMargin = 200
    ps.TopMargin = 200
    ps.BottomMargin = 200
    ps.HeaderIsOn = False
    ps.FooterIsOn = False
    ps.PrintGrid = False
    ps.CenterHorizontally = True
    lay.PageStyle = stName
    lay.getColumns().getByIndex(0).Width = (w - 4) * 100
    dp = lay.getDrawPage()
    For i = 0 To UBound(d)
        e = NormEI(CStr(d(i)(0)))
        If e <> "" Then
            n = CLng(Mid(e, 4))
            row = st.getCellRangeByPosition(0, n, 9, n).getDataArray()(0)
            If CStr(row(0)) <> e Then
                nBad = nBad + 1
                If bad = "" Then bad = e & " нет в снимке"
            Else
                c = 1
                If VarType(d(i)(1)) = 5 Then c = CLng(d(i)(1))
                If c < 1 Then c = 1
                job(nJob) = Array(e, c)
                nJob = nJob + 1
                For x = 1 To c
                    PutLabel(lay, dp, r, row, w, h)
                    r = r + ROWS_PER_LABEL
                    k = k + 1
                Next x
            End If
        ElseIf CStr(d(i)(0)) <> "" Then
            nBad = nBad + 1
            If bad = "" Then bad = "«" & d(i)(0) & "» — не номер ЕИ"
        End If
    Next i
    If nJob > 0 Then
        ReDim Preserve job(nJob - 1)
        gLblJob = job
    Else
        gLblJob = Array()
    End If
    If nBad > 0 Then
        BuildFrom = "ERR:ошибок в списке: " & nBad & " (" & bad & "); построено этикеток " & k
    Else
        BuildFrom = "OK:этикеток " & k & " (" & w & "×" & h & " мм) на листе «" & SH_LAYOUT & "»"
    End If
    Exit Function
EH:
    BuildFrom = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnLblPreview(Optional oEvent As Variant)
    Dim res As String
    res = LblPreview()
    If Left(res, 3) <> "OK:" Or gTbAuto Then
        TbMsg(res)
        Exit Sub
    End If
    ' the print preview of LibreOffice: the label page exactly as it prints
    CreateUnoService("com.sun.star.frame.DispatchHelper").executeDispatch(ThisComponent.getCurrentController().getFrame(), ".uno:PrintPreview", "", 0, Array())
End Sub

' «Предпросмотр»: one label of the first EI of the list on «Макет» (the sheet shown)
Function LblPreview() As String
    Dim lst As Object, last As Long, d As Variant, i As Long, res As String
    If Not ThisComponent.Sheets.hasByName("_stock") Then
        LblPreview = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    lst = ThisComponent.Sheets.getByName(SH_LIST)
    last = ListLast(lst)
    If last < 1 Then
        LblPreview = "ERR:на листе «" & SH_LIST & "» нет ЕИ (колонка A, начиная со строки 2)"
        Exit Function
    End If
    d = lst.getCellRangeByPosition(0, 1, 0, last).getDataArray()
    For i = 0 To UBound(d)
        If CStr(d(i)(0)) <> "" Then Exit For
    Next i
    res = BuildFrom(Array(Array(d(i)(0), 1)))
    If Left(res, 3) <> "OK:" Then
        LblPreview = res
        Exit Function
    End If
    ThisComponent.getCurrentController().setActiveSheet(ThisComponent.Sheets.getByName(SH_LAYOUT))
    LblPreview = "OK:предпросмотр " & NormEI(CStr(d(i)(0))) & " на листе «" & SH_LAYOUT & "» (для печати всего списка — «Построить этикетки»)"
End Function

' ================================================================ the history of the printing, «Повторная печать»

' the history of the printing: ../WMS_Labels/history.csv («задание;дата и время;ЕИ;копий;как», UTF-8) — it lives outside the
' tool book, which is never saved with a layout of thousands of barcode shapes
Private Function HistUrl() As String
    HistUrl = TbResolve("../WMS_Labels") & "history.csv"
End Function

' the rows of the history file (job, time, EI, copies, how); Array() when none
Private Function HistRows() As Variant
    Dim sfa As Object, lines As Variant, i As Long, f As Variant, rows() As Variant, n As Long
    HistRows = Array()
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(HistUrl()) Then Exit Function
    lines = Split(Replace(TbReadText(HistUrl()), Chr(13), ""), Chr(10))
    ReDim rows(UBound(lines))
    For i = 1 To UBound(lines)
        f = Split(lines(i), ";")
        If UBound(f) >= 4 Then
            rows(n) = Array(Val(f(0)), f(1), f(2), Val(f(3)), f(4))
            n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve rows(n - 1)
    HistRows = rows
End Function

' «История печати»: the history file shown
Private Sub ShowHistory()
    Dim h As Object, last As Long, rows As Variant
    h = ThisComponent.Sheets.getByName(SH_HIST)
    last = TbLastRow(h)
    If last >= 1 Then h.getCellRangeByPosition(0, 1, 4, last).clearContents(1023)
    rows = HistRows()
    If UBound(rows) >= 0 Then h.getCellRangeByPosition(0, 1, 4, 1 + UBound(rows)).setDataArray(rows)
End Sub

' the job of the layout built last appended to the history: its number, 0 — nothing built
Private Function RecordJob(sHow As String) As Long
    Dim rows As Variant, i As Long, nJob As Long, s As String, sfa As Object, stamp As String, d As String
    If IsEmpty(gLblJob) Then Exit Function
    If UBound(gLblJob) < 0 Then Exit Function
    rows = HistRows()
    For i = 0 To UBound(rows)
        If rows(i)(0) > nJob Then nJob = rows(i)(0)
    Next i
    nJob = nJob + 1
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbResolve("../WMS_Labels")
    If Not sfa.exists(d) Then sfa.createFolder(d)
    If sfa.exists(HistUrl()) Then s = TbReadText(HistUrl()) Else s = "задание;дата и время;ЕИ;копий;как" & Chr(10)
    If Right(s, 1) <> Chr(10) Then s = s & Chr(10)
    stamp = Format(Now(), "DD.MM.YYYY HH:MM")
    For i = 0 To UBound(gLblJob)
        s = s & nJob & ";" & stamp & ";" & gLblJob(i)(0) & ";" & gLblJob(i)(1) & ";" & sHow & Chr(10)
    Next i
    TbWriteText(HistUrl(), s)
    ShowHistory()
    RecordJob = nJob
End Function

Sub BtnLblReprint(Optional oEvent As Variant)
    TbMsg(LblReprint(Val(TbSetting(12))))
End Sub

' «Повторная печать»: the list of job nJob of the history (0 — the last one) back on «Этикетки», the layout built
Function LblReprint(nJob As Long) As String
    Dim d As Variant, i As Long, rows() As Variant, n As Long, lst As Object, res As String, ll As Long
    On Error GoTo EH
    ShowHistory()
    d = HistRows()
    If UBound(d) < 0 Then
        LblReprint = "ERR:в «" & SH_HIST & "» нет заданий (" & ConvertFromURL(HistUrl()) & ")"
        Exit Function
    End If
    If nJob <= 0 Then
        For i = 0 To UBound(d)
            If VarType(d(i)(0)) = 5 Then
                If d(i)(0) > nJob Then nJob = d(i)(0)
            End If
        Next i
    End If
    ReDim rows(UBound(d))
    For i = 0 To UBound(d)
        If VarType(d(i)(0)) = 5 Then
            If d(i)(0) = nJob Then
                rows(n) = Array(d(i)(2), d(i)(3))
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then
        LblReprint = "ERR:задания № " & nJob & " нет в «" & SH_HIST & "»"
        Exit Function
    End If
    ReDim Preserve rows(n - 1)
    lst = ThisComponent.Sheets.getByName(SH_LIST)
    ll = ListLast(lst)
    If ll >= 1 Then lst.getCellRangeByPosition(0, 1, 1, ll).clearContents(1023)
    lst.getCellRangeByPosition(0, 1, 1, n).setDataArray(rows)
    res = LblBuild()
    LblReprint = IIf(Left(res, 3) = "OK:", "OK:повтор задания № " & nJob & ": " & Mid(res, 4), res)
    Exit Function
EH:
    LblReprint = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function NormEI(ByVal s As String) As String
    Dim t As String
    t = UCase(Trim(s))
    If Left(t, 2) = "ЕИ" Or Left(t, 2) = "EI" Then t = Mid(t, 3)
    If Left(t, 1) = "-" Then t = Mid(t, 2)
    If t <> "" And Len(t) <= 8 And IsNumeric(t) And InStr(t, ",") = 0 And InStr(t, ".") = 0 Then NormEI = "ЕИ-" & Right("00000000" & CLng(t), 8)
End Function

' one label from row r: EI large, name, article, place, the barcode below; the first row starts a new page
Private Sub PutLabel(lay As Object, dp As Object, r As Long, row As Variant, w As Double, h As Double)
    Dim hrow As Double, rows As Object, i As Integer, c As Object, bars As String, x As Long, j As Integer, bw As Long, y As Long
    Dim shp As Object, pt As New com.sun.star.awt.Point, sz As New com.sun.star.awt.Size, modW As Long, bcH As Long
    hrow = (h - 4) * 100 / ROWS_PER_LABEL
    rows = lay.getRows()
    For i = 0 To ROWS_PER_LABEL - 1
        rows.getByIndex(r + i).Height = hrow
    Next i
    rows.getByIndex(r).IsStartOfNewPage = True
    c = lay.getCellByPosition(0, r)
    c.setString(row(0))
    c.CharHeight = Setting(11, 20)
    c.CharWeight = 150
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    If Flag(7) Then SetLine(lay, r + 1, Left(CStr(row(1)), 60), 9)
    If Flag(8) Then SetLine(lay, r + 2, IIf(CStr(row(2)) <> "", "Арт. " & row(2), ""), 9)
    If Flag(9) Then SetLine(lay, r + 3, IIf(CStr(row(5)) <> "", "Место: " & row(5), ""), 9)
    If Flag(10) Then
        bars = Code128C(Mid(CStr(row(0)), 4))
        modW = 30
        bcH = hrow * 0.8
        x = lay.getCellByPosition(0, r + 4).Position.X + ((w - 4) * 100 - modW * 101) / 2
        y = lay.getCellByPosition(0, r + 4).Position.Y + hrow * 0.1
        For j = 1 To Len(bars)
            bw = CLng(Mid(bars, j, 1)) * modW
            If j Mod 2 = 1 Then
                shp = ThisComponent.createInstance("com.sun.star.drawing.RectangleShape")
                pt.X = x
                pt.Y = y
                sz.Width = bw
                sz.Height = bcH
                shp.setPosition(pt)
                shp.setSize(sz)
                dp.add(shp)
                shp.FillColor = RGB(0, 0, 0)
                shp.LineStyle = com.sun.star.drawing.LineStyle.NONE
                ' anchored to the page (placed by the cell positions above): shapes anchored to cells of two and more labels make
                ' LibreOffice 24.2 write an endless file when the book is saved
                shp.Anchor = lay
            End If
            x = x + bw
        Next j
    End If
End Sub

Private Sub SetLine(lay As Object, r As Long, s As String, sz As Double)
    Dim c As Object
    c = lay.getCellByPosition(0, r)
    c.setString(s)
    c.CharHeight = sz
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    c.IsTextWrapped = True
End Sub

' «Печать»: the usual print window of LibreOffice for the layout (the system printer / CUPS)
Sub BtnLblPrint(Optional oEvent As Variant)
    Dim d As Object, args(0) As New com.sun.star.beans.PropertyValue
    If Not ThisComponent.Sheets.hasByName(SH_LAYOUT) Then
        TbMsg("ERR:сначала «Построить этикетки»")
        Exit Sub
    End If
    ThisComponent.getCurrentController().setActiveSheet(ThisComponent.Sheets.getByName(SH_LAYOUT))
    RecordJob("печать")
    d = CreateUnoService("com.sun.star.frame.DispatchHelper")
    d.executeDispatch(ThisComponent.getCurrentController().getFrame(), ".uno:Print", "", 0, args())
End Sub

' «В PDF»: the layout into ../WMS_Labels/labels_<time>.pdf (to check it without a printer)
Function LblPdf() As String
    Dim d As String, sfa As Object, url As String, args(1) As New com.sun.star.beans.PropertyValue, fd(0) As New com.sun.star.beans.PropertyValue, nJob As Long
    Dim lay As Object, cur As Object
    On Error GoTo EH
    If Not ThisComponent.Sheets.hasByName(SH_LAYOUT) Then
        LblPdf = "ERR:сначала «Построить этикетки»"
        Exit Function
    End If
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbResolve("../WMS_Labels")
    If Not sfa.exists(d) Then sfa.createFolder(d)
    url = d & "labels_" & Format(Now(), "YYYYMMDD-HHMMSS") & ".pdf"
    ' exactly the labels built: a cell range of whole labels (the used area ends at the last text, the barcode row of the last
    ' label below it would be cut off)
    lay = ThisComponent.Sheets.getByName(SH_LAYOUT)
    cur = lay.createCursor()
    cur.gotoEndOfUsedArea(False)
    fd(0).Name = "Selection"
    fd(0).Value = lay.getCellRangeByPosition(0, 0, 0, (cur.getRangeAddress().EndRow \ ROWS_PER_LABEL + 1) * ROWS_PER_LABEL - 1)
    args(0).Name = "FilterName"
    args(0).Value = "calc_pdf_Export"
    args(1).Name = "FilterData"
    args(1).Value = fd()
    ThisComponent.storeToURL(url, args())
    nJob = RecordJob("PDF")
    LblPdf = "OK:" & ConvertFromURL(url) & IIf(nJob > 0, "; задание № " & nJob & " в «" & SH_HIST & "»", "")
    Exit Function
EH:
    LblPdf = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnLblPdf(Optional oEvent As Variant)
    TbMsg(LblPdf())
End Sub
