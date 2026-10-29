' WmsItem — 1.0 (этап I): вещи и их свойства. Этикетка ЕИ прямо из WMS — номер крупно, полное наименование (шрифт
' уменьшается, пока название не уместится), единица измерения и штрихкод Code 128 (рисуется прямоугольниками, без шрифтов и
' интернета) в своей полосе — на странице «Печать» размера этикетки (Ф1; 1.0.1: 40×30 под термопринтер 203 dpi);
' свойства ЕИ — серийный или инвентарный номер, «годен до», примечание — своей операцией журнала EI_PROP, «Поиск» находит
' ЕИ по серийному номеру (Ф5); инструмент «на руки» со сроком возврата из «Номенклатуры» и список «Где инструмент» (Ф4);
' резерв под получателя или объект и вопрос при выдаче сверх свободного (Ф7); фото позиции из папки «WMS_Фото» рядом с
' книгой (Ф10). Library WmsPlus.
Option Explicit

' Ф5: _PROP — one row per EI (row = EI number), written only by EI_PROP
Public Const SH_PROP = "_PROP"
Private Const PR_EI = 0
Private Const PR_SERIAL = 1
Private Const PR_EXPIRY = 2
Private Const PR_NOTE = 3
Private Const PR_LAST = 3
' Ф7: «Резервы» — the storekeeper's own sheet: EI, quantity, for whom, until, note
Public Const SH_RESERVES = "Резервы"
Private Const RS_EI = 0
Private Const RS_QTY = 1
Private Const RS_WHO = 2
Private Const RS_UNTIL = 3
Private Const RS_NOTE = 4
Private Const RS_LAST = 4
' Ф4: «Номенклатура» Q — the days a returnable thing (a tool) may stay on the hands
Public Const GD_RETDAYS = 16
' 1.0.2: «Номенклатура» R — how many pieces come in one pack: the label of such goods says «Выдача по 1 шт · в уп. 10» (one label
' on the whole delivery, the pieces are given out one by one)
Public Const GD_PACKQ = 17
' 1.0.2: the labels of one printing (the page «Печать» holds PR_MAX rows), the labels «по количеству единиц» of one position, and
' the positions asked in one window
Private Const LB_MAX_UNITS = 200
Private Const LB_PAGE = 18
' Ф1: the label — the rows of one label on «Печать» (the EI, the name, the unit, the band of the barcode), the size in a hidden
' cell of «Печать» (AH1, «40×30»)
Private Const LB_ROWS = 4
Private Const LB_SIZE_COL = 33
Private Const LB_SHAPE = "wmsbar"
' 1.0.1: the margins of the label (1/100 mm — 1 mm, 8 dots of a 203 dpi printer) and a little slack that keeps the rows (rounded
' by Calc to its own units) inside one page
Public Const LB_MARGIN = 100
Public Const LB_SLACK = 30
' Ф10: the folder of the photos next to the book
Private Const PHOTO_DIR = "WMS_Фото"
Global gItPhotoTest As String
' Ф5: «годен до» this close (days) is shown in «Внимание» before it passes
Private Const PR_SOON_DAYS = 7
' Ф4: the overdue tools of the tile, kept until the next operation or the next day
Global gItLateOk As Boolean
Global gItLateSeq As Long
Global gItLateDay As Double
Global gItLate As Long

' ================================================================ the EI of the active sheet

' the EIs the cursor stands on: a row (or the selected rows) of «Наличие», «Заказы», «Иной приход», «Выдачи», «Корректировки»,
' the card on «Поиск». Array of EI numbers (empty — none)
Function CursorEIs() As Variant
    Dim nm As String, rows As Variant, i As Long, out() As Long, n As Long, s As String, k As Long, nextEI As Long, sh As Object, c As Integer
    CursorEIs = Array()
    nextEI = CLng(SysNum(SK_NEXT_EI))
    nm = gDoc.getCurrentController().getActiveSheet().getName()
    If nm = SH_FIND Then
        s = WmsFind.FindCardEI()
        If s <> "" Then CursorEIs = Array(WmsView.EiNo(s, nextEI))
        Exit Function
    End If
    Select Case nm
    Case SH_STOCK
        c = SC_EI
    Case SH_ORDERS
        c = OC_EI
    Case SH_SPECIAL
        c = XC_EI
    Case SH_ISSUES
        c = IC_EI
    Case SH_ADJUST
        c = AC_EI
    Case Else
        Exit Function
    End Select
    rows = WmsPrint.SelRows(nm)
    If UBound(rows) < 0 Then Exit Function
    sh = gDoc.Sheets.getByName(nm)
    ReDim out(UBound(rows))
    For i = 0 To UBound(rows)
        k = WmsView.EiNo(sh.getCellByPosition(c, rows(i)).getString(), nextEI)
        If k > 0 Then
            out(n) = k
            n = n + 1
        End If
        If n >= 200 Then Exit For
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve out(n - 1)
    CursorEIs = out
End Function

' ================================================================ Ф1 the label of an EI

Sub BtnLabel(Optional oEvent As Variant)
    Dim eis As Variant, s As String
    WmsInit()
    eis = CursorEIs()
    If UBound(eis) < 0 Then
        WmsUi.UiMessage("Встаньте на строку с ЕИ («Наличие», «Заказы», «Иной приход», «Выдачи», «Корректировки») или откройте карточку ЕИ на «Поиске» " _
            & "и нажмите «Этикетка». Выделите несколько строк (весь заказ) — окно спросит число этикеток для каждой позиции.")
        Exit Sub
    End If
    s = LabelWindow(eis)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' 1.0.2: the window of the labels of the EIs eis — for every position its own number of labels (1: one label on the position;
' 0: none), «Печатать»; «По количеству единиц» — as many labels as there are units (100 pcs — 100 labels, at most LB_MAX_UNITS
' of a position) for this and the positions that follow. More than LB_PAGE positions are asked page by page. OK:… / ERR:… / SKIP:…
Function LabelWindow(eis As Variant) As String
    Dim nAll As Long, counts() As Long, p0 As Long, p1 As Long, i As Long, k As Integer, nF As Integer, labels() As String, v() As String, lists() As String
    Dim sSize As String, info As String, page As Long, pages As Long, ln As String
    nAll = UBound(eis) + 1
    ReDim counts(nAll - 1)
    sSize = LabelSize()
    pages = (nAll + LB_PAGE - 1) \ LB_PAGE
    p0 = 0
    Do While p0 < nAll
        p1 = p0 + LB_PAGE - 1
        If p1 > nAll - 1 Then p1 = nAll - 1
        nF = p1 - p0 + 1
        If p0 = 0 Then nF = nF + 1
        ReDim labels(nF - 1)
        ReDim v(nF - 1)
        ReDim lists(nF - 1)
        For i = p0 To p1
            labels(i - p0) = LbPosText(CLng(eis(i)))
            v(i - p0) = "1"
        Next i
        If p0 = 0 Then
            labels(nF - 1) = "Размер, мм"
            v(nF - 1) = sSize
        End If
        page = p0 \ LB_PAGE + 1
        info = "Позиций: " & nAll & IIf(pages > 1, " (страница " & page & " из " & pages & ")", "") & ". Справа — сколько этикеток напечатать для каждой позиции " _
            & "(1 — одна на всю позицию, 0 — не печатать). «По количеству единиц» — этикеток столько, сколько единиц пришло (до " & LB_MAX_UNITS & " на позицию)."
        k = WmsPrimeUi.UiForm2("Этикетки", info, labels, v, lists, "Печатать", "По количеству единиц")
        If k = 0 Then
            LabelWindow = "SKIP:окно закрыто"
            gUiLastMsg = LabelWindow
            Exit Function
        End If
        If p0 = 0 Then sSize = Trim(v(nF - 1))
        If k = 2 Then
            For i = p0 To nAll - 1
                counts(i) = LbUnits(CLng(eis(i)))
            Next i
            Exit Do
        End If
        For i = p0 To p1
            ln = Trim(v(i - p0))
            If ln <> "" And Not IsNumeric(ln) Then
                LabelWindow = "ERR:число этикеток для «" & labels(i - p0) & "» — целое число, например 1 или 5 (сейчас «" & ln & "»)"
                gUiLastMsg = LabelWindow
                Exit Function
            End If
            counts(i) = CLng(Val(ln))
            If counts(i) < 0 Then counts(i) = 0
            If ln = "" Then counts(i) = 1
        Next i
        p0 = p1 + 1
    Loop
    LabelWindow = LabelPrintN(eis, counts, sSize)
End Function

' «ЕИ-123 · Краска белая 0,9 л · 100 шт» for the window of the labels (the line is short: the field of a window is narrow)
Private Function LbPosText(n As Long) As String
    Dim d As Variant, q As Double, c As Double, has As Boolean, nm As String
    d = WmsIssue.StockData(n)
    nm = Trim(CStr(d(SC_NAME)))
    If Len(nm) > 22 Then nm = Left(nm, 21) & "…"
    SiCameOne(n, CStr(d(SC_EI)), c, has)
    If has Then q = c Else q = CDbl(d(SC_QTY))
    LbPosText = "ЕИ-" & n & " · " & nm & " · " & WmsIssue.QtyText(q) & " " & Trim(CStr(d(SC_UNIT)))
End Function

' the labels «по количеству единиц» of EI n: what came into it (its balance when nothing is known), rounded, at least 1, at most LB_MAX_UNITS
Function LbUnits(n As Long) As Long
    Dim d As Variant, q As Double, c As Double, has As Boolean
    d = WmsIssue.StockData(n)
    SiCameOne(n, CStr(d(SC_EI)), c, has)
    If has Then q = c Else q = CDbl(d(SC_QTY))
    LbUnits = CLng(q + 0.4999)
    If LbUnits < 1 Then LbUnits = 1
    If LbUnits > LB_MAX_UNITS Then LbUnits = LB_MAX_UNITS
End Function

' the size of the label kept in the book («40×30» — the labels of the warehouse)
Function LabelSize() As String
    Dim s As String
    On Error Resume Next
    s = Trim(gDoc.Sheets.getByName(SH_PRINT).getCellByPosition(LB_SIZE_COL, 0).getString())
    If s = "" Then s = "40×30"
    LabelSize = s
End Function

' the labels of the EIs eis (numbers), nCopies each (see LabelPrintN)
Function LabelPrint(eis As Variant, nCopies As Long, sSize As String) As String
    Dim counts() As Long, i As Long
    If UBound(eis) < 0 Then
        LabelPrint = LabelPrintN(eis, Array(), sSize)
        Exit Function
    End If
    ReDim counts(UBound(eis))
    For i = 0 To UBound(eis)
        counts(i) = nCopies
    Next i
    LabelPrint = LabelPrintN(eis, counts, sSize)
End Function

' the labels of the EIs eis (numbers), counts(i) labels of eis(i) (its own number of every position), on «Печать» in its own page style
' of the label size sSize («40×30»): one label — one page. Below the unit: «Арт.: …» for the goods of the category «Детали» and
' «S/N: …» of an EI with a serial number; «Выдача по 1 шт · в уп. 10» for the goods that come in packs of «Номенклатуры» R (or an EI
' with a packing). OK:<labels> / ERR:…
Function LabelPrintN(eis As Variant, counts As Variant, sSize As String) As String
    Dim a As Variant, w As Double, h As Double, sh As Object, ps As Object, st As Object, i As Long, j As Long, r As Long, row As Variant, n As Long
    Dim pr(0) As New com.sun.star.table.CellRangeAddress, k As Long, d As Variant, yTop As Long, goods As Variant, ex As Variant, nx As Integer, need As Long
    Dim exs() As Variant
    On Error GoTo EH
    a = Split(Replace(Replace(LCase(sSize), "x", "×"), "*", "×"), "×")
    If UBound(a) <> 1 Then
        LabelPrintN = "ERR:размер этикетки — «ширина×высота» в мм, например 40×30"
        gUiLastMsg = LabelPrintN
        Exit Function
    End If
    w = Val(Replace(Trim(a(0)), ",", "."))
    h = Val(Replace(Trim(a(1)), ",", "."))
    If w < 30 Or w > 210 Or h < 20 Or h > 297 Then
        LabelPrintN = "ERR:размер этикетки " & sSize & " — ширина 30…210 мм, высота 20…297 мм"
        gUiLastMsg = LabelPrintN
        Exit Function
    End If
    st = WmsIssue.StockSheet()
    goods = LbGoods()
    ' what every label carries (the extra lines) and how many rows the printing takes — before the page is touched
    ReDim exs(UBound(eis))
    For i = 0 To UBound(eis)
        n = eis(i)
        row = st.getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
        If CStr(row(SC_EI)) <> "" Then
            exs(i) = LbExtra(n, row, goods)
            nx = 0
            If exs(i)(0) <> "" Then nx = nx + 1
            If exs(i)(1) <> "" Then nx = nx + 1
            If UBound(counts) >= i Then need = need + CLng(counts(i)) * (LB_ROWS + nx)
            If UBound(counts) >= i Then k = k + CLng(counts(i))
        End If
    Next i
    If need > PR_MAX - 2 Then
        LabelPrintN = "ERR:этикеток " & k & " — на странице «Печать» за один раз помещается меньше (около " & (PR_MAX \ (LB_ROWS + 1)) _
            & "). Напечатайте двумя-тремя частями: выделите меньше строк."
        gUiLastMsg = LabelPrintN
        Exit Function
    End If
    k = 0
    sh = WmsPrint.PrintBegin("", "Этикетки " & Format(Now(), "YYYY-MM-DD HHMM"))
    sh.getCellByPosition(LB_SIZE_COL, 0).setString(Format(w, "0") & "×" & Format(h, "0"))
    ps = LabelStyle(w, h)
    sh.PageStyle = "WMS_Этикетка"
    sh.getColumns().getByIndex(0).Width = LbInnerW(w)
    For i = 0 To UBound(eis)
        n = eis(i)
        row = st.getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
        If CStr(row(SC_EI)) <> "" And UBound(counts) >= i Then
            ex = exs(i)
            nx = 0
            If ex(0) <> "" Then nx = nx + 1
            If ex(1) <> "" Then nx = nx + 1
            For j = 1 To counts(i)
                PutLabel(sh, r, row, w, h, yTop, CStr(ex(0)), CStr(ex(1)))
                r = r + LB_ROWS + nx
                k = k + 1
            Next j
        End If
    Next i
    If k = 0 Then
        WmsView.ViewDone(sh)
        LabelPrintN = "ERR:ни для одной позиции не указано число этикеток (больше 0)"
        gUiLastMsg = LabelPrintN
        Exit Function
    End If
    pr(0).Sheet = sh.getRangeAddress().Sheet
    pr(0).StartColumn = 0
    pr(0).StartRow = 0
    pr(0).EndColumn = 0
    pr(0).EndRow = IIf(r > 0, r - 1, 0)
    sh.setPrintAreas(pr())
    WmsView.ViewDone(sh)
    WmsView.GoCell(SH_PRINT, 0, 0)
    LabelPrintN = "OK:этикеток " & k & " (" & Format(w, "0") & "×" & Format(h, "0") & " мм)"
    gUiLastMsg = LabelPrintN
    Exit Function
EH:
    LabelPrintN = "ERR:этикетки не построены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = LabelPrintN
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' «Номенклатура» as an array (A…R) for the packs of the labels; Array() when the sheet is not there
Private Function LbGoods() As Variant
    Dim gs As Object, last As Long
    LbGoods = Array()
    If Not gDoc.Sheets.hasByName(WmsSupply.SH_GOODS) Then Exit Function
    gs = gDoc.Sheets.getByName(WmsSupply.SH_GOODS)
    last = WmsOrders.LastRow(gs)
    If last < 1 Then Exit Function
    LbGoods = gs.getCellRangeByPosition(0, 1, GD_PACKQ, last).getDataArray()
End Function

' the extra lines of the label of EI n (its stock row `row`): (1) «Арт.: …» — only for the goods of the category «Детали» — and
' «S/N: …» of an EI with a serial number; (2) the way the pieces are given out: «Выдача по 1 шт · в уп. 10» — the packing of the EI,
' otherwise «В упаковке» of its goods in «Номенклатуре» (matched by the article, else by the name)
Private Function LbExtra(n As Long, row As Variant, goods As Variant) As Variant
    Dim l1 As String, l2 As String, p As Variant, pu As String, pn As Double, i As Long, art As String, nm As String, q As Double, unit As String
    Dim cArt As Long, cName As Long
    cArt = WmsSupply.GD_ART
    cName = WmsSupply.GD_NAME
    art = Trim(CStr(row(SC_ART)))
    nm = LCase(Trim(CStr(row(SC_NAME))))
    unit = Trim(CStr(row(SC_UNIT)))
    p = PropOf(n)
    ' 1.1.1: the article stands on the label of a part (the category «Детали») and of anything that has a serial number
    If (LCase(Trim(CStr(row(SC_CAT)))) = "детали" Or CStr(p(0)) <> "") And art <> "" Then l1 = "Арт.: " & art
    If CStr(p(0)) <> "" Then l1 = l1 & IIf(l1 <> "", " · ", "") & "S/N: " & p(0)
    If WmsPack.PackOf(n, pu, pn) Then
        l2 = "Выдача по 1 " & pu & " · в 1 " & unit & " — " & WmsIssue.QtyText(pn) & " " & pu
    ElseIf UBound(goods) >= 0 Then
        For i = 0 To UBound(goods)
            If VarType(goods(i)(GD_PACKQ)) = 5 Then
                If goods(i)(GD_PACKQ) > 1 Then
                    If (art <> "" And LCase(Trim(CStr(goods(i)(cArt)))) = LCase(art)) Or LCase(Trim(CStr(goods(i)(cName)))) = nm Then
                        q = goods(i)(GD_PACKQ)
                        l2 = "Выдача по 1 " & unit & " · в уп. " & WmsIssue.QtyText(q)
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    LbExtra = Array(l1, l2)
End Function

' the page style of the labels (made once; the size follows the last labels)
Private Function LabelStyle(w As Double, h As Double) As Object
    Dim fam As Object, ps As Object
    fam = gDoc.StyleFamilies.getByName("PageStyles")
    If Not fam.hasByName("WMS_Этикетка") Then
        ps = gDoc.createInstance("com.sun.star.style.PageStyle")
        fam.insertByName("WMS_Этикетка", ps)
    End If
    ps = fam.getByName("WMS_Этикетка")
    ps.Width = w * 100
    ps.Height = h * 100
    ' 1 mm all round (8 dots of a 203 dpi printer: the bars of the barcode stay on whole dots); the column is the printable width
    ps.LeftMargin = LB_MARGIN
    ps.RightMargin = LB_MARGIN
    ps.TopMargin = LB_MARGIN
    ps.BottomMargin = LB_MARGIN
    ps.HeaderIsOn = False
    ps.FooterIsOn = False
    ps.PrintGrid = False
    ps.CenterHorizontally = False
    ps.CenterVertically = False
    ps.ScaleToPagesX = 0
    ps.ScaleToPagesY = 0
    LabelStyle = ps
End Function

Private Function LbInnerW(w As Double) As Long
    LbInnerW = CLng(w * 100) - 2 * LB_MARGIN
End Function

' one label from row r: the EI big, the full name (the font made smaller until the whole name fits its part), the unit of
' measure, the extra lines ex1 (the article of a part, the serial number) and ex2 (the way the pieces are given out) when there are
' any, and the Code 128 of the 8 digits of the EI in its own band at the bottom — nothing drawn over the text. The label takes
' LB_ROWS rows, one more for every extra line. yTop: the top of the label on the sheet (the sum of the heights of the rows above,
' kept by LabelPrintN), moved to the next label
Private Sub PutLabel(sh As Object, r As Long, row As Variant, w As Double, h As Double, ByRef yTop As Long, ex1 As String, ex2 As String)
    Dim rows As Object, c As Object, bars As String, x As Long, j As Integer, bw As Long, y As Long, dp As Object, i As Integer, nx As Integer, nr As Integer
    Dim shp As Object, pt As New com.sun.star.awt.Point, sz As New com.sun.star.awt.Size, modW As Long, bcH As Long, nMod As Long, yAbove As Long
    Dim hTot As Long, hEI As Long, hName As Long, hUnit As Long, hBar As Long, hX As Long, wIn As Long, hReal(7) As Long, ptEI As Double, rx As Long
    nx = 0
    If ex1 <> "" Then nx = nx + 1
    If ex2 <> "" Then nx = nx + 1
    nr = LB_ROWS + nx
    hTot = CLng(h * 100) - 2 * LB_MARGIN - LB_SLACK
    wIn = LbInnerW(w)
    hEI = hTot * 0.19
    hUnit = hTot * 0.12
    hBar = hTot * 0.3
    hX = hTot * IIf(nx = 2, 0.1, 0.11)
    hName = hTot - hEI - hUnit - hBar - nx * hX
    rows = sh.getRows()
    If r > 0 Then rows.getByIndex(r).IsStartOfNewPage = True
    ' the EI: one line, as big as its part allows, never wider than the label (shrunk to fit if need be)
    rows.getByIndex(r).Height = hEI
    c = sh.getCellByPosition(0, r)
    c.setString(row(SC_EI))
    ptEI = Int(hEI / 100 / 0.3528 * 0.78)
    If ptEI > 16 Then ptEI = 16
    If ptEI < 7 Then ptEI = 7
    c.CharHeight = ptEI
    c.CharWeight = 150
    LbCell(c)
    c.ShrinkToFit = True
    ' the name: the whole of it, wrapped, the font made smaller until it fits its part
    rows.getByIndex(r + 1).Height = hName
    c = sh.getCellByPosition(0, r + 1)
    LbCell(c)
    c.IsTextWrapped = True
    LbFitName(sh, r + 1, c, Trim(CStr(row(SC_NAME))), hName)
    ' the unit of measure
    rows.getByIndex(r + 2).Height = hUnit
    c = sh.getCellByPosition(0, r + 2)
    c.setString(IIf(Trim(CStr(row(SC_UNIT))) <> "", "Ед. изм.: " & Trim(CStr(row(SC_UNIT))), ""))
    c.CharHeight = 7
    LbCell(c)
    c.ShrinkToFit = True
    ' the extra lines: the article / the serial number, then how the pieces are given out (bold: it is what the storekeeper must see)
    rx = r + 3
    If ex1 <> "" Then
        rows.getByIndex(rx).Height = hX
        c = sh.getCellByPosition(0, rx)
        c.setString(ex1)
        c.CharHeight = 8
        c.CharWeight = 150
        LbCell(c)
        c.ShrinkToFit = True
        rx = rx + 1
    End If
    If ex2 <> "" Then
        rows.getByIndex(rx).Height = hX
        c = sh.getCellByPosition(0, rx)
        c.setString(ex2)
        c.CharHeight = 8
        c.CharWeight = 150
        LbCell(c)
        c.ShrinkToFit = True
        rx = rx + 1
    End If
    ' the band of the barcode: an empty row, the bars drawn inside it (placed by the heights just set, read back as Calc keeps them)
    rows.getByIndex(rx).Height = hBar
    sh.getCellByPosition(0, rx).setString("")
    For i = 0 To nr - 1
        hReal(i) = rows.getByIndex(r + i).Height
    Next i
    bars = WmsItemCode128C(Mid(CStr(row(SC_EI)), 4))
    For j = 1 To Len(bars)
        nMod = nMod + CLng(Mid(bars, j, 1))
    Next j
    ' the module: a whole number of the dots of a 203 dpi printer (0,125 mm) — 4 dots on a wide label, else 2 — with the quiet
    ' zones of 10 modules on both sides and 2 mm to spare
    modW = 50
    If (nMod + 20) * modW > wIn - 200 Then modW = 25
    bcH = hReal(nr - 1) * 0.8
    dp = sh.getDrawPage()
    x = ((wIn - nMod * modW) \ 2 \ 25) * 25
    For i = 0 To nr - 2
        yAbove = yAbove + hReal(i)
    Next i
    y = yTop + yAbove + (hReal(nr - 1) - bcH) \ 2
    For j = 1 To Len(bars)
        bw = CLng(Mid(bars, j, 1)) * modW
        If j Mod 2 = 1 Then
            shp = gDoc.createInstance("com.sun.star.drawing.RectangleShape")
            pt.X = x
            pt.Y = y
            sz.Width = bw
            sz.Height = bcH
            ' added first, anchored to the page (shapes anchored to cells make LibreOffice 24.2 write an endless file), then sized
            ' and placed: a geometry given before the shape is on the page or before its anchor is not always kept (1.0: the bars
            ' of a label were shown at the top of the sheet, over the number of the EI)
            dp.add(shp)
            shp.Anchor = sh
            shp.setSize(sz)
            shp.setPosition(pt)
            shp.Name = LB_SHAPE
            shp.FillColor = RGB(0, 0, 0)
            shp.LineStyle = com.sun.star.drawing.LineStyle.NONE
        End If
        x = x + bw
    Next j
    yTop = yTop + yAbove + hReal(nr - 1)
End Sub

Private Sub LbCell(c As Object)
    ' pure black: a thermal printer prints a grey as scattered dots
    c.CharColor = RGB(0, 0, 0)
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    c.VertJustify = com.sun.star.table.CellVertJustify.CENTER
    c.IsTextWrapped = False
End Sub

' the name in its row of height hName: the biggest font of the list with which the wrapped text needs no more height (Calc
' measures it: the optimal height of the row); the very long name at the smallest font loses its tail words for «…»
Private Sub LbFitName(sh As Object, r As Long, c As Object, s As String, hName As Long)
    Dim sizes As Variant, i As Integer, rw As Object, fit As Boolean, words As Variant, n As Integer, t As String
    sizes = Array(11, 10, 9, 8, 7, 6.5, 6, 5.5, 5)
    rw = sh.getRows().getByIndex(r)
    c.setString(s)
    For i = 0 To UBound(sizes)
        c.CharHeight = sizes(i)
        rw.OptimalHeight = True
        If rw.Height <= hName Then
            fit = True
            Exit For
        End If
    Next i
    If Not fit Then
        words = Split(s, " ")
        n = UBound(words)
        Do While n > 0 And Not fit
            n = n - 1
            ReDim Preserve words(n)
            t = Join(words, " ") & "…"
            c.setString(t)
            rw.OptimalHeight = True
            fit = (rw.Height <= hName)
        Loop
    End If
    rw.Height = hName
End Sub

' the shapes of the labels of «Печать» removed and the page style of the documents back (WmsPrint.PrintBegin calls it)
Sub LabelClear(sh As Object)
    Dim dp As Object, i As Long, rows As Object
    dp = sh.getDrawPage()
    For i = dp.getCount() - 1 To 0 Step -1
        If dp.getByIndex(i).Name = LB_SHAPE Then dp.remove(dp.getByIndex(i))
    Next i
    If sh.PageStyle <> "WMS_Print" Then
        sh.PageStyle = "WMS_Print"
        rows = sh.getRows()
        rows.getByIndex(0).Height = 452
        sh.getCellRangeByPosition(0, 0, 0, PR_MAX).Rows.Height = 452
        sh.getCellRangeByPosition(0, 0, 0, PR_MAX).Rows.IsStartOfNewPage = False
        sh.getColumns().getByIndex(0).Width = 900
    End If
End Sub

' the bar widths (modules) of Code 128 C of an even string of digits: start C, the pairs, the checksum, stop
Function WmsItemCode128C(digits As String) As String
    Dim p As Variant, s As String, i As Integer, v As Integer, sum As Long, wgt As Integer
    p = Split("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 113222 123122 " _
        & "123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 131123 131321 " _
        & "112313 132113 132311 211313 231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 " _
        & "311321 331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 " _
        & "142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 412121 111143 " _
        & "111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 211232 2331112", " ")
    s = p(105)
    sum = 105
    wgt = 1
    For i = 1 To Len(digits) Step 2
        v = CInt(Mid(digits, i, 2))
        s = s & p(v)
        sum = sum + wgt * v
        wgt = wgt + 1
    Next i
    WmsItemCode128C = s & p(sum Mod 103) & p(106)
End Function

' after a receipt of the window «Принять…»: «Напечатать этикетку?» — asked only when labels were printed in this book before
' (the size of the label is kept), of a person or of a test that chose
Sub LabelAfterReceipt(canon As String)
    Dim k As Integer, n As Long
    If gUiAuto <> 0 And Not gPrChoiceSet Then Exit Sub
    If Trim(gDoc.Sheets.getByName(SH_PRINT).getCellByPosition(LB_SIZE_COL, 0).getString()) = "" Then Exit Sub
    k = WmsPrimeUi.UiButtons("Этикетка", "Приход проведён: " & canon & ". Напечатать этикетку?", Array("Напечатать этикетку", "Не сейчас"), Array(1, 0))
    If k <> 1 Then Exit Sub
    n = WmsView.EiNo(canon, CLng(SysNum(SK_NEXT_EI)))
    If n > 0 Then LabelPrint(Array(n), 1, LabelSize())
End Sub

' ================================================================ Ф5 the properties of an EI

' the properties of EI n: Array(serial, expiry (a date or 0), note); empty strings when none
Function PropOf(n As Long) As Variant
    Dim d As Variant, x As Double
    PropOf = Array("", 0, "")
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_PROP) Then Exit Function
    d = gDoc.Sheets.getByName(SH_PROP).getCellRangeByPosition(0, n, PR_LAST, n).getDataArray()(0)
    If VarType(d(PR_EXPIRY)) = 5 Then x = d(PR_EXPIRY)
    PropOf = Array(CStr(d(PR_SERIAL)), x, CStr(d(PR_NOTE)))
EH:
End Function

' the properties for the card («сер. № 123 · годен до 01.12.2026 · примечание»), "" — none
Function PropText(n As Long) As String
    Dim p As Variant, s As String
    p = PropOf(n)
    If p(0) <> "" Then s = "сер./инв. № " & p(0)
    If p(1) > 0 Then s = s & IIf(s <> "", " · ", "") & "годен до " & Format(p(1), "DD.MM.YYYY") & IIf(p(1) < WmsOrders.Today(), " (срок прошёл)", "")
    If p(2) <> "" Then s = s & IIf(s <> "", " · ", "") & p(2)
    PropText = s
End Function

Sub BtnProps(Optional oEvent As Variant)
    Dim eis As Variant, v(2) As String, p As Variant, s As String, n As Long
    WmsInit()
    eis = CursorEIs()
    If UBound(eis) < 0 Then
        WmsUi.UiMessage("Встаньте на строку ЕИ («Наличие») или откройте его карточку на «Поиске» и нажмите «Свойства ЕИ…».")
        Exit Sub
    End If
    n = eis(0)
    p = PropOf(n)
    v(0) = p(0)
    If p(1) > 0 Then v(1) = Format(p(1), "DD.MM.YYYY")
    v(2) = p(2)
    If Not WmsPrimeUi.UiForm("Свойства ЕИ", WmsIssue.EiCanon(n) & " — " & WmsIssue.StockSheet().getCellByPosition(SC_NAME, n).getString() _
        & ". Серийный или инвентарный номер (по нему «Поиск» найдёт ЕИ), срок годности, примечание. Пусто — убрать.", _
        Array("Серийный / инвентарный №", "Годен до", "Примечание"), v, Array("", "", ""), "Записать") Then Exit Sub
    s = PropSet(WmsIssue.EiCanon(n), Trim(v(0)), Trim(v(1)), Trim(v(2)))
    If Left(s, 3) = "OK:" Then
        ' 1.0.1: the serial number and «годен до» of «Наличие» follow at once when the window was opened there
        If gDoc.getCurrentController().getActiveSheet().getName() = SH_STOCK Then StockInfoRefresh(False)
        WmsOrdersUi.UiInfo("Свойства " & WmsIssue.EiCanon(n) & " записаны.")
    ElseIf Left(s, 5) <> "SKIP:" Then
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    End If
End Sub

' the properties of EI sEI written by the operation EI_PROP (the old values in the journal too); a serial number another EI
' already has is asked about. OK:<seq> / SKIP:… / ERR:…
Function PropSet(sEI As String, sSerial As String, sExpiry As String, sNote As String) As String
    Dim n As Long, canon As String, msg As String, p As String, old As Variant, dExp As Double, why As String, other As Long, k As Integer
    WmsInit()
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_PROP) Then
        PropSet = "ERR:в книге нет листа свойств " & SH_PROP & " — книга не обновлена до 1.0"
        Exit Function
    End If
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        PropSet = "ERR:" & msg
        Exit Function
    End If
    p = WmsIssue.StockRowProblem(n, canon)
    If p <> "" Then
        PropSet = "ERR:" & p
        Exit Function
    End If
    If sExpiry <> "" Then
        If Not WmsOrders.ParseDateTextAs(sExpiry, "годен до", dExp, msg) Then
            PropSet = "ERR:" & msg
            Exit Function
        End If
        dExp = Int(dExp)
    End If
    If Len(sSerial) > 60 Or Len(sNote) > 200 Then
        PropSet = "ERR:серийный номер — до 60 знаков, примечание — до 200"
        Exit Function
    End If
    old = PropOf(n)
    If old(0) = sSerial And old(1) = dExp And old(2) = sNote Then
        PropSet = "SKIP:свойства " & canon & " не изменились"
        Exit Function
    End If
    If sSerial <> "" And LCase(sSerial) <> LCase(CStr(old(0))) Then
        other = SerialEI(sSerial)
        If other > 0 And other <> n Then
            k = WmsPrimeUi.UiButtons("Свойства ЕИ", "Серийный номер «" & sSerial & "» уже записан у " & WmsIssue.EiCanon(other) & ". Записать и этому ЕИ?", _
                Array("Не записывать", "Записать всё равно"), Array(0, 1))
            If k <> 1 Then
                PropSet = "SKIP:серийный номер уже у " & WmsIssue.EiCanon(other)
                gUiLastMsg = PropSet
                Exit Function
            End If
        End If
    End If
    why = PostingBlockReason()
    If why <> "" Then
        PropSet = why
        Exit Function
    End If
    PlanBegin("EI_PROP", canon)
    PlanField("EI", canon)
    PlanField("SERIAL", sSerial)
    PlanField("EXPIRY", IIf(dExp > 0, Format(dExp, "YYYY-MM-DD"), ""))
    PlanField("NOTE", sNote)
    PlanField("OLD_SERIAL", CStr(old(0)))
    PlanField("OLD_EXPIRY", IIf(old(1) > 0, Format(old(1), "YYYY-MM-DD"), ""))
    PlanSetValue(SH_PROP, n, PR_EI, IIf(sSerial = "" And dExp = 0 And sNote = "", "", canon), False)
    PlanSetValue(SH_PROP, n, PR_SERIAL, sSerial, False)
    PlanSetValue(SH_PROP, n, PR_EXPIRY, IIf(dExp > 0, dExp, ""), False)
    PlanSetValue(SH_PROP, n, PR_NOTE, sNote, False)
    PropSet = ApplyOperation(0, 0)
    ' 1.0.1: the serial number and «годен до» of «Наличие» follow at the next opening of the sheet (the balance did not move)
    If Left(PropSet, 3) = "OK:" Then SiMark(n)
    gUiLastMsg = PropSet
    Exit Function
EH:
    PropSet = "ERR-SYS:свойства не записаны: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = PropSet
End Function

' the EI number of serial number s (the whole cell, without regard to case); 0 — none
Function SerialEI(s As String) As Long
    Dim sh As Object, last As Long, desc As Object, found As Object, a As Variant
    On Error GoTo EH
    If Trim(s) = "" Or Not gDoc.Sheets.hasByName(SH_PROP) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_PROP)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    desc = sh.createSearchDescriptor()
    desc.SearchString = Trim(s)
    desc.SearchCaseSensitive = False
    desc.SearchWords = True
    desc.SearchRegularExpression = False
    found = sh.getCellRangeByPosition(PR_SERIAL, 1, PR_SERIAL, last).findFirst(desc)
    If IsNull(found) Then Exit Function
    If LCase(Trim(found.getString())) = LCase(Trim(s)) Then SerialEI = found.getCellAddress().Row
EH:
End Function

' the EIs with a balance whose «годен до» passed or comes within PR_SOON_DAYS days: the EI numbers (the formula engine)
Function ExpiryEIs(ByRef nItems As Long) As Variant
    Dim last As Long, P As String, Q As String, nextEI As Long
    nItems = 0
    ExpiryEIs = Array()
    If Not gDoc.Sheets.hasByName(SH_PROP) Then Exit Function
    last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_PROP))
    nextEI = CLng(SysNum(SK_NEXT_EI))
    If last > nextEI - 1 Then last = nextEI - 1
    If last < 1 Then Exit Function
    P = WmsView.ColRng(SH_PROP, PR_EXPIRY, 1, last)
    Q = WmsView.ColRng(SH_STOCK, SC_QTY, 1, last)
    ExpiryEIs = WmsView.EngList("TEXTJOIN("";"";1;IF(ISNUMBER(" & P & ")*(" & P & "<" & Format(WmsOrders.Today() + PR_SOON_DAYS + 1, "0") & ")*ISNUMBER(" _
        & Q & ")*(" & Q & ">0.0000005);ROW(" & P & ")-1;""""))", nItems)
End Function

' ================================================================ Ф7 the reserves

' the reserved quantity of EI canon for the others than sWho (live: until today or later, or no date) and the first of them
Function ReservedFor(canon As String, sWho As String, ByRef sFirst As String) As Double
    Dim sh As Object, last As Long, d As Variant, i As Long, today As Double, q As Double, k As String
    sFirst = ""
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_RESERVES) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_RESERVES)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, RS_LAST, last).getDataArray()
    today = WmsOrders.Today()
    k = WmsAsk.WhoKey(sWho)
    For i = 0 To UBound(d)
        If UCase(Trim(CStr(d(i)(RS_EI)))) = canon Then
            If VarType(d(i)(RS_QTY)) = 5 Then
                If d(i)(RS_QTY) > 0 And (CStr(d(i)(RS_UNTIL)) = "" Or (VarType(d(i)(RS_UNTIL)) = 5 And d(i)(RS_UNTIL) >= today)) Then
                    If k = "" Or WmsAsk.WhoKey(CStr(d(i)(RS_WHO))) <> k Then
                        q = q + d(i)(RS_QTY)
                        If sFirst = "" Then sFirst = WmsIssue.QtyText(d(i)(RS_QTY)) & " для «" & d(i)(RS_WHO) & "»" & IIf(VarType(d(i)(RS_UNTIL)) = 5, " до " _
                            & Format(d(i)(RS_UNTIL), "DD.MM.YYYY"), "")
                    End If
                End If
            End If
        End If
    Next i
    ReservedFor = WmsIssue.Round3(q)
EH:
End Function

' WmsIssue (before an issue): q of EI canon to sWho when bal is on the EI; more than is free of the reserves of the others — the
' question. 0 — post, 1 — not posted (msg)
Function ReserveAsk(canon As String, q As Double, bal As Double, sWho As String, ByRef msg As String) As Integer
    Dim rsv As Double, sFirst As String, free As Double, k As Integer
    If Not gDoc.Sheets.hasByName(SH_RESERVES) Then Exit Function
    rsv = ReservedFor(canon, sWho, sFirst)
    If rsv <= 0.0000005 Then Exit Function
    free = WmsIssue.Round3(bal - rsv)
    If q <= free + 0.0000005 Then Exit Function
    If WmsAsk.MemoValue("R" & canon) = "1" Then Exit Function
    k = WmsPrimeUi.UiButtons("WMS — резерв", canon & ": из остатка " & WmsIssue.QtyText(bal) & " зарезервировано " & WmsIssue.QtyText(rsv) & " (" & sFirst & "), " _
        & "свободно " & WmsIssue.QtyText(IIf(free > 0, free, 0)) & "." & Chr(10) & "Выдать " & WmsIssue.QtyText(q) & " сверх свободного?", _
        Array("Выдать из резерва", "Не выдавать"), Array(0, 1))
    If k <> 0 Then
        msg = "сверх свободного (резерв " & sFirst & ")"
        ReserveAsk = 1
    Else
        WmsAsk.MemoSet("R" & canon, "1")
    End If
End Function

' for the card: «резерв 3 (для «объект Х» до 20.10.2026) · свободно 2», "" — no reserve
Function ReserveText(n As Long, bal As Double) As String
    Dim rsv As Double, sFirst As String
    rsv = ReservedFor(WmsIssue.EiCanon(n), "", sFirst)
    If rsv > 0.0000005 Then ReserveText = "резерв " & WmsIssue.QtyText(rsv) & " (" & sFirst & ") · свободно " & WmsIssue.QtyText(IIf(bal - rsv > 0, WmsIssue.Round3(bal - rsv), 0))
End Function

Sub NavReserves(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_RESERVES, 1, 0)
End Sub

' ================================================================ Ф4 the tools on the hands

' the returnable goods of «Номенклатура» (Q > 0): the keys (by article, otherwise by name) and their days
Private Function ItTools(keys() As String, days() As Double, isArt() As Boolean) As Long
    Dim sh As Object, last As Long, d As Variant, i As Long, n As Long
    If Not gDoc.Sheets.hasByName(WmsSupply.SH_GOODS) Then Exit Function
    sh = gDoc.Sheets.getByName(WmsSupply.SH_GOODS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, GD_RETDAYS, last).getDataArray()
    ReDim keys(UBound(d))
    ReDim days(UBound(d))
    ReDim isArt(UBound(d))
    For i = 0 To UBound(d)
        If VarType(d(i)(GD_RETDAYS)) = 5 Then
            If d(i)(GD_RETDAYS) > 0 Then
                If Trim(CStr(d(i)(GD_ART))) <> "" Then
                    keys(n) = LCase(Trim(CStr(d(i)(GD_ART))))
                    isArt(n) = True
                ElseIf Trim(CStr(d(i)(GD_NAME))) <> "" Then
                    keys(n) = LCase(Trim(CStr(d(i)(GD_NAME))))
                End If
                days(n) = d(i)(GD_RETDAYS)
                If keys(n) <> "" Then n = n + 1
            End If
        End If
    Next i
    ItTools = n
End Function

' the issues of returnable goods still on the hands (issued − returned > 0): rows of Array(who, EI, name, issued, date, due,
' overdue days, sheet row); nLate — overdue ones
Function ToolsOnHands(ByRef rows As Variant, ByRef nLate As Long) As Long
    Dim keys() As String, days() As Double, isArt() As Boolean, nk As Long, ish As Object, last As Long, arts As String, names As String, i As Long
    Dim cond As String, nItems As Long, d As Variant, r As Long, k As Long, dd As Double, retq As Double, isd As Object, no As Long
    Dim out() As Variant, n As Long, today As Double, q As Double, key As String, part As Variant, nPart As Long, seen() As Boolean, found() As Long
    nLate = 0
    rows = Array()
    nk = ItTools(keys, days, isArt)
    If nk = 0 Then Exit Function
    ish = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(ish)
    If last < 1 Then Exit Function
    ' the posted issues of these goods, by the formula engine; the keys in portions (a formula holds a few thousand characters)
    ReDim seen(last)
    ReDim found(last)
    i = 0
    Do While i < nk
        arts = ""
        names = ""
        Do While i < nk And Len(arts) + Len(names) < 2500
            If isArt(i) Then
                arts = arts & IIf(arts <> "", ";", "") & """" & Replace(keys(i), """", """""") & """"
            Else
                names = names & IIf(names <> "", ";", "") & """" & Replace(keys(i), """", """""") & """"
            End If
            i = i + 1
        Loop
        cond = ""
        If arts <> "" Then cond = "ISNUMBER(MATCH(LOWER(TRIM(" & WmsView.ColRng(SH_ISSUES, IC_ART, 1, last) & "));{" & arts & "};0))"
        If names <> "" Then cond = cond & IIf(cond <> "", "+", "") & "ISNUMBER(MATCH(LOWER(TRIM(" & WmsView.ColRng(SH_ISSUES, IC_NAME, 1, last) _
            & "));{" & names & "};0))"
        cond = "(" & cond & ")*" & WmsView.PostedExpr(WmsView.ColRng(SH_ISSUES, IC_CTL, 1, last))
        part = WmsView.EngList("TEXTJOIN("";"";1;IF(" & cond & ";ROW(" & WmsView.ColRng(SH_ISSUES, IC_CTL, 1, last) & ")-1;""""))", nPart)
        For k = 0 To nPart - 1
            r = CLng(part(k))
            If Not seen(r) Then
                seen(r) = True
                found(nItems) = r
                nItems = nItems + 1
            End If
        Next k
    Loop
    If nItems = 0 Then Exit Function
    isd = gDoc.Sheets.getByName(SH_ISS)
    today = WmsOrders.Today()
    ReDim out(nItems - 1)
    For i = 0 To nItems - 1
        r = found(i)
        d = ish.getCellRangeByPosition(0, r, IC_LAST, r).getDataArray()(0)
        If VarType(d(IC_NO)) = 5 And VarType(d(IC_QTY)) = 5 Then
            no = CLng(d(IC_NO))
            retq = 0
            If isd.getCellByPosition(IS_RET, no).getType() = com.sun.star.table.CellContentType.VALUE Then retq = isd.getCellByPosition(IS_RET, no).getValue()
            q = WmsIssue.Round3(d(IC_QTY) - retq)
            If q > 0.0000005 Then
                dd = 0
                For k = 0 To nk - 1
                    If isArt(k) Then
                        key = LCase(Trim(CStr(d(IC_ART))))
                    Else
                        key = LCase(Trim(CStr(d(IC_NAME))))
                    End If
                    If key = keys(k) Then
                        dd = days(k)
                        Exit For
                    End If
                Next k
                out(n) = Array(CStr(d(IC_WHO)), CStr(d(IC_EI)), CStr(d(IC_NAME)), WmsIssue.QtyText(q) & " " & d(IC_UNIT), WmsView.DateOf(d(IC_DATE)), _
                    WmsView.DateOf(d(IC_DATE)) + dd, IIf(WmsView.DateOf(d(IC_DATE)) + dd < today, CLng(today - WmsView.DateOf(d(IC_DATE)) - dd), 0), r)
                If WmsView.DateOf(d(IC_DATE)) + dd < today Then nLate = nLate + 1
                n = n + 1
            End If
        End If
    Next i
    If n > 0 Then
        ReDim Preserve out(n - 1)
        rows = out
    End If
    ToolsOnHands = n
End Function

' the overdue tools for the tile of «Главная» (counted again after an operation or on the next day)
Function ToolsLate() As Long
    Dim seq As Long, rows As Variant, nLate As Long
    seq = CLng(SysNum(SK_LAST_SEQ))
    If gItLateOk And gItLateSeq = seq And gItLateDay = WmsOrders.Today() Then
        ToolsLate = gItLate
        Exit Function
    End If
    ToolsOnHands(rows, nLate)
    gItLate = nLate
    gItLateSeq = seq
    gItLateDay = WmsOrders.Today()
    gItLateOk = True
    ToolsLate = nLate
End Function

Sub BtnTools(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = ToolsShow()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsView.GoCell(SH_FIND, FD_TABLE, 1)
End Sub

' «Где инструмент»: every returnable thing on the hands on «Поиск», the overdue first; a double click — the row of the issue
Function ToolsShow() As String
    Dim rows As Variant, n As Long, nLate As Long, keys() As Double, ord As Variant, i As Long, t() As Variant, x As Variant, today As Double
    n = ToolsOnHands(rows, nLate)
    If n = 0 Then
        ToolsShow = WmsFind.ShowTable("Где инструмент: на руках ничего нет", "Возвратные позиции — колонка Q «Вернуть через, дней» на листе «" & WmsSupply.SH_GOODS _
            & "».", "", Array("Кто", "ЕИ", "Наименование", "Количество", "Выдано", "Вернуть до", "Просрочено, дней", ""), Array(), 0, False)
        Exit Function
    End If
    today = WmsOrders.Today()
    ReDim keys(n - 1)
    For i = 0 To n - 1
        keys(i) = -rows(i)(6) * 100000 + rows(i)(5)
    Next i
    ord = WmsLedger.SortByCalc(keys, n)
    ReDim t(n - 1)
    For i = 0 To n - 1
        x = rows(ord(i))
        t(i) = Array(x(0), x(1), x(2), x(3), WmsView.ViewDay(x(4)), WmsView.ViewDay(x(5)), IIf(x(6) > 0, x(6), ""), "", "", SH_ISSUES, x(7), "")
    Next i
    ToolsShow = WmsFind.ShowTable("Где инструмент: на руках " & n & IIf(nLate > 0, ", просрочен возврат " & nLate, ""), "Возвратные позиции — колонка Q «Вернуть " _
        & "через, дней» на листе «" & WmsSupply.SH_GOODS & "». Двойной щелчок — строка выдачи.", "Просроченные — первыми.", Array("Кто", "ЕИ", "Наименование", _
        "Количество", "Выдано", "Вернуть до", "Просрочено, дней", ""), t, IIf(n > FD_MAX, FD_MAX, n), False)
End Function

' ================================================================ Ф10 the photo of a position

Sub BtnPhoto(Optional oEvent As Variant)
    Dim eis As Variant, s As String
    WmsInit()
    eis = CursorEIs()
    If UBound(eis) < 0 Then
        WmsUi.UiMessage("Откройте карточку ЕИ на «Поиске» (или встаньте на строку «Наличия») и нажмите «Фото».")
        Exit Sub
    End If
    s = PhotoOpen(eis(0))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the photo of EI n: the file named in «Номенклатура» K of its goods, otherwise «<артикул>.jpg/.png» or «<ЕИ>.jpg/.png» in the
' folder «WMS_Фото» next to the book; opened by the program of the computer. OK:<path> / ERR:… (how to add one)
Function PhotoOpen(n As Long) As String
    Dim st As Variant, sFolder As String, cand() As String, nc As Integer, i As Integer, f As String, ext As Variant, j As Integer, gs As Object, last As Long
    Dim d As Variant, k As Long, sh As Object
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        PhotoOpen = "ERR:книга не сохранена в файл — папку фото не найти"
        Exit Function
    End If
    st = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
    sFolder = gBaseDir & PHOTO_DIR & "/"
    If Not gSfa.exists(sFolder) Then gSfa.createFolder(sFolder)
    ReDim cand(9)
    ' the goods of «Номенклатура»: the file of K
    If gDoc.Sheets.hasByName(WmsSupply.SH_GOODS) Then
        gs = gDoc.Sheets.getByName(WmsSupply.SH_GOODS)
        last = WmsOrders.LastRow(gs)
        If last >= 1 Then
            d = gs.getCellRangeByPosition(0, 1, GD_PHOTO, last).getDataArray()
            For k = 0 To UBound(d)
                If Trim(CStr(d(k)(GD_PHOTO))) <> "" Then
                    If (CStr(st(SC_ART)) <> "" And LCase(Trim(CStr(d(k)(GD_ART)))) = LCase(Trim(CStr(st(SC_ART))))) _
                        Or LCase(Trim(CStr(d(k)(GD_NAME)))) = LCase(Trim(CStr(st(SC_NAME)))) Then
                        cand(nc) = Trim(CStr(d(k)(GD_PHOTO)))
                        nc = nc + 1
                        Exit For
                    End If
                End If
            Next k
        End If
    End If
    ext = Array(".jpg", ".jpeg", ".png")
    For j = 0 To 2
        If CStr(st(SC_ART)) <> "" Then
            cand(nc) = WmsPrint.FileSafe(CStr(st(SC_ART))) & ext(j)
            nc = nc + 1
        End If
        cand(nc) = CStr(st(SC_EI)) & ext(j)
        nc = nc + 1
    Next j
    For i = 0 To nc - 1
        f = sFolder & cand(i)
        If gSfa.exists(f) Then
            PhotoOpen = "OK:" & ConvertFromURL(f)
            gUiLastMsg = PhotoOpen
            If gUiAuto = 0 Then CreateUnoService("com.sun.star.system.SystemShellExecute").execute(f, "", 0)
            Exit Function
        End If
    Next i
    PhotoOpen = "ERR:фото " & st(SC_EI) & " нет. Положите снимок в папку «" & PHOTO_DIR & "» рядом с книгой с именем «" & IIf(CStr(st(SC_ART)) <> "", _
        WmsPrint.FileSafe(CStr(st(SC_ART))) & ".jpg» (по артикулу) или «", "") & st(SC_EI) & ".jpg», или впишите имя файла в K «" & WmsSupply.SH_GOODS & "»." _
        & Chr(10) & "Папка: " & ConvertFromURL(sFolder)
    gUiLastMsg = PhotoOpen
    Exit Function
EH:
    PhotoOpen = "ERR:фото не открыто: " & Error$
    gUiLastMsg = PhotoOpen
End Function

' ================================================================ 1.0.1 the information of «Наличие» (L…V)
' How much came in and how much is left, side by side (L «Пришло», M «Осталось» — the balance E), for whom the goods were
' ordered, where they came from (the supplier or the kind of the receipt and who passed them, the date, the document, the №
' of the order), the packing, the serial number, «годен до», the reserve. View only: filled from the sheets (no operation,
' nothing journaled, never inside a posting) when «Наличие» is opened after an operation and by «Обновить сведения». An
' opening computes only the rows that moved (SiLight: Calc itself finds them, 0,1 s on 20 000 EIs); the whole registry — the
' first fill of a book, the first opening of a session of a small registry, the button. Only the rows that differ are
' written (a write into «Наличие» is slow: its conditional formats and its filter). The constants of the other library are
' copied into locals before the loops: each use of one costs about 30 µs in LibreOffice Basic (a loop over 20 000 rows —
' seconds instead of tenths)
Public Const SI_COL0 = 11
Public Const SI_COLS = 11
Private Const SI_AUTO_MAX = 3000
' this session of WMS has filled the whole registry once (later openings take only the rows that moved); the EIs whose
' information changed without a movement (SiMark)
Global gSiSession As Boolean
Global gSiDirty As String

Sub OnStockFocus(Optional oEvent As Variant)
    On Error Resume Next
    StockInfoRefresh(False)
End Sub

Sub BtnStockInfo(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = StockInfoRefresh(True)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' OK:<rows written> — filled (the whole registry, or the rows that moved since: SiLight); SAME — nothing changed since the
' last fill; ERR:…
Function StockInfoRefresh(bForce As Boolean) As String
    Dim st As Object, last As Long, sh As Object, oLast As Long, i As Long, n As Long, s As String, stamp As String, um As Object
    Dim rowOf() As Long, spOf() As Long, cV As Variant, cA As Variant, cC As Variant, cL As Variant, cN As Variant, cT As Variant
    Dim xE As Variant, xT As Variant, xW As Variant, xD As Variant, xC As Variant, sp As Object, xLast As Long, pk As Variant, pr As Variant
    Dim rs As Variant, resQ() As Double, resWho() As String, resN() As Long, canon As String, msg As String, out() As Variant, stk As Variant
    Dim a(10) As Variant, k As Integer, locked As Boolean, came() As Double, hasCame() As Boolean, cur As Variant, nChg As Long
    Dim kEI As Integer, kName As Integer, kUnit As Integer, kQty As Integer, kPrEi As Integer, kPrSer As Integer, kPrExp As Integer, ei As String
    Dim kRsEi As Integer, kRsQty As Integer, kRsWho As Integer, kRsUntil As Integer, nPk As Long, nPr As Long, r0 As Long, r1 As Long, blk() As Variant
    Dim rw As Variant, rc As Variant, rp As Variant, x As Variant, j As Long, pos() As Long, old As String, p0 As Variant, p1 As Variant, bRes As Boolean
    Dim bLight As Boolean, res As String
    WmsInit()
    On Error GoTo EH
    st = gDoc.Sheets.getByName(SH_STOCK)
    last = WmsOrders.LastRow(st)
    If last < 1 Then
        StockInfoRefresh = "OK:0"
        Exit Function
    End If
    rs = Array()
    If gDoc.Sheets.hasByName(SH_RESERVES) Then
        sh = gDoc.Sheets.getByName(SH_RESERVES)
        i = WmsOrders.LastRow(sh)
        If i >= 1 Then rs = sh.getCellRangeByPosition(0, 1, RS_UNTIL, i).getDataArray()
    End If
    stamp = Format(SysNum(SK_LAST_SEQ), "0") & "|" & last & "|" & SiReserveStamp(rs) & "|" & gDoc.getURL()
    If Not bForce Then
        old = WmsSafe.StGet("STOCKINFO")
        If old = stamp Then
            StockInfoRefresh = "SAME"
            Exit Function
        End If
        ' only the rows that moved, when this book was filled before (the first fill of a book — the whole registry, once)
        ' and either this session filled it already (its first opening — the whole registry again: the edits of «Заказы»
        ' without an operation) or the registry is big (then the whole — by the button); a change of the reserves of a small
        ' registry — the whole
        p0 = Split(old, "|")
        p1 = Split(stamp, "|")
        bRes = True
        If UBound(p0) >= 3 Then bRes = (p0(2) <> p1(2))
        bLight = (old <> "" And Right(old, Len(gDoc.getURL()) + 1) = "|" & gDoc.getURL()) And (gSiSession Or last > SI_AUTO_MAX) _
            And (Not bRes Or last > SI_AUTO_MAX)
        If bLight Then
            res = SiLight(st, last, rs, bRes)
            If Left(res, 3) = "OK:" Then
                WmsSafe.StSet("STOCKINFO", stamp)
                gSiDirty = ""
            End If
            If Left(res, 3) = "OK:" Or Left(res, 4) = "ERR:" Then
                StockInfoRefresh = res
                Exit Function
            End If
            ' too many rows moved for one by one (a batch), or Calc did not answer: the whole registry
        End If
    End If
    kEI = SC_EI
    kName = SC_NAME
    kUnit = SC_UNIT
    kQty = SC_QTY
    kPrEi = PR_EI
    kPrSer = PR_SERIAL
    kPrExp = PR_EXPIRY
    kRsEi = RS_EI
    kRsQty = RS_QTY
    kRsWho = RS_WHO
    kRsUntil = RS_UNTIL
    ReDim rowOf(last) As Long
    ReDim spOf(last) As Long
    ' «Заказы»: the receipt row of each EI — the first row holding its key, as WMS finds it
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    oLast = WmsOrders.LastRow(sh)
    If oLast >= 1 Then
        cV = sh.getCellRangeByPosition(OC_EI, 1, OC_EI, oLast).getDataArray()
        For i = 0 To UBound(cV)
            s = CStr(cV(i)(0))
            If Left(s, 3) = "ЕИ-" Then
                n = Val(Mid(s, 4))
                If n >= 1 And n <= last Then
                    If rowOf(n) = 0 Then rowOf(n) = i + 1
                End If
            End If
        Next i
        cA = sh.getCellRangeByPosition(OC_ORDER, 1, OC_ORDER, oLast).getDataArray()
        cC = sh.getCellRangeByPosition(OC_DOC, 1, OC_DOC, oLast).getDataArray()
        cL = sh.getCellRangeByPosition(OC_SUPPLIER, 1, OC_SUPPLIER, oLast).getDataArray()
        cN = sh.getCellRangeByPosition(OC_RDATE, 1, OC_RDATE, oLast).getDataArray()
        cT = sh.getCellRangeByPosition(OC_ASSIGNED, 1, OC_ASSIGNED, oLast).getDataArray()
    End If
    ' «Иной приход»: the kind of the receipt and who passed the goods
    If gDoc.Sheets.hasByName(SH_SPECIAL) Then
        sp = gDoc.Sheets.getByName(SH_SPECIAL)
        xLast = WmsOrders.LastRow(sp)
        If xLast >= 1 Then
            xE = sp.getCellRangeByPosition(XC_EI, 1, XC_EI, xLast).getDataArray()
            For i = 0 To UBound(xE)
                s = CStr(xE(i)(0))
                If Left(s, 3) = "ЕИ-" Then
                    n = Val(Mid(s, 4))
                    If n >= 1 And n <= last Then
                        If spOf(n) = 0 And rowOf(n) = 0 Then spOf(n) = i + 1
                    End If
                End If
            Next i
            xT = sp.getCellRangeByPosition(XC_TYPE, 1, XC_TYPE, xLast).getDataArray()
            xW = sp.getCellRangeByPosition(XC_WHO, 1, XC_WHO, xLast).getDataArray()
            xD = sp.getCellRangeByPosition(XC_DATE, 1, XC_DATE, xLast).getDataArray()
            xC = sp.getCellRangeByPosition(XC_DOC, 1, XC_DOC, xLast).getDataArray()
        End If
    End If
    ' the packing (_PACK) and the properties (_PROP): one row per EI, the row of the EI
    pk = Array()
    If gDoc.Sheets.hasByName(SH_PACK) Then
        n = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_PACK))
        If n > last Then n = last
        If n >= 1 Then pk = gDoc.Sheets.getByName(SH_PACK).getCellRangeByPosition(0, 1, 2, n).getDataArray()
    End If
    nPk = UBound(pk)
    pr = Array()
    If gDoc.Sheets.hasByName(SH_PROP) Then
        n = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_PROP))
        If n > last Then n = last
        If n >= 1 Then pr = gDoc.Sheets.getByName(SH_PROP).getCellRangeByPosition(0, 1, kPrExp, n).getDataArray()
    End If
    nPr = UBound(pr)
    ' the reserves: the sum and the first «for whom» of every EI
    ReDim resQ(last) As Double
    ReDim resWho(last) As String
    ReDim resN(last) As Long
    For i = 0 To UBound(rs)
        If Trim(CStr(rs(i)(kRsEi))) <> "" And VarType(rs(i)(kRsQty)) = 5 Then
            If WmsIssue.NormalizeEI(rs(i)(kRsEi), n, canon, msg) Then
                If n >= 1 And n <= last Then
                    resQ(n) = resQ(n) + rs(i)(kRsQty)
                    resN(n) = resN(n) + 1
                    If resWho(n) = "" Then
                        resWho(n) = Trim(CStr(rs(i)(kRsWho)))
                        If VarType(rs(i)(kRsUntil)) = 5 Then resWho(n) = resWho(n) & " до " & Format(rs(i)(kRsUntil), "DD.MM.YYYY")
                    End If
                End If
            End If
        End If
    Next i
    ' how much came in: the receipts, the transfer, the splits and the kits
    ReDim came(last) As Double
    ReDim hasCame(last) As Boolean
    SiCame(last, came, hasCame)
    stk = st.getCellRangeByPosition(0, 1, kQty, last).getDataArray()
    cur = st.getCellRangeByPosition(SI_COL0, 1, SI_COL0 + SI_COLS - 1, last).getDataArray()
    ' every row computed and compared with what the sheet shows (in the same pass: the row of the sheet once); only the rows
    ' that differ are kept (out, their rows in pos)
    ReDim out(last - 1)
    ReDim pos(last) As Long
    For i = 1 To last
        rw = stk(i - 1)
        ei = CStr(rw(kEI))
        For k = 0 To 10
            a(k) = ""
        Next k
        If ei <> "" Then
            n = i
            If hasCame(n) Then a(0) = Int(came(n) * 1000 + 0.5) / 1000
            a(1) = rw(kQty)
            If rowOf(n) > 0 Then
                j = rowOf(n) - 1
                a(2) = cT(j)(0)
                a(3) = cL(j)(0)
                a(4) = cN(j)(0)
                a(5) = cC(j)(0)
                a(6) = cA(j)(0)
            ElseIf spOf(n) > 0 Then
                j = spOf(n) - 1
                s = Trim(CStr(xT(j)(0)))
                If Trim(CStr(xW(j)(0))) <> "" Then s = s & IIf(s <> "", " · ", "") & Trim(CStr(xW(j)(0)))
                a(3) = s
                a(4) = xD(j)(0)
                a(5) = xC(j)(0)
            End If
            If nPk >= i - 1 Then
                rp = pk(i - 1)
                If CStr(rp(0)) = ei And VarType(rp(2)) = 5 Then
                    If rp(2) > 0 Then a(7) = "в 1 " & rw(kUnit) & " — " & WmsIssue.QtyText(rp(2)) & " " & rp(1)
                End If
            End If
            If nPr >= i - 1 Then
                rp = pr(i - 1)
                If CStr(rp(kPrEi)) = ei Then
                    a(8) = rp(kPrSer)
                    a(9) = rp(kPrExp)
                End If
            End If
            If resN(n) > 0 Then a(10) = WmsIssue.QtyText(resQ(n)) & " · " & resWho(n) & IIf(resN(n) > 1, " и ещё " & (resN(n) - 1), "")
        End If
        rc = cur(i - 1)
        For k = 0 To 10
            x = rc(k)
            If VarType(x) <> VarType(a(k)) Then
                Exit For
            ElseIf x <> a(k) Then
                Exit For
            End If
        Next k
        If k <= 10 Then
            out(nChg) = Array(a(0), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10))
            pos(i) = nChg + 1
            nChg = nChg + 1
        End If
    Next i
    If nChg > 0 Then
        ' outside the undo of the user and without marking the book changed when it was not; the changed rows by runs (an
        ' unchanged row between two changed ones close together joins the run with what the sheet already shows)
        st = WmsView.ViewOpen(SH_STOCK)
        um = gDoc.getUndoManager()
        um.lock()
        locked = True
        i = 1
        Do While i <= last
            If pos(i) > 0 Then
                r0 = i
                r1 = i
                j = i + 1
                Do While j <= last And j - r1 <= 8
                    If pos(j) > 0 Then r1 = j
                    j = j + 1
                Loop
                ReDim blk(r1 - r0)
                For k = 0 To r1 - r0
                    If pos(r0 + k) > 0 Then blk(k) = out(pos(r0 + k) - 1) Else blk(k) = cur(r0 + k - 1)
                Next k
                st.getCellRangeByPosition(SI_COL0, r0, SI_COL0 + SI_COLS - 1, r1).setDataArray(blk())
                i = r1 + 1
            Else
                i = i + 1
            End If
        Loop
        um.unlock()
        locked = False
        WmsView.ViewDone(st)
    End If
    WmsSafe.StSet("STOCKINFO", stamp)
    gSiSession = True
    gSiDirty = ""
    StockInfoRefresh = "OK:" & nChg
    Exit Function
EH:
    StockInfoRefresh = "ERR:сведения «" & SH_STOCK & "» не обновлены: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    If locked Then um.unlock()
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_STOCK))
End Function

' ---------------------------------------------------------------- the light refresh: only the rows that moved

' an EI whose information WMS changed without moving its balance (its properties, its packing): its row is taken by the next
' light refresh
Sub SiMark(n As Long)
    gSiDirty = gSiDirty & ";" & n
End Sub

' the rows of «Наличие» to compute again, found by Calc itself (one formula over the registry, no loop of Basic): «Осталось»
' differs from «Остаток» (an operation moved the EI, a new EI has none yet), and the EIs marked by SiMark; with bRes — also
' the reserved ones and those that showed a reserve. OK:<rows written> / FULL — too many rows for one by one / "" — Calc did
' not answer (the caller computes everything)
Private Function SiLight(st As Object, last As Long, rs As Variant, bRes As Boolean) As String
    Dim cA As String, cE As String, cM As String, cV As String, body As String, ok As Boolean, s As String, items As Variant, i As Long, n As Long
    Dim seen As String, rows() As Long, nR As Long, a As Variant, rc As Variant, k As Integer, nW As Long, um As Object, locked As Boolean
    Dim opened As Boolean, canon As String, msg As String, t As String
    On Error GoTo EH
    cA = WmsView.ColRng(SH_STOCK, 0, 1, last)
    cE = WmsView.ColRng(SH_STOCK, SC_QTY, 1, last)
    cM = WmsView.ColRng(SH_STOCK, SI_COL0 + 1, 1, last)
    body = "TEXTJOIN("";"";1;IF((" & cA & "<>"""")*IF(ISNUMBER(" & cE & ");IF(ISNUMBER(" & cM & ");" & cM & "<>" & cE & ";1);ISNUMBER(" & cM _
        & "));ROW(" & cA & ")-1;""""))"
    s = WmsOrders.ScratchEval(body, ok)
    If Not ok Then Exit Function
    If bRes Then
        cV = WmsView.ColRng(SH_STOCK, SI_COL0 + SI_COLS - 1, 1, last)
        t = WmsOrders.ScratchEval("TEXTJOIN("";"";1;IF(" & cV & "<>"""";ROW(" & cV & ")-1;""""))", ok)
        If Not ok Then Exit Function
        If t <> "" Then s = s & IIf(s <> "", ";", "") & t
        For i = 0 To UBound(rs)
            If WmsIssue.NormalizeEI(rs(i)(RS_EI), n, canon, msg) Then s = s & IIf(s <> "", ";", "") & n
        Next i
    End If
    s = s & gSiDirty
    items = Split(s, ";")
    ReDim rows(UBound(items) + 1) As Long
    seen = ";"
    For i = 0 To UBound(items)
        If items(i) <> "" Then
            n = Val(items(i))
            If n >= 1 And n <= last And InStr(seen, ";" & n & ";") = 0 Then
                seen = seen & n & ";"
                rows(nR) = n
                nR = nR + 1
            End If
        End If
    Next i
    If nR > 400 Then
        SiLight = "FULL"
        Exit Function
    End If
    For i = 0 To nR - 1
        n = rows(i)
        a = SiRowOne(st, n, rs)
        rc = st.getCellRangeByPosition(SI_COL0, n, SI_COL0 + SI_COLS - 1, n).getDataArray()(0)
        For k = 0 To 10
            If VarType(rc(k)) <> VarType(a(k)) Then
                Exit For
            ElseIf rc(k) <> a(k) Then
                Exit For
            End If
        Next k
        If k <= 10 Then
            If Not opened Then
                ' outside the undo of the user and without marking the book changed when it was not
                st = WmsView.ViewOpen(SH_STOCK)
                opened = True
                um = gDoc.getUndoManager()
                um.lock()
                locked = True
            End If
            st.getCellRangeByPosition(SI_COL0, n, SI_COL0 + SI_COLS - 1, n).setDataArray(Array(a))
            nW = nW + 1
        End If
    Next i
    If locked Then
        um.unlock()
        locked = False
    End If
    If opened Then WmsView.ViewDone(st)
    SiLight = "OK:" & nW
    Exit Function
EH:
    SiLight = "ERR:сведения «" & SH_STOCK & "» не обновлены: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    If locked Then um.unlock()
    If opened Then WmsView.ViewDone(st)
End Function

' the information L…V of EI n alone, as the full pass computes it (the same sources, read for this row only)
Private Function SiRowOne(st As Object, n As Long, rs As Variant) As Variant
    Dim a(10) As Variant, d As Variant, ei As String, k As Integer, r As Long, o As Variant, c As Double, has As Boolean, s As String
    Dim i As Long, m As Long, canon As String, msg As String, resQ As Double, resN As Long, resWho As String
    For k = 0 To 10
        a(k) = ""
    Next k
    d = st.getCellRangeByPosition(0, n, SC_QTY, n).getDataArray()(0)
    ei = CStr(d(SC_EI))
    If ei <> "" Then
        SiCameOne(n, ei, c, has)
        If has Then a(0) = Int(c * 1000 + 0.5) / 1000
        a(1) = d(SC_QTY)
        r = SiFindRow(SH_ORDERS, OC_EI, n)
        If r > 0 Then
            o = gDoc.Sheets.getByName(SH_ORDERS).getCellRangeByPosition(0, r, OC_EI, r).getDataArray()(0)
            a(2) = o(OC_ASSIGNED)
            a(3) = o(OC_SUPPLIER)
            a(4) = o(OC_RDATE)
            a(5) = o(OC_DOC)
            a(6) = o(OC_ORDER)
        ElseIf gDoc.Sheets.hasByName(SH_SPECIAL) Then
            r = SiFindRow(SH_SPECIAL, XC_EI, n)
            If r > 0 Then
                o = gDoc.Sheets.getByName(SH_SPECIAL).getCellRangeByPosition(0, r, XC_EI, r).getDataArray()(0)
                s = Trim(CStr(o(XC_TYPE)))
                If Trim(CStr(o(XC_WHO))) <> "" Then s = s & IIf(s <> "", " · ", "") & Trim(CStr(o(XC_WHO)))
                a(3) = s
                a(4) = o(XC_DATE)
                a(5) = o(XC_DOC)
            End If
        End If
        If gDoc.Sheets.hasByName(SH_PACK) Then
            o = gDoc.Sheets.getByName(SH_PACK).getCellRangeByPosition(0, n, 2, n).getDataArray()(0)
            If CStr(o(0)) = ei And VarType(o(2)) = 5 Then
                If o(2) > 0 Then a(7) = "в 1 " & d(SC_UNIT) & " — " & WmsIssue.QtyText(o(2)) & " " & o(1)
            End If
        End If
        If gDoc.Sheets.hasByName(SH_PROP) Then
            o = gDoc.Sheets.getByName(SH_PROP).getCellRangeByPosition(0, n, PR_EXPIRY, n).getDataArray()(0)
            If CStr(o(PR_EI)) = ei Then
                a(8) = o(PR_SERIAL)
                a(9) = o(PR_EXPIRY)
            End If
        End If
        For i = 0 To UBound(rs)
            If Trim(CStr(rs(i)(RS_EI))) <> "" And VarType(rs(i)(RS_QTY)) = 5 Then
                If WmsIssue.NormalizeEI(rs(i)(RS_EI), m, canon, msg) Then
                    If m = n Then
                        resQ = resQ + rs(i)(RS_QTY)
                        resN = resN + 1
                        If resN = 1 Then
                            resWho = Trim(CStr(rs(i)(RS_WHO)))
                            If VarType(rs(i)(RS_UNTIL)) = 5 Then resWho = resWho & " до " & Format(rs(i)(RS_UNTIL), "DD.MM.YYYY")
                        End If
                    End If
                End If
            End If
        Next i
        If resN > 0 Then a(10) = WmsIssue.QtyText(resQ) & " · " & resWho & IIf(resN > 1, " и ещё " & (resN - 1), "")
    End If
    SiRowOne = Array(a(0), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10))
End Function

' how much came into EI n (canonical ei), as SiCame counts it for the whole registry
Private Sub SiCameOne(n As Long, ei As String, ByRef c As Double, ByRef has As Boolean)
    Dim d As Variant, sh As Object, rr As Variant, i As Long
    c = 0
    has = False
    d = WmsOrders.RcvSheet().getCellRangeByPosition(0, n, RV_QTY, n).getDataArray()(0)
    If CStr(d(RV_STATE)) = RV_LIVE And CStr(d(RV_EI)) <> "" And VarType(d(RV_QTY)) = 5 Then
        c = c + d(RV_QTY)
        has = True
    End If
    If gDoc.Sheets.hasByName(SH_LEG) Then
        d = WmsLegacy.LegSheet().getCellRangeByPosition(0, n, LG_BAL0, n).getDataArray()(0)
        If CStr(d(LG_KIND)) = LG_MIGRATE And VarType(d(LG_BAL0)) = 5 Then
            c = c + d(LG_BAL0)
            has = True
        End If
    End If
    If gDoc.Sheets.hasByName(SH_SPR) Then
        sh = gDoc.Sheets.getByName(SH_SPR)
        rr = SiFindAll(sh, SR_EI, "^" & ei & "$")
        For i = 0 To UBound(rr)
            d = sh.getCellRangeByPosition(0, rr(i), SR_STATE, rr(i)).getDataArray()(0)
            If CStr(d(SR_STATE)) = RV_LIVE And VarType(d(SR_QTY)) = 5 Then
                c = c + d(SR_QTY)
                has = True
            End If
        Next i
    End If
    If gDoc.Sheets.hasByName(SH_ADJ) Then
        sh = gDoc.Sheets.getByName(SH_ADJ)
        rr = SiFindAll(sh, AJ_NEW, "^" & ei & "$")
        For i = 0 To UBound(rr)
            d = sh.getCellRangeByPosition(0, rr(i), AJ_NEW, rr(i)).getDataArray()(0)
            If CStr(d(AJ_STATE)) = RV_LIVE And VarType(d(AJ_QTY)) = 5 Then
                If CStr(d(AJ_KIND)) = "SPLIT" Then
                    c = c - d(AJ_QTY)
                    has = True
                ElseIf CStr(d(AJ_KIND)) = "KIT" And CStr(d(AJ_EI)) = CStr(d(AJ_NEW)) Then
                    c = c + d(AJ_QTY)
                    has = True
                End If
            End If
        Next i
    End If
End Sub

' the first data row of sheet sName whose column col holds the key of EI n («ЕИ-…» with the number n), 0 — none
Private Function SiFindRow(sName As String, col As Integer, n As Long) As Long
    Dim rr As Variant, i As Long
    rr = SiFindAll(gDoc.Sheets.getByName(sName), col, "^ЕИ-0*" & n & "$")
    For i = 0 To UBound(rr)
        If SiFindRow = 0 Or rr(i) < SiFindRow Then SiFindRow = rr(i)
    Next i
End Function

' the data rows of sheet sh whose cell of column col matches the regular expression sRe (the search of Calc), Array() — none
Private Function SiFindAll(sh As Object, col As Integer, sRe As String) As Variant
    Dim l As Long, desc As Object, found As Object, ad As Variant, k As Long, r As Long, out() As Long, nO As Long
    SiFindAll = Array()
    l = WmsOrders.LastRow(sh)
    If l < 1 Then Exit Function
    desc = sh.createSearchDescriptor()
    desc.SearchString = sRe
    desc.SearchRegularExpression = True
    desc.SearchCaseSensitive = True
    found = sh.getCellRangeByPosition(col, 1, col, l).findAll(desc)
    If IsNull(found) Then Exit Function
    ad = found.getRangeAddresses()
    ReDim out(15)
    For k = 0 To UBound(ad)
        For r = ad(k).StartRow To ad(k).EndRow
            If nO > UBound(out) Then
                ReDim Preserve out(2 * nO)
            End If
            out(nO) = r
            nO = nO + 1
        Next r
    Next k
    If nO = 0 Then Exit Function
    ReDim Preserve out(nO - 1)
    SiFindAll = out
End Function

' how much came into EIs 1..last, in the unit of the EI: its receipt of «Заказы» (_RCV, a receipt of the transfer too — the
' whole quantity that came), a balance brought by the transfer without a receipt (MIGRATE), the posted lines of «Иной приход»
' (a part replenished — all of them), what a split moved into a new EI, what a kit was built with. A receipt or a line taken
' back (storno) does not count; hasCame(n) — WMS knows how EI n came in
Private Sub SiCame(last As Long, came() As Double, hasCame() As Boolean)
    Dim sh As Object, l As Long, d As Variant, i As Long, n As Long, k1 As Integer, k2 As Integer, k3 As Integer, k4 As Integer, sLive As String
    Dim sKind As String, s As String
    sLive = RV_LIVE
    sh = WmsOrders.RcvSheet()
    l = WmsOrders.LastRow(sh)
    If l > last Then l = last
    If l >= 1 Then
        k1 = RV_EI
        k2 = RV_STATE
        k3 = RV_QTY
        d = sh.getCellRangeByPosition(0, 1, k3, l).getDataArray()
        For i = 0 To UBound(d)
            If CStr(d(i)(k2)) = sLive Then
                If CStr(d(i)(k1)) <> "" And VarType(d(i)(k3)) = 5 Then
                    came(i + 1) = came(i + 1) + d(i)(k3)
                    hasCame(i + 1) = True
                End If
            End If
        Next i
    End If
    If gDoc.Sheets.hasByName(SH_LEG) Then
        sh = WmsLegacy.LegSheet()
        l = WmsOrders.LastRow(sh)
        If l > last Then l = last
        If l >= 1 Then
            k1 = LG_KIND
            k2 = LG_BAL0
            sKind = LG_MIGRATE
            d = sh.getCellRangeByPosition(0, 1, k2, l).getDataArray()
            For i = 0 To UBound(d)
                If CStr(d(i)(k1)) = sKind And VarType(d(i)(k2)) = 5 Then
                    came(i + 1) = came(i + 1) + d(i)(k2)
                    hasCame(i + 1) = True
                End If
            Next i
        End If
    End If
    If gDoc.Sheets.hasByName(SH_SPR) Then
        sh = gDoc.Sheets.getByName(SH_SPR)
        l = WmsOrders.LastRow(sh)
        If l >= 1 Then
            k1 = SR_EI
            k2 = SR_STATE
            k3 = SR_QTY
            d = sh.getCellRangeByPosition(0, 1, k2, l).getDataArray()
            For i = 0 To UBound(d)
                If CStr(d(i)(k2)) = sLive And VarType(d(i)(k3)) = 5 Then
                    n = WmsLegacy.EiNum(CStr(d(i)(k1)), last + 1)
                    If n > 0 Then
                        came(n) = came(n) + d(i)(k3)
                        hasCame(n) = True
                    End If
                End If
            Next i
        End If
    End If
    If gDoc.Sheets.hasByName(SH_ADJ) Then
        sh = gDoc.Sheets.getByName(SH_ADJ)
        l = WmsOrders.LastRow(sh)
        If l >= 1 Then
            k1 = AJ_KIND
            k2 = AJ_STATE
            k3 = AJ_QTY
            k4 = AJ_NEW
            d = sh.getCellRangeByPosition(0, 1, k4, l).getDataArray()
            For i = 0 To UBound(d)
                s = CStr(d(i)(k1))
                If s = "SPLIT" Or s = "KIT" Then
                    If CStr(d(i)(k2)) = sLive And VarType(d(i)(k3)) = 5 Then
                        n = WmsLegacy.EiNum(CStr(d(i)(k4)), last + 1)
                        If n > 0 And s = "SPLIT" Then
                            ' the new EI of a split: what was taken from the source
                            came(n) = came(n) - d(i)(k3)
                            hasCame(n) = True
                        ElseIf n > 0 And CStr(d(i)(AJ_EI)) = CStr(d(i)(k4)) Then
                            ' the kit itself: its own row of the build
                            came(n) = came(n) + d(i)(k3)
                            hasCame(n) = True
                        End If
                    End If
                End If
            Next i
        End If
    End If
End Sub

' the reserves as one short text: changed when a reserve is added, removed or changed
Private Function SiReserveStamp(rs As Variant) As String
    Dim i As Long, x As Double, s As String
    For i = 0 To UBound(rs)
        If VarType(rs(i)(RS_QTY)) = 5 Then x = x + rs(i)(RS_QTY) * (i + 1)
        s = s & Left(CStr(rs(i)(RS_EI)), 12) & Left(CStr(rs(i)(RS_WHO)), 4)
    Next i
    SiReserveStamp = (UBound(rs) + 1) & "/" & Format(x, "0.000") & "/" & Len(s)
End Function

' ================================================================ test seams (TEST books only)

' Т11: each string through the parsers of a quantity, of a quantity of an EI (issue, return, count) and of a place; the strings
' parsed without an internal error and without a negative quantity counted («OK:n»), the first one that broke a parser named
' (more than the balance is not the parser's to refuse: the posting refuses it)
Function TestFuzzParse(samples As Variant) As String
    Dim i As Long, n As Long, sBad As String, sWhy As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestFuzzParse = "REFUSED:не тестовая книга"
        Exit Function
    End If
    For i = LBound(samples) To UBound(samples)
        sWhy = ""
        If ItFuzzOne(CStr(samples(i)), sWhy) Then
            n = n + 1
        ElseIf sBad = "" Then
            sBad = " FIRST «" & samples(i) & "»: " & sWhy
        End If
    Next i
    TestFuzzParse = "OK:" & n & sBad
End Function

Private Function ItFuzzOne(s As String, ByRef sWhy As String) As Boolean
    Dim v As Double, q As Double, info As String, msg As String, k As Integer, mode As Variant
    On Error GoTo EH
    k = WmsCore.ParseQtyText(s, v, msg)
    For Each mode In Array("OUT", "IN", "FACT")
        k = WmsPack.QtyEIText(s, 0, 5, "шт", CStr(mode), 2, q, info, msg)
        If k = 0 And q < 0 Then
            sWhy = CStr(mode) & " дал " & q
            Exit Function
        End If
    Next mode
    msg = WmsMap.PlaceTextKey(s)
    ItFuzzOne = True
    Exit Function
EH:
    sWhy = Error$
End Function
