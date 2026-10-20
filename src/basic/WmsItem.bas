' WmsItem — 1.0 (этап I): вещи и их свойства. Этикетка ЕИ прямо из WMS — номер крупно, наименование, артикул, место и
' штрихкод Code 128 (рисуется прямоугольниками, без шрифтов и интернета) на странице «Печать» размера этикетки (Ф1);
' свойства ЕИ — серийный или инвентарный номер, «годен до», примечание — своей операцией журнала EI_PROP, «Поиск» находит
' ЕИ по серийному номеру (Ф5); инструмент «на руки» со сроком возврата из «Номенклатуры» и список «Где инструмент» (Ф4);
' резерв под получателя или объект и вопрос при выдаче сверх свободного (Ф7); фото позиции из папки «WMS_Фото» рядом с
' книгой (Ф10). Library WmsPlus.
Option Explicit

' Ф5: _PROP — one row per EI (row = EI number), written only by EI_PROP
Public Const SH_PROP = "_PROP"
Public Const PR_EI = 0
Public Const PR_SERIAL = 1
Public Const PR_EXPIRY = 2
Public Const PR_NOTE = 3
Public Const PR_LAST = 3
' Ф7: «Резервы» — the storekeeper's own sheet: EI, quantity, for whom, until, note
Public Const SH_RESERVES = "Резервы"
Public Const RS_EI = 0
Public Const RS_QTY = 1
Public Const RS_WHO = 2
Public Const RS_UNTIL = 3
Public Const RS_NOTE = 4
Public Const RS_LAST = 4
' Ф4: «Номенклатура» Q — the days a returnable thing (a tool) may stay on the hands
Public Const GD_RETDAYS = 16
' Ф1: the label — the rows of one label on «Печать», the size in a hidden cell of «Печать» (AH1, «58×40»)
Public Const LB_ROWS = 5
Public Const LB_SIZE_COL = 33
Public Const LB_SHAPE = "wmsbar"
' Ф10: the folder of the photos next to the book
Public Const PHOTO_DIR = "WMS_Фото"
Global gItPhotoTest As String
' Ф5: «годен до» this close (days) is shown in «Внимание» before it passes
Public Const PR_SOON_DAYS = 7
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
    Dim eis As Variant, v(1) As String, s As String, n As Long
    WmsInit()
    eis = CursorEIs()
    If UBound(eis) < 0 Then
        WmsUi.UiMessage("Встаньте на строку с ЕИ («Наличие», «Заказы», «Иной приход», «Выдачи», «Корректировки») или откройте карточку ЕИ на «Поиске» " _
            & "и нажмите «Этикетка». Несколько строк «Наличия» — несколько этикеток.")
        Exit Sub
    End If
    v(0) = "1"
    v(1) = LabelSize()
    If Not WmsPrimeUi.UiForm("Этикетка", "Этикеток ЕИ: " & (UBound(eis) + 1) & ". Номер крупно, наименование, артикул, место и штрихкод. Размер — как у "  _
        & "этикеток принтера (ширина×высота, мм).", Array("Копий каждой", "Размер, мм"), v, Array("", ""), "Сделать") Then Exit Sub
    n = Val(Trim(v(0)))
    s = LabelPrint(eis, IIf(n < 1, 1, n), Trim(v(1)))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the size of the label kept in the book («58×40»)
Function LabelSize() As String
    Dim s As String
    On Error Resume Next
    s = Trim(gDoc.Sheets.getByName(SH_PRINT).getCellByPosition(LB_SIZE_COL, 0).getString())
    If s = "" Then s = "58×40"
    LabelSize = s
End Function

' the labels of the EIs eis (numbers), nCopies each, on «Печать» in its own page style of the label size sSize («58×40»): one
' label — one page. OK:<labels> / ERR:…
Function LabelPrint(eis As Variant, nCopies As Long, sSize As String) As String
    Dim a As Variant, w As Double, h As Double, sh As Object, ps As Object, st As Object, i As Long, j As Long, r As Long, row As Variant, n As Long
    Dim pr(0) As New com.sun.star.table.CellRangeAddress, k As Long, d As Variant
    On Error GoTo EH
    a = Split(Replace(Replace(LCase(sSize), "x", "×"), "*", "×"), "×")
    If UBound(a) <> 1 Then
        LabelPrint = "ERR:размер этикетки — «ширина×высота» в мм, например 58×40"
        gUiLastMsg = LabelPrint
        Exit Function
    End If
    w = Val(Replace(Trim(a(0)), ",", "."))
    h = Val(Replace(Trim(a(1)), ",", "."))
    If w < 30 Or w > 210 Or h < 20 Or h > 297 Then
        LabelPrint = "ERR:размер этикетки " & sSize & " — ширина 30…210 мм, высота 20…297 мм"
        gUiLastMsg = LabelPrint
        Exit Function
    End If
    st = WmsIssue.StockSheet()
    sh = WmsPrint.PrintBegin("", "Этикетки " & Format(Now(), "YYYY-MM-DD HHMM"))
    sh.getCellByPosition(LB_SIZE_COL, 0).setString(Format(w, "0") & "×" & Format(h, "0"))
    ps = LabelStyle(w, h)
    sh.PageStyle = "WMS_Этикетка"
    sh.getColumns().getByIndex(0).Width = (w - 4) * 100
    For i = 0 To UBound(eis)
        n = eis(i)
        row = st.getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
        If CStr(row(SC_EI)) <> "" Then
            For j = 1 To nCopies
                If r + LB_ROWS > PR_MAX Then Exit For
                PutLabel(sh, r, row, w, h)
                r = r + LB_ROWS
                k = k + 1
            Next j
        End If
    Next i
    pr(0).Sheet = sh.getRangeAddress().Sheet
    pr(0).StartColumn = 0
    pr(0).StartRow = 0
    pr(0).EndColumn = 0
    pr(0).EndRow = IIf(r > 0, r - 1, 0)
    sh.setPrintAreas(pr())
    WmsView.ViewDone(sh)
    WmsView.GoCell(SH_PRINT, 0, 0)
    LabelPrint = "OK:этикеток " & k & " (" & Format(w, "0") & "×" & Format(h, "0") & " мм)"
    gUiLastMsg = LabelPrint
    Exit Function
EH:
    LabelPrint = "ERR:этикетки не построены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = LabelPrint
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
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
    ps.LeftMargin = 200
    ps.RightMargin = 200
    ps.TopMargin = 200
    ps.BottomMargin = 200
    ps.HeaderIsOn = False
    ps.FooterIsOn = False
    ps.PrintGrid = False
    ps.CenterHorizontally = True
    ps.ScaleToPagesX = 0
    ps.ScaleToPagesY = 0
    LabelStyle = ps
End Function

' one label from row r: the EI big, the name, the article, the place, the Code 128 of the 8 digits of the EI
Private Sub PutLabel(sh As Object, r As Long, row As Variant, w As Double, h As Double)
    Dim hrow As Double, rows As Object, i As Integer, c As Object, bars As String, x As Long, j As Integer, bw As Long, y As Long, dp As Object
    Dim shp As Object, pt As New com.sun.star.awt.Point, sz As New com.sun.star.awt.Size, modW As Long, bcH As Long
    hrow = (h - 4) * 100 / LB_ROWS
    rows = sh.getRows()
    For i = 0 To LB_ROWS - 1
        rows.getByIndex(r + i).Height = hrow
    Next i
    If r > 0 Then rows.getByIndex(r).IsStartOfNewPage = True
    c = sh.getCellByPosition(0, r)
    c.setString(row(SC_EI))
    c.CharHeight = 18
    c.CharWeight = 150
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    LabelLine(sh, r + 1, Left(CStr(row(SC_NAME)), 60), 9)
    LabelLine(sh, r + 2, IIf(CStr(row(SC_ART)) <> "", "Арт. " & row(SC_ART), ""), 9)
    LabelLine(sh, r + 3, IIf(CStr(row(SC_PLACE)) <> "", "Место: " & row(SC_PLACE), ""), 9)
    bars = WmsItemCode128C(Mid(CStr(row(SC_EI)), 4))
    modW = 30
    bcH = hrow * 0.8
    dp = sh.getDrawPage()
    x = sh.getCellByPosition(0, r + 4).Position.X + ((w - 4) * 100 - modW * 101) / 2
    y = sh.getCellByPosition(0, r + 4).Position.Y + hrow * 0.1
    For j = 1 To Len(bars)
        bw = CLng(Mid(bars, j, 1)) * modW
        If j Mod 2 = 1 Then
            shp = gDoc.createInstance("com.sun.star.drawing.RectangleShape")
            pt.X = x
            pt.Y = y
            sz.Width = bw
            sz.Height = bcH
            shp.setPosition(pt)
            shp.setSize(sz)
            dp.add(shp)
            shp.Name = LB_SHAPE
            shp.FillColor = RGB(0, 0, 0)
            shp.LineStyle = com.sun.star.drawing.LineStyle.NONE
            ' anchored to the page (placed by the cell positions): shapes anchored to cells make LibreOffice 24.2 write an endless file
            shp.Anchor = sh
        End If
        x = x + bw
    Next j
End Sub

Private Sub LabelLine(sh As Object, r As Long, s As String, sz As Double)
    Dim c As Object
    c = sh.getCellByPosition(0, r)
    c.setString(s)
    c.CharHeight = sz
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    c.IsTextWrapped = True
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
