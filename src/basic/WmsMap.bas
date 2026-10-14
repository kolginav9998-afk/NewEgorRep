' WmsMap — «Карта склада»: the heat map over the topology of the warehouse as the storekeeper drew it (PRIME 0.8, D-095).
' «Топология» holds the layout as is (pasted, or loaded from its file by «Загрузить топологию…»): every cell whose text is
' «ряд-секция-ярус» (5-12-3) is a place. The map is that layout copied cell for cell (sizes, colours, merged cells), each
' place coloured by the layer: Занятость, Выдачи за 30 дней, Без движения, Требует внимания, Свободные места. A click on a
' place shows what lies there; the EIs whose place is not on the map are listed. Places of «Наличие» are read as people
' write them: 5-12-3, 5.12.3, 5/12/3, «Р5 С12 Я3», 05-12-03. Read only: the data of the book is never changed.
Option Explicit

Public Const MP_TOP = 7
Public Const MP_INFO_ROW = 4
Public Const MP_SUM_ROW = 3
Public Const MP_LEGEND_ROW = 5
Public Const MP_MAX_PLACES = 20000

Global gMpLayer As String
Global gMpN As Long
Global gMpKey() As String
Global gMpRow() As Long
Global gMpCol() As Long
Global gMpEIs() As String
Global gMpCells As Object
Global gMpPanelCol As Long
Global gMpBusy As Boolean

Private Function TopoSheet() As Object
    TopoSheet = gDoc.Sheets.getByName(SH_TOPO)
End Function

' a place of the topology: the whole text is three numbers (5-12-3); a number alone (a label of a row) is not a place
Function IsPlaceText(s As String) As Boolean
    Dim k As String, a As Variant, i As Integer, nNum As Integer
    s = Trim(s)
    If Len(s) < 5 Or Len(s) > 20 Then Exit Function
    k = PlaceKey(s)
    If k = "" Then Exit Function
    a = Split(k, "-")
    If UBound(a) <> 2 Then Exit Function
    For i = 0 To 2
        If IsNumeric(a(i)) Then nNum = nNum + 1
    Next i
    IsPlaceText = (nNum >= 2)
End Function

' ================================================================ buttons

Sub BtnMapOcc(Optional oEvent As Variant)
    MapShow("OCC")
End Sub

Sub BtnMapIssues(Optional oEvent As Variant)
    MapShow("ISS")
End Sub

Sub BtnMapStale(Optional oEvent As Variant)
    MapShow("STALE")
End Sub

Sub BtnMapAttention(Optional oEvent As Variant)
    MapShow("ATT")
End Sub

Sub BtnMapFree(Optional oEvent As Variant)
    MapShow("FREE")
End Sub

Sub BtnMapRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"), True)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Не на карте»: the EIs with a balance whose place is not a place of the map — on «Поиск» (their card by a double click)
Sub BtnMapOff(Optional oEvent As Variant)
    Dim rows() As Variant, i As Long, n As Long, s As String
    WmsInit()
    If gMpN = 0 Or gMpStatSeq <> SysNum(SK_LAST_SEQ) Then MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"))
    If gMpN = 0 Then
        WmsUi.UiMessage("Карта ещё не построена: нужна топология склада (лист «" & SH_TOPO & "»)")
        Exit Sub
    End If
    n = gMpNOff
    If n > FD_MAX Then n = FD_MAX
    If n > 0 Then
        ReDim rows(n - 1)
        For i = 0 To n - 1
            rows(i) = Array(SH_STOCK, gMpOff(i)(6) + 1, gMpOff(i)(0), gMpOff(i)(1), gMpOff(i)(2), gMpOff(i)(3), gMpOff(i)(4), gMpOff(i)(5), "", _
                SH_STOCK, gMpOff(i)(6), gMpOff(i)(0))
        Next i
    End If
    s = WmsFind.ShowRows("ЕИ с остатком не на карте: " & NumText(gMpNOff), IIf(gMpNOff > 0, "Место не совпадает ни с одним местом карты — " _
        & "уточните его: «" & SH_ADJUST & "» → «Перемещение» (место «ряд-секция-ярус», например 5-12-3).", "Все ЕИ с остатком лежат на местах карты."), rows, n)
    WmsView.GoCell(SH_FIND, FD_TABLE, 1)
End Sub

' «Уместить на экран»: the zoom that shows the whole map (the header shrinks with it; Ctrl+колесо мыши — как удобно)
Sub BtnMapFit(Optional oEvent As Variant)
    Dim mp As Object, ctl As Object
    WmsInit()
    On Error GoTo EH
    If gMpN = 0 Then MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"))
    If gMpN = 0 Then Exit Sub
    mp = gDoc.Sheets.getByName(SH_MAP)
    ctl = gDoc.getCurrentController()
    ctl.setActiveSheet(mp)
    ctl.select(mp.getCellRangeByPosition(0, MP_TOP, gMpCols - 1, MP_TOP + gMpRows - 1))
    ctl.ZoomType = com.sun.star.view.DocumentZoomType.OPTIMAL
    ctl.select(mp.getCellByPosition(0, MP_INFO_ROW))
EH:
End Sub

Private Sub MapShow(sLayer As String)
    Dim s As String
    WmsInit()
    s = MapBuild(sLayer)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Загрузить топологию…»: a sheet of another file becomes «Топология» (cells, sizes, colours, merged cells, as drawn)
Sub BtnTopoLoad(Optional oEvent As Variant)
    Dim fp As Object, s As String
    WmsInit()
    On Error GoTo EH
    fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    fp.initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE))
    fp.setTitle("Файл с топологией склада (лист, где места подписаны «ряд-секция-ярус»)")
    fp.appendFilter("Таблицы (*.ods; *.xlsx; *.xls)", "*.ods;*.xlsx;*.xls")
    If fp.execute() <> 1 Then Exit Sub
    s = TopoLoad(fp.getFiles()(0))
    If Left(s, 3) = "OK:" Then
        MapBuild("OCC")
        WmsUi.UiMessage("Топология загружена: " & Mid(s, 4) & Chr(10) & "Карта построена по ней. Сохраните книгу (Ctrl+S), чтобы топология осталась.")
    Else
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    End If
    Exit Sub
EH:
    WmsUi.UiMessage("Топология не загружена: " & Error$)
End Sub

' the sheet of url with the most places becomes «Топология». OK:<places, sheet> / ERR:<why>
Function TopoLoad(url As String) As String
    Dim args(2) As New com.sun.star.beans.PropertyValue, src As Object, i As Integer, best As Integer, bestN As Long, n As Long, nm As String
    Dim pos As Integer, sh As Object, ok As Boolean, d As Variant, r As Long, c As Long, cur As Object
    WmsInit()
    On Error GoTo EH
    args(0).Name = "Hidden"
    args(0).Value = True
    args(1).Name = "ReadOnly"
    args(1).Value = True
    args(2).Name = "MacroExecutionMode"
    args(2).Value = 0
    src = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    If IsNull(src) Then
        TopoLoad = "ERR:файл не открылся: " & ConvertFromURL(url)
        Exit Function
    End If
    best = -1
    For i = 0 To src.Sheets.getCount() - 1
        n = CountPlaces(src.Sheets.getByIndex(i))
        If n > bestN Then
            bestN = n
            best = i
        End If
    Next i
    If best < 0 Then
        src.close(True)
        TopoLoad = "ERR:в файле нет мест вида «ряд-секция-ярус» (например 5-12-3) — это не топология склада"
        Exit Function
    End If
    nm = src.Sheets.getByIndex(best).Name
    gDoc.unprotect(PROTECT_PWD)
    pos = SheetIndex(SH_TOPO)
    gDoc.Sheets.importSheet(src, nm, pos)
    gDoc.Sheets.removeByName(SH_TOPO)
    gDoc.Sheets.getByIndex(pos).Name = SH_TOPO
    src.close(True)
    ' formulas of the file become their values (a layout refers to nothing)
    sh = TopoSheet()
    If Not sh.isProtected() Then
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        d = sh.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, cur.getRangeAddress().EndRow)
        d.setDataArray(d.getDataArray())
    End If
    gDoc.protect(PROTECT_PWD)
    TopoLoad = "OK:мест " & bestN & " (лист «" & nm & "» файла " & ConvertFromURL(url) & ")"
    Exit Function
EH:
    On Error Resume Next
    gDoc.protect(PROTECT_PWD)
    TopoLoad = "ERR:топология не загружена: " & Error$ & " (строка " & Erl & ")"
End Function

Private Function SheetIndex(sName As String) As Integer
    Dim i As Integer
    For i = 0 To gDoc.Sheets.getCount() - 1
        If gDoc.Sheets.getByIndex(i).Name = sName Then
            SheetIndex = i
            Exit Function
        End If
    Next i
    SheetIndex = gDoc.Sheets.getCount()
End Function

Private Function CountPlaces(sh As Object) As Long
    Dim cur As Object, d As Variant, r As Long, c As Long, a As Object
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.getRangeAddress()
    If a.EndRow > 3000 Or a.EndColumn > 500 Then Exit Function
    d = sh.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).getDataArray()
    For r = 0 To UBound(d)
        For c = 0 To UBound(d(r))
            If VarType(d(r)(c)) = 8 Then
                If IsPlaceText(CStr(d(r)(c))) Then CountPlaces = CountPlaces + 1
            End If
        Next c
    Next r
End Function

' ================================================================ the map
' MapLayout copies «Топология» under the header (only when it changed), MapStats counts the warehouse per place (only when an
' operation was posted since), MapPaint writes the colour code of every place into a hidden grid right of the map: the
' colours themselves are conditional formats over the map (styles MAP_<layer>_0 … 4), so a layer is switched at once.

Public Const MP_GRID_COL = 600

Global gMpSig As String
Global gMpStatSeq As Double
Global gMpRows As Long
Global gMpCols As Long
Global gMpOcc() As Long
Global gMpIss() As Long
Global gMpStale() As Double
Global gMpAtt() As Integer
Global gMpOff() As Variant
Global gMpNOff As Long
Global gMpNoLvl As Long

' builds (or recolours) the map in the layer sLayer (OCC, ISS, STALE, ATT, FREE). OK:<counts> / ERR:<why>
Function MapBuild(sLayer As String, Optional bForce As Variant) As String
    Dim mp As Object, why As String, force As Boolean, nOcc As Long, nFree As Long, nAtt As Long, i As Long, lbl As String
    If gMpBusy Then Exit Function
    gMpBusy = True
    On Error GoTo EH
    WmsInit()
    If Not IsMissing(bForce) Then force = bForce
    mp = ViewOpen(SH_MAP)
    why = MapLayout(mp, force)
    If why <> "" Then
        MapFail(mp, why)
        gMpBusy = False
        MapBuild = "ERR:" & why
        Exit Function
    End If
    If force Or gMpStatSeq <> SysNum(SK_LAST_SEQ) Or UBound(gMpOcc) < gMpN Then MapStats()
    MapPaint(mp, sLayer)
    For i = 0 To gMpN - 1
        If gMpOcc(i) > 0 Then nOcc = nOcc + 1 Else nFree = nFree + 1
        If gMpAtt(i) > 0 Then nAtt = nAtt + 1
    Next i
    gMpLayer = sLayer
    Select Case sLayer
    Case "OCC": lbl = "Занятость: сколько ЕИ с остатком лежит на месте"
    Case "ISS": lbl = "Выдачи за 30 дней: сколько раз выдавали ЕИ, что лежат на месте"
    Case "STALE": lbl = "Без движения: сколько дней не двигался самый «старый» ЕИ места"
    Case "ATT": lbl = "Требует внимания: «Требует разбора» — жёлтый, остаток ≠ движениям — красный"
    Case Else: lbl = "Свободные места: зелёные — пустые, серые — занятые"
    End Select
    mp.getCellByPosition(0, MP_SUM_ROW).setString(lbl & "   ·   мест " & NumText(gMpN) & ", занято " & NumText(nOcc) & " (" _
        & CLng(nOcc / gMpN * 100) & "%), свободно " & NumText(nFree) & IIf(nAtt > 0, ", требуют внимания " & NumText(nAtt), "") _
        & "   ·   ЕИ не на карте: " & NumText(gMpNOff) & IIf(gMpNOff > 0, " — кнопка «Не на карте»", ""))
    mp.getCellByPosition(0, MP_INFO_ROW).setString("Щёлкните по месту — здесь появится, что на нём лежит.   Обновлено " & Format(Now(), "DD.MM.YYYY HH:MM") _
        & " за " & ViewMs() & " мс")
    Legend(mp, sLayer)
    ViewDone(mp)
    gMpBusy = False
    MapBuild = "OK:мест " & gMpN & ", занято " & nOcc & ", свободно " & nFree & ", внимание " & nAtt & ", не на карте " & gMpNOff & ", " & ViewMs() & " мс"
    Exit Function
EH:
    MapBuild = "ERR:карта не построена: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    ViewDone(mp)
    gMpBusy = False
End Function

' the layout of «Топология» under the header of the map (copied again only when it changed): "" or why there is no map
Private Function MapLayout(mp As Object, force As Boolean) As String
    Dim topo As Object, cur As Object, r0 As Long, c0 As Long, r1 As Long, c1 As Long, d As Variant, r As Long, c As Long, pk As String
    Dim sig As String, nPl As Long, oldLast As Long, dest As New com.sun.star.table.CellAddress, srcA As New com.sun.star.table.CellRangeAddress
    topo = TopoSheet()
    cur = topo.createCursor()
    cur.gotoStartOfUsedArea(False)
    r0 = cur.getRangeAddress().StartRow
    c0 = cur.getRangeAddress().StartColumn
    cur.gotoEndOfUsedArea(False)
    r1 = cur.getRangeAddress().EndRow
    c1 = cur.getRangeAddress().EndColumn
    If r1 - r0 > 2000 Or c1 - c0 > 400 Then
        MapLayout = "лист «" & SH_TOPO & "» слишком большой для карты (" & (r1 - r0 + 1) & " строк, " & (c1 - c0 + 1) & " колонок; можно до 2000 × 400)"
        Exit Function
    End If
    d = topo.getCellRangeByPosition(c0, r0, c1, r1).getDataArray()
    ' the signature of the layout: its frame and its texts (a change of the drawing alone is taken by «Обновить»)
    For r = 0 To UBound(d)
        For c = 0 To UBound(d(r))
            If VarType(d(r)(c)) = 8 Then
                If IsPlaceText(CStr(d(r)(c))) Then
                    nPl = nPl + 1
                    sig = sig & r & "," & c & "=" & d(r)(c) & ";"
                    If Len(sig) > 60000 Then sig = CStr(Len(sig)) & Right(sig, 200)
                End If
            End If
        Next c
    Next r
    If nPl = 0 Then
        gMpN = 0
        MapLayout = "в листе «" & SH_TOPO & "» нет мест вида «ряд-секция-ярус» (например 5-12-3). Вставьте туда свою схему склада (Ctrl+V в ячейку A1) " _
            & "или нажмите «Загрузить топологию…» и выберите её файл"
        Exit Function
    End If
    sig = r0 & ":" & c0 & ":" & r1 & ":" & c1 & ":" & nPl & ":" & sig
    If Not force And sig = gMpSig And gMpN > 0 Then Exit Function
    ' the previous map: its rows removed whole (contents, colours, merged cells, the hidden grid)
    oldLast = WmsOrders.LastRow(mp)
    If oldLast >= MP_TOP Then mp.getRows().removeByIndex(MP_TOP, oldLast - MP_TOP + 1)
    gMpN = 0
    gMpCells = New Collection
    ReDim gMpKey(nPl)
    ReDim gMpRow(nPl)
    ReDim gMpCol(nPl)
    For r = 0 To UBound(d)
        For c = 0 To UBound(d(r))
            If VarType(d(r)(c)) = 8 Then
                If IsPlaceText(CStr(d(r)(c))) Then
                    gMpKey(gMpN) = PlaceKey(CStr(d(r)(c)))
                    gMpRow(gMpN) = MP_TOP + r
                    gMpCol(gMpN) = c
                    CollPut(gMpCells, gMpN, (MP_TOP + r) & ":" & c)
                    gMpN = gMpN + 1
                End If
            End If
        Next c
    Next r
    dest.Sheet = SheetIndex(SH_MAP)
    dest.Column = 0
    dest.Row = MP_TOP
    srcA.Sheet = SheetIndex(SH_TOPO)
    srcA.StartColumn = c0
    srcA.StartRow = r0
    srcA.EndColumn = c1
    srcA.EndRow = r1
    mp.copyRange(dest, srcA)
    For c = c0 To c1
        mp.getColumns().getByIndex(c - c0).Width = topo.getColumns().getByIndex(c).Width
    Next c
    For r = r0 To r1
        mp.getRows().getByIndex(MP_TOP + r - r0).Height = topo.getRows().getByIndex(r).Height
    Next r
    gMpRows = r1 - r0 + 1
    gMpCols = c1 - c0 + 1
    For c = MP_GRID_COL To MP_GRID_COL + gMpCols - 1
        mp.getColumns().getByIndex(c).IsVisible = False
    Next c
    gMpSig = sig
    gMpStatSeq = -1
End Function

' the colours of the map in the layer: a place whose code in the hidden grid is k gets the style MAP_<layer>_k
Private Sub MapConditions(mp As Object, sLayer As String)
    Dim rng As Object, cf As Object, k As Integer, ref As String, p(3) As New com.sun.star.beans.PropertyValue, srcPos As New com.sun.star.table.CellAddress
    EnsureMapStyles()
    rng = mp.getCellRangeByPosition(0, MP_TOP, gMpCols - 1, MP_TOP + gMpRows - 1)
    cf = rng.ConditionalFormat
    cf.clear()
    ref = MapColName(MP_GRID_COL) & (MP_TOP + 1)
    srcPos.Sheet = SheetIndex(SH_MAP)
    srcPos.Column = 0
    srcPos.Row = MP_TOP
    For k = 0 To 4
        p(0).Name = "Operator"
        p(0).Value = com.sun.star.sheet.ConditionOperator.FORMULA
        p(1).Name = "Formula1"
        p(1).Value = "AND(ISNUMBER(" & ref & ");" & ref & "=" & k & ")"
        p(2).Name = "StyleName"
        p(2).Value = "MAP_" & sLayer & "_" & k
        p(3).Name = "SourcePosition"
        p(3).Value = srcPos
        cf.addNew(p())
    Next k
    rng.ConditionalFormat = cf
End Sub

Private Function MapColName(ByVal c As Long) As String
    Dim s As String, m As Long
    c = c + 1
    Do While c > 0
        m = (c - 1) Mod 26
        s = Chr(65 + m) & s
        c = (c - 1 - m) \ 26
    Loop
    MapColName = s
End Function

' the 25 styles of the layers (made by the book build; made here once when missing — the structure of the book is unlocked
' for that moment only)
Private Sub EnsureMapStyles()
    Dim fam As Object, k As Integer, st As Object, layers As Variant, j As Integer, cols As Variant, ink As Variant, nm As String, opened As Boolean
    fam = gDoc.StyleFamilies.getByName("CellStyles")
    layers = Array("OCC", "ISS", "STALE", "ATT", "FREE")
    For j = 0 To 4
        LayerColors(layers(j), cols, ink)
        For k = 0 To 4
            nm = "MAP_" & layers(j) & "_" & k
            If Not fam.hasByName(nm) Then
                If Not opened Then
                    gDoc.unprotect(PROTECT_PWD)
                    opened = True
                End If
                st = gDoc.createInstance("com.sun.star.style.CellStyle")
                fam.insertByName(nm, st)
                st = fam.getByName(nm)
                st.CellBackColor = cols(k)
                st.CharColor = ink(k)
            End If
        Next k
    Next j
    If opened Then gDoc.protect(PROTECT_PWD)
End Sub

' the warehouse per place: EIs with a balance, issues in 30 days, the oldest movement, the checks; the EIs not on the map
Private Sub MapStats()
    Dim nextEI As Long, today As Double, lastMv() As Double, iss30() As Long, ex() As Double, mv() As Double, src() As Integer, places As Object
    Dim r As Long, r1x As Long, i As Long, n As Long, st As Variant, pk As String, idx As Long
    nextEI = CLng(SysNum(SK_NEXT_EI))
    today = Today0()
    places = New Collection
    For i = 0 To gMpN - 1
        CollPut(places, i, gMpKey(i))
    Next i
    ReDim gMpOcc(gMpN)
    ReDim gMpIss(gMpN)
    ReDim gMpStale(gMpN)
    ReDim gMpAtt(gMpN)
    ReDim gMpEIs(gMpN)
    gMpNOff = 0
    gMpNoLvl = 0
    ReDim gMpOff(IIf(nextEI > 1, nextEI - 1, 1))
    CollectMoves(nextEI, today, lastMv, iss30)
    WmsLegacy.ExpectedBalances(ex, mv, src, nextEI)
    For r = 1 To nextEI - 1 Step CHECK_CHUNK_ROWS
        r1x = r + CHECK_CHUNK_ROWS - 1
        If r1x > nextEI - 1 Then r1x = nextEI - 1
        st = WmsIssue.StockSheet().getCellRangeByPosition(0, r, SC_STATE, r1x).getDataArray()
        For i = 0 To r1x - r
            n = r + i
            If CStr(st(i)(SC_EI)) <> "" And CStr(st(i)(SC_STATE)) <> EI_ST_STORNO And VarType(st(i)(SC_QTY)) = 5 Then
                If st(i)(SC_QTY) > 0.0000005 Then
                    pk = PlaceKey(CStr(st(i)(SC_PLACE)))
                    idx = -1
                    If pk <> "" Then idx = CollGet(places, pk)
                    If idx >= 0 Then
                        gMpOcc(idx) = gMpOcc(idx) + 1
                        gMpIss(idx) = gMpIss(idx) + iss30(n)
                        If lastMv(n) > 0 And today - lastMv(n) > gMpStale(idx) Then gMpStale(idx) = today - lastMv(n)
                        If CStr(st(i)(SC_STATE)) = EI_ST_REVIEW And gMpAtt(idx) < 1 Then gMpAtt(idx) = 1
                        If src(n) <> 0 And Abs(WmsIssue.Round3(st(i)(SC_QTY) - ex(n))) > 0.0000005 Then gMpAtt(idx) = 2
                        If Len(gMpEIs(idx)) < 4000 Then gMpEIs(idx) = gMpEIs(idx) & n & ","
                    Else
                        gMpOff(gMpNOff) = Array(CStr(st(i)(SC_EI)), CStr(st(i)(SC_NAME)), st(i)(SC_QTY), CStr(st(i)(SC_UNIT)), CStr(st(i)(SC_PLACE)), _
                            IIf(Trim(CStr(st(i)(SC_PLACE))) = "", "место не указано", IIf(pk <> "" And UBound(Split(pk, "-")) = 1, _
                            "место без яруса — уточните «ряд-секция-ярус»", "такого места нет на карте")), n)
                        If pk <> "" And UBound(Split(pk, "-")) = 1 Then gMpNoLvl = gMpNoLvl + 1
                        gMpNOff = gMpNOff + 1
                    End If
                End If
            End If
        Next i
    Next r
    gMpStatSeq = SysNum(SK_LAST_SEQ)
End Sub

' the colour codes of the layer into the hidden grid, the colours of the styles, the list of the EIs not on the map
Private Sub MapPaint(mp As Object, sLayer As String)
    Dim g() As Variant, row() As Variant, r As Long, c As Long, i As Long
    ReDim g(gMpRows - 1)
    For r = 0 To gMpRows - 1
        ReDim row(gMpCols - 1)
        For c = 0 To gMpCols - 1
            row(c) = ""
        Next c
        g(r) = row
    Next r
    For i = 0 To gMpN - 1
        g(gMpRow(i) - MP_TOP)(gMpCol(i)) = Bucket(sLayer, i)
    Next i
    mp.getCellRangeByPosition(MP_GRID_COL, MP_TOP, MP_GRID_COL + gMpCols - 1, MP_TOP + gMpRows - 1).setDataArray(g)
    MapConditions(mp, sLayer)
End Sub

' the colour code 0 … 4 of place i in the layer
Private Function Bucket(sLayer As String, i As Long) As Integer
    Dim v As Double
    Select Case sLayer
    Case "OCC"
        v = gMpOcc(i)
        Bucket = IIf(v = 0, 0, IIf(v = 1, 1, IIf(v = 2, 2, IIf(v <= 4, 3, 4))))
    Case "ISS"
        v = gMpIss(i)
        If v = 0 Then
            Bucket = 0
        Else
            Bucket = IIf(v <= 2, 1, IIf(v <= 5, 2, IIf(v <= 10, 3, 4)))
        End If
    Case "STALE"
        v = gMpStale(i)
        If gMpOcc(i) = 0 Then
            Bucket = 0
        Else
            Bucket = IIf(v < 90, 1, IIf(v < 180, 2, IIf(v < 365, 3, 4)))
        End If
    Case "ATT"
        Bucket = IIf(gMpOcc(i) = 0, 0, IIf(gMpAtt(i) = 2, 3, IIf(gMpAtt(i) = 1, 2, 1)))
    Case Else
        Bucket = IIf(gMpOcc(i) = 0, 1, 0)
    End Select
End Function

' the five colours of a layer (backgrounds and texts): the same in the map and in its legend
Private Sub LayerColors(sLayer As String, ByRef cols As Variant, ByRef ink As Variant)
    Dim dk As Long, wh As Long, gray As Long
    dk = RGB(38, 50, 56)
    wh = RGB(255, 255, 255)
    gray = RGB(232, 236, 240)
    Select Case sLayer
    Case "OCC"
        cols = Array(RGB(244, 246, 248), RGB(198, 232, 201), RGB(143, 209, 158), RGB(76, 175, 106), RGB(46, 125, 79))
        ink = Array(dk, dk, dk, wh, wh)
    Case "ISS"
        cols = Array(gray, RGB(255, 232, 163), RGB(255, 192, 97), RGB(242, 140, 56), RGB(217, 71, 43))
        ink = Array(dk, dk, dk, wh, wh)
    Case "STALE"
        cols = Array(RGB(244, 246, 248), RGB(220, 235, 250), RGB(156, 195, 235), RGB(91, 151, 214), RGB(47, 109, 181))
        ink = Array(dk, dk, dk, wh, wh)
    Case "ATT"
        cols = Array(RGB(244, 246, 248), gray, RGB(255, 192, 97), RGB(217, 71, 43), gray)
        ink = Array(dk, dk, dk, wh, dk)
    Case Else
        cols = Array(gray, RGB(143, 209, 158), gray, gray, gray)
        ink = Array(dk, dk, dk, dk, dk)
    End Select
End Sub

Private Sub MapFail(mp As Object, why As String)
    mp.getCellByPosition(0, MP_SUM_ROW).setString(why)
    mp.getCellByPosition(0, MP_INFO_ROW).setString("")
    mp.getCellByPosition(0, MP_LEGEND_ROW).setString("")
    ViewDone(mp)
End Sub

' the legend of the layer in one line: coloured squares and their meaning
Private Sub Legend(mp As Object, sLayer As String)
    Dim cols As Variant, ink As Variant, txt As Variant, i As Integer, cell As Object, tx As Object, cur As Object
    LayerColors(sLayer, cols, ink)
    Select Case sLayer
    Case "OCC": txt = Array("пусто", "1 ЕИ", "2 ЕИ", "3–4 ЕИ", "5 и больше")
    Case "ISS": txt = Array("не выдавали", "1–2 раза", "3–5 раз", "6–10 раз", "больше 10")
    Case "STALE": txt = Array("пусто", "меньше 90 дней", "90–179 дней", "180–364 дня", "год и больше")
    Case "ATT": txt = Array("пусто", "в порядке", "требует разбора", "остаток ≠ движениям", "")
    Case Else: txt = Array("занято", "свободно", "", "", "")
    End Select
    cell = mp.getCellByPosition(0, MP_LEGEND_ROW)
    cell.setString("")
    tx = cell.getText()
    cur = tx.createTextCursor()
    For i = 0 To 4
        If txt(i) <> "" Then
            cur.gotoEnd(False)
            cur.CharColor = IIf(cols(i) = RGB(244, 246, 248), RGB(205, 212, 219), cols(i))
            tx.insertString(cur, "■■ ", False)
            cur.gotoEnd(False)
            cur.CharColor = RGB(55, 71, 79)
            tx.insertString(cur, txt(i) & "        ", False)
        End If
    Next i
End Sub

' the last movement of every EI and its issues in the last 30 days (posted rows of the sheets of movements)
Sub CollectMoves(nextEI As Long, today As Double, ByRef lastMv() As Double, ByRef iss30() As Long)
    Dim specs As Variant, i As Integer, sh As Object, last As Long, r0 As Long, r1 As Long, j As Long, d As Variant, n As Long, dt As Double
    Dim cEI As Integer, cDate As Integer, cCtl As Integer, cMax As Integer
    ReDim lastMv(nextEI)
    ReDim iss30(nextEI)
    ' sheet | EI | date | control (posted) — the receipts of «Заказы» have no control column: their EI is posted
    specs = Array(Array(SH_ISSUES, IC_EI, IC_DATE, IC_CTL), Array(SH_RETURNS, RC_EI, RC_DATE, RC_CTL), Array(SH_SPECIAL, XC_EI, XC_DATE, XC_STATUS), _
        Array(SH_ADJUST, AC_EI, AC_DATE, AC_CTL), Array(SH_ORDERS, OC_EI, OC_RDATE, -1))
    For i = 0 To UBound(specs)
        If gDoc.Sheets.hasByName(specs(i)(0)) Then
            sh = gDoc.Sheets.getByName(specs(i)(0))
            cEI = specs(i)(1)
            cDate = specs(i)(2)
            cCtl = specs(i)(3)
            cMax = cEI
            If cDate > cMax Then cMax = cDate
            If cCtl > cMax Then cMax = cCtl
            last = WmsOrders.LastRow(sh)
            For r0 = 1 To last Step CHECK_CHUNK_ROWS
                r1 = r0 + CHECK_CHUNK_ROWS - 1
                If r1 > last Then r1 = last
                d = sh.getCellRangeByPosition(0, r0, cMax, r1).getDataArray()
                For j = 0 To r1 - r0
                    n = EiNo(CStr(d(j)(cEI)), nextEI)
                    If n > 0 Then
                        If cCtl < 0 Or Left(CStr(d(j)(IIf(cCtl < 0, 0, cCtl))), Len(ST_POSTED)) = ST_POSTED Then
                            dt = DateOf(d(j)(cDate))
                            If dt > lastMv(n) Then lastMv(n) = dt
                            If i = 0 And dt > 0 And today - dt < 30 And today - dt >= 0 Then iss30(n) = iss30(n) + 1
                        End If
                    End If
                Next j
            Next r0
        End If
    Next i
End Sub

' ================================================================ a click on the map

' the selection changed: a place → what lies there (in the line over the map and in the panel)
Sub OnMapSelect(Optional oSel As Variant)
    Dim a As Object, idx As Long, mp As Object, s As String, eis As Variant, i As Integer, n As Long, st As Variant, rows() As Variant, k As Integer
    On Error GoTo EH
    If gMpBusy Then Exit Sub
    If IsMissing(oSel) Then Exit Sub
    If Not HasUnoInterfaces(oSel, "com.sun.star.sheet.XCellRangeAddressable") Then Exit Sub
    a = oSel.getRangeAddress()
    If a.StartRow < MP_TOP Then Exit Sub
    WmsInit()
    If gMpN = 0 Or IsNull(gMpCells) Then Exit Sub
    If UBound(gMpEIs) < gMpN Then Exit Sub
    idx = CollGet(gMpCells, a.StartRow & ":" & a.StartColumn)
    If idx < 0 Then Exit Sub
    mp = ViewOpen(SH_MAP)
    eis = Split(gMpEIs(idx), ",")
    k = UBound(eis)
    If k > 20 Then k = 20
    If k > 0 Then
        ReDim rows(k - 1)
    End If
    For i = 0 To k - 1
        n = CLng(eis(i))
        st = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_PLACE, n).getDataArray()(0)
        rows(i) = Array(CStr(st(SC_EI)), CStr(st(SC_NAME)), st(SC_QTY), CStr(st(SC_UNIT)), "")
        s = s & IIf(s <> "", "; ", "") & st(SC_EI) & " " & st(SC_NAME) & " — " & WmsIssue.QtyText(CDbl(st(SC_QTY))) & " " & st(SC_UNIT)
    Next i
    If UBound(eis) <= 0 Then s = "пусто — место свободно"
    mp.getCellByPosition(0, MP_INFO_ROW).setString("Место " & gMpKey(idx) & ": " & s & IIf(UBound(eis) > 20, " … ещё " & (UBound(eis) - 20), ""))
    ViewDone(mp)
    gUiLastMsg = "OK:" & gMpKey(idx) & ": " & s
    Exit Sub
EH:
    gUiLastMsg = "ERR:" & Error$ & " (строка " & Erl & ")"
End Sub

' the map is shown: built when it was never built in this session (the layer of the last time)
Sub OnMapFocus(Optional oEvent As Variant)
    On Error GoTo EH
    WmsInit()
    If gMpN = 0 And Not gMpBusy Then MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"))
EH:
End Sub

' ================================================================ «Куда положить?»
' For the row under the cursor of «Заказы» or «Иной приход»: the places where the same goods already lie (by the name, or by
' the article when there is one), then the free places of the map nearest to them (the same row and section first); with no
' such goods — the free places in the order of the map. The choice is written into «Место хранения» of a row not yet posted
' (an input cell, not an operation); a posted row is moved by «Корректировки» → «Перемещение».

Sub BtnWherePut(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = WherePut()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Function WherePut() As String
    Dim sName As String, r As Long, sh As Object, nm As String, art As String, cPlace As Integer, posted As Boolean, lines As Variant, opts() As String
    Dim k As Integer, c As Integer, n As Integer, pick As String, s As String
    sName = gDoc.getCurrentController().getActiveSheet().Name
    If sName <> SH_ORDERS And sName <> SH_SPECIAL Then
        WherePut = "ERR:«Куда положить?» работает на листах «" & SH_ORDERS & "» и «" & SH_SPECIAL & "»: встаньте на строку товара"
        Exit Function
    End If
    r = ViewRow(sName)
    If r < 1 Then
        WherePut = "ERR:встаньте на строку товара и нажмите «Куда положить?» снова"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(sName)
    If sName = SH_ORDERS Then
        nm = Trim(sh.getCellByPosition(OC_NAME, r).getString())
        art = Trim(sh.getCellByPosition(OC_ART, r).getString())
        cPlace = OC_PLACE
        posted = (sh.getCellByPosition(OC_EI, r).getString() <> "")
    Else
        nm = Trim(sh.getCellByPosition(XC_NAME, r).getString())
        art = Trim(sh.getCellByPosition(XC_ART, r).getString())
        cPlace = XC_PLACE
        posted = (sh.getCellByPosition(XC_EI, r).getString() <> "")
    End If
    If nm = "" And art = "" Then
        WherePut = "ERR:в строке " & (r + 1) & " нет наименования — впишите товар, затем «Куда положить?»"
        Exit Function
    End If
    s = WhereSuggest(nm, art)
    If Left(s, 4) = "ERR:" Then
        WherePut = s
        Exit Function
    End If
    lines = Split(s, Chr(10))
    n = UBound(lines) + 1
    ReDim opts(n - 1)
    For k = 0 To n - 1
        opts(k) = Mid(lines(k), InStr(lines(k), "|") + 1)
    Next k
    c = WmsPrimeUi.UiChoose("Куда положить — " & IIf(nm <> "", nm, art), "Места для «" & IIf(nm <> "", nm, art) & "»: сначала где этот товар уже лежит, " _
        & "затем свободные рядом." & IIf(posted, " Строка уже проведена — выбранное место будет подсказкой для перемещения.", ""), opts)
    If c < 0 Or c >= n Then
        gUiLastMsg = "SKIP:место не выбрано"
        WherePut = gUiLastMsg
        Exit Function
    End If
    pick = Left(lines(c), InStr(lines(c), "|") - 1)
    If posted Then
        gUiLastMsg = "OK:строка уже проведена: чтобы переложить ЕИ на " & pick & ", проведите «" & SH_ADJUST & "» → «Перемещение» (новое место " & pick & ")"
        WmsUi.UiMessage(Mid(gUiLastMsg, 4))
    Else
        sh.getCellByPosition(cPlace, r).setString(pick)
        gUiLastMsg = "OK:место " & pick & " записано в строку " & (r + 1)
    End If
    WherePut = gUiLastMsg
End Function

' the places for goods nm (article art): lines «place|text for a person», at most 8; ERR:… when there is nothing to propose
Function WhereSuggest(nm As String, art As String) As String
    Dim nextEI As Long, r As Long, r1x As Long, i As Long, st As Variant, same As Object, sKeys() As String, sCnt() As Long, nS As Long, idx As Long
    Dim pk As String, out As String, nOut As Integer, j As Long, best As Long, bestD As Double, d As Double, a As Variant, b As Variant, used() As Boolean
    Dim anchorKey As String, k As Integer
    nextEI = CLng(SysNum(SK_NEXT_EI))
    same = New Collection
    ReDim sKeys(50)
    ReDim sCnt(50)
    ' where the same goods lie now (EIs with a balance)
    For r = 1 To nextEI - 1 Step CHECK_CHUNK_ROWS
        r1x = r + CHECK_CHUNK_ROWS - 1
        If r1x > nextEI - 1 Then r1x = nextEI - 1
        st = WmsIssue.StockSheet().getCellRangeByPosition(0, r, SC_STATE, r1x).getDataArray()
        For i = 0 To r1x - r
            If VarType(st(i)(SC_QTY)) = 5 Then
                If st(i)(SC_QTY) > 0.0000005 And CStr(st(i)(SC_STATE)) <> EI_ST_STORNO Then
                    If (art <> "" And LCase(Trim(CStr(st(i)(SC_ART)))) = LCase(art)) Or (art = "" And LCase(Trim(CStr(st(i)(SC_NAME)))) = LCase(nm)) Then
                        pk = Trim(CStr(st(i)(SC_PLACE)))
                        If pk <> "" Then
                            idx = CollGet(same, LCase(pk))
                            If idx < 0 And nS <= 50 Then
                                idx = nS
                                CollPut(same, idx, LCase(pk))
                                sKeys(idx) = pk
                                nS = nS + 1
                            End If
                            If idx >= 0 Then sCnt(idx) = sCnt(idx) + 1
                        End If
                    End If
                End If
            End If
        Next i
    Next r
    For i = 0 To nS - 1
        If nOut < 3 Then
            out = out & IIf(out <> "", Chr(10), "") & sKeys(i) & "|" & sKeys(i) & " — уже лежит этот товар (" & sCnt(i) & " " & Plural(sCnt(i), "ЕИ", "ЕИ", "ЕИ") & ")"
            nOut = nOut + 1
            If anchorKey = "" Then anchorKey = PlaceKey(sKeys(i))
        End If
    Next i
    ' the free places of the map, nearest to the goods (the same row and section first)
    If gMpN = 0 Or gMpStatSeq <> SysNum(SK_LAST_SEQ) Then MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"))
    If gMpN > 0 Then
        ReDim used(gMpN)
        For k = 1 To 8 - nOut
            best = -1
            bestD = 1E+30
            For j = 0 To gMpN - 1
                If gMpOcc(j) = 0 And Not used(j) Then
                    d = PlaceDistance(anchorKey, gMpKey(j), j)
                    If d < bestD Then
                        bestD = d
                        best = j
                    End If
                End If
            Next j
            If best < 0 Then Exit For
            used(best) = True
            out = out & IIf(out <> "", Chr(10), "") & gMpKey(best) & "|" & gMpKey(best) & " — свободно" & IIf(anchorKey <> "", IIf(bestD < 1000, ", рядом", ""), "")
            nOut = nOut + 1
        Next k
    End If
    If out = "" Then
        WhereSuggest = "ERR:подсказать нечего: этого товара на складе нет, а карта склада не задана (лист «" & SH_TOPO & "») — впишите место вручную"
    Else
        WhereSuggest = out
    End If
End Function

' the distance between two places «ряд-секция-ярус» (numbers): rows far apart count most; with no anchor — the order of the map
Private Function PlaceDistance(sFrom As String, sTo As String, j As Long) As Double
    Dim a As Variant, b As Variant
    If sFrom = "" Then
        PlaceDistance = j
        Exit Function
    End If
    a = Split(sFrom, "-")
    b = Split(sTo, "-")
    If UBound(a) < 2 Or UBound(b) < 2 Or Not IsNumeric(a(0)) Or Not IsNumeric(b(0)) Or Not IsNumeric(a(1)) Or Not IsNumeric(b(1)) Then
        PlaceDistance = 100000 + j
        Exit Function
    End If
    PlaceDistance = Abs(CDbl(a(0)) - CDbl(b(0))) * 1000 + Abs(CDbl(a(1)) - CDbl(b(1))) * 10
    If IsNumeric(a(2)) And IsNumeric(b(2)) Then PlaceDistance = PlaceDistance + Abs(CDbl(a(2)) - CDbl(b(2)))
End Function

' ================================================================ test seams

Function TestMapClick(r As Long, c As Long) As String
    Dim cell As Object
    WmsInit()
    cell = gDoc.Sheets.getByName(SH_MAP).getCellByPosition(c, r)
    gUiLastMsg = ""
    OnMapSelect(cell)
    TestMapClick = gUiLastMsg
End Function

Function TestMapPlace(sKey As String) As String
    Dim i As Long
    For i = 0 To gMpN - 1
        If gMpKey(i) = sKey Then
            TestMapPlace = gMpRow(i) & ":" & gMpCol(i)
            Exit Function
        End If
    Next i
End Function

Function TestMapEIs(sKey As String) As String
    Dim i As Long
    For i = 0 To gMpN - 1
        If gMpKey(i) = sKey Then
            TestMapEIs = i & "|" & gMpEIs(i)
            Exit Function
        End If
    Next i
End Function
