' WmsMap — «Карта склада»: the heat map over the topology of the warehouse as the storekeeper drew it (PRIME 0.8, D-095).
' «Топология» holds the layout as is (pasted, or loaded from its file by «Загрузить топологию…»): every cell whose text is
' «ряд-секция-ярус» (5-12-3) is a place. The map is that layout copied cell for cell (sizes, colours, merged cells), each
' place coloured by the layer: Занятость, Выдачи за 30 дней, Без движения, Требует внимания, Свободные места. A click on a
' place shows what lies there; the EIs whose place is not on the map are listed. Places of «Наличие» are read as people
' write them: 5-12-3, 5.12.3, 5/12/3, «Р5 С12 Я3», 05-12-03. Read only: the data of the book is never changed.
Option Explicit

Public Const MP_TOP = 7
Private Const MP_INFO_ROW = 4
Private Const MP_SUM_ROW = 3
Private Const MP_LEGEND_ROW = 5
Private Const MP_MAX_PLACES = 20000
' the hidden sorted table of the texts of the places (MapKeys): text | number of the place (1-based)
Private Const MP_KEY_COL = 1100
' 1.0 (М2): the directory of the places outside the racks of the map (a floor, the yard, an office): known places, not
' «не на карте»; A the place, B its kind, C the zone, D the capacity in cells, E «закрыто», F a note
Public Const SH_PLACES = "Места"
' 1.0: the layout «Топология» had before the last «Сохранить карту» (hidden)
Public Const SH_TOPO_PREV = "_TOPO_PREV"
Private Const PD_PLACE = 0
Private Const PD_KIND = 1
Private Const PD_ZONE = 2
Private Const PD_CAP = 3
Private Const PD_CLOSED = 4
Private Const PD_NOTE = 5

Global gMpLayer As String
Global gMpN As Long
Global gMpKey() As String
Global gMpRow() As Long
Global gMpCol() As Long
Global gMpEIs() As String
Global gMpCells As Object
Global gMpPanelCol As Long
Global gMpBusy As Boolean
' 1.0: the map is being edited on the map itself (BtnMapEdit … BtnMapSave / BtnMapCancel); a cell of the map keeps the mode
' over a save and a reopening of the book
Global gMpEdit As Boolean
Private Const MP_EDIT_COL = 1105
' 1.0 (С5): the stamp of the map saved with the book — «LAST_SEQ|слой|когда»: a map up to date is not counted again at the
' next start, only its places are read
Private Const MP_STAMP_COL = 1109
Global gMpLazyNow As Boolean

Private Function TopoSheet() As Object
    TopoSheet = gDoc.Sheets.getByName(SH_TOPO)
End Function

' a place of the topology: the whole text is three numbers (5-12-3); a number alone (a label of a row) is not a place
Function IsPlaceText(s As String) As Boolean
    IsPlaceText = (PlaceTextKey(s) <> "")
End Function

' the key of a place of the topology (5-12-3), "" when the text is not a place
Function PlaceTextKey(ByVal s As String) As String
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
    If nNum >= 2 Then PlaceTextKey = k
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

Sub BtnMapFill(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = MapBuild("FILL")
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' 1.0 (М5): the places of the EIs eis (numbers, separated by «,») painted on the map — the layer «Товар»; the map is shown
Function ShowOnMap(sEIs As String, sName As String) As String
    Dim a As Variant, i As Long, n As Long, pls As Variant, j As Long, idx As Long, keyIdx As Object, st As Variant, nOn As Long, nOff As Long
    WmsInit()
    If MapEditing() Then
        ShowOnMap = "ERR:карта в режиме правки — сначала «Сохранить карту» или «Отменить правку»"
        Exit Function
    End If
    If gMpN = 0 Or gMpStatSeq <> SysNum(SK_LAST_SEQ) Then MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"))
    If gMpN = 0 Then
        ShowOnMap = "ERR:карта склада не построена — нужна топология (лист «" & SH_TOPO & "»)"
        Exit Function
    End If
    ReDim gMpFind(gMpN)
    keyIdx = New Collection
    For j = 0 To gMpN - 1
        If gMpFirst(j) = j Then CollPut(keyIdx, j, gMpKey(j))
    Next j
    a = Split(sEIs, ",")
    For i = 0 To UBound(a)
        n = CLng(Val(a(i)))
        If n > 0 Then
            st = WmsIssue.StockSheet().getCellByPosition(SC_PLACE, n).getString()
            pls = PlaceList(CStr(st))
            For j = 0 To UBound(pls)
                idx = CollGet(keyIdx, CStr(pls(j)))
                If idx >= 0 Then
                    gMpFind(idx) = True
                    nOn = nOn + 1
                Else
                    nOff = nOff + 1
                End If
            Next j
        End If
    Next i
    gMpFindName = sName
    ShowOnMap = MapBuild("FIND")
    gDoc.getCurrentController().setActiveSheet(gDoc.Sheets.getByName(SH_MAP))
    If Left(ShowOnMap, 3) = "OK:" Then ShowOnMap = "OK:мест на карте " & nOn & IIf(nOff > 0, ", вне карты " & nOff, "")
    gUiLastMsg = ShowOnMap
End Function

Sub BtnMapRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"), True)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Не на карте»: the EIs with a balance whose place is not a place of the map — on «Поиск» (their card by a double click)
Sub BtnMapOff(Optional oEvent As Variant)
    Dim rows() As Variant, i As Long, n As Long, s As String, k As Long, st As Variant, why As String
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
            k = gMpOff(i)(0)
            st = WmsIssue.StockSheet().getCellRangeByPosition(0, k, SC_PLACE, k).getDataArray()(0)
            Select Case gMpOff(i)(1)
            Case 0: why = "место не указано"
            Case -3: why = "место без яруса — уточните «ряд-секция-ярус»"
            Case Else: why = "такого места нет на карте и в «" & SH_PLACES & "» (место вне стеллажей запишите туда)"
            End Select
            rows(i) = Array(SH_STOCK, k + 1, CStr(st(SC_EI)), CStr(st(SC_NAME)), st(SC_QTY), CStr(st(SC_UNIT)), CStr(st(SC_PLACE)), why, "", SH_STOCK, k, CStr(st(SC_EI)))
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
    If MapEditing() Then
        WmsUi.UiMessage("Карта в режиме правки — сначала «Сохранить карту» или «Отменить правку».")
        Exit Sub
    End If
    fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    fp.initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE))
    fp.setTitle("Файл с топологией склада (лист, где места подписаны «ряд-секция-ярус»)")
    fp.appendFilter("Таблицы (*.ods; *.xlsx; *.xls)", "*.ods;*.xlsx;*.xls")
    If fp.execute() <> 1 Then Exit Sub
    ' 1.0 (Н6): a restore point — the previous topology stays in a copy
    WmsSafe.RestorePoint("перед загрузкой топологии")
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
    gTkValid = False
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

Private Const MP_GRID_COL = 600

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
Global gMpKeyN As Long
Global gMpPlaceOf As Variant
' 1.0: the first place of the map with the key of place i (one address drawn in several cells shows the same everywhere)
Global gMpFirst() As Long
' 1.0 (М2): the EIs on the places of «Места» (outside the racks of the map)
Global gMpNDir As Long
' 1.0: the keys of «Топология» and of «Места» read once and kept until those sheets change (their events): the check of a
' place typed at a posting needs no map. gTk: key → 1; gPd: key → the row of «Места»
Global gTkValid As Boolean
Global gTk As Object
Global gTkN As Long
Global gTkList() As String
Global gPdValid As Boolean
Global gPd As Object
Global gPdN As Long
Global gPdClosed() As Boolean
Global gPdName() As String
' 1.0 (М4): the capacity of a place in cells («Места» D; 1 when not given) — a rack place listed in «Места» gets its own
Global gPdCap() As Long
' 1.0 (М5): the places of the layer «Товар» (ShowOnMap)
Global gMpFind() As Boolean
Global gMpFindName As String

' builds (or recolours) the map in the layer sLayer (OCC, ISS, STALE, ATT, FREE). OK:<counts> / ERR:<why>
Function MapBuild(sLayer As String, Optional bForce As Variant) As String
    Dim mp As Object, why As String, force As Boolean, nOcc As Long, nFree As Long, nAtt As Long, i As Long, lbl As String
    If gMpBusy Then Exit Function
    ' 1.0: never over the edits of the storekeeper — the map is built again by «Сохранить карту» or «Отменить правку»
    If MapEditing() Then
        MapBuild = "ERR:карта в режиме правки — сначала «Сохранить карту» или «Отменить правку»"
        Exit Function
    End If
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
    If force Or gMpStatSeq <> SysNum(SK_LAST_SEQ) Or UBound(gMpOcc) < gMpN Then
        MapStats()
        If gVwEngFail <> "" Then
            gMpStatSeq = -1
            MapFail(mp, "карта не посчитана: " & gVwEngFail)
            gMpBusy = False
            MapBuild = "ERR:карта не посчитана: " & gVwEngFail
            Exit Function
        End If
    End If
    MapPaint(mp, sLayer)
    For i = 0 To gMpN - 1
        If gMpOcc(i) > 0 Then nOcc = nOcc + 1 Else nFree = nFree + 1
        If gMpAtt(i) > 0 Then nAtt = nAtt + 1
    Next i
    gMpLayer = sLayer
    Select Case sLayer
    Case "OCC": lbl = "Занятость: сколько ЕИ с остатком лежит на месте"
    Case "FILL": lbl = "Заполненность: ЕИ на месте против его вместимости (лист «" & SH_PLACES & "», колонка «Вместимость, ячеек»; без неё — 1)"
    Case "FIND": lbl = "Где лежит «" & gMpFindName & "»"
    Case "ISS": lbl = "Выдачи за 30 дней: сколько раз выдавали ЕИ, что лежат на месте"
    Case "STALE": lbl = "Без движения: сколько ЕИ места не двигались 90 дней и больше"
    Case "ATT": lbl = "Требует внимания: места, где лежит ЕИ «Требует разбора»"
    Case Else: lbl = "Свободные места: зелёные — пустые, серые — занятые"
    End Select
    mp.getCellByPosition(0, MP_SUM_ROW).setString(lbl & "   ·   мест " & NumText(gMpN) & ", занято " & NumText(nOcc) & " (" _
        & CLng(nOcc / gMpN * 100) & "%), свободно " & NumText(nFree) & IIf(nAtt > 0, ", требуют внимания " & NumText(nAtt), "") _
        & IIf(gMpNDir > 0, "   ·   вне стеллажей («" & SH_PLACES & "»): " & NumText(gMpNDir) & " ЕИ", "") _
        & "   ·   ЕИ не на карте: " & NumText(gMpNOff) & IIf(gMpNOff > 0, " — кнопка «Не на карте»", ""))
    mp.getCellByPosition(0, MP_INFO_ROW).setString("Щёлкните по месту — здесь появится, что на нём лежит.   Обновлено " & Format(Now(), "DD.MM.YYYY HH:MM"))
    Legend(mp, sLayer)
    ' С5: the stamp of this map; С4: its time
    mp.getCellByPosition(MP_STAMP_COL, 0).setString(NumStr(SysNum(SK_LAST_SEQ)) & "|" & sLayer & "|" & Format(Now(), "DD.MM.YYYY HH:MM"))
    mp.getCellByPosition(MP_STAMP_COL, 1).setString(MapSigHash(gMpSig))
    WmsCare.TimingSet("MAP", ViewMs())
    ViewDone(mp)
    gMpBusy = False
    MapBuild = "OK:мест " & gMpN & ", занято " & nOcc & ", свободно " & nFree & ", внимание " & nAtt & ", не на карте " & gMpNOff & ", вне стеллажей " _
        & gMpNDir & ", " & ViewMs() & " мс"
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
    Dim nTx As Long, pkR() As Long, pkC() As Long, pkK() As String, i As Long, nDt As Long, lazy As Boolean
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
                nTx = nTx + 1
                sig = sig & r & "," & c & "=" & d(r)(c) & ";"
                If Len(sig) > 60000 Then sig = CStr(Len(sig)) & Right(sig, 200)
            End If
        Next c
    Next r
    sig = r0 & ":" & c0 & ":" & r1 & ":" & c1 & ":" & nTx & ":" & sig
    ' 1.0: a layout written by a macro fires no event of the sheet — the keys of the check of places are read again too
    If sig <> gMpSig Then gTkValid = False
    If Not force And sig = gMpSig And gMpN > 0 Then Exit Function
    ' 1.0 (С5): the map saved with the book is of this very layout (its stamp) — the places read, the sheet not touched
    lazy = gMpLazyNow And Not force And mp.getCellByPosition(MP_STAMP_COL, 1).getString() = MapSigHash(sig)
    ' the places: the key of each text read once
    ReDim pkR(nTx)
    ReDim pkC(nTx)
    ReDim pkK(nTx)
    nDt = 0
    For r = 0 To UBound(d)
        For c = 0 To UBound(d(r))
            pk = ""
            If VarType(d(r)(c)) = 8 Then
                pk = PlaceTextKey(CStr(d(r)(c)))
            ElseIf VarType(d(r)(c)) = 5 Then
                ' 1.0: «5.12.3» that Calc took for a date (05.12.2003) is the place 5-12-3
                If d(r)(c) > 1 And d(r)(c) = Int(d(r)(c)) Then
                    pk = DateCellPlace(topo.getCellByPosition(c0 + c, r0 + r))
                    If pk <> "" Then nDt = nDt + 1
                End If
            End If
            If pk <> "" Then
                If nPl > UBound(pkK) Then
                    ReDim Preserve pkR(2 * nPl)
                    ReDim Preserve pkC(2 * nPl)
                    ReDim Preserve pkK(2 * nPl)
                End If
                pkR(nPl) = r
                pkC(nPl) = c
                pkK(nPl) = pk
                nPl = nPl + 1
            End If
        Next c
    Next r
    If nPl = 0 Then
        gMpN = 0
        MapLayout = "в листе «" & SH_TOPO & "» нет мест вида «ряд-секция-ярус» (например 5-12-3). Вставьте туда свою схему склада (Ctrl+V в ячейку A1) " _
            & "или нажмите «Загрузить топологию…» и выберите её файл"
        Exit Function
    End If
    ' the previous map: its rows removed whole (contents, colours, merged cells, the hidden grid)
    If Not lazy Then
        oldLast = WmsOrders.LastRow(mp)
        If oldLast >= MP_TOP Then mp.getRows().removeByIndex(MP_TOP, oldLast - MP_TOP + 1)
    End If
    gMpN = 0
    gMpCells = New Collection
    ReDim gMpKey(nPl)
    ReDim gMpRow(nPl)
    ReDim gMpCol(nPl)
    For i = 0 To nPl - 1
        gMpKey(gMpN) = pkK(i)
        gMpRow(gMpN) = MP_TOP + pkR(i)
        gMpCol(gMpN) = pkC(i)
        CollPut(gMpCells, gMpN, (MP_TOP + pkR(i)) & ":" & pkC(i))
        gMpN = gMpN + 1
    Next i
    If lazy Then
        gMpRows = r1 - r0 + 1
        gMpCols = c1 - c0 + 1
        MapKeys(mp, d, False)
        gMpSig = sig
        gMpStatSeq = -1
        Exit Function
    End If
    dest.Sheet = SheetIndex(SH_MAP)
    dest.Column = 0
    dest.Row = MP_TOP
    srcA.Sheet = SheetIndex(SH_TOPO)
    srcA.StartColumn = c0
    srcA.StartRow = r0
    srcA.EndColumn = c1
    srcA.EndRow = r1
    mp.copyRange(dest, srcA)
    ' the places Calc took for dates are shown on the map as places («5-12-3», not «05.12.03»)
    If nDt > 0 Then
        For i = 0 To nPl - 1
            If VarType(d(pkR(i))(pkC(i))) = 5 Then mp.getCellByPosition(pkC(i), MP_TOP + pkR(i)).setString(pkK(i))
        Next i
    End If
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
    MapKeys(mp, d, True)
    gMpSig = sig
    gMpStatSeq = -1
End Function

' С5: the stamp of a layout (its signature of texts and frame) kept beside the map
Private Function MapSigHash(sig As String) As String
    MapSigHash = WmsJournal.LineHash(sig)
    If MapSigHash = "" Then MapSigHash = "L" & Len(sig) & ":" & Right(sig, 200)
End Function

' The sorted table of the texts of the places in hidden columns of the map: the key of each place (5-12-3) and its text in
' «Топология» when it is written otherwise (05-12-03), with the number of the place (1-based; a key met twice — the first
' place). MapStats finds the place of every EI in it by a binary search of the Calc engine (LOOKUP); the table is sorted by
' Calc itself, so its order is the one LOOKUP expects.
Private Sub MapKeys(mp As Object, d As Variant, bWrite As Boolean)
    Dim first As Object, seen As Object, keys() As Variant, n As Long, i As Long, raw As String, j As Long, um As Object, rng As Object
    Dim desc As Variant, k As Integer, sortF(0) As New com.sun.star.table.TableSortField
    first = New Collection
    seen = New Collection
    ReDim keys(2 * gMpN)
    ReDim gMpFirst(gMpN)
    For i = 0 To gMpN - 1
        j = CollGet(first, gMpKey(i))
        If j < 0 Then
            j = i
            CollPut(first, i, gMpKey(i))
        End If
        gMpFirst(i) = j
        If CollPut(seen, 1, gMpKey(i)) Then
            keys(n) = Array(gMpKey(i), j + 1)
            n = n + 1
        End If
        raw = Trim(CStr(d(gMpRow(i) - MP_TOP)(gMpCol(i))))
        If CollPut(seen, 1, raw) Then
            keys(n) = Array(raw, j + 1)
            n = n + 1
        End If
    Next i
    gMpKeyN = 0
    If n = 0 Then Exit Sub
    ' С5: the sorted table is on the sheet already (a map saved with the book)
    If Not bWrite Then
        gMpKeyN = n
        Exit Sub
    End If
    ReDim Preserve keys(n - 1)
    um = gDoc.getUndoManager()
    um.lock()
    mp.getCellRangeByPosition(MP_KEY_COL, 0, MP_KEY_COL + 1, MP_TOP).clearContents(1 + 2 + 4)
    rng = mp.getCellRangeByPosition(MP_KEY_COL, 0, MP_KEY_COL + 1, n - 1)
    rng.setDataArray(keys)
    sortF(0).Field = 0
    sortF(0).IsAscending = True
    sortF(0).IsCaseSensitive = False
    desc = rng.createSortDescriptor()
    For k = 0 To UBound(desc)
        If desc(k).Name = "SortFields" Then desc(k).Value = sortF
        If desc(k).Name = "ContainsHeader" Then desc(k).Value = False
    Next k
    rng.sort(desc)
    um.unlock()
    mp.getColumns().getByIndex(MP_KEY_COL).IsVisible = False
    mp.getColumns().getByIndex(MP_KEY_COL + 1).IsVisible = False
    gMpKeyN = n
End Sub

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
    ' the build makes the styles; a missing one (a book of another release) is made when the book allows it — never an error
    On Error Resume Next
    fam = gDoc.StyleFamilies.getByName("CellStyles")
    layers = Array("OCC", "ISS", "STALE", "ATT", "FREE", "FILL", "FIND")
    For j = 0 To 6
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

' The warehouse per place: EIs with a balance, issues in 30 days, EIs without movement 90+ days, «Требует разбора»; the EIs
' not on the map. The Calc engine finds the place of every EI (a binary search in the table of MapKeys) and the rows that
' matter; Basic reads only the EIs whose place is written otherwise (PlaceKey) and the issues of 30 days.
Private Sub MapStats()
    Dim nextEI As Long, today As Double, last As Long, rgQ As String, rgF As String, cond As String, keyR As String, numR As String, lk As String
    Dim pl As Variant, nPl As Long, a As Variant, nItems As Long, d As Variant, i As Long, n As Long, k As Long, v As String, pk As String, idx As Long
    Dim places As Object, cache As Object, recent() As Boolean, occ() As Long, iss() As Long, stale() As Double, att() As Integer, eis() As String
    Dim nOff As Long, off() As Variant, offRows() As Long, noLvl As Long, cE As Integer, cval() As String, nC As Long, ci As Long, code As String, j As Long
    Dim b As Variant, nDir As Long
    nextEI = CLng(SysNum(SK_NEXT_EI))
    today = Today0()
    ReDim occ(gMpN)
    ReDim iss(gMpN)
    ReDim stale(gMpN)
    ReDim att(gMpN)
    ReDim eis(gMpN)
    ReDim off(IIf(nextEI > 1, nextEI - 1, 1))
    gVwEngFail = ""
    last = nextEI - 1
    If last >= 1 And gMpKeyN > 0 Then
        rgQ = ColRng(SH_STOCK, SC_QTY, 1, last)
        rgF = ColRng(SH_STOCK, SC_PLACE, 1, last)
        cond = "ISNUMBER(" & rgQ & ")*(" & rgQ & ">0.0000005)*(" & ColRng(SH_STOCK, SC_STATE, 1, last) & "<>""" & EI_ST_STORNO & """)"
        keyR = mp_KeyRange(0)
        numR = mp_KeyRange(1)
        lk = "IFERROR(LOOKUP(" & rgF & ";" & keyR & ";" & keyR & ")=" & rgF & ";0)"
        ' every EI (row = EI): −1 — no balance; the number of its place when its place is written as on the map; 0 — a balance,
        ' the place written otherwise or not on the map (read below)
        pl = EngList("TEXTJOIN("";"";0;" & cond & "*(" & lk & "*IFERROR(LOOKUP(" & rgF & ";" & keyR & ";" & numR & ");0)+1)-1)", nPl)
        If gVwEngFail = "" And nPl = last Then
            ReDim recent(nextEI)
            MarkRecent(recent, nextEI, today - 89)
            ReDim offRows(last)
            ' occupancy, EIs without movement 90+ days, the EIs of each place
            For i = 0 To last - 1
                v = pl(i)
                If v = "0" Then
                    offRows(nOff) = i + 1
                    nOff = nOff + 1
                ElseIf v <> "-1" Then
                    k = CLng(v) - 1
                    occ(k) = occ(k) + 1
                    If Not recent(i + 1) Then stale(k) = stale(k) + 1
                    If Len(eis(k)) < 4000 Then eis(k) = eis(k) & (i + 1) & ","
                End If
            Next i
            ' the EIs whose place is written otherwise: read as people write places (PlaceKey); the others are not on the map
            places = New Collection
            For i = 0 To gMpN - 1
                CollPut(places, i, gMpKey(i))
            Next i
            cache = New Collection
            ReDim cval(nOff)
            nC = 0
            d = RowsCol(WmsIssue.StockSheet(), offRows, nOff, SC_PLACE)
            n = nOff
            nOff = 0
            For i = 0 To n - 1
                v = Trim(CStr(d(i)))
                ci = CollGet(cache, "~" & v)
                If ci = -1 Then
                    ci = nC
                    CollPut(cache, ci, "~" & v)
                    cval(ci) = PlacesOn(v, places)
                    nC = nC + 1
                End If
                code = cval(ci)
                If Left(code, 1) <> "-" And code <> "0" Then
                    ' on the map: one place or several («1-2-3, 1-2-4» — the goods take more than one cell)
                    k = offRows(i)
                    pl(k - 1) = code
                    a = Split(code, ",")
                    For j = 0 To UBound(a)
                        idx = CLng(a(j)) - 1
                        occ(idx) = occ(idx) + 1
                        If Not recent(k) Then stale(idx) = stale(idx) + 1
                        If Len(eis(idx)) < 4000 Then eis(idx) = eis(idx) & k & ","
                    Next j
                ElseIf code = "-4" Then
                    nDir = nDir + 1
                Else
                    ' the EI and why (−2: no such place, −3: a place without its level, 0 — no place)
                    off(nOff) = Array(offRows(i), CLng(code))
                    If code = "-3" Then noLvl = noLvl + 1
                    nOff = nOff + 1
                End If
            Next i
            ' the issues of 30 days, each on the place where its EI lies now
            n = WmsOrders.LastRow(WmsIssue.IssuesSheet())
            If n >= 1 Then
                a = EngList("TEXTJOIN("";"";1;IF(" & PostedExpr(ColRng(SH_ISSUES, IC_CTL, 1, n)) & "*" & DaysExpr(ColRng(SH_ISSUES, IC_DATE, 1, n), today - 29, today) _
                    & ";ROW(" & ColRng(SH_ISSUES, IC_EI, 1, n) & ")-1;""""))", nItems)
                d = RowsCol(WmsIssue.IssuesSheet(), a, nItems, IC_EI)
                For i = 0 To nItems - 1
                    n = EiFast(CStr(d(i)), nextEI)
                    If n > 0 Then
                        v = pl(n - 1)
                        If v <> "0" And v <> "-1" Then
                            If InStr(v, ",") = 0 Then
                                k = CLng(v) - 1
                                iss(k) = iss(k) + 1
                            Else
                                a = Split(v, ",")
                                For j = 0 To UBound(a)
                                    iss(CLng(a(j)) - 1) = iss(CLng(a(j)) - 1) + 1
                                Next j
                            End If
                        End If
                    End If
                Next i
            End If
            ' «Требует разбора»
            a = EngList("TEXTJOIN("";"";1;IF(" & ColRng(SH_STOCK, SC_STATE, 1, last) & "=""" & EI_ST_REVIEW & """;ROW(" & rgQ & ")-1;""""))", nItems)
            For i = 0 To nItems - 1
                v = pl(CLng(a(i)) - 1)
                If v <> "0" And v <> "-1" Then
                    b = Split(v, ",")
                    For j = 0 To UBound(b)
                        att(CLng(b(j)) - 1) = 1
                    Next j
                End If
            Next i
        End If
    End If
    ' 1.0: an address drawn in several cells of «Топология» (a long place, both sides of an aisle): every cell of it shows what
    ' lies there — it was only the first one, the others looked free and «Куда положить?» proposed them
    If UBound(gMpFirst) >= gMpN - 1 Then
        For i = 0 To gMpN - 1
            j = gMpFirst(i)
            If j <> i Then
                occ(i) = occ(j)
                iss(i) = iss(j)
                stale(i) = stale(j)
                att(i) = att(j)
                eis(i) = eis(j)
            End If
        Next i
    End If
    gMpOcc = occ
    gMpIss = iss
    gMpStale = stale
    gMpAtt = att
    gMpEIs = eis
    gMpOff = off
    gMpNOff = nOff
    gMpNDir = nDir
    gMpNoLvl = noLvl
    If IsArray(pl) Then gMpPlaceOf = pl
    gMpStatSeq = SysNum(SK_LAST_SEQ)
End Sub

' the places of the map a place text of «Наличие» names: «3» or «3,7» (numbers of the places, 1-based); −2 — no such place,
' −3 — a place without its level (5-12), 0 — no place written
Private Function PlacesOn(v As String, places As Object) As String
    Dim keys As Variant, i As Integer, idx As Long, out As String, noLevel As Boolean
    If v = "" Then
        PlacesOn = "0"
        Exit Function
    End If
    keys = PlaceList(v)
    For i = 0 To UBound(keys)
        idx = CollGet(places, keys(i))
        If idx >= 0 Then
            If InStr("," & out & ",", "," & (idx + 1) & ",") = 0 Then out = out & IIf(out <> "", ",", "") & (idx + 1)
        ElseIf UBound(Split(keys(i), "-")) = 1 Then
            noLevel = True
        End If
    Next i
    If out <> "" Then
        PlacesOn = out
    ElseIf noLevel Then
        PlacesOn = "-3"
    ElseIf PlacesAllDir(v) Then
        ' 1.0 (М2): a place of the directory «Места» (a floor, the yard) is known — not «не на карте»
        PlacesOn = "-4"
    Else
        PlacesOn = "-2"
    End If
End Function

' every part of place text v is a place of «Места»
Private Function PlacesAllDir(v As String) As Boolean
    Dim parts As Variant, i As Integer
    parts = PlaceParts(v)
    For i = 0 To UBound(parts)
        If PlaceKnown(PlaceKeyAny(parts(i))) <> 2 Then Exit Function
    Next i
    PlacesAllDir = (UBound(parts) >= 0)
End Function

' the column i (0 — the texts, 1 — the numbers of the places) of the table of MapKeys
Private Function mp_KeyRange(i As Integer) As String
    mp_KeyRange = ColRng(SH_MAP, MP_KEY_COL + i, 0, gMpKeyN - 1)
End Function

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
    Dim v As Double, k As Long
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
            Bucket = IIf(v = 0, 1, IIf(v = 1, 2, IIf(v <= 4, 3, 4)))
        End If
    Case "ATT"
        Bucket = IIf(gMpOcc(i) = 0, 0, IIf(gMpAtt(i) = 1, 2, 1))
    Case "FILL"
        ' 1.0 (М4): the EIs of the place against its capacity
        v = gMpOcc(i)
        If v = 0 Then
            Bucket = 0
        Else
            k = PlaceCap(gMpKey(i))
            Bucket = IIf(v < k, 1, IIf(v = k, 3, 4))
        End If
    Case "FIND"
        ' 1.0 (М5): the places of the goods shown
        If UBound(gMpFind) >= i Then
            If gMpFind(i) Then
                Bucket = 4
                Exit Function
            End If
        End If
        Bucket = IIf(gMpOcc(i) > 0, 1, 0)
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
    Case "FILL"
        cols = Array(RGB(244, 246, 248), RGB(198, 232, 201), gray, RGB(255, 192, 97), RGB(217, 71, 43))
        ink = Array(dk, dk, dk, dk, wh)
    Case "FIND"
        cols = Array(RGB(244, 246, 248), gray, gray, gray, RGB(47, 109, 181))
        ink = Array(dk, dk, dk, dk, wh)
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
    Case "STALE": txt = Array("пусто", "всё двигалось", "1 ЕИ без движения", "2–4 ЕИ", "5 и больше")
    Case "ATT": txt = Array("пусто", "в порядке", "требует разбора", "", "")
    Case "FILL": txt = Array("пусто", "есть место", "", "заполнено", "больше вместимости")
    Case "FIND": txt = Array("пусто", "другой товар", "", "", "«" & gMpFindName & "»")
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

' ================================================================ 1.0: the map edited on the map itself

' True while the map is edited (the mode is kept in a cell of the map: it outlives a save and a reopening)
Function MapEditing() As Boolean
    On Error GoTo EH
    If gMpEdit Then
        MapEditing = True
        Exit Function
    End If
    If gDoc.Sheets.hasByName(SH_MAP) Then
        If gDoc.Sheets.getByName(SH_MAP).getCellByPosition(MP_EDIT_COL, 0).getString() = "EDIT" Then
            gMpEdit = True
            MapEditing = True
        End If
    End If
EH:
End Function

Sub BtnMapEdit(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = MapEditStart()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Править карту»: the map shows the plain layout of «Топология» (no layer), its cells may be typed, cleared, coloured, merged;
' a place is «ряд-секция-ярус» (5-12-3) — typed as text, Calc does not turn it into a date
Function MapEditStart() As String
    Dim mp As Object, rng As Object, cf As Object, nr As Long, nc As Long, nf As Object, loc As New com.sun.star.lang.Locale, k As Long
    If MapEditing() Then
        MapEditStart = "ERR:карта уже в режиме правки — «Сохранить карту» или «Отменить правку»"
        Exit Function
    End If
    If gMpN = 0 Then MapBuild("OCC", True)
    mp = gDoc.Sheets.getByName(SH_MAP)
    If mp.isProtected() Then mp.unprotect(PROTECT_PWD)
    nr = IIf(gMpN > 0, gMpRows, 0)
    nc = IIf(gMpN > 0, gMpCols, 0)
    If nr < 40 Then nr = 40
    If nc < 30 Then nc = 30
    rng = mp.getCellRangeByPosition(0, MP_TOP, nc - 1, MP_TOP + nr - 1)
    cf = rng.ConditionalFormat
    cf.clear()
    rng.ConditionalFormat = cf
    mp.getCellRangeByPosition(MP_GRID_COL, MP_TOP, MP_GRID_COL + nc, MP_TOP + nr).clearContents(1 + 2 + 4)
    ' empty cells of the edit area: text, so that «5.12.3» stays a place
    loc.Language = "ru"
    loc.Country = "RU"
    nf = gDoc.getNumberFormats()
    k = nf.queryKey("@", loc, False)
    If k < 0 Then k = nf.addNew("@", loc)
    rng.NumberFormat = k
    mp.getCellByPosition(MP_EDIT_COL, 0).setString("EDIT")
    mp.getCellByPosition(0, MP_INFO_ROW).setString("РЕЖИМ ПРАВКИ КАРТЫ: впишите места «ряд-секция-ярус» (5-12-3) в клетки, очистите лишние, раскрасьте и объедините " _
        & "клетки как на складе — затем «Сохранить карту» (WMS проверит схему) или «Отменить правку».")
    mp.getCellByPosition(0, MP_SUM_ROW).setString("Карта — как лист «" & SH_TOPO & "»: сохранённая правка заменяет его (прежняя схема сохраняется в скрытом листе).")
    gMpEdit = True
    gDoc.getCurrentController().setActiveSheet(mp)
    gDoc.getCurrentController().select(mp.getCellByPosition(0, MP_TOP))
    MapEditStart = "OK:правка карты"
    gUiLastMsg = MapEditStart
End Function

Sub BtnMapCancel(Optional oEvent As Variant)
    WmsInit()
    MapEditCancel()
End Sub

Function MapEditCancel() As String
    Dim mp As Object
    If Not MapEditing() Then
        MapEditCancel = "SKIP:карта не в режиме правки"
        Exit Function
    End If
    mp = gDoc.Sheets.getByName(SH_MAP)
    If mp.isProtected() Then mp.unprotect(PROTECT_PWD)
    mp.getCellByPosition(MP_EDIT_COL, 0).setString("")
    gMpEdit = False
    MapEditCancel = MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"), True)
    gUiLastMsg = MapEditCancel
End Function

' the edit area of the map: the frame of the layout it was built from and every cell written since
Private Sub MapEditFrame(mp As Object, ByRef nr As Long, ByRef nc As Long)
    Dim q As Object, a As Variant, i As Long
    nr = IIf(gMpN > 0, gMpRows, 1)
    nc = IIf(gMpN > 0, gMpCols, 1)
    q = mp.getCellRangeByPosition(0, MP_TOP, MP_GRID_COL - 1, MP_TOP + 2999).queryContentCells(1 + 2 + 4 + 16)
    a = q.getRangeAddresses()
    For i = 0 To UBound(a)
        If a(i).EndRow - MP_TOP + 1 > nr Then nr = a(i).EndRow - MP_TOP + 1
        If a(i).EndColumn + 1 > nc Then nc = a(i).EndColumn + 1
    Next i
End Sub

' М3: the check of a layout (the data of its cells): places, repeated addresses, places without a tier («5-12»), and what the
' current map loses — the places where EIs lie. nPl — the places found; the report "" when nothing needs a look
Function TopoReport(d As Variant, ByRef nPl As Long) As String
    Dim r As Long, c As Long, k As String, seen As Object, dup As String, nDup As Long, two As String, nTwo As Long, s As String, i As Long
    Dim lost As String, nLost As Long, nLostEI As Long, t As String
    seen = New Collection
    nPl = 0
    For r = 0 To UBound(d)
        For c = 0 To UBound(d(r))
            If VarType(d(r)(c)) = 8 Then
                t = Trim(CStr(d(r)(c)))
                k = PlaceTextKey(t)
                If k <> "" Then
                    nPl = nPl + 1
                    If Not CollPut(seen, 1, k) Then
                        nDup = nDup + 1
                        If nDup <= 5 Then dup = dup & IIf(dup <> "", ", ", "") & k
                    End If
                ElseIf TwoPartPlace(t) Then
                    nTwo = nTwo + 1
                    If nTwo <= 5 Then two = two & IIf(two <> "", ", ", "") & t
                End If
            End If
        Next c
    Next r
    For i = 0 To gMpN - 1
        If CollGet(seen, gMpKey(i)) < 0 And gMpFirst(i) = i Then
            If UBound(gMpOcc) >= i Then
                If gMpOcc(i) > 0 Then
                    nLost = nLost + 1
                    nLostEI = nLostEI + gMpOcc(i)
                    If nLost <= 5 Then lost = lost & IIf(lost <> "", ", ", "") & gMpKey(i)
                End If
            End If
        End If
    Next i
    If nDup > 0 Then s = s & "• адрес повторяется: " & dup & IIf(nDup > 5, " …", "") & " — одно место должно быть в одной клетке" & Chr(10)
    If nTwo > 0 Then s = s & "• похоже на место без яруса: " & two & IIf(nTwo > 5, " …", "") & " — место пишется «ряд-секция-ярус» (5-12-3)" & Chr(10)
    If nLost > 0 Then s = s & "• исчезнут места, где лежат ЕИ (" & nLostEI & "): " & lost & IIf(nLost > 5, " …", "") _
        & " — эти ЕИ станут «не на карте» (их место не меняется)" & Chr(10)
    TopoReport = s
End Function

Private Function TwoPartPlace(t As String) As Boolean
    Dim i As Long, ch As String, nParts As Integer, inNum As Boolean, other As Boolean
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            If Not inNum Then nParts = nParts + 1
            inNum = True
        ElseIf ch = "-" Or ch = "." Or ch = "/" Then
            inNum = False
        Else
            other = True
        End If
    Next i
    TwoPartPlace = (nParts = 2 And Not other And Len(t) >= 3)
End Function

' «Проверить топологию» (М3): the layout of «Топология» checked, nothing changed
Sub BtnTopoCheck(Optional oEvent As Variant)
    Dim topo As Object, cur As Object, d As Variant, nPl As Long, s As String, a As Object
    WmsInit()
    topo = TopoSheet()
    cur = topo.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.getRangeAddress()
    d = topo.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).getDataArray()
    If gMpN = 0 And Not MapEditing() Then MapBuild("OCC")
    s = TopoReport(d, nPl)
    WmsUi.UiMessage("Топология: мест " & nPl & "." & IIf(s <> "", Chr(10) & Chr(10) & s, Chr(10) & "Повторов адресов и мест без яруса нет."))
    gUiLastMsg = "OK:мест " & nPl & IIf(s <> "", Chr(10) & s, "")
End Sub

Sub BtnMapSave(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = MapEditSave()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Сохранить карту»: the edited map is checked (М3) and replaces «Топология» (the previous layout is kept in a hidden sheet);
' the map is built again from it
Function MapEditSave() As String
    Dim mp As Object, topo As Object, nr As Long, nc As Long, d As Variant, nPl As Long, rep As String, k As Integer, prev As Object
    Dim cur As Object, a As Object, dest As New com.sun.star.table.CellAddress, srcA As New com.sun.star.table.CellRangeAddress, c As Long, r As Long
    Dim wasProt As Boolean, wasDoc As Boolean, cp As Object
    On Error GoTo EH
    If Not MapEditing() Then
        MapEditSave = "ERR:карта не в режиме правки — сначала «Править карту»"
        Exit Function
    End If
    mp = gDoc.Sheets.getByName(SH_MAP)
    MapEditFrame(mp, nr, nc)
    d = mp.getCellRangeByPosition(0, MP_TOP, nc - 1, MP_TOP + nr - 1).getDataArray()
    rep = TopoReport(d, nPl)
    If nPl = 0 Then
        MapEditSave = "ERR:на карте нет ни одного места «ряд-секция-ярус» (например 5-12-3) — впишите места или нажмите «Отменить правку»"
        Exit Function
    End If
    k = WmsPrimeUi.UiButtons("Сохранить карту", "Мест на карте: " & nPl & IIf(gMpN > 0, " (было " & gMpN & ")", "") & "." & IIf(rep <> "", Chr(10) & Chr(10) _
        & "Проверьте:" & Chr(10) & rep, Chr(10) & "Повторов адресов и мест без яруса нет.") & Chr(10) & "Карта заменит схему на листе «" & SH_TOPO _
        & "»; прежняя схема сохранится в скрытом листе «" & SH_TOPO_PREV & "».", Array("Сохранить карту", "Вернуться к правке"), Array(1, 0))
    If k <> 1 Then
        MapEditSave = "SKIP:правка продолжается"
        gUiLastMsg = MapEditSave
        Exit Function
    End If
    topo = TopoSheet()
    wasProt = topo.isProtected()
    If wasProt Then topo.unprotect(PROTECT_PWD)
    ' the previous layout: a hidden copy (one), then «Топология» is emptied and gets the map
    wasDoc = gDoc.isProtected()
    If wasDoc Then gDoc.unprotect(PROTECT_PWD)
    If gDoc.Sheets.hasByName(SH_TOPO_PREV) Then gDoc.Sheets.removeByName(SH_TOPO_PREV)
    gDoc.Sheets.copyByName(SH_TOPO, SH_TOPO_PREV, gDoc.Sheets.getCount())
    prev = gDoc.Sheets.getByName(SH_TOPO_PREV)
    prev.IsVisible = False
    If wasDoc Then gDoc.protect(PROTECT_PWD)
    cur = topo.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.getRangeAddress()
    topo.getRows().removeByIndex(0, a.EndRow + 1)
    dest.Sheet = SheetIndex(SH_TOPO)
    dest.Column = 0
    dest.Row = 0
    srcA.Sheet = SheetIndex(SH_MAP)
    srcA.StartColumn = 0
    srcA.StartRow = MP_TOP
    srcA.EndColumn = nc - 1
    srcA.EndRow = MP_TOP + nr - 1
    If mp.isProtected() Then mp.unprotect(PROTECT_PWD)
    topo.copyRange(dest, srcA)
    For c = 0 To nc - 1
        topo.getColumns().getByIndex(c).Width = mp.getColumns().getByIndex(c).Width
    Next c
    For r = 0 To nr - 1
        topo.getRows().getByIndex(r).Height = mp.getRows().getByIndex(MP_TOP + r).Height
    Next r
    cp = topo.getCellRangeByPosition(0, 0, nc - 1, nr - 1).CellProtection
    cp.IsLocked = False
    topo.getCellRangeByPosition(0, 0, nc - 1, nr - 1).CellProtection = cp
    If wasProt Then topo.protect(PROTECT_PWD)
    mp.getCellByPosition(MP_EDIT_COL, 0).setString("")
    gMpEdit = False
    gTkValid = False
    MapEditSave = MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"), True)
    If Left(MapEditSave, 3) = "OK:" Then MapEditSave = "OK:карта сохранена: мест " & nPl
    gUiLastMsg = MapEditSave
    Exit Function
EH:
    MapEditSave = "ERR:карта не сохранена: " & Error$ & " (строка " & Erl & ") — режим правки остаётся; «Отменить правку» вернёт карту по листу «" & SH_TOPO & "»"
    gUiLastMsg = MapEditSave
End Function

' ================================================================ 1.0 (М7) a row, a section or a place renamed

' key k («5-12-3») renamed when it begins with the parts oldP («5», «5-12» or «5-12-3» split by «-»): those parts become newP;
' "" when k does not begin with them
Function KeyRenamed(k As String, oldP As Variant, newP As Variant) As String
    Dim a As Variant, i As Integer
    If k = "" Then Exit Function
    a = Split(k, "-")
    If UBound(a) < UBound(oldP) Then Exit Function
    For i = 0 To UBound(oldP)
        If a(i) <> oldP(i) Then Exit Function
    Next i
    For i = 0 To UBound(newP)
        a(i) = newP(i)
    Next i
    KeyRenamed = Join(a, "-")
End Function

' the places oldP… of the map («Топология») and of the directory «Места»: counted (bDo False) or renamed to newP… (bDo True;
' «Топология» is first copied to the hidden «_TOPO_PREV», as at «Сохранить карту»). The cells of the map, nDir — the rows of «Места»
Function TopoRename(oldP As Variant, newP As Variant, bDo As Boolean, ByRef nDir As Long) As Long
    Dim topo As Object, cur As Object, d As Variant, r As Long, c As Long, k As String, nk As String, n As Long, wasProt As Boolean, wasDoc As Boolean
    Dim vals As Object, addrs As Variant, i As Long, rr As Long, cc As Long, sh As Object, last As Long, cell As Object, prev As Object
    nDir = 0
    If gDoc.Sheets.hasByName(SH_TOPO) Then
        topo = TopoSheet()
        cur = topo.createCursor()
        cur.gotoEndOfUsedArea(False)
        If cur.getRangeAddress().EndRow <= 2000 And cur.getRangeAddress().EndColumn <= 400 Then
            d = topo.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, cur.getRangeAddress().EndRow).getDataArray()
            vals = topo.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, cur.getRangeAddress().EndRow).queryContentCells(com.sun.star.sheet.CellFlags.DATETIME)
            addrs = vals.getRangeAddresses()
            If bDo Then
                wasDoc = gDoc.isProtected()
                If wasDoc Then gDoc.unprotect(PROTECT_PWD)
                If gDoc.Sheets.hasByName(SH_TOPO_PREV) Then gDoc.Sheets.removeByName(SH_TOPO_PREV)
                gDoc.Sheets.copyByName(SH_TOPO, SH_TOPO_PREV, gDoc.Sheets.getCount())
                prev = gDoc.Sheets.getByName(SH_TOPO_PREV)
                prev.IsVisible = False
                If wasDoc Then gDoc.protect(PROTECT_PWD)
                wasProt = topo.isProtected()
                If wasProt Then topo.unprotect(PROTECT_PWD)
            End If
            For r = 0 To UBound(d)
                For c = 0 To UBound(d(r))
                    If VarType(d(r)(c)) = 8 Then
                        nk = KeyRenamed(PlaceTextKey(CStr(d(r)(c))), oldP, newP)
                        If nk <> "" Then
                            n = n + 1
                            If bDo Then topo.getCellByPosition(c, r).setString(nk)
                        End If
                    End If
                Next c
            Next r
            For i = 0 To UBound(addrs)
                For rr = addrs(i).StartRow To addrs(i).EndRow
                    For cc = addrs(i).StartColumn To addrs(i).EndColumn
                        cell = topo.getCellByPosition(cc, rr)
                        nk = KeyRenamed(DateCellPlace(cell), oldP, newP)
                        If nk <> "" Then
                            n = n + 1
                            If bDo Then cell.setString(nk)
                        End If
                    Next cc
                Next rr
            Next i
            If bDo And wasProt Then topo.protect(PROTECT_PWD)
        End If
    End If
    If gDoc.Sheets.hasByName(SH_PLACES) Then
        sh = gDoc.Sheets.getByName(SH_PLACES)
        last = WmsOrders.LastRow(sh)
        If last > 20000 Then last = 20000
        If last >= 1 Then
            d = sh.getCellRangeByPosition(PD_PLACE, 1, PD_PLACE, last).getDataArray()
            For i = 0 To UBound(d)
                nk = KeyRenamed(PlaceKey(Trim(CStr(d(i)(0)))), oldP, newP)
                If nk <> "" Then
                    nDir = nDir + 1
                    If bDo Then sh.getCellByPosition(PD_PLACE, i + 1).setString(nk)
                End If
            Next i
        End If
    End If
    If bDo Then
        gTkValid = False
        gPdValid = False
        gMpSig = ""
    End If
    TopoRename = n
End Function

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
    If gMpEdit Then Exit Sub
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
        s = s & IIf(s <> "", "; ", "") & st(SC_EI) & " " & st(SC_NAME) & " — " & WmsPack.QtyWithParts(n, CDbl(st(SC_QTY)), CStr(st(SC_UNIT)))
    Next i
    If UBound(eis) <= 0 Then s = "пусто — место свободно"
    mp.getCellByPosition(0, MP_INFO_ROW).setString("Место " & gMpKey(idx) & ": " & s & IIf(UBound(eis) > 20, " … ещё " & (UBound(eis) - 20), ""))
    ViewDone(mp)
    gUiLastMsg = "OK:" & gMpKey(idx) & ": " & s
    Exit Sub
EH:
    gUiLastMsg = "ERR:" & Error$ & " (строка " & Erl & ")"
End Sub

' 1.0 (М6): a double click on a place of the map — «Поиск» lists what lies there and the history of the place
Function OnMapDoubleClick(Optional oTarget As Variant) As Boolean
    Dim a As Object, idx As Long
    On Error GoTo EH
    WmsInit()
    If MapEditing() Then Exit Function
    a = oTarget.getCellAddress()
    If a.Row < MP_TOP Or gMpN = 0 Or IsNull(gMpCells) Then Exit Function
    idx = CollGet(gMpCells, a.Row & ":" & a.Column)
    If idx < 0 Then Exit Function
    OnMapDoubleClick = True
    WmsFind.FindRun(gMpKey(idx))
    WmsView.GoCell(SH_FIND, 3, 2)
    gUiLastMsg = "OK:" & gMpKey(idx)
    Exit Function
EH:
End Function

' the map is shown: built when it was never built in this session, and (1.0) counted again when operations were posted since
' it was counted — it showed the warehouse of the moment it was built (the layer of the last time)
Sub OnMapFocus(Optional oEvent As Variant)
    On Error GoTo EH
    WmsInit()
    If gMpBusy Or MapEditing() Then Exit Sub
    ' 1.0 (С5): the map saved with the book is up to date — its places read, nothing counted or painted again
    If gMpN = 0 Then
        If MapLazy() Then Exit Sub
    End If
    If gMpN = 0 Or gMpStatSeq <> SysNum(SK_LAST_SEQ) Then
        ' 1.0 (С4): a map that took more than 5 s is built again only by «Обновить»
        If WmsCare.MapSlow() Then
            MapStale()
            Exit Sub
        End If
        MapBuild(IIf(gMpLayer <> "", gMpLayer, "OCC"))
    End If
EH:
End Sub

' С5: the map of the book is of the last operation (its stamp): the layout and the counts read (for the clicks and «Куда
' положить?»), the painting kept as it was saved. True when so
Function MapLazy() As Boolean
    Dim mp As Object, a As Variant, why As String
    On Error GoTo EH
    mp = gDoc.Sheets.getByName(SH_MAP)
    a = Split(mp.getCellByPosition(MP_STAMP_COL, 0).getString(), "|")
    If UBound(a) < 2 Then Exit Function
    If Val(a(0)) <> SysNum(SK_LAST_SEQ) Then Exit Function
    gMpBusy = True
    mp = ViewOpen(SH_MAP)
    gMpLazyNow = True
    why = MapLayout(mp, False)
    gMpLazyNow = False
    If why = "" Then
        MapStats()
        gMpLayer = a(1)
        MapLazy = (gVwEngFail = "")
    End If
    ViewDone(mp)
    gMpBusy = False
    Exit Function
EH:
    gMpBusy = False
    MapLazy = False
End Function

' С4: the map is older than the last operation and slow to build — said on it, built by «Обновить»
Private Sub MapStale()
    Dim mp As Object, a As Variant
    On Error Resume Next
    mp = ViewOpen(SH_MAP)
    a = Split(mp.getCellByPosition(MP_STAMP_COL, 0).getString(), "|")
    mp.getCellByPosition(0, MP_INFO_ROW).setString("Карта " & IIf(UBound(a) >= 2, "от " & a(2), "старая") & ": после неё проводились операции. Карта строится " _
        & Format(WmsCare.TimingGet("MAP") / 1000, "0") & " с — обновляется только кнопкой «Обновить».")
    ViewDone(mp)
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
    Dim k As Integer, c As Integer, n As Integer, pick As String, s As String, cur As String, ans As Integer
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
        cur = Trim(sh.getCellByPosition(cPlace, r).getString())
        If cur <> "" And InStr(", " & LCase(cur) & ",", ", " & LCase(pick) & ",") = 0 And LCase(cur) <> LCase(pick) Then
            ' D-100: the goods may take more than one cell — the place is added to the list or replaces it
            ans = WmsPrimeUi.UiAsk3b("В строке " & (r + 1) & " уже указано место «" & cur & "»." & Chr(10) & "Товар может занимать несколько мест: «" _
                & cur & ", " & pick & "».", "Добавить " & pick & " (будет «" & cur & ", " & pick & "»)", "Заменить на " & pick, "Ничего не менять")
            If ans = 2 Then
                gUiLastMsg = "SKIP:место не изменено"
                WherePut = gUiLastMsg
                Exit Function
            End If
            If ans = 6 Then pick = cur & ", " & pick
        End If
        sh.getCellByPosition(cPlace, r).setString(pick)
        gUiLastMsg = "OK:место " & pick & " записано в строку " & (r + 1)
    End If
    WherePut = gUiLastMsg
End Function

' the places for goods nm (article art): lines «place|text for a person», at most 8; ERR:… when there is nothing to propose
Function WhereSuggest(nm As String, art As String) As String
    Dim nextEI As Long, r As Long, r1x As Long, i As Long, st As Variant, same As Object, sKeys() As String, sCnt() As Long, nS As Long, idx As Long
    Dim pk As String, out As String, nOut As Integer, j As Long, best As Long, bestD As Double, d As Double, a As Variant, b As Variant, used() As Boolean
    Dim anchorKey As String, k As Integer, pls As Variant, desc As Object, found As Object, addrs As Variant, q As Long
    Dim plan() As Long, cap() As Long, anchorIdx As Long
    nextEI = CLng(SysNum(SK_NEXT_EI))
    same = New Collection
    ReDim sKeys(50)
    ReDim sCnt(50)
    ' where the same goods lie now (EIs with a balance). 1.0: the rows of the goods are found by the search of the Calc engine
    ' (the whole cell of the article or the name, any case) and only they are read — it was a loop over all of «Наличие»
    If nextEI > 1 Then
        desc = WmsIssue.StockSheet().createSearchDescriptor()
        desc.SearchString = "^" & WmsFind.EscRegex(IIf(art <> "", art, nm)) & "$"
        desc.SearchRegularExpression = True
        desc.SearchCaseSensitive = False
        found = WmsIssue.StockSheet().getCellRangeByPosition(IIf(art <> "", SC_ART, SC_NAME), 1, IIf(art <> "", SC_ART, SC_NAME), nextEI - 1).findAll(desc)
        If Not IsNull(found) Then
            addrs = found.getRangeAddresses()
            For q = 0 To UBound(addrs)
                For r = addrs(q).StartRow To addrs(q).EndRow
                    st = WmsIssue.StockSheet().getCellRangeByPosition(0, r, SC_STATE, r).getDataArray()(0)
                    If VarType(st(SC_QTY)) = 5 Then
                        If st(SC_QTY) > 0.0000005 And CStr(st(SC_STATE)) <> EI_ST_STORNO Then
                            ' an EI on several places («1-2-3, 1-2-4») counts on each of them
                            pls = PlaceList(Trim(CStr(st(SC_PLACE))))
                            If UBound(pls) < 0 And Trim(CStr(st(SC_PLACE))) <> "" Then pls = Array(Trim(CStr(st(SC_PLACE))))
                            For j = 0 To UBound(pls)
                                pk = pls(j)
                                idx = CollGet(same, LCase(pk))
                                If idx < 0 And nS <= 50 Then
                                    idx = nS
                                    CollPut(same, idx, LCase(pk))
                                    sKeys(idx) = pk
                                    nS = nS + 1
                                End If
                                If idx >= 0 Then sCnt(idx) = sCnt(idx) + 1
                            Next j
                        End If
                    End If
                Next r
            Next q
        End If
    End If
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
        ' 1.0 (М4, М9): a place has room while its EIs and the rows not yet posted that name it are fewer than its capacity;
        ' «рядом» is measured on the map itself when the goods lie on it (the cells of the schema), otherwise by the address
        MapPlanned(plan)
        anchorIdx = -1
        If anchorKey <> "" Then
            For j = 0 To gMpN - 1
                If gMpKey(j) = anchorKey Then
                    anchorIdx = j
                    Exit For
                End If
            Next j
        End If
        ReDim cap(gMpN)
        For j = 0 To gMpN - 1
            cap(j) = PlaceCap(gMpKey(j))
        Next j
        ' the places where the goods lie are already named above
        For i = 0 To nS - 1
            For j = 0 To gMpN - 1
                If gMpKey(j) = PlaceKey(sKeys(i)) Then used(j) = True
            Next j
        Next i
        For k = 1 To 8 - nOut
            best = -1
            bestD = 1E+30
            For j = 0 To gMpN - 1
                If gMpOcc(j) + plan(j) < cap(j) And Not used(j) And gMpFirst(j) = j And Not PlaceClosed(gMpKey(j)) Then
                    If anchorIdx >= 0 Then
                        ' the same row of racks: the address (another tier of the same section first, then the next sections);
                        ' another row: as far as it is on the map (rows numbered apart may stand side by side)
                        d = SameRowDistance(anchorKey, gMpKey(j))
                        If d < 0 Then d = 100 + Abs(gMpRow(j) - gMpRow(anchorIdx)) * 3 + Abs(gMpCol(j) - gMpCol(anchorIdx))
                    Else
                        d = PlaceDistance(anchorKey, gMpKey(j), j)
                    End If
                    If d < bestD Then
                        bestD = d
                        best = j
                    End If
                End If
            Next j
            If best < 0 Then Exit For
            used(best) = True
            out = out & IIf(out <> "", Chr(10), "") & gMpKey(best) & "|" & gMpKey(best) & IIf(gMpOcc(best) + plan(best) = 0, " — свободно", " — есть место (занято " _
                & (gMpOcc(best) + plan(best)) & " из " & cap(best) & ")") & IIf(anchorKey <> "", IIf(bestD < IIf(anchorIdx >= 0, 21, 1000), ", рядом", ""), "")
            nOut = nOut + 1
        Next k
    End If
    If out = "" Then
        WhereSuggest = "ERR:подсказать нечего: этого товара на складе нет, а карта склада не задана (лист «" & SH_TOPO & "») — впишите место вручную"
    Else
        WhereSuggest = out
    End If
End Function

' 1.0 (М9): two places of one row of racks — 10 per section apart and 1 per tier; −1 for places of two rows (or not numbered)
Private Function SameRowDistance(sFrom As String, sTo As String) As Double
    Dim a As Variant, b As Variant
    SameRowDistance = -1
    a = Split(sFrom, "-")
    b = Split(sTo, "-")
    If UBound(a) < 2 Or UBound(b) < 2 Then Exit Function
    If a(0) <> b(0) Or Not IsNumeric(a(1)) Or Not IsNumeric(b(1)) Or Not IsNumeric(a(2)) Or Not IsNumeric(b(2)) Then Exit Function
    SameRowDistance = Abs(CDbl(a(1)) - CDbl(b(1))) * 10 + Abs(CDbl(a(2)) - CDbl(b(2)))
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

' ================================================================ 1.0: the known places (М2, the check of a typed place)

' «Топология» or «Места» changed: their keys are read again at the next use
Sub OnTopoChange(Optional oEvent As Variant)
    gTkValid = False
End Sub

Sub OnPlacesChange(Optional oEvent As Variant)
    gPdValid = False
End Sub

' the key of a place as the directory and the checks compare it: the key of «ряд-секция-ярус» (PlaceKey) or, for a named
' place without numbers («Улица», «Офис»), its text in upper case
Function PlaceKeyAny(ByVal s As String) As String
    s = Trim(s)
    PlaceKeyAny = PlaceKey(s)
    If PlaceKeyAny = "" Then PlaceKeyAny = UCase(s)
End Function

' the keys of the places of «Топология»: the texts that are places, and the cells Calc took for a date when «5.12.3» was typed
' (read back as 5-12-3)
Private Sub TopoKeysLoad()
    Dim topo As Object, cur As Object, d As Variant, r As Long, c As Long, pk As String, vals As Object, addrs As Variant, k As Long, cell As Object
    Dim ft As Long, x As Double, rr As Long, cc As Long
    gTk = New Collection
    gTkN = 0
    ReDim gTkList(0)
    gTkValid = True
    If Not gDoc.Sheets.hasByName(SH_TOPO) Then Exit Sub
    topo = TopoSheet()
    cur = topo.createCursor()
    cur.gotoEndOfUsedArea(False)
    If cur.getRangeAddress().EndRow > 2000 Or cur.getRangeAddress().EndColumn > 400 Then Exit Sub
    d = topo.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, cur.getRangeAddress().EndRow).getDataArray()
    ReDim gTkList(1000)
    For r = 0 To UBound(d)
        For c = 0 To UBound(d(r))
            If VarType(d(r)(c)) = 8 Then
                pk = PlaceTextKey(CStr(d(r)(c)))
                If pk <> "" Then TkAdd(pk)
            End If
        Next c
    Next r
    vals = topo.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, cur.getRangeAddress().EndRow).queryContentCells(com.sun.star.sheet.CellFlags.DATETIME)
    addrs = vals.getRangeAddresses()
    For k = 0 To UBound(addrs)
        For rr = addrs(k).StartRow To addrs(k).EndRow
            For cc = addrs(k).StartColumn To addrs(k).EndColumn
                pk = DateCellPlace(topo.getCellByPosition(cc, rr))
                If pk <> "" Then TkAdd(pk)
            Next cc
        Next rr
    Next k
End Sub

Private Sub TkAdd(pk As String)
    If CollPut(gTk, 1, pk) Then
        ' block form: after a one-line «If … Then ReDim …» LibreOffice Basic takes the next line into the If
        If gTkN > UBound(gTkList) Then
            ReDim Preserve gTkList(2 * gTkN + 10)
        End If
        gTkList(gTkN) = pk
        gTkN = gTkN + 1
    End If
End Sub

' a cell Calc made a date of «5.12.3» (05.12.2003): the place 5-12-3; "" for any other cell
Function DateCellPlace(cell As Object) As String
    Dim ft As Long, x As Double, dt As Date
    If cell.getType() <> com.sun.star.table.CellContentType.VALUE Then Exit Function
    ft = gDoc.getNumberFormats().getByKey(cell.NumberFormat).Type
    If (ft And com.sun.star.util.NumberFormat.DATE) = 0 Then Exit Function
    x = cell.getValue()
    If x < 1 Or x <> Int(x) Then Exit Function
    dt = CDate(x)
    DateCellPlace = Day(dt) & "-" & Month(dt) & "-" & (Year(dt) Mod 100)
End Function

Private Sub PlacesDirLoad()
    Dim sh As Object, last As Long, d As Variant, i As Long, k As String
    gPd = New Collection
    gPdN = 0
    ReDim gPdClosed(0)
    ReDim gPdName(0)
    gPdValid = True
    If Not gDoc.Sheets.hasByName(SH_PLACES) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_PLACES)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    If last > 20000 Then last = 20000
    d = sh.getCellRangeByPosition(0, 1, PD_NOTE, last).getDataArray()
    ReDim gPdClosed(last)
    ReDim gPdName(last)
    ReDim gPdCap(last)
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(PD_PLACE))) <> "" Then
            k = PlaceKeyAny(CStr(d(i)(PD_PLACE)))
            If CollPut(gPd, gPdN, k) Then
                gPdName(gPdN) = Trim(CStr(d(i)(PD_PLACE)))
                gPdCap(gPdN) = 1
                If VarType(d(i)(PD_CAP)) = 5 Then
                    If d(i)(PD_CAP) >= 1 Then gPdCap(gPdN) = CLng(Int(d(i)(PD_CAP)))
                End If
                gPdClosed(gPdN) = (LCase(Left(Trim(CStr(d(i)(PD_CLOSED))), 2)) = "да" Or LCase(Trim(CStr(d(i)(PD_CLOSED)))) = "закрыто")
                gPdN = gPdN + 1
            End If
        End If
    Next i
End Sub

' where place key k is known: 1 — on the map («Топология»), 2 — in «Места», 0 — nowhere
Function PlaceKnown(k As String) As Integer
    If Not gTkValid Then TopoKeysLoad()
    If Not gPdValid Then PlacesDirLoad()
    If CollGet(gTk, k) >= 0 Then
        PlaceKnown = 1
    ElseIf CollGet(gPd, k) >= 0 Then
        PlaceKnown = 2
    End If
End Function

' 1.0 (М4): the capacity of place key k in cells — «Места» D, 1 when the place is not there or D is empty
Function PlaceCap(k As String) As Long
    Dim i As Long
    If Not gPdValid Then PlacesDirLoad()
    i = CollGet(gPd, k)
    PlaceCap = 1
    If i >= 0 Then PlaceCap = gPdCap(i)
End Function

' 1.0 (М9): the places already promised by rows not yet posted («Заказы» U, «Иной приход» I, «Возврат» J, «Корректировки» J):
' plan(j) — how many of them name place j of the map
Sub MapPlanned(plan() As Long)
    Dim specs As Variant, i As Integer, sh As Object, last As Long, a As Variant, nItems As Long, k As Long, pls As Variant, j As Long, idx As Long
    Dim keyIdx As Object
    ReDim plan(gMpN)
    If gMpN = 0 Then Exit Sub
    keyIdx = New Collection
    For j = 0 To gMpN - 1
        If gMpFirst(j) = j Then CollPut(keyIdx, j, gMpKey(j))
    Next j
    specs = Array(Array(SH_ORDERS, OC_PLACE, OC_EI), Array(SH_SPECIAL, XC_PLACE, XC_NO), Array(SH_RETURNS, RC_PLACE, RC_NO), Array(SH_ADJUST, AC_PLACE, AC_NO))
    For i = 0 To UBound(specs)
        If gDoc.Sheets.hasByName(specs(i)(0)) Then
            sh = gDoc.Sheets.getByName(specs(i)(0))
            last = WmsOrders.LastRow(sh)
            If last >= 1 Then
                a = EngList("TEXTJOIN("";"";1;IF((" & ColRng(specs(i)(0), specs(i)(2), 1, last) & "="""")*(" & ColRng(specs(i)(0), specs(i)(1), 1, last) _
                    & "<>"""");" & ColRng(specs(i)(0), specs(i)(1), 1, last) & ";""""))", nItems)
                For k = 0 To nItems - 1
                    pls = PlaceList(CStr(a(k)))
                    For j = 0 To UBound(pls)
                        idx = CollGet(keyIdx, CStr(pls(j)))
                        If idx >= 0 Then plan(idx) = plan(idx) + 1
                    Next j
                Next k
            End If
        End If
    Next i
End Sub

' the place is in «Места» and marked «закрыто»
Function PlaceClosed(k As String) As Boolean
    Dim i As Long
    If Not gPdValid Then PlacesDirLoad()
    i = CollGet(gPd, k)
    If i >= 0 Then PlaceClosed = gPdClosed(i)
End Function

' 1.0: a place typed at a posting checked against the map and «Места» (there was no check: «1-24» instead of «1-2-4» went
' into the books). Nothing is asked when the book has neither a map nor a directory. A place nowhere known or closed:
' «Записать как есть», the nearest places of the map, «Не проводить». sPlace — in: as typed, out: as chosen. sNot —
' «приход не проведён». "" to post, SKIP:… otherwise
Function PlaceAsk(ByRef sPlace As String, sNot As String) As String
    Dim parts As Variant, i As Integer, k As String, bad As String, nBad As Integer, badKey As String, closed As String, sText As String
    Dim aNear() As String, nLike As Integer, j As Long, labels() As String, codes() As Integer, n As Integer, c As Integer, sep As String
    On Error GoTo EH
    If Trim(sPlace) = "" Then Exit Function
    If Not gTkValid Then TopoKeysLoad()
    If Not gPdValid Then PlacesDirLoad()
    If gTkN = 0 And gPdN = 0 Then Exit Function
    parts = PlaceParts(sPlace)
    For i = 0 To UBound(parts)
        k = PlaceKeyAny(parts(i))
        If PlaceKnown(k) = 0 Then
            bad = bad & IIf(bad <> "", ", ", "") & "«" & Trim(parts(i)) & "»"
            nBad = nBad + 1
            badKey = k
        ElseIf PlaceClosed(k) Then
            closed = closed & IIf(closed <> "", ", ", "") & "«" & Trim(parts(i)) & "»"
        End If
    Next i
    If nBad = 0 And closed = "" Then Exit Function
    If nBad > 0 Then
        sText = IIf(nBad = 1, "Места " & bad & " нет", "Мест " & bad & " нет") & " на карте склада (лист «" & SH_TOPO & "») и в справочнике «" & SH_PLACES & "»."
        If nBad = 1 Then
            ReDim aNear(2)
            For j = 0 To gTkN - 1
                If nLike > 2 Then Exit For
                If KeyNear(badKey, gTkList(j)) Then
                    aNear(nLike) = gTkList(j)
                    nLike = nLike + 1
                End If
            Next j
            If nLike > 0 Then
                ReDim Preserve aNear(nLike - 1)
                sText = sText & Chr(10) & "Похожие места карты: " & Join(aNear, ", ") & "."
            End If
        End If
        sText = sText & Chr(10) & "Место вне стеллажей (пол, улица, офис) можно один раз записать на лист «" & SH_PLACES & "» — тогда WMS будет его знать."
    End If
    If closed <> "" Then sText = sText & IIf(sText <> "", Chr(10), "") & "Место " & closed & " в справочнике «" & SH_PLACES & "» отмечено «закрыто»."
    n = nLike + 2
    ReDim labels(n - 1)
    ReDim codes(n - 1)
    labels(0) = "Записать «" & Trim(sPlace) & "» как есть"
    codes(0) = 0
    For c = 0 To nLike - 1
        labels(c + 1) = "Записать " & aNear(c)
        codes(c + 1) = 10 + c
    Next c
    labels(n - 1) = "Не проводить — исправлю место"
    codes(n - 1) = 1
    c = WmsPrimeUi.UiButtons("WMS — место хранения", sText, labels, codes)
    If c >= 10 Then
        ' the unknown part replaced by the chosen place, the other parts as typed
        sep = ""
        k = ""
        For i = 0 To UBound(parts)
            If PlaceKnown(PlaceKeyAny(parts(i))) = 0 Then k = k & sep & aNear(c - 10) Else k = k & sep & Trim(parts(i))
            sep = ", "
        Next i
        sPlace = k
    ElseIf c <> 0 Then
        PlaceAsk = "SKIP:" & sNot & " — проверьте место «" & Trim(sPlace) & "»"
    End If
    Exit Function
EH:
    ' the check never stops a posting by itself
End Function

' the parts of a place text as people write a list («1-2-3, 1-2-4; Улица»); one part when it is one place («5,12,3»); 1.0 (М10):
' a range «1-2-3…5» — its places
Function PlaceParts(ByVal s As String) As Variant
    PlaceParts = PlaceItems(s)
End Function

' two keys one typing slip apart: one character more, less or other, or a lost / doubled separator («1-24» ~ «1-2-4»)
Private Function KeyNear(a As String, b As String) As Boolean
    If a = b Then Exit Function
    If Replace(a, "-", "") = Replace(b, "-", "") Then
        KeyNear = True
        Exit Function
    End If
    If Abs(Len(a) - Len(b)) > 1 Then Exit Function
    KeyNear = (KeyDist1(a, b) <= 1)
End Function

Private Function KeyDist1(a As String, b As String) As Integer
    Dim i As Integer, j As Integer, d As Integer
    i = 1
    j = 1
    Do While i <= Len(a) And j <= Len(b)
        If Mid(a, i, 1) <> Mid(b, j, 1) Then
            d = d + 1
            If d > 1 Then
                KeyDist1 = d
                Exit Function
            End If
            If Len(a) > Len(b) Then
                i = i + 1
            ElseIf Len(b) > Len(a) Then
                j = j + 1
            Else
                i = i + 1
                j = j + 1
            End If
        Else
            i = i + 1
            j = j + 1
        End If
    Loop
    KeyDist1 = d + (Len(a) - i + 1) + (Len(b) - j + 1)
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

Function TestMapBucket(sKey As String, sLayer As String) As String
    Dim i As Long
    For i = 0 To gMpN - 1
        If gMpKey(i) = sKey Then
            TestMapBucket = CStr(Bucket(sLayer, i))
            Exit Function
        End If
    Next i
    TestMapBucket = "-"
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

' every cell of place sKey on the map: «i|EIs;i|EIs» (an address drawn twice shows the same)
Function TestMapEIsAll(sKey As String) As String
    Dim i As Long, r As String
    For i = 0 To gMpN - 1
        If gMpKey(i) = sKey Then r = r & IIf(r <> "", ";", "") & i & "|" & gMpEIs(i)
    Next i
    TestMapEIsAll = r
End Function

' the known places near key k (the suggestions of PlaceAsk) and how many keys the map has: «n|k1,k2»
Function TestPlaceNear(k As String) As String
    Dim j As Long, r As String
    If Not gTkValid Then TopoKeysLoad()
    For j = 0 To gTkN - 1
        If KeyNear(k, gTkList(j)) Then r = r & IIf(r <> "", ",", "") & gTkList(j)
    Next j
    TestPlaceNear = gTkN & "|" & r
End Function

' test seam: «мест|LAST_SEQ подсчёта|слой» of the map in memory
Function TestMapState() As String
    TestMapState = gMpN & "|" & gMpStatSeq & "|" & gMpLayer
End Function
