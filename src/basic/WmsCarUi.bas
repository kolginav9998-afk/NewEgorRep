' WmsCarUi — кнопки листа «Приход авто» (задание M6 PRIME, §1): ПРИЕХАЛ, УЕХАЛ и компактные «Найти», «Сегодня», «Все»,
' «Исправить», «Отменить визит». Рабочий сценарий — две ячейки и две кнопки: вписал машину → поставщика → ПРИЕХАЛ;
' при выезде выделил строку машины → УЕХАЛ. Успешная работа окон не открывает: новая строка выделяется, итог виден в
' «Контроль». Окно — когда действие невозможно, когда нужно выбрать машину из нескольких на территории и для исправления.
Option Explicit

' test seam (TEST books only): the answer of the next choice window (the line; -1 — «Отмена»), the text of the next
' «Найти», the answers of the next window of «Исправить» (CAR_KEEP — as proposed; the window is left with «Отмена» when
' the first answer is CAR_CANCEL_DLG)
Global gCarPickSet As Boolean
Global gCarPick As Integer
Global gCarFindSet As Boolean
Global gCarFind As String
Global gCarFixSet As Boolean
Global gCarFixIn(3) As String
' what the last window proposed: the lines of the choice / the values of «Исправить» (the test seam reads it)
Global gCarShown As String
Private Const CAR_KEEP = "<как предложено>"
Private Const CAR_CANCEL_DLG = "<отмена>"
Private Const CAR_DB = "WMS_CARS"

' ================================================================ helpers

Private Sub Refuse(s As String)
    UiMessage(s)
    gUiLastMsg = "ERR:" & s
End Sub

Private Function Payload(res As String) As String
    Payload = Mid(res, InStr(res, ":") + 1)
End Function

' after an operation: a blocked or critical result is also shown on «Главная»; a refusal in a window
Private Sub AfterOperation(sAction As String, res As String)
    If Left(res, 3) <> "OK:" Then
        If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 7) = "ERR-RB:" Then UiRefresh(sAction & ": " & res)
        UiMessage(Payload(res))
    End If
    gUiLastMsg = res
End Sub

' the row of the cursor on «Приход авто» (0-based); -1 with why when another sheet is active or several rows are selected
Private Function CursorRow(ByRef why As String) As Long
    Dim sel As Object, a As Object
    WmsInit()
    On Error GoTo EH
    CursorRow = -1
    If gDoc.getCurrentController().getActiveSheet().getName() <> SH_CARS Then
        why = "Перейдите на лист «" & SH_CARS & "»."
        Exit Function
    End If
    sel = gDoc.getCurrentSelection()
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        why = "Выделите одну строку машины (сейчас выделено несколько областей)."
        Exit Function
    End If
    a = sel.getRangeAddress()
    If a.EndRow > a.StartRow Then
        why = "Выделите одну строку машины (сейчас выделено несколько строк)."
        Exit Function
    End If
    CursorRow = a.StartRow
    Exit Function
EH:
    why = "Поставьте курсор в строку машины на листе «" & SH_CARS & "»."
    CursorRow = -1
End Function

' the row r is shown (a filter of the table may hide a new row) and selected A..M
Private Sub ShowAndSelect(r As Long)
    Dim sh As Object, um As Object, db As Object, locked As Boolean
    On Error GoTo EH
    sh = WmsCar.CarsSheet()
    um = gDoc.getUndoManager()
    If Not sh.getRows().getByIndex(r).IsVisible Then
        um.lock()
        locked = True
        If gDoc.DatabaseRanges.hasByName(CAR_DB) Then gDoc.DatabaseRanges.getByName(CAR_DB).refresh()
        If Not sh.getRows().getByIndex(r).IsVisible Then sh.getRows().getByIndex(r).IsVisible = True
        um.unlock()
        locked = False
    End If
    gDoc.getCurrentController().select(sh.getCellRangeByPosition(0, r, CC_NOTE, r))
    Exit Sub
EH:
    On Error Resume Next
    If locked Then um.unlock()
End Sub

' «№ 12 · А123ВС · Петрович · с 09:15 (1 ч 25 мин)»
Private Function OpenLine(n As Long) As String
    Dim v As Variant, d As Variant, s As String
    v = WmsCar.VisitCells(n)
    d = WmsCar.CarRow(n)
    s = "№ " & n & " · " & v(CC_PLATE) & " · " & v(CC_SUP)
    If VarType(d(CR_ARR)) = 5 Then
        If Int(d(CR_ARR)) <> Int(WmsCar.CarNow()) Then s = s & " · с " & Format(CDate(d(CR_ARR)), "DD.MM") & " " & WmsCar.HhMm(CDbl(d(CR_ARR))) _
            Else s = s & " · с " & WmsCar.HhMm(CDbl(d(CR_ARR)))
        s = s & " (" & WmsCar.DurText(WmsCar.CarNow() - d(CR_ARR)) & ")"
    End If
    OpenLine = s
End Function

' ================================================================ ПРИЕХАЛ

Sub BtnCarArrive(Optional oEvent As Variant)
    Dim sh As Object, res As String, a As Variant
    WmsInit()
    If gDoc.getCurrentController().getActiveSheet().getName() <> SH_CARS Then
        Refuse("Перейдите на лист «" & SH_CARS & "».")
        Exit Sub
    End If
    sh = WmsCar.CarsSheet()
    res = WmsCar.CarArrive(sh.getCellByPosition(CP_PLATE_COL, CP_IN_ROW).getString(), sh.getCellByPosition(CP_SUP_COL, CP_IN_ROW).getString())
    AfterOperation("ПРИЕХАЛ", res)
    If Left(res, 3) = "OK:" Then
        a = Split(res, "|")
        ShowAndSelect(CLng(a(2)))
    End If
End Sub

' ================================================================ УЕХАЛ

' The selected row when it is a visit (an open one; a closed or cancelled one is refused). Without such a row: the only
' vehicle on the territory is proposed, several are offered in a short list.
Sub BtnCarDepart(Optional oEvent As Variant)
    Dim r As Long, n As Long, why As String, opens As Variant, texts() As String, i As Integer, pick As Integer, res As String
    Dim d As Variant, a As Variant
    r = CursorRow(why)
    If r < 0 Then
        Refuse(why)
        Exit Sub
    End If
    If r >= CAR_FIRST Then n = WmsCar.VisitAtRow(r)
    If n > 0 Then
        d = WmsCar.CarRow(n)
        If d(CR_STATE) <> CR_OPEN Then
            res = WmsCar.CarDepart(n)
            AfterOperation("УЕХАЛ", res)
            Exit Sub
        End If
        If (WmsCar.CarNow() - d(CR_ARR)) * 1440 < CAR_QUICK_MIN Then
            If Not UiConfirm("Машина " & WmsCar.VisitText(n) & " приехала меньше " & CAR_QUICK_MIN & " минут назад. Записать её выезд?") Then
                gUiLastMsg = "CANCEL:выезд не записан"
                Exit Sub
            End If
        End If
    Else
        opens = WmsCar.OpenVisits()
        If UBound(opens) < 0 Then
            Refuse("На территории нет машин — отмечать выезд некому.")
            Exit Sub
        End If
        If UBound(opens) = 0 Then
            n = opens(0)
            If Not UiConfirm("Уехала машина " & WmsCar.VisitText(n) & "?") Then
                gUiLastMsg = "CANCEL:выезд не записан"
                Exit Sub
            End If
        Else
            ReDim texts(UBound(opens))
            For i = 0 To UBound(opens)
                texts(i) = OpenLine(CLng(opens(i)))
            Next i
            pick = 0
            If Not CarChoose("WMS — УЕХАЛ", "На территории машин: " & (UBound(opens) + 1) & ". Какая уехала?", texts, pick) Then
                gUiLastMsg = "CANCEL:выезд не записан"
                Exit Sub
            End If
            n = opens(pick)
        End If
    End If
    res = WmsCar.CarDepart(n)
    AfterOperation("УЕХАЛ", res)
    If Left(res, 3) = "OK:" Then
        a = Split(res, "|")
        ShowAndSelect(CLng(a(2)))
    End If
End Sub

' ================================================================ «Исправить», «Отменить визит»

Sub BtnCarFix(Optional oEvent As Variant)
    Dim r As Long, n As Long, why As String, d As Variant, v As Variant, f(3) As String, g(3) As String, res As String, bOpen As Boolean
    Dim sInfo As String
    r = CursorRow(why)
    If r < 0 Then
        Refuse(why)
        Exit Sub
    End If
    If r >= CAR_FIRST Then n = WmsCar.VisitAtRow(r)
    If n = 0 Then
        Refuse("Поставьте курсор в строку визита, который нужно исправить.")
        Exit Sub
    End If
    d = WmsCar.CarRow(n)
    If d(CR_STATE) = CR_CANCELLED Then
        Refuse("Визит № " & n & " отменён — исправлять нечего.")
        Exit Sub
    End If
    v = WmsCar.VisitCells(n)
    bOpen = (d(CR_STATE) = CR_OPEN)
    f(0) = CStr(v(CC_PLATE))
    f(1) = CStr(v(CC_SUP))
    f(2) = WmsCar.HhMm(CDbl(d(CR_ARR)))
    If Not bOpen Then
        If Int(d(CR_DEP)) = Int(d(CR_ARR)) Then f(3) = WmsCar.HhMm(CDbl(d(CR_DEP))) Else f(3) = Format(CDate(d(CR_DEP)), "DD.MM.YYYY HH:MM")
    End If
    For r = 0 To 3
        g(r) = f(r)
    Next r
    sInfo = "Визит № " & n & " от " & WmsCar.DayText(CDbl(d(CR_ARR))) & IIf(bOpen, " — машина на территории: выезд запишет УЕХАЛ.", ".") _
        & " Время — ЧЧ:ММ; выезд на другой день — ДД.ММ.ГГГГ ЧЧ:ММ."
    If Not CarFields("WMS — Исправить визит", sInfo, Array("Марка / госномер", "Поставщик", "Приезд, ЧЧ:ММ", "Выезд, ЧЧ:ММ"), g, IIf(bOpen, 3, 4)) Then
        gUiLastMsg = "CANCEL:визит не изменён"
        Exit Sub
    End If
    ' an unchanged time is not sent (the window shows it without seconds)
    If Trim(g(2)) = f(2) Then g(2) = ""
    If Trim(g(3)) = f(3) Then g(3) = ""
    res = WmsCar.CarFix(n, g(0), g(1), g(2), g(3))
    AfterOperation("Исправить", res)
End Sub

Sub BtnCarCancel(Optional oEvent As Variant)
    Dim r As Long, n As Long, why As String, d As Variant, res As String
    r = CursorRow(why)
    If r < 0 Then
        Refuse(why)
        Exit Sub
    End If
    If r >= CAR_FIRST Then n = WmsCar.VisitAtRow(r)
    If n = 0 Then
        Refuse("Поставьте курсор в строку визита, который нужно отменить.")
        Exit Sub
    End If
    d = WmsCar.CarRow(n)
    If d(CR_STATE) = CR_CANCELLED Then
        Refuse("Визит № " & n & " уже отменён.")
        Exit Sub
    End If
    If Not UiConfirm("Отменить визит № " & n & " " & WmsCar.VisitText(n) & "? Используйте для ошибочной записи: визит останется в истории " _
        & "со статусом «" & CS_CANCEL & "» и не попадёт в показатели.") Then
        gUiLastMsg = "CANCEL:визит не отменён"
        Exit Sub
    End If
    res = WmsCar.CarCancel(n)
    AfterOperation("Отменить визит", res)
End Sub

' ================================================================ «Найти», «Сегодня», «Все» — filters of the table

' the conditions of the table (WMS_CARS) replaced and applied; the undo stack gets nothing, the book is not «modified» by it
Private Function CarFilter(f As Variant) As Boolean
    Dim db As Object, fd As Object, um As Object, locked As Boolean, wasMod As Boolean
    On Error GoTo EH
    If Not gDoc.DatabaseRanges.hasByName(CAR_DB) Then Exit Function
    wasMod = gDoc.isModified()
    db = gDoc.DatabaseRanges.getByName(CAR_DB)
    um = gDoc.getUndoManager()
    um.lock()
    locked = True
    fd = db.getFilterDescriptor()
    fd.ContainsHeader = True
    fd.setFilterFields2(f)
    db.refresh()
    um.unlock()
    locked = False
    If Not wasMod Then gDoc.setModified(False)
    CarFilter = True
    Exit Function
EH:
    On Error Resume Next
    If locked Then um.unlock()
End Function

' «Найти»: part of a plate, a brand or a supplier — the rows whose vehicle or supplier contain it (the plate also in the
' normalized form: «a123» finds «А 123 ВС»)
Sub BtnCarFind(Optional oEvent As Variant)
    Dim s As String, f(2) As New com.sun.star.sheet.TableFilterField2, k As String
    WmsInit()
    If gUiAuto <> 0 Then
        If Not gCarFindSet Then Exit Sub
        s = gCarFind
        gCarFindSet = False
    Else
        s = InputBox("Часть госномера, марка или поставщик:", "WMS — Найти машину")
    End If
    s = WmsCar.CleanText(s)
    If s = "" Then
        gUiLastMsg = "CANCEL:поиск не задан"
        Exit Sub
    End If
    k = WmsCar.SearchKey(s)
    f(0).Field = CC_PLATE
    f(0).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
    f(0).IsNumeric = False
    f(0).StringValue = s
    f(1).Connection = 1          ' FilterConnection OR (the name is a Basic keyword)
    f(1).Field = CC_SUP
    f(1).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
    f(1).IsNumeric = False
    f(1).StringValue = s
    f(2).Connection = 1          ' FilterConnection OR (the name is a Basic keyword)
    f(2).Field = CC_KEY
    f(2).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
    f(2).IsNumeric = False
    f(2).StringValue = IIf(k <> "", k, s)
    If CarFilter(f()) Then gUiLastMsg = "OK:показаны строки с «" & s & "»" Else Refuse("Не удалось применить поиск — воспользуйтесь автофильтром заголовка.")
End Sub

' «Сегодня»: the visits of today and every vehicle still on the territory
Sub BtnCarToday(Optional oEvent As Variant)
    Dim f(1) As New com.sun.star.sheet.TableFilterField2
    WmsInit()
    f(0).Field = CC_DATE
    f(0).Operator = com.sun.star.sheet.FilterOperator2.EQUAL
    f(0).IsNumeric = True
    f(0).NumericValue = Int(WmsCar.CarNow())
    f(1).Connection = 1          ' FilterConnection OR (the name is a Basic keyword)
    f(1).Field = CC_STATUS
    f(1).Operator = com.sun.star.sheet.FilterOperator2.EQUAL
    f(1).IsNumeric = False
    f(1).StringValue = CS_OPEN
    If CarFilter(f()) Then gUiLastMsg = "OK:показаны визиты за сегодня и машины на территории" Else Refuse("Не удалось применить отбор.")
End Sub

' «Все»: the filter of the table is removed and every row of the table shown again (a filter also hides the empty rows below
' the data, the next ПРИЕХАЛ writes there)
Sub BtnCarAll(Optional oEvent As Variant)
    Dim f() As New com.sun.star.sheet.TableFilterField2, um As Object, wasMod As Boolean
    WmsInit()
    If Not CarFilter(f()) Then
        Refuse("Не удалось снять отбор — воспользуйтесь автофильтром заголовка.")
        Exit Sub
    End If
    On Error Resume Next
    wasMod = gDoc.isModified()
    um = gDoc.getUndoManager()
    um.lock()
    WmsCar.CarsSheet().getCellRangeByPosition(0, CAR_FIRST, 0, MAX_SHEET_ROW).getRows().IsVisible = True
    um.unlock()
    If Not wasMod Then gDoc.setModified(False)
    gUiLastMsg = "OK:показаны все визиты"
End Sub

' ================================================================ windows

' a short list to choose from; pick: in — the line proposed, out — the line chosen
Private Function CarChoose(sTitle As String, sInfo As String, texts() As String, ByRef pick As Integer) As Boolean
    Dim dm As Object, dlg As Object, m As Object, sel(0) As Integer, h As Long
    gCarShown = Join(texts, Chr(10))
    If gUiAuto <> 0 Then
        If gCarPickSet Then
            gCarPickSet = False
            If gCarPick >= 0 And gCarPick <= UBound(texts) Then
                pick = gCarPick
                CarChoose = True
            End If
        End If
        Exit Function
    End If
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.PositionX = 8
    m.PositionY = 8
    m.Width = 284
    m.Height = 12
    m.Label = sInfo
    dm.insertByName("info", m)
    h = 14 + 11 * (UBound(texts) + 1)
    If h > 150 Then h = 150
    m = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    m.PositionX = 8
    m.PositionY = 24
    m.Width = 284
    m.Height = h
    m.Dropdown = False
    m.MultiSelection = False
    m.StringItemList = texts
    sel(0) = pick
    m.SelectedItems = sel
    dm.insertByName("lst", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 160
    m.PositionY = 30 + h
    m.Width = 68
    m.Height = 15
    m.Label = "Уехала"
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 232
    m.PositionY = 30 + h
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dm.Height = 52 + h
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        If dlg.getControl("lst").getSelectedItemPos() >= 0 Then
            pick = dlg.getControl("lst").getSelectedItemPos()
            CarChoose = True
        End If
    End If
    dlg.dispose()
End Function

' a window of nF edit fields (labels lbl, values v in and out)
Private Function CarFields(sTitle As String, sInfo As String, lbl As Variant, v() As String, nF As Integer) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, m As Object, y As Long
    gCarShown = Join(v, Chr(9))
    If gUiAuto <> 0 Then
        If gCarFixSet Then
            gCarFixSet = False
            If gCarFixIn(0) = CAR_CANCEL_DLG Then Exit Function
            For i = 0 To nF - 1
                If gCarFixIn(i) <> CAR_KEEP Then v(i) = gCarFixIn(i)
            Next i
            CarFields = True
        End If
        Exit Function
    End If
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    y = 8
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.PositionX = 8
    m.PositionY = y
    m.Width = 284
    m.Height = 26
    m.MultiLine = True
    m.Label = sInfo
    dm.insertByName("info", m)
    y = y + 30
    For i = 0 To nF - 1
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = y + 3
        m.Width = 90
        m.Height = 12
        m.Label = lbl(i)
        dm.insertByName("l" & i, m)
        m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        m.PositionX = 100
        m.PositionY = y
        m.Width = 192
        m.Height = 14
        m.Text = v(i)
        dm.insertByName("e" & i, m)
        y = y + 18
    Next i
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 160
    m.PositionY = y + 4
    m.Width = 68
    m.Height = 15
    m.Label = "Исправить"
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 232
    m.PositionY = y + 4
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dm.Height = y + 26
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        For i = 0 To nF - 1
            v(i) = dlg.getControl("e" & i).getText()
        Next i
        CarFields = True
    End If
    dlg.dispose()
End Function

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestCarPick(i As Integer) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestCarPick = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gCarPick = i
    gCarPickSet = True
    TestCarPick = "OK"
End Function

Function TestCarFind(s As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestCarFind = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gCarFind = s
    gCarFindSet = True
    TestCarFind = "OK"
End Function

' the answers of the next window of «Исправить»: "<как предложено>" keeps a value, "<отмена>" as the first one closes the window
Function TestCarFixInput(sPlate As String, sSup As String, sArr As String, sDep As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestCarFixInput = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gCarFixIn(0) = sPlate
    gCarFixIn(1) = sSup
    gCarFixIn(2) = sArr
    gCarFixIn(3) = sDep
    gCarFixSet = True
    TestCarFixInput = "OK"
End Function

Function TestCarShown() As String
    TestCarShown = gCarShown
End Function
