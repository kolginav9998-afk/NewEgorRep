' WmsSpecialUi — кнопки листа «Иной приход» (задание Core Phase 5): «Проверить», «Провести», «Исправить», «Удалить»,
' «Разобрать», «Очистить». Как и на других листах, успешная работа окон не открывает: результат виден в строке (A № строки,
' C № поступления, N ЕИ, O P остаток до и после, Q статус, R контроль). Окно — для вопроса «одно поступление?» при
' проведении нескольких строк, для подтверждения перемещения ЕИ детали и сторно, для исправления и разбора и когда
' действие невозможно.
Option Explicit

' test seam (TEST books only): the answers of the next window of «Исправить» / «Разобрать» (one per field; SPC_KEEP: the
' value the window proposed stays, as when the user does not touch the field)
Global gSpcInSet As Boolean
Global gSpcIn() As String
' what the last window proposed and what it returned (tab-separated), the text of the window that ended the last block
Global gSpcShown As String
Global gSpcBack As String
Global gSpcBlockMsg As String
Private Const SPC_KEEP = "<как предложено>"

' ================================================================ the row under the cursor

' blocks of selected rows on «Иной приход»: Array(r0, r1) each; False with a message when the cursor is not on the sheet
Function ActiveSpecialBlocks(ByRef blocks As Variant, ByRef why As String) As Boolean
    Dim sel As Object, addrs As Variant, ctl As Object, i As Long, n As Long, out() As Variant, r0 As Long
    WmsInit()
    On Error GoTo EH
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().getName() <> SH_SPECIAL Then
        why = "Перейдите на лист «" & SH_SPECIAL & "» и поставьте курсор в строку прихода."
        Exit Function
    End If
    sel = gDoc.getCurrentSelection()
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        addrs = sel.getRangeAddresses()
    Else
        addrs = Array(sel.getRangeAddress())
    End If
    ReDim out(UBound(addrs))
    For i = 0 To UBound(addrs)
        r0 = addrs(i).StartRow
        If r0 < 1 Then r0 = 1
        If addrs(i).EndRow >= r0 Then
            out(n) = Array(r0, addrs(i).EndRow)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        why = "Поставьте курсор в строку прихода, а не в заголовок."
        Exit Function
    End If
    ReDim Preserve out(n - 1)
    blocks = out
    ActiveSpecialBlocks = True
    Exit Function
EH:
    why = "Поставьте курсор в строку прихода на листе «" & SH_SPECIAL & "»."
End Function

' exactly one row under the cursor
Function ActiveSpecialRow(ByRef r As Long, ByRef why As String) As Boolean
    Dim blocks As Variant
    If Not ActiveSpecialBlocks(blocks, why) Then Exit Function
    If UBound(blocks) > 0 Or blocks(0)(0) <> blocks(0)(1) Then
        why = "Выделите одну строку прихода (сейчас выделено несколько строк)."
        Exit Function
    End If
    r = blocks(0)(0)
    ActiveSpecialRow = True
End Function

' an action that is not possible for this row: the window explains it; the test seam records the refusal
Private Sub Refuse(s As String)
    UiMessage(s)
    gUiLastMsg = "ERR:" & s
End Sub

Private Function Payload(res As String) As String
    Payload = Mid(res, InStr(res, ":") + 1)
End Function

Private Sub AfterOperation(sAction As String, res As String)
    gUiLastMsg = res
    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Then UiRefresh(sAction & ": " & res)
End Sub

' the refusal for a row that is not a posted line
Private Function NotPostedText(kind As String) As String
    Select Case kind
    Case "OPEN"
        NotPostedText = "Строка не проведена — исправьте значения прямо в строке и нажмите «Провести»."
    Case "DELETED"
        NotPostedText = "Строка № " & gXn & " удалена (сторно) — она хранит историю."
    Case "EMPTY"
        NotPostedText = "Поставьте курсор в строку проведённого прихода."
    Case Else
        NotPostedText = "Строка — КОПИЯ, её можно только очистить кнопкой «Очистить»."
    End Select
End Function

' ================================================================ «Проверить»

Sub BtnSpcCheck(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, r As Long, nErr As Long, n As Long, sh As Object, b As Long, r0 As Long
    If Not ActiveSpecialBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsSpecial.SpecialCheckRow(r0)
        If WmsSpecial.SpecialRowKind(r0) <> "OPEN" Then
            If Left(res, 3) = "OK:" Then UiInfo(Payload(res)) Else UiMessage(Payload(res))
        End If
        gUiLastMsg = res
        Exit Sub
    End If
    sh = WmsSpecial.SpecialSheet()
    For b = 0 To UBound(blocks)
        For r = blocks(b)(0) To blocks(b)(1)
            If sh.getRows().getByIndex(r).IsVisible Then
                If WmsSpecial.SpecialRowKind(r) = "OPEN" Then
                    n = n + 1
                    res = WmsSpecial.SpecialCheckRow(r)
                    If Left(res, 3) <> "OK:" Then nErr = nErr + 1
                End If
            End If
        Next r
    Next b
    If nErr > 0 Then UiMessage("Проверено строк прихода: " & n & ", с ошибками: " & nErr & " — причины в колонке «Контроль».")
    gUiLastMsg = "OK:проверено строк " & n & ", с ошибками " & nErr
End Sub

' ================================================================ «Провести»

' The row under the cursor, or every visible row of the selected block(s) with a quantity F (D-059: each row its own
' operation; a problem of one row does not stop the others; a system error of WMS stops the whole block). Several rows:
' WMS asks whether they are one acceptance — then the lines of one source share one event № (OFF/PROD/DET/OLD/OTH).
' A part whose EI would move to another place is posted only row by row, after a confirmation.
Sub BtnSpcPost(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, a As Variant, b As Long, r0 As Long, nOk As Long, nErr As Long, nSkip As Long
    Dim first As String, sysErr As String, stopRow As Long, parts As Variant, msg As String, mv As String, bOk As Boolean, n As Long
    Dim nSrc As Integer, ns As Integer, ans As Integer, bShared As Boolean, i As Integer, evs As String
    If Not ActiveSpecialBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        mv = WmsSpecial.SpecialPlaceMove(r0)
        If mv <> "" Then
            If Not UiConfirm("Место ЕИ детали изменится: " & mv & Chr(10) & Chr(10) & "Провести приход и сделать это место текущим местом ЕИ?") Then
                gUiLastMsg = "SKIP:перемещение ЕИ не подтверждено — строка не проведена"
                Exit Sub
            End If
            bOk = True
        End If
        res = WmsSpecial.SpecialPostRow(r0, "", bOk)
        AfterOperation("«Провести» (иной приход)", res)
        If Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then UiMessage(Payload(res))
        gUiLastMsg = res
        Exit Sub
    End If
    For b = 0 To UBound(blocks)
        n = n + WmsSpecial.PostableRows(blocks(b)(0), blocks(b)(1), ns)
        If ns > nSrc Then nSrc = ns
    Next b
    If n = 0 Then
        UiMessage("В выделенных строках нечего проводить: нет непроведённых строк с количеством (F).")
        gUiLastMsg = "SKIP:нечего проводить"
        Exit Sub
    End If
    If n >= 2 Then
        ans = UiAsk3("Провести выделенные строки (" & n & ") как одно поступление?" & Chr(10) & Chr(10) _
            & "«Да» — один № поступления (OFF / PROD / DET / OLD / OTH) на все строки одного типа прихода" _
            & IIf(nSrc > 1, " (в выделении разные типы — у каждого типа свой №)", "") & ";" & Chr(10) _
            & "«Нет» — у каждой строки свой № поступления;" & Chr(10) & "«Отмена» — ничего не проводить.")
        If ans = 2 Then
            gUiLastMsg = "SKIP:проведение отменено"
            Exit Sub
        End If
        bShared = (ans = 6)
    End If
    For b = 0 To UBound(blocks)
        parts = Split(WmsSpecial.SpecialPostRange(blocks(b)(0), blocks(b)(1), bShared, b > 0), Chr(10))
        a = Split(parts(0), ";")
        nOk = nOk + Val(Mid(a(0), 4))
        nErr = nErr + Val(Mid(a(1), 5))
        nSkip = nSkip + Val(Mid(a(2), 6))
        If first = "" Then first = parts(1)
        stopRow = Val(Mid(a(3), 6))
        If stopRow > 0 Then
            sysErr = parts(2)
            Exit For
        End If
    Next b
    res = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";STOP=" & stopRow & Chr(10) & first & Chr(10) & sysErr
    msg = "«Провести» для выделенных строк иного прихода: проведено " & nOk & ", пропущено " & nSkip _
        & " (нет количества F, уже проведены, удалены или копии), отклонено с ошибкой " & nErr _
        & IIf(nErr > 0, " — причины в колонке «Контроль»", "") & "."
    If bShared Then
        For i = 0 To 4
            If gXev(i) <> "" Then evs = evs & IIf(evs <> "", ", ", "") & gXev(i)
        Next i
        If evs <> "" Then msg = msg & Chr(10) & "Одно поступление: " & evs & "."
    End If
    If first <> "" Then msg = msg & Chr(10) & "Первая ошибка: " & first
    gSpcBlockMsg = msg
    If stopRow > 0 Then
        msg = msg & Chr(10) & Chr(10) & "ОБРАБОТКА ОСТАНОВЛЕНА — системная ошибка WMS, " & SysReason(sysErr) _
            & Chr(10) & "Строки после неё не обрабатывались. Проверьте лист «" & SH_MAIN & "»."
        gSpcBlockMsg = msg
        UiMessage(msg)
        If InStr(sysErr, "BLOCKED:") > 0 Or InStr(sysErr, "ERR-CRITICAL:") > 0 Then UiRefresh("«Провести» (иной приход): " & sysErr)
    ElseIf nErr > 0 Then
        UiMessage(msg)
    Else
        UiInfo(msg)
    End If
    gUiLastMsg = res
End Sub

' the system error of a block ("строка N: <result>") for the window: the reason without the technical prefix
Private Function SysReason(sysErr As String) As String
    Dim p As Long, rest As String
    p = InStr(sysErr, ": ")
    If p = 0 Then
        SysReason = sysErr
        Exit Function
    End If
    rest = Mid(sysErr, p + 2)
    If Left(rest, 3) = "OK:" And InStr(rest, "(") > 0 Then
        SysReason = Left(sysErr, p + 1) & "приход проведён " & Mid(rest, InStr(rest, "("))
    Else
        SysReason = Left(sysErr, p + 1) & Payload(rest)
    End If
End Function

' ================================================================ «Исправить» — SP_FIX

Sub BtnSpcFix(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, sh As Object, v(9) As String, lbl(9) As String, sInfo As String
    Dim isAdd As Boolean, empty0() As String
    If Not ActiveSpecialRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsSpecial.SpecialRowKind(r)
    If kind <> "POSTED" Then
        Refuse(NotPostedText(kind))
        Exit Sub
    End If
    sh = WmsSpecial.SpecialSheet()
    isAdd = (Right(sh.getCellByPosition(XC_STATUS, r).getString(), Len(XK_ADD)) = XK_ADD)
    lbl(0) = "Количество (F)"
    lbl(1) = "Дата поступления (H), дд.мм.гггг"
    lbl(2) = "Место хранения (I)"
    lbl(3) = "Категория (J)"
    lbl(4) = "Документ (L)"
    lbl(5) = "Кто передал (K)"
    lbl(6) = "Старая маркировка (M)"
    lbl(7) = "Наименование (D)"
    lbl(8) = "Артикул (E)"
    lbl(9) = "Единица измерения (G)"
    v(0) = WmsIssue.QtyText(sh.getCellByPosition(XC_QTY, r).getValue())
    v(1) = Format(sh.getCellByPosition(XC_DATE, r).getValue(), "DD.MM.YYYY")
    v(2) = sh.getCellByPosition(XC_PLACE, r).getString()
    v(3) = sh.getCellByPosition(XC_CAT, r).getString()
    v(4) = sh.getCellByPosition(XC_DOC, r).getString()
    v(5) = sh.getCellByPosition(XC_WHO, r).getString()
    v(6) = sh.getCellByPosition(XC_MARK, r).getString()
    v(7) = sh.getCellByPosition(XC_NAME, r).getString()
    v(8) = sh.getCellByPosition(XC_ART, r).getString()
    v(9) = sh.getCellByPosition(XC_UNIT, r).getString()
    sInfo = "Строка № " & gXn & ", поступление " & sh.getCellByPosition(XC_EVENT, r).getString() & ", " & sh.getCellByPosition(XC_EI, r).getString() _
        & IIf(isAdd, ". Пополнение ЕИ детали: наименование, артикул, единица и категория — от карточки ЕИ, здесь не меняются.", _
        ". Строка создала ЕИ: наименование, артикул, категория и единица — это карточка ЕИ. Тип прихода, ЕИ и № поступления не меняются.")
    If Not SpcDialog("Исправить строку № " & gXn & " (тот же ЕИ)", sInfo, "Исправить", lbl, v, -1, empty0) Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        Exit Sub
    End If
    res = WmsSpecial.SpecialFixRow(r, v(0), v(1), v(2), v(3), v(4), v(5), v(6), v(7), v(8), v(9))
    AfterOperation("«Исправить» (иной приход)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Исправление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

' ================================================================ «Удалить» — SP_DEL: the check first, the confirmation only for an allowed storno

Sub BtnSpcDelete(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, sh As Object
    If Not ActiveSpecialRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsSpecial.SpecialRowKind(r)
    Select Case kind
    Case "POSTED"
    Case "DELETED"
        UiMessage("Строка № " & gXn & " уже удалена (сторно).")
        gUiLastMsg = "SKIP:строка уже удалена (сторно)"
        Exit Sub
    Case "OPEN"
        Refuse("Строка не проведена — чтобы убрать ввод, нажмите «Очистить».")
        Exit Sub
    Case "EMPTY"
        Refuse("Поставьте курсор в строку проведённого прихода.")
        Exit Sub
    Case Else
        Refuse("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    ' every check of the storno first (dependent issues and returns, the balance): a storno that is not allowed is refused
    ' at once, the confirmation is asked only for a storno that can be done
    res = WmsSpecial.SpecialDeleteCheck(r)
    If Left(res, 3) <> "OK:" Then
        AfterOperation("«Удалить» (иной приход)", res)
        If Left(res, 5) = "SKIP:" Then UiMessage(Payload(res)) Else UiMessage("Удалить нельзя: " & Payload(res))
        gUiLastMsg = res
        Exit Sub
    End If
    sh = WmsSpecial.SpecialSheet()
    If Not UiConfirm("Удалить строку прихода № " & gXn & "?" & Chr(10) & Chr(10) & sh.getCellByPosition(XC_EI, r).getString() & " — " _
        & sh.getCellByPosition(XC_NAME, r).getString() & ", " & sh.getCellByPosition(XC_QTY, r).getString() & " " & sh.getCellByPosition(XC_UNIT, r).getString() _
        & " (" & sh.getCellByPosition(XC_EVENT, r).getString() & ")" & Chr(10) & Chr(10) _
        & "Остаток ЕИ уменьшится на это количество. Строка останется в истории с отметкой «" & ST_DELETED & "», № не будет использован повторно.") Then
        gUiLastMsg = "SKIP:удаление не подтверждено"
        Exit Sub
    End If
    res = WmsSpecial.SpecialDeleteRow(r)
    AfterOperation("«Удалить» (иной приход)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Удаление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

' ================================================================ «Разобрать» — SP_IDENTIFY

Sub BtnSpcIdentify(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, sh As Object, v(3) As String, lbl(3) As String, items(4) As String, st As String
    If Not ActiveSpecialRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsSpecial.SpecialRowKind(r)
    If kind <> "POSTED" Then
        Refuse(NotPostedText(kind))
        Exit Sub
    End If
    sh = WmsSpecial.SpecialSheet()
    st = sh.getCellByPosition(XC_STATUS, r).getString()
    If Right(st, Len(XK_REVIEW)) <> XK_REVIEW Then
        Refuse("«Разобрать» — только для строки иного прихода, которая «требует разбора» (эта строка: «" & st & "»).")
        Exit Sub
    End If
    lbl(0) = "Что это за приход"
    lbl(1) = "Наименование"
    lbl(2) = "Артикул (для детали обязателен)"
    lbl(3) = "Категория"
    items(0) = STYPE_SUPPLIER
    items(1) = WmsSpecial.SrcName(0)
    items(2) = WmsSpecial.SrcName(1)
    items(3) = WmsSpecial.SrcName(2)
    items(4) = WmsSpecial.SrcName(3)
    ' no type is chosen for the user (D-081): «OK» without a choice is refused, nothing changes
    v(0) = ""
    v(1) = sh.getCellByPosition(XC_NAME, r).getString()
    v(2) = sh.getCellByPosition(XC_ART, r).getString()
    v(3) = sh.getCellByPosition(XC_CAT, r).getString()
    If Not SpcDialog("Разобрать иной приход, строка № " & gXn & " (" & sh.getCellByPosition(XC_EI, r).getString() & " сохраняется)", _
        "Та же физическая партия: ЕИ не меняется, уточняется карточка — источник, наименование, артикул, категория.", "Разобрать", lbl, v, 0, items) Then
        gUiLastMsg = "SKIP:окно закрыто без разбора"
        Exit Sub
    End If
    res = WmsSpecial.SpecialIdentifyRow(r, v(0), v(1), v(2), v(3))
    AfterOperation("«Разобрать» (иной приход)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Разбор не выполнен: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

Sub BtnSpcClear(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String
    If Not ActiveSpecialRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    res = WmsSpecial.SpecialClearRow(r)
    If Left(res, 3) <> "OK:" Then UiMessage(Payload(res))
    gUiLastMsg = res
End Sub

' ================================================================ the window: labelled fields, one of them may be a list

' labels(i), v(i): the fields (in: the proposed values, out: what «OK» returned); listField: the index of a field shown as
' a dropdown list of listItems (-1: none)
Private Function SpcDialog(sTitle As String, sInfo As String, sOk As String, labels() As String, v() As String, listField As Integer, _
    listItems() As String) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, m As Object, y As Long, sel(0) As Integer, k As Integer
    gSpcShown = Join(v, Chr(9))
    gSpcBack = ""
    If gUiAuto <> 0 Then
        If gSpcInSet Then
            For i = 0 To UBound(v)
                If i <= UBound(gSpcIn) Then
                    If gSpcIn(i) <> SPC_KEEP Then v(i) = gSpcIn(i)
                End If
            Next i
            gSpcInSet = False
            gSpcBack = Join(v, Chr(9))
            SpcDialog = True
        End If
        Exit Function
    End If
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 320
    y = 8
    If sInfo <> "" Then
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = y
        m.Width = 304
        m.Height = 26
        m.MultiLine = True
        m.Label = sInfo
        dm.insertByName("info", m)
        y = y + 30
    End If
    For i = 0 To UBound(labels)
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = y + 3
        m.Width = 120
        m.Height = 12
        m.Label = labels(i)
        dm.insertByName("l" & i, m)
        If i = listField Then
            m = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
            m.Dropdown = True
            m.MultiSelection = False
            m.StringItemList = listItems
            ' an item is selected only when the field already holds it: nothing is chosen for the user
            For k = 0 To UBound(listItems)
                If listItems(k) = v(i) Then
                    sel(0) = k
                    m.SelectedItems = sel
                End If
            Next k
        Else
            m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
            m.Text = v(i)
        End If
        m.PositionX = 130
        m.PositionY = y
        m.Width = 182
        m.Height = 14
        dm.insertByName("e" & i, m)
        y = y + 18
    Next i
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 180
    m.PositionY = y + 4
    m.Width = 68
    m.Height = 15
    m.Label = sOk
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 252
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
        For i = 0 To UBound(labels)
            If i = listField Then
                v(i) = dlg.getControl("e" & i).getSelectedItem()
            Else
                v(i) = dlg.getControl("e" & i).getText()
            End If
        Next i
        gSpcBack = Join(v, Chr(9))
        SpcDialog = True
    End If
    dlg.dispose()
End Function

' ================================================================ test seam (inert unless _SYS MODE = TEST)

' the answers of the next window of «Исправить» / «Разобрать»: tab-separated, one per field (SPC_KEEP — as proposed)
Function TestSpcInput(sValues As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSpcInput = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gSpcIn = Split(sValues, Chr(9))
    gSpcInSet = True
    TestSpcInput = "OK"
End Function

' the last window: what it proposed and what «OK» returned (tab-separated; the second part is empty when the window was
' closed without «OK»)
Function TestSpcDialog() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSpcDialog = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestSpcDialog = gSpcShown & Chr(10) & gSpcBack
End Function

' the text of the window that ended the last block of «Провести»
Function TestSpcBlockMessage() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSpcBlockMessage = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestSpcBlockMessage = gSpcBlockMsg
End Function
