' WmsOrdersUi — кнопки листа «Заказы» (задание Core Phase 3): «Проверить», «Провести приход», «Ещё поступление»,
' «Исправить», «Удалить», «Отменить заказ», «Отменить остаток», «Обновить статусы», «Очистить». Как и на «Выдачах»,
' успешная работа окон не открывает: результат виден в строке (V ЕИ, W статус, X наличие, Y контроль). Окно — для
' подтверждения, для ввода поступления или исправления и когда действие невозможно.
Option Explicit

' test seam (TEST books only): the values the next receipt dialog returns: F, G, C, O, N, U, J (RCV_KEEP — the value the
' window proposed stays, as when the user does not touch the field)
Global gRcvInSet As Boolean
Global gRcvIn(6) As String
' what the last receipt dialog proposed and what it returned (the test seam reads them; the GUI test checks the real window)
Global gRcvShown As String
Global gRcvBack As String
' the text of the window that ended the last block of «Провести приход» (the test seam reads it)
Global gRcvBlockMsg As String
Private Const RCV_KEEP = "<как предложено>"

' ================================================================ the row under the cursor

' blocks of selected rows on «Заказы»: Array(r0, r1) each (a selection across rows hidden by a filter consists of several
' blocks of visible rows); False with a message when the cursor is not on the sheet
Function ActiveOrderBlocks(ByRef blocks As Variant, ByRef why As String) As Boolean
    Dim sel As Object, addrs As Variant, ctl As Object, i As Long, n As Long, out() As Variant, r0 As Long
    WmsInit()
    On Error GoTo EH
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().getName() <> SH_ORDERS Then
        why = "Перейдите на лист «" & SH_ORDERS & "» и поставьте курсор в строку заказа."
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
        why = "Поставьте курсор в строку заказа, а не в заголовок."
        Exit Function
    End If
    ReDim Preserve out(n - 1)
    blocks = out
    ActiveOrderBlocks = True
    Exit Function
EH:
    why = "Поставьте курсор в строку заказа на листе «" & SH_ORDERS & "»."
End Function

' exactly one row under the cursor
Function ActiveOrderRow(ByRef r As Long, ByRef why As String) As Boolean
    Dim blocks As Variant
    If Not ActiveOrderBlocks(blocks, why) Then Exit Function
    If UBound(blocks) > 0 Or blocks(0)(0) <> blocks(0)(1) Then
        why = "Выделите одну строку заказа (сейчас выделено несколько строк)."
        Exit Function
    End If
    r = blocks(0)(0)
    ActiveOrderRow = True
End Function

' an action that is not possible for this row: the window explains it; the test seam records the refusal
Private Sub Refuse(s As String)
    UiMessage(s)
    gUiLastMsg = "ERR:" & s
End Sub

Private Function Payload(res As String) As String
    Payload = Mid(res, InStr(res, ":") + 1)
End Function

' an information window (the result of an explicit check or refresh); recorded only in the test mode
Sub UiInfo(s As String)
    gUiLastMsg = s
    If gUiAuto <> 0 Then Exit Sub
    MsgBox s, 64, "WMS"
End Sub

Private Sub AfterOperation(sAction As String, res As String)
    gUiLastMsg = res
    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Then UiRefresh(sAction & ": " & res)
End Sub

' ================================================================ buttons

' «Проверить»: open rows get the result in «Контроль»; for a row with a receipt the check result is shown
Sub BtnRcvCheck(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, r As Long, nErr As Long, n As Long, sh As Object, b As Long, r0 As Long
    If Not ActiveOrderBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsReceipt.CheckRow(r0)
        If WmsOrders.OrderRowKind(r0) <> "OPEN" Then
            If Left(res, 3) = "OK:" Then UiInfo(Payload(res)) Else UiMessage(Payload(res))
        End If
        gUiLastMsg = res
        Exit Sub
    End If
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    For b = 0 To UBound(blocks)
        For r = blocks(b)(0) To blocks(b)(1)
            If sh.getRows().getByIndex(r).IsVisible Then
                If WmsOrders.OrderRowKind(r) = "OPEN" Then
                    n = n + 1
                    res = WmsReceipt.CheckRow(r)
                    If Left(res, 3) <> "OK:" Then nErr = nErr + 1
                End If
            End If
        Next r
    Next b
    If nErr > 0 Then UiMessage("Проверено строк заказа: " & n & ", с ошибками: " & nErr & " — причины в колонке «Контроль».")
    gUiLastMsg = "OK:проверено строк " & n & ", с ошибками " & nErr
End Sub

' «Провести приход»: the row under the cursor, or every visible row of the selected block(s) with a fact F (D-059: each row
' its own operation; a problem of one row does not stop the others; a system error of WMS stops the whole block). A block
' ends with a window showing how many rows were posted, skipped and rejected.
Sub BtnRcvPost(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, a As Variant, b As Long, r0 As Long, nOk As Long, nErr As Long, nSkip As Long
    Dim first As String, sysErr As String, stopRow As Long, lines As Variant, msg As String
    If Not ActiveOrderBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsReceipt.ReceiptPostRow(r0)
        AfterOperation("«Провести приход»", res)
        If Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then UiMessage(Payload(res))
        gUiLastMsg = res
        Exit Sub
    End If
    For b = 0 To UBound(blocks)
        lines = Split(WmsReceipt.ReceiptPostRange(blocks(b)(0), blocks(b)(1)), Chr(10))
        a = Split(lines(0), ";")
        nOk = nOk + Val(Mid(a(0), 4))
        nErr = nErr + Val(Mid(a(1), 5))
        nSkip = nSkip + Val(Mid(a(2), 6))
        If first = "" Then first = lines(1)
        stopRow = Val(Mid(a(3), 6))
        If stopRow > 0 Then
            sysErr = lines(2)
            Exit For
        End If
    Next b
    res = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";STOP=" & stopRow & Chr(10) & first & Chr(10) & sysErr
    msg = "«Провести приход» для выделенных строк: проведено " & nOk & ", пропущено " & nSkip _
        & " (нет факта F, уже проведены, отменены или копии), отклонено с ошибкой " & nErr _
        & IIf(nErr > 0, " — причины в колонке «Контроль»", "") & "."
    If first <> "" Then msg = msg & Chr(10) & "Первая ошибка: " & first
    gRcvBlockMsg = msg
    If stopRow > 0 Then
        msg = msg & Chr(10) & Chr(10) & "ОБРАБОТКА ОСТАНОВЛЕНА — системная ошибка WMS, " & SysReason(sysErr) _
            & Chr(10) & "Строки после неё не обрабатывались. Проверьте лист «" & SH_MAIN & "»."
        gRcvBlockMsg = msg
        UiMessage(msg)
        If InStr(sysErr, "BLOCKED:") > 0 Or InStr(sysErr, "ERR-CRITICAL:") > 0 Then UiRefresh("«Провести приход»: " & sysErr)
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

' «Ещё поступление»: the next receipt of the position of the row under the cursor, in a new row
Sub BtnRcvMore(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, v(6) As String, sh As Object, a As Variant, nr As Long
    If Not ActiveOrderRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "RECEIVED", "STORNO"
    Case "OPEN"
        Refuse("По этой позиции ещё не было прихода — заполните F, N, U в этой строке и нажмите «Провести приход».")
        Exit Sub
    Case "CANCELLED"
        Refuse("Позиция отменена — поступление по ней не проводится. Оформите новую строку заказа.")
        Exit Sub
    Case "EMPTY"
        Refuse("Поставьте курсор в строку позиции заказа, по которой пришла следующая поставка.")
        Exit Sub
    Case Else
        Refuse("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    ' the quantity, the document and its date are new for every delivery; the receipt date N is proposed as today (D-049:
    ' a visible suggestion the user confirms or changes, never a date fixed without the user); place and price are
    ' proposed from the row
    v(4) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    v(5) = sh.getCellByPosition(OC_PLACE, r).getString()
    If sh.getCellByPosition(OC_PRICE, r).getType() = com.sun.star.table.CellContentType.VALUE Then v(6) = WmsIssue.QtyText(sh.getCellByPosition(OC_PRICE, r).getValue())
    If Not ReceiptDialog("Ещё поступление: " & sh.getCellByPosition(OC_NAME, r).getString() & " (заказ " & sh.getCellByPosition(OC_ORDER, r).getString() & ")", _
        "Провести", v) Then
        gUiLastMsg = "SKIP:окно закрыто без проведения"
        Exit Sub
    End If
    res = WmsReceipt.ReceiptAddRow(r, v(0), v(1), v(2), v(3), v(4), v(5), v(6))
    AfterOperation("«Ещё поступление»", res)
    If Left(res, 3) <> "OK:" Then
        UiMessage("Поступление не проведено: " & Payload(res))
        gUiLastMsg = res
        Exit Sub
    End If
    ' the new row is shown to the user (its new EI in V); a row inserted inside a filtered range may be hidden by the filter
    ' (M6 §18: the delivery row stays in the block of its order) — it is shown, the filter itself stays
    a = Split(res, "|")
    If UBound(a) >= 1 Then
        nr = CLng(Val(Mid(a(1), 8)))
        On Error Resume Next
        If Not sh.getRows().getByIndex(nr - 1).IsVisible Then
            gDoc.getUndoManager().lock()
            sh.getRows().getByIndex(nr - 1).IsVisible = True
            gDoc.getUndoManager().unlock()
        End If
        gDoc.getCurrentController().select(sh.getCellByPosition(OC_EI, nr - 1))
    End If
End Sub

Sub BtnRcvFix(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, v(6) As String, sh As Object
    If Not ActiveOrderRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "RECEIVED"
    Case "OPEN"
        Refuse("Строка не проведена — исправьте значения прямо в строке и нажмите «Провести приход».")
        Exit Sub
    Case "STORNO"
        Refuse("Приход удалён (сторно) — исправлять нечего.")
        Exit Sub
    Case "CANCELLED"
        Refuse("Позиция отменена — прихода в этой строке нет.")
        Exit Sub
    Case "EMPTY"
        Refuse("Поставьте курсор в строку проведённого прихода.")
        Exit Sub
    Case Else
        Refuse("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    v(0) = WmsIssue.QtyText(sh.getCellByPosition(OC_FACT, r).getValue())
    If sh.getCellByPosition(OC_DOCQTY, r).getType() = com.sun.star.table.CellContentType.VALUE Then v(1) = WmsIssue.QtyText(sh.getCellByPosition(OC_DOCQTY, r).getValue())
    v(2) = sh.getCellByPosition(OC_DOC, r).getString()
    If sh.getCellByPosition(OC_DDATE, r).getType() = com.sun.star.table.CellContentType.VALUE Then v(3) = Format(sh.getCellByPosition(OC_DDATE, r).getValue(), "DD.MM.YYYY")
    v(4) = Format(sh.getCellByPosition(OC_RDATE, r).getValue(), "DD.MM.YYYY")
    v(5) = sh.getCellByPosition(OC_PLACE, r).getString()
    If sh.getCellByPosition(OC_PRICE, r).getType() = com.sun.star.table.CellContentType.VALUE Then v(6) = WmsIssue.QtyText(sh.getCellByPosition(OC_PRICE, r).getValue())
    If Not ReceiptDialog("Исправить приход " & gKcanon & " (тот же ЕИ)", "Исправить", v) Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        Exit Sub
    End If
    res = WmsReceipt.ReceiptFixRow(r, v(0), v(1), v(2), v(3), v(4), v(5), v(6))
    AfterOperation("«Исправить»", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Исправление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

Sub BtnRcvDelete(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, sh As Object
    If Not ActiveOrderRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "RECEIVED"
    Case "STORNO"
        ' a repeated click after the storno: nothing to do
        UiMessage("Приход уже удалён (сторно).")
        gUiLastMsg = "SKIP:приход уже удалён (сторно)"
        Exit Sub
    Case "OPEN"
        Refuse("Строка не проведена. Чтобы убрать ввод — «Очистить»; чтобы отменить заказ — «Отменить заказ».")
        Exit Sub
    Case "CANCELLED"
        Refuse("Позиция отменена — прихода в этой строке нет.")
        Exit Sub
    Case "EMPTY"
        Refuse("Поставьте курсор в строку проведённого прихода.")
        Exit Sub
    Case Else
        Refuse("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    ' every check of the storno first (D-069: live issues / returns of the EI, the balance, the WMS state): a storno that is
    ' not allowed is refused at once, the confirmation is asked only for a storno that can be done
    res = WmsReceipt.ReceiptDeleteCheck(r)
    If Left(res, 3) <> "OK:" Then
        AfterOperation("«Удалить»", res)
        If Left(res, 5) = "SKIP:" Then
            UiMessage(Payload(res))
        Else
            UiMessage("Удалить приход нельзя: " & Payload(res))
        End If
        gUiLastMsg = res
        Exit Sub
    End If
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    If Not UiConfirm("Удалить приход " & gKcanon & "?" & Chr(10) & Chr(10) & sh.getCellByPosition(OC_NAME, r).getString() & ", " _
        & sh.getCellByPosition(OC_FACT, r).getString() & " " & sh.getCellByPosition(OC_UNIT, r).getString() & Chr(10) & Chr(10) _
        & "Остаток этого ЕИ станет 0. Строка останется в истории с отметкой «Приход удалён (сторно)», ЕИ не будет использован повторно.") Then Exit Sub
    res = WmsReceipt.ReceiptDeleteRow(r)
    AfterOperation("«Удалить»", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Удаление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

Sub BtnOrdCancel(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, sh As Object
    If Not ActiveOrderRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    If Not UiConfirm("Отменить заказ?" & Chr(10) & Chr(10) & "Заказ " & sh.getCellByPosition(OC_ORDER, r).getString() & ": " _
        & sh.getCellByPosition(OC_NAME, r).getString() & Chr(10) & Chr(10) & "Позиция получит статус «Отменено», строка останется в истории. " _
        & "ЕИ не создаётся.") Then Exit Sub
    res = WmsReceipt.OrderCancelRow(r)
    AfterOperation("«Отменить заказ»", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Заказ не отменён: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

Sub BtnOrdCancelRest(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, sh As Object
    If Not ActiveOrderRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    If Not UiConfirm("Отменить остаток позиции?" & Chr(10) & Chr(10) & "Заказ " & sh.getCellByPosition(OC_ORDER, r).getString() & ": " _
        & sh.getCellByPosition(OC_NAME, r).getString() & Chr(10) & Chr(10) & "Полученное количество и заказанное (H) не меняются, " _
        & "неполученный остаток больше не ожидается.") Then Exit Sub
    res = WmsReceipt.OrderCancelRestRow(r)
    AfterOperation("«Отменить остаток»", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Остаток не отменён: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

' «Обновить статусы»: order rows without a status get it, the date statuses of an earlier core become the status of their
' data; the note counts the expected positions whose expected date has passed (D-089: the status never depends on the date)
Sub BtnOrdRefresh(Optional oEvent As Variant)
    Dim res As String, um As Object
    WmsInit()
    If gState = "" Then WmsStartup()
    On Error GoTo EH
    UndoBegin()
    res = WmsOrders.RefreshStatuses()
    UndoEnd()
    UiInfo(res)
    Exit Sub
EH:
    res = "статусы не обновлены: " & Error$
    UndoEnd()
    UiMessage(res)
End Sub

Sub BtnRcvClear(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String
    If Not ActiveOrderRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    res = WmsOrders.ClearOrderRow(r)
    If Left(res, 3) <> "OK:" Then UiMessage(Payload(res))
    gUiLastMsg = res
End Sub

' ================================================================ the receipt dialog (F, G, C, O, N, U, J)

Private Function ReceiptDialog(sTitle As String, sOk As String, v() As String) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, lbl As Variant, m As Object
    gRcvShown = Join(v, Chr(9))
    gRcvBack = ""
    If gUiAuto <> 0 Then
        If gRcvInSet Then
            For i = 0 To 6
                If gRcvIn(i) <> RCV_KEEP Then v(i) = gRcvIn(i)
            Next i
            gRcvInSet = False
            gRcvBack = Join(v, Chr(9))
            ReceiptDialog = True
        End If
        Exit Function
    End If
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 250
    dm.Height = 170
    lbl = Array("Фактически принято (F)", "Количество по документу (G)", "Номер документа (C)", "Дата документа (O), дд.мм.гггг", _
        "Дата поступления (N), дд.мм.гггг", "Место хранения (U)", "Цена (J)")
    For i = 0 To 6
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = 11 + i * 18
        m.Width = 108
        m.Height = 12
        m.Label = lbl(i)
        dm.insertByName("l" & i, m)
        m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        m.PositionX = 118
        m.PositionY = 8 + i * 18
        m.Width = 124
        m.Height = 14
        m.Text = v(i)
        dm.insertByName("e" & i, m)
    Next i
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 118
    m.PositionY = 148
    m.Width = 60
    m.Height = 15
    m.Label = sOk
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 182
    m.PositionY = 148
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        For i = 0 To 6
            v(i) = dlg.getControl("e" & i).getText()
        Next i
        gRcvBack = Join(v, Chr(9))
        ReceiptDialog = True
    End If
    dlg.dispose()
End Function

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestRcvInput(sF As String, sG As String, sC As String, sO As String, sN As String, sU As String, sJ As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRcvInput = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gRcvIn(0) = sF
    gRcvIn(1) = sG
    gRcvIn(2) = sC
    gRcvIn(3) = sO
    gRcvIn(4) = sN
    gRcvIn(5) = sU
    gRcvIn(6) = sJ
    gRcvInSet = True
    TestRcvInput = "OK"
End Function

' the fields of the last receipt dialog: what the window proposed and what «OK» returned (tab-separated F G C O N U J;
' the returned part is empty when the window was closed without «OK»)
Function TestRcvBlockMessage() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRcvBlockMessage = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestRcvBlockMessage = gRcvBlockMsg
End Function

Function TestRcvDialog() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRcvDialog = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestRcvDialog = gRcvShown & Chr(10) & gRcvBack
End Function
