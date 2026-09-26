' WmsAdjustUi — кнопки листа «Корректировки» (задание «FINAL WMS MARATHON», §2): «Проверить», «Провести», «Исправить»,
' «Удалить», «Очистить». Как на других листах, успешная работа окон не открывает: результат виден в строке (A №,
' M N O место и остаток до/после, P разница, Q контроль). Окно — для исправления, для подтверждения сторно и когда
' действие невозможно.
Option Explicit

' test seam (TEST books only): the answers of the next correction window — the main value (new place / quantity /
' actual balance), the date and the reason (ADJ_KEEP: the value the window proposed stays)
Global gAdjInSet As Boolean
Global gAdjIn(2) As String
' what the last correction window proposed and what it returned (the test seam reads them)
Global gAdjShown As String
Global gAdjBack As String
' the text of the window that ended the last block of «Провести»
Global gAdjBlockMsg As String
Private Const ADJ_KEEP = "<как предложено>"

' ================================================================ the rows under the cursor

Function ActiveAdjustBlocks(ByRef blocks As Variant, ByRef why As String) As Boolean
    Dim sel As Object, addrs As Variant, ctl As Object, i As Long, n As Long, out() As Variant, r0 As Long
    WmsInit()
    On Error GoTo EH
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().getName() <> SH_ADJUST Then
        why = "Перейдите на лист «" & SH_ADJUST & "» и поставьте курсор в строку корректировки."
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
        why = "Поставьте курсор в строку корректировки, а не в заголовок."
        Exit Function
    End If
    ReDim Preserve out(n - 1)
    blocks = out
    ActiveAdjustBlocks = True
    Exit Function
EH:
    why = "Поставьте курсор в строку корректировки на листе «" & SH_ADJUST & "»."
End Function

Function ActiveAdjustRow(ByRef r As Long, ByRef why As String) As Boolean
    Dim blocks As Variant
    If Not ActiveAdjustBlocks(blocks, why) Then Exit Function
    If UBound(blocks) > 0 Or blocks(0)(0) <> blocks(0)(1) Then
        why = "Выделите одну строку корректировки (сейчас выделено несколько строк)."
        Exit Function
    End If
    r = blocks(0)(0)
    ActiveAdjustRow = True
End Function

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

' the refusal for a row that is not a posted correction (for «Исправить» and «Удалить»)
Private Function NotPostedText(kind As String, sAction As String) As String
    Select Case kind
    Case "OPEN"
        If sAction = "Исправить" Then
            NotPostedText = "Строка не проведена — исправьте значения прямо в строке и нажмите «Провести»."
        Else
            NotPostedText = "Строка не проведена — чтобы убрать ввод, нажмите «Очистить»."
        End If
    Case "DELETED"
        NotPostedText = "Корректировка № " & gAjN & " удалена (сторно) — " & IIf(sAction = "Исправить", "исправлять нечего.", "строка хранит историю.")
    Case "EMPTY"
        NotPostedText = "Поставьте курсор в строку проведённой корректировки."
    Case Else
        NotPostedText = "Строка — КОПИЯ, её можно только очистить кнопкой «Очистить»."
    End Select
End Function

' ================================================================ buttons

Sub BtnAdjCheck(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, r As Long, nErr As Long, n As Long, sh As Object, b As Long, r0 As Long
    If Not ActiveAdjustBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsAdjust.AdjustCheckRow(r0)
        If WmsAdjust.AdjustRowKind(r0) <> "OPEN" Then
            If Left(res, 3) = "OK:" Then UiInfo(Payload(res)) Else UiMessage(Payload(res))
        End If
        gUiLastMsg = res
        Exit Sub
    End If
    sh = WmsAdjust.AdjustSheet()
    For b = 0 To UBound(blocks)
        For r = blocks(b)(0) To blocks(b)(1)
            If sh.getRows().getByIndex(r).IsVisible Then
                If WmsAdjust.AdjustRowKind(r) = "OPEN" Then
                    n = n + 1
                    res = WmsAdjust.AdjustCheckRow(r)
                    If Left(res, 3) <> "OK:" Then nErr = nErr + 1
                End If
            End If
        Next r
    Next b
    If nErr > 0 Then UiMessage("Проверено строк корректировок: " & n & ", с ошибками: " & nErr & " — причины в колонке «Контроль».")
    gUiLastMsg = "OK:проверено строк " & n & ", с ошибками " & nErr
End Sub

' «Провести»: the row under the cursor, or every visible open row with a kind (B) of the selected block(s) (D-059)
Sub BtnAdjPost(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, a As Variant, b As Long, r0 As Long, nOk As Long, nErr As Long, nSkip As Long
    Dim first As String, sysErr As String, stopRow As Long, parts As Variant, msg As String
    If Not ActiveAdjustBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsAdjust.AdjustPostRow(r0)
        AfterOperation("«Провести» (корректировка)", res)
        If Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then UiMessage(Payload(res))
        gUiLastMsg = res
        Exit Sub
    End If
    For b = 0 To UBound(blocks)
        parts = Split(WmsAdjust.AdjustPostRange(blocks(b)(0), blocks(b)(1)), Chr(10))
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
    msg = "«Провести» для выделенных строк корректировок: проведено " & nOk & ", пропущено " & nSkip _
        & " (нет вида B, уже проведены, удалены или копии), отклонено с ошибкой " & nErr _
        & IIf(nErr > 0, " — причины в колонке «Контроль»", "") & "."
    If first <> "" Then msg = msg & Chr(10) & "Первая ошибка: " & first
    gAdjBlockMsg = msg
    If stopRow > 0 Then
        msg = msg & Chr(10) & Chr(10) & "ОБРАБОТКА ОСТАНОВЛЕНА — системная ошибка WMS: " & sysErr _
            & Chr(10) & "Строки после неё не обрабатывались. Проверьте лист «" & SH_MAIN & "»."
        gAdjBlockMsg = msg
        UiMessage(msg)
        If InStr(sysErr, "BLOCKED:") > 0 Or InStr(sysErr, "ERR-CRITICAL:") > 0 Then UiRefresh("«Провести» (корректировка): " & sysErr)
    ElseIf nErr > 0 Then
        UiMessage(msg)
    Else
        UiInfo(msg)
    End If
    gUiLastMsg = res
End Sub

' «Исправить»: the main value (new place / quantity / actual balance), the date and the reason of a posted correction in a
' window prefilled with the posted values
Sub BtnAdjFix(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, v(2) As String, sh As Object, n As Long, code As String, lbl0 As String
    Dim sInfo As String
    If Not ActiveAdjustRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsAdjust.AdjustRowKind(r)
    If kind <> "POSTED" Then
        Refuse(NotPostedText(kind, "Исправить"))
        Exit Sub
    End If
    n = gAjN
    sh = WmsAdjust.AdjustSheet()
    code = WmsAdjust.KindCode(sh.getCellByPosition(AC_KIND, r).getString())
    Select Case code
    Case "MOVE"
        lbl0 = "Новое место (J)"
        v(0) = sh.getCellByPosition(AC_PLACE, r).getString()
        sInfo = "Перемещение " & sh.getCellByPosition(AC_EI, r).getString() & " из «" & sh.getCellByPosition(AC_FROM, r).getString() & "»."
    Case "WRITE_OFF"
        lbl0 = "Списать, количество (G)"
        v(0) = WmsIssue.QtyText(sh.getCellByPosition(AC_QTY, r).getValue())
        sInfo = "Списание " & sh.getCellByPosition(AC_EI, r).getString() & " " & sh.getCellByPosition(AC_NAME, r).getString() & "."
    Case "INV_ADJ"
        lbl0 = "Фактический остаток (H)"
        v(0) = WmsIssue.QtyText(sh.getCellByPosition(AC_FACT, r).getValue())
        sInfo = "Инвентаризация " & sh.getCellByPosition(AC_EI, r).getString() & ": учётный остаток при пересчёте " _
            & WmsIssue.QtyText(sh.getCellByPosition(AC_BOOK, r).getValue()) & " (не меняется)."
    Case Else
        Refuse("Вид корректировки № " & n & " не распознан — нужна самопроверка.")
        Exit Sub
    End Select
    v(1) = Format(sh.getCellByPosition(AC_DATE, r).getValue(), "DD.MM.YYYY")
    v(2) = sh.getCellByPosition(AC_REASON, r).getString()
    If Not AdjustDialog("Исправить корректировку № " & n & " (тот же ЕИ, тот же вид)", sInfo, Array(lbl0, "Дата (K), дд.мм.гггг", _
        IIf(code = "INV_ADJ", "Основание (L)", "Причина (L)")), v) Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        Exit Sub
    End If
    res = WmsAdjust.AdjustFixRow(r, v(0), v(1), v(2))
    AfterOperation("«Исправить» (корректировка)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Исправление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

' «Удалить»: the storno of a posted correction after a confirmation; a refusal is the first window (no confirmation)
Sub BtnAdjDelete(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, sh As Object, rd As Variant, code As String, what As String
    If Not ActiveAdjustRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsAdjust.AdjustRowKind(r)
    If kind = "DELETED" Then
        UiMessage("Корректировка № " & gAjN & " уже удалена (сторно).")
        gUiLastMsg = "SKIP:корректировка уже удалена (сторно)"
        Exit Sub
    End If
    If kind <> "POSTED" Then
        Refuse(NotPostedText(kind, "Удалить"))
        Exit Sub
    End If
    why = WmsAdjust.DeleteProblem(r, rd)
    If why <> "" Then
        Refuse("Удаление невозможно: " & Payload(why))
        Exit Sub
    End If
    sh = WmsAdjust.AdjustSheet()
    code = CStr(rd(AJ_KIND))
    Select Case code
    Case "MOVE"
        what = "ЕИ вернётся на место «" & rd(AJ_FROM) & "»."
    Case "WRITE_OFF"
        what = "Списанное количество " & WmsIssue.QtyText(-rd(AJ_QTY)) & " вернётся в остаток."
    Case Else
        what = "Разница " & IIf(rd(AJ_QTY) > 0, "+", "") & WmsIssue.QtyText(rd(AJ_QTY)) & " будет снята с остатка."
    End Select
    If Not UiConfirm("Удалить корректировку № " & gAjN & " (" & WmsAdjust.KindName(code) & ")?" & Chr(10) & Chr(10) _
        & sh.getCellByPosition(AC_EI, r).getString() & " — " & sh.getCellByPosition(AC_NAME, r).getString() & Chr(10) & Chr(10) & what _
        & " Строка останется в истории с отметкой «" & ST_DELETED & "», № не будет использован повторно.") Then
        gUiLastMsg = "SKIP:удаление не подтверждено"
        Exit Sub
    End If
    res = WmsAdjust.AdjustDeleteRow(r)
    AfterOperation("«Удалить» (корректировка)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Удаление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

Sub BtnAdjClear(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String
    If Not ActiveAdjustRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    res = WmsAdjust.AdjustClearRow(r)
    If Left(res, 3) <> "OK:" Then UiMessage(Payload(res))
    gUiLastMsg = res
End Sub

' ================================================================ the correction window: three fields

Private Function AdjustDialog(sTitle As String, sInfo As String, lbl As Variant, v() As String) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, m As Object, y As Long
    gAdjShown = Join(v, Chr(9))
    gAdjBack = ""
    If gUiAuto <> 0 Then
        If gAdjInSet Then
            For i = 0 To 2
                If gAdjIn(i) <> ADJ_KEEP Then v(i) = gAdjIn(i)
            Next i
            gAdjInSet = False
            gAdjBack = Join(v, Chr(9))
            AdjustDialog = True
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
    For i = 0 To 2
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = y + 3
        m.Width = 110
        m.Height = 12
        m.Label = lbl(i)
        dm.insertByName("l" & i, m)
        m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        m.PositionX = 120
        m.PositionY = y
        m.Width = 172
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
        For i = 0 To 2
            v(i) = dlg.getControl("e" & i).getText()
        Next i
        gAdjBack = Join(v, Chr(9))
        AdjustDialog = True
    End If
    dlg.dispose()
End Function

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestAdjInput(sMain As String, sDate As String, sReason As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestAdjInput = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gAdjIn(0) = sMain
    gAdjIn(1) = sDate
    gAdjIn(2) = sReason
    gAdjInSet = True
    TestAdjInput = "OK"
End Function

' the last correction window: what it proposed and what «OK» returned (tab-separated; empty when closed without «OK»)
Function TestAdjDialog() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestAdjDialog = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestAdjDialog = gAdjShown & Chr(10) & gAdjBack
End Function

Function TestAdjBlockMessage() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestAdjBlockMessage = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestAdjBlockMessage = gAdjBlockMsg
End Function
