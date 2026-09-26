' WmsReturnUi — кнопки листа «Возврат» (задание Core Phase 4): «Найти выдачу», «Проверить», «Провести», «Исправить»,
' «Удалить», «Очистить». Как и на «Выдачах» и «Заказах», успешная работа окон не открывает: результат виден в строке
' (A № возврата, L M остаток до и после, N контроль). Окно — для выбора выдачи и ввода количества, даты и места, для
' исправления, для подтверждения сторно и когда действие невозможно.
Option Explicit

' test seam (TEST books only): the answers of the next return window — quantity, date, place (RET_KEEP: the value the
' window proposed stays, as when the user does not touch the field), the line picked in the list of issues (-1: as
' proposed) — and the answer of the question «ЕИ» of «Найти выдачу»
Global gRetInSet As Boolean
Global gRetIn(2) As String
Global gRetPick As Integer
Global gRetKeySet As Boolean
Global gRetKey As String
' what the last return window proposed and what it returned, its list of issues (the test seam reads them)
Global gRetShown As String
Global gRetBack As String
Global gRetList As String
' the text of the window that ended the last block of «Провести»
Global gRetBlockMsg As String
Private Const RET_KEEP = "<как предложено>"

' ================================================================ the row under the cursor

' blocks of selected rows on «Возврат»: Array(r0, r1) each; False with a message when the cursor is not on the sheet
Function ActiveReturnBlocks(ByRef blocks As Variant, ByRef why As String) As Boolean
    Dim sel As Object, addrs As Variant, ctl As Object, i As Long, n As Long, out() As Variant, r0 As Long
    WmsInit()
    On Error GoTo EH
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().getName() <> SH_RETURNS Then
        why = "Перейдите на лист «" & SH_RETURNS & "» и поставьте курсор в строку возврата."
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
        why = "Поставьте курсор в строку возврата, а не в заголовок."
        Exit Function
    End If
    ReDim Preserve out(n - 1)
    blocks = out
    ActiveReturnBlocks = True
    Exit Function
EH:
    why = "Поставьте курсор в строку возврата на листе «" & SH_RETURNS & "»."
End Function

' exactly one row under the cursor
Function ActiveReturnRow(ByRef r As Long, ByRef why As String) As Boolean
    Dim blocks As Variant
    If Not ActiveReturnBlocks(blocks, why) Then Exit Function
    If UBound(blocks) > 0 Or blocks(0)(0) <> blocks(0)(1) Then
        why = "Выделите одну строку возврата (сейчас выделено несколько строк)."
        Exit Function
    End If
    r = blocks(0)(0)
    ActiveReturnRow = True
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

' the refusal for a row that is not an open (new) return
Private Function NotOpenText(kind As String) As String
    Select Case kind
    Case "POSTED"
        NotOpenText = "Строка — проведённый возврат № " & gRtN & ". Для нового возврата выберите пустую строку; изменить этот — «Исправить»."
    Case "DELETED"
        NotOpenText = "Строка — удалённый (сторно) возврат № " & gRtN & ", она хранит историю. Для нового возврата выберите пустую строку."
    Case Else
        NotOpenText = "Строка — КОПИЯ, её можно только очистить кнопкой «Очистить»."
    End Select
End Function

' ================================================================ «Найти выдачу»

' Fills the row under the cursor for a new return: the issue (by its № in B, otherwise from the list of the posted
' issues of the EI in C, or of the EI asked for, that can still be returned; the recipient in I narrows the list), the
' quantity, the date of the return (today is proposed, it can be changed) and the place (the current place of the EI is
' proposed). Nothing is posted: the row shows what the issue allows, «Провести» posts it.
Sub BtnRetFind(Optional oEvent As Variant)
    Dim r As Long, why As String, kind As String, sh As Object, k As Long, n As Long, canon As String, msg As String, st As Integer
    Dim sKey As String, found As String, ks As Variant, texts() As String, i As Integer, v(2) As String, pick As Integer, ok As Boolean
    Dim sInfo As String, sPlace As String, avail As Double, who As String, full As String, note As String, d As Double, um As Object
    Dim inCtx As Boolean, c As Object, sd As Variant
    If Not ActiveReturnRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsReturn.ReturnRowKind(r)
    If kind <> "EMPTY" And kind <> "OPEN" Then
        Refuse(NotOpenText(kind))
        Exit Sub
    End If
    sh = WmsReturn.ReturnsSheet()
    who = Trim(sh.getCellByPosition(RC_WHO, r).getString())
    full = ""
    If who <> "" Then
        st = WmsIssue.ResolveRecipient(who, full, note)
        If st = 3 Then full = ""
    End If
    k = WmsReturn.IssueNoOf(sh.getCellByPosition(RC_ISSUE, r))
    If k < 0 Then
        Refuse("№ выдачи (B) «" & sh.getCellByPosition(RC_ISSUE, r).getString() & "» — не номер. Исправьте его или очистите B, чтобы выбрать выдачу по ЕИ.")
        Exit Sub
    End If
    If k > 0 Then
        ' the issue is known: its window (the list holds just this issue)
        why = WmsReturn.IssueForDialog(k, sInfo, sPlace, avail)
        If why <> "" Then
            Refuse("Возврат по выдаче № " & k & " невозможен: " & why & ".")
            Exit Sub
        End If
        If avail <= 0.0000001 Then
            Refuse("По выдаче № " & k & " всё уже возвращено: " & WmsReturn.IssueLine(k) & ".")
            Exit Sub
        End If
        canon = gRtIssueEI
        c = sh.getCellByPosition(RC_EI, r)
        If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
            If Not WmsIssue.NormalizeEI(c.getString(), n, canon, msg) Then
                Refuse("ЕИ в строке (C): " & msg & ".")
                Exit Sub
            End If
            If canon <> gRtIssueEI Then
                Refuse("В строке ЕИ " & canon & ", а выдача № " & k & " — на " & gRtIssueEI & ". Возврат относится к ЕИ исходной выдачи: " _
                    & "исправьте C или № выдачи в B.")
                Exit Sub
            End If
        End If
        If full <> "" And LCase(full) <> LCase(gRtIssueWho) Then
            Refuse("В строке получатель «" & full & "», а выдача № " & k & " была «" & gRtIssueWho & "». Возврат от другого получателя " _
                & "не поддерживается (будущий сценарий передачи).")
            Exit Sub
        End If
        ks = Array(k)
    Else
        ' the EI: from C, otherwise asked
        c = sh.getCellByPosition(RC_EI, r)
        If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
            sKey = c.getString()
        Else
            sKey = AskEI()
            If Trim(sKey) = "" Then
                gUiLastMsg = "SKIP:ЕИ не указан"
                Exit Sub
            End If
        End If
        If Not WmsIssue.NormalizeEI(sKey, n, canon, msg) Then
            Refuse(msg & ".")
            Exit Sub
        End If
        why = WmsIssue.StockRowProblem(n, canon)
        If why <> "" Then
            Refuse(why & ". Возврат относится только к ЕИ, который есть в реестре.")
            Exit Sub
        End If
        found = WmsReturn.IssuesOfEI(canon, full, ok)
        If Not ok Then
            Refuse("Список выдач не удалось получить формулой. Введите № выдачи в B вручную (он есть в колонке A листа «" & SH_ISSUES & "»).")
            Exit Sub
        End If
        If found = "" Then
            Refuse("По " & canon & IIf(full <> "", " для «" & full & "»", "") & " нет проведённых выдач, по которым можно вернуть товар " _
                & "(возвращать можно только по действующей выдаче, в пределах выданного).")
            Exit Sub
        End If
        ks = Split(found, ";")
        sd = WmsIssue.StockData(n)
        sInfo = canon & " " & sd(SC_NAME) & IIf(full <> "", " — выдачи «" & full & "»", "") & ": выберите выдачу, по которой возвращают" _
            & IIf(UBound(ks) + 1 >= ISSUE_LIST_MAX, " (показаны последние " & ISSUE_LIST_MAX & ")", "")
        sPlace = CStr(sd(SC_PLACE))
    End If
    ReDim texts(UBound(ks))
    For i = 0 To UBound(ks)
        texts(i) = WmsReturn.IssueLine(CLng(ks(i)))
    Next i
    ' the values proposed: what the row already holds, otherwise today and the current place
    v(0) = sh.getCellByPosition(RC_QTY, r).getString()
    If sh.getCellByPosition(RC_DATE, r).getType() = com.sun.star.table.CellContentType.VALUE Then
        v(1) = Format(sh.getCellByPosition(RC_DATE, r).getValue(), "DD.MM.YYYY")
    ElseIf sh.getCellByPosition(RC_DATE, r).getString() <> "" Then
        v(1) = sh.getCellByPosition(RC_DATE, r).getString()
    Else
        v(1) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    End If
    v(2) = Trim(sh.getCellByPosition(RC_PLACE, r).getString())
    If v(2) = "" Then v(2) = sPlace
    pick = 0
    If Not ReturnDialog("Найти выдачу — возврат", sInfo, "Заполнить строку", texts, pick, v) Then
        gUiLastMsg = "SKIP:окно закрыто без выбора"
        Exit Sub
    End If
    If pick < 0 Or pick > UBound(ks) Then
        Refuse("Не выбрана выдача.")
        Exit Sub
    End If
    k = CLng(ks(pick))
    ' the row is filled like the user would type it (one undo step); the check shows the result in «Контроль»
    On Error GoTo EH
    WmsReturn.HandlerOff(True)
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: найти выдачу")
    inCtx = True
    sh.getCellByPosition(RC_ISSUE, r).setValue(k)
    If sh.getCellByPosition(RC_EI, r).getType() = com.sun.star.table.CellContentType.EMPTY Then sh.getCellByPosition(RC_EI, r).setString(canon)
    If Trim(v(0)) <> "" Then
        sh.getCellByPosition(RC_QTY, r).setString(Trim(v(0)))
    ElseIf sh.getCellByPosition(RC_QTY, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        ClearCell(sh.getCellByPosition(RC_QTY, r))
    End If
    If Trim(v(1)) = "" Then
        If sh.getCellByPosition(RC_DATE, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(RC_DATE, r))
    ElseIf WmsOrders.ParseDateTextAs(v(1), "дата возврата", d, msg) Then
        sh.getCellByPosition(RC_DATE, r).setValue(d)
    Else
        sh.getCellByPosition(RC_DATE, r).setString(Trim(v(1)))
    End If
    If Trim(v(2)) <> "" Then
        WmsOrders.SetIfDiff(sh.getCellByPosition(RC_PLACE, r), Trim(v(2)))
    ElseIf sh.getCellByPosition(RC_PLACE, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        ClearCell(sh.getCellByPosition(RC_PLACE, r))
    End If
    WmsReturn.PreviewReturnRow(r, False)
    um.leaveUndoContext()
    inCtx = False
    WmsReturn.HandlerOff(False)
    gUiLastMsg = "OK:" & sh.getCellByPosition(RC_CTL, r).getString()
    Exit Sub
EH:
    gUiLastMsg = "ERR:строка не заполнена: " & Error$
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    WmsReturn.HandlerOff(False)
    UiMessage(Mid(gUiLastMsg, 5))
End Sub

' the EI for «Найти выдачу» when the row has neither an issue № nor an EI
Private Function AskEI() As String
    If gUiAuto <> 0 Then
        If gRetKeySet Then AskEI = gRetKey
        gRetKeySet = False
        Exit Function
    End If
    AskEI = InputBox("ЕИ товара, который возвращают (например ЕИ-00000123 или 123):", "WMS — найти выдачу")
End Function

' ================================================================ buttons

' «Проверить»: open rows get the result in «Контроль»; for a posted return the check result is shown
Sub BtnRetCheck(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, r As Long, nErr As Long, n As Long, sh As Object, b As Long, r0 As Long
    If Not ActiveReturnBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsReturn.ReturnCheckRow(r0)
        If WmsReturn.ReturnRowKind(r0) <> "OPEN" Then
            If Left(res, 3) = "OK:" Then UiInfo(Payload(res)) Else UiMessage(Payload(res))
        End If
        gUiLastMsg = res
        Exit Sub
    End If
    sh = WmsReturn.ReturnsSheet()
    For b = 0 To UBound(blocks)
        For r = blocks(b)(0) To blocks(b)(1)
            If sh.getRows().getByIndex(r).IsVisible Then
                If WmsReturn.ReturnRowKind(r) = "OPEN" Then
                    n = n + 1
                    res = WmsReturn.ReturnCheckRow(r)
                    If Left(res, 3) <> "OK:" Then nErr = nErr + 1
                End If
            End If
        Next r
    Next b
    If nErr > 0 Then UiMessage("Проверено строк возврата: " & n & ", с ошибками: " & nErr & " — причины в колонке «Контроль».")
    gUiLastMsg = "OK:проверено строк " & n & ", с ошибками " & nErr
End Sub

' «Провести»: the row under the cursor, or every visible row of the selected block(s) with a quantity F (D-059: each row
' its own operation; a problem of one row does not stop the others; a system error of WMS stops the whole block)
Sub BtnRetPost(Optional oEvent As Variant)
    Dim blocks As Variant, why As String, res As String, a As Variant, b As Long, r0 As Long, nOk As Long, nErr As Long, nSkip As Long
    Dim first As String, sysErr As String, stopRow As Long, parts As Variant, msg As String
    If Not ActiveReturnBlocks(blocks, why) Then
        UiMessage(why)
        Exit Sub
    End If
    r0 = blocks(0)(0)
    If UBound(blocks) = 0 And blocks(0)(1) = r0 Then
        res = WmsReturn.ReturnPostRow(r0)
        AfterOperation("«Провести» (возврат)", res)
        If Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then UiMessage(Payload(res))
        gUiLastMsg = res
        Exit Sub
    End If
    For b = 0 To UBound(blocks)
        parts = Split(WmsReturn.ReturnPostRange(blocks(b)(0), blocks(b)(1)), Chr(10))
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
    msg = "«Провести» для выделенных строк возврата: проведено " & nOk & ", пропущено " & nSkip _
        & " (нет количества F, уже проведены, удалены или копии), отклонено с ошибкой " & nErr _
        & IIf(nErr > 0, " — причины в колонке «Контроль»", "") & "."
    If first <> "" Then msg = msg & Chr(10) & "Первая ошибка: " & first
    gRetBlockMsg = msg
    If stopRow > 0 Then
        msg = msg & Chr(10) & Chr(10) & "ОБРАБОТКА ОСТАНОВЛЕНА — системная ошибка WMS, " & SysReason(sysErr) _
            & Chr(10) & "Строки после неё не обрабатывались. Проверьте лист «" & SH_MAIN & "»."
        gRetBlockMsg = msg
        UiMessage(msg)
        If InStr(sysErr, "BLOCKED:") > 0 Or InStr(sysErr, "ERR-CRITICAL:") > 0 Then UiRefresh("«Провести» (возврат): " & sysErr)
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
        SysReason = Left(sysErr, p + 1) & "возврат проведён " & Mid(rest, InStr(rest, "("))
    Else
        SysReason = Left(sysErr, p + 1) & Payload(rest)
    End If
End Function

' «Исправить»: quantity, date and place of a posted return in a window prefilled with the posted values (RETURN_FIX)
Sub BtnRetFix(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, v(2) As String, sh As Object, k As Long, n As Long, sInfo As String
    Dim sPlace As String, avail As Double, q1 As Double, pick As Integer, empty0() As String
    If Not ActiveReturnRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsReturn.ReturnRowKind(r)
    Select Case kind
    Case "POSTED"
    Case "OPEN"
        Refuse("Строка не проведена — исправьте значения прямо в строке и нажмите «Провести».")
        Exit Sub
    Case "DELETED"
        Refuse("Возврат № " & gRtN & " удалён (сторно) — исправлять нечего.")
        Exit Sub
    Case "EMPTY"
        Refuse("Поставьте курсор в строку проведённого возврата.")
        Exit Sub
    Case Else
        Refuse("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    n = gRtN
    sh = WmsReturn.ReturnsSheet()
    k = CLng(sh.getCellByPosition(RC_ISSUE, r).getValue())
    q1 = sh.getCellByPosition(RC_QTY, r).getValue()
    why = WmsReturn.IssueForDialog(k, sInfo, sPlace, avail)
    If why <> "" Then
        Refuse("Возврат № " & n & " не исправить: " & why & ".")
        Exit Sub
    End If
    sInfo = sInfo & ". Сейчас возвращено этим возвратом " & WmsIssue.QtyText(q1) & ", его можно увеличить до " & WmsIssue.QtyText(WmsIssue.Round3(avail + q1))
    v(0) = WmsIssue.QtyText(q1)
    v(1) = Format(sh.getCellByPosition(RC_DATE, r).getValue(), "DD.MM.YYYY")
    v(2) = sh.getCellByPosition(RC_PLACE, r).getString()
    pick = -1
    If Not ReturnDialog("Исправить возврат № " & n & " (та же выдача, тот же ЕИ)", sInfo, "Исправить", empty0, pick, v) Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        Exit Sub
    End If
    res = WmsReturn.ReturnFixRow(r, v(0), v(1), v(2))
    AfterOperation("«Исправить» (возврат)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Исправление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

' «Удалить»: the storno of a posted return (RETURN_DEL) after a confirmation
Sub BtnRetDelete(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, kind As String, sh As Object
    If Not ActiveReturnRow(r, why) Then
        Refuse(why)
        Exit Sub
    End If
    kind = WmsReturn.ReturnRowKind(r)
    Select Case kind
    Case "POSTED"
    Case "DELETED"
        ' a repeated click after the storno: nothing to do
        UiMessage("Возврат № " & gRtN & " уже удалён (сторно).")
        gUiLastMsg = "SKIP:возврат уже удалён (сторно)"
        Exit Sub
    Case "OPEN"
        Refuse("Строка не проведена — чтобы убрать ввод, нажмите «Очистить».")
        Exit Sub
    Case "EMPTY"
        Refuse("Поставьте курсор в строку проведённого возврата.")
        Exit Sub
    Case Else
        Refuse("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    sh = WmsReturn.ReturnsSheet()
    If Not UiConfirm("Удалить возврат № " & gRtN & "?" & Chr(10) & Chr(10) & sh.getCellByPosition(RC_EI, r).getString() & " — " _
        & sh.getCellByPosition(RC_NAME, r).getString() & ", " & sh.getCellByPosition(RC_QTY, r).getString() & " " & sh.getCellByPosition(RC_UNIT, r).getString() _
        & " (выдача № " & sh.getCellByPosition(RC_ISSUE, r).getString() & ", «" & sh.getCellByPosition(RC_WHO, r).getString() & "»)" & Chr(10) & Chr(10) _
        & "Остаток ЕИ уменьшится на это количество, по выдаче его снова можно будет вернуть. Строка останется в истории с отметкой «" _
        & ST_DELETED & "», № не будет использован повторно.") Then
        gUiLastMsg = "SKIP:удаление не подтверждено"
        Exit Sub
    End If
    res = WmsReturn.ReturnDeleteRow(r)
    AfterOperation("«Удалить» (возврат)", res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage("Удаление не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

Sub BtnRetClear(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String
    If Not ActiveReturnRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    res = WmsReturn.ReturnClearRow(r)
    If Left(res, 3) <> "OK:" Then UiMessage(Payload(res))
    gUiLastMsg = res
End Sub

' ================================================================ the return window: a list of issues (optional), F, H, J

' sInfo: the text above the fields; texts: the issues to choose from (none: no list); pick: in — the line proposed, out —
' the line chosen; v(0..2): quantity, date, place — in the proposed values, out what «OK» returned
Private Function ReturnDialog(sTitle As String, sInfo As String, sOk As String, texts() As String, ByRef pick As Integer, v() As String) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, lbl As Variant, m As Object, y As Long, nList As Integer, sel(0) As Integer
    nList = UBound(texts) + 1
    gRetShown = Join(v, Chr(9)) & Chr(9) & pick
    gRetList = ""
    If nList > 0 Then gRetList = Join(texts, Chr(10))
    gRetBack = ""
    If gUiAuto <> 0 Then
        If gRetInSet Then
            For i = 0 To 2
                If gRetIn(i) <> RET_KEEP Then v(i) = gRetIn(i)
            Next i
            If gRetPick >= 0 Then pick = gRetPick
            gRetInSet = False
            gRetBack = Join(v, Chr(9)) & Chr(9) & pick
            ReturnDialog = True
        End If
        Exit Function
    End If
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    y = 8
    If sInfo <> "" Then
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = y
        m.Width = 284
        m.Height = 26
        m.MultiLine = True
        m.Label = sInfo
        dm.insertByName("info", m)
        y = y + 30
    End If
    If nList > 0 Then
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = y
        m.Width = 284
        m.Height = 10
        m.Label = "Выдача, по которой возвращают:"
        dm.insertByName("lList", m)
        y = y + 11
        m = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
        m.PositionX = 8
        m.PositionY = y
        m.Width = 284
        m.Height = 60
        m.Dropdown = False
        m.MultiSelection = False
        m.StringItemList = texts
        If pick < 0 Or pick >= nList Then pick = 0
        sel(0) = pick
        m.SelectedItems = sel
        dm.insertByName("lst", m)
        y = y + 66
    End If
    lbl = Array("Количество возврата (F)", "Дата возврата (H), дд.мм.гггг", "Место хранения (J)")
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
    m.Label = sOk
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
        If nList > 0 Then pick = dlg.getControl("lst").getSelectedItemPos()
        gRetBack = Join(v, Chr(9)) & Chr(9) & pick
        ReturnDialog = True
    End If
    dlg.dispose()
End Function

' ================================================================ test seam (inert unless _SYS MODE = TEST)

' the answers of the next return window: quantity, date, place (RET_KEEP — as proposed), the line of the list (-1 — as
' proposed)
Function TestRetInput(sQty As String, sDate As String, sPlace As String, iPick As Integer) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRetInput = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gRetIn(0) = sQty
    gRetIn(1) = sDate
    gRetIn(2) = sPlace
    gRetPick = iPick
    gRetInSet = True
    TestRetInput = "OK"
End Function

' the answer of the question «ЕИ» of «Найти выдачу» ("" — the question is cancelled)
Function TestRetKey(sKey As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRetKey = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gRetKey = sKey
    gRetKeySet = True
    TestRetKey = "OK"
End Function

' the last return window: what it proposed and what «OK» returned (tab-separated quantity, date, place, line of the list;
' the returned part is empty when the window was closed without «OK»), then its list of issues (one per line)
Function TestRetDialog() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRetDialog = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestRetDialog = gRetShown & Chr(10) & gRetBack & Chr(10) & gRetList
End Function

' the text of the window that ended the last block of «Провести»
Function TestRetBlockMessage() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRetBlockMessage = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestRetBlockMessage = gRetBlockMsg
End Function
