' WmsUi — минимальный пользовательский слой (задание Core Phase 2): состояние WMS и причина блокировки на листе
' «Главная», кнопки восстановления Phase 1 и кнопки «Провести», «Исправить», «Удалить», «Очистить» на листе «Выдачи»
' (кнопки «Заказов» — WmsOrdersUi, «Возврата» — WmsReturnUi).
' Успешная работа окон не открывает: результат виден в «Контроль» строки. Окно появляется только для подтверждения
' опасного действия, для ввода исправления и когда действие невозможно.
Option Explicit

' «Главная»: column B of these rows (0-based) holds the values, column A the captions (written by tools/build_ods.py)
Private Const MR_STATE = 2
Private Const MR_REASON = 3
Private Const MR_HINT = 4
Private Const MR_UNPOSTED = 5
Private Const MR_NOTES = 6
Private Const MR_LAST = 7
Private Const MR_TIME = 8

' test seam (TEST books only): 0 interactive, 1 confirmations answer «Да», 2 answer «Нет»; messages are only recorded.
' Global on purpose: LibreOffice Basic resets module-level Private variables at every new macro call.
Global gUiAuto As Integer
Global gUiLastMsg As String
Global gUiConfirms As Long
Global gUiFixSet As Boolean
Global gUiFixEI As String
Global gUiFixQty As String
Global gUiFixWho As String
Global gUiFixDate As String

' filtered WMS ranges whose rows were shown for a save (FiltersStash) and are filtered again after it (FiltersRestore)
Global gFltN As Integer
Global gFltName(6) As String

' ================================================================ status panel «Главная»

Private Sub PutB(sh As Object, r As Integer, s As String)
    Dim c As Object
    c = sh.getCellByPosition(1, r)
    If s = "" Then
        If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(c)
    ElseIf c.getString() <> s Then
        c.setString(s)
    End If
End Sub

Function BlockHint(code As String) As String
    Select Case code
    Case "TAIL"
        BlockHint = "Нажмите «Восстановить»: операции из журнала будут доведены в книге. Если проводить их нельзя — «Отложить хвост журнала» " _
            & "(они останутся в журнале для ручной сверки)."
    Case "LOCKED"
        BlockHint = "Если WMS точно не открыта в другом окне или на другом компьютере — нажмите «Снять блокировку WMS»."
    Case "LOCK_LOST"
        BlockHint = "Блокировку перехватил другой экземпляр WMS. Закройте эту книгу без сохранения и откройте WMS заново."
    Case "NOT_REGISTERED"
        BlockHint = "Если этот файл — рабочая WMS (перенесена намеренно), нажмите «Сделать рабочим файлом». Резервную копию так открывать нельзя: " _
            & "скопируйте её на место рабочей книги."
    Case "JOURNAL_BEHIND"
        BlockHint = "Журнал отстаёт от книги. Если WMS перенесена без папки WMS_Journal — «Сделать рабочим файлом» (будет начат новый журнал). " _
            & "Иначе верните папку WMS_Journal рядом с книгой."
    Case "ROLLBACK_FAILED", "COMMIT_FAILED", "JOURNAL_UNCERTAIN", "JOURNAL_FOREIGN_WRITE", "STARTUP_ERROR"
        BlockHint = "Закройте книгу и откройте WMS заново — запуск сам разберёт незавершённую операцию."
    Case "NO_FILE", "READ_ONLY", "NOT_ODS"
        BlockHint = "Откройте рабочую книгу WMS (.ods) для записи."
    Case ""
        BlockHint = ""
    Case Else
        BlockHint = "Проведение остановлено для защиты данных. Не сохраняйте книгу и обратитесь к ответственному за WMS: " _
            & "нужна проверка по журналу и резервной копии. Кнопка «Самопроверка» покажет подробности."
    End Select
End Function

' The panel is derived state (recomputed at every start and action): writing it does not make an unchanged book
' «modified», otherwise LibreOffice would ask to save after merely opening and closing WMS.
Sub UiRefresh(sLast As String)
    Dim sh As Object, n As Long, c As Object, wasModified As Boolean, color As Long, s As String
    On Error GoTo EH
    If IsNull(gDoc) Or IsEmpty(gDoc) Then Exit Sub
    If Not gDoc.Sheets.hasByName(SH_MAIN) Then Exit Sub
    wasModified = gDoc.isModified()
    sh = gDoc.Sheets.getByName(SH_MAIN)
    c = sh.getCellByPosition(1, MR_STATE)
    If gState = "CLEAN" Then
        PutB(sh, MR_STATE, "Работа разрешена")
        PutB(sh, MR_REASON, "")
        PutB(sh, MR_HINT, "")
        color = RGB(198, 239, 206)
    Else
        PutB(sh, MR_STATE, "ЗАБЛОКИРОВАНО" & IIf(gBlock <> "", " [" & gBlock & "]", "") & " — проведение запрещено")
        PutB(sh, MR_REASON, gBlockText)
        PutB(sh, MR_HINT, BlockHint(gBlock))
        color = RGB(255, 199, 206)
    End If
    If c.CellBackColor <> color Then c.CellBackColor = color
    n = WmsIssue.UnpostedCount()
    If n > 0 Then
        s = "выдачи: " & IIf(gUnpostedMore, "не менее ", "") & n & " строк(и) с ЕИ без № — ещё не проведены"
    ElseIf n = 0 Then
        s = "выдачи: нет"
    Else
        s = "выдачи: не удалось посчитать — отфильтруйте «№» = пусто"
    End If
    n = WmsOrders.UnpostedReceipts()
    If n > 0 Then
        s = s & Chr(10) & "приходы: " & IIf(gRcvUnpostedMore, "не менее ", "") & n & " строк(и) «Заказы» с фактом (F) без ЕИ — приход ещё не проведён"
    ElseIf n = 0 Then
        s = s & Chr(10) & "приходы: нет"
    Else
        s = s & Chr(10) & "приходы: не удалось посчитать — отфильтруйте «Внутренний код» = пусто"
    End If
    n = WmsReturn.UnpostedReturns()
    If n > 0 Then
        s = s & Chr(10) & "возвраты: " & n & " строк(и) «Возврат» с количеством (F) без № — ещё не проведены"
    ElseIf n = 0 Then
        s = s & Chr(10) & "возвраты: нет"
    Else
        s = s & Chr(10) & "возвраты: не удалось посчитать — отфильтруйте «№ возврата» = пусто"
    End If
    n = WmsSpecial.UnpostedSpecial()
    If n > 0 Then
        s = s & Chr(10) & "иной приход: " & n & " строк(и) «" & SH_SPECIAL & "» с количеством (F) без № — ещё не проведены"
    ElseIf n = 0 Then
        s = s & Chr(10) & "иной приход: нет"
    Else
        s = s & Chr(10) & "иной приход: не удалось посчитать — отфильтруйте «№ строки» = пусто"
    End If
    n = WmsAdjust.UnpostedAdjustments()
    If n > 0 Then
        s = s & Chr(10) & "корректировки: " & n & " строк(и) «" & SH_ADJUST & "» с видом (B) без № — ещё не проведены"
    ElseIf n = 0 Then
        s = s & Chr(10) & "корректировки: нет"
    Else
        s = s & Chr(10) & "корректировки: не удалось посчитать — отфильтруйте «№» = пусто"
    End If
    PutB(sh, MR_UNPOSTED, s)
    PutB(sh, MR_NOTES, Replace(gReport, " | ", Chr(10)))
    If sLast <> "" Then
        PutB(sh, MR_LAST, sLast)
        PutB(sh, MR_TIME, Replace(TS(), "T", " "))
    End If
    If Not wasModified Then gDoc.setModified(False)
    Exit Sub
EH:
End Sub

' end of every start: the panel is refreshed; a blocked WMS shows «Главная» so that the reason is seen at once
Sub UiAfterStartup()
    On Error Resume Next
    UiRefresh("")
    If gState <> "CLEAN" And gDoc.Sheets.hasByName(SH_MAIN) Then
        gDoc.getCurrentController().setActiveSheet(gDoc.Sheets.getByName(SH_MAIN))
    End If
End Sub

' ================================================================ messages (never on success)

Sub UiMessage(s As String)
    gUiLastMsg = s
    If gUiAuto <> 0 Then Exit Sub
    MsgBox s, 48, "WMS"
End Sub

Function UiConfirm(s As String) As Boolean
    gUiLastMsg = s
    gUiConfirms = gUiConfirms + 1
    If gUiAuto = 1 Then
        UiConfirm = True
    ElseIf gUiAuto = 2 Then
        UiConfirm = False
    Else
        UiConfirm = (MsgBox(s, 4 + 32, "WMS") = 6)
    End If
End Function

' a question with three answers: 6 «Да», 7 «Нет», 2 «Отмена» (test seam: gUiAuto 1 → «Да», 2 → «Нет»)
Function UiAsk3(s As String) As Integer
    gUiLastMsg = s
    gUiConfirms = gUiConfirms + 1
    If gUiAuto = 1 Then
        UiAsk3 = 6
    ElseIf gUiAuto = 2 Then
        UiAsk3 = 7
    Else
        UiAsk3 = MsgBox(s, 3 + 32, "WMS")
    End If
End Function

Private Function FirstPart(s As String) As String
    Dim p As Long
    p = InStr(s, " || ")
    If p > 0 Then FirstPart = Left(s, p - 1) Else FirstPart = s
End Function

' ================================================================ «Главная»: recovery actions of Phase 1 (spec §21)

Sub BtnRecover(Optional oEvent As Variant)
    Dim res As String
    WmsInit()
    res = WmsCore.ActionRecover()
    UiRefresh("«Восстановить»: " & FirstPart(res))
End Sub

Sub BtnAbandon(Optional oEvent As Variant)
    Dim res As String
    WmsInit()
    If gBlock <> "TAIL" Then
        UiMessage("«Отложить хвост журнала» нужно только когда в журнале есть операции, которых нет в книге. Сейчас: " & StateLine())
        Exit Sub
    End If
    If Not UiConfirm("Отложить хвост журнала?" & Chr(10) & Chr(10) & gBlockText & Chr(10) & Chr(10) _
        & "Эти операции НЕ будут применены к книге и останутся в журнале для ручной сверки.") Then Exit Sub
    res = WmsCore.ActionAbandonTail()
    UiRefresh("«Отложить хвост журнала»: " & FirstPart(res))
End Sub

Sub BtnUnlock(Optional oEvent As Variant)
    Dim res As String
    WmsInit()
    If Not UiConfirm("Снять блокировку WMS?" & Chr(10) & Chr(10) & "Делайте это, только если WMS точно не открыта в другом окне " _
        & "или на другом компьютере — иначе две копии начнут проводить операции одновременно.") Then Exit Sub
    res = WmsCore.ActionForceUnlock()
    UiRefresh("«Снять блокировку WMS»: " & FirstPart(res))
End Sub

Sub BtnRegister(Optional oEvent As Variant)
    Dim res As String, bNew As Boolean
    WmsInit()
    If Not UiConfirm("Сделать этот файл рабочей WMS?" & Chr(10) & Chr(10) & ConvertFromURL(gDoc.getURL()) & Chr(10) & Chr(10) _
        & "Делайте это, только если WMS перенесена сюда намеренно. Резервную копию так открывать нельзя.") Then Exit Sub
    If gJDir <> "" And UBound(WmsJournal.ListFiles()) < 0 Then
        bNew = UiConfirm("Рядом с книгой нет журнала WMS_Journal. Начать новый журнал с текущего состояния книги?")
    End If
    res = WmsCore.ActionRegisterHere(bNew)
    UiRefresh("«Сделать рабочим файлом»: " & FirstPart(res))
End Sub

Sub BtnSelfCheck(Optional oEvent As Variant)
    Dim res As String
    WmsInit()
    res = WmsCore.ActionSelfCheck()
    UiRefresh(res)
End Sub

Sub BtnBackup(Optional oEvent As Variant)
    Dim res As String
    WmsInit()
    res = WmsCore.ActionBackupNow()
    UiRefresh(res)
End Sub

' ================================================================ «Выдачи»: buttons for the row under the cursor

' row (0-based) of the cell cursor on «Выдачи»; False with a message when there is no single issue row
Function ActiveIssueRow(ByRef r As Long, ByRef why As String) As Boolean
    Dim sel As Object, ra As Variant, ctl As Object
    WmsInit()
    On Error GoTo EH
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().getName() <> SH_ISSUES Then
        why = "Перейдите на лист «" & SH_ISSUES & "» и поставьте курсор в строку выдачи."
        Exit Function
    End If
    sel = gDoc.getCurrentSelection()
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        why = "Выделите одну строку выдачи."
        Exit Function
    End If
    ra = sel.getRangeAddress()
    If ra.StartRow <> ra.EndRow Then
        why = "Выделите одну строку выдачи (сейчас выделено несколько строк)."
        Exit Function
    End If
    r = ra.StartRow
    If r < 1 Then
        why = "Поставьте курсор в строку выдачи, а не в заголовок."
        Exit Function
    End If
    ActiveIssueRow = True
    Exit Function
EH:
    why = "Поставьте курсор в строку выдачи на листе «" & SH_ISSUES & "»."
End Function

' «Провести»: the result is written into «Контроль» of the row; nothing modal on success
Sub BtnPost(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String
    If Not ActiveIssueRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    res = WmsIssue.IssuePostRow(r)
    gUiLastMsg = res
    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Then UiRefresh("«Провести»: " & res)
    If Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Then UiMessage(Mid(res, InStr(res, ":") + 1))
End Sub

Sub BtnFix(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, sh As Object, sEI As String, sQty As String, sWho As String, sDate As String, k As Long
    If Not ActiveIssueRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    sh = gDoc.Sheets.getByName(SH_ISSUES)
    Select Case WmsIssue.RowKind(r)
    Case "POSTED"
    Case "OPEN"
        UiMessage("Строка не проведена — исправьте значения прямо в строке и нажмите «Провести».")
        Exit Sub
    Case "DELETED"
        UiMessage("Выдача удалена (сторно) — исправлять нечего.")
        Exit Sub
    Case Else
        UiMessage("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    k = CLng(sh.getCellByPosition(IC_NO, r).getValue())
    sEI = sh.getCellByPosition(IC_EI, r).getString()
    sQty = WmsIssue.QtyText(sh.getCellByPosition(IC_QTY, r).getValue())
    sWho = sh.getCellByPosition(IC_WHO, r).getString()
    sDate = Format(sh.getCellByPosition(IC_DATE, r).getValue(), "DD.MM.YYYY")
    If Not FixDialog(k, sEI, sQty, sWho, sDate) Then Exit Sub
    res = WmsIssue.IssueFixRow(r, sEI, sQty, sWho, sDate)
    gUiLastMsg = res
    If Left(res, 3) <> "OK:" Then UiMessage("Исправление не выполнено: " & Mid(res, InStr(res, ":") + 1))
    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Then UiRefresh("«Исправить»: " & res)
End Sub

Sub BtnDelete(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String, sh As Object
    If Not ActiveIssueRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    sh = gDoc.Sheets.getByName(SH_ISSUES)
    Select Case WmsIssue.RowKind(r)
    Case "POSTED"
    Case "OPEN"
        UiMessage("Строка не проведена — чтобы убрать ввод, нажмите «Очистить».")
        Exit Sub
    Case "DELETED"
        UiMessage("Выдача уже удалена (сторно).")
        Exit Sub
    Case Else
        UiMessage("Строка — КОПИЯ, её можно только очистить кнопкой «Очистить».")
        Exit Sub
    End Select
    If Not UiConfirm("Удалить выдачу № " & sh.getCellByPosition(IC_NO, r).getString() & "?" & Chr(10) & Chr(10) _
        & sh.getCellByPosition(IC_EI, r).getString() & " — " & sh.getCellByPosition(IC_NAME, r).getString() & ", " _
        & sh.getCellByPosition(IC_QTY, r).getString() & " " & sh.getCellByPosition(IC_UNIT, r).getString() & Chr(10) & Chr(10) _
        & "Остаток вернётся на склад. Строка останется в истории с отметкой «" & ST_DELETED & "», № не будет использован повторно.") Then Exit Sub
    res = WmsIssue.IssueDeleteRow(r)
    gUiLastMsg = res
    If Left(res, 3) <> "OK:" Then UiMessage("Удаление не выполнено: " & Mid(res, InStr(res, ":") + 1))
    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Then UiRefresh("«Удалить»: " & res)
End Sub

Sub BtnClear(Optional oEvent As Variant)
    Dim r As Long, why As String, res As String
    If Not ActiveIssueRow(r, why) Then
        UiMessage(why)
        Exit Sub
    End If
    res = WmsIssue.IssueClearRow(r)
    gUiLastMsg = res
    If Left(res, 3) <> "OK:" Then UiMessage(Mid(res, InStr(res, ":") + 1))
End Sub

' the correction dialog: EI, quantity, recipient, date prefilled with the posted values
Private Function FixDialog(k As Long, ByRef sEI As String, ByRef sQty As String, ByRef sWho As String, ByRef sDate As String) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, lbl As Variant, nm As Variant, vals As Variant, m As Object
    If gUiAuto <> 0 Then
        If gUiFixSet Then
            sEI = gUiFixEI
            sQty = gUiFixQty
            sWho = gUiFixWho
            sDate = gUiFixDate
            gUiFixSet = False
            FixDialog = True
        End If
        Exit Function
    End If
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = "Исправить выдачу № " & k
    dm.Width = 236
    dm.Height = 122
    lbl = Array("ЕИ (Внутренний код)", "Количество", "Кому выдано", "Дата выдачи (дд.мм.гггг)")
    nm = Array("eEI", "eQty", "eWho", "eDate")
    vals = Array(sEI, sQty, sWho, sDate)
    For i = 0 To 3
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = 11 + i * 18
        m.Width = 94
        m.Height = 12
        m.Label = lbl(i)
        dm.insertByName("l" & i, m)
        m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        m.PositionX = 104
        m.PositionY = 8 + i * 18
        m.Width = 124
        m.Height = 14
        m.Text = vals(i)
        dm.insertByName(nm(i), m)
    Next i
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 104
    m.PositionY = 100
    m.Width = 60
    m.Height = 15
    m.Label = "Исправить"
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 168
    m.PositionY = 100
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        sEI = dlg.getControl("eEI").getText()
        sQty = dlg.getControl("eQty").getText()
        sWho = dlg.getControl("eWho").getText()
        sDate = dlg.getControl("eDate").getText()
        FixDialog = True
    End If
    dlg.dispose()
End Function

' ================================================================ «Главная»: переходы на рабочие листы (M6 PRIME §11)
' The sheet is shown with the cursor where the work starts: the two input cells of «Приход авто», the first free row of a
' sheet of rows (its first input column), the top of «Наличие».

Private Sub GoSheet(sName As String, inCol As Integer)
    Dim sh As Object, r As Long, ctl As Object
    WmsInit()
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(sName) Then Exit Sub
    sh = gDoc.Sheets.getByName(sName)
    ctl = gDoc.getCurrentController()
    ctl.setActiveSheet(sh)
    If sName = SH_CARS Then
        ctl.select(sh.getCellByPosition(CP_PLATE_COL, CP_IN_ROW))
    ElseIf inCol >= 0 Then
        r = WmsOrders.LastRow(sh) + 1
        If r < 1 Then r = 1
        If r > MAX_SHEET_ROW Then r = MAX_SHEET_ROW
        ctl.select(sh.getCellByPosition(inCol, r))
    Else
        ctl.select(sh.getCellByPosition(0, 1))
    End If
    gUiLastMsg = "OK:" & sName
EH:
End Sub

Sub NavOrders(Optional oEvent As Variant)
    GoSheet(SH_ORDERS, OC_ORDER)
End Sub

Sub NavSpecial(Optional oEvent As Variant)
    GoSheet(SH_SPECIAL, XC_TYPE)
End Sub

Sub NavCars(Optional oEvent As Variant)
    GoSheet(SH_CARS, CP_PLATE_COL)
End Sub

Sub NavIssues(Optional oEvent As Variant)
    GoSheet(SH_ISSUES, IC_EI)
End Sub

Sub NavReturns(Optional oEvent As Variant)
    GoSheet(SH_RETURNS, RC_ISSUE)
End Sub

Sub NavAdjust(Optional oEvent As Variant)
    GoSheet(SH_ADJUST, AC_KIND)
End Sub

Sub NavStock(Optional oEvent As Variant)
    GoSheet(SH_STOCK, -1)
End Sub

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestUiAuto(mode As Integer) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestUiAuto = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gUiAuto = mode
    TestUiAuto = "OK"
End Function

Function TestFixInput(sEI As String, sQty As String, sWho As String, sDate As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestFixInput = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gUiFixEI = sEI
    gUiFixQty = sQty
    gUiFixWho = sWho
    gUiFixDate = sDate
    gUiFixSet = True
    TestFixInput = "OK"
End Function

Function TestUiLastMessage() As String
    TestUiLastMessage = gUiLastMsg
End Function

' how many confirmations were asked since the book was opened (an action that is refused must not ask)
Function TestUiConfirmCount() As Long
    TestUiConfirmCount = gUiConfirms
End Function

' ================================================================ saving with an active filter (D-041)
' LibreOffice 24.2 opens a file with many rows hidden by a filter in quadratic time: it recalculates the page breaks for
' every run of hidden rows (40 000 rows with every fifth one shown: 55 s; 100 000: minutes). So before every save WMS
' temporarily unhides the rows of a filtered WMS range (all rows and their data are saved as usual, just not hidden);
' the filter itself (its conditions, kept by the database range) is not removed and is applied again right after the
' save and at every start, so the user keeps the filter. The conditions are never rebuilt through the API (getFilterFields3 does not return what
' setFilterFields3 expects). Nothing here may prevent a save or a start: errors are ignored. The undo manager is
' locked, so no undo entries appear.

' names of the WMS database ranges that currently have filter conditions
Private Function FilteredRanges() As Variant
    Dim names As Variant, i As Integer, out() As String, n As Integer, db As Object
    names = Array("WMS_ORDERS", "WMS_SPECIAL", "WMS_CARS", "WMS_ISSUES", "WMS_RETURNS", "WMS_ADJUST", "WMS_STOCK")
    ReDim out(UBound(names))
    For i = 0 To UBound(names)
        If gDoc.DatabaseRanges.hasByName(names(i)) Then
            db = gDoc.DatabaseRanges.getByName(names(i))
            If UBound(db.getFilterDescriptor().getFilterFields3()) >= 0 Then
                out(n) = names(i)
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then
        FilteredRanges = Array()
    Else
        ReDim Preserve out(n - 1)
        FilteredRanges = out
    End If
End Function

Sub FiltersStash()
    Dim f As Variant, i As Integer, db As Object, a As Object, sh As Object, um As Object, locked As Boolean
    On Error GoTo EH
    gFltN = 0
    f = FilteredRanges()
    um = gDoc.getUndoManager()
    For i = 0 To UBound(f)
        db = gDoc.DatabaseRanges.getByName(f(i))
        a = db.getDataArea()
        sh = gDoc.Sheets.getByIndex(a.Sheet)
        gFltName(gFltN) = f(i)
        gFltN = gFltN + 1
        ' every row of the range, including the empty ones below the data (they are hidden by the filter too)
        If a.EndRow > a.StartRow Then
            um.lock()
            locked = True
            sh.getCellRangeByPosition(0, a.StartRow + 1, 0, a.EndRow).getRows().IsVisible = True
            um.unlock()
            locked = False
        End If
    Next i
    Exit Sub
EH:
    On Error Resume Next
    If locked Then um.unlock()
End Sub

' runs the kept filter conditions of the given ranges again (after a save, at a start); the modified flag is kept
Private Sub FiltersRun(f As Variant)
    Dim i As Integer, um As Object, wasModified As Boolean, locked As Boolean
    On Error GoTo EH
    wasModified = gDoc.isModified()
    um = gDoc.getUndoManager()
    For i = 0 To UBound(f)
        If gDoc.DatabaseRanges.hasByName(f(i)) Then
            um.lock()
            locked = True
            gDoc.DatabaseRanges.getByName(f(i)).refresh()
            um.unlock()
            locked = False
        End If
    Next i
    If Not wasModified Then gDoc.setModified(False)
    Exit Sub
EH:
    On Error Resume Next
    If locked Then um.unlock()
End Sub

Sub FiltersRestore()
    Dim f() As String, i As Integer
    If gFltN = 0 Then Exit Sub
    ReDim f(gFltN - 1)
    For i = 0 To gFltN - 1
        f(i) = gFltName(i)
    Next i
    gFltN = 0
    FiltersRun(f)
End Sub

' a start: the file holds no hidden rows, the filter the user left is shown again
Sub FiltersAtStart()
    On Error Resume Next
    FiltersRun(FilteredRanges())
End Sub
