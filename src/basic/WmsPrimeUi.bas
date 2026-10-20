' WmsPrimeUi — PRIME 0.8: the windows of the new corrections. «Исправить» on «Заказы» offers what can be corrected in the
' row under the cursor (the receipt itself, the position, the transfer, the cancellation); «Исправить ЕИ переноса» and
' «Сверка переноса» on «Наличие»; the button of the sheet «Сверка переноса». Every change goes through the operations
' of WmsReceipt / WmsLegacy (ApplyOperation and the journal); the windows only collect the values.
Option Explicit

' test seam (TEST books only): the answer of the next choice window (0-based; -1 — closed), the values the next form returns
' (FORM_KEEP — the value the form proposed), what the last windows showed
Global gPrChoiceSet As Boolean
Global gPrChoice As Integer
Global gPrFormSet As Boolean
Global gPrFormIn(11) As String
Global gPrShown As String
' UiButtons: the window being shown and the button pressed (Global: the button listener is a call of its own)
Global gWmsBtnDlg As Object
Global gWmsBtnPick As Integer
' 1.0: the test seam of UiForm2 — the next form is answered by its second action button
Global gPrFormNext As Boolean

Public Const FORM_KEEP = "<как предложено>"

' ================================================================ generic windows

' a list of choices; the 0-based index of the chosen line or -1 (closed). In the test mode without a preset answer: -1.
Function UiChoose(sTitle As String, sPrompt As String, opts As Variant) As Integer
    Dim dm As Object, dlg As Object, m As Object, lst As Object, i As Integer, h As Integer, n As Integer
    gPrShown = sTitle & Chr(10) & sPrompt & Chr(10) & Join(opts, Chr(10))
    UiChoose = -1
    If gUiAuto <> 0 Then
        If gPrChoiceSet Then
            UiChoose = gPrChoice
            gPrChoiceSet = False
        End If
        Exit Function
    End If
    n = UBound(opts) + 1
    h = 14 * n + 8
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    dm.Height = 58 + h
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.PositionX = 8
    m.PositionY = 6
    m.Width = 284
    m.Height = 22
    m.MultiLine = True
    m.Label = sPrompt
    dm.insertByName("lPrompt", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    m.PositionX = 8
    m.PositionY = 30
    m.Width = 284
    m.Height = h
    m.StringItemList = opts
    m.SelectedItems = Array(0)
    dm.insertByName("lst", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 168
    m.PositionY = 36 + h
    m.Width = 60
    m.Height = 15
    m.Label = "Выбрать"
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 232
    m.PositionY = 36 + h
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        lst = dlg.getControl("lst")
        UiChoose = lst.getSelectedItemPos()
    End If
    dlg.dispose()
End Function

' 1.0: a question whose buttons name what they do («Выдать 1 шт (это 4 кан)», «Выдать 1 кан», «Не проводить») instead of
' «Да / Нет / Отмена» where «Нет» did something too. codes(i) — what button i means; the result is the code of the button
' pressed, -1 when the window was closed. Test seam: a preset choice (TestPrimeChoice: the index of the button), otherwise
' gUiAuto 1 → the first button, 2 → the second.

Function UiButtons(sTitle As String, sText As String, labels As Variant, codes As Variant) As Integer
    Dim dm As Object, dlg As Object, m As Object, i As Integer, n As Integer, lines As Integer, a As Variant, h As Integer, lis As Object
    gUiLastMsg = sText
    gUiConfirms = gUiConfirms + 1
    gPrShown = sTitle & Chr(10) & sText & Chr(10) & Join(labels, Chr(10))
    n = UBound(labels) + 1
    UiButtons = -1
    If gUiAuto <> 0 Then
        If gPrChoiceSet Then
            If gPrChoice >= 0 And gPrChoice < n Then UiButtons = codes(gPrChoice)
            gPrChoiceSet = False
        ElseIf gUiAuto = 1 Then
            UiButtons = codes(0)
        ElseIf gUiAuto = 2 And n > 1 Then
            UiButtons = codes(1)
        End If
        Exit Function
    End If
    a = Split(sText, Chr(10))
    lines = 0
    For i = 0 To UBound(a)
        lines = lines + 1 + Len(a(i)) \ 62
    Next i
    h = 10 * lines + 4
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    dm.Height = 16 + h + n * 20 + 4
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.PositionX = 10
    m.PositionY = 8
    m.Width = 280
    m.Height = h
    m.MultiLine = True
    m.Label = sText
    dm.insertByName("lText", m)
    For i = 0 To n - 1
        m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        m.PositionX = 10
        m.PositionY = 12 + h + i * 20
        m.Width = 280
        m.Height = 16
        m.Label = labels(i)
        m.PushButtonType = 0
        If i = 0 Then m.DefaultButton = True
        dm.insertByName("b" & i, m)
    Next i
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    lis = CreateUnoListener("WmsBtn_", "com.sun.star.awt.XActionListener")
    For i = 0 To n - 1
        dlg.getControl("b" & i).setActionCommand(CStr(i))
        dlg.getControl("b" & i).addActionListener(lis)
    Next i
    gWmsBtnDlg = dlg
    gWmsBtnPick = -1
    dlg.execute()
    If gWmsBtnPick >= 0 And gWmsBtnPick < n Then UiButtons = codes(gWmsBtnPick)
    gWmsBtnDlg = Nothing
    dlg.dispose()
End Function

Sub WmsBtn_actionPerformed(oEv As Object)
    gWmsBtnPick = CInt(oEv.ActionCommand)
    If Not IsNull(gWmsBtnDlg) Then gWmsBtnDlg.endExecute()
End Sub

Sub WmsBtn_disposing(oEv As Object)
End Sub

' the same with the answers of UiAsk3 (6, 7, 2): sYes, sNo — the two actions, sCancel — not doing it (the window closed too)
Function UiAsk3b(s As String, sYes As String, sNo As String, sCancel As String) As Integer
    Dim k As Integer
    k = UiButtons("WMS", s, Array(sYes, sNo, sCancel), Array(6, 7, 2))
    If k = -1 Then k = 2
    UiAsk3b = k
End Function

' a form of labelled fields: v() — the values proposed (in) and entered (out); lists(i) — "" for a text field or the choices
' «a|b|c» of a drop-down field; sInfo — a line of explanation above the fields. True when «OK» was pressed.
Function UiForm(sTitle As String, sInfo As String, labels As Variant, v() As String, lists As Variant, sOk As String) As Boolean
    Dim dm As Object, dlg As Object, i As Integer, m As Object, n As Integer, top As Integer, infoH As Integer, nm As String
    n = UBound(labels) + 1
    gPrShown = sTitle & Chr(10) & sInfo & Chr(10) & Join(v, Chr(9))
    If gUiAuto <> 0 Then
        If gPrFormSet Then
            For i = 0 To n - 1
                If gPrFormIn(i) <> FORM_KEEP Then v(i) = gPrFormIn(i)
            Next i
            gPrFormSet = False
            UiForm = True
        End If
        Exit Function
    End If
    infoH = 0
    If sInfo <> "" Then infoH = 12 * (1 + Len(sInfo) \ 70) + 4
    top = 8 + infoH
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    dm.Height = top + n * 18 + 30
    If sInfo <> "" Then
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = 6
        m.Width = 284
        m.Height = infoH
        m.MultiLine = True
        m.Label = sInfo
        dm.insertByName("lInfo", m)
    End If
    For i = 0 To n - 1
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = top + 3 + i * 18
        m.Width = 120
        m.Height = 12
        m.Label = labels(i)
        dm.insertByName("l" & i, m)
        If CStr(lists(i)) <> "" Then
            m = dm.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
            m.Dropdown = True
            m.StringItemList = Split(CStr(lists(i)), "|")
        Else
            m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        End If
        m.PositionX = 130
        m.PositionY = top + i * 18
        m.Width = 162
        m.Height = 14
        m.Text = v(i)
        dm.insertByName("e" & i, m)
    Next i
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 168
    m.PositionY = top + n * 18 + 8
    m.Width = 60
    m.Height = 15
    m.Label = sOk
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("bOK", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 232
    m.PositionY = top + n * 18 + 8
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        For i = 0 To n - 1
            v(i) = dlg.getControl("e" & i).getText()
        Next i
        UiForm = True
    End If
    dlg.dispose()
End Function

' 1.0 (З5, К1): a question with the reason — «Отменить заказ?», «Удалить?»: sText and one field «Причина» (sReason — in: the
' proposed one, out: as typed; may stay empty). True — «sOk» pressed. The window is «WMS» like the confirmation it replaces
' (the buttons «yes» / «close» as the tests of the windows know them). In the tests: a form set by TestPrimeForm answers it;
' without one the answer is that of UiConfirm (TestUiAuto 1 — «Да» with the proposed reason, 2 — «Нет»)
Function ReasonAsk(sTitle As String, sText As String, sOk As String, ByRef sReason As String) As Boolean
    Dim dm As Object, dlg As Object, m As Object, a As Variant, i As Integer, lines As Integer, h As Integer
    gUiLastMsg = sText
    gUiConfirms = gUiConfirms + 1
    gPrShown = sTitle & Chr(10) & sText & Chr(10) & sReason
    If gUiAuto <> 0 Then
        If gPrFormSet Then
            If gPrFormIn(0) <> FORM_KEEP Then sReason = Trim(gPrFormIn(0))
            gPrFormSet = False
            ReasonAsk = True
        Else
            ReasonAsk = (gUiAuto = 1)
        End If
        Exit Function
    End If
    a = Split(sText, Chr(10))
    For i = 0 To UBound(a)
        lines = lines + 1 + Len(a(i)) \ 62
    Next i
    h = 10 * lines + 4
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = "WMS"
    dm.Width = 300
    dm.Height = h + 64
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.PositionX = 10
    m.PositionY = 8
    m.Width = 280
    m.Height = h
    m.MultiLine = True
    m.Label = sText
    dm.insertByName("lText", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.PositionX = 10
    m.PositionY = h + 13
    m.Width = 80
    m.Height = 12
    m.Label = "Причина:"
    dm.insertByName("lWhy", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    m.PositionX = 90
    m.PositionY = h + 10
    m.Width = 200
    m.Height = 14
    m.Text = sReason
    dm.insertByName("why", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 130
    m.PositionY = h + 34
    m.Width = 78
    m.Height = 16
    m.Label = sOk
    m.PushButtonType = 1
    m.DefaultButton = True
    dm.insertByName("yes", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 212
    m.PositionY = h + 34
    m.Width = 78
    m.Height = 16
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("close", m)
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    If dlg.execute() = 1 Then
        sReason = Trim(dlg.getControl("why").getText())
        ReasonAsk = True
    End If
    dlg.dispose()
End Function

' 1.0: UiForm with two action buttons (sOk, sOk2 — «Выдать», «Выдать и следующую»): 1 — sOk, 2 — sOk2, 0 — «Отмена» or
' the window closed. The values as for UiForm
Function UiForm2(sTitle As String, sInfo As String, labels As Variant, v() As String, lists As Variant, sOk As String, sOk2 As String) As Integer
    Dim dm As Object, dlg As Object, i As Integer, m As Object, n As Integer, top As Integer, infoH As Integer, lis As Object, r As Integer
    n = UBound(labels) + 1
    gPrShown = sTitle & Chr(10) & sInfo & Chr(10) & Join(v, Chr(9))
    If gUiAuto <> 0 Then
        If gPrFormSet Then
            For i = 0 To n - 1
                If gPrFormIn(i) <> FORM_KEEP Then v(i) = gPrFormIn(i)
            Next i
            gPrFormSet = False
            UiForm2 = IIf(gPrFormNext, 2, 1)
            gPrFormNext = False
        End If
        Exit Function
    End If
    infoH = 0
    If sInfo <> "" Then infoH = 12 * (1 + Len(sInfo) \ 70) + 4
    top = 8 + infoH
    dm = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    dm.Title = sTitle
    dm.Width = 300
    dm.Height = top + n * 18 + 30
    If sInfo <> "" Then
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = 6
        m.Width = 284
        m.Height = infoH
        m.MultiLine = True
        m.Label = sInfo
        dm.insertByName("lInfo", m)
    End If
    For i = 0 To n - 1
        m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        m.PositionX = 8
        m.PositionY = top + 3 + i * 18
        m.Width = 120
        m.Height = 12
        m.Label = labels(i)
        dm.insertByName("l" & i, m)
        If CStr(lists(i)) <> "" Then
            m = dm.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
            m.Dropdown = True
            m.Autocomplete = True
            m.StringItemList = Split(CStr(lists(i)), "|")
        Else
            m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        End If
        m.PositionX = 130
        m.PositionY = top + i * 18
        m.Width = 162
        m.Height = 14
        m.Text = v(i)
        dm.insertByName("e" & i, m)
    Next i
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 8
    m.PositionY = top + n * 18 + 8
    m.Width = 70
    m.Height = 15
    m.Label = sOk
    m.DefaultButton = True
    dm.insertByName("b1", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 82
    m.PositionY = top + n * 18 + 8
    m.Width = 140
    m.Height = 15
    m.Label = sOk2
    dm.insertByName("b2", m)
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = 232
    m.PositionY = top + n * 18 + 8
    m.Width = 60
    m.Height = 15
    m.Label = "Отмена"
    m.PushButtonType = 2
    dm.insertByName("bCancel", m)
    dlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    dlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Null)
    lis = CreateUnoListener("WmsBtn_", "com.sun.star.awt.XActionListener")
    dlg.getControl("b1").setActionCommand("1")
    dlg.getControl("b1").addActionListener(lis)
    dlg.getControl("b2").setActionCommand("2")
    dlg.getControl("b2").addActionListener(lis)
    gWmsBtnDlg = dlg
    gWmsBtnPick = -1
    dlg.execute()
    r = gWmsBtnPick
    If r = 1 Or r = 2 Then
        For i = 0 To n - 1
            v(i) = dlg.getControl("e" & i).getText()
        Next i
        UiForm2 = r
    End If
    gWmsBtnDlg = Nothing
    dlg.dispose()
End Function

Function TestPrimeFormNext() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestPrimeFormNext = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gPrFormNext = True
    TestPrimeFormNext = "OK"
End Function

Private Function Payload(res As String) As String
    Payload = Mid(res, InStr(res, ":") + 1)
End Function

' the result of an operation started from a window: nothing on success, the reason otherwise; a blocked WMS refreshes «Главная»
Private Sub Report(sAction As String, res As String)
    gUiLastMsg = res
    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Then UiRefresh(sAction & ": " & res)
    If Left(res, 5) = "SKIP:" Then
        UiMessage(Payload(res))
    ElseIf Left(res, 3) <> "OK:" Then
        UiMessage(sAction & " не выполнено: " & Payload(res))
    End If
    gUiLastMsg = res
End Sub

' ================================================================ «Заказы»: what «Исправить» can correct in a row

' the corrections offered for row r of kind sKind (codes RCV, POS, LEG, UNC), as «code|text» lines
Function FixChoices(r As Long, sKind As String) As String
    Dim od As Variant, s As String, olid As Long
    Select Case sKind
    Case "RECEIVED", "STORNO"
        olid = gKol
        If sKind = "RECEIVED" Then s = "RCV|Этот приход: количество, документ, даты, место, цена" & Chr(10)
        s = s & "POS|Позицию заказа: наименование, артикул, единица, категория, поставщик, заказано, № заказа, счёт, дата заказа" & Chr(10)
        If sKind = "RECEIVED" Then
            If WmsReceipt.LegacyFixInfo(r) <> "" Then s = s & "LEG|Коррекция переноса: получено по старой таблице и остаток на момент переноса" & Chr(10)
        End If
        If WmsOrders.OlValid(olid, od) Then
            If CStr(od(OD_CANCEL)) = OD_CANCEL_ORDER Then s = s & "UNC|Снять отмену заказа" & Chr(10)
            If CStr(od(OD_CANCEL)) = OD_CANCEL_REST Then s = s & "UNC|Снять отмену остатка (снова ждать недостающее)" & Chr(10)
        End If
    Case "CANCELLED"
        s = "UNC|Снять отмену заказа — строка снова станет заказом, который ждём" & Chr(10)
    End Select
    FixChoices = s
End Function

' «Исправить» → the chosen correction for row r (called by WmsOrdersUi.BtnRcvFix); "RCV" — the receipt window of WmsOrdersUi
' is shown by the caller; otherwise the result of the correction (gUiLastMsg)
Function RcvFixMenu(r As Long, sKind As String) As String
    Dim lines As Variant, opts As Variant, codes As Variant, i As Integer, k As Integer, c As Integer, p As Long, so As String, sc As String
    ' the lines «code|text» → two lists (no ReDim Preserve of a String array: see the traps of LibreOffice Basic)
    lines = Split(FixChoices(r, sKind), Chr(10))
    For i = 0 To UBound(lines)
        If lines(i) <> "" Then
            p = InStr(lines(i), "|")
            sc = sc & IIf(k > 0, Chr(10), "") & Left(lines(i), p - 1)
            so = so & IIf(k > 0, Chr(10), "") & Mid(lines(i), p + 1)
            k = k + 1
        End If
    Next i
    If k = 0 Then
        RcvFixMenu = ""
        Exit Function
    End If
    opts = Split(so, Chr(10))
    codes = Split(sc, Chr(10))
    c = UiChoose("Исправить — строка " & (r + 1), "Что исправить в строке " & (r + 1) & IIf(gKcanon <> "", " (" & gKcanon & ")", "") & "?", opts)
    If c < 0 Or c >= k Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        RcvFixMenu = "SKIP"
        Exit Function
    End If
    Select Case codes(c)
    Case "RCV"
        RcvFixMenu = "RCV"
    Case "POS"
        RcvFixMenu = PositionFixUi(r)
    Case "LEG"
        RcvFixMenu = LegacyFixUi(r)
    Case "UNC"
        RcvFixMenu = UncancelUi(r)
    End Select
End Function

' «Исправить позицию»: the window with the identity of the position (from its source row)
Function PositionFixUi(r As Long) As String
    Dim sh As Object, src As Long, od As Variant, v(8) As String, labels As Variant, lists As Variant, res As String, c As Object
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    src = r
    If Not gKsrc Then
        If WmsOrders.OlValid(gKol, od) Then src = WmsOrders.SourceRowOf(od)
        If src < 1 Then src = r
    End If
    v(0) = CellShown(sh.getCellByPosition(OC_ORDER, src))
    v(1) = sh.getCellByPosition(OC_NAME, src).getString()
    v(2) = CellShown(sh.getCellByPosition(OC_INVOICE, src))
    v(3) = sh.getCellByPosition(OC_ART, src).getString()
    c = sh.getCellByPosition(OC_ORDQTY, src)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then v(4) = WmsIssue.QtyText(c.getValue()) Else v(4) = c.getString()
    v(5) = sh.getCellByPosition(OC_UNIT, src).getString()
    v(6) = sh.getCellByPosition(OC_SUPPLIER, src).getString()
    c = sh.getCellByPosition(OC_ODATE, src)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then v(7) = Format(c.getValue(), "DD.MM.YYYY")
    v(8) = sh.getCellByPosition(OC_CAT, src).getString()
    labels = Array("№ заказа (A)", "Наименование (B)", "№ счёта (D)", "Артикул (E)", "Заказано (H)", "Единица (I)", "Поставщик (L)", _
        "Дата заказа (P), дд.мм.гггг", "Категория (S)")
    lists = Array("", "", "", "", "", "", "", "", "")
    If Not UiForm("Исправить позицию заказа", "Исправление попадёт во все строки позиции и в карточки её ЕИ в «" & SH_STOCK _
        & "». Выдачи и возвраты сохранят то, что было в момент операции. Единица меняется, только пока по ЕИ ничего не выдано.", labels, v, lists, "Исправить") Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        PositionFixUi = gUiLastMsg
        Exit Function
    End If
    res = WmsReceipt.PositionFixRow(r, v(0), v(1), v(2), v(3), v(4), v(5), v(6), v(7), v(8))
    Report("«Исправить позицию»", res)
    PositionFixUi = res
End Function

' «Коррекция переноса»: the window with the numbers of the transfer and the act
Function LegacyFixUi(r As Long) As String
    Dim info As Variant, v(4) As String, labels As Variant, lists As Variant, res As String, sInfo As String
    info = Split(WmsReceipt.LegacyFixInfo(r), "|")
    If UBound(info) < 4 Then
        UiMessage("В строке нет перенесённого прихода.")
        gUiLastMsg = "ERR:в строке нет перенесённого прихода"
        LegacyFixUi = gUiLastMsg
        Exit Function
    End If
    sInfo = gKcanon & ": по переносу получено " & info(0) & ", остаток на момент переноса " & info(1) & ", израсходовано до переноса " & info(2) _
        & IIf(info(3) <> "", ", в старой таблице X = «" & info(3) & "»", "") & "; остаток сейчас " & info(4) _
        & ". Остаток сейчас изменится на столько же, на сколько изменится остаток на момент переноса."
    v(0) = info(0)
    v(1) = info(1)
    v(3) = Format(Now(), "DD.MM.YYYY")
    labels = Array("Получено по старой таблице (F)", "Остаток на момент переноса", "№ акта / документа", "Дата акта, дд.мм.гггг", "Причина")
    lists = Array("", "", "", "", "")
    If Not UiForm("Коррекция переноса " & gKcanon, sInfo, labels, v, lists, "Провести") Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        LegacyFixUi = gUiLastMsg
        Exit Function
    End If
    res = WmsReceipt.LegacyFixRow(r, v(0), v(1), v(2), v(3), v(4))
    Report("«Коррекция переноса»", res)
    LegacyFixUi = res
End Function

' «Снять отмену»: asked once, then ORDER_UNCANCEL
Function UncancelUi(r As Long) As String
    Dim res As String
    If Not UiConfirm("Снять отмену в строке " & (r + 1) & "? Позиция снова станет ожидаемой; история отмены останется в журнале.") Then
        gUiLastMsg = "SKIP:снятие отмены не подтверждено"
        UncancelUi = gUiLastMsg
        Exit Function
    End If
    res = WmsReceipt.OrderUncancelRow(r)
    Report("«Снять отмену»", res)
    UncancelUi = res
End Function

' the text of a cell as typed (a number without a thousands separator)
Private Function CellShown(c As Object) As String
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then CellShown = NumStr(c.getValue()) Else CellShown = c.getString()
End Function

' ================================================================ «Наличие»

' the EI of the row under the cursor of «Наличие» (row = EI number), or ""
Private Function ActiveStockEI() As String
    Dim sel As Object, a As Object, sh As Object
    sel = gDoc.getCurrentController().getSelection()
    If Not sel.supportsService("com.sun.star.sheet.SheetCellRange") Then Exit Function
    a = sel.getRangeAddress()
    sh = gDoc.Sheets.getByIndex(a.Sheet)
    If sh.Name <> SH_STOCK Or a.StartRow < 1 Then Exit Function
    ActiveStockEI = sh.getCellByPosition(SC_EI, a.StartRow).getString()
End Function

' «Исправить ЕИ переноса»: the card and the balance at the transfer of a transferred EI without an order receipt
Sub BtnMigrateFix(Optional oEvent As Variant)
    Dim sEI As String, info As Variant, v(8) As String, labels As Variant, lists As Variant, res As String, sInfo As String, lg As Variant
    Dim n As Long, canon As String, msg As String
    WmsInit()
    sEI = ActiveStockEI()
    If sEI = "" Then
        UiMessage("Поставьте курсор в строку ЕИ на листе «" & SH_STOCK & "».")
        gUiLastMsg = "ERR:не выбран ЕИ"
        Exit Sub
    End If
    info = Split(WmsLegacy.MigrateFixInfo(sEI), "|")
    If UBound(info) < 8 Then
        msg = sEI & " перенесён не как ЕИ без заказа."
        If WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
            If WmsLegacy.LegRow(n, lg) Then
                msg = sEI & " перенесён приходом по заказу: его исправляет лист «" & SH_ORDERS & "» → строка прихода → «Исправить» → «Коррекция переноса»."
            Else
                msg = sEI & " не из переноса старой таблицы. Карточку исправляет приход, которым создан ЕИ («Заказы» или «Иной приход» → «Исправить»), " _
                    & "остаток — «Корректировки» → «Инвентаризация»."
            End If
        End If
        UiMessage(msg)
        gUiLastMsg = "ERR:" & msg
        Exit Sub
    End If
    sInfo = sEI & ": остаток сейчас " & info(6) & ", место «" & info(8) & "», состояние «" & info(7) & "»" _
        & IIf(info(7) = EI_ST_REVIEW, " — выберите тип источника, чтобы разобрать ЕИ", "") _
        & ". Остаток на момент переноса меняется только с № и датой акта."
    v(0) = info(0)
    v(1) = info(1)
    v(2) = info(2)
    v(3) = info(3)
    v(4) = info(4)
    v(5) = info(5)
    labels = Array("Наименование", "Артикул", "Единица", "Категория", "Тип источника", "Остаток на момент переноса", "№ акта / документа", _
        "Дата акта, дд.мм.гггг", "Причина")
    lists = Array("", "", "", "", STYPE_SUPPLIER & "|" & SRC_STYPES, "", "", "", "")
    If Not UiForm("Исправить ЕИ переноса " & sEI, sInfo, labels, v, lists, "Исправить") Then
        gUiLastMsg = "SKIP:окно закрыто без исправления"
        Exit Sub
    End If
    ' 1.0 (Н6): a restore point before the correction of the transfer
    WmsSafe.RestorePoint("перед коррекцией переноса " & sEI)
    res = WmsLegacy.MigrateFixEI(sEI, v(0), v(1), v(2), v(3), v(4), v(5), v(6), v(7), v(8))
    Report("«Исправить ЕИ переноса»", res)
End Sub

' «Сверка переноса» / «Обновить сверку»: the report is rebuilt and shown
Sub BtnLegacyReport(Optional oEvent As Variant)
    Dim res As String, sh As Object
    WmsInit()
    res = WmsLegacy.LegacyReport()
    gUiLastMsg = res
    sh = gDoc.Sheets.getByName(SH_LEGREP)
    gDoc.getCurrentController().setActiveSheet(sh)
    If Left(res, 3) = "OK:" Then
        If gUiAuto = 0 Then MsgBox "Сверка переноса обновлена: " & Payload(res) & "." & Chr(10) & Chr(10) _
            & "Сначала идут строки, где остаток не сходится с движениями, затем — где остаток при переносе стоит проверить по полке.", 64, "WMS"
    Else
        UiMessage("Сверка не построена: " & Payload(res))
    End If
End Sub

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestPrimeChoice(k As Integer) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestPrimeChoice = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gPrChoice = k
    gPrChoiceSet = True
    TestPrimeChoice = "OK"
End Function

' the values of the next form, separated by Tab (FORM_KEEP keeps the proposed value)
Function TestPrimeForm(sValues As String) As String
    Dim a As Variant, i As Integer
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestPrimeForm = "REFUSED:не тестовая книга"
        Exit Function
    End If
    a = Split(sValues, Chr(9))
    For i = 0 To 11
        If i <= UBound(a) Then gPrFormIn(i) = a(i) Else gPrFormIn(i) = FORM_KEEP
    Next i
    gPrFormSet = True
    TestPrimeForm = "OK"
End Function

Function TestPrimeShown() As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestPrimeShown = "REFUSED:не тестовая книга"
        Exit Function
    End If
    TestPrimeShown = gPrShown
End Function
