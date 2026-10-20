' WmsDaily — 1.0: the everyday work made shorter (the choice of the storekeeper, stage B): several issues posted at once
' under one document (Р1), and the helpers of the next steps. Every row stays its own operation of the core (ISSUE): the
' block only chooses the rows, the common document and remembers an answer given once for the rest of the block.
' Library WmsPlus.
Option Explicit

' the last block posted (for the tests and for «Накладная»): its document and the numbers of its issues
Global gDyDoc As String
Global gDyNos As String

' ================================================================ Р1 «Провести» for several selected rows of «Выдачи»

' the rows (0-based) of a selection of several rows on «Выдачи»; False — not a block (one row, another sheet, several ranges)
Function IssueSelection(ByRef r0 As Long, ByRef r1 As Long) As Boolean
    Dim sel As Object, ra As Variant, ctl As Object
    On Error GoTo EH
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().getName() <> SH_ISSUES Then Exit Function
    sel = gDoc.getCurrentSelection()
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then Exit Function
    ra = sel.getRangeAddress()
    If ra.EndRow <= ra.StartRow Then Exit Function
    r0 = IIf(ra.StartRow < 1, 1, ra.StartRow)
    r1 = ra.EndRow
    If r1 - r0 > 2000 Then r1 = r0 + 2000
    IssueSelection = (r1 > r0)
    Exit Function
EH:
End Function

' the rows of r0..r1 that can be posted: not posted, the EI (L) and a quantity (E or F) typed
Private Function OpenRows(r0 As Long, r1 As Long) As Variant
    Dim sh As Object, d As Variant, i As Long, n As Long, rows() As Long
    sh = WmsIssue.IssuesSheet()
    d = sh.getCellRangeByPosition(0, r0, IC_LAST, r1).getDataArray()
    ReDim rows(r1 - r0)
    n = 0
    For i = 0 To r1 - r0
        If CStr(d(i)(IC_NO)) = "" And Trim(CStr(d(i)(IC_EI))) <> "" And (CStr(d(i)(IC_QTY)) <> "" Or CStr(d(i)(IC_PCT)) <> "") Then
            rows(n) = r0 + i
            n = n + 1
        End If
    Next i
    If n = 0 Then
        OpenRows = Array()
        Exit Function
    End If
    ReDim Preserve rows(n - 1)
    OpenRows = rows
End Function

' the common document of the block: the first «Документ» (B) typed in its rows, otherwise «ВД-<№ of the first issue>»
Private Function BlockDoc(rows As Variant) As String
    Dim sh As Object, i As Long, s As String
    sh = WmsIssue.IssuesSheet()
    For i = 0 To UBound(rows)
        s = Trim(sh.getCellByPosition(IC_DOC, rows(i)).getString())
        If s <> "" Then
            BlockDoc = s
            Exit Function
        End If
    Next i
    BlockDoc = "ВД-" & Format(CLng(SysNum(SK_NEXT_NO)), "000000")
End Function

' «Провести» with several rows selected: one question (one document / each apart / not now), then every row by the
' ordinary posting; a question answered for one row (an unknown recipient, an unusual date) is not asked again in the block.
' The summary is shown only when something was not posted. The result: «OK=…;ERR=…;SKIP=…» and the lines of the rows
Function PostIssueBlock(r0 As Long, r1 As Long) As String
    Dim rows As Variant, n As Long, i As Long, k As Integer, sDoc As String, res As String, nOk As Long, nErr As Long, nSkip As Long
    Dim lines As String, bStop As Boolean, sh As Object
    WmsInit()
    rows = OpenRows(r0, r1)
    n = UBound(rows) + 1
    If n = 0 Then
        WmsUi.UiMessage("В выделенных строках нечего проводить: нет непроведённых строк с ЕИ (L) и количеством (E или F).")
        PostIssueBlock = "SKIP:нечего проводить"
        gUiLastMsg = PostIssueBlock
        Exit Function
    End If
    sDoc = BlockDoc(rows)
    If n >= 2 Then
        k = WmsPrimeUi.UiButtons("WMS — провести выдачи", "Выделено непроведённых выдач: " & n & "." & Chr(10) & Chr(10) _
            & "Одним документом — у всех строк в «Документ» (B) будет «" & sDoc & "» (где B пусто); по нему можно напечатать накладную." _
            & Chr(10) & "Каждая выдача остаётся своей операцией со своим №.", _
            Array("Провести " & n & " выдач одним документом «" & sDoc & "»", "Провести каждую отдельно", "Не проводить"), Array(0, 1, 2))
        If k = 2 Or k = -1 Then
            PostIssueBlock = "SKIP:проведение отменено"
            gUiLastMsg = PostIssueBlock
            Exit Function
        End If
        If k = 1 Then sDoc = ""
    Else
        sDoc = ""
    End If
    sh = WmsIssue.IssuesSheet()
    gDyDoc = sDoc
    gDyNos = ""
    WmsAsk.AskMemoBegin()
    For i = 0 To n - 1
        res = WmsIssue.IssuePostRow(rows(i), sDoc)
        If Left(res, 3) = "OK:" Then
            nOk = nOk + 1
            gDyNos = gDyNos & IIf(gDyNos <> "", ",", "") & CStr(CLng(sh.getCellByPosition(IC_NO, rows(i)).getValue()))
        ElseIf Left(res, 5) = "SKIP:" Then
            nSkip = nSkip + 1
            lines = lines & Chr(10) & "строка " & (rows(i) + 1) & ": " & Mid(res, 6)
        Else
            nErr = nErr + 1
            lines = lines & Chr(10) & "строка " & (rows(i) + 1) & ": " & Mid(res, InStr(res, ":") + 1)
            ' a stop of WMS itself ends the block: the next rows would be refused the same way
            If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 7) = "ERR-SYS" Then bStop = True
        End If
        If bStop Then Exit For
    Next i
    WmsAsk.AskMemoEnd()
    PostIssueBlock = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & lines
    If nErr + nSkip > 0 Then
        WmsUi.UiMessage("Проведено выдач: " & nOk & " из " & n & IIf(sDoc <> "" And nOk > 0, " (документ «" & sDoc & "»)", "") & "." _
            & IIf(bStop, Chr(10) & "Проведение остановлено: WMS не может проводить дальше — см. лист «" & SH_MAIN & "».", "") & Chr(10) _
            & "Не проведены:" & lines)
    End If
    gUiLastMsg = PostIssueBlock
    If bStop Then WmsUi.UiRefresh("«Провести»: " & Mid(lines, 2))
End Function

' ================================================================ Р3 «Повторить выдачу», Р9 «Выдать» from «Наличие» and the card

' the first row of «Выдачи» below every used one (a new issue goes there)
Function FreeIssueRow() As Long
    FreeIssueRow = WmsOrders.LastRow(WmsIssue.IssuesSheet()) + 1
    If FreeIssueRow < 1 Then FreeIssueRow = 1
End Function

' a new row of «Выдачи» prepared, not posted: the EI, the quantity as the storekeeper would type it ("" — none), the
' recipient, the date today; «Контроль» shows what «Провести» will do; the cursor stops on cCursor
Private Function PrepareIssue(sEI As String, sQty As String, sWho As String, cCursor As Integer) As Long
    Dim sh As Object, r As Long
    sh = WmsIssue.IssuesSheet()
    r = FreeIssueRow()
    sh.getCellByPosition(IC_EI, r).setString(sEI)
    If sQty <> "" Then sh.getCellByPosition(IC_QTY, r).setString(sQty)
    If sWho <> "" Then sh.getCellByPosition(IC_WHO, r).setString(sWho)
    sh.getCellByPosition(IC_DATE, r).setValue(CDbl(Date()))
    WmsIssue.PreviewRow(r, True, sWho <> "")
    WmsView.GoCell(SH_ISSUES, r, cCursor)
    PrepareIssue = r
End Function

Sub BtnRepeatIssue(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = RepeatIssue()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Повторить»: the issue under the cursor — a new row with the same EI, quantity (in its parts when the EI has a packing)
' and recipient, dated today, not posted: the storekeeper checks it and presses «Провести»
Function RepeatIssue() As String
    Dim r As Long, why As String, sh As Object, sEI As String, n As Long, canon As String, msg As String, sQty As String, r2 As Long
    WmsInit()
    If Not WmsUi.ActiveIssueRow(r, why) Then
        RepeatIssue = "ERR:" & why
        gUiLastMsg = RepeatIssue
        Exit Function
    End If
    If WmsIssue.RowKind(r) <> "POSTED" Then
        RepeatIssue = "ERR:встаньте на проведённую выдачу — «Повторить» делает такую же новую строку с сегодняшней датой"
        gUiLastMsg = RepeatIssue
        Exit Function
    End If
    sh = WmsIssue.IssuesSheet()
    sEI = sh.getCellByPosition(IC_EI, r).getString()
    If WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        sQty = WmsPack.QtyForInput(n, sh.getCellByPosition(IC_QTY, r).getValue())
    Else
        sQty = WmsIssue.QtyText(sh.getCellByPosition(IC_QTY, r).getValue())
    End If
    r2 = PrepareIssue(sEI, sQty, sh.getCellByPosition(IC_WHO, r).getString(), IC_QTY)
    RepeatIssue = "OK:строка " & (r2 + 1) & ": " & sEI & ", " & sQty & ", " & sh.getCellByPosition(IC_WHO, r).getString() & " — проверьте и нажмите «Провести»"
    gUiLastMsg = RepeatIssue
End Function

Sub BtnIssueFrom(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = IssueFrom()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Выдать» on «Наличие» (the EI of the row under the cursor) or on «Поиск» (the EI of the card shown): a new row of «Выдачи»
' with this EI, the cursor on «Количество»
Function IssueFrom() As String
    Dim sName As String, r As Long, sEI As String, n As Long, b As Double, r2 As Long
    WmsInit()
    sName = gDoc.getCurrentController().getActiveSheet().Name
    If sName = SH_STOCK Then
        r = WmsView.ViewRow(SH_STOCK)
        If r >= 1 Then sEI = WmsIssue.StockSheet().getCellByPosition(SC_EI, r).getString()
    ElseIf sName = SH_FIND Then
        sEI = WmsFind.FindCardEI()
    End If
    If sEI = "" Then
        IssueFrom = "ERR:встаньте на строку ЕИ в «" & SH_STOCK & "» или откройте карточку ЕИ на «" & SH_FIND & "» — затем «Выдать»"
        gUiLastMsg = IssueFrom
        Exit Function
    End If
    n = CLng(Val(Mid(sEI, Len(EI_PREFIX) + 1)))
    If WmsIssue.StockQty(n, b) Then
        If b <= 0.0000005 Then
            IssueFrom = "ERR:остаток " & sEI & " — 0, выдавать нечего"
            gUiLastMsg = IssueFrom
            Exit Function
        End If
    End If
    r2 = PrepareIssue(sEI, "", "", IC_QTY)
    IssueFrom = "OK:строка " & (r2 + 1) & " «" & SH_ISSUES & "»: " & sEI & " — впишите количество и получателя"
    gUiLastMsg = IssueFrom
End Function

' ================================================================ Р2 the window «Выдать»

' the recipient «Выдать и следующую» keeps for the next issue (Global: every button is a macro call of its own)
Global gDyWho As String

Sub BtnIssueWindow(Optional oEvent As Variant)
    WmsInit()
    IssueWindow()
End Sub

' the names of «Получатели» for the list of the window (at most 400)
Private Function RecipientList() As String
    Dim i As Long, s As String
    If Not gRcLoaded Then WmsIssue.RecipientsCount()
    For i = 0 To gRcN - 1
        If i >= 400 Then Exit For
        If gRcName(i) <> "" Then s = s & IIf(s <> "", "|", "") & Replace(gRcName(i), "|", "/")
    Next i
    RecipientList = s
End Function

' the EI of a text of the window: a number or a scan («427», «ЕИ-00000427») — that EI; anything else — a part of a name, an
' article or a place: one EI with a balance — it, several — chosen in a list. "" and msg when none
Function ResolveEIText(s As String, ByRef msg As String) As String
    Dim n As Long, canon As String, c As String, lines As Variant, opts() As String, k As Integer, i As Integer
    If s = "" Then
        msg = "впишите ЕИ: номер, скан этикетки или часть наименования"
        Exit Function
    End If
    If WmsIssue.NormalizeEI(s, n, canon, msg) Then
        ResolveEIText = canon
        Exit Function
    End If
    c = WmsFind.PickCandidates(s)
    If Left(c, 4) = "ERR:" Then
        msg = Mid(c, 5)
        Exit Function
    End If
    lines = Split(c, Chr(10))
    If UBound(lines) = 0 Then
        ResolveEIText = Left(lines(0), InStr(lines(0), "|") - 1)
        Exit Function
    End If
    ReDim opts(UBound(lines))
    For i = 0 To UBound(lines)
        opts(i) = Mid(lines(i), InStr(lines(i), "|") + 1)
    Next i
    k = WmsPrimeUi.UiChoose("Выдать — «" & s & "»", "Найдено несколько ЕИ с остатком (старые первыми)" _
        & IIf(gPickAll > UBound(lines) + 1, "; показано " & (UBound(lines) + 1) & " из " & gPickAll, "") & ". Выберите:", opts)
    If k < 0 Or k > UBound(lines) Then
        msg = "ЕИ не выбран"
        Exit Function
    End If
    ResolveEIText = Left(lines(k), InStr(lines(k), "|") - 1)
End Function

' «Выдать»: one window instead of the columns — the EI (a number, a scan or a name), the quantity (the parts, «всё»), the
' recipient from the list, the date (today). The issue goes into a new row of «Выдачи» and is posted at once; a refusal
' brings the window back with the values and the reason (the row is not left half-done). «Выдать и следующую» keeps the
' recipient and asks for the next one
Function IssueWindow() As String
    Dim v(3) As String, lab As Variant, lists As Variant, k As Integer, sInfo As String, res As String, sEI As String, r As Long
    Dim msg As String, sh As Object, d As Double, nDone As Long, sLast As String
    WmsInit()
    lab = Array("ЕИ (номер, скан или название)", "Количество", "Кому", "Дата")
    lists = Array("", "", RecipientList(), "")
    v(2) = gDyWho
    v(3) = Format(Date(), "DD.MM.YYYY")
    sInfo = "Количество — число в единице документа, «1 кан», «1,5 м», «25%» или «всё» (весь остаток). Получателя можно выбрать из списка."
    sh = WmsIssue.IssuesSheet()
    Do
        k = WmsPrimeUi.UiForm2("WMS — выдать", sInfo, lab, v, lists, "Выдать", "Выдать и следующую")
        If k = 0 Then Exit Do
        sEI = ResolveEIText(Trim(v(0)), msg)
        If sEI = "" Then
            sInfo = "Не выдано: " & msg
        ElseIf Trim(v(3)) <> "" And Not WmsIssue.ParseDateText(Trim(v(3)), d, msg) Then
            sInfo = "Не выдано: дата — " & msg
        Else
            r = FreeIssueRow()
            sh.getCellByPosition(IC_EI, r).setString(sEI)
            sh.getCellByPosition(IC_QTY, r).setString(Trim(v(1)))
            sh.getCellByPosition(IC_WHO, r).setString(Trim(v(2)))
            If Trim(v(3)) = "" Then d = CDbl(Date())
            sh.getCellByPosition(IC_DATE, r).setValue(d)
            res = WmsIssue.IssuePostRow(r)
            If Left(res, 3) = "OK:" Then
                ' 1.0 (М11): an EI on several places — did a place become empty?
                PlaceAfterIssue(r)
                nDone = nDone + 1
                sLast = "№ " & CLng(sh.getCellByPosition(IC_NO, r).getValue()) & " — " & sEI & ", " _
                    & WmsPack.QtyWithParts(CLng(Val(Mid(sEI, Len(EI_PREFIX) + 1))), sh.getCellByPosition(IC_QTY, r).getValue(), sh.getCellByPosition(IC_UNIT, r).getString()) _
                    & " → " & sh.getCellByPosition(IC_WHO, r).getString()
                gDyWho = sh.getCellByPosition(IC_WHO, r).getString()
                If k = 1 Then Exit Do
                v(0) = ""
                v(1) = ""
                v(2) = gDyWho
                sInfo = "Проведено: " & sLast & ". Следующая выдача — тому же получателю (его можно сменить):"
            Else
                ' not posted: the row is cleared, the values and the reason go back into the window
                WmsIssue.IssueClearRow(r)
                If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Then
                    WmsUi.UiMessage("Выдача не проведена: " & Mid(res, InStr(res, ":") + 1))
                    Exit Do
                End If
                sInfo = "Не выдано: " & Mid(res, InStr(res, ":") + 1)
            End If
        End If
    Loop
    IssueWindow = "OK:выдано " & nDone & IIf(sLast <> "", "; последняя " & sLast, "")
    If k = 0 And nDone = 0 Then IssueWindow = "SKIP:окно закрыто" & IIf(sInfo <> "" And Left(sInfo, 10) = "Не выдано:", " — " & sInfo, "")
    gUiLastMsg = IssueWindow
End Function

' ================================================================ test seams

Function TestDyWho() As String
    TestDyWho = gDyWho
End Function

' ================================================================ Р6 «Принять», Р7 «Принять всё по заказу» («Заказы»)

' a cell kept to be put back (a window that did not post leaves the row as it was)
Private Function CellKeep(c As Object) As Variant
    Select Case c.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        CellKeep = Array(0, "")
    Case com.sun.star.table.CellContentType.VALUE
        CellKeep = Array(1, c.getValue())
    Case Else
        CellKeep = Array(2, c.getString())
    End Select
End Function

Private Sub CellBack(c As Object, k As Variant)
    If k(0) = 0 Then
        c.clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING)
    ElseIf k(0) = 1 Then
        c.setValue(k(1))
    Else
        c.setString(k(1))
    End If
End Sub

' a date of a window into a date cell; "" leaves the cell empty. False and msg when it is not a date
Private Function PutDate(c As Object, s As String, sField As String, ByRef msg As String) As Boolean
    Dim d As Double
    If Trim(s) = "" Then
        c.clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING)
        PutDate = True
        Exit Function
    End If
    If Not WmsOrders.ParseDateTextAs(Trim(s), sField, d, msg) Then Exit Function
    c.setValue(d)
    PutDate = True
End Function

Private Sub PutText(c As Object, s As String)
    If Trim(s) = "" Then
        c.clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING)
    Else
        c.setString(Trim(s))
    End If
End Sub

' the first place «Куда положить?» proposes for the goods of the row ("" — none)
Private Function FirstSuggestion(sName As String, sArt As String) As String
    Dim s As String
    s = WmsMap.WhereSuggest(sName, sArt)
    If Left(s, 4) = "ERR:" Or s = "" Then Exit Function
    s = Split(s, Chr(10))(0)
    FirstSuggestion = Left(s, InStr(s, "|") - 1)
End Function

Sub BtnReceiveWindow(Optional oEvent As Variant)
    WmsInit()
    ReceiveWindow()
End Sub

' «Принять» on a position not received yet: one window — accepted (the ordered quantity proposed), the document and its
' date, the date of the receipt (today), the place («Куда положить?» proposes one), the price; then the ordinary
' «Провести приход» of the row. Not posted — the row is put back as it was, the window comes back with the reason. On a
' position already received it is «Ещё поступление»
Function ReceiveWindow() As String
    Dim r As Long, why As String, kind As String, sh As Object, v(5) As String, lab As Variant, sInfo As String, keep(5) As Variant, cols As Variant
    Dim i As Integer, res As String, msg As String, ok As Boolean, sName As String
    WmsInit()
    If Not WmsOrdersUi.ActiveOrderRow(r, why) Then
        WmsUi.UiMessage(why)
        ReceiveWindow = "ERR:" & why
        gUiLastMsg = ReceiveWindow
        Exit Function
    End If
    kind = WmsOrders.OrderRowKind(r)
    If kind = "RECEIVED" Or kind = "STORNO" Then
        WmsOrdersUi.BtnRcvMore()
        ReceiveWindow = gUiLastMsg
        Exit Function
    End If
    If kind <> "OPEN" Then
        why = IIf(kind = "CANCELLED", "позиция отменена — приход по ней не проводится", "поставьте курсор в строку позиции заказа (наименование, количество)")
        WmsUi.UiMessage("«Принять»: " & why & ".")
        ReceiveWindow = "ERR:" & why
        gUiLastMsg = ReceiveWindow
        Exit Function
    End If
    sh = WmsOrders.OrdersSheet()
    sName = sh.getCellByPosition(OC_NAME, r).getString()
    cols = Array(OC_FACT, OC_DOC, OC_DDATE, OC_RDATE, OC_PLACE, OC_PRICE)
    For i = 0 To 5
        keep(i) = CellKeep(sh.getCellByPosition(cols(i), r))
    Next i
    If sh.getCellByPosition(OC_FACT, r).getString() <> "" Then
        v(0) = sh.getCellByPosition(OC_FACT, r).getString()
    ElseIf OrderedQty(r) > 0 Then
        v(0) = WmsIssue.QtyText(OrderedQty(r))
    End If
    v(1) = sh.getCellByPosition(OC_DOC, r).getString()
    If sh.getCellByPosition(OC_DDATE, r).getType() = com.sun.star.table.CellContentType.VALUE Then v(2) = Format(sh.getCellByPosition(OC_DDATE, r).getValue(), "DD.MM.YYYY")
    v(3) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    v(4) = sh.getCellByPosition(OC_PLACE, r).getString()
    If v(4) = "" Then v(4) = FirstSuggestion(sName, sh.getCellByPosition(OC_ART, r).getString())
    v(5) = sh.getCellByPosition(OC_PRICE, r).getString()
    lab = Array("Принято (F)", "№ документа (C)", "Дата документа (O)", "Дата поступления (N)", "Место (U)", "Цена (J)")
    sInfo = sName & " — заказано " & sh.getCellByPosition(OC_ORDQTY, r).getString() & " " & sh.getCellByPosition(OC_UNIT, r).getString() _
        & IIf(sh.getCellByPosition(OC_SUPPLIER, r).getString() <> "", ", " & sh.getCellByPosition(OC_SUPPLIER, r).getString(), "") _
        & ". Место предложено «Куда положить?» — его можно заменить."
    Do
        If Not WmsPrimeUi.UiForm("WMS — принять", sInfo, lab, v, Array("", "", "", "", "", ""), "Принять") Then
            ReceiveWindow = "SKIP:окно закрыто" & IIf(Left(sInfo, 11) = "Не принято:", " — " & sInfo, "")
            Exit Do
        End If
        ok = True
        PutText(sh.getCellByPosition(OC_FACT, r), v(0))
        PutText(sh.getCellByPosition(OC_DOC, r), v(1))
        If Not PutDate(sh.getCellByPosition(OC_DDATE, r), v(2), "дата документа", msg) Then ok = False
        If ok Then
            If Not PutDate(sh.getCellByPosition(OC_RDATE, r), v(3), "дата поступления", msg) Then ok = False
        End If
        PutText(sh.getCellByPosition(OC_PLACE, r), v(4))
        PutText(sh.getCellByPosition(OC_PRICE, r), v(5))
        If ok Then
            res = WmsReceipt.ReceiptPostRow(r)
        Else
            res = "ERR:" & msg
        End If
        If Left(res, 3) = "OK:" Then
            ReceiveWindow = res
            If gRcvNote <> "" Then WmsUi.UiMessage(gRcvNote)
            WmsItem.LabelAfterReceipt(sh.getCellByPosition(OC_EI, r).getString())
            Exit Do
        End If
        For i = 0 To 5
            CellBack(sh.getCellByPosition(cols(i), r), keep(i))
        Next i
        WmsOrders.PreviewOrderRow(r, False)
        If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then
            WmsUi.UiMessage("Приход не проведён: " & Mid(res, InStr(res, ":") + 1))
            ReceiveWindow = res
            Exit Do
        End If
        sInfo = "Не принято: " & Mid(res, InStr(res, ":") + 1)
    Loop
    gUiLastMsg = ReceiveWindow
End Function

' the ordered quantity (H) of row r of «Заказы» (a text column: read as WMS reads a quantity); 0 — none or not a number
Private Function OrderedQty(r As Long) As Double
    Dim q As Double, msg As String, c As Object
    c = WmsOrders.OrdersSheet().getCellByPosition(OC_ORDQTY, r)
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then Exit Function
    If WmsOrders.ParseQtyAs(c, "заказано (H)", q, msg) Then OrderedQty = q
End Function

' the rows of the order of row r: up to the first position (A = 1) above, down to the row before the next first position
Sub OrderBlockOf(r As Long, ByRef r0 As Long, ByRef r1 As Long)
    Dim sh As Object, last As Long, a As String
    sh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(sh)
    r0 = r
    Do While r0 > 1
        a = Trim(sh.getCellByPosition(OC_ORDER, r0).getString())
        If a = "1" Or a = "" Then Exit Do
        r0 = r0 - 1
    Loop
    r1 = r
    Do While r1 < last
        a = Trim(sh.getCellByPosition(OC_ORDER, r1 + 1).getString())
        If a = "1" Or Not WmsOrders.RowHasInput(r1 + 1) Then Exit Do
        r1 = r1 + 1
    Loop
End Sub

Sub BtnReceiveAll(Optional oEvent As Variant)
    WmsInit()
    ReceiveAll()
End Sub

' «Принять всё по заказу»: the selected positions (one row — its whole order): what is still awaited of each — the ordered
' quantity for a position not received yet, the rest (ordered − received) for a position received in part — is shown,
' confirmed once with the document and the dates, and posted position by position (a first receipt — «Провести приход» of
' the row, a rest — «Ещё поступление»). The summary names what was not posted
Function ReceiveAll() As String
    Dim blocks As Variant, why As String, b As Long, r As Long, r0 As Long, r1 As Long, sh As Object, kind As String, od As Variant
    Dim rows() As Long, how() As String, qty() As Double, n As Long, sInfo As String, v(2) As String, i As Long, res As String, msg As String
    Dim nOk As Long, lines As String, ordq As Double, rest As Double, keep(5) As Variant, cols As Variant, j As Integer, ok As Boolean
    WmsInit()
    If Not WmsOrdersUi.ActiveOrderBlocks(blocks, why) Then
        WmsUi.UiMessage(why)
        ReceiveAll = "ERR:" & why
        gUiLastMsg = ReceiveAll
        Exit Function
    End If
    sh = WmsOrders.OrdersSheet()
    If UBound(blocks) = 0 And blocks(0)(0) = blocks(0)(1) Then
        OrderBlockOf(blocks(0)(0), r0, r1)
        blocks = Array(Array(r0, r1))
    End If
    ReDim rows(200)
    ReDim how(200)
    ReDim qty(200)
    For b = 0 To UBound(blocks)
        For r = blocks(b)(0) To blocks(b)(1)
            If n > 200 Then Exit For
            kind = WmsOrders.OrderRowKind(r)
            If kind = "OPEN" Then
                If sh.getCellByPosition(OC_FACT, r).getType() = com.sun.star.table.CellContentType.EMPTY Then
                    ordq = OrderedQty(r)
                    If ordq > 0 Then
                        rows(n) = r
                        how(n) = "FIRST"
                        qty(n) = ordq
                        n = n + 1
                    End If
                End If
            ElseIf kind = "RECEIVED" And gKsrc Then
                If WmsOrders.OlValid(gKol, od) Then
                    If CStr(od(OD_CANCEL)) = "" And VarType(od(OD_ORD)) = 5 And VarType(od(OD_RCV)) = 5 Then
                        rest = WmsIssue.Round3(od(OD_ORD) - od(OD_RCV))
                        If rest > 0.0000005 Then
                            rows(n) = r
                            how(n) = "MORE"
                            qty(n) = rest
                            n = n + 1
                        End If
                    End If
                End If
            End If
        Next r
    Next b
    If n = 0 Then
        WmsUi.UiMessage("Принимать нечего: в выделенных позициях всё получено, отменено или не указано заказанное количество (H).")
        ReceiveAll = "SKIP:принимать нечего"
        gUiLastMsg = ReceiveAll
        Exit Function
    End If
    For i = 0 To n - 1
        If i < 12 Then sInfo = sInfo & Chr(10) & "строка " & (rows(i) + 1) & ": " & sh.getCellByPosition(OC_NAME, rows(i)).getString() & " — " _
            & WmsIssue.QtyText(qty(i)) & " " & sh.getCellByPosition(OC_UNIT, rows(i)).getString() & IIf(how(i) = "MORE", " (остаток заказа)", "")
    Next i
    If n > 12 Then sInfo = sInfo & Chr(10) & "… и ещё " & (n - 12)
    v(0) = sh.getCellByPosition(OC_DOC, rows(0)).getString()
    v(2) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    If Not WmsPrimeUi.UiForm("WMS — принять всё по заказу", "Принять полностью, как заказано (" & n & " поз.):" & sInfo, _
        Array("№ документа (для всех)", "Дата документа", "Дата поступления"), v, Array("", "", ""), "Принять " & n) Then
        ReceiveAll = "SKIP:окно закрыто"
        gUiLastMsg = ReceiveAll
        Exit Function
    End If
    cols = Array(OC_FACT, OC_DOC, OC_DDATE, OC_RDATE, OC_PLACE, OC_PRICE)
    WmsAsk.AskMemoBegin()
    ' from the bottom up: «Ещё поступление» inserts its row below its position — the rows above keep their numbers
    For i = n - 1 To 0 Step -1
        r = rows(i)
        If how(i) = "FIRST" Then
            For j = 0 To 5
                keep(j) = CellKeep(sh.getCellByPosition(cols(j), r))
            Next j
            ok = True
            sh.getCellByPosition(OC_FACT, r).setString(WmsIssue.QtyText(qty(i)))
            If Trim(v(0)) <> "" And sh.getCellByPosition(OC_DOC, r).getString() = "" Then sh.getCellByPosition(OC_DOC, r).setString(Trim(v(0)))
            If Trim(v(1)) <> "" Then ok = PutDate(sh.getCellByPosition(OC_DDATE, r), v(1), "дата документа", msg)
            If ok Then ok = PutDate(sh.getCellByPosition(OC_RDATE, r), v(2), "дата поступления", msg)
            If ok And sh.getCellByPosition(OC_PLACE, r).getString() = "" Then
                sh.getCellByPosition(OC_PLACE, r).setString(FirstSuggestion(sh.getCellByPosition(OC_NAME, r).getString(), sh.getCellByPosition(OC_ART, r).getString()))
            End If
            If ok Then res = WmsReceipt.ReceiptPostRow(r) Else res = "ERR:" & msg
            If Left(res, 3) <> "OK:" Then
                For j = 0 To 5
                    CellBack(sh.getCellByPosition(cols(j), r), keep(j))
                Next j
                WmsOrders.PreviewOrderRow(r, False)
            End If
        Else
            res = WmsReceipt.ReceiptAddRow(r, WmsIssue.QtyText(qty(i)), "", Trim(v(0)), Trim(v(1)), Trim(v(2)), sh.getCellByPosition(OC_PLACE, r).getString(), _
                sh.getCellByPosition(OC_PRICE, r).getString())
        End If
        If Left(res, 3) = "OK:" Then
            nOk = nOk + 1
        Else
            lines = lines & Chr(10) & "строка " & (r + 1) & ": " & Mid(res, InStr(res, ":") + 1)
            If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then Exit For
        End If
    Next i
    WmsAsk.AskMemoEnd()
    ReceiveAll = "OK=" & nOk & ";ERR=" & (n - nOk) & lines
    If nOk < n Then WmsUi.UiMessage("Принято позиций: " & nOk & " из " & n & "." & Chr(10) & "Не приняты:" & lines)
    gUiLastMsg = ReceiveAll
End Function

' ================================================================ Р8 the return of a list («на руках у …» on «Поиск»)

Sub BtnReturnHands(Optional oEvent As Variant)
    WmsInit()
    ReturnHands()
End Sub

' the lines selected on «Поиск» that are issues (the list «на руках у …», a found list, a card): «Возврат» rows are filled —
' the issue №, what is still on hands, the date — confirmed once and posted one by one (each its own return)
Function ReturnHands() As String
    Dim ctl As Object, sel As Object, ra As Variant, r As Long, fs As Object, iss As Object, sr As Long, k As Long, q As Double, cnt As Long
    Dim rows() As Long, rest() As Double, ks() As Long, n As Long, i As Long, sInfo As String, v(1) As String, rr As Long, res As String
    Dim nOk As Long, lines As String, rsh As Object, n2 As Long, canon As String, msg As String, d As Double, seen As String
    WmsInit()
    ctl = gDoc.getCurrentController()
    If ctl.getActiveSheet().Name <> SH_FIND Then
        ReturnHands = "ERR:откройте «" & SH_FIND & "», найдите получателя («на руках у …») и выделите строки выдач"
        WmsUi.UiMessage(Mid(ReturnHands, 5))
        gUiLastMsg = ReturnHands
        Exit Function
    End If
    sel = gDoc.getCurrentSelection()
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        ra = sel.getRangeAddresses()(0)
    Else
        ra = sel.getRangeAddress()
    End If
    fs = gDoc.Sheets.getByName(SH_FIND)
    iss = WmsIssue.IssuesSheet()
    ReDim rows(500)
    ReDim rest(500)
    ReDim ks(500)
    seen = "|"
    For r = IIf(ra.StartRow < FD_TABLE, FD_TABLE, ra.StartRow) To ra.EndRow
        If n > 500 Then Exit For
        If fs.getCellByPosition(FD_SRC_SHEET, r).getString() = SH_ISSUES Then
            sr = CLng(fs.getCellByPosition(FD_SRC_ROW, r).getValue())
            If sr >= 1 And WmsIssue.RowKind(sr) = "POSTED" Then
                k = CLng(iss.getCellByPosition(IC_NO, sr).getValue())
                If InStr(seen, "|" & k & "|") = 0 Then
                    seen = seen & k & "|"
                    WmsReturn.IssueReturns(k, q, cnt)
                    q = WmsIssue.Round3(iss.getCellByPosition(IC_QTY, sr).getValue() - q)
                    If q > 0.0000005 Then
                        rows(n) = sr
                        rest(n) = q
                        ks(n) = k
                        n = n + 1
                    End If
                End If
            End If
        End If
    Next r
    If n = 0 Then
        ReturnHands = "SKIP:возвращать нечего — выделите строки выдач, которые ещё на руках (список «на руках у …» на «" & SH_FIND & "»)"
        WmsUi.UiMessage(Mid(ReturnHands, 6))
        gUiLastMsg = ReturnHands
        Exit Function
    End If
    For i = 0 To n - 1
        If i < 12 Then
            canon = iss.getCellByPosition(IC_EI, rows(i)).getString()
            If WmsIssue.NormalizeEI(canon, n2, canon, msg) Then
                sInfo = sInfo & Chr(10) & "выдача № " & ks(i) & ": " & canon & " " & iss.getCellByPosition(IC_NAME, rows(i)).getString() & " — вернуть " _
                    & WmsPack.QtyWithParts(n2, rest(i), iss.getCellByPosition(IC_UNIT, rows(i)).getString())
            End If
        End If
    Next i
    If n > 12 Then sInfo = sInfo & Chr(10) & "… и ещё " & (n - 12)
    v(0) = Format(Date(), "DD.MM.YYYY")
    v(1) = ""
    If Not WmsPrimeUi.UiForm("WMS — вернуть списком", "Вернуть всё, что на руках, по выдачам (" & n & "):" & sInfo, _
        Array("Дата возврата", "Место (пусто — как у ЕИ)"), v, Array("", ""), "Вернуть " & n) Then
        ReturnHands = "SKIP:окно закрыто"
        gUiLastMsg = ReturnHands
        Exit Function
    End If
    If Not WmsIssue.ParseDateText(Trim(v(0)), d, msg) Then
        ReturnHands = "ERR:дата возврата — " & msg
        WmsUi.UiMessage(Mid(ReturnHands, 5))
        gUiLastMsg = ReturnHands
        Exit Function
    End If
    rsh = WmsReturn.ReturnsSheet()
    WmsAsk.AskMemoBegin()
    For i = 0 To n - 1
        rr = WmsOrders.LastRow(rsh) + 1
        rsh.getCellByPosition(RC_ISSUE, rr).setValue(ks(i))
        canon = iss.getCellByPosition(IC_EI, rows(i)).getString()
        If WmsIssue.NormalizeEI(canon, n2, canon, msg) Then
            rsh.getCellByPosition(RC_QTY, rr).setString(WmsPack.QtyForInput(n2, rest(i)))
        Else
            rsh.getCellByPosition(RC_QTY, rr).setString(WmsIssue.QtyText(rest(i)))
        End If
        rsh.getCellByPosition(RC_DATE, rr).setValue(d)
        If Trim(v(1)) <> "" Then rsh.getCellByPosition(RC_PLACE, rr).setString(Trim(v(1)))
        res = WmsReturn.ReturnPostRow(rr)
        If Left(res, 3) = "OK:" Then
            nOk = nOk + 1
        Else
            lines = lines & Chr(10) & "выдача № " & ks(i) & ": " & Mid(res, InStr(res, ":") + 1)
            If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Or Left(res, 8) = "ERR-SYS:" Then Exit For
        End If
    Next i
    WmsAsk.AskMemoEnd()
    ReturnHands = "OK=" & nOk & ";ERR=" & (n - nOk) & lines
    If nOk < n Then WmsUi.UiMessage("Возвращено: " & nOk & " из " & n & "." & Chr(10) & "Не проведены (строки остались на «" & SH_RETURNS & "»):" & lines)
    gUiLastMsg = ReturnHands
End Function

' ================================================================ Р11 keyboard shortcuts (tools/build_ods.py HOTKEYS_10)

' Ctrl+Shift+P — «Провести» of the sheet in view
Sub KeyPost(Optional oEvent As Variant)
    Dim sName As String
    WmsInit()
    sName = gDoc.getCurrentController().getActiveSheet().Name
    Select Case sName
    Case SH_ISSUES
        WmsUi.BtnPost()
    Case SH_ORDERS
        WmsOrdersUi.BtnRcvPost()
    Case SH_RETURNS
        WmsReturnUi.BtnRetPost()
    Case SH_SPECIAL
        WmsSpecialUi.BtnSpcPost()
    Case SH_ADJUST
        WmsAdjustUi.BtnAdjPost()
    Case Else
        WmsUi.UiMessage("Ctrl+Shift+P — «Провести» на листах «" & SH_ISSUES & "», «" & SH_ORDERS & "», «" & SH_RETURNS & "», «" & SH_SPECIAL & "», «" _
            & SH_ADJUST & "».")
    End Select
End Sub

' Ctrl+Shift+F — «Поиск», the cursor in «Что найти»
Sub KeyFind(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_FIND, FD_INPUT_ROW, FD_INPUT_COL)
End Sub

' ================================================================ М1 «Разделить ЕИ…»

' the EI of the row under the cursor of «Наличие», «Выдачи», «Корректировки» or of the card of «Поиск» ("" — none)
Function CursorEI() As String
    Dim sName As String, r As Long
    sName = gDoc.getCurrentController().getActiveSheet().Name
    r = WmsView.ViewRow(sName)
    Select Case sName
    Case SH_STOCK
        If r >= 1 Then CursorEI = WmsIssue.StockSheet().getCellByPosition(SC_EI, r).getString()
    Case SH_ISSUES
        If r >= 1 Then CursorEI = WmsIssue.IssuesSheet().getCellByPosition(IC_EI, r).getString()
    Case SH_ADJUST
        If r >= 1 Then CursorEI = gDoc.Sheets.getByName(SH_ADJUST).getCellByPosition(AC_EI, r).getString()
    Case SH_FIND
        CursorEI = WmsFind.FindCardEI()
    End Select
End Function

Sub BtnSplitWindow(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SplitWindow()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Разделить ЕИ…»: part of an EI becomes a new EI with the same card and packing (its own place and label) — a row
' «Разделение» of «Корректировки», posted by the usual checks; the new EI is named in the result
Function SplitWindow() As String
    Dim sEI As String, n As Long, canon As String, msg As String, d As Variant, v(3) As String, ash As Object, r As Long, res As String, ctl As String
    sEI = CursorEI()
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        v(0) = sEI
        canon = ""
    Else
        v(0) = canon
    End If
    If canon <> "" Then
        d = WmsIssue.StockData(n)
        v(2) = CStr(d(SC_PLACE))
    End If
    v(3) = "часть ЕИ — отдельно"
    If Not WmsPrimeUi.UiForm("Разделить ЕИ", "Часть ЕИ станет новым ЕИ с той же карточкой и фасовкой — со своим местом и этикеткой (например: 2 кан из 4 " _
        & "на другое место, обрезок профиля 4,5 м). Учёт: строка «Разделение» на листе «" & SH_ADJUST & "».", _
        Array("ЕИ", "Сколько отделить (2 кан, 4,5 м, 0,5)", "Место нового ЕИ", "Причина"), v, Array("", "", "", ""), "Разделить") Then
        SplitWindow = "SKIP:окно закрыто"
        gUiLastMsg = SplitWindow
        Exit Function
    End If
    If Trim(v(1)) = "" Then
        SplitWindow = "ERR:укажите, сколько отделить"
        gUiLastMsg = SplitWindow
        Exit Function
    End If
    SplitWindow = SplitPost(Trim(v(0)), Trim(v(1)), Trim(v(2)), IIf(Trim(v(3)) <> "", Trim(v(3)), "разделение"))
    gUiLastMsg = SplitWindow
End Function

' a row «Разделение» of «Корректировки» filled and posted: OK:<new EI>|строка N, or the refusal (said in a window)
Function SplitPost(sEI As String, sQty As String, sPlace As String, sReason As String) As String
    Dim ash As Object, r As Long, res As String, ctl As String
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    r = WmsOrders.LastRow(ash) + 1
    If r < 1 Then r = 1
    ash.getCellByPosition(AC_KIND, r).setString("Разделение")
    ash.getCellByPosition(AC_EI, r).setString(sEI)
    ash.getCellByPosition(AC_QTY, r).setString(sQty)
    If sPlace <> "" Then ash.getCellByPosition(AC_PLACE, r).setString(sPlace)
    ash.getCellByPosition(AC_DATE, r).setValue(CDbl(Date()))
    ash.getCellByPosition(AC_REASON, r).setString(sReason)
    res = WmsAdjust.AdjustPostRow(r)
    ctl = ash.getCellByPosition(AC_CTL, r).getString()
    If Left(res, 3) = "OK:" Then
        SplitPost = "OK:" & Mid(ctl, InStr(ctl, "новый ") + 6) & "|строка " & (r + 1)
        WmsUi.UiMessage("Готово: " & Mid(ctl, InStr(ctl, "новый ")) & " — " & sQty & " на месте «" & ash.getCellByPosition(AC_PLACE, r).getString() _
            & "». Наклейте на него свою этикетку.")
    Else
        SplitPost = res
        If Left(res, 5) <> "SKIP:" Then WmsUi.UiMessage("Разделение не проведено: " & Mid(res, InStr(res, ":") + 1) & " (строка " & (r + 1) & " «" & SH_ADJUST & "»)")
    End If
End Function

' ================================================================ Ч7 offcuts of metal

' after an issue in metres from a bar (the packing «1 шт = 6 м»): whole bars and a piece remain — the piece may become an
' EI of its own (an offcut with its own label). Asked once; nothing happens without «Оформить обрезок»
Sub OffcutAfterIssue(r As Long)
    Dim ctl As String, sEI As String, n As Long, canon As String, msg As String, pu As String, pn As Double, b As Double, whole As Double, rest As Double
    Dim k As Integer, d As Variant, sRest As String
    On Error GoTo EH
    ctl = WmsIssue.IssuesSheet().getCellByPosition(IC_CTL, r).getString()
    If Right(ctl, 2) <> " м" Or InStr(ctl, "·") = 0 Then Exit Sub
    sEI = WmsIssue.IssuesSheet().getCellByPosition(IC_EI, r).getString()
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then Exit Sub
    If Not WmsPack.PackOf(n, pu, pn) Then Exit Sub
    If pu <> "м" Or pn <= 1 Then Exit Sub
    If Not WmsIssue.StockQty(n, b) Then Exit Sub
    whole = Int(b + 0.0000001)
    rest = WmsPack.PkRound((b - whole) * pn, 2)
    If whole < 1 Or rest < 0.005 Then Exit Sub
    sRest = WmsIssue.QtyText(rest) & " м"
    d = WmsIssue.StockData(n)
    k = WmsPrimeUi.UiButtons("WMS — обрезок", canon & " · «" & d(SC_NAME) & "»: осталось " & WmsIssue.QtyText(whole) & " " & d(SC_UNIT) & " по " _
        & WmsIssue.QtyText(pn) & " м и кусок " & sRest & "." & Chr(10) & Chr(10) & "Оформить кусок отдельным ЕИ — обрезком со своей этикеткой? Тогда «Обрезок…» " _
        & "найдёт его по длине.", Array("Оформить обрезок " & sRest, "Не надо"), Array(1, 0))
    If k <> 1 Then Exit Sub
    SplitPost(canon, sRest, CStr(d(SC_PLACE)), "обрезок после выдачи № " & WmsIssue.IssuesSheet().getCellByPosition(IC_NO, r).getString())
    Exit Sub
EH:
End Sub

Sub BtnOffcut(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PickOffcut()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Обрезок…» on «Выдачи»: the pieces of the product of the row (the same name, the packing in metres) at least as long as
' needed, the shortest first; the chosen one goes into the row with the length
Function PickOffcut() As String
    Dim sh As Object, r As Long, sEI As String, n As Long, canon As String, msg As String, d As Variant, v(0) As String, L As Double, st As Integer
    Dim last As Long, s As String, a As Variant, i As Long, nC As Long, keys() As Double, idx() As Long, m As Long, pu As String, pn As Double
    Dim labels() As String, codes() As Integer, k As Integer, bm As Double, lenM As Double, nm As String
    sh = WmsIssue.IssuesSheet()
    r = WmsView.ViewRow(SH_ISSUES)
    If r < 1 Then
        PickOffcut = "ERR:встаньте на строку «" & SH_ISSUES & "» с ЕИ профиля (трубы, прутка) — WMS подберёт кусок нужной длины"
        gUiLastMsg = PickOffcut
        Exit Function
    End If
    sEI = sh.getCellByPosition(IC_EI, r).getString()
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        PickOffcut = "ERR:в строке нет ЕИ — впишите ЕИ любого куска этого профиля, WMS подберёт кусок нужной длины"
        gUiLastMsg = PickOffcut
        Exit Function
    End If
    d = WmsIssue.StockData(n)
    nm = CStr(d(SC_NAME))
    v(0) = ""
    If Not WmsPrimeUi.UiForm("Обрезок нужной длины", "«" & nm & "»: WMS найдёт кусок не короче нужного (самый короткий подходящий — первым).", _
        Array("Нужная длина, м"), v, Array(""), "Подобрать") Then
        PickOffcut = "SKIP:окно закрыто"
        gUiLastMsg = PickOffcut
        Exit Function
    End If
    st = ParseQtyText(Trim(Replace(LCase(v(0)), "м", "")), L, msg)
    If st <> 0 Then
        PickOffcut = "ERR:длина «" & v(0) & "»: " & msg
        gUiLastMsg = PickOffcut
        Exit Function
    End If
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    s = WmsView.Eng("TEXTJOIN("";"";1;IF((" & WmsView.ColRng(SH_STOCK, SC_NAME, 1, last) & "=""" & Replace(nm, """", """""") & """)*(" _
        & WmsView.ColRng(SH_PACK, PK_EI, 1, last) & "=" & WmsView.ColRng(SH_STOCK, SC_EI, 1, last) & ")*(" & WmsView.ColRng(SH_PACK, PK_UNIT, 1, last) _
        & "=""м"")*ISNUMBER(" & WmsView.ColRng(SH_STOCK, SC_QTY, 1, last) & ")*(" & WmsView.ColRng(SH_STOCK, SC_QTY, 1, last) & "*" _
        & WmsView.ColRng(SH_PACK, PK_N, 1, last) & ">=" & Replace(CStr(L - 0.0000001), ",", ".") & ");ROW(" & WmsView.ColRng(SH_STOCK, SC_QTY, 1, last) & ")-1;""""))")
    If s = "" Then
        PickOffcut = "ERR:куска «" & nm & "» не короче " & WmsIssue.QtyText(L) & " м нет"
        gUiLastMsg = PickOffcut
        Exit Function
    End If
    a = Split(s, ";")
    nC = UBound(a) + 1
    ReDim keys(nC - 1)
    ReDim idx(nC - 1)
    For i = 0 To nC - 1
        m = CLng(a(i))
        WmsIssue.StockQty(m, bm)
        WmsPack.PackOf(m, pu, pn)
        keys(i) = bm * pn + m * 0.000000001
        idx(i) = i
    Next i
    WmsView.SortIdx(keys, idx, nC, False)
    If nC > 3 Then nC = 3
    ReDim labels(nC)
    ReDim codes(nC)
    For i = 0 To nC - 1
        m = CLng(a(idx(i)))
        lenM = Int(keys(idx(i)) * 100 + 0.5) / 100
        labels(i) = WmsIssue.EiCanon(m) & " — " & WmsIssue.QtyText(lenM) & " м · " & WmsIssue.StockSheet().getCellByPosition(SC_PLACE, m).getString()
        codes(i) = i + 1
    Next i
    labels(nC) = "Отмена"
    codes(nC) = 0
    k = WmsPrimeUi.UiButtons("Обрезок «" & nm & "» от " & WmsIssue.QtyText(L) & " м", "Куски не короче " & WmsIssue.QtyText(L) & " м (самый короткий — первым):", labels, codes)
    If k < 1 Then
        PickOffcut = "SKIP:кусок не выбран"
        gUiLastMsg = PickOffcut
        Exit Function
    End If
    m = CLng(a(idx(k - 1)))
    sh.getCellByPosition(IC_EI, r).setString(WmsIssue.EiCanon(m))
    sh.getCellByPosition(IC_QTY, r).setString(WmsIssue.QtyText(L) & " м")
    WmsIssue.PreviewRow(r, True, False)
    PickOffcut = "OK:" & WmsIssue.EiCanon(m)
    gUiLastMsg = PickOffcut
End Function

' ================================================================ М11 the issue «с места»

' after an issue posted by «Провести»: the rest of a bar (Ч7), then the places of an EI on several places (М11)
Sub AfterIssue(r As Long)
    OffcutAfterIssue(r)
    PlaceAfterIssue(r)
End Sub

' 1.0 (М11): an EI on several places and something left after the issue — did one of the places become empty? The chosen
' place leaves the list of the EI by a row «Перемещение» («−1-2-4»), posted at once. Enter — «Места не менялись»
Sub PlaceAfterIssue(r As Long)
    Dim sEI As String, n As Long, canon As String, msg As String, b As Double, d As Variant, pls As Variant, labels() As String, codes() As Integer
    Dim i As Integer, m As Integer, k As Integer, sNo As String
    On Error GoTo EH
    sEI = WmsIssue.IssuesSheet().getCellByPosition(IC_EI, r).getString()
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then Exit Sub
    If Not WmsIssue.StockQty(n, b) Then Exit Sub
    If b <= 0.0000005 Then Exit Sub
    d = WmsIssue.StockData(n)
    pls = WmsView.PlaceList(CStr(d(SC_PLACE)))
    If UBound(pls) < 1 Then Exit Sub
    m = UBound(pls) + 1
    If m > 6 Then m = 6
    ReDim labels(m)
    ReDim codes(m)
    labels(0) = "Места не менялись"
    codes(0) = 0
    For i = 0 To m - 1
        labels(i + 1) = "Место " & pls(i) & " опустело — убрать его из списка"
        codes(i + 1) = 10 + i
    Next i
    sNo = WmsIssue.IssuesSheet().getCellByPosition(IC_NO, r).getString()
    k = WmsPrimeUi.UiButtons("WMS — место", canon & " · «" & d(SC_NAME) & "» лежит на " & (UBound(pls) + 1) & " местах: " & Join(pls, ", ") _
        & ". После выдачи № " & sNo & " осталось " & WmsPack.QtyWithParts(n, b, CStr(d(SC_UNIT))) & "." & Chr(10) & Chr(10) _
        & "Одно из мест опустело? WMS уберёт его из списка мест ЕИ (строка «Перемещение» на листе «" & SH_ADJUST & "»).", labels, codes)
    If k < 10 Then Exit Sub
    MovePost(canon, "−" & pls(k - 10), "место " & pls(k - 10) & " опустело после выдачи № " & sNo)
    Exit Sub
EH:
End Sub

' a row «Перемещение» of «Корректировки» filled and posted (the new place may be a change of the list: «−1-2-4», «+1-2-5»):
' OK:<the new place>|строка N, or the refusal (said in a window)
Function MovePost(sEI As String, sPlace As String, sReason As String) As String
    Dim ash As Object, r As Long, res As String
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    r = WmsOrders.LastRow(ash) + 1
    If r < 1 Then r = 1
    ash.getCellByPosition(AC_KIND, r).setString("Перемещение")
    ash.getCellByPosition(AC_EI, r).setString(sEI)
    ash.getCellByPosition(AC_PLACE, r).setString(sPlace)
    ash.getCellByPosition(AC_DATE, r).setValue(CDbl(Date()))
    ash.getCellByPosition(AC_REASON, r).setString(sReason)
    res = WmsAdjust.AdjustPostRow(r)
    If Left(res, 3) = "OK:" Then
        MovePost = "OK:" & ash.getCellByPosition(AC_PLACE, r).getString() & "|строка " & (r + 1)
    Else
        MovePost = res
        If Left(res, 5) <> "SKIP:" Then WmsUi.UiMessage("Перемещение не проведено: " & Mid(res, InStr(res, ":") + 1) & " (строка " & (r + 1) & " «" & SH_ADJUST & "»)")
    End If
    gUiLastMsg = MovePost
End Function

' the test seam of WmsView.PlaceEdit: OK:<the new list> / ERR:<the refusal>
Function TestPlaceEdit(sBase As String, sTyped As String) As String
    Dim why As String, sOut As String
    why = WmsView.PlaceEdit(sBase, sTyped, sOut)
    TestPlaceEdit = IIf(why = "", "OK:" & sOut, "ERR:" & why)
End Function

' ================================================================ М7 a row, a section or a place renamed

Sub BtnRenamePlaces(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = RenamePlaces()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the parts of «5», «5-12», «5-12-3» (a row, a section, a place) as the keys of the places have them; Array() when it is none
Private Function RenameParts(ByVal s As String) As Variant
    Dim k As String, a As Variant
    RenameParts = Array()
    s = Trim(s)
    If s = "" Then Exit Function
    k = WmsView.PlaceKey(s)
    If k = "" Then
        ' one part — a row alone («5», «Р5», «А»): read as the first part of a key
        k = WmsView.PlaceKey(s & "-0")
        If k = "" Then Exit Function
        a = Split(k, "-")
        RenameParts = Array(a(0))
        Exit Function
    End If
    RenameParts = Split(k, "-")
End Function

' «Переименовать ряд…»: the row 5 became 6 (or a section, a place): rows «Перемещение» are prepared on «Корректировки» for every
' EI with a balance on those places — not posted: the storekeeper checks them and posts the selected block by «Провести».
' The map («Топология») and «Места» are renamed too when asked (the previous map kept in «_TOPO_PREV»). OK:<rows>|<first>-<last> /
' SKIP:… / ERR:…
Function RenamePlaces() As String
    Dim v(2) As String, oldP As Variant, newP As Variant, st As Object, last As Long, r0 As Long, r1 As Long, d As Variant, i As Long
    Dim items As Variant, j As Integer, nk As String, t As String, seen As String, changed As Boolean, eis() As String, pls() As String, n As Long
    Dim nTopo As Long, nDir As Long, nBusy As Long, k As Integer, ash As Object, r As Long, a() As Variant, b() As Variant, sOld As String, sNew As String
    Dim labels As Variant, codes As Variant, sText As String, bMap As Boolean, ky As String
    On Error GoTo EH
    If WmsMap.MapEditing() Then
        RenamePlaces = "ERR:карта в режиме правки — сначала «Сохранить карту» или «Отменить правку»"
        gUiLastMsg = RenamePlaces
        Exit Function
    End If
    v(2) = "переименование мест"
    If Not WmsPrimeUi.UiForm("Переименовать ряд или место", "Ряд, секция или место получили другой номер (перевесили таблички, перенумеровали ряды). " _
        & "WMS подготовит строки «Перемещение» на листе «" & SH_ADJUST & "» для всех ЕИ с остатком, что там лежат, — вы проверите их и проведёте сами. " _
        & "Ряд — «5», секция — «5-12», место — «5-12-3».", Array("Было", "Стало", "Причина"), v, Array("", "", ""), "Подготовить") Then
        RenamePlaces = "SKIP:окно закрыто"
        gUiLastMsg = RenamePlaces
        Exit Function
    End If
    oldP = RenameParts(v(0))
    newP = RenameParts(v(1))
    If UBound(oldP) < 0 Then
        RenamePlaces = "ERR:«Было»: напишите ряд («5»), секцию («5-12») или место («5-12-3»)"
    ElseIf UBound(newP) <> UBound(oldP) Then
        RenamePlaces = "ERR:«Стало» пишется так же, как «Было»: ряд на ряд («5» → «6»), секция на секцию («5-12» → «5-14»), место на место"
    ElseIf Join(oldP, "-") = Join(newP, "-") Then
        RenamePlaces = "ERR:«Было» и «Стало» совпадают — переименовывать нечего"
    End If
    If RenamePlaces <> "" Then
        gUiLastMsg = RenamePlaces
        Exit Function
    End If
    sOld = Join(oldP, "-")
    sNew = Join(newP, "-")
    ' the EIs with a balance on those places: their new lists (the other places of a list stay as written)
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    ReDim eis(100)
    ReDim pls(100)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = st.getCellRangeByPosition(0, r0, SC_PLACE, r1).getDataArray()
        For i = 0 To r1 - r0
            If VarType(d(i)(SC_QTY)) = 5 And Trim(CStr(d(i)(SC_PLACE))) <> "" Then
                If d(i)(SC_QTY) > 0.0000005 Then
                    items = WmsView.PlaceItems(CStr(d(i)(SC_PLACE)))
                    t = ""
                    seen = "|"
                    changed = False
                    For j = 0 To UBound(items)
                        ky = WmsView.PlaceKey(CStr(items(j)))
                        nk = WmsMap.KeyRenamed(ky, oldP, newP)
                        If nk <> "" Then
                            changed = True
                        Else
                            nk = Trim(CStr(items(j)))
                            If WmsMap.KeyRenamed(ky, newP, newP) <> "" Then nBusy = nBusy + 1
                        End If
                        If InStr(seen, "|" & UCase(nk) & "|") = 0 Then
                            t = t & IIf(t <> "", ", ", "") & nk
                            seen = seen & UCase(nk) & "|"
                        End If
                    Next j
                    If changed Then
                        If n > UBound(eis) Then
                            ReDim Preserve eis(2 * n)
                            ReDim Preserve pls(2 * n)
                        End If
                        eis(n) = CStr(d(i)(SC_EI))
                        pls(n) = t
                        n = n + 1
                    End If
                End If
            End If
        Next i
    Next r0
    nTopo = WmsMap.TopoRename(oldP, newP, False, nDir)
    If n = 0 And nTopo = 0 And nDir = 0 Then
        RenamePlaces = "ERR:мест «" & sOld & "…» нет ни у одного ЕИ с остатком, ни на карте склада, ни в справочнике «" & SH_PLACES & "»"
        gUiLastMsg = RenamePlaces
        Exit Function
    End If
    sText = "«" & sOld & "» → «" & sNew & "»." & Chr(10) & Chr(10) & "ЕИ с остатком на этих местах: " & n _
        & IIf(n > 0, " — для них будут подготовлены строки «Перемещение» на листе «" & SH_ADJUST & "» (ещё не проведены).", ".")
    If nBusy > 0 Then sText = sText & Chr(10) & "Внимание: на местах «" & sNew & "…» уже лежат ЕИ (" & nBusy & ") — после переименования места совпадут."
    If nTopo + nDir > 0 Then sText = sText & Chr(10) & "На карте склада таких мест: " & nTopo & ", в справочнике «" & SH_PLACES & "»: " & nDir & "."
    If n > 0 And nTopo + nDir > 0 Then
        labels = Array("Подготовить строки и переименовать на карте", "Только подготовить строки", "Отмена")
        codes = Array(3, 1, 0)
    ElseIf n > 0 Then
        labels = Array("Подготовить строки", "Отмена")
        codes = Array(1, 0)
    Else
        labels = Array("Переименовать на карте", "Отмена")
        codes = Array(2, 0)
    End If
    k = WmsPrimeUi.UiButtons("Переименовать места", sText, labels, codes)
    If k = 0 Or k = -1 Then
        RenamePlaces = "SKIP:ничего не изменено"
        gUiLastMsg = RenamePlaces
        Exit Function
    End If
    bMap = (k = 2 Or k = 3)
    If bMap Then WmsMap.TopoRename(oldP, newP, True, nDir)
    If n = 0 Then
        RenamePlaces = "OK:0|карта"
        WmsUi.UiMessage("Готово: на карте склада и в «" & SH_PLACES & "» места «" & sOld & "…» стали «" & sNew & "…» (прежняя карта — в скрытом листе «" & SH_TOPO_PREV & "»).")
        gUiLastMsg = RenamePlaces
        Exit Function
    End If
    ' the rows: B «Перемещение», C the EI; J the new places, K today, L the reason (only the input columns)
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    r = WmsOrders.LastRow(ash) + 1
    If r < 1 Then r = 1
    ReDim a(n - 1)
    ReDim b(n - 1)
    For i = 0 To n - 1
        a(i) = Array("Перемещение", eis(i))
        b(i) = Array(pls(i), CDbl(Date()), IIf(Trim(v(2)) <> "", Trim(v(2)), "переименование мест") & ": «" & sOld & "» → «" & sNew & "»")
    Next i
    ash.getCellRangeByPosition(AC_KIND, r, AC_EI, r + n - 1).setDataArray(a)
    ash.getCellRangeByPosition(AC_PLACE, r, AC_REASON, r + n - 1).setDataArray(b)
    gDoc.getCurrentController().setActiveSheet(ash)
    gDoc.getCurrentController().select(ash.getCellRangeByPosition(0, r, AC_LAST, r + n - 1))
    RenamePlaces = "OK:" & n & "|" & (r + 1) & "-" & (r + n)
    WmsUi.UiMessage("Подготовлено строк «Перемещение»: " & n & " (строки " & (r + 1) & "–" & (r + n) & " листа «" & SH_ADJUST & "»), «" & sOld & "» → «" _
        & sNew & "». Ничего ещё не проведено." & Chr(10) & Chr(10) & "Строки выделены: проверьте новые места (J) и нажмите «Провести» — WMS проведёт их блоком." _
        & IIf(bMap, Chr(10) & "Карта склада и «" & SH_PLACES & "» уже переименованы (прежняя карта — в скрытом листе «" & SH_TOPO_PREV & "»).", ""))
    gUiLastMsg = RenamePlaces
    Exit Function
EH:
    RenamePlaces = "ERR:строки не подготовлены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = RenamePlaces
End Function

' 1.0 (В4): «Справка» from «Главная»
Sub NavHelp(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell("Справка", 0, 0)
    gUiLastMsg = "OK:Справка"
End Sub

' ================================================================ В6 «Почему?»

Sub BtnWhy(Optional oEvent As Variant)
    WmsInit()
    WhyRow()
End Sub

' the column «Контроль» of a working sheet (-1 — none)
Private Function CtlCol(sName As String) As Integer
    Select Case sName
    Case SH_ISSUES
        CtlCol = IC_CTL
    Case SH_ORDERS
        CtlCol = OC_CTL
    Case SH_RETURNS
        CtlCol = RC_CTL
    Case SH_SPECIAL
        CtlCol = XC_CTL
    Case SH_ADJUST
        CtlCol = AC_CTL
    Case Else
        CtlCol = -1
    End Select
End Function

' what to do about a text of «Контроль» (the words of the commonest reasons)
Function WhyAdvice(ByVal s As String) As String
    Dim t As String
    t = LCase(s)
    If InStr(t, "заблокирован") > 0 Then
        WhyAdvice = "Проведение остановлено самой WMS — на листе «" & SH_MAIN & "» написано почему и какую кнопку нажать."
    ElseIf InStr(t, "копия") > 0 Then
        WhyAdvice = "Строка скопирована с другой (тот же номер) — нажмите «Очистить» и введите строку заново."
    ElseIf InStr(t, "недостаточно") > 0 Or InStr(t, "остаток") > 0 Then
        WhyAdvice = "Проверьте количество и ЕИ: на странице «" & SH_FIND & "» карточка ЕИ покажет остаток и все движения. Частями — «1 кан», «1,5 м»; весь остаток — «всё»."
    ElseIf InStr(t, "не найден") > 0 Or InStr(t, "не выдавал") > 0 Or InStr(t, "неверный еи") > 0 Then
        WhyAdvice = "Проверьте номер ЕИ — «Подобрать ЕИ» (Ctrl+Shift+E) найдёт его по наименованию, артикулу или месту."
    ElseIf InStr(t, "дата") > 0 Then
        WhyAdvice = "Дата пишется так: 15.10.2026. Будущая дата или дата больше года назад — WMS переспросит."
    ElseIf InStr(t, "получател") > 0 Then
        WhyAdvice = "Получатель — из листа «" & SH_RCPT & "» (можно сокращение) или полностью, как в справочнике."
    ElseIf InStr(t, "мест") > 0 Then
        WhyAdvice = "Место — как на «Топологии» (например 1-2-3) или из листа «Места»; несколько мест — через запятую."
    ElseIf InStr(t, "провед") > 0 And InStr(t, "не провед") = 0 Then
        WhyAdvice = "Строка проведена. Изменить — «Исправить», отменить — «Удалить» (сторно)."
    Else
        WhyAdvice = "Исправьте значения в строке и нажмите «Провести» ещё раз."
    End If
End Function

' «Почему?»: the whole text of «Контроль» of the row under the cursor and what to do about it
Function WhyRow() As String
    Dim sName As String, r As Long, c As Integer, s As String
    sName = gDoc.getCurrentController().getActiveSheet().Name
    c = CtlCol(sName)
    r = WmsView.ViewRow(sName)
    If c < 0 Or r < 1 Then
        WhyRow = "ERR:встаньте на строку листа «" & SH_ISSUES & "», «" & SH_ORDERS & "», «" & SH_RETURNS & "», «" & SH_SPECIAL & "» или «" & SH_ADJUST & "»"
        WmsUi.UiMessage("«Почему?»: " & Mid(WhyRow, 5) & " — WMS покажет полный текст «Контроль» и что сделать.")
        gUiLastMsg = WhyRow
        Exit Function
    End If
    s = Trim(gDoc.Sheets.getByName(sName).getCellByPosition(c, r).getString())
    If s = "" Then
        WhyRow = "OK:строка " & (r + 1) & ": в «Контроль» пусто — строка ещё не проверялась (заполните её и нажмите «Провести»)"
    Else
        WhyRow = "OK:строка " & (r + 1) & ": " & s & Chr(10) & Chr(10) & "Что сделать: " & WhyAdvice(s)
    End If
    WmsUi.UiMessage(Mid(WhyRow, 4))
    gUiLastMsg = WhyRow
End Function
