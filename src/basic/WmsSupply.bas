' WmsSupply — 1.0, этап F «Заказы и поставщики» (выбор кладовщика): возврат поставщику (З1), страница «Снабжение» —
' «Ожидаем» (З2) и «Нужно заказать» (З10), справочник «Поставщики» (З8), повтор (З6) и отмена (З5) заказа, машина и её
' приход (З9), акт расхождения (З12). Library WmsPlus. Nothing here writes the accounting by itself: the operations of the
' core do (WmsCore.ApplyOperation); the rows of «Заказы» WMS fills here are rows not yet received, as the storekeeper types them.
Option Explicit

Public Const SH_SUPPLIERS = "Поставщики"
' «Поставщики»: A the supplier (one spelling), B phone, C contact, D the usual delivery time in days, E note
Private Const SU_NAME = 0
Private Const SU_PHONE = 1
Private Const SU_CONTACT = 2
Private Const SU_LEAD = 3
Private Const SU_NOTE = 4

' the result of the last return window (for the tests): the row of «Корректировки», the row of the replacement in «Заказы»
Global gSuRetRow As Long
Global gSuNewRow As Long

' ================================================================ the supplier of an EI

' the supplier of EI n: «Заказы» L of its receipt; an EI split from another (М1) — the supplier of that one; "" when the EI
' came otherwise («Иной приход», the transfer) or the row is not found
Function EiSupplier(ByVal n As Long) As String
    Dim rv As Variant, r As Long, canon As String, k As Integer, m As Long
    On Error GoTo EH
    canon = WmsIssue.EiCanon(n)
    For k = 0 To 5
        If WmsOrders.RcvRegistered(n, rv) Then
            r = WmsOrders.ReceiptRowOf(canon, rv(RV_ROW))
            If r >= 1 Then EiSupplier = Trim(WmsOrders.OrdersSheet().getCellByPosition(OC_SUPPLIER, r).getString())
            Exit Function
        End If
        m = SplitSourceOf(canon)
        If m < 1 Then Exit Function
        n = m
        canon = WmsIssue.EiCanon(n)
    Next k
EH:
End Function

' the EI a split (М1) made canon from; 0 when canon was not made by a split
Function SplitSourceOf(canon As String) As Long
    Dim sh As Object, last As Long, desc As Object, found As Object, a As Variant, d As Variant, m As Long
    On Error GoTo EH
    sh = gDoc.Sheets.getByName(SH_ADJ)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    desc = sh.createSearchDescriptor()
    desc.SearchString = "^" & canon & "$"
    desc.SearchRegularExpression = True
    desc.SearchCaseSensitive = True
    found = sh.getCellRangeByPosition(AJ_NEW, 1, AJ_NEW, last).findAll(desc)
    If IsNull(found) Then Exit Function
    a = found.getRangeAddresses()
    d = sh.getCellRangeByPosition(0, a(0).StartRow, AJ_LAST, a(0).StartRow).getDataArray()(0)
    If CStr(d(AJ_KIND)) <> "SPLIT" Then Exit Function
    If WmsOrders.StrictEI(CStr(d(AJ_EI)), m) Then SplitSourceOf = m
EH:
End Function

' ================================================================ З8 the directory «Поставщики»

' the row (1-based in the sheet: 1 … ) of supplier s in «Поставщики» — the same name in any case and spacing; 0 when absent
Function SupplierRow(ByVal s As String) As Long
    Dim sh As Object, last As Long, d As Variant, i As Long, k As String
    On Error GoTo EH
    k = SupKey(s)
    If k = "" Or Not gDoc.Sheets.hasByName(SH_SUPPLIERS) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_SUPPLIERS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(SU_NAME, 1, SU_NAME, last).getDataArray()
    For i = 0 To UBound(d)
        If SupKey(CStr(d(i)(0))) = k Then
            SupplierRow = i + 1
            Exit Function
        End If
    Next i
EH:
End Function

' a supplier as it is compared: upper case, «ё» = «е», the spaces, the quotes and the dots do not part two spellings
Function SupKey(ByVal s As String) As String
    Dim i As Integer, ch As String, r As String
    s = UCase(Trim(s))
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "Ё" Then ch = "Е"
        If ch <> " " And ch <> """" And ch <> "«" And ch <> "»" And ch <> "." And ch <> "," And ch <> "'" Then r = r & ch
    Next i
    SupKey = r
End Function

' the usual delivery time of supplier s in days («Поставщики» D); 0 when not known
Function SupplierLead(s As String) As Long
    Dim r As Long, c As Object
    r = SupplierRow(s)
    If r < 1 Then Exit Function
    c = gDoc.Sheets.getByName(SH_SUPPLIERS).getCellByPosition(SU_LEAD, r)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then
        If c.getValue() >= 1 And c.getValue() < 1000 Then SupplierLead = CLng(Int(c.getValue()))
    End If
End Function

Sub BtnSuppliersFill(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SuppliersFill()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsOrdersUi.UiInfo(Mid(s, 4))
End Sub

' «Пополнить из заказов»: the suppliers of «Заказы» (L) and of «Приход авто» (D) not yet in «Поставщики» are added below
' (each once, as first written). OK:<text> / ERR:…
Function SuppliersFill() As String
    Dim sh As Object, known As Object, last As Long, d As Variant, i As Long, k As String, add() As String, nAdd As Long, src As Variant, j As Integer
    Dim r As Long, a() As Variant
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_SUPPLIERS) Then
        SuppliersFill = "ERR:в книге нет листа «" & SH_SUPPLIERS & "» — его создаёт обновление до 1.0"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_SUPPLIERS)
    known = New Collection
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        d = sh.getCellRangeByPosition(SU_NAME, 1, SU_NAME, last).getDataArray()
        For i = 0 To UBound(d)
            k = SupKey(CStr(d(i)(0)))
            If k <> "" Then CollPut(known, 1, k)
        Next i
    End If
    ReDim add(100)
    src = Array(Array(SH_ORDERS, OC_SUPPLIER, 1), Array(SH_CARS, CC_SUP, CAR_FIRST))
    For j = 0 To 1
        If gDoc.Sheets.hasByName(src(j)(0)) Then
            r = WmsOrders.LastRow(gDoc.Sheets.getByName(src(j)(0)))
            If r >= src(j)(2) Then
                d = gDoc.Sheets.getByName(src(j)(0)).getCellRangeByPosition(src(j)(1), src(j)(2), src(j)(1), r).getDataArray()
                For i = 0 To UBound(d)
                    k = SupKey(CStr(d(i)(0)))
                    If k <> "" Then
                        If CollPut(known, 1, k) Then
                            If nAdd > UBound(add) Then
                                ReDim Preserve add(2 * nAdd)
                            End If
                            add(nAdd) = Trim(CStr(d(i)(0)))
                            nAdd = nAdd + 1
                        End If
                    End If
                Next i
            End If
        End If
    Next j
    If nAdd = 0 Then
        SuppliersFill = "OK:все поставщики «" & SH_ORDERS & "» и «" & SH_CARS & "» уже есть в «" & SH_SUPPLIERS & "»"
        gUiLastMsg = SuppliersFill
        Exit Function
    End If
    ReDim a(nAdd - 1)
    For i = 0 To nAdd - 1
        a(i) = Array(add(i))
    Next i
    sh.getCellRangeByPosition(SU_NAME, last + 1, SU_NAME, last + nAdd).setDataArray(a)
    SuppliersFill = "OK:в «" & SH_SUPPLIERS & "» добавлено поставщиков: " & nAdd & " (строки " & (last + 2) & "–" & (last + nAdd + 1) _
        & "). Впишите телефон и обычный срок поставки — срок подставит ожидаемую дату новых заказов."
    gUiLastMsg = SuppliersFill
    Exit Function
EH:
    SuppliersFill = "ERR:поставщики не добавлены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = SuppliersFill
End Function

' ================================================================ new rows of «Заказы» (З1 the replacement, З6 the repeat)

' a new row of «Заказы» below the last one, as the storekeeper would type it: the position of row src (A order, B name,
' D invoice, E article, I unit, J price, L supplier, M seller, R buyer, S category, T assigned), H — qty, P — today, Q — today
' plus the usual delivery time of the supplier when «Поставщики» knows it, Z — note. The row (0-based); −1 on failure (why)
Function NewOrderRowLike(src As Long, qty As Double, note As String, ByRef why As String) As Long
    Dim sh As Object, r As Long, d As Variant, a(20) As Variant, i As Integer, lead As Long, t As Double, cc As Variant
    On Error GoTo EH
    NewOrderRowLike = -1
    sh = WmsOrders.OrdersSheet()
    r = WmsOrders.LastRow(sh) + 1
    If r >= MAX_SHEET_ROW Then
        why = "на листе «" & SH_ORDERS & "» нет места для новой строки"
        Exit Function
    End If
    d = sh.getCellRangeByPosition(0, src, OC_PLACE, src).getDataArray()(0)
    For i = 0 To 20
        a(i) = ""
    Next i
    For Each cc In Array(OC_ORDER, OC_NAME, OC_INVOICE, OC_ART, OC_UNIT, OC_PRICE, OC_SUPPLIER, OC_SELLER, OC_BUYER, OC_CAT, OC_ASSIGNED)
        a(cc) = d(cc)
    Next cc
    a(OC_ORDQTY) = WmsIssue.QtyText(qty)
    t = WmsOrders.Today()
    a(OC_ODATE) = t
    lead = SupplierLead(CStr(d(OC_SUPPLIER)))
    If lead > 0 Then a(OC_EDATE) = t + lead
    WmsOrders.HandlerOff(True)
    sh.getCellRangeByPosition(0, r, OC_PLACE, r).setDataArray(Array(a))
    sh.getCellByPosition(OC_NOTE, r).setString(note)
    WmsOrders.HandlerOff(False)
    WmsOrders.PreviewOrderRow(r, True)
    NewOrderRowLike = r
    Exit Function
EH:
    why = "строка заказа не создана: " & Error$
    On Error Resume Next
    WmsOrders.HandlerOff(False)
End Function

' ================================================================ З1 the return to the supplier

Sub BtnSupReturn(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SupReturnWindow()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Вернуть поставщику…»: the EI under the cursor (or typed), how much, the reason, the document of the return, «ждём замену»:
' a row «Возврат поставщику» of «Корректировки» posted by the usual checks (the balance goes down, the supplier and the
' document are in the journal); «ждём замену — да» adds a new row of «Заказы» for the same goods and supplier — the position
' is expected again. OK:<row of «Корректировки»>|<row of the replacement or 0> / SKIP:… / ERR:…
Function SupReturnWindow() As String
    Dim sEI As String, n As Long, canon As String, msg As String, d As Variant, v(4) As String, sup As String, sInfo As String
    sEI = WmsDaily.CursorEI()
    v(0) = sEI
    If WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        v(0) = canon
        d = WmsIssue.StockData(n)
        sup = EiSupplier(n)
        sInfo = canon & " · «" & d(SC_NAME) & "» · остаток " & WmsPack.QtyWithParts(n, CDbl(d(SC_QTY)), CStr(d(SC_UNIT))) & " · поставщик " _
            & IIf(sup <> "", "«" & sup & "»", "не известен (ЕИ пришёл не по заказу)") & "."
    End If
    v(4) = "нет"
    If Not WmsPrimeUi.UiForm("Вернуть поставщику", IIf(sInfo <> "", sInfo & " ", "") & "Брак, пересорт, лишнее: остаток уменьшится, поставщик и документ " _
        & "запишутся. «Ждём замену — да» — позиция снова появится в «Ожидаем» (новая строка «" & SH_ORDERS & "»).", _
        Array("ЕИ", "Сколько вернуть (2, 1 кан, 1,5 м)", "Причина (брак, пересорт…)", "Документ возврата (если есть)", "Ждём замену?"), v, _
        Array("", "", "брак|пересорт|лишнее|не тот товар", "", "нет|да"), "Вернуть") Then
        SupReturnWindow = "SKIP:окно закрыто"
        gUiLastMsg = SupReturnWindow
        Exit Function
    End If
    SupReturnWindow = SupReturnPost(Trim(v(0)), Trim(v(1)), Trim(v(2)), Trim(v(3)), (LCase(Trim(v(4))) = "да"))
    gUiLastMsg = SupReturnWindow
End Function

' a row «Возврат поставщику» filled and posted; bWait — the replacement row in «Заказы» after the posting
Function SupReturnPost(sEI As String, sQty As String, sReason As String, sDoc As String, bWait As Boolean) As String
    Dim ash As Object, r As Long, res As String, n As Long, canon As String, msg As String, nr As Long, why As String, src As Long, rv As Variant
    Dim q As Double, nAdj As String
    gSuRetRow = 0
    gSuNewRow = 0
    If sQty = "" Then
        SupReturnPost = "ERR:укажите, сколько вернуть"
        Exit Function
    End If
    If sReason = "" Then
        SupReturnPost = "ERR:укажите причину возврата (брак, пересорт…)"
        Exit Function
    End If
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    r = WmsOrders.LastRow(ash) + 1
    If r < 1 Then r = 1
    ash.getCellByPosition(AC_KIND, r).setString("Возврат поставщику")
    ash.getCellByPosition(AC_EI, r).setString(sEI)
    ash.getCellByPosition(AC_QTY, r).setString(sQty)
    If sDoc <> "" Then ash.getCellByPosition(AC_PLACE, r).setString(sDoc)
    ash.getCellByPosition(AC_DATE, r).setValue(CDbl(Date()))
    ash.getCellByPosition(AC_REASON, r).setString(sReason & IIf(bWait, "; ждём замену", ""))
    res = WmsAdjust.AdjustPostRow(r)
    gSuRetRow = r
    If Left(res, 3) <> "OK:" Then
        SupReturnPost = res
        If Left(res, 5) <> "SKIP:" Then WmsUi.UiMessage("Возврат поставщику не проведён: " & Mid(res, InStr(res, ":") + 1) & " (строка " & (r + 1) & " «" & SH_ADJUST & "»)")
        Exit Function
    End If
    nAdj = ash.getCellByPosition(AC_NO, r).getString()
    q = ash.getCellByPosition(AC_QTY, r).getValue()
    canon = ash.getCellByPosition(AC_EI, r).getString()
    msg = "Возврат поставщику № " & nAdj & " проведён: " & canon & " — " & WmsPack.QtyWithParts(WmsView.EiNo(canon, CLng(SysNum(SK_NEXT_EI))), q, _
        ash.getCellByPosition(AC_UNIT, r).getString()) & "."
    If bWait Then
        src = -1
        If WmsOrders.StrictEI(canon, n) Then
            If WmsOrders.RcvRegistered(n, rv) Then src = WmsOrders.ReceiptRowOf(canon, rv(RV_ROW))
        End If
        If src < 1 Then
            msg = msg & Chr(10) & "Замену записать некуда: ЕИ пришёл не по заказу — оформите новую строку «" & SH_ORDERS & "» сами."
        Else
            nr = NewOrderRowLike(src, q, "замена по возврату поставщику № " & nAdj & " (" & canon & "): " & sReason, why)
            If nr < 0 Then
                msg = msg & Chr(10) & "Строка замены не создана: " & why & " — оформите её на листе «" & SH_ORDERS & "» сами."
            Else
                gSuNewRow = nr
                msg = msg & Chr(10) & "Замена ожидается: строка " & (nr + 1) & " листа «" & SH_ORDERS & "» (заказано " & WmsIssue.QtyText(q) & ")."
            End If
        End If
    End If
    SupReturnPost = "OK:" & (r + 1) & "|" & IIf(gSuNewRow > 0, gSuNewRow + 1, 0)
    WmsUi.UiMessage(msg)
End Function

' ================================================================ З6 the repeat of a position or an order

Sub BtnRepeatOrder(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = RepeatOrder()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Повторить заказ»: the position under the cursor or its whole order — new rows of «Заказы» with the same goods, supplier
' and price, the quantity as ordered then (one position — asked), the order date today. Nothing is posted. OK:<rows> / …
Function RepeatOrder() As String
    Dim sh As Object, r As Long, r0 As Long, r1 As Long, k As Integer, i As Long, src As Long, q As Double, msg As String, why As String
    Dim v(0) As String, n As Long, first As Long, last As Long, nm As String, kind As String
    sh = WmsOrders.OrdersSheet()
    If gDoc.getCurrentController().getActiveSheet().getName() <> SH_ORDERS Then
        RepeatOrder = "ERR:встаньте на строку позиции на листе «" & SH_ORDERS & "»"
        gUiLastMsg = RepeatOrder
        Exit Function
    End If
    r = gDoc.getCurrentController().getSelection().getCellAddress().Row
    kind = WmsOrders.OrderRowKind(r)
    If r < 1 Or kind = "EMPTY" Or kind = "COPY" Or kind = "FOREIGN" Then
        RepeatOrder = "ERR:встаньте на строку позиции заказа (с наименованием и заказанным количеством)"
        gUiLastMsg = RepeatOrder
        Exit Function
    End If
    src = PositionRowOf(r)
    nm = sh.getCellByPosition(OC_NAME, src).getString()
    WmsDaily.OrderBlockOf(src, r0, r1)
    k = WmsPrimeUi.UiButtons("Повторить заказ", "Заказ «" & sh.getCellByPosition(OC_ORDER, r0).getString() & "» · позиция «" & nm & "»." & Chr(10) & Chr(10) _
        & "Новые строки «" & SH_ORDERS & "» внизу листа: тот же товар, поставщик и цена, дата заказа — сегодня. Ничего не проводится.", _
        Array("Эту позицию", "Весь заказ (позиций: " & PositionsIn(r0, r1) & ")", "Отмена"), Array(1, 2, 0))
    If k = 1 Then
        v(0) = WmsIssue.QtyText(OrderedOf(src))
        If Not WmsPrimeUi.UiForm("Повторить позицию", "«" & nm & "» — сколько заказать?", Array("Заказать (H)"), v, Array(""), "Создать строку") Then
            RepeatOrder = "SKIP:окно закрыто"
            gUiLastMsg = RepeatOrder
            Exit Function
        End If
        If Not WmsOrders.ParseQtyValueAs(v(0), "заказать", q, msg) Then
            RepeatOrder = "ERR:" & msg
            gUiLastMsg = RepeatOrder
            Exit Function
        End If
        If q <= 0 Then
            RepeatOrder = "ERR:количество должно быть больше 0"
            gUiLastMsg = RepeatOrder
            Exit Function
        End If
        first = NewOrderRowLike(src, q, "повтор позиции строки " & (src + 1), why)
        If first < 0 Then
            RepeatOrder = "ERR:" & why
            gUiLastMsg = RepeatOrder
            Exit Function
        End If
        last = first
        n = 1
    ElseIf k = 2 Then
        first = -1
        For i = r0 To r1
            If IsPositionRow(i) Then
                last = NewOrderRowLike(i, OrderedOf(i), "повтор заказа (строка " & (i + 1) & ")", why)
                If last < 0 Then
                    RepeatOrder = "ERR:" & why & IIf(n > 0, " (созданы строки " & (first + 1) & "–" & (first + n) & ")", "")
                    gUiLastMsg = RepeatOrder
                    Exit Function
                End If
                If first < 0 Then first = last
                n = n + 1
            End If
        Next i
    Else
        RepeatOrder = "SKIP:ничего не создано"
        gUiLastMsg = RepeatOrder
        Exit Function
    End If
    gDoc.getCurrentController().select(sh.getCellByPosition(OC_ORDQTY, first))
    RepeatOrder = "OK:" & n & "|" & (first + 1) & "-" & (last + 1)
    WmsUi.UiMessage("Готово: новых строк «" & SH_ORDERS & "» — " & n & " (строки " & (first + 1) & IIf(n > 1, "–" & (last + 1), "") & "). Проверьте количество и " _
        & "ожидаемую дату (Q); пока не пришло — они в «Ожидаем».")
    gUiLastMsg = RepeatOrder
End Function

' the source row of the position of row r (a row of «Ещё поступление» → the first row of its position); r itself otherwise
Function PositionRowOf(r As Long) As Long
    Dim mark As String, sh As Object, olid As Long, od As Variant, src As Long
    PositionRowOf = r
    sh = WmsOrders.OrdersSheet()
    mark = sh.getCellByPosition(OC_BLOCK, r).getString()
    If Left(mark, 2) <> "OL" Then Exit Function
    olid = CLng(Val(Mid(mark, 3)))
    If WmsOrders.OlValid(olid, od) Then
        src = WmsOrders.SourceRowOf(od)
        If src >= 1 Then PositionRowOf = src
    End If
End Function

' row i of «Заказы» is the first row of a position (a name and an ordered quantity, not a row of «Ещё поступление»)
Function IsPositionRow(i As Long) As Boolean
    Dim sh As Object
    sh = WmsOrders.OrdersSheet()
    If Left(sh.getCellByPosition(OC_BLOCK, i).getString(), 2) = "OL" Then Exit Function
    If Left(sh.getCellByPosition(OC_CTL, i).getString(), 5) = "КОПИЯ" Then Exit Function
    IsPositionRow = (Trim(sh.getCellByPosition(OC_NAME, i).getString()) <> "" And OrderedOf(i) > 0)
End Function

Private Function PositionsIn(r0 As Long, r1 As Long) As Long
    Dim i As Long
    For i = r0 To r1
        If IsPositionRow(i) Then PositionsIn = PositionsIn + 1
    Next i
End Function

' the ordered quantity (H) of row r of «Заказы»; 0 — none or not a number
Function OrderedOf(r As Long) As Double
    Dim q As Double, msg As String, c As Object
    c = WmsOrders.OrdersSheet().getCellByPosition(OC_ORDQTY, r)
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then Exit Function
    If WmsOrders.ParseQtyAs(c, "заказано (H)", q, msg) Then OrderedOf = q
End Function

' ================================================================ З5 the whole order cancelled, with the reason

Sub BtnCancelOrderAll(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = CancelOrderAll()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Отменить весь заказ»: every position of the order of the row under the cursor — not received: «Отменено»; received in part:
' the rest is cancelled; received in full — stays. One reason for all, in the journal of every cancellation. OK:<counts> / …
Function CancelOrderAll() As String
    Dim sh As Object, r As Long, r0 As Long, r1 As Long, i As Long, res As String, nC As Long, nR As Long, nKeep As Long, nErr As Long, first As String
    Dim reason As String, kind As String, od As Variant, nPos As Long, msg As String
    sh = WmsOrders.OrdersSheet()
    If gDoc.getCurrentController().getActiveSheet().getName() <> SH_ORDERS Then
        CancelOrderAll = "ERR:встаньте на строку заказа на листе «" & SH_ORDERS & "»"
        gUiLastMsg = CancelOrderAll
        Exit Function
    End If
    r = gDoc.getCurrentController().getSelection().getCellAddress().Row
    If r < 1 Or Not WmsOrders.RowHasInput(r) Then
        CancelOrderAll = "ERR:встаньте на строку заказа"
        gUiLastMsg = CancelOrderAll
        Exit Function
    End If
    WmsDaily.OrderBlockOf(PositionRowOf(r), r0, r1)
    nPos = PositionsIn(r0, r1)
    If nPos = 0 Then
        CancelOrderAll = "ERR:в заказе нет позиций с заказанным количеством"
        gUiLastMsg = CancelOrderAll
        Exit Function
    End If
    If Not WmsPrimeUi.ReasonAsk("Отменить весь заказ", "Заказ «" & sh.getCellByPosition(OC_ORDER, r0).getString() & "» (строки " & (r0 + 1) & "–" & (r1 + 1) _
        & ", позиций " & nPos & "): не полученные позиции получат статус «Отменено», у частично полученных отменится остаток; полученные полностью " _
        & "не меняются. Строки останутся в истории.", "Отменить заказ", reason) Then
        CancelOrderAll = "SKIP:ничего не отменено"
        gUiLastMsg = CancelOrderAll
        Exit Function
    End If
    gCancelReason = IIf(reason <> "", reason, "заказ отменён целиком")
    For i = r0 To r1
        If IsPositionRow(i) Then
            kind = WmsOrders.OrderRowKind(i)
            res = ""
            If kind = "OPEN" Then
                res = WmsReceipt.OrderCancelRow(i)
                If Left(res, 3) = "OK:" Then nC = nC + 1
            ElseIf kind = "RECEIVED" Or kind = "STORNO" Then
                If WmsOrders.OlValid(gKol, od) Then
                    If od(OD_RCV) <= 0.0000001 And CStr(od(OD_CANCEL)) <> OD_CANCEL_ORDER Then
                        res = WmsReceipt.OrderCancelRow(i)
                        If Left(res, 3) = "OK:" Then nC = nC + 1
                    ElseIf od(OD_RCV) < od(OD_ORD) - 0.0000001 And CStr(od(OD_CANCEL)) = "" Then
                        res = WmsReceipt.OrderCancelRestRow(i)
                        If Left(res, 3) = "OK:" Then nR = nR + 1
                    Else
                        nKeep = nKeep + 1
                    End If
                End If
            Else
                nKeep = nKeep + 1
            End If
            If res <> "" And Left(res, 3) <> "OK:" Then
                If Left(res, 5) = "SKIP:" Then
                    nKeep = nKeep + 1
                Else
                    nErr = nErr + 1
                    If first = "" Then first = "строка " & (i + 1) & ": " & Mid(res, InStr(res, ":") + 1)
                    If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Then Exit For
                End If
            End If
        End If
    Next i
    gCancelReason = ""
    CancelOrderAll = "OK:" & nC & "|" & nR & "|" & nKeep & "|" & nErr
    msg = "Заказ «" & sh.getCellByPosition(OC_ORDER, r0).getString() & "»: отменено позиций " & nC & ", отменён остаток у " & nR & ", не изменено " & nKeep _
        & IIf(nErr > 0, ", отказ " & nErr & Chr(10) & "Первая причина: " & first, "") & "." & Chr(10) & "Причина: " & IIf(reason <> "", reason, "заказ отменён целиком") & "."
    If nErr > 0 Then WmsUi.UiMessage(msg) Else WmsOrdersUi.UiInfo(msg)
    gUiLastMsg = CancelOrderAll
End Function

' ================================================================ the page «Снабжение»: «Ожидаем» (З2) and «Нужно заказать» (З10)

Public Const SH_SUPPLY = "Снабжение"
Public Const SH_GOODS = "Номенклатура"
Private Const SP_INPUT_ROW = 3
Private Const SP_INPUT_COL = 2
Private Const SP_HEAD = 6
Private Const SP_INFO = 7
Private Const SP_TABLE_HEAD = 9
Private Const SP_TABLE = 10
Private Const SP_MAX = 3000
' hidden columns of the table: the sheet and the row of the source, the mark of the row (G — a group, L — late)
Public Const SP_SRC_SHEET = 12
Public Const SP_SRC_ROW = 13
Public Const SP_MARK = 14
' «Номенклатура»: B name, C article, D unit, H minimum, I batch, J supplier; M..P written by «Обновить»
Private Const GD_CODE = 0
Public Const GD_NAME = 1
Public Const GD_ART = 2
Public Const GD_UNIT = 3
Public Const GD_CAT = 4
Public Const GD_PACKN = 5
Public Const GD_PACKU = 6
Private Const GD_MIN = 7
Private Const GD_BATCH = 8
Private Const GD_SUP = 9
Public Const GD_PHOTO = 10
Private Const GD_NOTE = 11
Public Const GD_STOCK = 12
Private Const GD_NEI = 13
Private Const GD_TRANSIT = 14
Public Const GD_UPD = 15

' the mode of the page shown (WAIT / NEED) and the open positions of the last «Ожидаем»: the row of «Заказы», the key of the
' goods (UCase article or name), the rest still expected
Global gSpMode As String
Global gSpBusy As Boolean
Global gSpN As Long
Global gSpRow() As Long
Global gSpKey() As String
Global gSpRest() As Double

Sub NavSupply(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_SUPPLY, SP_INPUT_ROW, SP_INPUT_COL)
    gUiLastMsg = "OK:" & SH_SUPPLY
End Sub

Sub NavSuppliers(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_SUPPLIERS, 1, 0)
    gUiLastMsg = "OK:" & SH_SUPPLIERS
End Sub

Sub NavGoods(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_GOODS, 1, 1)
    gUiLastMsg = "OK:" & SH_GOODS
End Sub

Sub BtnWait(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = WaitShow(SupplyFilter())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub BtnNeed(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = NeedShow(SupplyFilter())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Private Function SupplyFilter() As String
    SupplyFilter = Trim(gDoc.Sheets.getByName(SH_SUPPLY).getCellByPosition(SP_INPUT_COL, SP_INPUT_ROW).getString())
End Function

' Enter in the field of the supplier: the page of the mode shown, for that supplier
Sub OnSupplyChange(Optional oTarget As Variant)
    On Error GoTo EH
    If gSpBusy Then Exit Sub
    If oTarget.getCellAddress().Row <> SP_INPUT_ROW Or oTarget.getCellAddress().Column <> SP_INPUT_COL Then Exit Sub
    WmsInit()
    If gSpMode = "NEED" Then NeedShow(SupplyFilter()) Else WaitShow(SupplyFilter())
EH:
End Sub

' a double click on a line: its row of «Заказы» or «Номенклатура». True: no editing of the cell
Function OnSupplyDoubleClick(Optional oTarget As Variant) As Boolean
    Dim r As Long
    On Error GoTo EH
    WmsInit()
    r = oTarget.getCellAddress().Row
    If r < SP_TABLE Then Exit Function
    OnSupplyDoubleClick = True
    SupplyGo(r)
EH:
End Function

Sub BtnSupplyGo(Optional oEvent As Variant)
    Dim r As Long
    WmsInit()
    r = WmsView.ViewRow(SH_SUPPLY)
    If r < SP_TABLE Or Not SupplyGo(r) Then WmsUi.UiMessage("Встаньте на строку позиции и нажмите «Перейти к строке» — откроется её строка «" & SH_ORDERS & "»")
End Sub

Private Function SupplyGo(r As Long) As Boolean
    Dim sh As Object, s As String
    sh = gDoc.Sheets.getByName(SH_SUPPLY)
    s = sh.getCellByPosition(SP_SRC_SHEET, r).getString()
    If s = "" Then Exit Function
    WmsView.GoCell(s, CLng(sh.getCellByPosition(SP_SRC_ROW, r).getValue()), IIf(s = SH_ORDERS, OC_NAME, GD_NAME))
    gUiLastMsg = "OK:" & s & " " & (CLng(sh.getCellByPosition(SP_SRC_ROW, r).getValue()) + 1)
    SupplyGo = True
End Function

' the expression of the Calc engine «row of «Заказы» is a position still expected» (as «Дашборд» counts it), rows 1..last
Function OpenCond(last As Long) As String
    Dim rg As String, ei As String
    rg = WmsView.ColRng(SH_ORDERS, OC_STATUS, 1, last)
    ei = WmsView.ColRng(SH_ORDERS, OC_EI, 1, last)
    OpenCond = "(((" & ei & "<>"""")*((" & rg & "=""" & OS_WAITING & """)+(" & rg & "=""" & OS_PARTIAL & """)+(" & rg & "=""" & OS_OVERDUE & """)+(" & rg _
        & "=""" & OS_PARTIAL_OVERDUE & """))+(" & ei & "="""")*(" & rg & "<>""" & OS_CANCELLED & """)*(LEFT(" & WmsView.ColRng(SH_ORDERS, OC_CTL, 1, last) _
        & ";5)<>""КОПИЯ"")*(LEFT(" & WmsView.ColRng(SH_ORDERS, OC_BLOCK, 1, last) & ";2)<>""OL"")*(" & WmsView.ColRng(SH_ORDERS, OC_ORDQTY, 1, last) _
        & "<>"""")*(" & WmsView.ColRng(SH_ORDERS, OC_NAME, 1, last) & "<>""""))>0)"
End Function

Private Function Esc(s As String) As String
    Esc = Replace(s, """", """""")
End Function

' the goods of a row: its article, or its name when there is none (upper case — as Calc compares)
Function GoodsKey(sArt As String, sName As String) As String
    If Trim(sArt) <> "" Then GoodsKey = "A|" & UCase(Trim(sArt)) Else GoodsKey = "N|" & UCase(Trim(sName))
End Function

' 1.0 (З7): the hint of «Заказы» Y for open row r — the same goods of the same supplier already expected on another row; ""
Function WaitingHint(r As Long) As String
    Dim sh As Object, sup As String, art As String, nm As String, last As Long, cond As String, x As String, k As Long, ok As Boolean
    On Error GoTo EH
    sh = WmsOrders.OrdersSheet()
    sup = Trim(sh.getCellByPosition(OC_SUPPLIER, r).getString())
    art = Trim(sh.getCellByPosition(OC_ART, r).getString())
    nm = Trim(sh.getCellByPosition(OC_NAME, r).getString())
    If sup = "" Or (art = "" And nm = "") Then Exit Function
    last = WmsOrders.LastRow(sh)
    If last < 2 Then Exit Function
    cond = OpenCond(last) & "*(" & WmsView.ColRng(SH_ORDERS, OC_SUPPLIER, 1, last) & "=""" & Esc(sup) & """)*(ROW(" _
        & WmsView.ColRng(SH_ORDERS, OC_SUPPLIER, 1, last) & ")<>" & (r + 1) & ")"
    If art <> "" Then
        cond = cond & "*(" & WmsView.ColRng(SH_ORDERS, OC_ART, 1, last) & "=""" & Esc(art) & """)"
    Else
        cond = cond & "*(" & WmsView.ColRng(SH_ORDERS, OC_NAME, 1, last) & "=""" & Esc(nm) & """)"
    End If
    x = WmsOrders.ScratchEval("IFERROR(MATCH(1;" & cond & ";0);0)", ok)
    If Not ok Then Exit Function
    k = CLng(Val(x))
    If k < 1 Then Exit Function
    WaitingHint = "Внимание: «" & IIf(nm <> "", nm, art) & "» у «" & sup & "» уже ожидается — строка " & (k + 1) & " (заказ «" _
        & sh.getCellByPosition(OC_ORDER, k).getString() & "», заказано " & sh.getCellByPosition(OC_ORDQTY, k).getString() & "). Нужен ещё один заказ?"
    Exit Function
EH:
    WaitingHint = ""
End Function

' the open positions of «Заказы» (sSup — one supplier, "" — all): gSpN rows in gSpRow, the goods in gSpKey, the rest in gSpRest
Function OpenPositions(sSup As String) As String
    Dim sh As Object, last As Long, cond As String, a As Variant, nItems As Long, i As Long, r As Long, d As Variant, n As Long, rv As Variant
    Dim od As Variant, ord As Double, rcv As Double, msg As String, canon As String
    gSpN = 0
    ReDim gSpRow(0)
    ReDim gSpKey(0)
    ReDim gSpRest(0)
    gVwEngFail = ""
    sh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    cond = OpenCond(last)
    If sSup <> "" Then cond = cond & "*(" & WmsView.ColRng(SH_ORDERS, OC_SUPPLIER, 1, last) & "=""" & Esc(sSup) & """)"
    a = WmsView.EngList("TEXTJOIN("";"";1;IF(" & cond & ";ROW(" & WmsView.ColRng(SH_ORDERS, OC_NAME, 1, last) & ")-1;""""))", nItems)
    If gVwEngFail <> "" Then
        OpenPositions = "ERR:" & gVwEngFail
        Exit Function
    End If
    If nItems = 0 Then Exit Function
    ReDim gSpRow(nItems - 1)
    ReDim gSpKey(nItems - 1)
    ReDim gSpRest(nItems - 1)
    For i = 0 To nItems - 1
        r = CLng(a(i))
        d = sh.getCellRangeByPosition(0, r, OC_EI, r).getDataArray()(0)
        ord = 0
        rcv = 0
        canon = CStr(d(OC_EI))
        If canon <> "" Then
            If WmsOrders.StrictEI(canon, n) Then
                If WmsOrders.RcvRegistered(n, rv) Then
                    If WmsOrders.OlValid(CLng(rv(RV_OL)), od) Then
                        ord = od(OD_ORD)
                        rcv = od(OD_RCV)
                    End If
                End If
            End If
        Else
            If Not WmsOrders.ParseQtyValueAs(d(OC_ORDQTY), "заказано", ord, msg) Then ord = 0
        End If
        gSpRow(i) = r
        gSpKey(i) = GoodsKey(CStr(d(OC_ART)), CStr(d(OC_NAME)))
        gSpRest(i) = WmsIssue.Round3(ord - rcv)
        If gSpRest(i) < 0 Then gSpRest(i) = 0
    Next i
    gSpN = nItems
End Function

' «Ожидаем»: the positions still expected, grouped by supplier (the late ones first in a group). OK:<n> / ERR:…
Function WaitShow(sSup As String) As String
    Dim sh As Object, osh As Object, res As String, i As Long, j As Long, d As Variant, today As Double, q As Double, late As Long, nLate As Long
    Dim sups() As String, nS As Long, sIdx As Object, k As Long, key() As Double, ord() As Long, rows() As Variant, nOut As Long, g As Long, gPrev As Long
    Dim rcv As Double, ordered As Double, phone As String, rs As Long, nGroup As Long, sRest As Double, gLate As Long, pos As Long, nm As String
    gSpBusy = True
    On Error GoTo EH
    res = OpenPositions(sSup)
    If Left(res, 4) = "ERR:" Then
        WaitShow = res
        gSpBusy = False
        Exit Function
    End If
    osh = WmsOrders.OrdersSheet()
    today = WmsOrders.Today()
    ' the suppliers in alphabetical order
    sIdx = New Collection
    ReDim sups(IIf(gSpN > 0, gSpN - 1, 0))
    ReDim key(IIf(gSpN > 0, gSpN - 1, 0))
    ReDim ord(IIf(gSpN > 0, gSpN - 1, 0))
    For i = 0 To gSpN - 1
        nm = Trim(osh.getCellByPosition(OC_SUPPLIER, gSpRow(i)).getString())
        If CollGet(sIdx, SupKey(nm)) < 0 Then
            CollPut(sIdx, nS, SupKey(nm))
            sups(nS) = nm
            nS = nS + 1
        End If
    Next i
    SortStrings(sups, nS)
    sIdx = New Collection
    For i = 0 To nS - 1
        CollPut(sIdx, i, SupKey(sups(i)))
    Next i
    For i = 0 To gSpN - 1
        g = CollGet(sIdx, SupKey(osh.getCellByPosition(OC_SUPPLIER, gSpRow(i)).getString()))
        q = WmsOrders.QSerial(osh.getCellByPosition(OC_EDATE, gSpRow(i)).getDataArray()(0)(0))
        key(i) = g * 1E9 + IIf(q > 0, q - 30000, 99999) * 10000 + (i Mod 10000)
        ord(i) = i
    Next i
    WmsView.SortIdx(key, ord, gSpN, False)
    ReDim rows(gSpN + nS + 1)
    gPrev = -1
    For j = 0 To gSpN - 1
        i = ord(j)
        If nOut > SP_MAX Then Exit For
        g = Int(key(i) / 1E9)
        d = osh.getCellRangeByPosition(0, gSpRow(i), OC_EI, gSpRow(i)).getDataArray()(0)
        If g <> gPrev Then
            ' the line of the group: the supplier, its phone, how many positions, how many late
            gPrev = g
            nGroup = 0
            gLate = 0
            For k = j To gSpN - 1
                If Int(key(ord(k)) / 1E9) <> g Then Exit For
                nGroup = nGroup + 1
                q = WmsOrders.QSerial(osh.getCellByPosition(OC_EDATE, gSpRow(ord(k))).getDataArray()(0)(0))
                If q > 0 And q < today Then gLate = gLate + 1
            Next k
            phone = ""
            rs = SupplierRow(sups(g))
            If rs > 0 Then phone = Trim(gDoc.Sheets.getByName(SH_SUPPLIERS).getCellByPosition(SU_PHONE, rs).getString())
            rows(nOut) = Array(IIf(sups(g) <> "", sups(g), "поставщик не указан"), NumText(nGroup) & " " & Plural(nGroup, "позиция", "позиции", "позиций") _
                & IIf(gLate > 0, " · опаздывают " & gLate, "") & IIf(phone <> "", " · тел. " & phone, ""), "", "", "", "", "", "", "", "", "", "", "", "G")
            nOut = nOut + 1
        End If
        q = WmsOrders.QSerial(d(OC_EDATE))
        late = 0
        If q > 0 And q < today Then
            late = CLng(today - q)
            nLate = nLate + 1
        End If
        ordered = gSpRest(i)
        rcv = 0
        If CStr(d(OC_EI)) <> "" Then
            rows(nOut) = Array("заказ " & d(OC_ORDER), CStr(d(OC_NAME)), CStr(d(OC_ART)), "", "", gSpRest(i), CStr(d(OC_UNIT)), IIf(q > 0, ViewDay(q), "—"), _
                IIf(late > 0, late & " дн.", ""), "строка " & (gSpRow(i) + 1), "", SH_ORDERS, gSpRow(i), IIf(late > 0, "L", ""))
            PartFill(rows(nOut), CStr(d(OC_EI)), gSpRest(i))
        Else
            rows(nOut) = Array("заказ " & d(OC_ORDER), CStr(d(OC_NAME)), CStr(d(OC_ART)), gSpRest(i), 0, gSpRest(i), CStr(d(OC_UNIT)), IIf(q > 0, ViewDay(q), "—"), _
                IIf(late > 0, late & " дн.", ""), "строка " & (gSpRow(i) + 1), "", SH_ORDERS, gSpRow(i), IIf(late > 0, "L", ""))
        End If
        nOut = nOut + 1
    Next j
    sh = WmsView.ViewOpen(SH_SUPPLY)
    SupplyClear(sh)
    gSpMode = "WAIT"
    sh.getCellByPosition(1, SP_HEAD).setString("Ожидаем: " & NumText(gSpN) & " " & Plural(gSpN, "позиция", "позиции", "позиций") & IIf(sSup <> "", " от «" & sSup & "»", _
        " от " & NumText(nS) & " " & Plural(nS, "поставщика", "поставщиков", "поставщиков")) & IIf(nLate > 0, " · опаздывают " & NumText(nLate), ""))
    sh.getCellByPosition(1, SP_INFO).setString("Сначала опаздывающие и ближайшие по ожидаемой дате (Q). Пришло — «Принять…» или «Ещё поступление» на строке «" _
        & SH_ORDERS & "»; не придёт — «Отменить заказ» / «Отменить остаток».")
    sh.getCellByPosition(1, SP_INFO + 1).setString("Двойной щелчок по строке — строка «" & SH_ORDERS & "».")
    sh.getCellRangeByPosition(1, SP_TABLE_HEAD, 10, SP_TABLE_HEAD).setDataArray(Array(Array("Поставщик / заказ", "Наименование", "Артикул", "Заказано", _
        "Получено", "Осталось", "Ед.", "Ожидаем к", "Опоздание", "Строка")))
    If nOut > 0 Then
        sh.getCellRangeByPosition(1, SP_TABLE, SP_MARK, SP_TABLE + nOut - 1).setDataArray(SubRows(rows, nOut))
    Else
        sh.getCellByPosition(2, SP_TABLE).setString(IIf(sSup <> "", "от «" & sSup & "» ничего не ожидается", "ничего не ожидается — все заказы получены или отменены"))
    End If
    WmsView.ViewDone(sh)
    gSpBusy = False
    WaitShow = "OK:" & gSpN
    gUiLastMsg = WaitShow
    Exit Function
EH:
    WaitShow = "ERR:«Ожидаем» не построено: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = WaitShow
    On Error Resume Next
    WmsView.ViewDone(sh)
    gSpBusy = False
End Function

' the ordered and received quantities of a position with receipts (its EI on the source row)
Private Sub PartFill(row As Variant, canon As String, rest As Double)
    Dim n As Long, rv As Variant, od As Variant
    If Not WmsOrders.StrictEI(canon, n) Then Exit Sub
    If Not WmsOrders.RcvRegistered(n, rv) Then Exit Sub
    If Not WmsOrders.OlValid(CLng(rv(RV_OL)), od) Then Exit Sub
    row(3) = od(OD_ORD)
    row(4) = od(OD_RCV)
End Sub

Private Sub SupplyClear(sh As Object)
    sh.getCellRangeByPosition(1, SP_HEAD, SP_MARK, SP_TABLE + SP_MAX + 2).clearContents(1 + 2 + 4)
End Sub

Private Function SubRows(a As Variant, n As Long) As Variant
    Dim r() As Variant, i As Long
    ReDim r(n - 1)
    For i = 0 To n - 1
        r(i) = a(i)
    Next i
    SubRows = r
End Function

Private Sub SortStrings(a() As String, n As Long)
    Dim i As Long, j As Long, t As String
    For i = 1 To n - 1
        t = a(i)
        j = i - 1
        Do While j >= 0
            If UCase(a(j)) <= UCase(t) Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub

' ---------------------------------------------------------------- «Нужно заказать»

' a goods of «Номенклатура» (art / name) on «Наличие»: the balance of its EIs (a storno EI has 0) and how many EIs have one
Private Function GoodsStock(art As String, nm As String, ByRef nEI As Long) As Double
    Dim last As Long, m As String, a As Variant, rg As String
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last < 1 Then Exit Function
    If art <> "" Then m = "(" & WmsView.ColRng(SH_STOCK, SC_ART, 1, last) & "=""" & Esc(art) & """)" Else m = "(" & WmsView.ColRng(SH_STOCK, SC_NAME, 1, last) & "=""" & Esc(nm) & """)"
    rg = WmsView.ColRng(SH_STOCK, SC_QTY, 1, last)
    a = Split(WmsView.Eng("SUMPRODUCT(ISNUMBER(" & rg & ")*(" & rg & ">0.0000005)*" & m & "*IF(ISNUMBER(" & rg & ");" & rg & ";0))&""|""&SUMPRODUCT(ISNUMBER(" & rg _
        & ")*(" & rg & ">0.0000005)*" & m & ")"), "|")
    GoodsStock = WmsView.ListNum(a, 0)
    nEI = CLng(WmsView.ListNum(a, 1))
End Function

' the quantity issued (posted) of a goods in the last days
Private Function GoodsIssued(art As String, nm As String, days As Long) As Double
    Dim last As Long, m As String, today As Double
    last = WmsOrders.LastRow(WmsIssue.IssuesSheet())
    If last < 1 Then Exit Function
    today = WmsOrders.Today()
    If art <> "" Then m = "(" & WmsView.ColRng(SH_ISSUES, IC_ART, 1, last) & "=""" & Esc(art) & """)" Else m = "(" & WmsView.ColRng(SH_ISSUES, IC_NAME, 1, last) & "=""" & Esc(nm) & """)"
    GoodsIssued = WmsView.ListNum(Array(WmsView.Eng("SUMPRODUCT(" & WmsView.PostedExpr(WmsView.ColRng(SH_ISSUES, IC_CTL, 1, last)) & "*" _
        & WmsView.DaysExpr(WmsView.ColRng(SH_ISSUES, IC_DATE, 1, last), today - days + 1, today) & "*" & m & "*IF(ISNUMBER(" _
        & WmsView.ColRng(SH_ISSUES, IC_QTY, 1, last) & ");" & WmsView.ColRng(SH_ISSUES, IC_QTY, 1, last) & ";0))")), 0)
End Function

' the goods below their minimum (stock + in transit < «Минимум»): how many to order (up to the minimum, by the batch). The rows
' found are in gNdRow (the rows of «Номенклатура») and gNdQty. bShow — the page; False — only counted (the tile of «Дашборд»)
Global gNdN As Long
Global gNdRow() As Long
Global gNdQty() As Double

Function NeedCompute(sSup As String) As String
    Dim gsh As Object, last As Long, d As Variant, i As Long, k As Long, mn As Double, st As Double, tr As Double, need As Double, bt As Double, nEI As Long
    Dim transit As Object, tq() As Double, nT As Long, key As String, res As String
    gNdN = 0
    ReDim gNdRow(0)
    ReDim gNdQty(0)
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then Exit Function
    gsh = gDoc.Sheets.getByName(SH_GOODS)
    last = WmsOrders.LastRow(gsh)
    If last < 1 Then Exit Function
    If last > 20000 Then last = 20000
    d = gsh.getCellRangeByPosition(0, 1, GD_NOTE, last).getDataArray()
    ' in transit: the rest of the open positions by goods
    res = OpenPositions("")
    If Left(res, 4) = "ERR:" Then
        NeedCompute = res
        Exit Function
    End If
    transit = New Collection
    ReDim tq(IIf(gSpN > 0, gSpN - 1, 0))
    For i = 0 To gSpN - 1
        k = CollGet(transit, gSpKey(i))
        If k < 0 Then
            k = nT
            CollPut(transit, k, gSpKey(i))
            nT = nT + 1
        End If
        tq(k) = tq(k) + gSpRest(i)
    Next i
    ReDim gNdRow(UBound(d))
    ReDim gNdQty(UBound(d))
    For i = 0 To UBound(d)
        mn = 0
        If VarType(d(i)(GD_MIN)) = 5 Then mn = d(i)(GD_MIN)
        If mn > 0 And (Trim(CStr(d(i)(GD_ART))) <> "" Or Trim(CStr(d(i)(GD_NAME))) <> "") Then
            If sSup = "" Or SupKey(CStr(d(i)(GD_SUP))) = SupKey(sSup) Then
                st = GoodsStock(Trim(CStr(d(i)(GD_ART))), Trim(CStr(d(i)(GD_NAME))), nEI)
                k = CollGet(transit, GoodsKey(CStr(d(i)(GD_ART)), CStr(d(i)(GD_NAME))))
                tr = 0
                If k >= 0 Then tr = tq(k)
                need = WmsIssue.Round3(mn - st - tr)
                If need > 0.0000005 Then
                    bt = 0
                    If VarType(d(i)(GD_BATCH)) = 5 Then bt = d(i)(GD_BATCH)
                    If bt > 0 Then need = bt * (Int(need / bt - 0.0000001) + 1)
                    gNdRow(gNdN) = i + 1
                    gNdQty(gNdN) = need
                    gNdN = gNdN + 1
                End If
            End If
        End If
    Next i
    NeedCompute = "OK:" & gNdN
End Function

' «Нужно заказать»: the goods of «Номенклатура» below their minimum — the balance of all their EIs, what is in transit, the issues
' of 30 days, how many to order and from whom. OK:<n> / ERR:…
Function NeedShow(sSup As String) As String
    Dim sh As Object, gsh As Object, res As String, i As Long, d As Variant, rows() As Variant, nEI As Long, st As Double, tr As Double, used As Double
    Dim nm As String, art As String, k As Long, nNo As Long
    gSpBusy = True
    On Error GoTo EH
    res = NeedCompute(sSup)
    If Left(res, 4) = "ERR:" Then
        NeedShow = res
        gSpBusy = False
        Exit Function
    End If
    If gDoc.Sheets.hasByName(SH_GOODS) Then gsh = gDoc.Sheets.getByName(SH_GOODS)
    ReDim rows(IIf(gNdN > 0, gNdN - 1, 0))
    For i = 0 To gNdN - 1
        d = gsh.getCellRangeByPosition(0, gNdRow(i), GD_NOTE, gNdRow(i)).getDataArray()(0)
        nm = Trim(CStr(d(GD_NAME)))
        art = Trim(CStr(d(GD_ART)))
        st = GoodsStock(art, nm, nEI)
        used = GoodsIssued(art, nm, 30)
        tr = 0
        For k = 0 To gSpN - 1
            If gSpKey(k) = GoodsKey(art, nm) Then tr = tr + gSpRest(k)
        Next k
        rows(i) = Array(IIf(nm <> "", nm, art), art, st, tr, d(GD_MIN), used, gNdQty(i), CStr(d(GD_UNIT)), CStr(d(GD_SUP)), "строка " & (gNdRow(i) + 1), "", _
            SH_GOODS, gNdRow(i), IIf(st + tr <= 0, "L", ""))
        If st + tr <= 0 Then nNo = nNo + 1
    Next i
    If gDoc.Sheets.hasByName(SH_GOODS) Then
        If WmsOrders.LastRow(gsh) < 1 Then nNo = -1
    End If
    sh = WmsView.ViewOpen(SH_SUPPLY)
    SupplyClear(sh)
    gSpMode = "NEED"
    sh.getCellByPosition(1, SP_HEAD).setString("Нужно заказать: " & NumText(gNdN) & " " & Plural(gNdN, "товар", "товара", "товаров") & IIf(sSup <> "", " у «" & sSup & "»", "") _
        & IIf(nNo > 0, " · нет совсем " & NumText(nNo), ""))
    sh.getCellByPosition(1, SP_INFO).setString("Товары «" & SH_GOODS & "» с «Минимумом», у которых остаток всех ЕИ и то, что уже в пути, меньше минимума. «Заказать» — " _
        & "до минимума, кратно «Заказывать по».")
    sh.getCellByPosition(1, SP_INFO + 1).setString("Двойной щелчок — строка «" & SH_GOODS & "». Заказ: «" & SH_ORDERS & "» → новая строка (или «Повторить заказ» на прошлой).")
    sh.getCellRangeByPosition(1, SP_TABLE_HEAD, 10, SP_TABLE_HEAD).setDataArray(Array(Array("Товар", "Артикул", "Остаток", "В пути", "Минимум", "Выдано за 30 дн.", _
        "Заказать", "Ед.", "Поставщик", "Строка")))
    If gNdN > 0 Then
        sh.getCellRangeByPosition(1, SP_TABLE, SP_MARK, SP_TABLE + gNdN - 1).setDataArray(SubRows(rows, gNdN))
    ElseIf Not gDoc.Sheets.hasByName(SH_GOODS) Then
        sh.getCellByPosition(2, SP_TABLE).setString("в книге нет листа «" & SH_GOODS & "»")
    Else
        sh.getCellByPosition(2, SP_TABLE).setString("ниже минимума ничего нет — или в «" & SH_GOODS & "» не задан «Минимум» (колонка H)")
    End If
    WmsView.ViewDone(sh)
    gSpBusy = False
    NeedShow = "OK:" & gNdN
    gUiLastMsg = NeedShow
    Exit Function
EH:
    NeedShow = "ERR:«Нужно заказать» не построено: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = NeedShow
    On Error Resume Next
    WmsView.ViewDone(sh)
    gSpBusy = False
End Function

' ---------------------------------------------------------------- «Номенклатура»: «Обновить», «Пополнить из «Наличия»»

Sub BtnGoodsRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = GoodsRefresh()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsOrdersUi.UiInfo(Mid(s, 4))
End Sub

' M..P of every goods: the balance of its EIs, how many EIs have one, what is in transit, today. OK:<text> / ERR:…
Function GoodsRefresh() As String
    Dim gsh As Object, last As Long, d As Variant, i As Long, k As Long, out() As Variant, nEI As Long, st As Double, tr As Double, res As String, n As Long
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then
        GoodsRefresh = "ERR:в книге нет листа «" & SH_GOODS & "»"
        Exit Function
    End If
    gsh = gDoc.Sheets.getByName(SH_GOODS)
    last = WmsOrders.LastRow(gsh)
    If last < 1 Then
        GoodsRefresh = "OK:«" & SH_GOODS & "» пуста — впишите товары или нажмите «Пополнить из «Наличия»»"
        gUiLastMsg = GoodsRefresh
        Exit Function
    End If
    d = gsh.getCellRangeByPosition(0, 1, GD_NOTE, last).getDataArray()
    res = OpenPositions("")
    ReDim out(UBound(d))
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(GD_ART))) <> "" Or Trim(CStr(d(i)(GD_NAME))) <> "" Then
            st = GoodsStock(Trim(CStr(d(i)(GD_ART))), Trim(CStr(d(i)(GD_NAME))), nEI)
            tr = 0
            For k = 0 To gSpN - 1
                If gSpKey(k) = GoodsKey(CStr(d(i)(GD_ART)), CStr(d(i)(GD_NAME))) Then tr = tr + gSpRest(k)
            Next k
            out(i) = Array(st, nEI, tr, WmsOrders.Today())
            n = n + 1
        Else
            out(i) = Array("", "", "", "")
        End If
    Next i
    gsh.getCellRangeByPosition(GD_STOCK, 1, GD_UPD, last).setDataArray(out)
    GoodsRefresh = "OK:«" & SH_GOODS & "» обновлена: товаров " & n & " — остаток всех ЕИ, ЕИ с остатком, в пути."
    gUiLastMsg = GoodsRefresh
    Exit Function
EH:
    GoodsRefresh = "ERR:«" & SH_GOODS & "» не обновлена: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = GoodsRefresh
End Function

Sub BtnGoodsFill(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = GoodsFill()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsOrdersUi.UiInfo(Mid(s, 4))
End Sub

' «Пополнить из «Наличия»»: every goods of «Наличие» (an article, or a name without one) not yet in «Номенклатура» is added below —
' the name, the article, the unit, the category, the packing of its last EI. OK:<text> / ERR:…
Function GoodsFill() As String
    Dim gsh As Object, known As Object, last As Long, d As Variant, i As Long, k As String, add() As Variant, nAdd As Long, r0 As Long, r1 As Long
    Dim st As Object, lastS As Long, pu As String, pn As Double, idx As Object, j As Long, row As Variant
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then
        GoodsFill = "ERR:в книге нет листа «" & SH_GOODS & "»"
        Exit Function
    End If
    gsh = gDoc.Sheets.getByName(SH_GOODS)
    known = New Collection
    last = WmsOrders.LastRow(gsh)
    If last >= 1 Then
        d = gsh.getCellRangeByPosition(0, 1, GD_ART, last).getDataArray()
        For i = 0 To UBound(d)
            If Trim(CStr(d(i)(GD_ART))) <> "" Or Trim(CStr(d(i)(GD_NAME))) <> "" Then CollPut(known, 1, GoodsKey(CStr(d(i)(GD_ART)), CStr(d(i)(GD_NAME))))
        Next i
    End If
    st = WmsIssue.StockSheet()
    lastS = CLng(SysNum(SK_NEXT_EI)) - 1
    ReDim add(200)
    idx = New Collection
    For r0 = 1 To lastS Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > lastS Then r1 = lastS
        d = st.getCellRangeByPosition(0, r0, SC_STATE, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(d(i)(SC_STATE)) <> EI_ST_STORNO And (Trim(CStr(d(i)(SC_ART))) <> "" Or Trim(CStr(d(i)(SC_NAME))) <> "") Then
                k = GoodsKey(CStr(d(i)(SC_ART)), CStr(d(i)(SC_NAME)))
                If CollGet(known, k) < 0 Then
                    j = CollGet(idx, k)
                    pn = 0
                    pu = ""
                    WmsPack.PackOf(r0 + i, pu, pn)
                    row = Array("", Trim(CStr(d(i)(SC_NAME))), Trim(CStr(d(i)(SC_ART))), Trim(CStr(d(i)(SC_UNIT))), Trim(CStr(d(i)(SC_CAT))), IIf(pn > 0, pn, ""), pu)
                    If j < 0 Then
                        If nAdd > UBound(add) Then
                            ReDim Preserve add(2 * nAdd)
                        End If
                        CollPut(idx, nAdd, k)
                        add(nAdd) = row
                        nAdd = nAdd + 1
                    Else
                        ' the last EI of the goods gives the card (the newest packing)
                        add(j) = row
                    End If
                End If
            End If
        Next i
    Next r0
    If nAdd = 0 Then
        GoodsFill = "OK:все товары «" & SH_STOCK & "» уже есть в «" & SH_GOODS & "»"
        gUiLastMsg = GoodsFill
        Exit Function
    End If
    ReDim Preserve add(nAdd - 1)
    gsh.getCellRangeByPosition(0, last + 1, GD_PACKU, last + nAdd).setDataArray(add)
    GoodsFill = "OK:в «" & SH_GOODS & "» добавлено товаров: " & nAdd & " (строки " & (last + 2) & "–" & (last + nAdd + 1) & "). Впишите «Минимум» тем, " _
        & "что нельзя допускать до нуля, — они появятся в «Нужно заказать»."
    gUiLastMsg = GoodsFill
    Exit Function
EH:
    GoodsFill = "ERR:товары не добавлены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = GoodsFill
End Function

' ================================================================ З9 the vehicle and its receipts

Sub BtnCarAccept(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = CarAccept()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Принять товар» on «Приход авто»: the visit under the cursor (or the supplier typed on the panel) — «Ожидаем» of its supplier,
' and how many positions of that supplier were received on the day of the visit. OK:… / ERR:…
Function CarAccept() As String
    Dim csh As Object, r As Long, sup As String, dt As Double, nRcv As Long, last As Long, s As String, sh As Object, res As String
    csh = gDoc.Sheets.getByName(SH_CARS)
    r = WmsView.ViewRow(SH_CARS)
    If r >= CAR_FIRST Then
        sup = Trim(csh.getCellByPosition(CC_SUP, r).getString())
        If csh.getCellByPosition(CC_DATE, r).getType() = com.sun.star.table.CellContentType.VALUE Then dt = Int(csh.getCellByPosition(CC_DATE, r).getValue())
    End If
    If sup = "" Then sup = Trim(csh.getCellByPosition(CP_SUP_COL, CP_IN_ROW).getString())
    If sup = "" Then
        CarAccept = "ERR:встаньте на строку визита машины (или впишите поставщика на панели) и нажмите «Принять товар» — откроется, что ждём от этого поставщика"
        gUiLastMsg = CarAccept
        Exit Function
    End If
    If dt = 0 Then dt = WmsOrders.Today()
    nRcv = ReceivedOn(sup, dt)
    gDoc.Sheets.getByName(SH_SUPPLY).getCellByPosition(SP_INPUT_COL, SP_INPUT_ROW).setString(sup)
    res = WaitShow(sup)
    If Left(res, 4) = "ERR:" Then
        CarAccept = res
        gUiLastMsg = res
        Exit Function
    End If
    sh = WmsView.ViewOpen(SH_SUPPLY)
    sh.getCellByPosition(1, SP_INFO).setString("Машина «" & sup & "» " & ViewDay(dt) & ": принято позиций " & nRcv & ". Пришло — встаньте на позицию, двойной " _
        & "щелчок → строка «" & SH_ORDERS & "» → «Принять…».")
    WmsView.ViewDone(sh)
    WmsView.GoCell(SH_SUPPLY, SP_TABLE, 1)
    CarAccept = "OK:" & sup & "|принято " & nRcv & "|ждём " & gSpN
    gUiLastMsg = CarAccept
End Function

' how many receipts of supplier sup were posted with the receipt date dt («Заказы» L, N; a row with its EI)
Function ReceivedOn(sup As String, dt As Double) As Long
    Dim last As Long
    last = WmsOrders.LastRow(WmsOrders.OrdersSheet())
    If last < 1 Then Exit Function
    ReceivedOn = CLng(WmsView.ListNum(Array(WmsView.Eng("SUMPRODUCT((" & WmsView.ColRng(SH_ORDERS, OC_SUPPLIER, 1, last) & "=""" & Esc(sup) & """)*(" _
        & WmsView.ColRng(SH_ORDERS, OC_EI, 1, last) & "<>"""")*(LEFT(" & WmsView.ColRng(SH_ORDERS, OC_STATUS, 1, last) & ";14)<>""Поступление уд"")*" _
        & WmsView.DaysExpr(WmsView.ColRng(SH_ORDERS, OC_RDATE, 1, last), dt, dt) & ")")), 0))
End Function

' the row (0-based) of «Заказы» the last return window wrote for the replacement (0 — none)
Function TestSupNewRow() As Long
    TestSupNewRow = gSuNewRow
End Function
