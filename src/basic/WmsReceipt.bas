' WmsReceipt — обычный приход по заказу (задание Core Phase 3; MASTER SPEC v0.3 §4–§6, §12, §14–§18; D-005…D-007, D-033).
' «Провести приход» (первый приход позиции в её исходной строке), «Ещё поступление» (следующий УПД — новая строка той же
' позиции), «Исправить» (RECEIPT_FIX, тот же ЕИ), «Удалить» (RECEIPT_DEL, сторно), «Отменить заказ» (ORDER_CANCEL, без ЕИ),
' «Отменить остаток» (ORDER_CANCEL_REST), «Проверить». Каждая операция — одна запись журнала через WmsCore.ApplyOperation:
' новый ЕИ (NEXT_EI), его строка «Наличие», запись прихода (_RCV), позиция заказа (_ORD) и строка «Заказы» меняются вместе
' или не меняются вовсе. Порядок записей в плане: _SYS → «Наличие» → _RCV → _ORD → «Заказы».
Option Explicit

' validated receipt values (filled by the checks)
Private mFact As Double
Private mHasDocQty As Boolean
Private mDocQty As Double
Private mDoc As String
Private mHasDdate As Boolean
Private mDdate As Double
Private mRdate As Double
Private mPlace As String
Private mHasPrice As Boolean
Private mPrice As Double
' validated order position (an open source row)
Private mOrdQty As Double
Private mHasOdate As Boolean
Private mOdate As Double
Private mHasEdate As Boolean
Private mEdate As Double
Private mName As String
Private mUnit As String

' ================================================================ helpers

Private Function OSh() As Object
    OSh = WmsOrders.OrdersSheet()
End Function

Private Function Txt(r As Long, c As Integer) As String
    Txt = WmsOrders.OrdersSheet().getCellByPosition(c, r).getString()
End Function

Private Function NumOrEmpty(v As Variant) As Variant
    If VarType(v) = 5 Then NumOrEmpty = v Else NumOrEmpty = ""
End Function

Private Function IsoDate(d As Double) As String
    IsoDate = Format(d, "YYYY-MM-DD")
End Function

' ================================================================ the row of «Ещё поступление» inside the block of its order (M6 §18)
' The new delivery row is placed right under its source row and the earlier delivery rows of the same position, so the
' order stays one visual block. A row is inserted there before the operation starts (operations themselves only write
' cells); the journal names it (INSROW) and the recovery inserts it again before the writes are replayed
' (WmsRecovery.ReplayInsertRow). When the operation does not happen, the inserted row is removed again.

' True when row r of «Заказы» holds nothing in A..AC
Function OrderRowEmpty(r As Long) As Boolean
    Dim d As Variant, i As Integer
    d = OSh().getCellRangeByPosition(0, r, OC_BLOCK, r).getDataArray()(0)
    For i = 0 To OC_BLOCK
        If CStr(d(i)) <> "" Then Exit Function
    Next i
    OrderRowEmpty = True
End Function

' the row under the source row src (0-based) and the delivery rows of position olid that follow it: where the next
' delivery row of the position goes
Function DeliveryRowAfter(src As Long, olid As Long) As Long
    Dim r As Long, sh As Object, mark As String
    sh = OSh()
    mark = "OL" & olid
    r = src + 1
    Do While r < MAX_SHEET_ROW
        If sh.getCellByPosition(OC_BLOCK, r).getString() <> mark Then Exit Do
        r = r + 1
    Loop
    DeliveryRowAfter = r
End Function

' inserts an empty row at r of «Заказы» (the rows from r move down); "" or the error. The sheet allows the user to insert
' rows too; the protection is lifted only for the moment of the insertion.
Function InsertOrderRow(r As Long) As String
    Dim sh As Object, wasProt As Boolean
    On Error GoTo EH
    sh = OSh()
    If Not OrderRowEmpty(MAX_SHEET_ROW) Then
        InsertOrderRow = "на листе «" & SH_ORDERS & "» нет места для новой строки (последняя строка листа занята)"
        Exit Function
    End If
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    sh.getRows().insertByIndex(r, 1)
    If wasProt Then sh.protect(PROTECT_PWD)
    InsertOrderRow = ""
    Exit Function
EH:
    InsertOrderRow = "не удалось вставить строку в «" & SH_ORDERS & "»: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
End Function

' removes row r of «Заказы» when it is still empty (a delivery that did not happen); True when removed
Private Function RemoveEmptyOrderRow(r As Long) As Boolean
    Dim sh As Object, wasProt As Boolean
    On Error GoTo EH
    If Not OrderRowEmpty(r) Then Exit Function
    sh = OSh()
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    sh.getRows().removeByIndex(r, 1)
    If wasProt Then sh.protect(PROTECT_PWD)
    RemoveEmptyOrderRow = True
    Exit Function
EH:
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
End Function

' «Контроль» of an open row (a result or a problem); rows with a receipt are written only by operations
Private Sub SetCtl(r As Long, s As String)
    WmsOrders.SetIfDiff(OSh().getCellByPosition(OC_CTL, r), s)
End Sub

' documents of a receipt: number (C) and date (O) both present
Private Function DocsMissing(sDoc As String, bHasDdate As Boolean) As Integer
    If Trim(sDoc) <> "" And bHasDdate Then DocsMissing = 0 Else DocsMissing = 1
End Function

' ================================================================ checks (phase 1: no change of the document)

' A B H I L P Q of an open order row; "" or the reason
Private Function CheckOrderPart(r As Long, bNeedUnit As Boolean) As String
    Dim sh As Object, msg As String, sp As String, c As Object
    sh = OSh()
    If Trim(Txt(r, OC_ORDER)) = "" Then
        CheckOrderPart = "не указан № заказа (A)"
        Exit Function
    End If
    If Trim(Txt(r, OC_NAME)) = "" Then
        CheckOrderPart = "не указано наименование товара (B)"
        Exit Function
    End If
    If Not WmsOrders.ParseQtyAs(sh.getCellByPosition(OC_ORDQTY, r), "Заказанное количество (H)", mOrdQty, msg) Then
        CheckOrderPart = msg
        Exit Function
    End If
    If bNeedUnit And Trim(Txt(r, OC_UNIT)) = "" Then
        CheckOrderPart = "не указана единица измерения (I)"
        Exit Function
    End If
    sp = WmsOrders.SpecialSupplier(Txt(r, OC_SUPPLIER))
    If sp <> "" Then
        CheckOrderPart = "«" & sp & "» — специальный приход: он проводится на листе «" & SH_SPECIAL & "» (тип прихода «" & sp & "»), " _
            & "обычным приходом «Заказов» не проводится"
        Exit Function
    End If
    mHasOdate = False
    c = sh.getCellByPosition(OC_ODATE, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not WmsOrders.ParseDateCellAs(c, "дата заказа (P)", mOdate, msg) Then
            CheckOrderPart = msg
            Exit Function
        End If
        mHasOdate = True
    End If
    mHasEdate = False
    c = sh.getCellByPosition(OC_EDATE, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not WmsOrders.ParseDateCellAs(c, "ожидаемая дата поступления (Q)", mEdate, msg) Then
            CheckOrderPart = msg
            Exit Function
        End If
        mHasEdate = True
    End If
    mName = Trim(Txt(r, OC_NAME))
    mUnit = Trim(Txt(r, OC_UNIT))
    CheckOrderPart = ""
End Function

' F G C O N U J of an open row being received; "" or the reason
Private Function CheckReceiptCells(r As Long) As String
    Dim sh As Object, msg As String, c As Object
    sh = OSh()
    If Not WmsOrders.ParseQtyAs(sh.getCellByPosition(OC_FACT, r), "Фактическое количество (F)", mFact, msg) Then
        CheckReceiptCells = msg
        Exit Function
    End If
    mHasDocQty = False
    c = sh.getCellByPosition(OC_DOCQTY, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not WmsOrders.ParseQtyAs(c, "Количество по документу (G)", mDocQty, msg) Then
            CheckReceiptCells = msg
            Exit Function
        End If
        mHasDocQty = True
    End If
    mDoc = Trim(Txt(r, OC_DOC))
    mHasDdate = False
    c = sh.getCellByPosition(OC_DDATE, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not WmsOrders.ParseDateCellAs(c, "дата документа (O)", mDdate, msg) Then
            CheckReceiptCells = msg
            Exit Function
        End If
        mHasDdate = True
    End If
    If Not WmsOrders.ParseDateCellAs(sh.getCellByPosition(OC_RDATE, r), "дата поступления (N)", mRdate, msg) Then
        CheckReceiptCells = msg
        Exit Function
    End If
    mPlace = Trim(Txt(r, OC_PLACE))
    If mPlace = "" Then
        CheckReceiptCells = "не указано место хранения (U)"
        Exit Function
    End If
    If Not WmsOrders.ParsePriceCell(sh.getCellByPosition(OC_PRICE, r), mPrice, mHasPrice, msg) Then
        CheckReceiptCells = msg
        Exit Function
    End If
    CheckReceiptCells = ""
End Function

' the same fields given as values (dialog text; a number in tests); "" or the reason
Private Function CheckReceiptValues(vFact As Variant, vDocQty As Variant, vDoc As Variant, vDdate As Variant, vRdate As Variant, _
    vPlace As Variant, vPrice As Variant) As String
    Dim msg As String
    If Not WmsOrders.ParseQtyValueAs(vFact, "Фактическое количество (F)", mFact, msg) Then
        CheckReceiptValues = msg
        Exit Function
    End If
    mHasDocQty = False
    If Trim(CStr(vDocQty)) <> "" Then
        If Not WmsOrders.ParseQtyValueAs(vDocQty, "Количество по документу (G)", mDocQty, msg) Then
            CheckReceiptValues = msg
            Exit Function
        End If
        mHasDocQty = True
    End If
    mDoc = Trim(CStr(vDoc))
    mHasDdate = False
    If Trim(CStr(vDdate)) <> "" Then
        If Not WmsOrders.ParseDateTextAs(CStr(vDdate), "дата документа (O)", mDdate, msg) Then
            CheckReceiptValues = msg
            Exit Function
        End If
        mHasDdate = True
    End If
    If Trim(CStr(vRdate)) = "" Then
        CheckReceiptValues = "не указана дата поступления (N)"
        Exit Function
    End If
    If Not WmsOrders.ParseDateTextAs(CStr(vRdate), "дата поступления (N)", mRdate, msg) Then
        CheckReceiptValues = msg
        Exit Function
    End If
    mPlace = Trim(CStr(vPlace))
    If mPlace = "" Then
        CheckReceiptValues = "не указано место хранения (U)"
        Exit Function
    End If
    If Not WmsOrders.ParsePriceText(CStr(vPrice), mPrice, mHasPrice, msg) Then
        CheckReceiptValues = msg
        Exit Function
    End If
    CheckReceiptValues = ""
End Function

' the date of the antidubl key: the document date when there is one, else the receipt date
Private Function DupDate() As Double
    If mHasDdate Then DupDate = mDdate Else DupDate = mRdate
End Function

' ================================================================ plan pieces

' a new «Наличие» row: EI, name, article, unit, balance, place, category, state, source, kind of source (Поставщик)
Private Sub PlanStockRow(n As Long, canon As String, qty As Double, sName As String, sArt As String, sUnit As String, sPlace As String, _
    sCat As String, sSrc As String)
    PlanSetValue(SH_STOCK, n, SC_EI, canon, False)
    PlanSetValue(SH_STOCK, n, SC_NAME, sName, False)
    If sArt <> "" Then PlanSetValue(SH_STOCK, n, SC_ART, sArt, False)
    PlanSetValue(SH_STOCK, n, SC_UNIT, sUnit, False)
    PlanSetValue(SH_STOCK, n, SC_QTY, qty, False)
    PlanSetValue(SH_STOCK, n, SC_PLACE, sPlace, False)
    If sCat <> "" Then PlanSetValue(SH_STOCK, n, SC_CAT, sCat, False)
    PlanSetValue(SH_STOCK, n, SC_STATE, EI_ST_ACTIVE, False)
    PlanSetValue(SH_STOCK, n, SC_SRC, sSrc, False)
    PlanSetValue(SH_STOCK, n, SC_STYPE, STYPE_SUPPLIER, False)
End Sub

' a new _RCV row: the receipt of EI n
Private Sub PlanRcvRow(n As Long, canon As String, olid As Long, r As Long, sKey As String, sKind As String, qty As Double, nodoc As Integer)
    PlanSetValue(SH_RCV, n, RV_EI, canon, False)
    PlanSetValue(SH_RCV, n, RV_OL, olid, False)
    PlanSetValue(SH_RCV, n, RV_ROW, r, False)
    PlanSetValue(SH_RCV, n, RV_DUP, WmsOrders.DupHash(sKey), False)
    PlanSetValue(SH_RCV, n, RV_STATE, RV_LIVE, False)
    PlanSetValue(SH_RCV, n, RV_KIND, sKind, False)
    PlanSetValue(SH_RCV, n, RV_QTY, qty, False)
    PlanSetValue(SH_RCV, n, RV_NODOC, nodoc, False)
    PlanSetValue(SH_RCV, n, RV_DUPKEY, sKey, False)
End Sub

' a new _ORD row: the order position olid
Private Sub PlanOrdRow(olid As Long, r As Long, sKey As String, fp As String, ordQty As Double, rcvQty As Double, cnt As Long, nodoc As Long, cancel As String)
    PlanSetValue(SH_ORD, olid, OD_ID, olid, False)
    PlanSetValue(SH_ORD, olid, OD_ROW, r, False)
    If sKey <> "" Then PlanSetValue(SH_ORD, olid, OD_KEY, sKey, False)
    PlanSetValue(SH_ORD, olid, OD_FP, fp, False)
    PlanSetValue(SH_ORD, olid, OD_ORD, ordQty, False)
    PlanSetValue(SH_ORD, olid, OD_RCV, rcvQty, False)
    PlanSetValue(SH_ORD, olid, OD_CNT, cnt, False)
    PlanSetValue(SH_ORD, olid, OD_NODOC, nodoc, False)
    If cancel <> "" Then PlanSetValue(SH_ORD, olid, OD_CANCEL, cancel, False)
End Sub

' the user inputs of an open row that the operation keeps (restored from the journal if lost in a crash)
Private Sub PlanOrderInputs(r As Long)
    Dim cols As Variant, i As Integer, sh As Object
    sh = OSh()
    cols = Array(OC_ORDER, OC_NAME, OC_DOC, OC_INVOICE, OC_ART, OC_UNIT, OC_SUM, OC_SUPPLIER, OC_SELLER, OC_BUYER, OC_CAT, OC_ASSIGNED, _
        OC_PLACE, OC_NOTE, OC_DAYS)
    For i = 0 To UBound(cols)
        If sh.getCellByPosition(cols(i), r).getType() <> com.sun.star.table.CellContentType.EMPTY Then PlanInput(SH_ORDERS, r, cols(i))
    Next i
End Sub

' H P Q of an open row as numbers / dates (the text typed by the user becomes the checked value)
Private Sub PlanOrderValues(r As Long)
    PlanSetValue(SH_ORDERS, r, OC_ORDQTY, mOrdQty, True)
    If mHasOdate Then PlanSetValue(SH_ORDERS, r, OC_ODATE, mOdate, True)
    If mHasEdate Then PlanSetValue(SH_ORDERS, r, OC_EDATE, mEdate, True)
End Sub

' F G J N O of a receipt row as numbers / dates; bRestorable for an open row (the input may be lost in a crash)
Private Sub PlanReceiptValues(r As Long, bRestorable As Boolean, bClearEmpty As Boolean)
    PlanSetValue(SH_ORDERS, r, OC_FACT, mFact, bRestorable)
    If mHasDocQty Then
        PlanSetValue(SH_ORDERS, r, OC_DOCQTY, mDocQty, bRestorable)
    ElseIf bClearEmpty Then
        PlanSetValue(SH_ORDERS, r, OC_DOCQTY, "", False)
    End If
    If mHasPrice Then
        PlanSetValue(SH_ORDERS, r, OC_PRICE, mPrice, bRestorable)
    ElseIf bClearEmpty Then
        PlanSetValue(SH_ORDERS, r, OC_PRICE, "", False)
    End If
    PlanSetValue(SH_ORDERS, r, OC_RDATE, mRdate, bRestorable)
    If mHasDdate Then
        PlanSetValue(SH_ORDERS, r, OC_DDATE, mDdate, bRestorable)
    ElseIf bClearEmpty Then
        PlanSetValue(SH_ORDERS, r, OC_DDATE, "", False)
    End If
End Sub

' the logical fields of a receipt for the journal (audit, oracle)
Private Sub PlanReceiptFields(sOrder As String, sName As String, sArt As String, sUnit As String, sSupplier As String)
    PlanField("ORDER", sOrder)
    PlanField("NAME", sName)
    PlanField("ART", sArt)
    PlanField("UNIT", sUnit)
    PlanField("SUPPLIER", sSupplier)
    PlanField("QTY", mFact)
    If mHasDocQty Then PlanField("DOC_QTY", mDocQty)
    PlanField("DOC", mDoc)
    If mHasDdate Then PlanField("DOC_DATE", IsoDate(mDdate))
    PlanField("DATE", IsoDate(mRdate))
    PlanField("PLACE", mPlace)
    If mHasPrice Then PlanField("PRICE", mPrice)
End Sub

' «Контроль» of a receipt row that the operation keeps as it is (its own receipt part), from the row's values
Private Function RowOwnControl(r As Long, bStorno As Boolean) As String
    Dim sh As Object, g As Object, hasG As Boolean, gq As Double, f As Double
    If bStorno Then
        RowOwnControl = "Приход удалён (сторно)"
        Exit Function
    End If
    sh = OSh()
    f = sh.getCellByPosition(OC_FACT, r).getValue()
    g = sh.getCellByPosition(OC_DOCQTY, r)
    hasG = (g.getType() = com.sun.star.table.CellContentType.VALUE)
    If hasG Then gq = g.getValue()
    RowOwnControl = WmsOrders.ReceiptControl(f, hasG, gq, Trim(Txt(r, OC_DOC)) <> "", _
        sh.getCellByPosition(OC_DDATE, r).getType() = com.sun.star.table.CellContentType.VALUE, Trim(Txt(r, OC_UNIT)))
End Function

' W and Y of the source row of a position after the operation (the source row's own receipt part is kept)
Private Sub PlanSourceStatus(src As Long, od As Variant, rcvQty As Double, nodoc As Double, cancel As String)
    Dim n As Long, bStorno As Boolean, d As Variant, own As String
    If CStr(od(OD_KEY)) <> "" Then
        If WmsOrders.StrictEI(CStr(od(OD_KEY)), n) Then
            If WmsOrders.RcvRegistered(n, d) Then bStorno = (CStr(d(RV_STATE)) = RV_STORNO)
        End If
        own = RowOwnControl(src, bStorno)
    End If
    PlanDerived(SH_ORDERS, src, OC_STATUS, WmsOrders.PositionStatus(od(OD_ORD), rcvQty, nodoc, cancel))
    PlanDerived(SH_ORDERS, src, OC_CTL, WmsOrders.SourceControl(own, od(OD_ORD), rcvQty, cancel, Trim(Txt(src, OC_UNIT))))
End Sub

' ================================================================ «Провести приход» — the first receipt of a position in its row

' Posts row r (0-based) of «Заказы». OK:<seq>, SKIP:<why> (nothing done), ERR:<why> (a problem of the row: nothing done,
' the reason is in Y), ERR-SYS:<why> (a problem of WMS itself: the EI or position counter is not safe, an internal error;
' nothing done), BLOCKED:<why> or an ApplyOperation error (ERR-RB, ERR-CRITICAL). IsSystemResult tells them apart.
Function ReceiptPostRow(r As Long) As String
    Dim why As String, kind As String, n As Long, canon As String, olid As Long, res As String, sKey As String, dup As String
    Dim fp As String, st As String, ctl As String, nodoc As Integer, own As String, sOrder As String, sArt As String, sSupplier As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReceiptPostRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "RECEIVED"
        ReceiptPostRow = "SKIP:уже проведено (" & gKcanon & ")"
        Exit Function
    Case "STORNO"
        ReceiptPostRow = "SKIP:приход " & gKcanon & " удалён (сторно) — новое поступление этой позиции: «Ещё поступление»"
        Exit Function
    Case "CANCELLED"
        ReceiptPostRow = "SKIP:позиция отменена — приход в эту строку не проводится"
        Exit Function
    Case "COPY"
        ReceiptPostRow = "SKIP:строка помечена как КОПИЯ — очистите её кнопкой «Очистить»"
        Exit Function
    Case "FOREIGN"
        WmsOrders.CopyCheckOrder(r)
        ReceiptPostRow = "SKIP:в строке ЕИ, который WMS здесь не проводила — строка не проводится"
        Exit Function
    Case "EMPTY"
        ReceiptPostRow = "ERR:строка пуста"
        Exit Function
    End Select
    why = PostingBlockReason()
    If why <> "" Then
        If Left(why, 8) = "BLOCKED:" Then SetCtl(r, "Не проведено: WMS заблокирована — " & Mid(why, 9) & ". См. лист «" & SH_MAIN & "»")
        ReceiptPostRow = why
        Exit Function
    End If
    why = CheckOrderPart(r, True)
    If why = "" Then why = CheckReceiptCells(r)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        ReceiptPostRow = "ERR:" & why
        Exit Function
    End If
    ' the new EI and the new position number must be safe: a failure here is a problem of WMS, not of this row (D-059)
    n = CLng(SysNum(SK_NEXT_EI))
    why = WmsOrders.NewEIProblem(n)
    olid = WmsOrders.NextOl()
    If why = "" Then why = NewOlProblem(olid)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        ReceiptPostRow = "ERR-SYS:" & why
        Exit Function
    End If
    canon = WmsIssue.EiCanon(n)
    sOrder = Trim(Txt(r, OC_ORDER))
    sArt = Trim(Txt(r, OC_ART))
    sSupplier = Trim(Txt(r, OC_SUPPLIER))
    sKey = WmsOrders.DupKeyText(sSupplier, mDoc, sArt, mName, mFact, DupDate())
    dup = WmsOrders.FindDuplicate(sKey, 0)
    nodoc = DocsMissing(mDoc, mHasDdate)
    st = WmsOrders.PositionStatus(mOrdQty, mFact, nodoc, "")
    own = WmsOrders.ReceiptControl(mFact, mHasDocQty, mDocQty, mDoc <> "", mHasDdate, mUnit)
    ctl = WmsOrders.SourceControl(own, mOrdQty, mFact, "", mUnit)
    fp = WmsOrders.Fingerprint(r, mOrdQty, IIf(mHasOdate, mOdate, 0))
    ' phase 2: one operation — the new EI, its registry row, the receipt, the position and the order row change together
    PlanBegin("RECEIPT", canon)
    PlanField("EI", canon)
    PlanField("OL", olid)
    PlanField("ROW", r + 1)
    PlanField("ORD_QTY", mOrdQty)
    PlanReceiptFields(sOrder, mName, sArt, mUnit, sSupplier)
    PlanField("BAL_AFTER", mFact)
    PlanField("STATUS", st)
    If dup <> "" Then PlanField("DUP", dup)
    PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, n + 1, False)
    PlanStockRow(n, canon, mFact, mName, sArt, mUnit, mPlace, Trim(Txt(r, OC_CAT)), "Заказ " & sOrder)
    PlanRcvRow(n, canon, olid, r, sKey, RV_SRC, mFact, nodoc)
    PlanSetValue(SH_ORD, 0, OD_NEXT_COL, olid + 1, False)
    PlanOrdRow(olid, r, canon, fp, mOrdQty, mFact, 1, nodoc, "")
    PlanOrderInputs(r)
    PlanOrderValues(r)
    PlanReceiptValues(r, True, False)
    PlanSetValue(SH_ORDERS, r, OC_EI, canon, False)
    PlanDerived(SH_ORDERS, r, OC_STATUS, st)
    PlanDerived(SH_ORDERS, r, OC_STOCK, mFact)
    PlanDerived(SH_ORDERS, r, OC_CTL, ctl)
    PlanDerived(SH_ORDERS, r, OC_DUP, IIf(dup <> "", "Возможный дубль: " & dup, ""))
    PlanLockBits(SH_ORDERS, r, 0, OC_LAST, ORDER_LOCKS_POSTED)
    res = ApplyOperation(0, 0)
    If Left(res, 6) = "ERR-RB" Then SetCtl(r, "Не проведено: " & Mid(res, 8))
    ReceiptPostRow = res
    Exit Function
EH:
    ' only the checks can get here (ApplyOperation handles its own errors): nothing was written
    ReceiptPostRow = "ERR-SYS:внутренняя ошибка проверки строки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    On Error Resume Next
    SetCtl(r, "Не проведено: " & Mid(ReceiptPostRow, 9))
End Function

' D-059: True when a result means that WMS itself failed, not the row: posting blocked (fail-closed) or busy, a rollback
' of the operation (a write or the journal failed), an uncertain journal or an unfinished completion, an unsafe counter or
' service table, an internal error — and an operation that was completed only after a failure. Anything unknown counts as
' a system error too.
Function IsSystemResult(res As String) As Boolean
    If Left(res, 3) = "OK:" Then
        IsSystemResult = (InStr(res, "(операция уже в журнале") > 0)
    ElseIf Left(res, 4) = "ERR:" Or Left(res, 5) = "SKIP:" Then
        IsSystemResult = False
    Else
        IsSystemResult = True
    End If
End Function

' "" when OrderLineID olid may be created now (its _ORD row is empty)
Private Function NewOlProblem(olid As Long) As String
    If olid < 1 Or olid >= MAX_SHEET_ROW Then
        NewOlProblem = "номера позиций заказов исчерпаны (NEXT_OL " & olid & ")"
    ElseIf WmsOrders.LastRow(WmsOrders.OrdSheet()) >= olid Then
        NewOlProblem = "служебная таблица позиций " & SH_ORD & " не согласована со счётчиком NEXT_OL " & olid & " — нужна самопроверка"
    Else
        NewOlProblem = ""
    End If
End Function

' ================================================================ «Ещё поступление» — the next receipt of the same position

' r: any row of the position (its source row or one of its receipt rows). The new receipt gets its own row after the
' used rows of the sheet: A D and the identity of the position (B E I L M P R S T, J) are copied, H stays empty.
Function ReceiptAddRow(r As Long, vFact As Variant, vDocQty As Variant, vDoc As Variant, vDdate As Variant, vRdate As Variant, _
    vPlace As Variant, vPrice As Variant) As String
    Dim why As String, kind As String, olid As Long, od As Variant, src As Long, srcMoved As Boolean, sd As Variant, newRow As Long
    Dim n As Long, canon As String, sKey As String, dup As String, nodoc As Integer, rcv2 As Double, nodoc2 As Double, res As String
    Dim cols As Variant, i As Integer, own As String, curN As Long, curMoved As Boolean, bIns As Boolean
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReceiptAddRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "RECEIVED", "STORNO"
    Case "OPEN"
        ReceiptAddRow = "ERR:по этой позиции ещё не было прихода — первый приход проводится в этой строке кнопкой «Провести приход»"
        Exit Function
    Case "CANCELLED"
        ReceiptAddRow = "ERR:позиция отменена — поступление по ней не проводится; оформите новую строку заказа"
        Exit Function
    Case "COPY", "FOREIGN"
        ReceiptAddRow = "ERR:строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
        Exit Function
    Case Else
        ReceiptAddRow = "ERR:строка пуста — поставьте курсор в строку позиции заказа"
        Exit Function
    End Select
    olid = gKol
    curN = gKn
    curMoved = gKmoved
    why = PostingBlockReason()
    If why <> "" Then
        ReceiptAddRow = why
        Exit Function
    End If
    If Not WmsOrders.OlValid(olid, od) Then
        ReceiptAddRow = "ERR:позиция заказа (OLID " & olid & ") не найдена в служебной таблице — нужна самопроверка"
        Exit Function
    End If
    If CStr(od(OD_CANCEL)) = OD_CANCEL_ORDER Then
        ReceiptAddRow = "ERR:позиция отменена — поступление по ней не проводится; оформите новую строку заказа"
        Exit Function
    End If
    src = WmsOrders.SourceRowOf(od)
    srcMoved = gSmoved
    If src < 1 Then
        ReceiptAddRow = "ERR:исходная строка позиции (OLID " & olid & ") не найдена на листе «" & SH_ORDERS & "» — нужна самопроверка"
        Exit Function
    End If
    why = CheckReceiptValues(vFact, vDocQty, vDoc, vDdate, vRdate, vPlace, vPrice)
    ' the new row: inside the block of the order, under the source row and its earlier deliveries (M6 §18)
    newRow = DeliveryRowAfter(src, olid)
    If why = "" And newRow > MAX_SHEET_ROW Then why = "на листе «" & SH_ORDERS & "» нет свободной строки"
    If why <> "" Then
        ReceiptAddRow = "ERR:" & why
        Exit Function
    End If
    ' the new EI must be safe: a failure here is a problem of WMS, not of this delivery (D-059)
    n = CLng(SysNum(SK_NEXT_EI))
    why = WmsOrders.NewEIProblem(n)
    If why <> "" Then
        ReceiptAddRow = "ERR-SYS:" & why
        Exit Function
    End If
    ' the row is inserted (unless it is free) before the plan reads the cells: the plan describes the book after the
    ' insertion, exactly as the recovery rebuilds it. Nothing WMS does here stays in the undo stack.
    bIns = Not OrderRowEmpty(newRow)
    If bIns Then
        UndoBegin()
        why = InsertOrderRow(newRow)
        If why <> "" Then
            UndoEnd()
            ReceiptAddRow = "ERR-SYS:" & why
            Exit Function
        End If
        If r >= newRow Then r = r + 1
    End If
    canon = WmsIssue.EiCanon(n)
    sd = OSh().getCellRangeByPosition(0, src, OC_LAST, src).getDataArray()(0)
    sKey = WmsOrders.DupKeyText(CStr(sd(OC_SUPPLIER)), mDoc, CStr(sd(OC_ART)), CStr(sd(OC_NAME)), mFact, DupDate())
    dup = WmsOrders.FindDuplicate(sKey, 0)
    nodoc = DocsMissing(mDoc, mHasDdate)
    rcv2 = WmsIssue.Round3(od(OD_RCV) + mFact)
    nodoc2 = od(OD_NODOC) + nodoc
    own = WmsOrders.ReceiptControl(mFact, mHasDocQty, mDocQty, mDoc <> "", mHasDdate, Trim(CStr(sd(OC_UNIT))))
    PlanBegin("RECEIPT_ADD", canon)
    PlanField("EI", canon)
    PlanField("OL", olid)
    PlanField("ROW", newRow + 1)
    If bIns Then PlanField("INSROW", newRow + 1)
    PlanField("SRC_ROW", src + 1)
    PlanReceiptFields(Trim(CStr(sd(OC_ORDER))), Trim(CStr(sd(OC_NAME))), Trim(CStr(sd(OC_ART))), Trim(CStr(sd(OC_UNIT))), Trim(CStr(sd(OC_SUPPLIER))))
    PlanField("BAL_AFTER", mFact)
    PlanField("POS_RCV", rcv2)
    If dup <> "" Then PlanField("DUP", dup)
    PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, n + 1, False)
    PlanStockRow(n, canon, mFact, Trim(CStr(sd(OC_NAME))), Trim(CStr(sd(OC_ART))), Trim(CStr(sd(OC_UNIT))), mPlace, Trim(CStr(sd(OC_CAT))), _
        "Заказ " & Trim(CStr(sd(OC_ORDER))))
    PlanRcvRow(n, canon, olid, newRow, sKey, RV_ADD, mFact, nodoc)
    ' the receipt of the row under the cursor was found after rows moved: its row hint is corrected too
    If curMoved And curN <> 0 Then PlanSetValue(SH_RCV, curN, RV_ROW, r, False)
    PlanSetValue(SH_ORD, olid, OD_RCV, rcv2, False)
    PlanSetValue(SH_ORD, olid, OD_CNT, od(OD_CNT) + 1, False)
    PlanSetValue(SH_ORD, olid, OD_NODOC, nodoc2, False)
    If srcMoved Then PlanSetValue(SH_ORD, olid, OD_ROW, src, False)
    ' the new row: the identity of the position is copied by WMS, H stays empty (the ordered quantity is not repeated)
    cols = Array(OC_ORDER, OC_NAME, OC_INVOICE, OC_ART, OC_UNIT, OC_SUPPLIER, OC_SELLER, OC_ODATE, OC_BUYER, OC_CAT, OC_ASSIGNED)
    For i = 0 To UBound(cols)
        If CStr(sd(cols(i))) <> "" Then PlanSetValue(SH_ORDERS, newRow, cols(i), sd(cols(i)), False)
    Next i
    If mDoc <> "" Then PlanSetValue(SH_ORDERS, newRow, OC_DOC, mDoc, False)
    PlanReceiptValues(newRow, False, False)
    PlanSetValue(SH_ORDERS, newRow, OC_PLACE, mPlace, False)
    PlanSetValue(SH_ORDERS, newRow, OC_EI, canon, False)
    PlanSetValue(SH_ORDERS, newRow, OC_BLOCK, "OL" & olid, False)
    PlanDerived(SH_ORDERS, newRow, OC_STATUS, OS_ADD)
    PlanDerived(SH_ORDERS, newRow, OC_STOCK, mFact)
    PlanDerived(SH_ORDERS, newRow, OC_CTL, own)
    If dup <> "" Then PlanDerived(SH_ORDERS, newRow, OC_DUP, "Возможный дубль: " & dup)
    PlanLockBits(SH_ORDERS, newRow, 0, OC_LAST, ORDER_LOCKS_POSTED)
    ' the source row shows the status of the whole position
    PlanSourceStatus(src, od, rcv2, nodoc2, CStr(od(OD_CANCEL)))
    res = ApplyOperation(0, 0)
    If Left(res, 3) = "OK:" Then
        res = res & "|строка " & (newRow + 1) & "|" & canon
    ElseIf bIns And (Left(res, 7) = "ERR-RB:" Or Left(res, 8) = "BLOCKED:" Or Left(res, 5) = "BUSY:") Then
        ' the delivery did not happen (rolled back or refused): the inserted row goes away again. After ERR-CRITICAL the
        ' operation may be in the journal — the row stays for the recovery
        RemoveEmptyOrderRow(newRow)
        UndoEnd()
    End If
    ReceiptAddRow = res
    Exit Function
EH:
    ReceiptAddRow = "ERR-SYS:внутренняя ошибка проверки поступления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    On Error Resume Next
    If bIns Then
        RemoveEmptyOrderRow(newRow)
        UndoEnd()
    End If
End Function

' ================================================================ «Исправить» — RECEIPT_FIX: the same EI, the same physical batch
' The old receipt is reversed and the new one applied in one journal line: quantity, document (C G O), receipt date, place
' and price may change; the EI stays. The new quantity may not fall below what was already issued from this EI.

Function ReceiptFixRow(r As Long, vFact As Variant, vDocQty As Variant, vDoc As Variant, vDdate As Variant, vRdate As Variant, _
    vPlace As Variant, vPrice As Variant) As String
    Dim why As String, kind As String, n As Long, canon As String, olid As Long, isSrc As Boolean, moved As Boolean, rv As Variant, od As Variant
    Dim oldF As Double, s As Double, issued As Double, newBal As Double, nodocOld As Double, nodocNew As Integer, rcv2 As Double, nodoc2 As Double
    Dim src As Long, srcMoved As Boolean, sKey As String, dup As String, res As String, sh As Object, own As String, st As String
    Dim oldG As Object, oldO As Object, oldJ As Object, same As Boolean, oldPlace As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReceiptFixRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = OSh()
    kind = WmsOrders.OrderRowKind(r)
    why = ReceiptRowProblem(kind, "Исправить")
    If why <> "" Then
        ReceiptFixRow = "ERR:" & why
        Exit Function
    End If
    n = gKn
    canon = gKcanon
    olid = gKol
    isSrc = gKsrc
    moved = gKmoved
    why = PostingBlockReason()
    If why <> "" Then
        ReceiptFixRow = why
        Exit Function
    End If
    why = ReceiptStateProblem(r, n, canon, olid, rv, od, s)
    If why <> "" Then
        ReceiptFixRow = "ERR:" & why
        Exit Function
    End If
    oldF = rv(RV_QTY)
    why = CheckReceiptValues(vFact, vDocQty, vDoc, vDdate, vRdate, vPlace, vPrice)
    If why <> "" Then
        ReceiptFixRow = "ERR:" & why
        Exit Function
    End If
    issued = WmsIssue.Round3(oldF - s)
    If mFact < issued - 0.0000001 Then
        ReceiptFixRow = "ERR:по " & canon & " уже выдано " & WmsIssue.QtyText(issued) & " " & Trim(Txt(r, OC_UNIT)) _
            & " — приход нельзя уменьшить ниже выданного (сейчас остаток " & WmsIssue.QtyText(s) & ")"
        Exit Function
    End If
    ' nothing changed?
    oldG = sh.getCellByPosition(OC_DOCQTY, r)
    oldO = sh.getCellByPosition(OC_DDATE, r)
    oldJ = sh.getCellByPosition(OC_PRICE, r)
    oldPlace = Trim(Txt(r, OC_PLACE))
    same = (Abs(mFact - oldF) < 0.0000001 And mDoc = Trim(Txt(r, OC_DOC)) And mPlace = oldPlace _
        And mRdate = sh.getCellByPosition(OC_RDATE, r).getValue())
    If same Then same = (mHasDocQty = (oldG.getType() = com.sun.star.table.CellContentType.VALUE))
    If same And mHasDocQty Then same = (Abs(oldG.getValue() - mDocQty) < 0.0000001)
    If same Then same = (mHasDdate = (oldO.getType() = com.sun.star.table.CellContentType.VALUE))
    If same And mHasDdate Then same = (oldO.getValue() = mDdate)
    If same Then same = (mHasPrice = (oldJ.getType() = com.sun.star.table.CellContentType.VALUE))
    If same And mHasPrice Then same = (Abs(oldJ.getValue() - mPrice) < 0.0000001)
    If same Then
        ReceiptFixRow = "SKIP:ничего не изменилось"
        Exit Function
    End If
    newBal = WmsIssue.Round3(s + mFact - oldF)
    nodocOld = rv(RV_NODOC)
    nodocNew = DocsMissing(mDoc, mHasDdate)
    rcv2 = WmsIssue.Round3(od(OD_RCV) + mFact - oldF)
    nodoc2 = od(OD_NODOC) - nodocOld + nodocNew
    If isSrc Then
        src = r
        srcMoved = (CLng(od(OD_ROW)) <> r)
    Else
        src = WmsOrders.SourceRowOf(od)
        srcMoved = gSmoved
        If src < 1 Then
            ReceiptFixRow = "ERR:исходная строка позиции (OLID " & olid & ") не найдена на листе «" & SH_ORDERS & "» — нужна самопроверка"
            Exit Function
        End If
    End If
    sKey = WmsOrders.DupKeyText(Trim(Txt(r, OC_SUPPLIER)), mDoc, Trim(Txt(r, OC_ART)), Trim(Txt(r, OC_NAME)), mFact, DupDate())
    dup = WmsOrders.FindDuplicate(sKey, n)
    own = WmsOrders.ReceiptControl(mFact, mHasDocQty, mDocQty, mDoc <> "", mHasDdate, Trim(Txt(r, OC_UNIT)))
    PlanBegin("RECEIPT_FIX", canon)
    PlanField("EI", canon)
    PlanField("OL", olid)
    PlanField("ROW", r + 1)
    PlanField("OLD_QTY", oldF)
    PlanField("OLD_DOC", Trim(Txt(r, OC_DOC)))
    PlanField("OLD_PLACE", oldPlace)
    PlanReceiptFields(Trim(Txt(r, OC_ORDER)), Trim(Txt(r, OC_NAME)), Trim(Txt(r, OC_ART)), Trim(Txt(r, OC_UNIT)), Trim(Txt(r, OC_SUPPLIER)))
    PlanField("BAL_BEFORE", s)
    PlanField("BAL_AFTER", newBal)
    PlanField("POS_RCV", rcv2)
    If dup <> "" Then PlanField("DUP", dup)
    PlanSetValue(SH_STOCK, n, SC_QTY, newBal, False)
    If mPlace <> oldPlace Then PlanSetValue(SH_STOCK, n, SC_PLACE, mPlace, False)
    PlanSetValue(SH_RCV, n, RV_QTY, mFact, False)
    PlanSetValue(SH_RCV, n, RV_NODOC, nodocNew, False)
    PlanSetValue(SH_RCV, n, RV_DUP, WmsOrders.DupHash(sKey), False)
    PlanSetValue(SH_RCV, n, RV_DUPKEY, sKey, False)
    If moved Then PlanSetValue(SH_RCV, n, RV_ROW, r, False)
    PlanSetValue(SH_ORD, olid, OD_RCV, rcv2, False)
    PlanSetValue(SH_ORD, olid, OD_NODOC, nodoc2, False)
    If srcMoved Then PlanSetValue(SH_ORD, olid, OD_ROW, src, False)
    PlanSetValue(SH_ORDERS, r, OC_DOC, mDoc, False)
    PlanReceiptValues(r, False, True)
    PlanSetValue(SH_ORDERS, r, OC_PLACE, mPlace, False)
    PlanDerived(SH_ORDERS, r, OC_STOCK, newBal)
    PlanDerived(SH_ORDERS, r, OC_DUP, IIf(dup <> "", "Возможный дубль: " & dup, ""))
    If isSrc Then
        st = WmsOrders.PositionStatus(od(OD_ORD), rcv2, nodoc2, CStr(od(OD_CANCEL)))
        PlanDerived(SH_ORDERS, r, OC_STATUS, st)
        PlanDerived(SH_ORDERS, r, OC_CTL, WmsOrders.SourceControl(own, od(OD_ORD), rcv2, CStr(od(OD_CANCEL)), Trim(Txt(r, OC_UNIT))))
    Else
        PlanDerived(SH_ORDERS, r, OC_CTL, own)
        PlanSourceStatus(src, od, rcv2, nodoc2, CStr(od(OD_CANCEL)))
    End If
    PlanLockBits(SH_ORDERS, r, 0, OC_LAST, ORDER_LOCKS_POSTED)
    res = ApplyOperation(0, 0)
    ReceiptFixRow = res
    Exit Function
EH:
    ReceiptFixRow = "ERR:внутренняя ошибка проверки исправления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' "" when a row of this kind holds a live receipt that «Исправить» / «Удалить» may change
Private Function ReceiptRowProblem(kind As String, sAction As String) As String
    Select Case kind
    Case "RECEIVED"
        ReceiptRowProblem = ""
    Case "STORNO"
        ReceiptRowProblem = "приход " & gKcanon & " уже удалён (сторно)"
    Case "OPEN"
        ReceiptRowProblem = "строка не проведена — «" & sAction & "» нужно только для проведённого прихода; исправьте значения в строке и нажмите «Провести приход»"
    Case "CANCELLED"
        ReceiptRowProblem = "позиция отменена — прихода в этой строке нет"
    Case "COPY", "FOREIGN"
        ReceiptRowProblem = "строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
    Case Else
        ReceiptRowProblem = "строка пуста"
    End Select
End Function

' the receipt of row r (EI n) against its service data: _RCV row, position, registry balance; "" or the problem
Private Function ReceiptStateProblem(r As Long, n As Long, canon As String, olid As Long, rv As Variant, od As Variant, ByRef s As Double) As String
    Dim p As String, f As Object
    If Not WmsOrders.RcvRegistered(n, rv) Then
        ReceiptStateProblem = canon & " не найден среди приходов — нужна самопроверка"
        Exit Function
    End If
    If Not WmsOrders.OlValid(olid, od) Then
        ReceiptStateProblem = "позиция заказа (OLID " & olid & ") не найдена в служебной таблице — нужна самопроверка"
        Exit Function
    End If
    p = WmsIssue.StockRowProblem(n, canon)
    If p <> "" Then
        ReceiptStateProblem = p
        Exit Function
    End If
    If Not WmsIssue.StockQty(n, s) Then
        ReceiptStateProblem = "остаток " & canon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    f = OSh().getCellByPosition(OC_FACT, r)
    If f.getType() <> com.sun.star.table.CellContentType.VALUE Or VarType(rv(RV_QTY)) <> 5 Then
        ReceiptStateProblem = "строка " & (r + 1) & ": количество прихода " & canon & " повреждено — нужна самопроверка"
        Exit Function
    End If
    If Abs(f.getValue() - rv(RV_QTY)) > 0.0000001 Then
        ReceiptStateProblem = "строка " & (r + 1) & ": F = " & WmsIssue.QtyText(f.getValue()) & ", а приход " & canon & " записан на " _
            & WmsIssue.QtyText(rv(RV_QTY)) & " — нужна самопроверка"
        Exit Function
    End If
    ReceiptStateProblem = ""
End Function

' ================================================================ «Удалить» — RECEIPT_DEL: storno of a receipt
' The balance of the EI goes back to 0 and the EI keeps its registry row with the state «Приход удалён (сторно)»; it is
' never reused. Refused while the EI has live issues or live returns (D-069, even when everything issued was returned) —
' and, as a safety net, when the balance is below the receipt (the result would be negative).

Function ReceiptDeleteRow(r As Long) As String
    ReceiptDeleteRow = RcvDelete(r, False)
End Function

' Every check of «Удалить» (D-069 and the rest) without the operation: "OK:" when the storno may be done — the button asks
' for the confirmation only then; otherwise exactly the refusal ReceiptDeleteRow would give. Nothing is written.
Function ReceiptDeleteCheck(r As Long) As String
    ReceiptDeleteCheck = RcvDelete(r, True)
End Function

Private Function RcvDelete(r As Long, checkOnly As Boolean) As String
    Dim why As String, kind As String, n As Long, canon As String, olid As Long, isSrc As Boolean, moved As Boolean, rv As Variant
    Dim od As Variant, s As Double, f As Double, newBal As Double, rcv2 As Double, nodoc2 As Double, src As Long, srcMoved As Boolean
    Dim st As String, ctl As String, cancel As String, nIss As Long, nRet As Long, nAdj As Long
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        RcvDelete = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = WmsOrders.OrderRowKind(r)
    If kind = "STORNO" Then
        RcvDelete = "SKIP:приход " & gKcanon & " уже удалён (сторно)"
        Exit Function
    End If
    why = ReceiptRowProblem(kind, "Удалить")
    If why <> "" Then
        RcvDelete = "ERR:" & why
        Exit Function
    End If
    n = gKn
    canon = gKcanon
    olid = gKol
    isSrc = gKsrc
    moved = gKmoved
    why = PostingBlockReason()
    If why <> "" Then
        RcvDelete = why
        Exit Function
    End If
    why = ReceiptStateProblem(r, n, canon, olid, rv, od, s)
    If why <> "" Then
        RcvDelete = "ERR:" & why
        Exit Function
    End If
    f = rv(RV_QTY)
    ' D-069: while the EI has live dependent movements (issues, returns) its receipt is not cancelled — even when all that
    ' was issued came back: first the returns of the EI, then its issues are cancelled (their rows stay as history, their
    ' numbers are never reused); a live chain «issue → return» never refers to a cancelled receipt
    If Not WmsReturn.LiveDependents(canon, nIss, nRet) Then
        RcvDelete = "ERR-SYS:не удалось проверить связанные выдачи и возвраты " & canon & " (формула движка Calc) — сторно не выполнено"
        Exit Function
    End If
    If nIss > 0 Or nRet > 0 Then
        RcvDelete = "ERR:По этому ЕИ (" & canon & ") есть связанные выдачи/возвраты: действующих выдач " & nIss & ", возвратов " & nRet _
            & IIf(s < f - 0.0000001, " (уже выдано " & WmsIssue.QtyText(WmsIssue.Round3(f - s)) & " " & Trim(Txt(r, OC_UNIT)) & ")", "") _
            & ". Сначала выполните сторно зависимых операций: " & IIf(nRet > 0, "возвратов этого ЕИ (лист «" & SH_RETURNS & "»), затем ", "") _
            & "выдач (лист «" & SH_ISSUES & "»)"
        Exit Function
    End If
    ' D-069 for the corrections (Final Core): a live move, write-off or inventory correction refers to the receipt too
    If Not WmsAdjust.LiveAdjustments(canon, nAdj) Then
        RcvDelete = "ERR-SYS:не удалось проверить корректировки " & canon & " (формула движка Calc) — сторно не выполнено"
        Exit Function
    End If
    If nAdj > 0 Then
        RcvDelete = "ERR:По этому ЕИ (" & canon & ") есть действующие корректировки: " & nAdj & " (лист «" & SH_ADJUST _
            & "»: перемещения, списания, инвентаризация). Сначала выполните их сторно"
        Exit Function
    End If
    If s < f - 0.0000001 Then
        RcvDelete = "ERR:по " & canon & " уже выдано " & WmsIssue.QtyText(WmsIssue.Round3(f - s)) & " " & Trim(Txt(r, OC_UNIT)) _
            & " — удаление прихода дало бы отрицательный остаток; сначала удалите (сторно) выдачи по этому ЕИ"
        Exit Function
    End If
    newBal = WmsIssue.Round3(s - f)
    rcv2 = WmsIssue.Round3(od(OD_RCV) - f)
    nodoc2 = od(OD_NODOC) - rv(RV_NODOC)
    cancel = CStr(od(OD_CANCEL))
    If isSrc Then
        src = r
        srcMoved = (CLng(od(OD_ROW)) <> r)
    Else
        src = WmsOrders.SourceRowOf(od)
        srcMoved = gSmoved
        If src < 1 Then
            RcvDelete = "ERR:исходная строка позиции (OLID " & olid & ") не найдена на листе «" & SH_ORDERS & "» — нужна самопроверка"
            Exit Function
        End If
    End If
    If checkOnly Then
        RcvDelete = "OK:сторно прихода " & canon & " допустимо"
        Exit Function
    End If
    PlanBegin("RECEIPT_DEL", canon)
    PlanField("EI", canon)
    PlanField("OL", olid)
    PlanField("ROW", r + 1)
    PlanField("QTY", f)
    PlanField("BAL_BEFORE", s)
    PlanField("BAL_AFTER", newBal)
    PlanField("POS_RCV", rcv2)
    PlanSetValue(SH_STOCK, n, SC_QTY, newBal, False)
    PlanSetValue(SH_STOCK, n, SC_STATE, EI_ST_STORNO, False)
    PlanSetValue(SH_RCV, n, RV_STATE, RV_STORNO, False)
    PlanSetValue(SH_RCV, n, RV_DUP, "", False)
    PlanSetValue(SH_RCV, n, RV_DUPKEY, "", False)
    If moved Then PlanSetValue(SH_RCV, n, RV_ROW, r, False)
    PlanSetValue(SH_ORD, olid, OD_RCV, rcv2, False)
    PlanSetValue(SH_ORD, olid, OD_CNT, od(OD_CNT) - 1, False)
    PlanSetValue(SH_ORD, olid, OD_NODOC, nodoc2, False)
    If srcMoved Then PlanSetValue(SH_ORD, olid, OD_ROW, src, False)
    PlanDerived(SH_ORDERS, r, OC_STOCK, newBal)
    PlanDerived(SH_ORDERS, r, OC_DUP, "")
    If isSrc Then
        st = WmsOrders.PositionStatus(od(OD_ORD), rcv2, nodoc2, cancel)
        ctl = WmsOrders.SourceControl("Приход удалён (сторно)", od(OD_ORD), rcv2, cancel, Trim(Txt(r, OC_UNIT)))
        PlanDerived(SH_ORDERS, r, OC_STATUS, st)
        PlanDerived(SH_ORDERS, r, OC_CTL, ctl)
    Else
        PlanDerived(SH_ORDERS, r, OC_STATUS, OS_ADD_STORNO)
        PlanDerived(SH_ORDERS, r, OC_CTL, "Приход удалён (сторно)")
        PlanSourceStatus(src, od, rcv2, nodoc2, cancel)
    End If
    PlanLockBits(SH_ORDERS, r, 0, OC_LAST, ORDER_LOCKS_POSTED)
    RcvDelete = ApplyOperation(0, 0)
    Exit Function
EH:
    RcvDelete = "ERR:внутренняя ошибка проверки удаления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Отменить заказ» — ORDER_CANCEL: no EI, NEXT_EI untouched

Function OrderCancelRow(r As Long) As String
    Dim why As String, kind As String, olid As Long, od As Variant, fp As String, src As Long, res As String, sh As Object
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        OrderCancelRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = OSh()
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "CANCELLED"
        OrderCancelRow = "SKIP:позиция уже отменена"
        Exit Function
    Case "COPY", "FOREIGN"
        OrderCancelRow = "ERR:строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
        Exit Function
    Case "EMPTY"
        OrderCancelRow = "ERR:строка пуста"
        Exit Function
    Case "RECEIVED", "STORNO"
        olid = gKol
        If Not WmsOrders.OlValid(olid, od) Then
            OrderCancelRow = "ERR:позиция заказа (OLID " & olid & ") не найдена в служебной таблице — нужна самопроверка"
            Exit Function
        End If
        If CStr(od(OD_CANCEL)) = OD_CANCEL_ORDER Then
            OrderCancelRow = "SKIP:позиция уже отменена"
            Exit Function
        End If
        If od(OD_RCV) > 0.0000001 Then
            OrderCancelRow = "ERR:по позиции уже получено " & WmsIssue.QtyText(od(OD_RCV)) & " — заказ целиком не отменяется; чтобы не ждать остаток, " _
                & "используйте «Отменить остаток»"
            Exit Function
        End If
    End Select
    why = PostingBlockReason()
    If why <> "" Then
        If kind = "OPEN" And Left(why, 8) = "BLOCKED:" Then SetCtl(r, "Не отменено: WMS заблокирована — " & Mid(why, 9))
        OrderCancelRow = why
        Exit Function
    End If
    If kind = "OPEN" Then
        why = CheckOrderPart(r, False)
        If why = "" And sh.getCellByPosition(OC_FACT, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
            why = "в строке указано фактическое количество (F) — если товар пришёл, проведите приход; иначе очистите F и отмените заказ"
        End If
        olid = WmsOrders.NextOl()
        If why = "" Then why = NewOlProblem(olid)
        If why <> "" Then
            SetCtl(r, "Не отменено: " & why)
            OrderCancelRow = "ERR:" & why
            Exit Function
        End If
        fp = WmsOrders.Fingerprint(r, mOrdQty, IIf(mHasOdate, mOdate, 0))
        PlanBegin("ORDER_CANCEL", "OL" & olid)
        PlanField("OL", olid)
        PlanField("ROW", r + 1)
        PlanField("ORDER", Trim(Txt(r, OC_ORDER)))
        PlanField("NAME", Trim(Txt(r, OC_NAME)))
        PlanField("ORD_QTY", mOrdQty)
        PlanSetValue(SH_ORD, 0, OD_NEXT_COL, olid + 1, False)
        PlanOrdRow(olid, r, "", fp, mOrdQty, 0, 0, 0, OD_CANCEL_ORDER)
        PlanOrderInputs(r)
        PlanOrderValues(r)
        PlanDerived(SH_ORDERS, r, OC_STATUS, OS_CANCELLED)
        PlanDerived(SH_ORDERS, r, OC_CTL, "Заказ отменён; ЕИ не создавался")
        PlanDerived(SH_ORDERS, r, OC_STOCK, "")
        PlanDerived(SH_ORDERS, r, OC_DUP, "")
        PlanLockBits(SH_ORDERS, r, 0, OC_LAST, ORDER_LOCKS_POSTED)
    Else
        ' every receipt of the position was cancelled (storno): the position itself is cancelled now
        src = WmsOrders.SourceRowOf(od)
        If src < 1 Then
            OrderCancelRow = "ERR:исходная строка позиции (OLID " & olid & ") не найдена на листе «" & SH_ORDERS & "» — нужна самопроверка"
            Exit Function
        End If
        PlanBegin("ORDER_CANCEL", "OL" & olid)
        PlanField("OL", olid)
        PlanField("ROW", src + 1)
        PlanField("ORD_QTY", od(OD_ORD))
        PlanSetValue(SH_ORD, olid, OD_CANCEL, OD_CANCEL_ORDER, False)
        If gSmoved Then PlanSetValue(SH_ORD, olid, OD_ROW, src, False)
        PlanSourceStatus(src, od, 0, 0, OD_CANCEL_ORDER)
    End If
    res = ApplyOperation(0, 0)
    If kind = "OPEN" And Left(res, 6) = "ERR-RB" Then SetCtl(r, "Не отменено: " & Mid(res, 8))
    OrderCancelRow = res
    Exit Function
EH:
    OrderCancelRow = "ERR:внутренняя ошибка проверки отмены: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Отменить остаток» — ORDER_CANCEL_REST
' Received quantity and H stay; the rest is no longer expected. History is kept.

Function OrderCancelRestRow(r As Long) As String
    Dim why As String, kind As String, olid As Long, od As Variant, src As Long
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        OrderCancelRestRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "RECEIVED", "STORNO"
    Case "OPEN"
        OrderCancelRestRow = "ERR:по позиции ещё нет прихода — чтобы отменить заказ целиком, используйте «Отменить заказ»"
        Exit Function
    Case "CANCELLED"
        OrderCancelRestRow = "SKIP:позиция отменена"
        Exit Function
    Case "COPY", "FOREIGN"
        OrderCancelRestRow = "ERR:строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
        Exit Function
    Case Else
        OrderCancelRestRow = "ERR:строка пуста"
        Exit Function
    End Select
    olid = gKol
    If Not WmsOrders.OlValid(olid, od) Then
        OrderCancelRestRow = "ERR:позиция заказа (OLID " & olid & ") не найдена в служебной таблице — нужна самопроверка"
        Exit Function
    End If
    Select Case CStr(od(OD_CANCEL))
    Case OD_CANCEL_ORDER
        OrderCancelRestRow = "SKIP:позиция отменена"
        Exit Function
    Case OD_CANCEL_REST
        OrderCancelRestRow = "SKIP:остаток уже отменён"
        Exit Function
    End Select
    If od(OD_RCV) <= 0.0000001 Then
        OrderCancelRestRow = "ERR:по позиции нет действующих приходов — чтобы отменить заказ целиком, используйте «Отменить заказ»"
        Exit Function
    End If
    If od(OD_RCV) >= od(OD_ORD) - 0.0000001 Then
        OrderCancelRestRow = "ERR:позиция получена полностью — отменять нечего"
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        OrderCancelRestRow = why
        Exit Function
    End If
    src = WmsOrders.SourceRowOf(od)
    If src < 1 Then
        OrderCancelRestRow = "ERR:исходная строка позиции (OLID " & olid & ") не найдена на листе «" & SH_ORDERS & "» — нужна самопроверка"
        Exit Function
    End If
    PlanBegin("ORDER_CANCEL_REST", "OL" & olid)
    PlanField("OL", olid)
    PlanField("ROW", src + 1)
    PlanField("ORD_QTY", od(OD_ORD))
    PlanField("RCV_QTY", od(OD_RCV))
    PlanField("REST", WmsIssue.Round3(od(OD_ORD) - od(OD_RCV)))
    PlanSetValue(SH_ORD, olid, OD_CANCEL, OD_CANCEL_REST, False)
    If gSmoved Then PlanSetValue(SH_ORD, olid, OD_ROW, src, False)
    PlanSourceStatus(src, od, od(OD_RCV), od(OD_NODOC), OD_CANCEL_REST)
    OrderCancelRestRow = ApplyOperation(0, 0)
    Exit Function
EH:
    OrderCancelRestRow = "ERR:внутренняя ошибка проверки отмены остатка: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Проверить» — no accounting change
' An open row: every check of «Провести приход» (or of the order part only, when F is empty), the EI it would get and the
' possible duplicate; the result goes into Y (and AB). A row with a receipt: its receipt against the registry and the
' service data; the result is returned (the row is not changed).

Function CheckRow(r As Long) As String
    Dim kind As String, why As String, n As Long, sKey As String, dup As String, rv As Variant, od As Variant, s As Double, x As Object
    Dim src As Long, sh As Object, st As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        CheckRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = OSh()
    kind = WmsOrders.OrderRowKind(r)
    Select Case kind
    Case "EMPTY"
        CheckRow = "ERR:строка пуста"
    Case "COPY", "FOREIGN"
        If kind = "FOREIGN" Then WmsOrders.CopyCheckOrder(r)
        CheckRow = "ERR:строка — КОПИЯ, не проводится; очистите её кнопкой «Очистить»"
    Case "CANCELLED"
        CheckRow = "OK:позиция отменена"
    Case "OPEN"
        why = CheckOrderPart(r, True)
        If why = "" And sh.getCellByPosition(OC_FACT, r).getType() = com.sun.star.table.CellContentType.EMPTY Then
            WmsOrders.SetIfDiff(sh.getCellByPosition(OC_STATUS, r), WmsOrders.OpenRowStatus(r))
            SetCtl(r, "Заказ в порядке; для прихода укажите F, N, U и нажмите «Провести приход»")
            CheckRow = "OK:заказ в порядке, прихода ещё нет"
            Exit Function
        End If
        If why = "" Then why = CheckReceiptCells(r)
        n = CLng(SysNum(SK_NEXT_EI))
        If why = "" Then why = WmsOrders.NewEIProblem(n)
        If why <> "" Then
            SetCtl(r, "Ошибка: " & why)
            CheckRow = "ERR:" & why
            Exit Function
        End If
        sKey = WmsOrders.DupKeyText(Trim(Txt(r, OC_SUPPLIER)), mDoc, Trim(Txt(r, OC_ART)), mName, mFact, DupDate())
        dup = WmsOrders.FindDuplicate(sKey, 0)
        WmsOrders.SetIfDiff(sh.getCellByPosition(OC_DUP, r), IIf(dup <> "", "Возможный дубль: " & dup, ""))
        SetCtl(r, "Готово к проведению: будет создан " & WmsIssue.EiCanon(n) & "; " _
            & WmsOrders.ReceiptControl(mFact, mHasDocQty, mDocQty, mDoc <> "", mHasDdate, mUnit))
        CheckRow = "OK:готово к проведению (" & WmsIssue.EiCanon(n) & ")" & IIf(dup <> "", "; возможный дубль: " & dup, "")
    Case Else
        why = ReceiptStateProblem(r, gKn, gKcanon, gKol, rv, od, s)
        If why = "" Then
            x = sh.getCellByPosition(OC_STOCK, r)
            If x.getType() <> com.sun.star.table.CellContentType.VALUE Then
                why = "«Наличие» (X) строки не число"
            ElseIf Abs(x.getValue() - s) > 0.0000001 Then
                why = "«Наличие» (X) строки " & WmsIssue.QtyText(x.getValue()) & ", а остаток " & gKcanon & " в реестре " & WmsIssue.QtyText(s)
            End If
        End If
        If why = "" Then
            src = WmsOrders.SourceRowOf(od)
            If src < 1 Then why = "исходная строка позиции не найдена"
        End If
        If why <> "" Then
            CheckRow = "ERR:" & why
            Exit Function
        End If
        st = WmsOrders.PositionStatus(od(OD_ORD), od(OD_RCV), od(OD_NODOC), CStr(od(OD_CANCEL)))
        CheckRow = "OK:" & gKcanon & IIf(kind = "STORNO", " (приход удалён, сторно)", "") & ": остаток " & WmsIssue.QtyText(s) & " " _
            & Trim(Txt(r, OC_UNIT)) & "; статус позиции «" & st & "», " & WmsOrders.PositionSummary(od(OD_ORD), od(OD_RCV), CStr(od(OD_CANCEL)), _
            Trim(Txt(r, OC_UNIT))) & "; исходная строка " & (src + 1)
    End Select
    Exit Function
EH:
    CheckRow = "ERR:внутренняя ошибка проверки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ several rows at once («Проверить», «Провести приход»)

' «Провести приход» for a block of rows (D-059). Every visible row with a fact F is posted by its own atomic operation;
' a row without F is skipped (the order still waits for the goods), so is a row already posted, cancelled or a copy;
' hidden rows (a filter) are not touched. A problem of one row (the reason in its Y) does not stop the others. A system
' error of WMS (IsSystemResult) stops the whole block at once: the rows after it are not processed.
' Result: "OK=<posted>;ERR=<rejected>;SKIP=<skipped>;STOP=<1-based row of the stop, 0 = none>", a line with the first
' rejection ("строка N: reason") and a line with the system error ("строка N: result").
Function ReceiptPostRange(r0 As Long, r1 As Long) As String
    Dim r As Long, res As String, nOk As Long, nErr As Long, nSkip As Long, sh As Object, firstErr As String, sysErr As String
    Dim stopRow As Long
    On Error GoTo EH
    sh = OSh()
    For r = r0 To r1
        If sh.getRows().getByIndex(r).IsVisible Then
            If sh.getCellByPosition(OC_FACT, r).getType() = com.sun.star.table.CellContentType.EMPTY Then
                nSkip = nSkip + 1
            ElseIf WmsOrders.OrderRowKind(r) = "OPEN" Then
                res = ReceiptPostRow(r)
                If IsSystemResult(res) Then
                    ' the operation of this row may even have been completed (after a failure): it counts as posted
                    If Left(res, 3) = "OK:" Then nOk = nOk + 1
                    sysErr = "строка " & (r + 1) & ": " & res
                    stopRow = r + 1
                    Exit For
                ElseIf Left(res, 3) = "OK:" Then
                    nOk = nOk + 1
                ElseIf Left(res, 5) = "SKIP:" Then
                    nSkip = nSkip + 1
                Else
                    nErr = nErr + 1
                    If firstErr = "" Then firstErr = "строка " & (r + 1) & ": " & Mid(res, InStr(res, ":") + 1)
                End If
            Else
                nSkip = nSkip + 1
            End If
        End If
    Next r
DONE:
    ReceiptPostRange = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";STOP=" & stopRow & Chr(10) & firstErr & Chr(10) & sysErr
    Exit Function
EH:
    sysErr = "строка " & (r + 1) & ": ERR-SYS:внутренняя ошибка проведения блока: " & Error$ & " (код " & Err & ")"
    stopRow = r + 1
    Resume DONE
End Function

' ================================================================ перенос старой таблицы (M7 PRIME: инструмент WMS_LEGACY_TRANSFER)
' Приход позиции старого листа «Заказы», уже бывший до WMS: строка вписана инструментом переноса как обычная строка заказа
' (ввод A..U, Z, AA), операция проводит её как приход WMS — позиция _ORD, запись прихода _RCV, строка «Наличие», V W X Y —
' но с прежним номером ЕИ (без перенумерации) и с текущим остатком этого ЕИ (после прихода по старой таблице могли быть
' выдачи: остаток ≤ количества прихода; разница — расход до переноса, как уже выданное). Журнал хранит происхождение:
' строку staging, её SHA-256, старые статус, контроль и наличие, откуда взят остаток. Дальше позиция живёт обычной
' жизнью WMS: «Ещё поступление», выдачи, возвраты, корректировки, «Исправить» (уже расходованное считается выданным).
' Вызывается только инструментом переноса на копии чистой книги-кандидата.

' "" when EI n (any number, not only NEXT_EI) may be given to a transferred receipt: its «Наличие» and _RCV rows are empty
' and «Заказы» does not hold it; otherwise the reason
Private Function LegacyEIProblem(n As Long, canon As String) As String
    Dim d As Variant
    If n < 1 Or n >= MAX_SHEET_ROW Then
        LegacyEIProblem = "номер " & canon & " вне допустимого диапазона"
        Exit Function
    End If
    d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
    If CStr(d(SC_EI)) <> "" Or CStr(d(SC_QTY)) <> "" Then
        LegacyEIProblem = canon & " уже есть в «" & SH_STOCK & "» — перенос не перенумеровывает и не сливает ЕИ"
        Exit Function
    End If
    d = WmsOrders.RcvRow(n)
    If CStr(d(RV_EI)) <> "" Or CStr(d(RV_OL)) <> "" Then
        LegacyEIProblem = canon & " уже есть среди приходов (" & SH_RCV & ")"
        Exit Function
    End If
    If WmsOrders.CountEI(canon) > 0 Then
        LegacyEIProblem = canon & " уже записан в «" & SH_ORDERS & "»"
        Exit Function
    End If
    LegacyEIProblem = ""
End Function

' the current balance of a transferred EI: 0 ≤ balance ≤ its receipt; "" or the reason
Private Function LegacyBalance(sBal As String, fact As Double, ByRef bal As Double) As String
    Dim st As Integer, msg As String
    st = ParseQtyText(sBal, bal, msg)
    If st = 5 And bal = 0 And InStr(sBal, "-") = 0 Then
        bal = 0
    ElseIf st <> 0 Then
        LegacyBalance = "остаток «" & sBal & "» — " & msg
        Exit Function
    End If
    If bal > fact + 0.0000001 Then
        LegacyBalance = "остаток " & WmsIssue.QtyText(bal) & " больше количества прихода " & WmsIssue.QtyText(fact) & " — такой остаток перенос не создаёт"
        Exit Function
    End If
    LegacyBalance = ""
End Function

' the audit fields of a transferred receipt (the journal: the history of the transfer, the reconciliation)
Private Sub PlanLegacyFields(sOrigin As String, sLegStatus As String, sLegCtl As String, sLegStock As String, sLegDup As String, _
    sStockSrc As String, sEiSrc As String)
    PlanField("ORIGIN", sOrigin)
    PlanField("LEGACY_STATUS", sLegStatus)
    PlanField("LEGACY_CTL", sLegCtl)
    PlanField("LEGACY_STOCK", sLegStock)
    PlanField("LEGACY_DUP", sLegDup)
    PlanField("STOCK_SRC", sStockSrc)
    PlanField("EI_SRC", sEiSrc)
End Sub

' LEGACY_RECEIPT: the first receipt of a transferred position in its source row r (0-based, kind OPEN: written by the tool,
' V empty). sEI — the EI of the old table (or the new number the tool planned above all old ones: sEiSrc NEW); sBal — the
' current balance; sPlaceNow — the current place of the EI when it differs from U (""); the rest — the audit fields (the old
' W, Y, X and AB of the row — LEGACY_STATUS, LEGACY_CTL, LEGACY_STOCK, LEGACY_DUP — and the sources of the balance and the EI).
' OK:<seq>, SKIP:<why> (already done — a repeated run), ERR:<why> (nothing written), or an ApplyOperation error.
Function ReceiptLegacyRow(r As Long, sEI As String, sBal As String, sPlaceNow As String, sOrigin As String, sLegStatus As String, _
    sLegCtl As String, sLegStock As String, sLegDup As String, sStockSrc As String, sEiSrc As String) As String
    Dim why As String, kind As String, n As Long, canon As String, msg As String, olid As Long, res As String, sKey As String, dup As String
    Dim fp As String, st As String, ctl As String, nodoc As Integer, own As String, sOrder As String, sArt As String, sSupplier As String
    Dim bal As Double, placeNow As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ReceiptLegacyRow = "ERR:строка вне листа"
        Exit Function
    End If
    On Error GoTo EH
    kind = WmsOrders.OrderRowKind(r)
    If kind = "RECEIVED" Or kind = "STORNO" Then
        ReceiptLegacyRow = "SKIP:строка уже проведена (" & gKcanon & ")"
        Exit Function
    ElseIf kind <> "OPEN" Then
        ReceiptLegacyRow = "ERR:строка " & (r + 1) & " не строка заказа без прихода (" & kind & ")"
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        ReceiptLegacyRow = why
        Exit Function
    End If
    why = CheckOrderPart(r, True)
    If why = "" Then why = CheckReceiptCells(r)
    If why <> "" Then
        ReceiptLegacyRow = "ERR:" & why
        Exit Function
    End If
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        ReceiptLegacyRow = "ERR:" & msg
        Exit Function
    End If
    why = LegacyEIProblem(n, canon)
    If why = "" Then why = LegacyBalance(sBal, mFact, bal)
    olid = WmsOrders.NextOl()
    If why = "" Then why = NewOlProblem(olid)
    If why <> "" Then
        ReceiptLegacyRow = "ERR:" & why
        Exit Function
    End If
    placeNow = Trim(sPlaceNow)
    If placeNow = "" Then placeNow = mPlace
    sOrder = Trim(Txt(r, OC_ORDER))
    sArt = Trim(Txt(r, OC_ART))
    sSupplier = Trim(Txt(r, OC_SUPPLIER))
    sKey = WmsOrders.DupKeyText(sSupplier, mDoc, sArt, mName, mFact, DupDate())
    dup = WmsOrders.FindDuplicate(sKey, 0)
    nodoc = DocsMissing(mDoc, mHasDdate)
    st = WmsOrders.PositionStatus(mOrdQty, mFact, nodoc, "")
    own = WmsOrders.ReceiptControl(mFact, mHasDocQty, mDocQty, mDoc <> "", mHasDdate, mUnit)
    ctl = WmsOrders.SourceControl(own, mOrdQty, mFact, "", mUnit)
    fp = WmsOrders.Fingerprint(r, mOrdQty, IIf(mHasOdate, mOdate, 0))
    PlanBegin("LEGACY_RECEIPT", canon)
    PlanField("EI", canon)
    PlanField("OL", olid)
    PlanField("ROW", r + 1)
    PlanField("ORD_QTY", mOrdQty)
    PlanReceiptFields(sOrder, mName, sArt, mUnit, sSupplier)
    PlanField("BAL_AFTER", bal)
    If placeNow <> mPlace Then PlanField("PLACE_NOW", placeNow)
    PlanField("STATUS", st)
    If dup <> "" Then PlanField("DUP", dup)
    PlanLegacyFields(sOrigin, sLegStatus, sLegCtl, sLegStock, sLegDup, sStockSrc, sEiSrc)
    If n >= SysNum(SK_NEXT_EI) Then PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, n + 1, False)
    PlanStockRow(n, canon, bal, mName, sArt, mUnit, placeNow, Trim(Txt(r, OC_CAT)), "Перенос: заказ " & sOrder)
    PlanRcvRow(n, canon, olid, r, sKey, RV_SRC, mFact, nodoc)
    PlanSetValue(SH_ORD, 0, OD_NEXT_COL, olid + 1, False)
    PlanOrdRow(olid, r, canon, fp, mOrdQty, mFact, 1, nodoc, "")
    PlanOrderInputs(r)
    PlanOrderValues(r)
    PlanReceiptValues(r, True, False)
    PlanSetValue(SH_ORDERS, r, OC_EI, canon, False)
    PlanDerived(SH_ORDERS, r, OC_STATUS, st)
    PlanDerived(SH_ORDERS, r, OC_STOCK, bal)
    PlanDerived(SH_ORDERS, r, OC_CTL, ctl)
    PlanDerived(SH_ORDERS, r, OC_DUP, IIf(dup <> "", "Возможный дубль: " & dup, ""))
    PlanLockBits(SH_ORDERS, r, 0, OC_LAST, ORDER_LOCKS_POSTED)
    res = ApplyOperation(0, 0)
    ReceiptLegacyRow = res
    Exit Function
EH:
    ReceiptLegacyRow = "ERR-SYS:внутренняя ошибка переноса строки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' LEGACY_RECEIPT_ADD: a further receipt of a transferred position (a delivery row of the old table). Row r (0-based, OPEN,
' written by the tool right under the source row srcRow and the earlier deliveries of its position, H empty) becomes the
' delivery row of the position, as «Ещё поступление» makes it — only without inserting a row, with the old EI and the
' current balance. The identity of the position (A B D E I L M P R S T) is taken from the source row.
Function ReceiptLegacyAdd(r As Long, srcRow As Long, sEI As String, sBal As String, sPlaceNow As String, sOrigin As String, _
    sLegStatus As String, sLegCtl As String, sLegStock As String, sLegDup As String, sStockSrc As String, sEiSrc As String) As String
    Dim why As String, kind As String, olid As Long, od As Variant, sd As Variant, n As Long, canon As String, msg As String
    Dim sKey As String, dup As String, nodoc As Integer, rcv2 As Double, nodoc2 As Double, res As String, cols As Variant, i As Integer
    Dim own As String, bal As Double, placeNow As String, want As Long
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Or srcRow < 1 Or srcRow >= r Then
        ReceiptLegacyAdd = "ERR:строка поступления должна быть ниже исходной строки позиции"
        Exit Function
    End If
    On Error GoTo EH
    kind = WmsOrders.OrderRowKind(r)
    If kind = "RECEIVED" Or kind = "STORNO" Then
        ReceiptLegacyAdd = "SKIP:строка уже проведена (" & gKcanon & ")"
        Exit Function
    ElseIf kind <> "OPEN" Then
        ReceiptLegacyAdd = "ERR:строка " & (r + 1) & " не строка поступления без прихода (" & kind & ")"
        Exit Function
    End If
    If WmsOrders.OrderRowKind(srcRow) <> "RECEIVED" Or Not gKsrc Then
        ReceiptLegacyAdd = "ERR:исходная строка " & (srcRow + 1) & " — не проведённая исходная строка позиции"
        Exit Function
    End If
    olid = gKol
    If Not WmsOrders.OlValid(olid, od) Then
        ReceiptLegacyAdd = "ERR:позиция заказа (OLID " & olid & ") не найдена в служебной таблице"
        Exit Function
    End If
    If CStr(od(OD_CANCEL)) = OD_CANCEL_ORDER Then
        ReceiptLegacyAdd = "ERR:позиция отменена — поступление по ней не проводится"
        Exit Function
    End If
    want = DeliveryRowAfter(srcRow, olid)
    If want <> r Then
        ReceiptLegacyAdd = "ERR:строка поступления " & (r + 1) & " не сразу под позицией (ожидалась строка " & (want + 1) & ")"
        Exit Function
    End If
    If Trim(Txt(r, OC_ORDQTY)) <> "" Then
        ReceiptLegacyAdd = "ERR:у строки поступления заполнено заказанное количество (H)"
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        ReceiptLegacyAdd = why
        Exit Function
    End If
    why = CheckReceiptCells(r)
    If why <> "" Then
        ReceiptLegacyAdd = "ERR:" & why
        Exit Function
    End If
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        ReceiptLegacyAdd = "ERR:" & msg
        Exit Function
    End If
    why = LegacyEIProblem(n, canon)
    If why = "" Then why = LegacyBalance(sBal, mFact, bal)
    If why <> "" Then
        ReceiptLegacyAdd = "ERR:" & why
        Exit Function
    End If
    placeNow = Trim(sPlaceNow)
    If placeNow = "" Then placeNow = mPlace
    sd = OSh().getCellRangeByPosition(0, srcRow, OC_LAST, srcRow).getDataArray()(0)
    sKey = WmsOrders.DupKeyText(CStr(sd(OC_SUPPLIER)), mDoc, CStr(sd(OC_ART)), CStr(sd(OC_NAME)), mFact, DupDate())
    dup = WmsOrders.FindDuplicate(sKey, 0)
    nodoc = DocsMissing(mDoc, mHasDdate)
    rcv2 = WmsIssue.Round3(od(OD_RCV) + mFact)
    nodoc2 = od(OD_NODOC) + nodoc
    own = WmsOrders.ReceiptControl(mFact, mHasDocQty, mDocQty, mDoc <> "", mHasDdate, Trim(CStr(sd(OC_UNIT))))
    PlanBegin("LEGACY_RECEIPT_ADD", canon)
    PlanField("EI", canon)
    PlanField("OL", olid)
    PlanField("ROW", r + 1)
    PlanField("SRC_ROW", srcRow + 1)
    PlanReceiptFields(Trim(CStr(sd(OC_ORDER))), Trim(CStr(sd(OC_NAME))), Trim(CStr(sd(OC_ART))), Trim(CStr(sd(OC_UNIT))), Trim(CStr(sd(OC_SUPPLIER))))
    PlanField("BAL_AFTER", bal)
    If placeNow <> mPlace Then PlanField("PLACE_NOW", placeNow)
    PlanField("POS_RCV", rcv2)
    If dup <> "" Then PlanField("DUP", dup)
    PlanLegacyFields(sOrigin, sLegStatus, sLegCtl, sLegStock, sLegDup, sStockSrc, sEiSrc)
    If n >= SysNum(SK_NEXT_EI) Then PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, n + 1, False)
    PlanStockRow(n, canon, bal, Trim(CStr(sd(OC_NAME))), Trim(CStr(sd(OC_ART))), Trim(CStr(sd(OC_UNIT))), placeNow, Trim(CStr(sd(OC_CAT))), _
        "Перенос: заказ " & Trim(CStr(sd(OC_ORDER))))
    PlanRcvRow(n, canon, olid, r, sKey, RV_ADD, mFact, nodoc)
    PlanSetValue(SH_ORD, olid, OD_RCV, rcv2, False)
    PlanSetValue(SH_ORD, olid, OD_CNT, od(OD_CNT) + 1, False)
    PlanSetValue(SH_ORD, olid, OD_NODOC, nodoc2, False)
    ' the identity of the position, as «Ещё поступление» copies it; the old row's own values of these columns are replaced
    cols = Array(OC_ORDER, OC_NAME, OC_INVOICE, OC_ART, OC_UNIT, OC_SUPPLIER, OC_SELLER, OC_ODATE, OC_BUYER, OC_CAT, OC_ASSIGNED)
    For i = 0 To UBound(cols)
        PlanSetValue(SH_ORDERS, r, cols(i), sd(cols(i)), False)
    Next i
    ' the delivery's own sum, comment and delivery time, as the old table had them: kept as inputs of the row
    If OSh().getCellByPosition(OC_SUM, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then PlanInput(SH_ORDERS, r, OC_SUM)
    If OSh().getCellByPosition(OC_NOTE, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then PlanInput(SH_ORDERS, r, OC_NOTE)
    If OSh().getCellByPosition(OC_DAYS, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then PlanInput(SH_ORDERS, r, OC_DAYS)
    If mDoc <> "" Then PlanSetValue(SH_ORDERS, r, OC_DOC, mDoc, False)
    PlanReceiptValues(r, True, False)
    PlanSetValue(SH_ORDERS, r, OC_PLACE, mPlace, False)
    PlanSetValue(SH_ORDERS, r, OC_EI, canon, False)
    PlanSetValue(SH_ORDERS, r, OC_BLOCK, "OL" & olid, False)
    PlanDerived(SH_ORDERS, r, OC_STATUS, OS_ADD)
    PlanDerived(SH_ORDERS, r, OC_STOCK, bal)
    PlanDerived(SH_ORDERS, r, OC_CTL, own)
    PlanDerived(SH_ORDERS, r, OC_DUP, IIf(dup <> "", "Возможный дубль: " & dup, ""))
    PlanLockBits(SH_ORDERS, r, 0, OC_LAST, ORDER_LOCKS_POSTED)
    PlanSourceStatus(srcRow, od, rcv2, nodoc2, CStr(od(OD_CANCEL)))
    res = ApplyOperation(0, 0)
    ReceiptLegacyAdd = res
    Exit Function
EH:
    ReceiptLegacyAdd = "ERR-SYS:внутренняя ошибка переноса поступления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function
