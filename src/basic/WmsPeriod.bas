' WmsPeriod — 1.1: the working calendar and the totals of a period. «Календарь» — the days that are not as the schedule 5/2 says (a holiday on a weekday, a working Saturday): the storekeeper's
' own sheet. IsWorkDay is what the reports of the day ask («задачи на следующий рабочий день», «выдачи за рабочие дни»). PeriodTotals reads the journal for a period and counts what was
' done in it — the lines for the end of «Отчёт за период», of a week or of a month. The journal is only read. Library WmsPlus.
Option Explicit

Public Const SH_CAL = "Календарь"
Private Const CA_ROWS = 400

' the mark of the calendar for day d: 1 — a working day (listed as one), -1 — a day off (listed), 0 — the day is not listed
Private Function CalMark(d As Double) As Integer
    Dim sh As Object, dat As Variant, i As Long, v As Double, msg As String, t As String, c As Variant
    If Not gDoc.Sheets.hasByName(SH_CAL) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_CAL)
    dat = sh.getCellRangeByPosition(0, 1, 1, CA_ROWS).getDataArray()
    For i = 0 To UBound(dat)
        c = dat(i)(0)
        v = 0
        If VarType(c) = 5 Then
            v = Int(c)
        ElseIf Trim(CStr(c)) <> "" Then
            If Not WmsOrders.ParseDateTextAs(Trim(CStr(c)), "дата календаря", v, msg) Then v = 0
            v = Int(v)
        End If
        If v = Int(d) Then
            t = LCase(Trim(CStr(dat(i)(1))))
            If InStr(t, "раб") > 0 Then CalMark = 1 Else CalMark = -1
            Exit Function
        End If
    Next i
End Function

' True when day d is a working day: the schedule 5/2 (Monday … Friday), corrected by «Календарь» (a holiday on a weekday, a working Saturday)
Function IsWorkDay(d As Double) As Boolean
    Dim m As Integer
    m = CalMark(d)
    If m = 1 Then
        IsWorkDay = True
    ElseIf m = -1 Then
        IsWorkDay = False
    Else
        IsWorkDay = (Weekday(d, 2) <= 5)
    End If
End Function

' ================================================================ 1.4.0: names, prices and money for the texts of the reports

' The reports are read by people who do not use the table: money is written in whole roubles, the goods by their names. The names of
' the positions and the unit prices (of the order a position came with) are read from the sheets once for a report.
Global gRfOk As Boolean
Global gRfHi As Long
Global gRfName() As String
Global gRfPrice() As Double
Global gRfUnit() As String

' «12 340 руб.» — whole roubles with the groups of digits (the spaces, whatever the language of the PC)
Function MoneyRub(x As Double) As String
    Dim s As String
    s = WmsPrint.Money(CDbl(Int(Abs(x) + 0.5)))
    If Right(s, 3) = ",00" Then s = Left(s, Len(s) - 3)
    MoneyRub = IIf(x < 0, "−", "") & s & " руб."
End Function

' a new report reads the names and the prices again
Sub RefsReset()
    gRfOk = False
End Sub

Private Sub RefsLoad()
    Dim st As Object, sh As Object, last As Long, d As Variant, i As Long, n As Long, hi As Long, p As Variant
    gRfOk = True
    gRfHi = 0
    ReDim gRfName(0)
    ReDim gRfPrice(0)
    ReDim gRfUnit(0)
    On Error GoTo EH
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    If last < 1 Then Exit Sub
    d = st.getCellRangeByPosition(0, 1, SC_NAME, last).getDataArray()
    For i = 0 To UBound(d)
        If WmsOrders.StrictEI(CStr(d(i)(SC_EI)), n) Then
            If n > hi Then hi = n
        End If
    Next i
    If hi < 1 Then Exit Sub
    ReDim gRfName(hi)
    ReDim gRfPrice(hi)
    ReDim gRfUnit(hi)
    gRfHi = hi
    For i = 0 To UBound(d)
        If WmsOrders.StrictEI(CStr(d(i)(SC_EI)), n) Then gRfName(n) = Trim(CStr(d(i)(SC_NAME)))
    Next i
    sh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    d = sh.getCellRangeByPosition(OC_UNIT, 1, OC_EI, last).getDataArray()
    For i = 0 To UBound(d)
        If WmsOrders.StrictEI(CStr(d(i)(OC_EI - OC_UNIT)), n) Then
            If n <= hi Then
                p = d(i)(OC_PRICE - OC_UNIT)
                If VarType(p) = 5 Then
                    If p > 0 Then
                        gRfPrice(n) = p
                        gRfUnit(n) = CStr(d(i)(0))
                    End If
                End If
            End If
        End If
    Next i
    Exit Sub
EH:
End Sub

' the name of the position with the code number n; "" when it is not known
Function RefName(n As Long) As String
    If Not gRfOk Then RefsLoad()
    If n >= 1 And n <= gRfHi Then RefName = gRfName(n)
End Function

' the price of one unit of the position n by the order it came with — when the unit asked is the unit of that order (otherwise the
' amount would be a guess); 0 — not known
Function RefPrice(n As Long, sUnit As String) As Double
    If Not gRfOk Then RefsLoad()
    If n < 1 Or n > gRfHi Then Exit Function
    If gRfPrice(n) <= 0 Then Exit Function
    If WmsPack.UnitKey(sUnit) <> WmsPack.UnitKey(gRfUnit(n)) Then Exit Function
    RefPrice = gRfPrice(n)
End Function

' a key for the lists «name=count;…»: no separators of the list inside
Function RefKey(s As String) As String
    RefKey = Trim(Replace(Replace(s, ";", ","), "=", "-"))
End Function

' ================================================================ the totals

' the sums by name as a text «шт=105;м=3» (an array does not come back from a procedure here): q of name un added
Function SumAdd(acc As String, un As String, q As Double) As String
    Dim a As Variant, i As Integer, found As Boolean, r As String, p As Integer, v As Double
    If acc <> "" Then
        a = Split(acc, ";")
        For i = 0 To UBound(a)
            p = InStr(a(i), "=")
            If Left(a(i), p - 1) = un Then
                v = Val(Mid(a(i), p + 1)) + q
                r = r & IIf(r <> "", ";", "") & un & "=" & Str(v)
                found = True
            Else
                r = r & IIf(r <> "", ";", "") & a(i)
            End If
        Next i
    End If
    If Not found Then r = r & IIf(r <> "", ";", "") & un & "=" & Str(q)
    SumAdd = r
End Function

' «105 шт, 3 м» from the sums by unit
Private Function SumText(acc As String) As String
    Dim a As Variant, i As Integer, p As Integer, s As String, un As String
    If acc = "" Then Exit Function
    a = Split(acc, ";")
    For i = 0 To UBound(a)
        p = InStr(a(i), "=")
        un = Left(a(i), p - 1)
        s = s & IIf(s <> "", ", ", "") & WmsIssue.QtyText(WmsPack.PkRound(Val(Mid(a(i), p + 1)), 3)) & IIf(un <> "", " " & un, "")
    Next i
    SumText = s
End Function

' the first n names of a list «name=count;…» by the count, as «name — 5, name — 3»
Function TopText(acc As String, n As Integer) As String
    Dim a As Variant, i As Integer, j As Integer, best As Integer, used As String, s As String, p As Integer, v As Double, bv As Double, k As Integer
    If acc = "" Then Exit Function
    a = Split(acc, ";")
    For k = 1 To n
        best = -1
        bv = -1
        For i = 0 To UBound(a)
            If InStr(used, "|" & i & "|") = 0 Then
                p = InStr(a(i), "=")
                v = Val(Mid(a(i), p + 1))
                If v > bv Then
                    bv = v
                    best = i
                End If
            End If
        Next i
        If best < 0 Then Exit For
        used = used & "|" & best & "|"
        p = InStr(a(best), "=")
        s = s & IIf(s <> "", ", ", "") & Left(a(best), p - 1) & " — " & WmsIssue.QtyText(bv)
    Next k
    TopText = s
End Function

' What was done from d0 to d1 (inclusive), by the journal: the lines «• …» one under another (Chr(10)); "" — nothing was done. nDays — the days that had operations
Function PeriodTotals(d0 As Double, d1 As Double, ByRef nDays As Long) As String
    Dim d As Double, files As String, f As Variant, url As String, txt As String, a As Variant, i As Long, seq As Long, ts As String, inst As String, sType As String, fields As Variant
    Dim lo As String, hi As String, day As String, days As String, q As Double, un As String, w As String, nOps As Long, seqFirst As Long, seqLast As Long
    Dim nRcv As Long, rcvU As String, nSp As Long, nIss As Long, issU As String, nRet As Long, retU As String, nCars As Long, nMove As Long, nWo As Long, woU As String, nInv As Long
    Dim nSplit As Long, nSup As Long, nKit As Long, nAsm As Long, nFix As Long, nDel As Long, nCancel As Long, whoAcc As String, delNos As String, nDoc As Long, docs As String, nIssDel As Long
    Dim iNo() As String, iWho() As String, iQty() As Double, iUn() As String, iEi() As Long, n2 As Long, k As Long, out As String, nNew As Long
    Dim rcvSum As Double, nNoPr As Long, issSum As Double, nIssNoPr As Long, supAcc As String, itemAcc As String, price As Double, nm As String, dq As Double, nDocChk As Long, nDiff As Long, shortU As String, overU As String
    Dim rcEi() As Long, rcQ() As Double, rcU() As String, rcSp() As String, rcPr() As Double, rcDq() As Double, rcDel() As Boolean, nRc As Long, slot As Long, k2 As Long, j As Long
    Dim rtNo() As Long, rtQ() As Double, rtU() As String, rtDel() As Boolean, nRt As Long
    On Error GoTo EH
    ReDim rcEi(500)
    ReDim rcQ(500)
    ReDim rcU(500)
    ReDim rcSp(500)
    ReDim rcPr(500)
    ReDim rcDq(500)
    ReDim rcDel(500)
    ReDim rtNo(500)
    ReDim rtQ(500)
    ReDim rtU(500)
    ReDim rtDel(500)
    ReDim iNo(20000)
    ReDim iWho(20000)
    ReDim iQty(20000)
    ReDim iUn(20000)
    ReDim iEi(20000)
    RefsReset()
    lo = Format(CDate(d0), "YYYY-MM-DD")
    hi = Format(CDate(d1), "YYYY-MM-DD")
    ' the files of the months of the period
    d = d0
    Do While d <= d1 + 31
        url = WmsJournal.JournalFileName(CDate(d))
        If InStr(files, "|" & url & "|") = 0 Then files = files & "|" & url & "|"
        d = d + 28
        If d > d1 + 28 Then Exit Do
    Loop
    url = WmsJournal.JournalFileName(CDate(d1))
    If InStr(files, "|" & url & "|") = 0 Then files = files & "|" & url & "|"
    For Each f In Split(Replace(files, "||", "|"), "|")
        If CStr(f) <> "" Then
            url = gJDir & CStr(f)
            If gSfa.exists(url) Then
                txt = WmsAudit.ReadText(url)
                a = Split(Replace(txt, Chr(13), ""), Chr(10))
                For i = 0 To UBound(a)
                    If a(i) <> "" Then
                        If WmsJournal.ParseLine(CStr(a(i)), seq, ts, inst, sType, fields) Then
                            day = Left(ts, 10)
                            If day >= lo And day <= hi Then
                                nOps = nOps + 1
                                If InStr(days, "|" & day & "|") = 0 Then days = days & "|" & day & "|"
                                If seqFirst = 0 Or seq < seqFirst Then seqFirst = seq
                                If seq > seqLast Then seqLast = seq
                                q = Val(WmsView.JField(fields, "QTY"))
                                un = WmsView.JField(fields, "UNIT")
                                Select Case sType
                                Case "RECEIPT", "RECEIPT_ADD", "RECEIPT_FIX"
                                    ' 1.4.0: the receipts of the period are kept one by one: a corrected one replaces its figures, a deleted one is not counted
                                    slot = -1
                                    If sType = "RECEIPT_FIX" Then
                                        k2 = WmsView.EiNo(WmsView.JField(fields, "EI"), 99999999)
                                        For j = nRc - 1 To 0 Step -1
                                            If rcEi(j) = k2 And Not rcDel(j) Then
                                                slot = j
                                                Exit For
                                            End If
                                        Next j
                                    Else
                                        If nRc > UBound(rcEi) Then
                                            ReDim Preserve rcEi(2 * nRc)
                                            ReDim Preserve rcQ(2 * nRc)
                                            ReDim Preserve rcU(2 * nRc)
                                            ReDim Preserve rcSp(2 * nRc)
                                            ReDim Preserve rcPr(2 * nRc)
                                            ReDim Preserve rcDq(2 * nRc)
                                            ReDim Preserve rcDel(2 * nRc)
                                        End If
                                        slot = nRc
                                        nRc = nRc + 1
                                        rcEi(slot) = WmsView.EiNo(WmsView.JField(fields, "EI"), 99999999)
                                        rcDel(slot) = False
                                    End If
                                    If slot >= 0 Then
                                        rcQ(slot) = q
                                        rcU(slot) = un
                                        rcSp(slot) = RefKey(WmsView.JField(fields, "SUPPLIER"))
                                        w = WmsView.JField(fields, "PRICE")
                                        rcPr(slot) = -1
                                        If w <> "" Then rcPr(slot) = Val(w)
                                        w = WmsView.JField(fields, "DOC_QTY")
                                        rcDq(slot) = -1
                                        If w <> "" Then rcDq(slot) = Val(w)
                                    End If
                                Case "RECEIPT_DEL"
                                    k2 = WmsView.EiNo(WmsView.JField(fields, "EI"), 99999999)
                                    For j = nRc - 1 To 0 Step -1
                                        If rcEi(j) = k2 And Not rcDel(j) Then
                                            rcDel(j) = True
                                            Exit For
                                        End If
                                    Next j
                                Case "SP_RECEIPT", "SP_REFILL"
                                    nSp = nSp + 1
                                Case "CAR_ARRIVE"
                                    nCars = nCars + 1
                                Case "ISSUE"
                                    ' (the arrays are long enough for a month of any real work; beyond them the issue is counted without its name)
                                    If nIss < UBound(iNo) Then
                                        iNo(nIss) = WmsView.JField(fields, "NO")
                                        iWho(nIss) = WmsView.JField(fields, "WHO")
                                        iQty(nIss) = q
                                        iUn(nIss) = un
                                        iEi(nIss) = WmsView.EiNo(WmsView.JField(fields, "EI"), 99999999)
                                        nIss = nIss + 1
                                    End If
                                Case "ISSUE_FIX"
                                    ' 1.4.0: an issue of the period that was corrected counts as corrected — its new quantity, unit and recipient replace the posted ones
                                    ' (a correction of an issue posted before the period is only counted below)
                                    w = WmsView.JField(fields, "NO")
                                    If w <> "" Then
                                        For k = 0 To nIss - 1
                                            If iNo(k) = w Then
                                                iWho(k) = WmsView.JField(fields, "WHO")
                                                iQty(k) = q
                                                iUn(k) = un
                                                If WmsView.JField(fields, "EI") <> "" Then iEi(k) = WmsView.EiNo(WmsView.JField(fields, "EI"), 99999999)
                                            End If
                                        Next k
                                    End If
                                Case "ISSUE_DEL"
                                    delNos = delNos & ";" & WmsView.JField(fields, "NO") & ";"
                                    nIssDel = nIssDel + 1
                                Case "ISSUE_DOC"
                                    w = WmsView.JField(fields, "DOC")
                                    If w <> "" And InStr(";" & docs & ";", ";" & w & ";") = 0 Then
                                        docs = docs & IIf(docs <> "", ";", "") & w
                                        nDoc = nDoc + 1
                                    End If
                                Case "RETURN", "RETURN_FIX"
                                    slot = -1
                                    k2 = CLng(Val(WmsView.JField(fields, "RET")))
                                    If sType = "RETURN_FIX" Then
                                        For j = nRt - 1 To 0 Step -1
                                            If rtNo(j) = k2 And Not rtDel(j) Then
                                                slot = j
                                                Exit For
                                            End If
                                        Next j
                                    Else
                                        If nRt > UBound(rtNo) Then
                                            ReDim Preserve rtNo(2 * nRt)
                                            ReDim Preserve rtQ(2 * nRt)
                                            ReDim Preserve rtU(2 * nRt)
                                            ReDim Preserve rtDel(2 * nRt)
                                        End If
                                        slot = nRt
                                        nRt = nRt + 1
                                        rtNo(slot) = k2
                                        rtDel(slot) = False
                                    End If
                                    If slot >= 0 Then
                                        rtQ(slot) = q
                                        ' (the correction of a return has no unit of its own: the unit of the return stays)
                                        If un <> "" Then rtU(slot) = un
                                    End If
                                Case "RETURN_DEL"
                                    k2 = CLng(Val(WmsView.JField(fields, "RET")))
                                    For j = nRt - 1 To 0 Step -1
                                        If rtNo(j) = k2 And Not rtDel(j) Then
                                            rtDel(j) = True
                                            Exit For
                                        End If
                                    Next j
                                Case "MOVE"
                                    nMove = nMove + 1
                                Case "WRITE_OFF"
                                    nWo = nWo + 1
                                    woU = SumAdd(woU, un, q)
                                Case "INV_ADJ"
                                    nInv = nInv + 1
                                Case "SPLIT"
                                    nSplit = nSplit + 1
                                Case "SUP_RET"
                                    nSup = nSup + 1
                                Case "KIT_BUILD"
                                    nKit = nKit + 1
                                Case "ASM_BUILD"
                                    nAsm = nAsm + 1
                                Case "ORDER_CANCEL", "ORDER_CANCEL_REST"
                                    nCancel = nCancel + 1
                                End Select
                                If Right(sType, 4) = "_FIX" Then nFix = nFix + 1
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next f
    ' the receipts and the returns of the period, the corrected and the deleted ones taken into account
    For k = 0 To nRc - 1
        If Not rcDel(k) Then
            nRcv = nRcv + 1
            rcvU = SumAdd(rcvU, rcU(k), rcQ(k))
            If rcSp(k) <> "" Then supAcc = SumAdd(supAcc, rcSp(k), 1)
            If rcPr(k) >= 0 Then
                rcvSum = rcvSum + rcPr(k) * rcQ(k)
            Else
                nNoPr = nNoPr + 1
            End If
            If rcDq(k) >= 0 Then
                nDocChk = nDocChk + 1
                If Abs(rcDq(k) - rcQ(k)) > 0.0000001 Then
                    nDiff = nDiff + 1
                    If rcDq(k) > rcQ(k) Then
                        shortU = SumAdd(shortU, rcU(k), rcDq(k) - rcQ(k))
                    Else
                        overU = SumAdd(overU, rcU(k), rcQ(k) - rcDq(k))
                    End If
                End If
            End If
        End If
    Next k
    For k = 0 To nRt - 1
        If Not rtDel(k) Then
            nRet = nRet + 1
            retU = SumAdd(retU, rtU(k), rtQ(k))
        End If
    Next k
    ' the issues of the period without those deleted (storno)
    n2 = 0
    For k = 0 To nIss - 1
        If iNo(k) = "" Or InStr(delNos, ";" & iNo(k) & ";") = 0 Then
            n2 = n2 + 1
            issU = SumAdd(issU, iUn(k), iQty(k))
            If Trim(iWho(k)) <> "" Then whoAcc = SumAdd(whoAcc, RefKey(Trim(iWho(k))), 1)
            ' 1.4.0: the amount by the price of the order the position came with (only when the unit is that of the order) and the goods most issued
            price = RefPrice(iEi(k), iUn(k))
            If price > 0 Then
                issSum = issSum + price * iQty(k)
            Else
                nIssNoPr = nIssNoPr + 1
            End If
            nm = RefName(iEi(k))
            If nm <> "" Then itemAcc = SumAdd(itemAcc, RefKey(nm), 1)
        End If
    Next k
    nDays = UBound(Split(Replace(days, "||", "|"), "|")) - 1
    If days = "" Then nDays = 0
    If nOps = 0 Then Exit Function
    out = "• Проведено складских операций: " & nOps & " (дней с операциями — " & nDays & ")"
    If nCars > 0 Then out = out & Chr(10) & "• Принято автомашин с грузом от поставщиков: " & nCars
    If nRcv > 0 Then
        w = "• Приёмка товара по заказам: принято позиций — " & nRcv & "; количество: " & SumText(rcvU)
        If rcvSum > 0 Then w = w & "; на сумму около " & MoneyRub(rcvSum) & IIf(nNoPr > 0, " (по " & nNoPr & " поз. цена не указана)", "")
        out = out & Chr(10) & w
        If supAcc <> "" Then out = out & Chr(10) & "• Поставщики (по числу принятых позиций): " & TopText(supAcc, 5)
        If nDocChk > 0 Then
            If nDiff = 0 Then
                out = out & Chr(10) & "• Сверка с документами поставщиков: проверено позиций — " & nDocChk & ", расхождений нет"
            Else
                w = ""
                If shortU <> "" Then w = "недопоставлено — " & SumText(shortU)
                If overU <> "" Then w = w & IIf(w <> "", "; ", "") & "излишек — " & SumText(overU)
                out = out & Chr(10) & "• Сверка с документами поставщиков: проверено позиций — " & nDocChk & ", расхождения по " & nDiff & " (" & w & ")"
            End If
        End If
    End If
    If nSp > 0 Then out = out & Chr(10) & "• Прочее оприходование (остатки по результатам пересчёта, принятое от производства и др.): записей — " & nSp
    If n2 > 0 Then
        w = "• Отпуск товара со склада: выдач — " & n2 & "; отпущено: " & SumText(issU)
        If issSum > 0 Then w = w & "; на сумму около " & MoneyRub(issSum) & IIf(nIssNoPr > 0, " (оценено без " & nIssNoPr & " поз.)", "")
        out = out & Chr(10) & w
        If whoAcc <> "" Then out = out & Chr(10) & "• Кому выдавали чаще всего (по числу выдач): " & TopText(whoAcc, 5)
        If itemAcc <> "" Then out = out & Chr(10) & "• Чаще всего выдавали (по числу выдач): " & TopText(itemAcc, 5)
    End If
    If nRet > 0 Then out = out & Chr(10) & "• Принято возвратов от получателей: " & nRet & "; количество: " & SumText(retU)
    k = nMove + nWo + nInv + nSplit + nSup + nKit + nAsm
    If k > 0 Then
        w = ""
        If nMove > 0 Then w = w & IIf(w <> "", "; ", "") & "перемещений между местами хранения — " & nMove
        If nWo > 0 Then w = w & IIf(w <> "", "; ", "") & "списаний — " & nWo & " (" & SumText(woU) & ")"
        If nInv > 0 Then w = w & IIf(w <> "", "; ", "") & "позиций пересчитано при инвентаризации — " & nInv
        If nSplit > 0 Then w = w & IIf(w <> "", "; ", "") & "позиций разделено на части — " & nSplit
        If nSup > 0 Then w = w & IIf(w <> "", "; ", "") & "возвратов поставщикам — " & nSup
        If nKit > 0 Then w = w & IIf(w <> "", "; ", "") & "наборов скомплектовано — " & nKit
        If nAsm > 0 Then w = w & IIf(w <> "", "; ", "") & "узлов собрано из деталей — " & nAsm
        out = out & Chr(10) & "• Внутрискладские операции: " & w
    End If
    If nDoc > 0 Then out = out & Chr(10) & "• Оформлено документов выдачи и актов: " & nDoc
    If nFix + nIssDel + nCancel > 0 Then out = out & Chr(10) & "• Исправления в учёте: скорректировано записей — " & nFix & ", выдач аннулировано (сторно) — " & nIssDel & IIf(nCancel > 0, ", отменено заказов — " & nCancel, "")
    PeriodTotals = out
    Exit Function
EH:
    PeriodTotals = "• Итоги не посчитаны полностью: " & Error$ & " (строка " & Erl & ")"
End Function

' test seam
Function TestTotals(d0 As Double, d1 As Double) As String
    Dim n As Long
    TestTotals = PeriodTotals(d0, d1, n) & "|" & n
End Function

Function TestIsWork(d As Double) As String
    TestIsWork = IIf(IsWorkDay(d), "1", "0")
End Function
