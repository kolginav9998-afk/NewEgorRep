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

' ================================================================ the totals

' the sums by name as a text «шт=105;м=3» (an array does not come back from a procedure here): q of name un added
Private Function SumAdd(acc As String, un As String, q As Double) As String
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
Private Function TopText(acc As String, n As Integer) As String
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
    Dim iNo() As String, iWho() As String, iQty() As Double, iUn() As String, n2 As Long, k As Long, out As String, nNew As Long
    On Error GoTo EH
    ReDim iNo(20000)
    ReDim iWho(20000)
    ReDim iQty(20000)
    ReDim iUn(20000)
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
                                Case "RECEIPT", "RECEIPT_ADD"
                                    nRcv = nRcv + 1
                                    rcvU = SumAdd(rcvU, un, q)
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
                                        nIss = nIss + 1
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
                                Case "RETURN"
                                    nRet = nRet + 1
                                    retU = SumAdd(retU, un, q)
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
    ' the issues of the period without those deleted (storno)
    n2 = 0
    For k = 0 To nIss - 1
        If iNo(k) = "" Or InStr(delNos, ";" & iNo(k) & ";") = 0 Then
            n2 = n2 + 1
            issU = SumAdd(issU, iUn(k), iQty(k))
            If Trim(iWho(k)) <> "" Then whoAcc = SumAdd(whoAcc, Trim(iWho(k)), 1)
        End If
    Next k
    nDays = UBound(Split(Replace(days, "||", "|"), "|")) - 1
    If days = "" Then nDays = 0
    If nOps = 0 Then Exit Function
    out = "• Операций проведено: " & nOps & " в " & nDays & " " & WmsView.Plural(nDays, "день", "дня", "дней") & " (журнал № " & seqFirst & "–" & seqLast & ")"
    If nCars > 0 Then out = out & Chr(10) & "• Принято машин с поставками: " & nCars
    If nRcv > 0 Then out = out & Chr(10) & "• Приёмка по заказам: принято позиций — " & nRcv & "; количество: " & SumText(rcvU)
    If nSp > 0 Then out = out & Chr(10) & "• Иной приход: записей — " & nSp
    If n2 > 0 Then
        out = out & Chr(10) & "• Выдачи: " & n2 & "; количество: " & SumText(issU)
        If whoAcc <> "" Then out = out & Chr(10) & "• Кому выдавали чаще всего (по числу выдач): " & TopText(whoAcc, 5)
    End If
    If nRet > 0 Then out = out & Chr(10) & "• Возвраты: " & nRet & "; количество: " & SumText(retU)
    k = nMove + nWo + nInv + nSplit + nSup + nKit + nAsm
    If k > 0 Then
        w = ""
        If nMove > 0 Then w = w & IIf(w <> "", ", ", "") & "перемещений — " & nMove
        If nWo > 0 Then w = w & IIf(w <> "", ", ", "") & "списаний — " & nWo & " (" & SumText(woU) & ")"
        If nInv > 0 Then w = w & IIf(w <> "", ", ", "") & "позиций инвентаризации — " & nInv
        If nSplit > 0 Then w = w & IIf(w <> "", ", ", "") & "разделений ЕИ — " & nSplit
        If nSup > 0 Then w = w & IIf(w <> "", ", ", "") & "возвратов поставщику — " & nSup
        If nKit > 0 Then w = w & IIf(w <> "", ", ", "") & "сборок комплектов — " & nKit
        If nAsm > 0 Then w = w & IIf(w <> "", ", ", "") & "сборок деталей — " & nAsm
        out = out & Chr(10) & "• Корректировки склада: " & w
    End If
    If nDoc > 0 Then out = out & Chr(10) & "• Оформлено документов выдачи / актов: " & nDoc
    If nFix + nIssDel + nCancel > 0 Then out = out & Chr(10) & "• Исправления учёта: исправлено операций — " & nFix & ", выдач удалено (сторно) — " & nIssDel & IIf(nCancel > 0, ", отменено заказов — " & nCancel, "")
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
