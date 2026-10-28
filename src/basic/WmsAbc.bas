' WmsAbc — 1.1: «ABC-анализ» (what is issued most often, by the number of issues over a period) and «Без движения» (what lies in the stock and has not moved for a given number of days).
' Both are views made from «Наличие» and the journal (the issues, their deletions, the returns): they only read; they write nothing into the books, nothing into the journal. A name with an
' article is one position (as in «Метизы»). Library WmsPlus.
Option Explicit

Public Const SH_ABC = "ABC-анализ"
Public Const AB_PAR_ROW = 1
Public Const AB_STAT_ROW = 2
Public Const AB_HEAD = 3
Public Const AB_ROW0 = 4
Public Const AB_COLS = 10
Public Const AB_MAX = 5000

Public Const SH_IDLE = "Без движения"
Public Const ID_PAR_ROW = 1
Public Const ID_STAT_ROW = 2
Public Const ID_HEAD = 3
Public Const ID_ROW0 = 4
Public Const ID_COLS = 8
Public Const ID_MAX = 5000

' the positions: name, article, unit, balance, the places, the last receipt / issue / return (a Calc date, 0 — none), the issues of the period (number and quantity)
Global gAgN As Long
Global gAgName As Variant
Global gAgArt As Variant
Global gAgUnit As Variant
Global gAgBal As Variant
Global gAgPlaces As Variant
Global gAgRcv As Variant
Global gAgIss As Variant
Global gAgRet As Variant
Global gAgCnt As Variant
Global gAgQty As Variant
Global gAbSeq As Long
Global gIdSeq As Long
Global gAbBusy As Boolean
Global gIdBusy As Boolean

Private Function CollIdx(coll As Object, key As String) As Long
    On Error GoTo Miss
    CollIdx = CLng(coll.Item(key))
    Exit Function
Miss:
    CollIdx = -1
End Function

' the date of a text «2026-10-26» as a Calc date; 0 when it is not one
Private Function IsoDate(s As String) As Double
    On Error GoTo Bad
    If Len(s) < 10 Then Exit Function
    IsoDate = CDbl(DateSerial(Val(Left(s, 4)), Val(Mid(s, 6, 2)), Val(Mid(s, 9, 2))))
    Exit Function
Bad:
    IsoDate = 0
End Function

' The positions of the registry and what the journal says about them: the number and the quantity of the issues since dSince (a Calc date), the last issue, the last return.
' OK:<positions> / ERR:…
Function Aggregate(dSince As Double) As String
    Dim st As Object, last As Long, d As Variant, i As Long, r As Variant, nm As String, art As String, key As String, g As Long, n As Long, coll As Object, q As Double, pl As String, dt As Double
    Dim cEi As Long, cNm As Long, cArt As Long, cUn As Long, cQty As Long, cPl As Long, cRcv As Long
    Dim gN() As String, gA() As String, gU() As String, gB() As Double, gP() As String, gR() As Double, gI() As Double, gT() As Double, gC() As Long, gQ() As Double
    Dim eiG() As Long, fl As Variant, f As Long, txt As String, lines As Variant, seq As Long, ts As String, inst As String, sType As String, fields As Variant, canon As String, en As Long
    Dim iG() As Long, iQ() As Double, iD() As Double, iNo() As String, nIss As Long, delNos As String, k As Long, c10 As String, c13 As String, c9 As String, c160 As String
    On Error GoTo EH
    gAgN = 0
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    If last < 1 Then
        Aggregate = "OK:0"
        Exit Function
    End If
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cUn = SC_UNIT
    cQty = SC_QTY
    cPl = SC_PLACE
    cRcv = SI_COL0 + 4
    c10 = Chr(10)
    c13 = Chr(13)
    c9 = Chr(9)
    c160 = Chr(160)
    d = st.getCellRangeByPosition(0, 1, cRcv, last).getDataArray()
    ReDim gN(UBound(d))
    ReDim gA(UBound(d))
    ReDim gU(UBound(d))
    ReDim gB(UBound(d))
    ReDim gP(UBound(d))
    ReDim gR(UBound(d))
    ReDim gI(UBound(d))
    ReDim gT(UBound(d))
    ReDim gC(UBound(d))
    ReDim gQ(UBound(d))
    ReDim eiG(UBound(d) + 2)
    coll = New Collection
    For i = 0 To UBound(d)
        r = d(i)
        If CStr(r(cEi)) <> "" Then
            nm = CStr(r(cNm))
            If InStr(nm, c10) + InStr(nm, c13) + InStr(nm, c9) + InStr(nm, c160) + InStr(nm, "  ") > 0 Then nm = WmsNames.NormName(nm)
            art = Trim(CStr(r(cArt)))
            key = LCase(nm) & "|" & LCase(art)
            g = CollIdx(coll, key)
            If g < 0 Then
                g = n
                coll.Add(g, key)
                gN(g) = nm
                gA(g) = art
                gU(g) = Trim(CStr(r(cUn)))
                n = n + 1
            End If
            eiG(i + 1) = g + 1
            q = 0
            If VarType(r(cQty)) = 5 Then q = r(cQty)
            gB(g) = gB(g) + q
            pl = Trim(CStr(r(cPl)))
            If q > 0 And pl <> "" Then
                If InStr("; " & gP(g) & ";", "; " & pl & ";") = 0 Then gP(g) = gP(g) & IIf(gP(g) <> "", "; ", "") & pl
            End If
            If VarType(r(cRcv)) = 5 Then
                If r(cRcv) > gR(g) Then gR(g) = Int(r(cRcv))
            End If
        End If
    Next i
    ' the journal: the issues, the deletions of them, the returns
    ReDim iG(50000)
    ReDim iQ(50000)
    ReDim iD(50000)
    ReDim iNo(50000)
    If gJDir <> "" Then
        If gSfa.exists(gJDir) Then
            fl = gSfa.getFolderContents(gJDir, False)
            For f = 0 To UBound(fl)
                If LCase(Right(CStr(fl(f)), 4)) = ".csv" Then
                    txt = WmsAudit.ReadText(CStr(fl(f)))
                    If InStr(txt, "ISSUE") + InStr(txt, "RETURN") + InStr(txt, "RECEIPT") + InStr(txt, "SP_REFILL") > 0 Then
                        lines = Split(Replace(txt, Chr(13), ""), Chr(10))
                        For i = 0 To UBound(lines)
                            If InStr(lines(i), ";ISSUE") > 0 Or InStr(lines(i), ";RETURN") > 0 Or InStr(lines(i), "RECEIPT") > 0 Or InStr(lines(i), ";SP_REFILL") > 0 Then
                                If WmsJournal.ParseLine(CStr(lines(i)), seq, ts, inst, sType, fields) Then
                                    If sType = "RECEIPT" Or sType = "RECEIPT_ADD" Or sType = "SP_RECEIPT" Or sType = "SP_REFILL" Then
                                        ' the date of a receipt of an EI: the last receipt of the position
                                        If WmsOrders.StrictEI(WmsView.JField(fields, "EI"), en) Then
                                            If en >= 1 And en <= UBound(eiG) Then
                                                g = eiG(en) - 1
                                                If g >= 0 Then
                                                    dt = IsoDate(WmsView.JField(fields, "DATE"))
                                                    If dt > gR(g) Then gR(g) = dt
                                                End If
                                            End If
                                        End If
                                    ElseIf sType = "ISSUE" Or sType = "ISSUE_DEL" Or sType = "RETURN" Then
                                        canon = WmsView.JField(fields, "EI")
                                        If sType = "ISSUE_DEL" Then
                                            delNos = delNos & ";" & WmsView.JField(fields, "NO") & ";"
                                        ElseIf WmsOrders.StrictEI(canon, en) Then
                                            If en >= 1 And en <= UBound(eiG) Then
                                                g = eiG(en) - 1
                                                If g >= 0 Then
                                                    dt = IsoDate(WmsView.JField(fields, "DATE"))
                                                    If sType = "RETURN" Then
                                                        If dt > gT(g) Then gT(g) = dt
                                                    ElseIf nIss < UBound(iG) Then
                                                        iG(nIss) = g
                                                        iQ(nIss) = Val(WmsView.JField(fields, "QTY"))
                                                        iD(nIss) = dt
                                                        iNo(nIss) = WmsView.JField(fields, "NO")
                                                        nIss = nIss + 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
            Next f
        End If
    End If
    For k = 0 To nIss - 1
        If iNo(k) = "" Or InStr(delNos, ";" & iNo(k) & ";") = 0 Then
            g = iG(k)
            If iD(k) > gI(g) Then gI(g) = iD(k)
            If iD(k) >= dSince Then
                gC(g) = gC(g) + 1
                gQ(g) = gQ(g) + iQ(k)
            End If
        End If
    Next k
    gAgName = gN
    gAgArt = gA
    gAgUnit = gU
    gAgBal = gB
    gAgPlaces = gP
    gAgRcv = gR
    gAgIss = gI
    gAgRet = gT
    gAgCnt = gC
    gAgQty = gQ
    gAgN = n
    Aggregate = "OK:" & n
    Exit Function
EH:
    Aggregate = "ERR:" & Error$ & " (строка " & Erl & ")"
End Function

' the indexes idx(0…n-1) of the positions put in the order of k1 descending, then k2 descending, then the name (a shell sort)
Private Sub SortIdx(idx() As Long, n As Long, k1 As Variant, k2 As Variant, nm As Variant)
    Dim gap As Long, i As Long, j As Long, t As Long, ok As Boolean
    gap = n \ 2
    Do While gap > 0
        For i = gap To n - 1
            t = idx(i)
            j = i
            Do While j >= gap
                ok = True
                If k1(idx(j - gap)) < k1(t) Then
                    ok = False
                ElseIf k1(idx(j - gap)) = k1(t) Then
                    If k2(idx(j - gap)) < k2(t) Then
                        ok = False
                    ElseIf k2(idx(j - gap)) = k2(t) Then
                        If LCase(CStr(nm(idx(j - gap)))) > LCase(CStr(nm(t))) Then ok = False
                    End If
                End If
                If ok Then Exit Do
                idx(j) = idx(j - gap)
                j = j - gap
            Loop
            idx(j) = t
        Next i
        gap = gap \ 2
    Loop
End Sub

' ================================================================ ABC

Sub BtnAbcRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = AbcRefresh()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub OnAbcFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gAbBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gAbSeq And gAbSeq > 0 Then Exit Sub
    AbcRefresh()
    Exit Sub
EH:
End Sub

Sub OnAbcChange(Optional oTarget As Variant)
    Dim ra As Object
    If gAbBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> AB_PAR_ROW Or (ra.Column <> 1 And ra.Column <> 3 And ra.Column <> 5) Then Exit Sub
    WmsInit()
    AbcRefresh()
    Exit Sub
EH:
End Sub

' a whole number from a cell of the parameters (the default when it is empty or not a number; min…max)
Private Function ParNum(sh As Object, c As Integer, r As Integer, def As Double, lo As Double, hi As Double) As Double
    Dim cell As Object, s As String, v As Double
    cell = sh.getCellByPosition(c, r)
    v = def
    If cell.getType() = com.sun.star.table.CellContentType.VALUE Then
        v = cell.getValue()
    Else
        s = Replace(Trim(cell.getString()), ",", ".")
        If s <> "" And IsNumeric(s) Then v = Val(s)
    End If
    If v < lo Then v = lo
    If v > hi Then v = hi
    ParNum = v
End Function

' The positions issued over the last months, by the number of issues, with the class: A — the head that makes up the first part of all the issues (80 % by default), B — the next part (to 95 %),
' C — the rest. OK:<rows> / ERR:…
Function AbcRefresh() As String
    Dim sh As Object, months As Double, pA As Double, pB As Double, since As Double, res As String, i As Long, n As Long, idx() As Long, total As Double, cum As Double, before As Double, cls As String
    Dim out() As Variant, rows As Long, g As Long, idxAll() As Long, cnt As Variant, qty As Variant, nm As Variant, k As Long, vArt As Variant, vUnit As Variant, vBal As Variant
    On Error GoTo EH
    gAbBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_ABC) Then
        AbcRefresh = "ERR:в книге нет листа «" & SH_ABC & "» — книга не обновлена до 1.1"
        gAbBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_ABC)
    months = Int(ParNum(sh, 1, AB_PAR_ROW, 12, 1, 120))
    pA = ParNum(sh, 3, AB_PAR_ROW, 80, 1, 99)
    pB = ParNum(sh, 5, AB_PAR_ROW, 95, 2, 100)
    If pB <= pA Then pB = pA + 1
    since = CDbl(DateAdd("m", -months, Date()))
    res = Aggregate(since)
    If Left(res, 4) = "ERR:" Then
        AbcRefresh = res
        gAbBusy = False
        Exit Function
    End If
    cnt = gAgCnt
    qty = gAgQty
    nm = gAgName
    vArt = gAgArt
    vUnit = gAgUnit
    vBal = gAgBal
    ReDim idxAll(IIf(gAgN > 0, gAgN - 1, 0))
    For i = 0 To gAgN - 1
        If cnt(i) > 0 Then
            idxAll(n) = i
            total = total + cnt(i)
            n = n + 1
        End If
    Next i
    If n > 1 Then SortIdx(idxAll, n, cnt, qty, nm)
    rows = n
    If rows > AB_MAX Then rows = AB_MAX
    ReDim out(IIf(rows > 0, rows - 1, 0))
    For k = 0 To rows - 1
        g = idxAll(k)
        before = cum * 100 / total
        If before < pA Then
            cls = "A"
        ElseIf before < pB Then
            cls = "B"
        Else
            cls = "C"
        End If
        cum = cum + cnt(g)
        out(k) = Array(k + 1, cls, nm(g), vArt(g), cnt(g), WmsPack.PkRound(CDbl(qty(g)), 3), vUnit(g), Int(cnt(g) * 1000 / total + 0.5) / 10, Int(cum * 1000 / total + 0.5) / 10, WmsPack.PkRound(CDbl(vBal(g)), 3))
    Next k
    sh.getCellRangeByPosition(0, AB_ROW0, AB_COLS - 1, AB_ROW0 + AB_MAX).clearContents(1 + 2 + 4 + 16)
    If rows > 0 Then sh.getCellRangeByPosition(0, AB_ROW0, AB_COLS - 1, AB_ROW0 + rows - 1).setDataArray(out)
    sh.getCellByPosition(0, AB_STAT_ROW).setString("Позиций с выдачами за " & CLng(months) & " " & WmsView.Plural(CLng(months), "месяц", "месяца", "месяцев") & ": " & n & "; всего выдач: " & CLng(total) _
        & ". A — до " & CLng(pA) & " % выдач, B — до " & CLng(pB) & " %, C — остальное. Считано из журнала на " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
    gAbSeq = CLng(SysNum(SK_LAST_SEQ))
    If gAbSeq = 0 Then gAbSeq = 1
    gAbBusy = False
    AbcRefresh = "OK:" & n
    gUiLastMsg = AbcRefresh
    Exit Function
EH:
    gAbBusy = False
    AbcRefresh = "ERR:ABC-анализ не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = AbcRefresh
End Function

' ================================================================ without movement

Sub BtnIdleRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = IdleRefresh()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub OnIdleFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gIdBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gIdSeq And gIdSeq > 0 Then Exit Sub
    IdleRefresh()
    Exit Sub
EH:
End Sub

Sub OnIdleChange(Optional oTarget As Variant)
    Dim ra As Object
    If gIdBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> ID_PAR_ROW Or ra.Column <> 1 Then Exit Sub
    WmsInit()
    IdleRefresh()
    Exit Sub
EH:
End Sub

' The positions with a balance that have not moved for the given number of days (the last receipt, issue or return counts as a movement); the longest first. OK:<rows> / ERR:…
Function IdleRefresh() As String
    Dim sh As Object, days As Double, res As String, i As Long, n As Long, idx() As Long, lastMove As Double, how As String, today As Double, out() As Variant, k As Long, g As Long
    Dim ageA() As Double, ageB() As Double, dm() As Double, hw() As String, nm As Variant, vArt As Variant, vUnit As Variant, vBal As Variant, vPl As Variant, vRcv As Variant, vIss As Variant, vRet As Variant
    On Error GoTo EH
    gIdBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_IDLE) Then
        IdleRefresh = "ERR:в книге нет листа «" & SH_IDLE & "» — книга не обновлена до 1.1"
        gIdBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_IDLE)
    days = Int(ParNum(sh, 1, ID_PAR_ROW, 90, 1, 3650))
    today = CDbl(Date())
    res = Aggregate(today - 36500)
    If Left(res, 4) = "ERR:" Then
        IdleRefresh = res
        gIdBusy = False
        Exit Function
    End If
    nm = gAgName
    vArt = gAgArt
    vUnit = gAgUnit
    vBal = gAgBal
    vPl = gAgPlaces
    vRcv = gAgRcv
    vIss = gAgIss
    vRet = gAgRet
    ReDim idx(IIf(gAgN > 0, gAgN - 1, 0))
    ReDim ageA(IIf(gAgN > 0, gAgN - 1, 0))
    ReDim ageB(IIf(gAgN > 0, gAgN - 1, 0))
    ReDim dm(IIf(gAgN > 0, gAgN - 1, 0))
    ReDim hw(IIf(gAgN > 0, gAgN - 1, 0))
    For i = 0 To gAgN - 1
        If vBal(i) > 0.0000005 Then
            lastMove = vRcv(i)
            how = "приход"
            If vIss(i) > lastMove Then
                lastMove = vIss(i)
                how = "выдача"
            End If
            If vRet(i) > lastMove Then
                lastMove = vRet(i)
                how = "возврат"
            End If
            If lastMove = 0 Then
                ageA(i) = 100000
                how = "дат нет"
            Else
                ageA(i) = today - lastMove
            End If
            If ageA(i) >= days Then
                idx(n) = i
                dm(i) = lastMove
                hw(i) = how
                n = n + 1
            End If
        End If
    Next i
    If n > 1 Then SortIdx(idx, n, ageA, ageB, nm)
    k = n
    If k > ID_MAX Then k = ID_MAX
    ReDim out(IIf(k > 0, k - 1, 0))
    For i = 0 To k - 1
        g = idx(i)
        out(i) = Array(nm(g), vArt(g), WmsPack.PkRound(CDbl(vBal(g)), 3), vUnit(g), vPl(g), IIf(dm(g) > 0, dm(g), ""), IIf(ageA(g) < 100000, ageA(g), ""), hw(g))
    Next i
    sh.getCellRangeByPosition(0, ID_ROW0, ID_COLS - 1, ID_ROW0 + ID_MAX).clearContents(1 + 2 + 4 + 16)
    If k > 0 Then
        sh.getCellRangeByPosition(0, ID_ROW0, ID_COLS - 1, ID_ROW0 + k - 1).setDataArray(out)
    End If
    sh.getCellByPosition(0, ID_STAT_ROW).setString("Позиций с остатком, которые не двигались " & CLng(days) & " " & WmsView.Plural(CLng(days), "день", "дня", "дней") & " и больше: " & n & IIf(n > ID_MAX, " (показаны первые " & ID_MAX & ")", "") _
        & ". Движение — последний приход, выдача или возврат. Считано на " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
    gIdSeq = CLng(SysNum(SK_LAST_SEQ))
    If gIdSeq = 0 Then gIdSeq = 1
    gIdBusy = False
    IdleRefresh = "OK:" & n
    gUiLastMsg = IdleRefresh
    Exit Function
EH:
    gIdBusy = False
    IdleRefresh = "ERR:список «Без движения» не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = IdleRefresh
End Function
