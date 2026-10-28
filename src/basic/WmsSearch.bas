' WmsSearch — 1.1: the search of «Поиск» in the registry «Наличие» that understands the storekeeper. Several words (every one has to be found), the beginning of a word
' (the stem: «болты» finds «болт», «лампа» finds «лампочка»), the Latin letters that look like the Russian ones («M8x30» is «М8х30»), «м8 30» finds «М8х30», «8х30» too,
' «din 933» also finds what is written as «ISO 4017» or «ГОСТ 7805» (the sheet «Стандарты»), the words that mean the same (the sheet «Синонимы»: «саморез» = «шуруп»),
' «-слово» leaves out what has the word, «категория:метизы», «место:5-1», «артикул:gr-40», «серия:sn-7» narrow the search, the most fitting first; a word that is
' not found at all is looked for with one letter wrong («болд» → «болт»): «Возможно, вы искали». What the old search found by the text of the whole query is
' found now too (WmsFind adds it). The registry itself is only read; its rows are not parsed one by one: the query is taken apart (cheap), the rows are only compared
' with it, so the search stays fast on thousands of EIs. Library WmsPlus.
Option Explicit

Public Const SH_SYN = "Синонимы"
Public Const SX_MAX = 500

' the registry as it was read at the journal position gSxSeq, and the folded text of every row for the search
Global gSxSeq As Long
Global gSxN As Long
Global gSxData As Variant
Global gSxHay As Variant
Global gSxNm As Variant
' the raw texts the folded ones were made from (to keep the rows that did not change)
Global gSxRawN As Variant
Global gSxRawR As Variant
' the folded columns for the filters (place, category, article, serial) — made at the first filter
Global gSxCol As Variant
Global gSxColSeq As Long
' the synonyms
Global gSxSyn As Object
Global gSxGroups As Variant
Global gSxSynLoaded As Boolean
' the query as the groups of alternatives: a word (kind 0) or a token / a phrase with a number (kind 1)
Global gQAlts As Variant
Global gQKind As Variant
Global gQN As Long
Global gQNote As String
' True when the last query was words only (no «-слово», no filters)
Global gQWords As Boolean
' the hits of the last search
Global gSxHit As Variant
Global gSxScore As Variant
Global gSxNHit As Long
Global gSxNote As String
' the rows that were hit (by the row of the sheet, from 0) and the balance of them
Global gSxIs As Variant
Global gSxSum As Double
Global gSxNSum As Long
Global gSxUnit As String
Global gSxSame As Boolean

' ================================================================ words

Private Function CyrOnly(w As String) As Boolean
    Dim i As Long, c As String
    CyrOnly = (Len(w) > 0)
    For i = 1 To Len(w)
        c = Mid(w, i, 1)
        If Not ((c >= "а" And c <= "я") Or c = "ё") Then
            CyrOnly = False
            Exit Function
        End If
    Next i
End Function

' digits and points only («30», «0.9»)
Private Function IsNumber(w As String) As Boolean
    Dim i As Long, c As String
    IsNumber = (Len(w) > 0)
    For i = 1 To Len(w)
        c = Mid(w, i, 1)
        If Not ((c >= "0" And c <= "9") Or c = ".") Then
            IsNumber = False
            Exit Function
        End If
    Next i
End Function

Private Function HasDigit(w As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(w)
        c = Mid(w, i, 1)
        If c >= "0" And c <= "9" Then
            HasDigit = True
            Exit Function
        End If
    Next i
End Function

' the stem of a folded word: the ending of a case or of a number off (no shorter than three letters); the words that are not Russian are left as they are
Function Stem(ByVal w As String) As String
    Dim ends As Variant, i As Long
    Stem = w
    If Len(w) < 5 Or Not CyrOnly(w) Then Exit Function
    ends = Array("ами", "ями", "ого", "его", "ому", "ему", "ыми", "ими", "ов", "ев", "ей", "ой", "ый", "ий", "ая", "яя", "ое", "ее", "ые", "ие", "ых", "их", "ом", "ем", "ам", "ям", _
        "ах", "ях", "ую", "юю", "ию", "а", "я", "ы", "и", "е", "о", "у", "ю", "ь", "й")
    For i = 0 To UBound(ends)
        If Len(w) - Len(ends(i)) >= 3 And Right(w, Len(ends(i))) = ends(i) Then
            Stem = Left(w, Len(w) - Len(ends(i)))
            Exit Function
        End If
    Next i
End Function

' the signs between the words are spaces (after Fold); the block form makes it once for all the rows
Private Function Spaced(ByVal s As String) As String
    Dim k As Variant
    For Each k In Array(";", "(", ")", "/", Chr(34), "«", "»", ":", "|", "[", "]")
        s = Replace(s, CStr(k), " ")
    Next k
    Spaced = s
End Function

' ================================================================ the synonyms

Private Sub SynLoad()
    Dim sh As Object, last As Long, d As Variant, i As Long, j As Long, a As Variant, n As Long, g As String, w As String, gr() As String
    gSxSyn = New Collection
    gSxSynLoaded = True
    gSxGroups = Array()
    If Not gDoc.Sheets.hasByName(SH_SYN) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_SYN)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    If last > 300 Then last = 300
    d = sh.getCellRangeByPosition(0, 1, 0, last).getDataArray()
    ReDim gr(UBound(d))
    For i = 0 To UBound(d)
        a = Split(CStr(d(i)(0)), ",")
        g = ""
        For j = 0 To UBound(a)
            w = WmsNames.Fold(Trim(CStr(a(j))))
            If w <> "" Then g = g & IIf(g <> "", ",", "") & Stem(w)
        Next j
        If InStr(g, ",") > 0 Then
            gr(n) = g
            a = Split(g, ",")
            For j = 0 To UBound(a)
                On Error Resume Next
                gSxSyn.Add(n, CStr(a(j)))
                On Error GoTo 0
            Next j
            n = n + 1
        End If
    Next i
    If n > 0 Then
        ReDim Preserve gr(n - 1)
        gSxGroups = gr
    End If
End Sub

Sub OnSynChange(Optional oTarget As Variant)
    gSxSynLoaded = False
    gSxSeq = 0
End Sub

' the stems that stand for the stem st (st itself first); an array
Private Function SynAlts(st As String) As Variant
    Dim k As Long, a As Variant, i As Long, o As String
    If Not gSxSynLoaded Then SynLoad()
    On Error GoTo None
    k = gSxSyn.Item(st)
    a = Split(gSxGroups(k), ",")
    o = st
    For i = 0 To UBound(a)
        If CStr(a(i)) <> st Then o = o & "," & a(i)
    Next i
    SynAlts = Split(o, ",")
    Exit Function
None:
    SynAlts = Array(st)
End Function

' ================================================================ the index

' the registry read and the text of each of its rows made, when the journal moved since the last time. A row: its name (the rule of the single line), the EI, the article, the place,
' the category, the source, the serial — all folded (see WmsNames.Fold). A posting changes a few rows: the text of a row that is the same as at the last time is kept, the few
' that changed are folded one by one (a whole new registry is folded in two blocks — one pass of the language engine over a long text, not a call per row)
Private Sub Ensure()
    Dim st As Object, cur As Object, last As Long, i As Long, d As Variant, cEi As Long, cNm As Long, cArt As Long, cPl As Long, cCat As Long, cSrc As Long, cSer As Long
    Dim seq As Long, aN() As String, aR() As String, fn As Variant, fr As Variant, hay() As String, nm() As String, n As Long, r As Variant, has() As Boolean, nmv As String
    Dim c10 As String, c13 As String, c9 As String, c160 As String, pvN As Variant, pvR As Variant, pvH As Variant, pvM As Variant, nOld As Long, nChg As Long, kept() As Boolean
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = 0 Then seq = 1
    If seq = gSxSeq And gSxN > 0 Then Exit Sub
    st = WmsIssue.StockSheet()
    cur = st.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    nOld = gSxN
    If nOld > 0 And Not IsEmpty(gSxRawN) Then
        pvN = gSxRawN
        pvR = gSxRawR
        pvH = gSxHay
        pvM = gSxNm
    Else
        nOld = 0
    End If
    gSxN = 0
    gSxSeq = 0
    gSxColSeq = -1
    If last < 1 Then Exit Sub
    d = st.getCellRangeByPosition(0, 1, SI_COL0 + SI_COLS - 1, last).getDataArray()
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cPl = SC_PLACE
    cCat = SC_CAT
    cSrc = SC_SRC
    cSer = SI_COL0 + 8
    n = UBound(d) + 1
    ReDim aN(n - 1)
    ReDim aR(n - 1)
    ReDim has(n - 1)
    ReDim kept(n - 1)
    ' (a row of the array is taken once: d(i)(c) copies the whole row at every use; the names that need the rule of the single line are the rare ones)
    c10 = Chr(10)
    c13 = Chr(13)
    c9 = Chr(9)
    c160 = Chr(160)
    For i = 0 To n - 1
        r = d(i)
        If CStr(r(cEi)) <> "" Then
            nmv = CStr(r(cNm))
            If InStr(nmv, c10) + InStr(nmv, c13) + InStr(nmv, c9) + InStr(nmv, c160) + InStr(nmv, "  ") > 0 Then nmv = WmsNames.NormName(nmv)
            aN(i) = nmv
            aR(i) = r(cEi) & " " & r(cArt) & " " & r(cPl) & " " & r(cCat) & " " & r(cSrc) & " " & r(cSer)
            has(i) = True
            If i < nOld Then
                If CStr(pvN(i)) = nmv And CStr(pvR(i)) = aR(i) Then
                    kept(i) = True
                Else
                    nChg = nChg + 1
                End If
            Else
                nChg = nChg + 1
            End If
        End If
    Next i
    ReDim hay(n - 1)
    ReDim nm(n - 1)
    If nOld > 0 And nChg <= 200 Then
        For i = 0 To n - 1
            If has(i) Then
                If kept(i) Then
                    hay(i) = pvH(i)
                    nm(i) = pvM(i)
                Else
                    nm(i) = " " & Trim(Spaced(WmsNames.Fold(aN(i)))) & " "
                    hay(i) = nm(i) & Trim(Spaced(WmsNames.Fold(aR(i)))) & " "
                End If
            End If
        Next i
    Else
        fn = Split(Spaced(WmsNames.Fold(Join(aN, "¦"))), "¦")
        fr = Split(Spaced(WmsNames.Fold(Join(aR, "¦"))), "¦")
        For i = 0 To n - 1
            If has(i) Then
                nm(i) = " " & Trim(fn(i)) & " "
                hay(i) = nm(i) & Trim(fr(i)) & " "
            End If
        Next i
    End If
    gSxData = d
    gSxRawN = aN
    gSxRawR = aR
    gSxHay = hay
    gSxNm = nm
    gSxN = n
    gSxSeq = seq
End Sub

' the folded columns for the filters, once per state of the journal: 0 place, 1 category, 2 article, 3 serial
Private Sub EnsureCols()
    Dim k As Long, i As Long, a0() As String, a1() As String, a2() As String, a3() As String, d As Variant, n As Long, r As Variant, res(3) As Variant, c0 As Long, c1 As Long, c2 As Long, c3 As Long
    If gSxColSeq = gSxSeq Then Exit Sub
    d = gSxData
    n = gSxN
    c0 = SC_PLACE
    c1 = SC_CAT
    c2 = SC_ART
    c3 = SI_COL0 + 8
    ReDim a0(n - 1)
    ReDim a1(n - 1)
    ReDim a2(n - 1)
    ReDim a3(n - 1)
    For i = 0 To n - 1
        r = d(i)
        a0(i) = CStr(r(c0))
        a1(i) = CStr(r(c1))
        a2(i) = CStr(r(c2))
        a3(i) = CStr(r(c3))
    Next i
    res(0) = Split(Spaced(WmsNames.Fold(Join(a0, "¦"))), "¦")
    res(1) = Split(Spaced(WmsNames.Fold(Join(a1, "¦"))), "¦")
    res(2) = Split(Spaced(WmsNames.Fold(Join(a2, "¦"))), "¦")
    res(3) = Split(Spaced(WmsNames.Fold(Join(a3, "¦"))), "¦")
    gSxCol = res
    gSxColSeq = gSxSeq
End Sub

' the distance of two words (one letter put in, left out or changed — 1; two letters changed — 2): stops when it is more than nMax
Private Function Dist(a As String, b As String, nMax As Long) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, c As Long, prev() As Long, cur() As Long, m As Long
    la = Len(a)
    lb = Len(b)
    If Abs(la - lb) > nMax Then
        Dist = nMax + 1
        Exit Function
    End If
    ReDim prev(lb)
    ReDim cur(lb)
    For j = 0 To lb
        prev(j) = j
    Next j
    For i = 1 To la
        cur(0) = i
        m = cur(0)
        For j = 1 To lb
            c = IIf(Mid(a, i, 1) = Mid(b, j, 1), 0, 1)
            cur(j) = prev(j - 1) + c
            If prev(j) + 1 < cur(j) Then cur(j) = prev(j) + 1
            If cur(j - 1) + 1 < cur(j) Then cur(j) = cur(j - 1) + 1
            If cur(j) < m Then m = cur(j)
        Next j
        If m > nMax Then
            Dist = nMax + 1
            Exit Function
        End If
        For j = 0 To lb
            prev(j) = cur(j)
        Next j
    Next i
    Dist = prev(lb)
End Function

' the word of the registry closest to the stem st (no more than nMax letters away, its ending counted off; "" — none). The words are looked at as they come: a registry of
' thousands of names has this many words, and the search for a letter wrong is rare
Private Function Closest(st As String) As String
    Dim all As String, a As Variant, i As Long, w As String, dd As Long, d2 As Long, best As Long, nMax As Long
    nMax = IIf(Len(st) >= 7, 2, 1)
    best = nMax + 1
    all = Join(gSxNm, " ")
    a = Split(all, " ")
    For i = 0 To UBound(a)
        w = CStr(a(i))
        If Len(w) >= 4 And Len(w) <= Len(st) + nMax + 4 Then
            If Not IsNumber(w) Then
                dd = Dist(st, w, nMax)
                If Len(w) > Len(st) Then
                    d2 = Dist(st, Left(w, Len(st)), nMax)
                    If d2 < dd Then dd = d2
                End If
                If dd < best Then
                    best = dd
                    Closest = Stem(w)
                    If best = 1 Then Exit For
                End If
            End If
        End If
    Next i
End Function

' ================================================================ the query

' the query taken apart: the words (folded, stemmed), the words that must not be there, the filters
Private Sub Parse(sQuery As String, words() As String, ByRef nW As Long, excl() As String, ByRef nX As Long, ByRef fCat As String, ByRef fPlace As String, ByRef fArt As String, ByRef fSer As String)
    Dim q As String, a As Variant, i As Long, t As String, k As String, v As String, p As Long, inQ As Boolean, cur As String
    nW = 0
    nX = 0
    ReDim words(30)
    ReDim excl(30)
    q = WmsNames.NormName(sQuery)
    ' a phrase in quotes is one word of the search (its spaces are held)
    For i = 1 To Len(q)
        t = Mid(q, i, 1)
        If t = Chr(34) Or t = "«" Or t = "»" Then
            inQ = Not inQ
        ElseIf t = " " And inQ Then
            cur = cur & Chr(160)
        Else
            cur = cur & t
        End If
    Next i
    a = Split(cur, " ")
    For i = 0 To UBound(a)
        t = Trim(CStr(a(i)))
        If t <> "" Then
            p = InStr(t, ":")
            If p > 1 And p < Len(t) Then
                k = LCase(Left(t, p - 1))
                v = WmsNames.Fold(Replace(Mid(t, p + 1), Chr(160), " "))
                Select Case k
                Case "категория", "кат", "category"
                    fCat = v
                Case "место", "м", "place"
                    fPlace = v
                Case "артикул", "арт", "art"
                    fArt = v
                Case "серия", "серийный", "sn", "серия№"
                    fSer = v
                Case Else
                    GoTo PLAIN
                End Select
            ElseIf Left(t, 1) = "-" And Len(t) > 1 Then
                v = WmsNames.Fold(Replace(Mid(t, 2), Chr(160), " "))
                If v <> "" And nX <= UBound(excl) Then
                    excl(nX) = Stem(v)
                    nX = nX + 1
                End If
            Else
PLAIN:
                v = WmsNames.Fold(Replace(t, Chr(160), " "))
                If v <> "" And nW <= UBound(words) Then
                    words(nW) = Stem(v)
                    nW = nW + 1
                End If
            End If
        End If
    Next i
End Sub

' «din», «iso» (folded: «isо» — the last letter is the Russian one), «исо», «гост» → the standard as the sheet «Стандарты» writes it; "" — not a standard
Private Function StdPrefix(w As String) As String
    Select Case w
    Case "din"
        StdPrefix = "din"
    Case "isо", "iso", "исо"
        StdPrefix = "iso"
    Case "гост"
        StdPrefix = "гост"
    End Select
End Function

' a number of a standard: digits, and the year after a dash («7805-70»)
Private Function IsStdNum(w As String) As Boolean
    Dim i As Long, c As String
    IsStdNum = (Len(w) > 0)
    For i = 1 To Len(w)
        c = Mid(w, i, 1)
        If Not ((c >= "0" And c <= "9") Or (c = "-" And i > 1 And i < Len(w))) Then
            IsStdNum = False
            Exit Function
        End If
    Next i
End Function

' «м8», «м10», «м4.5» — the beginning of a size of a thread (folded)
Private Function IsSizeHead(w As String) As Boolean
    IsSizeHead = (Len(w) >= 2 And Left(w, 1) = "м")
    If IsSizeHead Then IsSizeHead = IsNumber(Mid(w, 2))
End Function

' «8х30», «4.2х32», «6х1.0х20» — a size without the letter
Private Function IsDimNum(w As String) As Boolean
    Dim p As Long
    p = InStr(w, "х")
    If p < 2 Or p = Len(w) Then Exit Function
    IsDimNum = IsNumber(Replace(w, "х", ""))
End Function

' the standard in a query and what is the same (the sheet «Стандарты») as the phrases to look for: «din 933», «din933», «iso 4017», «гост 7805» …
Private Function StdAlts(sPf As String, sNum As String) As Variant
    Dim sd As String, eq As String, a As Variant, i As Long, lst As String, f As String, bs As String, p As Long
    bs = sNum
    p = InStr(bs, "-")
    If p > 0 Then bs = Left(bs, p - 1)
    sd = sPf & " " & bs
    eq = WmsSpec.SpEquiv(sd)
    lst = sd
    If eq <> "" Then lst = lst & ";" & eq
    ' a number with a year («гост 7805-70») is looked for as it is written too
    If bs <> sNum Then lst = lst & ";" & sPf & " " & sNum
    a = Split(lst, ";")
    lst = ""
    For i = 0 To UBound(a)
        f = WmsNames.Fold(CStr(a(i)))
        lst = lst & IIf(lst <> "", ";", "") & f & ";" & Replace(f, " ", "")
        If Left(f, 3) = "isо" Then lst = lst & ";исо" & Mid(f, 4) & ";исо" & Replace(Mid(f, 4), " ", "")
    Next i
    If eq <> "" Then
        gQNote = gQNote & IIf(gQNote <> "", "; ", "") & UCase(sPf) & " " & bs & " = " & UCase(Replace(eq, ";", " / "))
    End If
    StdAlts = Split(lst, ";")
End Function

' a word of the query that is a standard written together («din933»): its prefix and number
Private Function SplitGlued(w As String, ByRef sPf As String, ByRef sNum As String) As Boolean
    Dim p As Long, pf As String
    For p = 3 To 4
        pf = StdPrefix(Left(w, p))
        If pf <> "" And Len(w) > p Then
            If IsStdNum(Mid(w, p + 1)) Then
                sPf = pf
                sNum = Mid(w, p + 1)
                SplitGlued = True
                Exit Function
            End If
        End If
    Next p
End Function

' the words of the query as the groups of alternatives: a standard with its number («din 933» → the same in ISO and ГОСТ), a size («м8 30» → «м8х30»), a size without
' the letter («8х30» → also «м8х30»), a word with the words that mean the same
Private Sub BuildGroups(words() As String, nW As Long)
    Dim j As Long, w As String, nx As String, pf As String, num As String, ga() As Variant, gk() As Long, alts As Variant, k As Long, s As String
    gQN = 0
    gQNote = ""
    If nW = 0 Then
        gQAlts = Array()
        gQKind = Array()
        Exit Sub
    End If
    ReDim ga(nW)
    ReDim gk(nW)
    j = 0
    Do While j < nW
        w = words(j)
        nx = ""
        If j + 1 < nW Then nx = words(j + 1)
        If StdPrefix(w) <> "" And IsStdNum(nx) Then
            ga(gQN) = StdAlts(StdPrefix(w), nx)
            gk(gQN) = 1
            j = j + 2
        ElseIf SplitGlued(w, pf, num) Then
            ga(gQN) = StdAlts(pf, num)
            gk(gQN) = 1
            j = j + 1
        ElseIf IsSizeHead(w) And IsNumber(nx) Then
            ga(gQN) = Array(w & "х" & nx)
            gk(gQN) = 1
            j = j + 2
        ElseIf IsDimNum(w) Then
            ga(gQN) = Array(w, "м" & w)
            gk(gQN) = 1
            j = j + 1
        ElseIf HasDigit(w) Then
            ga(gQN) = Array(w)
            gk(gQN) = 1
            j = j + 1
        Else
            alts = SynAlts(w)
            ga(gQN) = alts
            gk(gQN) = 0
            If UBound(alts) > 0 Then
                s = ""
                For k = 1 To UBound(alts)
                    s = s & IIf(s <> "", " / ", "") & alts(k)
                Next k
                gQNote = gQNote & IIf(gQNote <> "", "; ", "") & w & " = " & s
            End If
            j = j + 1
        End If
        gQN = gQN + 1
    Loop
    gQAlts = ga
    gQKind = gk
End Sub

' how well the word w (a stem, no digits) fits the text hay: 3 — the whole word, 2 — the beginning of a word, 1 — a part of a word; 0 — no
Private Function Fit(hay As String, w As String) As Long
    If InStr(hay, " " & w & " ") > 0 Then
        Fit = 3
    ElseIf InStr(hay, " " & w) > 0 Then
        Fit = 2
    ElseIf Len(w) >= 3 Then
        If InStr(hay, w) > 0 Then Fit = 1
    End If
End Function

' a token with a digit («м8», «30», «м8х30», «din 933», «4.2х32») that is in hay from the position p (InStr): 3 when it stands as a whole in hay — not a part of a longer number («30» is not «300», «м8» is not «м80»,
' «5» is not «0.5»); «30» also as a part of a size («м8х30»). 0 — no
Private Function FitTok(hay As String, w As String, ByVal p As Long) As Long
    Dim lw As Long, pre As String, pre2 As String, nx As String, nx2 As String, okB As Boolean, okA As Boolean, digit1 As Boolean
    lw = Len(w)
    digit1 = (Left(w, 1) >= "0" And Left(w, 1) <= "9")
    Do While p > 0
        pre = Mid(hay, p - 1, 1)
        okB = False
        okA = False
        If pre = " " Or pre = "-" Or pre = "(" Or pre = "/" Or pre = "," Then
            okB = True
        ElseIf pre = "х" And p >= 3 And digit1 Then
            pre2 = Mid(hay, p - 2, 1)
            okB = (pre2 >= "0" And pre2 <= "9")
        ElseIf pre = "." And p >= 3 And digit1 Then
            pre2 = Mid(hay, p - 2, 1)
            okB = Not (pre2 >= "0" And pre2 <= "9")
        End If
        If okB Then
            nx = Mid(hay, p + lw, 1)
            If nx >= "0" And nx <= "9" Then
                okA = False
            ElseIf nx = "." Then
                nx2 = Mid(hay, p + lw + 1, 1)
                okA = Not (nx2 >= "0" And nx2 <= "9")
            Else
                okA = True
            End If
        End If
        If okB And okA Then
            FitTok = 3
            Exit Function
        End If
        p = InStr(p + 1, hay, w)
    Loop
End Function

' the rows of the registry that fit the query: gSxHit, gSxScore, gSxNHit (the most fitting first). True when the query had something to search by
Private Function Hits(excl() As String, nX As Long, fCat As String, fPlace As String, fArt As String, fSer As String) As Boolean
    Dim i As Long, j As Long, k As Long, sc As Long, f As Long, g As Long, ok As Boolean, hay As String, d As Variant, n As Long, qty As Double, cQ As Long, nm As Variant, cols As Variant, hays As Variant
    Dim hit() As Long, score() As Long, flat() As String, st() As Long, en() As Long, kind() As Long, w As String, p As Long, nF As Long, al As Variant, w0() As String
    Hits = (gQN > 0 Or nX > 0 Or fCat <> "" Or fPlace <> "" Or fArt <> "" Or fSer <> "")
    gSxNHit = 0
    If Not Hits Or gSxN = 0 Then Exit Function
    d = gSxData
    nm = gSxNm
    hays = gSxHay
    cQ = SC_QTY
    If fCat <> "" Or fPlace <> "" Or fArt <> "" Or fSer <> "" Then
        EnsureCols()
        cols = gSxCol
    End If
    ' the alternatives of the groups in one flat list: group j is flat(st(j)) … flat(en(j))
    If gQN > 0 Then
        ReDim st(gQN)
        ReDim en(gQN)
        ReDim kind(gQN)
        ReDim w0(gQN)
        ' (the size is counted first: a ReDim Preserve inside a loop loses the elements written after it in this Basic)
        For j = 0 To gQN - 1
            al = gQAlts(j)
            nF = nF + UBound(al) + 1
        Next j
        ReDim flat(nF)
        nF = 0
        For j = 0 To gQN - 1
            al = gQAlts(j)
            st(j) = nF
            For k = 0 To UBound(al)
                flat(nF) = CStr(al(k))
                nF = nF + 1
            Next k
            en(j) = nF - 1
            kind(j) = gQKind(j)
            w0(j) = CStr(al(0))
        Next j
    End If
    ReDim hit(gSxN)
    ReDim score(gSxN)
    For i = 0 To gSxN - 1
        hay = hays(i)
        If hay <> "" Then
            ok = True
            sc = 0
            For j = 0 To gQN - 1
                f = 0
                For k = st(j) To en(j)
                    w = flat(k)
                    p = InStr(hay, w)
                    If p > 0 Then
                        If kind(j) = 0 Then
                            g = Fit(hay, w)
                        Else
                            g = FitTok(hay, w, p)
                        End If
                        If g > f Then f = g
                        If f = 3 Then Exit For
                    End If
                Next k
                If f = 0 Then
                    ok = False
                    Exit For
                End If
                sc = sc + f
                If kind(j) = 0 Then
                    If InStr(nm(i), " " & w0(j)) > 0 Then sc = sc + 1
                End If
            Next j
            If ok Then
                For j = 0 To nX - 1
                    If Fit(hay, excl(j)) > 0 Then
                        ok = False
                        Exit For
                    End If
                Next j
            End If
            If ok And fCat <> "" Then
                If InStr(cols(1)(i), fCat) = 0 Then ok = False
            End If
            If ok And fPlace <> "" Then
                If InStr(" " & cols(0)(i), " " & fPlace) = 0 Then ok = False
            End If
            If ok And fArt <> "" Then
                If InStr(cols(2)(i), fArt) = 0 Then ok = False
            End If
            If ok And fSer <> "" Then
                If InStr(cols(3)(i), fSer) = 0 Then ok = False
            End If
            If ok Then
                qty = 0
                If VarType(d(i)(cQ)) = 5 Then qty = d(i)(cQ)
                If qty > 0.0000005 Then sc = sc + 2
                If sc > 99 Then sc = 99
                hit(n) = i
                score(n) = sc
                n = n + 1
            End If
        End If
    Next i
    gSxNHit = n
    SortHits(hit, score, n)
End Function

' the hits from the highest score to the lowest (the same score — the order of the registry): counting by the score, no sort of n·log n steps
Private Sub SortHits(hit() As Long, score() As Long, n As Long)
    Dim cnt(100) As Long, i As Long, s As Long, start(100) As Long, pos As Long, outH() As Long, outS() As Long
    If n = 0 Then Exit Sub
    ReDim outH(n)
    ReDim outS(n)
    For i = 0 To n - 1
        cnt(score(i)) = cnt(score(i)) + 1
    Next i
    pos = 0
    For s = 100 To 0 Step -1
        start(s) = pos
        pos = pos + cnt(s)
    Next s
    For i = 0 To n - 1
        s = score(i)
        outH(start(s)) = hit(i)
        outS(start(s)) = s
        start(s) = start(s) + 1
    Next i
    gSxHit = outH
    gSxScore = outS
End Sub

' ================================================================ the rows for the page «Поиск»

' The smart search of the registry for the page «Поиск»: the rows found are put into out (the most fitting first), nAll — how many, totals — the balance of the found EIs
' when they have one unit; sNote — what to tell under the title (what the words stand for, «Возможно, вы искали»). False — the query has nothing for this search to hold on to.
Function SmartRun(sQuery As String, out() As Variant, ByRef nOut As Long, ByRef nAll As Long, ByRef totals As String, ByRef sNote As String) As Boolean
    Dim words() As String, excl() As String, nW As Long, nX As Long, fCat As String, fPlace As String, fArt As String, fSer As String, i As Long, k As Long, sum As Double, nSum As Long
    Dim unit As String, sameUnit As Boolean, d As Variant, r As Variant, q As Variant, det As String, cEi As Long, cNm As Long, cQ As Long, cSt As Long, cUn As Long, cPl As Long, cSrc As Long, cArt As Long, cCat As Long, sStorno As String, sStock As String, fixed As String, c As String, changed As Boolean, al As Variant, allNm As String, hit() As Boolean, w As String
    On Error GoTo EH
    Ensure()
    If gSxN = 0 Then Exit Function
    Parse(sQuery, words, nW, excl, nX, fCat, fPlace, fArt, fSer)
    BuildGroups(words, nW)
    gQWords = (nW > 0 And nX = 0 And fCat = "" And fPlace = "" And fArt = "" And fSer = "")
    If Not Hits(excl, nX, fCat, fPlace, fArt, fSer) Then Exit Function
    SmartRun = True
    sNote = gQNote
    ' nothing found: a word that is not in the registry is looked for with a letter wrong
    If gSxNHit = 0 And gQN > 0 Then
        allNm = Join(gSxNm, " ")
        For i = 0 To gQN - 1
            If gQKind(i) = 0 Then
                al = gQAlts(i)
                w = CStr(al(0))
                If UBound(al) = 0 And Len(w) >= 4 And InStr(allNm, w) = 0 Then
                    c = Closest(w)
                    If c <> "" Then
                        gQAlts(i) = Array(c)
                        changed = True
                        fixed = fixed & IIf(fixed <> "", " ", "") & c
                    End If
                End If
            End If
        Next i
        If changed Then
            Hits(excl, nX, fCat, fPlace, fArt, fSer)
            If gSxNHit > 0 Then sNote = "Возможно, вы искали: «" & fixed & "» — показано по нему."
        End If
    End If
    d = gSxData
    ReDim hit(gSxN + 1)
    ' (the columns as local numbers: a constant of another library costs ~30 µs a use)
    cEi = SC_EI
    cNm = SC_NAME
    cQ = SC_QTY
    cSt = SC_STATE
    cUn = SC_UNIT
    cPl = SC_PLACE
    cSrc = SC_SRC
    cArt = SC_ART
    cCat = SC_CAT
    sStorno = EI_ST_STORNO
    sStock = SH_STOCK
    For k = 0 To gSxNHit - 1
        i = gSxHit(k)
        hit(i + 1) = True
        nAll = nAll + 1
        r = d(i)
        q = ""
        If VarType(r(cQ)) = 5 Then q = r(cQ)
        If nOut < SX_MAX Then
            ' one hit as the row of the table of the page (the form of WmsFind.RowView): sheet, row, EI, name, quantity, unit, place, details, "", the hidden source (sheet, row, EI)
            det = CStr(r(cSt))
            If CStr(r(cSrc)) <> "" Then det = det & " · " & r(cSrc)
            If Trim(CStr(r(cArt))) <> "" Then det = det & " · арт. " & Trim(CStr(r(cArt)))
            If Trim(CStr(r(cCat))) <> "" Then det = det & " · " & Trim(CStr(r(cCat)))
            out(nOut) = Array(sStock, i + 2, CStr(r(cEi)), WmsNames.NormNameFast(CStr(r(cNm))), q, CStr(r(cUn)), CStr(r(cPl)), det, "", sStock, i + 1, CStr(r(cEi)))
            nOut = nOut + 1
        End If
        If VarType(r(cQ)) = 5 And CStr(r(cSt)) <> sStorno Then
            nSum = nSum + 1
            If nSum = 1 Then
                unit = CStr(r(cUn))
                sameUnit = True
            ElseIf CStr(r(cUn)) <> unit Then
                sameUnit = False
            End If
            sum = sum + r(cQ)
        End If
    Next k
    gSxIs = hit
    gSxSum = sum
    gSxNSum = nSum
    gSxUnit = unit
    gSxSame = sameUnit
    If nSum > 0 Then
        If sameUnit Then
            totals = "ЕИ " & NumText(nSum) & ", остаток " & WmsIssue.QtyText(sum) & " " & unit
        Else
            totals = "ЕИ " & NumText(nSum) & " (единицы разные — общей суммы нет)"
        End If
    End If
    Exit Function
EH:
    gSxNote = "ERR:" & Error$ & " (строка " & Erl & ")"
    SmartRun = False
End Function

' did the last smart search hit the row r of «Наличие» (from 0)? and the balance of its hits
Function IsHit(r As Long) As Boolean
    If IsEmpty(gSxIs) Then Exit Function
    If r >= 0 And r <= UBound(gSxIs) Then IsHit = gSxIs(r)
End Function

Function HitSum() As Double
    HitSum = gSxSum
End Function

Function HitNSum() As Long
    HitNSum = gSxNSum
End Function

Function HitUnit() As String
    HitUnit = gSxUnit
End Function

Function HitSame() As Boolean
    HitSame = gSxSame
End Function

' ================================================================ the words in the other sheets

' How many groups the last query has, and whether the words may be looked for in the other sheets (there are words; no «-слово» and no filters — those are for the registry)
Function WordGroups() As Long
    WordGroups = IIf(gQWords, gQN, 0)
End Function

' the regular expression of the alternative w (folded) for the engine of the sheet: a Russian letter also as the Latin one that looks like it, a point also as a comma, a space as
' any spaces; kind 1 — a token with a number: not a part of a longer number
Private Function RxOf(w As String, kind As Long) As String
    Dim i As Long, c As String, o As String
    For i = 1 To Len(w)
        c = Mid(w, i, 1)
        Select Case c
        Case "а"
            o = o & "[aа]"
        Case "в"
            o = o & "[bв]"
        Case "с"
            o = o & "[cс]"
        Case "е"
            o = o & "[eеё]"
        Case "н"
            o = o & "[hн]"
        Case "к"
            o = o & "[kк]"
        Case "м"
            o = o & "[mм]"
        Case "о"
            o = o & "[oо]"
        Case "р"
            o = o & "[pр]"
        Case "т"
            o = o & "[tт]"
        Case "х"
            o = o & "[xх×*]"
        Case "у"
            o = o & "[yу]"
        Case "."
            o = o & "[.,]"
        Case " "
            o = o & "\s+"
        Case "\", "^", "$", "|", "?", "+", "(", ")", "[", "]", "{", "}", "*"
            o = o & "\" & c
        Case Else
            o = o & c
        End Select
    Next i
    If kind = 1 Then o = "(?<![0-9])" & o & "(?![0-9])"
    RxOf = o
End Function

' the regular expression of the group j of the last query (its alternatives: the words that mean the same, the equivalent standards)
Function GroupRegex(j As Long) As String
    Dim al As Variant, k As Long, o As String
    If j < 0 Or j >= gQN Then Exit Function
    al = gQAlts(j)
    For k = 0 To UBound(al)
        o = o & IIf(o <> "", "|", "") & RxOf(CStr(al(k)), CLng(gQKind(j)))
    Next k
    GroupRegex = "(" & o & ")"
End Function

' test seam: the EIs found for a query, as «ЕИ,ЕИ,…» (the most fitting first); the note after «|»
Function TestSmart(sQuery As String) As String
    Dim out() As Variant, nOut As Long, nAll As Long, tot As String, note As String, i As Long, s As String
    ReDim out(SX_MAX)
    gSxNote = ""
    If Not SmartRun(sQuery, out, nOut, nAll, tot, note) Then
        TestSmart = IIf(Left(gSxNote, 4) = "ERR:", gSxNote, "NONE")
        Exit Function
    End If
    For i = 0 To nOut - 1
        s = s & IIf(s <> "", ",", "") & out(i)(2)
    Next i
    TestSmart = s & "|" & note & "|" & nAll & "|" & tot
End Function
