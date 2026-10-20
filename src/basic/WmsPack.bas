' WmsPack — учёт по документу, выдача — как удобно (PRIME 0.8, D-099). The balance of an EI is always in the unit of its
' document (1 шт). An EI may have a packing (_PACK): «в 1 шт — 4 кан» (a document of 1 piece, 4 canisters in fact),
' «1 шт = 6 м» (a profile cut into parts). A quantity of an issue, a return, a write-off or a count can then be typed in
' the parts («1 кан», «1,5 м», «3 канистры»), as a percent of what came in by the EI («25%», v0.1 §21) or in the unit of
' the document (0,25); WMS turns it into the unit of the document so that the last part closes the balance exactly (no
' «хвост» 0,001 after the third of three canisters). The packing is set and changed by the operation EI_PACK only.
Option Explicit

' how the last quantity was given (QtyEIText / QtyEICell): "" — a number in the unit of the document, PARTS — parts of the
' packing, PCT — a percent of what came in; the journal of the operation keeps it (PlanQtyFields)
Global gPkKind As String
Global gPkParts As Double
Global gPkUnit As String
Global gPkN As Double
Global gPkPct As Double
Global gPkBase As Double
' 1.0: the last quantity was a bare number («1», no unit word) — for an EI with a packing it may mean 1 шт or 1 кан (QtyAsk)
Global gPkPlain As Boolean

' ================================================================ the packing of an EI

Function PackSheet() As Object
    If gDoc.Sheets.hasByName(SH_PACK) Then PackSheet = gDoc.Sheets.getByName(SH_PACK)
End Function

' the packing of EI n: pu — the name of a part («кан», «м»), pn — parts in one unit of the document. False when none
Function PackOf(n As Long, ByRef pu As String, ByRef pn As Double) As Boolean
    Dim sh As Object, d As Variant
    pu = ""
    pn = 0
    If n < 1 Then Exit Function
    sh = PackSheet()
    If IsNull(sh) Then Exit Function
    d = sh.getCellRangeByPosition(0, n, PK_N, n).getDataArray()(0)
    If CStr(d(PK_EI)) <> WmsIssue.EiCanon(n) Or VarType(d(PK_N)) <> 5 Then Exit Function
    If d(PK_N) <= 0 Or Trim(CStr(d(PK_UNIT))) = "" Then Exit Function
    pu = Trim(CStr(d(PK_UNIT)))
    pn = d(PK_N)
    PackOf = True
End Function

' the precision of a count of parts: a unit of the document is kept to 0,001, so 0,001 × pn parts is what can be told
' apart — two decimals of a part while pn < 10, one below 100, whole parts above
Function PackDec(pn As Double) As Integer
    If pn < 10 Then
        PackDec = 2
    ElseIf pn < 100 Then
        PackDec = 1
    Else
        PackDec = 0
    End If
End Function

Function PkRound(x As Double, dec As Integer) As Double
    Dim m As Double
    m = 10 ^ dec
    If x >= 0 Then PkRound = Int(x * m + 0.5) / m Else PkRound = -Int(-x * m + 0.5) / m
End Function

' the parts in a quantity b of the unit of the document
Function PartsOf(b As Double, pn As Double) As Double
    PartsOf = PkRound(b * pn, PackDec(pn))
End Function

' «3 кан» / «4,5 м» — a number of parts as text
Function PartsText(x As Double, pu As String) As String
    PartsText = WmsIssue.QtyText(x) & " " & pu
End Function

' a quantity of EI n with its parts when it has a packing: «0,75 шт (3 кан)»; without — «0,75 шт»
Function QtyWithParts(n As Long, q As Double, sUnit As String) As String
    Dim pu As String, pn As Double
    QtyWithParts = WmsIssue.QtyText(q) & IIf(sUnit <> "", " " & sUnit, "")
    If PackOf(n, pu, pn) Then QtyWithParts = QtyWithParts & " (" & PartsText(PartsOf(q, pn), pu) & ")"
End Function

' the packing as a line for people: «в 1 шт — 4 кан»
Function PackLine(n As Long, sUnit As String) As String
    Dim pu As String, pn As Double
    If PackOf(n, pu, pn) Then PackLine = "в 1 " & sUnit & " — " & PartsText(pn, pu)
End Function

' ================================================================ the words of a unit

' the stem of a word: without the endings of Russian forms (канистра, канистры, канистр → канистр)
Private Function PkStem(ByVal w As String) As String
    Do While Len(w) > 2 And InStr("аяыиоеуюьй", Right(w, 1)) > 0
        w = Left(w, Len(w) - 1)
    Loop
    PkStem = w
End Function

' the word w (as typed) names the unit u: «кан», «канистра», «канистры» for «кан» or «канистра»; «м», «метр», «метра» for
' «м» (but not «мм», «см»); «шт», «штук» for «шт»
Function UnitMatch(ByVal w As String, ByVal u As String) As Boolean
    Dim syn As Variant, i As Integer, a As Variant, sw As String, su As String
    w = LCase(Trim(w))
    u = LCase(Trim(u))
    If Right(w, 1) = "." Then w = Left(w, Len(w) - 1)
    If Right(u, 1) = "." Then u = Left(u, Len(u) - 1)
    If w = "" Or u = "" Then Exit Function
    If w = u Then
        UnitMatch = True
        Exit Function
    End If
    ' the short units and their words
    syn = Array("м|метр", "л|литр", "г|грамм", "т|тонн", "кг|килограмм", "шт|штук", "уп|упаков", "пар|пар", "кан|канистр", "бут|бутыл", "рул|рулон", _
        "меш|мешк", "кор|коробк", "пач|пачк", "лист|лист", "бан|банк", "мл|миллилитр")
    For i = 0 To UBound(syn)
        a = Split(syn(i), "|")
        If u = a(0) Or Left(u, Len(a(1))) = a(1) Then
            If w = a(0) Or Left(w, Len(a(1))) = a(1) Then
                UnitMatch = True
                Exit Function
            End If
        End If
    Next i
    If Len(u) < 3 Or Len(w) < 3 Then Exit Function
    sw = PkStem(w)
    su = PkStem(u)
    UnitMatch = (Left(sw, Len(su)) = su) Or (Left(su, Len(sw)) = sw)
End Function

' a unit word: it starts with a letter («кан», «м») or is the percent sign
Private Function PkIsWord(w As String) As Boolean
    Dim ch As String
    ch = Left(w, 1)
    PkIsWord = (ch = "%") Or (LCase(ch) <> UCase(ch))
    ' «1e3» stays a number in the exponent form (refused by the strict parser), not «1 e3»
    If (ch = "e" Or ch = "E") And Len(w) > 1 Then
        If Mid(w, 2, 1) >= "0" And Mid(w, 2, 1) <= "9" Then PkIsWord = False
    End If
End Function

' «1,5 м» → the number «1,5» and the word «м»; «25%» → «25» and «%»; «2» → «2» and ""
Private Sub PkSplit(ByVal s As String, ByRef sNum As String, ByRef sWord As String)
    Dim i As Long, ch As String
    s = Trim(Replace(s, Chr(160), " "))
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or ch = "," Or ch = ".") Then Exit For
    Next i
    sNum = Trim(Left(s, i - 1))
    sWord = Trim(Mid(s, i))
End Sub

' a unit word with something after it: a digit or a second word («кан 2», «кан и»); a final dot is part of the word («кан.»)
Private Function PkHasTail(w As String) As Boolean
    Dim i As Long, ch As String
    If w = "%" Then Exit Function
    For i = 1 To Len(w)
        ch = Mid(w, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = " " Or ch = Chr(9) Or ch = "%" Or ch = "," Or ch = ";" Or ch = "+" Then
            PkHasTail = True
            Exit Function
        End If
    Next i
End Function

' the word is exactly the unit of the document or of the packing («м2», «м³» keep their digits)
Private Function PkExactUnit(w As String, n As Long, sUnit As String) As Boolean
    Dim pu As String, pn As Double
    If LCase(Trim(w)) = LCase(Trim(sUnit)) Then
        PkExactUnit = True
    ElseIf PackOf(n, pu, pn) Then
        PkExactUnit = (LCase(Trim(w)) = LCase(pu))
    End If
End Function

Private Function PkFirstWord(w As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(w)
        ch = Mid(w, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = " " Or ch = Chr(9) Or ch = "%" Or ch = "," Or ch = ";" Or ch = "+" Then Exit For
    Next i
    PkFirstWord = Left(w, i - 1)
End Function

' ================================================================ a quantity typed for an EI

' A quantity typed for EI n (balance b in the unit of its document), as WMS reads it:
'   0,25 / 2           — in the unit of the document;  «2 шт» — the unit of the document named;
'   1 кан / 1,5 м      — in the parts of the packing of the EI, turned into the unit of the document;
'   25%                — a percent of what came in by the EI (issues and write-offs only).
' sMode: OUT (an issue, a write-off: taken from b), IN (a return: added to b; capQty > 0 — what can come back at most:
' the parts of it close it exactly), FACT (a count: the new balance). 0 with q (in the unit of the document) and info
' («1 кан = 0,25 шт»; "" for a plain number of an EI without a packing), or the code of ParseQtyText and msg.
Function QtyEIText(ByVal s As String, n As Long, b As Double, sUnit As String, sMode As String, capQty As Double, ByRef q As Double, _
    ByRef info As String, ByRef msg As String) As Integer
    Dim sNum As String, sWord As String, st As Integer, x As Double, pu As String, pn As Double, hasPack As Boolean
    q = 0
    info = ""
    gPkKind = ""
    gPkPlain = False
    ' 1.0 (Р12): «всё» — the whole balance in one record (no tail of 0,001 left by typing the number)
    If sMode = "OUT" And IsAllWord(s) Then
        If b <= 0.0000005 Then
            msg = "«" & Trim(s) & "»: остаток " & WmsIssue.EiCanon(n) & " — 0, выдавать нечего"
            QtyEIText = 2
            Exit Function
        End If
        q = b
        gPkKind = "ALL"
        info = "всё: " & QtyWithParts(n, b, sUnit)
        QtyEIText = 0
        Exit Function
    End If
    ' 1.0 (Ч9): the words of a storekeeper — «пол кан», «1/4 кан», «2 + 1 кан» (a count by places)
    If InStr(s, "+") > 0 Then
        QtyEIText = QtySum(s, n, b, sUnit, sMode, capQty, q, info, msg)
        Exit Function
    End If
    s = PkWords(s)
    PkSplit(s, sNum, sWord)
    ' a number and a word after it («1 кан», «25%»); anything else — the strict parser of a quantity with its own messages
    ' («-1»: отрицательное, «1/2»: дробь, «abc»: недопустимый символ)
    If sWord = "" Or sNum = "" Or Not PkIsWord(sWord) Then
        st = ParseQtyText(s, q, msg)
        If st = 0 Then
            info = PlainInfo(n, q, sUnit)
            gPkPlain = True
        End If
        QtyEIText = st
        Exit Function
    End If
    ' 1.0: one quantity and one word — «1 кан 2» or «1 кан и 2» is not guessed
    If PkHasTail(sWord) And Not PkExactUnit(sWord, n, sUnit) Then
        msg = "«" & Trim(s) & "»: после единицы лишнее — напишите одно количество, например «" & sNum & " " & PkFirstWord(sWord) & "»"
        QtyEIText = 2
        Exit Function
    End If
    If sMode = "FACT" And Replace(Replace(Replace(sNum, "0", ""), ",", ""), ".", "") = "" Then
        ' a count: nothing found («0 кан»)
        x = 0
    Else
        st = ParseQtyText(sNum, x, msg)
        If st <> 0 Then
            QtyEIText = st
            Exit Function
        End If
    End If
    If sWord = "%" Then
        QtyEIText = QtyPercent(x, n, sUnit, sMode, q, info, msg)
        Exit Function
    End If
    If UnitMatch(sWord, sUnit) Then
        q = x
        info = PlainInfo(n, q, sUnit)
        QtyEIText = 0
        Exit Function
    End If
    hasPack = PackOf(n, pu, pn)
    If hasPack Then
        If UnitMatch(sWord, pu) Then
            QtyEIText = QtyParts(x, pu, pn, b, sUnit, sMode, capQty, q, info, msg)
            Exit Function
        End If
    End If
    If hasPack Then
        msg = "«" & Trim(s) & "»: единица «" & sWord & "» не подходит — " & WmsIssue.EiCanon(n) & " считается в «" & sUnit & "» (по документу), " _
            & "выдаётся в «" & pu & "» (в 1 " & sUnit & " — " & PartsText(pn, pu) & ")"
    Else
        msg = "«" & Trim(s) & "»: единица «" & sWord & "» не подходит — " & WmsIssue.EiCanon(n) & " считается в «" & sUnit & "». Чтобы выдавать " _
            & "частями (канистрами, метрами), задайте фасовку: кнопка «Фасовка ЕИ»"
    End If
    QtyEIText = 2
End Function

' 1.0 (Ч9): «половина», «четверть», «пол кан», «полметра» → 0,5 / 0,25 of the part; a fraction before a word («1/4 кан»,
' «1 1/2 м») → its decimal; a bare «1/2» stays as typed (the strict parser refuses it with its own message)
Function PkWords(ByVal s As String) As String
    Dim t As String, lt As String, p As Long, a As Long, e As Long, sA As String, sB As String, rest As String, head As String, x As Double
    t = Trim(Replace(s, Chr(160), " "))
    lt = LCase(t)
    If Left(lt, 9) = "половина " Then
        t = "0,5 " & Trim(Mid(t, 10))
    ElseIf Left(lt, 9) = "четверть " Then
        t = "0,25 " & Trim(Mid(t, 10))
    ElseIf Left(lt, 4) = "пол " Or Left(lt, 4) = "пол-" Then
        t = "0,5 " & Trim(Mid(t, 5))
    ElseIf Left(lt, 3) = "пол" And Len(lt) > 3 And Not PkHasDigit(lt) Then
        t = "0,5 " & Mid(t, 4)
    End If
    p = InStr(t, "/")
    If p > 1 Then
        a = p - 1
        Do While a >= 1
            If Not PkDigits(Mid(t, a, 1)) Then Exit Do
            a = a - 1
        Loop
        e = p + 1
        Do While e <= Len(t)
            If Not PkDigits(Mid(t, e, 1)) Then Exit Do
            e = e + 1
        Loop
        sA = Mid(t, a + 1, p - a - 1)
        sB = Mid(t, p + 1, e - p - 1)
        rest = Trim(Mid(t, e))
        head = Trim(Left(t, a))
        If sA <> "" And sB <> "" And rest <> "" And (head = "" Or PkDigits(head)) Then
            If PkIsWord(rest) And CDbl(sB) > 0 Then
                x = CDbl(sA) / CDbl(sB)
                If head <> "" Then x = x + CDbl(head)
                t = WmsIssue.QtyText(x) & " " & rest
            End If
        End If
    End If
    PkWords = t
End Function

Private Function PkDigits(s As String) As Boolean
    Dim i As Long
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then Exit Function
    Next i
    PkDigits = True
End Function

Private Function PkHasDigit(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            PkHasDigit = True
            Exit Function
        End If
    Next i
End Function

' «2 + 1 кан»: a sum (the count of one EI on two places). The word of the last term is that of a term without one; the same
' word everywhere — one quantity of that word (exact in parts); different words — only in a count (FACT), in the unit of the
' document
Private Function QtySum(ByVal s As String, n As Long, b As Double, sUnit As String, sMode As String, capQty As Double, ByRef q As Double, _
    ByRef info As String, ByRef msg As String) As Integer
    Dim parts As Variant, i As Long, sN As String, sW As String, wLast As String, w0 As String, same As Boolean, tot As Double, x As Double
    Dim st As Integer, qi As Double, inf As String
    parts = Split(s, "+")
    For i = 0 To UBound(parts)
        parts(i) = PkWords(Trim(CStr(parts(i))))
        If parts(i) = "" Then
            msg = "«" & Trim(s) & "»: между «+» нет количества"
            QtySum = 2
            Exit Function
        End If
    Next i
    PkSplit(CStr(parts(UBound(parts))), sN, wLast)
    same = True
    For i = 0 To UBound(parts)
        PkSplit(CStr(parts(i)), sN, sW)
        If sN = "" Then
            msg = "«" & Trim(s) & "»: в слагаемом «" & parts(i) & "» нет числа"
            QtySum = 2
            Exit Function
        End If
        If sW = "" Then sW = wLast
        If i = 0 Then w0 = sW
        If LCase(sW) <> LCase(w0) Then same = False
        parts(i) = sN & IIf(sW = "", "", " " & sW)
    Next i
    If same Then
        For i = 0 To UBound(parts)
            PkSplit(CStr(parts(i)), sN, sW)
            If Replace(Replace(Replace(sN, "0", ""), ",", ""), ".", "") = "" Then
                x = 0
            Else
                st = ParseQtyText(sN, x, msg)
                If st <> 0 Then
                    QtySum = st
                    Exit Function
                End If
            End If
            tot = tot + x
        Next i
        st = QtyEIText(WmsIssue.QtyText(tot) & IIf(w0 = "", "", " " & w0), n, b, sUnit, sMode, capQty, q, info, msg)
        If st = 0 Then info = Trim(s) & ": " & info
        QtySum = st
        Exit Function
    End If
    If sMode <> "FACT" Then
        msg = "«" & Trim(s) & "»: разные единицы складываются только при пересчёте (инвентаризация) — напишите одно количество"
        QtySum = 2
        Exit Function
    End If
    For i = 0 To UBound(parts)
        st = QtyEIText(CStr(parts(i)), n, b, sUnit, "FACT", 0, qi, inf, msg)
        If st <> 0 Then
            QtySum = st
            Exit Function
        End If
        tot = tot + qi
    Next i
    q = WmsIssue.Round3(tot)
    info = Trim(s) & " = " & WmsIssue.QtyText(q) & " " & sUnit
    gPkKind = ""
    gPkPlain = False
    QtySum = 0
End Function

' the same for a cell: a number in a plain format — the unit of the document; a number Calc formatted as a percent
' (25% typed) — a percent; a text — QtyEIText; a date, time, fraction or formula — refused as by ParseQtyCell
Function QtyEICell(oCell As Object, n As Long, b As Double, sUnit As String, sMode As String, capQty As Double, ByRef q As Double, _
    ByRef info As String, ByRef msg As String) As Integer
    Dim ft As Long, st As Integer
    q = 0
    info = ""
    gPkKind = ""
    gPkPlain = False
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.TEXT
        QtyEICell = QtyEIText(oCell.getString(), n, b, sUnit, sMode, capQty, q, info, msg)
    Case com.sun.star.table.CellContentType.VALUE
        ft = gDoc.getNumberFormats().getByKey(oCell.NumberFormat).Type
        If (ft And 128) <> 0 And (ft And (2 Or 4 Or 32 Or 64 Or 1024)) = 0 Then
            QtyEICell = QtyPercent(WmsIssue.Round3(oCell.getValue() * 100), n, sUnit, sMode, q, info, msg)
        Else
            st = ParseQtyCell(oCell, q, msg)
            If st = 0 Then
                info = PlainInfo(n, q, sUnit)
                gPkPlain = True
            End If
            QtyEICell = st
        End If
    Case Else
        QtyEICell = ParseQtyCell(oCell, q, msg)
    End Select
End Function

' a quantity of a field (sField: «количество возврата (F)») for EI n — from a cell, or from the text or number of a dialog
' (oCell Nothing) — with the name of the field in the message and the threshold MAX_QTY. True when it can be used
Function QtyAsEI(oCell As Object, v As Variant, sField As String, n As Long, b As Double, sUnit As String, sMode As String, capQty As Double, _
    ByRef q As Double, ByRef info As String, ByRef msg As String) As Boolean
    Dim st As Integer
    gPkKind = ""
    gPkPlain = False
    info = ""
    If Not (oCell Is Nothing) Then
        st = QtyEICell(oCell, n, b, sUnit, sMode, capQty, q, info, msg)
    ElseIf VarType(v) = 8 Then
        st = QtyEIText(CStr(v), n, b, sUnit, sMode, capQty, q, info, msg)
    Else
        QtyAsEI = WmsOrders.ParseQtyValueAs(v, sField, q, msg)
        gPkPlain = True
        Exit Function
    End If
    If st <> 0 Then
        msg = sField & ": " & msg
        Exit Function
    End If
    If q > SysNum(SK_MAX_QTY) Then
        msg = sField & ": больше порога MAX_QTY (" & SysStr(SK_MAX_QTY) & ")"
        Exit Function
    End If
    QtyAsEI = True
End Function

' the line of a plain quantity: its parts when the EI has a packing («= 1 кан»)
Private Function PlainInfo(n As Long, q As Double, sUnit As String) As String
    Dim pu As String, pn As Double
    If PackOf(n, pu, pn) Then PlainInfo = WmsIssue.QtyText(q) & " " & sUnit & " = " & PartsText(PartsOf(q, pn), pu)
End Function

' x parts of the packing (pu, pn) in the unit of the document; the balance b and its parts keep step: what is left is
' always a whole representation of the parts left, the last part takes the balance to 0 exactly
Private Function QtyParts(x As Double, pu As String, pn As Double, b As Double, sUnit As String, sMode As String, capQty As Double, _
    ByRef q As Double, ByRef info As String, ByRef msg As String) As Integer
    Dim dec As Integer, bp As Double, rest As Double, capP As Double
    dec = PackDec(pn)
    If Abs(PkRound(x, dec) - x) > 0.0000001 Then
        msg = "«" & WmsIssue.QtyText(x) & " " & pu & "»: точнее " & IIf(dec = 0, "целых", IIf(dec = 1, "0,1", "0,01")) & " " & pu & " WMS не различает " _
            & "(в 1 " & sUnit & " — " & PartsText(pn, pu) & ")"
        QtyParts = 4
        Exit Function
    End If
    bp = PartsOf(b, pn)
    Select Case sMode
    Case "OUT"
        If x > bp + 0.0000001 Then
            msg = "недостаточно: остаток " & PartsText(bp, pu) & " (" & WmsIssue.QtyText(b) & " " & sUnit & "), запрошено " & PartsText(x, pu)
            QtyParts = 5
            Exit Function
        End If
        rest = PkRound(bp - x, dec)
        q = WmsIssue.Round3(b - WmsIssue.Round3(rest / pn))
    Case "IN"
        q = WmsIssue.Round3(WmsIssue.Round3((bp + x) / pn) - b)
        If capQty > 0 Then
            capP = PartsOf(capQty, pn)
            If Abs(x - capP) < 0.0000001 Then q = capQty
        End If
    Case Else
        q = WmsIssue.Round3(x / pn)
    End Select
    If q <= 0 And sMode <> "FACT" Then
        msg = PartsText(x, pu) & " — меньше 0,001 " & sUnit & ": так мало WMS не учитывает"
        QtyParts = 5
        Exit Function
    End If
    info = PartsText(x, pu) & " = " & WmsIssue.QtyText(q) & " " & sUnit
    gPkKind = "PARTS"
    gPkParts = x
    gPkUnit = pu
    gPkN = pn
    QtyParts = 0
End Function

' pct percent of what came in by EI n (v0.1 §21: «пришло 20 кг, выдали 25% → 5 кг»)
Private Function QtyPercent(pct As Double, n As Long, sUnit As String, sMode As String, ByRef q As Double, ByRef info As String, ByRef msg As String) As Integer
    Dim qIn As Double, pu As String, pn As Double
    If sMode <> "OUT" Then
        msg = "в процентах вводятся только выдача и списание — здесь укажите количество"
        QtyPercent = 2
        Exit Function
    End If
    If pct <= 0 Or pct > 100 Then
        msg = WmsIssue.QtyText(pct) & "% — процент должен быть больше 0 и не больше 100"
        QtyPercent = 5
        Exit Function
    End If
    qIn = EiReceived(n)
    If qIn <= 0 Then
        msg = "у " & WmsIssue.EiCanon(n) & " в книге нет прихода — проценты считать не от чего; укажите количество"
        QtyPercent = 2
        Exit Function
    End If
    q = WmsIssue.Round3(qIn * pct / 100)
    If q <= 0 Then
        msg = WmsIssue.QtyText(pct) & "% от " & WmsIssue.QtyText(qIn) & " " & sUnit & " — меньше 0,001 " & sUnit
        QtyPercent = 5
        Exit Function
    End If
    info = WmsIssue.QtyText(pct) & "% от пришедших " & WmsIssue.QtyText(qIn) & " " & sUnit & " = " & WmsIssue.QtyText(q) & " " & sUnit
    If pct < 1 Then info = info & " — проверьте: это меньше 1% (25% пишется как 25)"
    If PackOf(n, pu, pn) Then info = info & " (" & PartsText(PartsOf(q, pn), pu) & ")"
    gPkKind = "PCT"
    gPkPct = pct
    gPkBase = qIn
    QtyPercent = 0
End Function

' the fields of the journal line of the operation being planned: how its quantity was given
' «всё», «весь остаток» (Р12)
Function IsAllWord(ByVal s As String) As Boolean
    s = Replace(LCase(Trim(s)), "ё", "е")
    IsAllWord = (s = "все" Or s = "весь" Or s = "весь остаток" Or s = "все целиком" Or s = "остаток")
End Function

Sub PlanQtyFields()
    If gPkKind = "ALL" Then
        PlanField("HOW", "ALL")
    ElseIf gPkKind = "PARTS" Then
        PlanField("PARTS", gPkParts)
        PlanField("PART_UNIT", gPkUnit)
        PlanField("PART_N", gPkN)
    ElseIf gPkKind = "PCT" Then
        PlanField("PCT", gPkPct)
        PlanField("PCT_BASE", gPkBase)
    End If
End Sub

' 1.0 (Ч1): the column «Остаток частями» of «Наличие» (K, after the ten columns of the registry): «3 кан» — the balance E
' times the parts of the packing of the same row of _PACK (A the EI, B the part, C parts in 1 unit), "" without a packing
Private Const SC_PARTS = 10

Function PartsFormula(n As Long) As String
    Dim r As String, v As String, pk As String
    r = CStr(n + 1)
    pk = "$'" & SH_PACK & "'."
    v = "ROUND($E" & r & "*" & pk & "$C" & r & ";2)"
    PartsFormula = "=IF(AND($A" & r & "<>"""";" & pk & "$A" & r & "=$A" & r & ";N(" & pk & "$C" & r & ")>0);IF(" & v & "=INT(" & v & ");FIXED(" _
        & v & ";0;1);IF(" & v & "*10=INT(" & v & "*10);FIXED(" & v & ";1;1);FIXED(" & v & ";2;1)))&"" ""&" & pk & "$B" & r & ";"""")"
End Function

' 1.0 (Ч4): how the quantity was given, for the status of a posted row of «Выдачи» («Проведено · 1 кан») — written by the
' operation itself, so the replay of the journal writes the same text
Function PostedSuffix() As String
    Select Case gPkKind
    Case "PARTS"
        PostedSuffix = " · " & PartsText(gPkParts, gPkUnit)
    Case "PCT"
        PostedSuffix = " · " & WmsIssue.QtyText(gPkPct) & " %"
    Case "ALL"
        PostedSuffix = " · весь остаток"
    End Select
End Function

' ================================================================ 1.0: the two readings WMS never guesses

' After a quantity of EI n was read (QtyEIText / QtyEICell / QtyAsEI: q, info and gPk*), before it is used:
'   a bare whole number for an EI with a packing — «1»: 1 шт (4 кан) or 1 кан;
'   a percent below 1 — «0,25%»: 0,25 % or 25 %.
' sVerb names the action on the buttons («Выдать», «Списать», «Вернуть», «Факт»); sMode, b, capQty — as for QtyEIText.
' bFits False: the quantity as read does not fit (more than the balance, msg holds why) — only the parts may.
' 0 — use q (the other reading when chosen: info and the journal fields follow it), 1 — not posted by the storekeeper
' (msg), 2 — nothing fits (msg).
Function QtyAsk(n As Long, sName As String, sUnit As String, sVerb As String, sMode As String, b As Double, capQty As Double, bFits As Boolean, _
    ByRef q As Double, ByRef info As String, ByRef msg As String) As Integer
    Dim pu As String, pn As Double, x As Double, q2 As Double, info2 As String, msg2 As String, k As Integer, sText As String, sHead As String
    Dim labels As Variant, codes As Variant, alt As String, pct As Double, pctBase As Double
    sHead = WmsIssue.EiCanon(n) & IIf(sName <> "", " «" & sName & "»", "")
    If gPkKind = "PCT" And bFits Then
        If gPkPct >= 1 Then Exit Function
        pct = gPkPct
        pctBase = gPkBase
        sText = sHead & ": " & WmsIssue.QtyText(pct) & "% от пришедших " & WmsIssue.QtyText(pctBase) & " " & sUnit & " = " & WmsIssue.QtyText(q) & " " _
            & sUnit & " — это меньше одного процента." & Chr(10) & "Если имелось в виду " & WmsIssue.QtyText(WmsIssue.Round3(pct * 100)) _
            & "%, выберите вторую кнопку."
        labels = Array(sVerb & " " & WmsIssue.QtyText(pct) & "% — это " & WmsIssue.QtyText(q) & " " & sUnit, "", "Не проводить")
        codes = Array(0, 1, 2)
        If WmsIssue.Round3(pct * 100) <= 100 Then
            alt = WmsIssue.QtyText(WmsIssue.Round3(pct * 100)) & "%"
            If QtyEIText(alt, n, b, sUnit, sMode, capQty, q2, info2, msg2) = 0 And q2 <= b + 0.0000001 Then
                labels(1) = sVerb & " " & alt & " — это " & WmsIssue.QtyText(q2) & " " & sUnit
            End If
        End If
        If labels(1) = "" Then
            labels = Array(labels(0), "Не проводить")
            codes = Array(0, 2)
        End If
        ' QtyEIText of the other reading changed gPk*: the reading as typed is restored unless the other one is chosen
        gPkKind = "PCT"
        gPkPct = pct
        gPkBase = pctBase
        k = WmsPrimeUi.UiButtons("WMS — проверьте процент", sText, labels, codes)
        If k = 1 Then
            QtyAsk = QtyAskOther(alt, n, b, sUnit, sMode, capQty, q, info, msg)
            Exit Function
        End If
        If k <> 0 Then
            msg = "проверьте процент «" & WmsIssue.QtyText(pct) & "%» (может быть, имелось в виду " & WmsIssue.QtyText(WmsIssue.Round3(pct * 100)) & "%)"
            QtyAsk = 1
        End If
        Exit Function
    End If
    If Not gPkPlain Or gPkKind <> "" Then
        If Not bFits Then QtyAsk = 2
        Exit Function
    End If
    x = q
    If x < 1 Or Abs(x - Int(x)) > 0.0000001 Or Not PackOf(n, pu, pn) Then
        If Not bFits Then QtyAsk = 2
        Exit Function
    End If
    alt = WmsIssue.QtyText(x) & " " & pu
    If QtyEIText(alt, n, b, sUnit, sMode, capQty, q2, info2, msg2) <> 0 Then
        ' the parts do not fit either: as typed (or the refusal of it)
        gPkKind = ""
        gPkPlain = True
        If Not bFits Then QtyAsk = 2
        Exit Function
    End If
    gPkKind = ""
    gPkPlain = True
    sText = sHead & ": в 1 " & sUnit & " — " & PartsText(pn, pu) & ". " & IIf(sMode = "FACT", "Учётный остаток", "Остаток") & " " _
        & QtyWithParts(n, b, sUnit) & "." & Chr(10) & "Написано «" & WmsIssue.QtyText(x) & "» — " & IIf(sMode = "FACT", "сколько на самом деле?", "это штуки документа или части?")
    If bFits Then
        labels = Array(sVerb & " " & WmsIssue.QtyText(x) & " " & sUnit & " — это " & PartsText(PartsOf(x, pn), pu) & AskRest(n, b, x, sUnit, sMode), _
            sVerb & " " & alt & " — это " & WmsIssue.QtyText(q2) & " " & sUnit & AskRest(n, b, q2, sUnit, sMode), "Не проводить")
        codes = Array(0, 1, 2)
    Else
        sText = sText & Chr(10) & WmsIssue.QtyText(x) & " " & sUnit & " нельзя: " & msg
        labels = Array(sVerb & " " & alt & " — это " & WmsIssue.QtyText(q2) & " " & sUnit & AskRest(n, b, q2, sUnit, sMode), "Не проводить")
        codes = Array(1, 2)
    End If
    k = WmsPrimeUi.UiButtons("WMS — штуки или части?", sText, labels, codes)
    If k = 1 Then
        QtyAsk = QtyAskOther(alt, n, b, sUnit, sMode, capQty, q, info, msg)
    ElseIf k = 0 Then
        QtyAsk = 0
    Else
        msg = "уточните количество: «" & WmsIssue.QtyText(x) & " " & sUnit & "» или «" & alt & "»"
        QtyAsk = 1
    End If
End Function

' the other reading chosen: read again so that q, info and gPk* (the journal fields) are its own
Private Function QtyAskOther(alt As String, n As Long, b As Double, sUnit As String, sMode As String, capQty As Double, ByRef q As Double, _
    ByRef info As String, ByRef msg As String) As Integer
    If QtyEIText(alt, n, b, sUnit, sMode, capQty, q, info, msg) <> 0 Then
        QtyAskOther = 2
    ElseIf sMode = "OUT" And q > b + 0.0000001 Then
        msg = "недостаточно: остаток " & QtyWithParts(n, b, sUnit) & ", запрошено " & WmsIssue.QtyText(q) & " " & sUnit
        QtyAskOther = 2
    End If
End Function

' «, останется 3 кан» after an issue or a write-off of q; «, станет …» after a return; "" for a count
Private Function AskRest(n As Long, b As Double, q As Double, sUnit As String, sMode As String) As String
    If sMode = "OUT" Then
        AskRest = ", останется " & QtyWithParts(n, WmsIssue.Round3(b - q), sUnit)
    ElseIf sMode = "IN" Then
        AskRest = ", станет " & QtyWithParts(n, WmsIssue.Round3(b + q), sUnit)
    End If
End Function

' the name of a part of EI n («кан»), "" without a packing
Function PartUnitOf(n As Long) As String
    Dim pu As String, pn As Double
    If PackOf(n, pu, pn) Then PartUnitOf = pu
End Function

' 1.0: a quantity of EI n for a field of a window («Исправить»): in the parts when it is a whole count of them («1 кан»),
' otherwise as a number of the unit of the document («0,25»)
Function QtyForInput(n As Long, q As Double) As String
    Dim pu As String, pn As Double, x As Double
    QtyForInput = WmsIssue.QtyText(q)
    If Not PackOf(n, pu, pn) Then Exit Function
    x = PartsOf(q, pn)
    If Abs(WmsIssue.Round3(x / pn) - q) < 0.0000001 Then QtyForInput = PartsText(x, pu)
End Function

' what is left after an issue of q from b: «0,75 шт (3 кан)» (1.0: «Контроль» shows it before the posting)
Function RestText(n As Long, b As Double, q As Double, sUnit As String) As String
    RestText = QtyWithParts(n, WmsIssue.Round3(b - q), sUnit)
End Function

' what came in by EI n in the unit of its document (the base of a percent): its receipt of «Заказы» (_RCV, a transferred
' receipt too), the lines of «Иной приход» posted for it (a part replenished — all of them), a migrated balance (_LEG);
' 0 when the book holds none
Function EiReceived(n As Long) As Double
    Dim x As Double, rv As Variant, lg As Variant, last As Long, s As String
    rv = WmsOrders.RcvSheet().getCellRangeByPosition(0, n, RV_QTY, n).getDataArray()(0)
    If CStr(rv(RV_EI)) <> "" And CStr(rv(RV_STATE)) = RV_LIVE And VarType(rv(RV_QTY)) = 5 Then x = x + rv(RV_QTY)
    If gDoc.Sheets.hasByName(SH_LEG) Then
        lg = WmsLegacy.LegSheet().getCellRangeByPosition(0, n, LG_BAL0, n).getDataArray()(0)
        If CStr(lg(LG_KIND)) = LG_MIGRATE And VarType(lg(LG_BAL0)) = 5 Then x = x + lg(LG_BAL0)
    End If
    If gDoc.Sheets.hasByName(SH_SPECIAL) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_SPECIAL))
        If last >= 1 Then
            s = WmsView.Eng("SUMPRODUCT(" & WmsView.PostedExpr(WmsView.ColRng(SH_SPECIAL, XC_STATUS, 1, last)) & "*(" & WmsView.ColRng(SH_SPECIAL, XC_EI, 1, last) _
                & "=""" & WmsIssue.EiCanon(n) & """)*ISNUMBER(" & WmsView.ColRng(SH_SPECIAL, XC_QTY, 1, last) & ")*" & WmsView.ColRng(SH_SPECIAL, XC_QTY, 1, last) & ")")
            ' a number of the engine as text in the language of the book (2,5): read with the point
            If s <> "" Then x = x + Val(Replace(s, ",", "."))
        End If
    End If
    EiReceived = WmsIssue.Round3(x)
End Function

' ================================================================ «Фасовка ЕИ» — the operation EI_PACK

' Sets the packing of EI sEI: sN parts called sPart in one unit of its document; sN empty or 0 — no packing. Changing it
' changes no balance: the balances and the movements stay in the unit of the document. OK:<seq> / SKIP / ERR / BLOCKED
Function PackSet(sEI As String, sN As String, sPart As String, sNote As String) As String
    Dim n As Long, canon As String, msg As String, p As String, d As Variant, pn As Double, pu As String, oldU As String, oldN As Double, st As Integer
    Dim why As String, sh As Object
    WmsInit()
    On Error GoTo EH
    sh = PackSheet()
    If IsNull(sh) Then
        PackSet = "ERR:в книге нет листа фасовки " & SH_PACK & " — книга не обновлена до 0.8 (обновите её мастером WMS_UPGRADE)"
        Exit Function
    End If
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        PackSet = "ERR:" & msg
        Exit Function
    End If
    p = WmsIssue.StockRowProblem(n, canon)
    If p <> "" Then
        PackSet = "ERR:" & p
        Exit Function
    End If
    d = WmsIssue.StockData(n)
    If CStr(d(SC_STATE)) = EI_ST_STORNO Then
        PackSet = "ERR:" & canon & " — приход удалён (сторно): фасовка не нужна"
        Exit Function
    End If
    pu = Trim(sPart)
    If Trim(sN) = "" Or Trim(sN) = "0" Then
        pn = 0
        pu = ""
    Else
        st = ParseQtyText(Trim(sN), pn, msg)
        If st = 4 And PkDecimal(Trim(sN)) Then
            ' 1.0 (Ч8): parts computed from a weight («1 м = 7,85 кг» → 0,127389 м in 1 кг) keep six decimals
            pn = Val(Replace(Trim(sN), ",", "."))
            st = 0
        End If
        If st <> 0 Then
            PackSet = "ERR:сколько частей в 1 " & d(SC_UNIT) & ": " & msg
            Exit Function
        End If
        If pn > PACK_MAX_N Then
            PackSet = "ERR:" & WmsIssue.QtyText(pn) & " частей в 1 " & d(SC_UNIT) & " — больше " & PACK_MAX_N & ". Возьмите части крупнее (профиль — в метрах, не в миллиметрах)"
            Exit Function
        End If
        If pu = "" Then
            PackSet = "ERR:укажите, как называть часть (кан, бут, м, кг, л)"
            Exit Function
        End If
        If Len(pu) > 12 Or pu <> Replace(Replace(Replace(pu, ";", ""), "|", ""), "=", "") Then
            PackSet = "ERR:название части «" & pu & "» — до 12 знаков, без ; | ="
            Exit Function
        End If
        For st = 1 To Len(pu)
            If Mid(pu, st, 1) >= "0" And Mid(pu, st, 1) <= "9" Then
                PackSet = "ERR:в названии части «" & pu & "» не должно быть цифр — число частей укажите отдельно"
                Exit Function
            End If
        Next st
        If UnitMatch(pu, CStr(d(SC_UNIT))) And Abs(pn - 1) > 0.0000001 Then
            PackSet = "ERR:часть «" & pu & "» — та же единица, что и по документу («" & d(SC_UNIT) & "»): в 1 " & d(SC_UNIT) & " не может быть " & WmsIssue.QtyText(pn) & " " & pu
            Exit Function
        End If
    End If
    PackOf(n, oldU, oldN)
    If oldU = pu And Abs(oldN - pn) < 0.0000001 And Trim(sNote) = Trim(CStr(sh.getCellByPosition(PK_NOTE, n).getString())) Then
        PackSet = "SKIP:фасовка " & canon & " не изменилась"
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        PackSet = why
        Exit Function
    End If
    PlanBegin("EI_PACK", canon)
    PlanField("EI", canon)
    PlanField("UNIT", pu)
    PlanField("N", pn)
    PlanField("DOC_UNIT", CStr(d(SC_UNIT)))
    PlanField("OLD_UNIT", oldU)
    PlanField("OLD_N", oldN)
    If Trim(sNote) <> "" Then PlanField("NOTE", Trim(sNote))
    If pn > 0 Then
        PlanSetValue(SH_PACK, n, PK_EI, canon, False)
        PlanSetValue(SH_PACK, n, PK_UNIT, pu, False)
        PlanSetValue(SH_PACK, n, PK_N, pn, False)
        PlanSetValue(SH_PACK, n, PK_NOTE, Trim(sNote), False)
        ' 1.0 (Ч1): «Остаток частями» of «Наличие» — a formula of the row (it follows the balance by itself)
        If gDoc.Sheets.getByName(SH_STOCK).getCellByPosition(SC_PARTS, n).getType() <> com.sun.star.table.CellContentType.FORMULA Then
            PlanFormula(SH_STOCK, n, SC_PARTS, PartsFormula(n))
        End If
    Else
        PlanSetValue(SH_PACK, n, PK_EI, "", False)
        PlanSetValue(SH_PACK, n, PK_UNIT, "", False)
        PlanSetValue(SH_PACK, n, PK_N, "", False)
        PlanSetValue(SH_PACK, n, PK_NOTE, "", False)
    End If
    PackSet = ApplyOperation(0, 0)
    Exit Function
EH:
    PackSet = "ERR:фасовка не записана: " & Error$ & " (строка " & Erl & ")"
End Function

' a guess of the packing from the name of the EI: «Незамерзайка 4 канистры по 5 л» → 4 кан, «Профиль 40х20 6 м» → 6 м,
' «… L=6000 мм» → 6 м. pn = 0 when the name says nothing
' 1.0: a comma between two digits is a decimal comma («2,5 м», «11,7 м»): it becomes a dot, so that the separators that
' follow do not cut the number in two (it was read as «5 м», «7 м»)
Private Function PkDecimalCommas(s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "," And i > 1 And i < Len(s) Then
            If Mid(s, i - 1, 1) >= "0" And Mid(s, i - 1, 1) <= "9" And Mid(s, i + 1, 1) >= "0" And Mid(s, i + 1, 1) <= "9" Then ch = "."
        End If
        r = r & ch
    Next i
    PkDecimalCommas = r
End Function

Sub PackGuess(sName As String, ByRef pn As Double, ByRef pu As String)
    Dim s As String, t As Variant, i As Integer, w As Variant, x As Double, msg As String, nm As Variant, k As Integer, num As String, rest As String
    pn = 0
    pu = ""
    s = PkDecimalCommas(LCase(sName))
    For Each w In Array("(", ")", ",", ";", "/", "\", ":", "=", "+", "×", "*")
        s = Replace(s, w, " ")
    Next w
    s = Replace(s, "l ", " ")
    t = Split(Trim(s), " ")
    nm = Array("канистр|кан", "кан|кан", "бутыл|бут", "бут|бут", "банк|бан", "упаков|уп", "уп|уп", "рулон|рул", "рул|рул", "пачк|пач", "мешк|меш", "мешок|меш", _
        "ведр|вед", "ведер|вед", "короб|кор", "флакон|фл", "баллон|бал", "тюбик|тюб", "пар|пар")
    For i = 0 To UBound(t)
        num = ""
        rest = ""
        PkSplit(t(i), num, rest)
        If num <> "" Then
            If rest = "" And i < UBound(t) Then rest = t(i + 1)
            If ParseQtyText(Replace(num, ".", ","), x, msg) = 0 Then
                If rest = "м" Or rest = "м." Or Left(rest, 4) = "метр" Then
                    pn = x
                    pu = "м"
                ElseIf rest = "мм" And x >= 100 Then
                    pn = WmsIssue.Round3(x / 1000)
                    pu = "м"
                Else
                    For k = 0 To UBound(nm)
                        If Left(rest, Len(Split(nm(k), "|")(0))) = Split(nm(k), "|")(0) Then
                            pn = x
                            pu = Split(nm(k), "|")(1)
                            Exit For
                        End If
                    Next k
                End If
                If pn > 0 And pn <= PACK_MAX_N Then Exit Sub
                pn = 0
                pu = ""
            End If
        End If
    Next i
End Sub

' «Фасовка ЕИ»: the EI of the current row («Наличие», «Выдачи», «Возврат», «Корректировки», the card of «Поиск») or asked
Sub BtnPackEI(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PackEI()
    If Left(s, 4) = "ERR:" Or Left(s, 8) = "BLOCKED:" Or Left(s, 3) = "ERR" Then
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    ElseIf Left(s, 3) = "OK:" Then
        WmsUi.UiMessage(gUiLastMsg)
    End If
End Sub

' 1.0 (Ч6): «Фасовка пачкой…» — the EIs with a balance and without a packing whose name tells one («4 канистры», «L=6000»),
' each shown with its guess and written only when the storekeeper presses «Записать»
Sub BtnPackBatch(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PackBatch()
    If Left(s, 4) = "ERR:" Then
        WmsUi.UiMessage(Mid(s, 5))
    ElseIf Left(s, 3) = "OK=" Then
        WmsUi.UiMessage(gUiLastMsg)
    End If
End Sub

Function PackBatch() As String
    Dim st As Object, pk As Object, last As Long, d As Variant, p As Variant, i As Long, n As Long, gn As Double, gu As String, nC As Long
    Dim cand() As Long, cn() As Double, cu() As String, k As Integer, res As String, nOk As Long, nSkip As Long, nErr As Long, sErr As String, x As String
    st = WmsIssue.StockSheet()
    pk = PackSheet()
    If IsNull(pk) Then
        PackBatch = "ERR:в книге нет листа фасовки " & SH_PACK & " — книга не обновлена"
        Exit Function
    End If
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last < 1 Then
        PackBatch = "ERR:в «" & SH_STOCK & "» нет ЕИ"
        Exit Function
    End If
    d = st.getCellRangeByPosition(0, 1, SC_STATE, last).getDataArray()
    p = pk.getCellRangeByPosition(0, 1, PK_N, last).getDataArray()
    ReDim cand(last)
    ReDim cn(last)
    ReDim cu(last)
    For i = 0 To last - 1
        If VarType(d(i)(SC_QTY)) = 5 And CStr(d(i)(SC_STATE)) <> EI_ST_STORNO And CStr(d(i)(SC_EI)) <> "" Then
            If d(i)(SC_QTY) > 0.0000005 And Not (CStr(p(i)(PK_EI)) = CStr(d(i)(SC_EI)) And VarType(p(i)(PK_N)) = 5) Then
                PackGuess(CStr(d(i)(SC_NAME)), gn, gu)
                If gn > 1 And Not UnitMatch(gu, CStr(d(i)(SC_UNIT))) Then
                    cand(nC) = i + 1
                    cn(nC) = gn
                    cu(nC) = gu
                    nC = nC + 1
                End If
            End If
        End If
    Next i
    If nC = 0 Then
        PackBatch = "ERR:ЕИ с остатком, у которых наименование подсказывает фасовку и она ещё не задана, нет"
        Exit Function
    End If
    For i = 0 To nC - 1
        n = cand(i)
        x = "в 1 " & d(n - 1)(SC_UNIT) & " — " & PartsText(cn(i), cu(i))
        k = WmsPrimeUi.UiButtons("Фасовка пачкой — " & (i + 1) & " из " & nC, WmsIssue.EiCanon(n) & " · «" & d(n - 1)(SC_NAME) & "» · остаток " _
            & WmsIssue.QtyText(CDbl(d(n - 1)(SC_QTY))) & " " & d(n - 1)(SC_UNIT) & Chr(10) & Chr(10) & "По наименованию: «" & x & "» (остаток — " _
            & PartsText(WmsIssue.Round3(CDbl(d(n - 1)(SC_QTY)) * cn(i)), cu(i)) & ")." & Chr(10) & "Записать эту фасовку?", _
            Array("Записать «" & x & "»", "Пропустить", "Остановить"), Array(1, 0, 2))
        If k = 2 Or k < 0 Then Exit For
        If k = 1 Then
            res = PackSet(WmsIssue.EiCanon(n), PkNumText6(cn(i)), cu(i), "по наименованию (фасовка пачкой)")
            If Left(res, 3) = "OK:" Then
                nOk = nOk + 1
            Else
                nErr = nErr + 1
                sErr = sErr & Chr(10) & WmsIssue.EiCanon(n) & ": " & Mid(res, InStr(res, ":") + 1)
                If Left(res, 4) <> "ERR:" And Left(res, 5) <> "SKIP:" Then Exit For
            End If
        Else
            nSkip = nSkip + 1
        End If
    Next i
    PackBatch = "OK=" & nOk & ";SKIP=" & nSkip & ";ERR=" & nErr
    gUiLastMsg = "Фасовка пачкой: записано " & nOk & ", пропущено " & nSkip & IIf(nErr > 0, ", не записано " & nErr & ":" & sErr, "") _
        & IIf(nOk + nSkip + nErr < nC, " (остановлено: просмотрено " & (nOk + nSkip + nErr) & " из " & nC & ")", "")
End Function

' 1.0 (Ч2): the packing of the last EI of the same product (the same article — or the same name when there is no article —
' and the same unit of the document) is offered to a new EI of a first receipt. "" to post (pn > 0 — with that packing),
' SKIP:… when the receipt is not posted
Function ProductPackAsk(sName As String, sArt As String, sUnit As String, ByRef pn As Double, ByRef pu As String, ByRef sFrom As String) As String
    Dim sh As Object, cur As Object, last As Long, f As String, s As String, a As Variant, n As Long, k As Integer, x As String
    sFrom = ""
    If Trim(sArt) = "" And Trim(sName) = "" Then Exit Function
    sh = PackSheet()
    If Not IsNull(sh) Then
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        last = cur.getRangeAddress().EndRow
    End If
    If last >= 1 Then
        If Trim(sArt) <> "" Then
            f = "(" & WmsView.ColRng(SH_STOCK, SC_ART, 1, last) & "=""" & Replace(Trim(sArt), """", """""") & """)"
        Else
            f = "(" & WmsView.ColRng(SH_STOCK, SC_NAME, 1, last) & "=""" & Replace(Trim(sName), """", """""") & """)"
        End If
        s = WmsView.Eng("TEXTJOIN("";"";1;IF(ISNUMBER(" & WmsView.ColRng(SH_PACK, PK_N, 1, last) & ")*(" & WmsView.ColRng(SH_PACK, PK_N, 1, last) & ">0)*" & f _
            & "*(" & WmsView.ColRng(SH_STOCK, SC_UNIT, 1, last) & "=""" & Replace(Trim(sUnit), """", """""") & """);ROW(" & WmsView.ColRng(SH_PACK, PK_N, 1, last) _
            & ")-1;""""))")
    End If
    ' 1.0 (Ф3): no earlier EI of the product has a packing — the packing written for the goods in «Номенклатура»
    If s = "" Then
        ProductPackAsk = PkGoodsAsk(sName, sArt, sUnit, pn, pu, sFrom)
        Exit Function
    End If
    a = Split(s, ";")
    n = CLng(Val(a(UBound(a))))
    If Not PackOf(n, pu, pn) Then
        pn = 0
        pu = ""
        Exit Function
    End If
    x = "в 1 " & IIf(Trim(sUnit) <> "", Trim(sUnit), "шт") & " — " & PartsText(pn, pu)
    k = WmsPrimeUi.UiButtons("WMS — фасовка нового ЕИ", "«" & sName & "»: у прошлой поставки этого товара (" & WmsIssue.EiCanon(n) & ") фасовка «" & x _
        & "»." & Chr(10) & "Задать такую же новому ЕИ — чтобы выдавать по 1 " & pu & " сразу?", Array("Да, фасовка «" & x & "»", "Без фасовки"), Array(1, 0))
    If k <> 1 Then
        pn = 0
        pu = ""
        Exit Function
    End If
    sFrom = WmsIssue.EiCanon(n)
End Function

' 1.0 (Ф3): the packing of «Номенклатура» (F parts in 1 unit, G the part) of the same article — or of the same name when one of
' them has no article — and of the same unit (or no unit there) offered to the new EI; sFrom = «Номенклатуре» when taken
Private Function PkGoodsAsk(sName As String, sArt As String, sUnit As String, ByRef pn As Double, ByRef pu As String, ByRef sFrom As String) As String
    Dim gs As Object, last As Long, a As Variant, nItems As Long, i As Integer, d As Variant, r As Long, x As String, k As Integer, cond As String
    Dim F As String, G As String, A0 As String, N0 As String, sA As String, sN As String
    If Not gDoc.Sheets.hasByName(WmsSupply.SH_GOODS) Then Exit Function
    gs = gDoc.Sheets.getByName(WmsSupply.SH_GOODS)
    last = WmsOrders.LastRow(gs)
    If last < 1 Then Exit Function
    F = WmsView.ColRng(WmsSupply.SH_GOODS, GD_PACKN, 1, last)
    G = WmsView.ColRng(WmsSupply.SH_GOODS, GD_PACKU, 1, last)
    A0 = "LOWER(TRIM(" & WmsView.ColRng(WmsSupply.SH_GOODS, GD_ART, 1, last) & "))"
    N0 = "LOWER(TRIM(" & WmsView.ColRng(WmsSupply.SH_GOODS, GD_NAME, 1, last) & "))"
    sA = """" & Replace(LCase(Trim(sArt)), """", """""") & """"
    sN = """" & Replace(LCase(Trim(sName)), """", """""") & """"
    If Trim(sArt) <> "" Then
        cond = "((" & A0 & "=" & sA & ")+(" & A0 & "="""")*(" & N0 & "=" & sN & ")>0)"
    Else
        cond = "(" & N0 & "=" & sN & ")"
    End If
    a = WmsView.EngList("TEXTJOIN("";"";1;IF(ISNUMBER(" & F & ")*(" & F & ">0)*(TRIM(" & G & ")<>"""")*" & cond & ";ROW(" & F & ")-1;""""))", nItems)
    For i = 0 To nItems - 1
        r = CLng(a(i))
        d = gs.getCellRangeByPosition(0, r, GD_PACKU, r).getDataArray()(0)
        If Trim(CStr(d(GD_UNIT))) = "" Or UnitMatch(CStr(d(GD_UNIT)), sUnit) Then
            pn = d(GD_PACKN)
            pu = Trim(CStr(d(GD_PACKU)))
            Exit For
        End If
    Next i
    If pn <= 0 Or pu = "" Then
        pn = 0
        pu = ""
        Exit Function
    End If
    x = "в 1 " & IIf(Trim(sUnit) <> "", Trim(sUnit), "шт") & " — " & PartsText(pn, pu)
    k = WmsPrimeUi.UiButtons("WMS — фасовка нового ЕИ", "«" & sName & "»: в «" & WmsSupply.SH_GOODS & "» у этого товара фасовка «" & x & "»." & Chr(10) _
        & "Задать её новому ЕИ — чтобы выдавать по 1 " & pu & " сразу?", Array("Да, фасовка «" & x & "»", "Без фасовки"), Array(1, 0))
    If k <> 1 Then
        pn = 0
        pu = ""
        Exit Function
    End If
    sFrom = "«Номенклатуре»"
End Function

' 1.0 (Ч8): a unit of weight in kilograms (1 for кг, 1000 for т, 0,001 for г); 0 — not a weight
Function PkWeightUnit(sUnit As String) As Double
    If UnitMatch(sUnit, "кг") Then
        PkWeightUnit = 1
    ElseIf UnitMatch(sUnit, "т") Then
        PkWeightUnit = 1000
    ElseIf UnitMatch(sUnit, "г") Then
        PkWeightUnit = 0.001
    End If
End Function

' «1 м = 2,5 кг» (or «2,5 кг = 1 м») for a document in sDocUnit (a weight): pn — parts (metres) in 1 unit of the document, pu — the part
Function PkWeight(ByVal s As String, sDocUnit As String, ByRef pn As Double, ByRef pu As String, ByRef msg As String) As Boolean
    Dim a As Variant, nL As String, wL As String, nR As String, wR As String, xL As Double, xR As Double, t As String, kd As Double, kw As Double
    a = Split(PkDecimalCommas(Trim(s)), "=")
    If UBound(a) <> 1 Then
        msg = "«" & Trim(s) & "»: напишите вес одного метра так: «1 м = 2,5 кг»"
        Exit Function
    End If
    PkSplit(Replace(Trim(CStr(a(0))), ".", ","), nL, wL)
    PkSplit(Replace(Trim(CStr(a(1))), ".", ","), nR, wR)
    If PkWeightUnit(wL) > 0 And PkWeightUnit(wR) = 0 Then
        t = nL: nL = nR: nR = t
        t = wL: wL = wR: wR = t
    End If
    kd = PkWeightUnit(sDocUnit)
    kw = PkWeightUnit(wR)
    If kd = 0 Then
        msg = "вес метра нужен, когда документ в кг или т, а у этого ЕИ — «" & sDocUnit & "»: впишите просто, сколько частей в 1 " & sDocUnit
        Exit Function
    End If
    If kw = 0 Or wL = "" Or PkWeightUnit(wL) > 0 Then
        msg = "«" & Trim(s) & "»: напишите вес одного метра так: «1 м = 2,5 кг» (слева длина, справа вес в кг, т или г)"
        Exit Function
    End If
    If nL = "" Then nL = "1"
    If ParseQtyText(nL, xL, msg) <> 0 Or ParseQtyText(nR, xR, msg) <> 0 Then
        msg = "«" & Trim(s) & "»: " & msg
        Exit Function
    End If
    pn = xL / (xR * kw / kd)
    pu = wL
    PkWeight = True
End Function

' a number with up to six decimals and a decimal comma (the parts of a packing computed from a weight)
Function PkNumText6(x As Double) As String
    Dim ip As Double, fr As Long, s As String
    ip = Int(x)
    fr = Int((x - ip) * 1000000 + 0.5)
    If fr >= 1000000 Then
        ip = ip + 1
        fr = 0
    End If
    s = Format(ip, "0")
    If fr > 0 Then
        s = s & "," & Right("000000" & fr, 6)
        Do While Right(s, 1) = "0"
            s = Left(s, Len(s) - 1)
        Loop
    End If
    PkNumText6 = s
End Function

Private Function PkDecimal(s As String) As Boolean
    Dim i As Long, ch As String, nSep As Integer
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "," Or ch = "." Then
            nSep = nSep + 1
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    PkDecimal = (nSep <= 1 And Len(s) <= 12)
End Function

' 1.0 (Ч3): what to check before a packing is changed — the balance in the new parts is not whole (for canisters, bottles,
' pieces), and how many issues of this EI were given in the parts of the old packing. "" — nothing to ask
Function PackChangeQuestion(n As Long, canon As String, d As Variant, sN As String, sPart As String) As String
    Dim pn As Double, msg As String, b As Double, x As Double, pu As String, oldN As Double, oldU As String, s As String, cnt As String, sh As Object
    Dim cur As Object, last As Long, st As Integer
    If Trim(sN) = "" Or Trim(sN) = "0" Or Trim(sPart) = "" Then Exit Function
    st = ParseQtyText(Trim(sN), pn, msg)
    If st = 4 And PkDecimal(Trim(sN)) Then
        pn = Val(Replace(Trim(sN), ",", "."))
        st = 0
    End If
    If st <> 0 Or pn <= 0 Then Exit Function
    If VarType(d(SC_QTY)) = 5 Then b = d(SC_QTY)
    pu = Trim(sPart)
    If b > 0 And PkCountable(pu) Then
        x = b * pn
        If Abs(x - Int(x + 0.5)) > 0.005 Then
            s = "Остаток " & WmsIssue.QtyText(b) & " " & d(SC_UNIT) & " = " & WmsIssue.QtyText(x) & " " & pu & " при новой фасовке — не целое число " & pu _
                & ". Проверьте: столько ли " & pu & " лежит на самом деле?"
        End If
    End If
    If PackOf(n, oldU, oldN) Then
        If oldU <> pu Or Abs(oldN - pn) > 0.0000001 Then
            sh = WmsIssue.IssuesSheet()
            cur = sh.createCursor()
            cur.gotoEndOfUsedArea(False)
            last = cur.getRangeAddress().EndRow
            If last >= 1 Then
                cnt = WmsView.Eng("SUMPRODUCT((" & WmsView.ColRng(SH_ISSUES, IC_EI, 1, last) & "=""" & canon & """)*" _
                    & WmsView.PostedExpr(WmsView.ColRng(SH_ISSUES, IC_CTL, 1, last)) & "*ISNUMBER(SEARCH(""·"";" _
                    & WmsView.ColRng(SH_ISSUES, IC_CTL, 1, last) & "))*(RIGHT(" & WmsView.ColRng(SH_ISSUES, IC_CTL, 1, last) & ";" & (Len(oldU) + 1) _
                    & ")="" " & Replace(oldU, """", """""") & """))")
            End If
            If Val(cnt) > 0 Then
                s = s & IIf(s <> "", Chr(10) & Chr(10), "") & "По старой фасовке (в 1 " & d(SC_UNIT) & " — " & PartsText(oldN, oldU) & ") этот ЕИ уже выдавали частями: " _
                    & cnt & " раз. Проведённые выдачи не меняются — новая фасовка действует для следующих."
            ElseIf s <> "" Then
                s = s & Chr(10) & Chr(10) & "Сейчас: в 1 " & d(SC_UNIT) & " — " & PartsText(oldN, oldU) & "."
            End If
        End If
    End If
    PackChangeQuestion = s
End Function

' a part counted in whole pieces (a canister, a bottle, a box) — not a length, a weight or a volume
Private Function PkCountable(pu As String) As Boolean
    Dim w As Variant
    For Each w In Array("м", "мм", "см", "кг", "г", "т", "л", "мл", "м2", "м3", "м²", "м³")
        If LCase(pu) = w Or UnitMatch(pu, CStr(w)) Then Exit Function
    Next w
    PkCountable = True
End Function

Function PackEI() As String
    Dim sName As String, r As Long, sEI As String, n As Long, canon As String, msg As String, d As Variant, v(2) As String, pu As String, pn As Double
    Dim gu As String, gn As Double, info As String, res As String, bal As String
    sName = gDoc.getCurrentController().getActiveSheet().Name
    r = WmsView.ViewRow(sName)
    Select Case sName
    Case SH_STOCK
        If r >= 1 Then sEI = WmsIssue.StockSheet().getCellByPosition(SC_EI, r).getString()
    Case SH_ISSUES
        If r >= 1 Then sEI = WmsIssue.IssuesSheet().getCellByPosition(IC_EI, r).getString()
    Case SH_RETURNS
        If r >= 1 Then sEI = gDoc.Sheets.getByName(SH_RETURNS).getCellByPosition(RC_EI, r).getString()
    Case SH_ADJUST
        If r >= 1 Then sEI = gDoc.Sheets.getByName(SH_ADJUST).getCellByPosition(AC_EI, r).getString()
    Case SH_FIND
        sEI = WmsFind.FindCardEI()
    End Select
    If Trim(sEI) = "" Then
        v(0) = ""
        If Not WmsPrimeUi.UiForm("Фасовка ЕИ", "Впишите ЕИ (например 12 или ЕИ-00000012) — или встаньте на его строку в «Наличие», «Выдачи», «Корректировки».", _
            Array("ЕИ"), v, Array(""), "Дальше") Then
            gUiLastMsg = "SKIP:окно закрыто"
            PackEI = gUiLastMsg
            Exit Function
        End If
        sEI = v(0)
    End If
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        PackEI = "ERR:" & msg
        Exit Function
    End If
    msg = WmsIssue.StockRowProblem(n, canon)
    If msg <> "" Then
        PackEI = "ERR:" & msg
        Exit Function
    End If
    d = WmsIssue.StockData(n)
    If PackOf(n, pu, pn) Then
        v(0) = WmsIssue.QtyText(pn)
        v(1) = pu
        info = "Сейчас: в 1 " & d(SC_UNIT) & " — " & PartsText(pn, pu) & ". Измените или очистите число (без фасовки)."
    Else
        PackGuess(CStr(d(SC_NAME)), gn, gu)
        If gn > 0 Then
            v(0) = WmsIssue.QtyText(gn)
            v(1) = gu
            info = "По наименованию похоже: в 1 " & d(SC_UNIT) & " — " & PartsText(gn, gu) & ". Проверьте и нажмите «Записать»."
        ElseIf InStr(LCase(CStr(d(SC_CAT))), "метал") > 0 Then
            v(0) = ""
            v(1) = "м"
            info = "Металл: укажите длину одного прутка, трубы или профиля в метрах (в 1 шт — 6 м) и выдавайте метрами: «1,5 м»."
        Else
            v(0) = ""
            v(1) = ""
            info = "Например: в 1 шт — 4 кан (незамерзайка: по документу 1 шт, по факту 4 канистры); в 1 шт — 6 м (профиль 6 м)."
        End If
    End If
    ' 1.0 (Ч8): a document in kg or t, issued in metres — the weight of one metre is enough («1 м = 2,5 кг»)
    If PkWeightUnit(CStr(d(SC_UNIT))) > 0 And Not PackOf(n, pu, pn) Then
        info = info & " Документ в «" & d(SC_UNIT) & "», а выдаёте метрами? Впишите вес погонного метра: «1 м = 2,5 кг» — WMS посчитает сам."
        If PkWeightUnit(CStr(d(SC_UNIT))) >= 1000 Then info = info & " (В тоннах WMS считает до 0,001 т = 1 кг: выдача метрами округляется до килограмма.)"
    End If
    v(2) = PackSheet().getCellByPosition(PK_NOTE, n).getString()
    If VarType(d(SC_QTY)) = 5 Then bal = " · остаток " & WmsIssue.QtyText(CDbl(d(SC_QTY))) & " " & d(SC_UNIT)
    If Not WmsPrimeUi.UiForm("Фасовка ЕИ — " & canon, d(SC_NAME) & bal & " (учёт — по документу, в «" & d(SC_UNIT) & "»). " & info _
        & " Потом в «Количество» можно писать «1 " & IIf(v(1) <> "", v(1), "кан") & "», а WMS переведёт в " & d(SC_UNIT) & ".", _
        Array("Сколько частей в 1 " & d(SC_UNIT), "Как называть часть (кан, бут, м, кг, л)", "Примечание"), v, Array("", "", ""), "Записать") Then
        gUiLastMsg = "SKIP:окно закрыто"
        PackEI = gUiLastMsg
        Exit Function
    End If
    If InStr(v(0), "=") > 0 Then
        If Not PkWeight(v(0), CStr(d(SC_UNIT)), gn, gu, msg) Then
            PackEI = "ERR:" & msg
            Exit Function
        End If
        v(0) = PkNumText6(gn)
        If Trim(v(1)) = "" Then v(1) = gu
    End If
    msg = PackChangeQuestion(n, canon, d, v(0), v(1))
    If msg <> "" Then
        If WmsPrimeUi.UiButtons("Фасовка ЕИ — проверьте", msg, Array("Записать фасовку", "Не записывать"), Array(1, 0)) <> 1 Then
            gUiLastMsg = "SKIP:фасовка не записана"
            PackEI = gUiLastMsg
            Exit Function
        End If
    End If
    res = PackSet(canon, v(0), v(1), v(2))
    If Left(res, 3) = "OK:" Then
        If PackOf(n, pu, pn) Then
            gUiLastMsg = "OK:" & canon & ": в 1 " & d(SC_UNIT) & " — " & PartsText(pn, pu) & ". Выдавайте как удобно: «1 " & pu & "» или в " & d(SC_UNIT)
        Else
            gUiLastMsg = "OK:" & canon & ": фасовка снята — выдача в " & d(SC_UNIT)
        End If
        PackEI = gUiLastMsg
    Else
        PackEI = res
        If Left(res, 5) = "SKIP:" Then gUiLastMsg = res
    End If
End Function

' "" when every row of _PACK is the packing of its own EI with a positive number of parts and a name; otherwise what is wrong
' (a row moved or typed by hand: WMS ignores it, «Фасовка ЕИ» writes it again)
Function PackProblems() As String
    Dim sh As Object, last As Long, a As String, s As String
    sh = PackSheet()
    If IsNull(sh) Then
        PackProblems = "нет листа фасовки " & SH_PACK & " — книга не обновлена до 0.8"
        Exit Function
    End If
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    a = WmsView.ColRng(SH_PACK, PK_EI, 1, last)
    s = WmsView.Eng("SUMPRODUCT((" & a & "<>"""")*(((" & a & "<>""" & EI_PREFIX & """&TEXT(ROW(" & a & ")-1;""00000000""))+NOT(ISNUMBER(" _
        & WmsView.ColRng(SH_PACK, PK_N, 1, last) & "))+(" & WmsView.ColRng(SH_PACK, PK_UNIT, 1, last) & "=""""))>0))")
    If s = "" Then
        PackProblems = "фасовка не проверена (формула Calc)"
    ElseIf Val(s) > 0 Then
        PackProblems = "строк фасовки не на своём ЕИ или без числа частей: " & s & " — такие строки WMS не использует; задайте фасовку этих ЕИ заново кнопкой «Фасовка ЕИ»"
    End If
End Function

' ================================================================ test seams

Function TestQtyEI(sText As String, n As Long, sMode As String, capQty As Double) As String
    Dim q As Double, info As String, msg As String, st As Integer, b As Double, d As Variant
    WmsInit()
    d = WmsIssue.StockData(n)
    If VarType(d(SC_QTY)) = 5 Then b = d(SC_QTY)
    st = QtyEIText(sText, n, b, CStr(d(SC_UNIT)), sMode, capQty, q, info, msg)
    If st = 0 Then TestQtyEI = "OK:" & WmsIssue.QtyText(q) & "|" & info Else TestQtyEI = "ERR:" & msg
End Function

Function TestPackChange(sEI As String, sN As String, sPart As String) As String
    Dim n As Long, canon As String, msg As String
    WmsInit()
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        TestPackChange = "ERR:" & msg
        Exit Function
    End If
    TestPackChange = PackChangeQuestion(n, canon, WmsIssue.StockData(n), sN, sPart)
End Function

Function TestPackGuess(sName As String) As String
    Dim pn As Double, pu As String
    PackGuess(sName, pn, pu)
    TestPackGuess = WmsIssue.QtyText(pn) & "|" & pu
End Function
