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
    PkSplit(s, sNum, sWord)
    If sWord = "" Then
        st = ParseQtyText(s, q, msg)
        If st = 0 Then info = PlainInfo(n, q, sUnit)
        QtyEIText = st
        Exit Function
    End If
    If sNum = "" Then
        msg = "«" & Trim(s) & "»: сначала число, потом единица — например 1 кан или 1,5 м"
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

' the same for a cell: a number in a plain format — the unit of the document; a number Calc formatted as a percent
' (25% typed) — a percent; a text — QtyEIText; a date, time, fraction or formula — refused as by ParseQtyCell
Function QtyEICell(oCell As Object, n As Long, b As Double, sUnit As String, sMode As String, capQty As Double, ByRef q As Double, _
    ByRef info As String, ByRef msg As String) As Integer
    Dim ft As Long, st As Integer
    q = 0
    info = ""
    gPkKind = ""
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.TEXT
        QtyEICell = QtyEIText(oCell.getString(), n, b, sUnit, sMode, capQty, q, info, msg)
    Case com.sun.star.table.CellContentType.VALUE
        ft = gDoc.getNumberFormats().getByKey(oCell.NumberFormat).Type
        If (ft And 128) <> 0 And (ft And (2 Or 4 Or 32 Or 64 Or 1024)) = 0 Then
            QtyEICell = QtyPercent(WmsIssue.Round3(oCell.getValue() * 100), n, sUnit, sMode, q, info, msg)
        Else
            st = ParseQtyCell(oCell, q, msg)
            If st = 0 Then info = PlainInfo(n, q, sUnit)
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
    info = ""
    If Not (oCell Is Nothing) Then
        st = QtyEICell(oCell, n, b, sUnit, sMode, capQty, q, info, msg)
    ElseIf VarType(v) = 8 Then
        st = QtyEIText(CStr(v), n, b, sUnit, sMode, capQty, q, info, msg)
    Else
        QtyAsEI = WmsOrders.ParseQtyValueAs(v, sField, q, msg)
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
    If PackOf(n, pu, pn) Then info = info & " (" & PartsText(PartsOf(q, pn), pu) & ")"
    gPkKind = "PCT"
    gPkPct = pct
    gPkBase = qIn
    QtyPercent = 0
End Function

' the fields of the journal line of the operation being planned: how its quantity was given
Sub PlanQtyFields()
    If gPkKind = "PARTS" Then
        PlanField("PARTS", gPkParts)
        PlanField("PART_UNIT", gPkUnit)
        PlanField("PART_N", gPkN)
    ElseIf gPkKind = "PCT" Then
        PlanField("PCT", gPkPct)
        PlanField("PCT_BASE", gPkBase)
    End If
End Sub

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
Sub PackGuess(sName As String, ByRef pn As Double, ByRef pu As String)
    Dim s As String, t As Variant, i As Integer, w As Variant, x As Double, msg As String, nm As Variant, k As Integer, num As String, rest As String
    pn = 0
    pu = ""
    s = LCase(sName)
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
        Else
            v(0) = ""
            v(1) = ""
            info = "Например: в 1 шт — 4 кан (незамерзайка: по документу 1 шт, по факту 4 канистры); в 1 шт — 6 м (профиль 6 м)."
        End If
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

' ================================================================ test seams

Function TestQtyEI(sText As String, n As Long, sMode As String, capQty As Double) As String
    Dim q As Double, info As String, msg As String, st As Integer, b As Double, d As Variant
    WmsInit()
    d = WmsIssue.StockData(n)
    If VarType(d(SC_QTY)) = 5 Then b = d(SC_QTY)
    st = QtyEIText(sText, n, b, CStr(d(SC_UNIT)), sMode, capQty, q, info, msg)
    If st = 0 Then TestQtyEI = "OK:" & WmsIssue.QtyText(q) & "|" & info Else TestQtyEI = "ERR:" & msg
End Function

Function TestPackGuess(sName As String) As String
    Dim pn As Double, pu As String
    PackGuess(sName, pn, pu)
    TestPackGuess = WmsIssue.QtyText(pn) & "|" & pu
End Function
