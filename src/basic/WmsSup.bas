' WmsSup — 1.3.0: поставщики одним написанием. The same supplier is written in several ways («Все инструменты», «ВсеИнструменты», «все инструменты.» …); a screen that compares the text exactly
' (the panel of «Приход авто», «Принять товар», «Уже ожидается») takes them for different suppliers. This module (1) gives ONE key of a supplier for all the comparisons (SupId: the case, «ё», the Latin
' letters that look like the Russian ones, the spaces and the signs do not part two spellings), (2) shows on the sheet «Написания поставщиков» every spelling that is in the book — «Заказы» (not received,
' received), «Приход авто», «Номенклатура», «Поставщики» — with the spellings that mean one supplier together and a column «Привести к» for the storekeeper's decision, (3) applies the decisions
' when he presses the button: the not received rows of «Заказы» get the new text as if he typed it; a received position — through «Исправить позицию» (POSITION_FIX), a visit of «Приход авто» —
' through «Исправить» (CAR_FIX): each in the journal, nothing but the supplier changes, quantities and EIs are not touched. A restore point of the book is made before. Nothing here runs by itself.
' Library WmsPlus.
Option Explicit

Public Const SH_SPEL = "Написания поставщиков"
Private Const SPL_STATUS_ROW = 2
Private Const SPL_ROW0 = 4
Private Const SPL_COLS = 11
Private Const SPL_MAX = 3000
' columns: A the spelling, B not received rows, C received rows, D cancelled rows, E visits, F goods, G «Поставщики», H what WMS noticed, I «Привести к» (the input), J what will change, K the spelling as it is (hidden)
Private Const SPL_NOTE_COL = 7
Private Const SPL_TO_COL = 8
Private Const SPL_PREV_COL = 9
Private Const SPL_RAW_COL = 10

Global gSpelSeq As Long
Global gSpelBusy As Boolean
Global gSupSeps As String

' the spellings read from the book (parallel arrays; the exact text is the identity of a spelling), the chain pNext joins the spellings that differ only by the case
Private pTxt() As String
Private pOpen() As Long
Private pRcv() As Long
Private pCan() As Long
Private pCar() As Long
Private pGoods() As Long
Private pDir() As Long
Private pNext() As Long
Private pCnt As Long
Private pBk As Object

' ================================================================ the key of a supplier

' the signs that do not part two spellings of one supplier
Private Function SupSeps() As String
    If gSupSeps = "" Then
        gSupSeps = " " & Chr(9) & Chr(10) & Chr(13) & Chr(160) & Chr(34) & "'" & "«»„“”‘’" & ".,;:-–—_()[]{}/\|№#!?*+=~`^%<>"
    End If
    SupSeps = gSupSeps
End Function

' the key of a supplier for every comparison of two spellings: lower case, «ё» = «е», the Latin letters that look like the Russian ones are Russian (WmsNames.Fold), the spaces, the quotes, the
' dots, the commas, the hyphens and the brackets are left out. «Все инструменты», «ВсеИнструменты», «все инструменты.» and «"Все-Инструменты"» have one key
Public Function SupId(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, out As String, seps As String
    t = WmsNames.Fold(Trim(s))
    seps = SupSeps()
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If InStr(1, seps, ch, 0) = 0 Then out = out & ch
    Next i
    If out = "" Then out = t
    SupId = out
End Function

' the key for «it looks like the same supplier»: SupId of the name without the part in brackets, without the end «.ru», «.рф», «.com» and without the words of the firm form (ООО, ИП, АО …).
' Only a hint: two spellings with one such key are shown together, the storekeeper decides
Public Function SupLike(ByVal s As String) As String
    Dim t As String, p As Long, q As Long, w As Variant, f As String, i As Long, ch As String, seps As String, u As String
    t = Trim(s)
    p = InStr(1, t, "(", 0)
    Do While p > 0
        q = InStr(p + 1, t, ")", 0)
        If q = 0 Then
            t = Left(t, p - 1)
        Else
            t = Left(t, p - 1) & " " & Mid(t, q + 1)
        End If
        p = InStr(1, t, "(", 0)
    Loop
    t = WmsNames.Fold(t)
    For Each w In Array(".ru", ".рф", ".ру", ".com", ".su", ".net", ".org", ".рус")
        f = WmsNames.Fold(CStr(w))
        If Len(t) > Len(f) Then
            If Right(t, Len(f)) = f Then t = Left(t, Len(t) - Len(f))
        End If
    Next w
    seps = SupSeps()
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If InStr(1, seps, ch, 0) > 0 Then ch = " "
        u = u & ch
    Next i
    t = " " & u & " "
    Do While InStr(1, t, "  ", 0) > 0
        t = Replace(t, "  ", " ", 1, -1, 0)
    Loop
    For Each w In Array("ооо", "ип", "ао", "зао", "пао", "оао", "ано", "нпо", "чп", "тоо", "ooo", "llc", "ltd", "inc", "gmbh", "corp", "тд", "тк")
        f = " " & WmsNames.Fold(CStr(w)) & " "
        Do While InStr(1, t, f, 0) > 0
            t = Replace(t, f, " ", 1, -1, 0)
        Loop
    Next w
    If Trim(t) = "" Then
        SupLike = SupId(s)
    Else
        SupLike = SupId(t)
    End If
End Function

' a directory spelling of the supplier typed as s ("" when «Поставщики» has none): the same supplier written otherwise gets the one spelling of the directory
Public Function SupSnap(ByVal s As String) As String
    Dim r As Long, sh As Object
    SupSnap = s
    If Trim(s) = "" Then Exit Function
    r = WmsSupply.SupplierRow(s)
    If r < 1 Then Exit Function
    sh = gDoc.Sheets.getByName(WmsSupply.SH_SUPPLIERS)
    If Trim(sh.getCellByPosition(WmsSupply.SU_NAME, r).getString()) <> "" Then SupSnap = Trim(sh.getCellByPosition(WmsSupply.SU_NAME, r).getString())
End Function

' the supplier typed into a cell (sheet sSheet, row r, column c) gets the spelling of «Поставщики» when it is the same supplier by its key; True when the cell was changed. Special receipts
' («Офис», «Иной приход» …) are not touched
Function SupFixCell(sSheet As String, r As Long, c As Integer) As Boolean
    Dim cell As Object, t As String, u As String
    On Error GoTo EH
    cell = gDoc.Sheets.getByName(sSheet).getCellByPosition(c, r)
    If cell.getType() <> com.sun.star.table.CellContentType.TEXT Then Exit Function
    t = cell.getString()
    If Trim(t) = "" Then Exit Function
    If WmsOrders.SpecialSupplier(t) <> "" Then Exit Function
    u = SupSnap(t)
    If u <> t Then
        cell.setString(u)
        SupFixCell = True
    End If
    Exit Function
EH:
End Function

' ================================================================ reading the spellings from the book

Private Function CollTry(coll As Object, key As String, ByRef v As Variant) As Boolean
    On Error GoTo Miss
    v = coll.Item(key)
    CollTry = True
    Exit Function
Miss:
    CollTry = False
End Function

Private Sub SpReset()
    pCnt = 0
    ReDim pTxt(63)
    ReDim pOpen(63)
    ReDim pRcv(63)
    ReDim pCan(63)
    ReDim pCar(63)
    ReDim pGoods(63)
    ReDim pDir(63)
    ReDim pNext(63)
    pBk = New Collection
End Sub

Private Function SpAdd(ByVal raw As String) As Long
    Dim m As Long
    If pCnt > UBound(pTxt) Then
        m = 2 * pCnt + 1
        ReDim Preserve pTxt(m)
        ReDim Preserve pOpen(m)
        ReDim Preserve pRcv(m)
        ReDim Preserve pCan(m)
        ReDim Preserve pCar(m)
        ReDim Preserve pGoods(m)
        ReDim Preserve pDir(m)
        ReDim Preserve pNext(m)
    End If
    pTxt(pCnt) = raw
    pNext(pCnt) = -1
    SpAdd = pCnt
    pCnt = pCnt + 1
End Function

' the number of the spelling raw (exactly this text, the case counts); added when new
Private Function SpIndex(ByVal raw As String) As Long
    Dim k As String, i As Long, last As Long, v As Variant, n As Long
    k = LCase(raw)
    If CollTry(pBk, k, v) Then
        i = CLng(v)
        Do
            If pTxt(i) = raw Then
                SpIndex = i
                Exit Function
            End If
            last = i
            i = pNext(i)
        Loop While i >= 0
        n = SpAdd(raw)
        pNext(last) = n
        SpIndex = n
        Exit Function
    End If
    n = SpAdd(raw)
    pBk.Add(n, k)
    SpIndex = n
End Function

' the number of the spelling raw when the book has it, otherwise -1 (nothing is added)
Private Function SpFind(ByVal raw As String) As Long
    Dim i As Long, v As Variant
    SpFind = -1
    If Not CollTry(pBk, LCase(raw), v) Then Exit Function
    i = CLng(v)
    Do
        If pTxt(i) = raw Then
            SpFind = i
            Exit Function
        End If
        i = pNext(i)
    Loop While i >= 0
End Function

' the kind of a row of «Заказы»: 0 — not a position (a copy, an empty row), 1 — not received, 2 — received (any state of the receipt), 3 — cancelled before any receipt
Private Function SupRowKind(r As Variant) As Long
    If Left(CStr(r(OC_CTL)), 5) = "КОПИЯ" Then Exit Function
    If CStr(r(OC_EI)) <> "" Then
        SupRowKind = 2
    ElseIf CStr(r(OC_STATUS)) = OS_CANCELLED Then
        SupRowKind = 3
    ElseIf Left(CStr(r(OC_BLOCK)), 2) <> "OL" Then
        SupRowKind = 1
    End If
End Function

' every spelling of a supplier that the book has, with how many times where (SpReset first)
Private Sub SpRead()
    Dim sh As Object, last As Long, d As Variant, i As Long, sup As String, k As Long, j As Long
    SpReset()
    sh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        d = sh.getCellRangeByPosition(0, 1, OC_BLOCK, last).getDataArray()
        For i = 0 To UBound(d)
            sup = CStr(d(i)(OC_SUPPLIER))
            If Trim(sup) <> "" Then
                If WmsOrders.SpecialSupplier(sup) = "" Then
                    k = SupRowKind(d(i))
                    If k > 0 Then
                        j = SpIndex(sup)
                        If k = 1 Then
                            pOpen(j) = pOpen(j) + 1
                        ElseIf k = 2 Then
                            pRcv(j) = pRcv(j) + 1
                        Else
                            pCan(j) = pCan(j) + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If
    If gDoc.Sheets.hasByName(SH_CARS) Then
        sh = gDoc.Sheets.getByName(SH_CARS)
        last = WmsOrders.LastRow(sh)
        If last >= CAR_FIRST Then
            d = sh.getCellRangeByPosition(0, CAR_FIRST, CC_STATUS, last).getDataArray()
            For i = 0 To UBound(d)
                sup = CStr(d(i)(CC_SUP))
                If Trim(sup) <> "" Then
                    j = SpIndex(sup)
                    If CStr(d(i)(CC_STATUS)) = CS_CANCEL Then
                        pCan(j) = pCan(j) + 1
                    Else
                        pCar(j) = pCar(j) + 1
                    End If
                End If
            Next i
        End If
    End If
    If gDoc.Sheets.hasByName(SH_GOODS) Then
        sh = gDoc.Sheets.getByName(SH_GOODS)
        last = WmsOrders.LastRow(sh)
        If last >= 1 Then
            d = sh.getCellRangeByPosition(0, 1, WmsSupply.GD_SUP, last).getDataArray()
            For i = 0 To UBound(d)
                sup = CStr(d(i)(WmsSupply.GD_SUP))
                If Trim(sup) <> "" Then
                    j = SpIndex(sup)
                    pGoods(j) = pGoods(j) + 1
                End If
            Next i
        End If
    End If
    If gDoc.Sheets.hasByName(WmsSupply.SH_SUPPLIERS) Then
        sh = gDoc.Sheets.getByName(WmsSupply.SH_SUPPLIERS)
        last = WmsOrders.LastRow(sh)
        If last >= 1 Then
            d = sh.getCellRangeByPosition(WmsSupply.SU_NAME, 1, WmsSupply.SU_NAME, last).getDataArray()
            For i = 0 To UBound(d)
                sup = CStr(d(i)(0))
                If Trim(sup) <> "" Then
                    j = SpIndex(sup)
                    If pDir(j) = 0 Then pDir(j) = i + 2
                End If
            Next i
        End If
    End If
End Sub

' ================================================================ the sheet «Написания поставщиков»

Private Function EdgeSpaces(s As String) As Boolean
    EdgeSpaces = (s <> Trim(s)) Or (InStr(s, "  ") > 0) Or (InStr(s, Chr(160)) > 0)
End Function

' the text of "what will change" for a spelling: the counts of the rows (open, received, cancelled, visits, goods), whether the directory has it, and the new text
Private Function Preview(nOpen As Long, nRcv As Long, nCan As Long, nCarV As Long, nGds As Long, bDir As Boolean, sTo As String) As String
    Dim s As String
    If nOpen > 0 Then s = s & ", ожидаем (строк): " & nOpen
    If nRcv > 0 Then s = s & ", принято (строк): " & nRcv
    If nCarV > 0 Then s = s & ", визитов: " & nCarV
    If nGds > 0 Then s = s & ", товаров: " & nGds
    If bDir Then s = s & ", строка справочника"
    If nCan > 0 Then s = s & " (отменённых строк и визитов: " & nCan & " — не меняются)"
    If s = "" Then
        Preview = "нигде не встречается"
    Else
        Preview = "станет «" & sTo & "»:" & Mid(s, 2)
    End If
End Function

' how much a spelling is used (rows, visits, goods)
Private Function Used(j As Long) As Long
    Used = pOpen(j) + pRcv(j) + pCan(j) + pCar(j) + pGoods(j)
End Function

' True when spelling i is a better main spelling than c: the one the directory has, then the more used, then the first by the binary order of the text
Private Function Better(i As Long, c As Long) As Boolean
    If (pDir(i) > 0) <> (pDir(c) > 0) Then
        Better = (pDir(i) > 0)
    ElseIf Used(i) <> Used(c) Then
        Better = (Used(i) > Used(c))
    Else
        Better = (pTxt(i) < pTxt(c))
    End If
End Function

' the number j of the spellings sorted: the groups that look like one supplier first (the more used first), then the others alphabetically; inside a group the more used first
Private Sub SpOrder(ord() As Long, idk() As String, lk() As String, ByRef nMulti As Long)
    Dim i As Long, j As Long, g As Object, v As Variant, gi As Long, nGr As Long, gUse() As Long, gCnt() As Long, gOf() As Long, key() As String, use As Long, h As Long, t As Long, tmp As String
    g = New Collection
    ReDim gOf(pCnt - 1)
    ReDim gUse(pCnt)
    ReDim gCnt(pCnt)
    For i = 0 To pCnt - 1
        If CollTry(g, lk(i), v) Then
            gi = CLng(v)
        Else
            gi = nGr
            g.Add(gi, lk(i))
            nGr = nGr + 1
        End If
        gOf(i) = gi
        gCnt(gi) = gCnt(gi) + 1
        gUse(gi) = gUse(gi) + Used(i)
    Next i
    ReDim key(pCnt - 1)
    ReDim ord(pCnt - 1)
    For i = 0 To pCnt - 1
        gi = gOf(i)
        use = Used(i)
        If use > 999999 Then use = 999999
        h = gUse(gi)
        If h > 999999 Then h = 999999
        If gCnt(gi) >= 2 Then
            key(i) = "0" & Right("000000" & CStr(999999 - h), 6) & lk(i) & Chr(1) & Right("000000" & CStr(999999 - use), 6) & pTxt(i)
        Else
            key(i) = "1" & lk(i) & Chr(1) & pTxt(i)
        End If
        ord(i) = i
    Next i
    ' shell sort by the order of the keys (the operators < and <= compare the letters as they are, the case counts)
    h = 1
    Do While h < pCnt \ 3
        h = 3 * h + 1
    Loop
    Do While h >= 1
        For i = h To pCnt - 1
            t = ord(i)
            tmp = key(t)
            j = i
            Do While j >= h
                If key(ord(j - h)) <= tmp Then Exit Do
                ord(j) = ord(j - h)
                j = j - h
            Loop
            ord(j) = t
        Next i
        h = h \ 3
    Loop
    nMulti = 0
    For i = 0 To nGr - 1
        If gCnt(i) >= 2 Then nMulti = nMulti + 1
    Next i
End Sub

Sub BtnSpelRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SpelRefresh()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the sheet is opened: the list is made again when the journal moved since it was made (a new book — the first time)
Sub OnSpelFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gSpelBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gSpelSeq And gSpelSeq > 0 Then Exit Sub
    SpelRefresh()
    Exit Sub
EH:
End Sub

' a cell of "Привести к" was changed: "what will change" is written again for the row (from the counts of the row itself)
Sub OnSpelChange(Optional oTarget As Variant)
    Dim ra As Object, sh As Object, raw As String, sTo As String, d As Variant
    If gSpelBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Column <> SPL_TO_COL Or ra.Row < SPL_ROW0 Then Exit Sub
    WmsInit()
    sh = gDoc.Sheets.getByName(SH_SPEL)
    d = sh.getCellRangeByPosition(0, ra.Row, SPL_COLS - 1, ra.Row).getDataArray()(0)
    raw = CStr(d(SPL_RAW_COL))
    If raw = "" Then Exit Sub
    sTo = Trim(CStr(d(SPL_TO_COL)))
    gSpelBusy = True
    If sTo = "" Or sTo = raw Then
        sh.getCellByPosition(SPL_PREV_COL, ra.Row).setString("")
    Else
        sh.getCellByPosition(SPL_PREV_COL, ra.Row).setString(Preview(CLng(Val(CStr(d(1)))), CLng(Val(CStr(d(2)))), CLng(Val(CStr(d(3)))), CLng(Val(CStr(d(4)))), CLng(Val(CStr(d(5)))), (Left(CStr(d(6)), 6) = "строка"), sTo))
    End If
    gSpelBusy = False
    Exit Sub
EH:
    gSpelBusy = False
End Sub

' the list of the spellings: the spellings of one supplier (the key SupId is the same) together, the proposal «Привести к» only for those; the others that look alike are only named. OK:<rows> / ERR:…
Function SpelRefresh() As String
    Dim sh As Object, i As Long, j As Long, k As Long, n As Long, idk() As String, lk() As String, ord() As Long, nMulti As Long, out() As Variant, sNote As String, sProp As String
    Dim canon() As Long, likeBest() As Long, grp As Object, lcn As Object, v As Variant, gi As Long, nId As Long, nLk As Long, gOf() As Long, lOf() As Long, gBest() As Long, lBest() As Long, nSame As Long
    Dim nTotal As Long, sDir As String, sE As String, nE As Long
    On Error GoTo EH
    gSpelBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_SPEL) Then
        SpelRefresh = "ERR:в книге нет листа «" & SH_SPEL & "» — книга не обновлена до 1.3"
        gSpelBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_SPEL)
    SpRead()
    n = pCnt
    sh.getCellRangeByPosition(0, SPL_ROW0, SPL_COLS - 1, SPL_ROW0 + SPL_MAX).clearContents(1 + 2 + 4 + 16)
    If n = 0 Then
        sh.getCellByPosition(0, SPL_STATUS_ROW).setString("Поставщиков в книге нет: ни в «Заказах», ни в «Приходе авто», ни в «Номенклатуре», ни в «Поставщиках». Список на " & Format(Now(), "DD.MM.YYYY HH:MM") _
            & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
        gSpelSeq = CLng(SysNum(SK_LAST_SEQ))
        If gSpelSeq = 0 Then gSpelSeq = 1
        gSpelBusy = False
        SpelRefresh = "OK:0"
        gUiLastMsg = SpelRefresh
        Exit Function
    End If
    ReDim idk(n - 1)
    ReDim lk(n - 1)
    ReDim gOf(n - 1)
    ReDim lOf(n - 1)
    grp = New Collection
    lcn = New Collection
    For i = 0 To n - 1
        idk(i) = SupId(pTxt(i))
        lk(i) = SupLike(pTxt(i))
        If CollTry(grp, idk(i), v) Then
            gOf(i) = CLng(v)
        Else
            gOf(i) = nId
            grp.Add(nId, idk(i))
            nId = nId + 1
        End If
        If CollTry(lcn, lk(i), v) Then
            lOf(i) = CLng(v)
        Else
            lOf(i) = nLk
            lcn.Add(nLk, lk(i))
            nLk = nLk + 1
        End If
    Next i
    ' the main spelling of every set with one key SupId, and the most used of every set that looks alike
    ReDim gBest(nId)
    ReDim lBest(nLk)
    For i = 0 To nId
        gBest(i) = -1
    Next i
    For i = 0 To nLk
        lBest(i) = -1
    Next i
    For i = 0 To n - 1
        If gBest(gOf(i)) < 0 Then
            gBest(gOf(i)) = i
        ElseIf Better(i, gBest(gOf(i))) Then
            gBest(gOf(i)) = i
        End If
        If lBest(lOf(i)) < 0 Then
            lBest(lOf(i)) = i
        ElseIf Better(i, lBest(lOf(i))) Then
            lBest(lOf(i)) = i
        End If
    Next i
    SpOrder(ord, idk, lk, nMulti)
    ReDim out(n - 1)
    nTotal = n
    For k = 0 To n - 1
        i = ord(k)
        sNote = ""
        sProp = ""
        nSame = 0
        For j = 0 To n - 1
            If gOf(j) = gOf(i) And j <> i Then nSame = nSame + 1
        Next j
        If pOpen(i) + pRcv(i) + pCar(i) + pGoods(i) + IIf(pDir(i) > 0, 1, 0) = 0 And pCan(i) > 0 Then
            sNote = "встречается только в отменённых строках и визитах — они не меняются"
        ElseIf gBest(gOf(i)) <> i Then
            sNote = "то же, что «" & pTxt(gBest(gOf(i))) & "» (регистр, пробелы, знаки, русские и латинские буквы)"
            sProp = pTxt(gBest(gOf(i)))
        ElseIf nSame > 0 Then
            sNote = "основное написание (так записано чаще или так в справочнике)"
        ElseIf gOf(lBest(lOf(i))) <> gOf(i) Then
            sNote = "похоже на «" & pTxt(lBest(lOf(i))) & "» — если это один поставщик, впишите справа то же написание"
        End If
        If EdgeSpaces(pTxt(i)) Then sNote = sNote & IIf(sNote <> "", "; ", "") & "в тексте лишние пробелы"
        sDir = "нет"
        If pDir(i) > 0 Then sDir = "строка " & pDir(i)
        out(k) = Array(pTxt(i), pOpen(i), pRcv(i), pCan(i), pCar(i), pGoods(i), sDir, sNote, sProp, IIf(sProp <> "", Preview(pOpen(i), pRcv(i), pCan(i), pCar(i), pGoods(i), (pDir(i) > 0), sProp), ""), pTxt(i))
    Next k
    If n > SPL_MAX Then
        ReDim Preserve out(SPL_MAX - 1)
        n = SPL_MAX
    End If
    sh.getCellRangeByPosition(0, SPL_ROW0, SPL_COLS - 1, SPL_ROW0 + n - 1).setDataArray(out)
    sh.getCellByPosition(0, SPL_STATUS_ROW).setString("Написаний: " & nTotal & IIf(nTotal > n, " (показано " & n & ")", "") & "; наборов, которые выглядят как один поставщик: " & nMulti _
        & ". Список на " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
    gSpelSeq = CLng(SysNum(SK_LAST_SEQ))
    If gSpelSeq = 0 Then gSpelSeq = 1
    gSpelBusy = False
    SpelRefresh = "OK:" & n
    gUiLastMsg = SpelRefresh
    Exit Function
EH:
    sE = Error$
    nE = Erl
    gSpelBusy = False
    SpelRefresh = "ERR:список написаний не построен: " & sE & " (строка " & nE & ")"
    gUiLastMsg = SpelRefresh
End Function

' ================================================================ applying the decisions

Sub BtnSpelApply(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SpelApply()
    If Left(s, 4) = "ERR:" Then
        WmsUi.UiMessage(Mid(s, 5))
    ElseIf Left(s, 3) = "OK:" Then
        WmsOrdersUi.UiInfo(Mid(s, 4))
    End If
End Sub

Private Function SupCellShown(c As Object) As String
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then SupCellShown = NumStr(c.getValue()) Else SupCellShown = c.getString()
End Function

' a received row r of «Заказы»: the position gets the supplier sTo through «Исправить позицию» (POSITION_FIX); everything else of the position is passed as it is written, and if anything else would
' change the position is not touched (gPosFixOnly). OK:<seq> / SKIP:… / ERR:…
Private Function PositionSupplier(r As Long, sTo As String) As String
    Dim sh As Object, od As Variant, src As Long, kind As String, c As Object, vQty As String, vOdate As String, res As String, sE As String, nE As Long
    On Error GoTo EH
    sh = WmsOrders.OrdersSheet()
    kind = WmsOrders.OrderRowKind(r)
    If kind <> "RECEIVED" And kind <> "STORNO" Then
        PositionSupplier = "SKIP:строка " & (r + 1) & " сейчас не принятая (" & kind & ")"
        Exit Function
    End If
    src = r
    If Not gKsrc Then
        If WmsOrders.OlValid(gKol, od) Then src = WmsOrders.SourceRowOf(od)
        If src < 1 Then src = r
    End If
    c = sh.getCellByPosition(OC_ORDQTY, src)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then vQty = WmsIssue.QtyText(c.getValue()) Else vQty = c.getString()
    c = sh.getCellByPosition(OC_ODATE, src)
    vOdate = ""
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then vOdate = Format(c.getValue(), "DD.MM.YYYY")
    gPosFixOnly = "L"
    res = WmsReceipt.PositionFixRow(r, SupCellShown(sh.getCellByPosition(OC_ORDER, src)), sh.getCellByPosition(OC_NAME, src).getString(), SupCellShown(sh.getCellByPosition(OC_INVOICE, src)), _
        sh.getCellByPosition(OC_ART, src).getString(), vQty, sh.getCellByPosition(OC_UNIT, src).getString(), sTo, vOdate, sh.getCellByPosition(OC_CAT, src).getString())
    gPosFixOnly = ""
    PositionSupplier = res
    Exit Function
EH:
    sE = Error$
    nE = Erl
    gPosFixOnly = ""
    PositionSupplier = "ERR:строка " & (r + 1) & ": " & sE & " (строка кода " & nE & ")"
End Function

' the rows of the directory «Поставщики» with the spelling sFrom become the row of sTo: the phone, the contact, the time, the note go to the cells of sTo that are empty; the row of sFrom goes
' (the first of them is renamed when sTo is not in the directory yet). The number of the rows removed
Private Function DirectoryMerge(sFrom As String, sTo As String) As Long
    Dim sh As Object, last As Long, d As Variant, i As Long, t As Long, c As Long, tc As Object, sc As Object, n As Long, del() As Long, kTo As String
    If Not gDoc.Sheets.hasByName(WmsSupply.SH_SUPPLIERS) Then Exit Function
    sh = gDoc.Sheets.getByName(WmsSupply.SH_SUPPLIERS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(WmsSupply.SU_NAME, 1, WmsSupply.SU_NAME, last).getDataArray()
    t = -1
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(0))) = Trim(sTo) Then
            t = i
            Exit For
        End If
    Next i
    ReDim del(UBound(d))
    For i = 0 To UBound(d)
        If CStr(d(i)(0)) = sFrom And i <> t Then
            If t < 0 Then
                sh.getCellByPosition(WmsSupply.SU_NAME, i + 1).setString(sTo)
                t = i
            Else
                For c = WmsSupply.SU_PHONE To WmsSupply.SU_NOTE
                    tc = sh.getCellByPosition(c, t + 1)
                    sc = sh.getCellByPosition(c, i + 1)
                    If tc.getType() = com.sun.star.table.CellContentType.EMPTY And sc.getType() <> com.sun.star.table.CellContentType.EMPTY Then
                        If sc.getType() = com.sun.star.table.CellContentType.VALUE Then tc.setValue(sc.getValue()) Else tc.setString(sc.getString())
                    End If
                Next c
                del(n) = i + 1
                n = n + 1
            End If
        End If
    Next i
    For i = n - 1 To 0 Step -1
        sh.getRows().removeByIndex(del(i), 1)
    Next i
    DirectoryMerge = n
End Function

' the decisions of the sheet, applied. The sheet is read as it is: a row with a text in «Привести к» that is not the spelling of the row itself. OK:<text> / ERR:<why> / SKIP:<why>
Function SpelApply() As String
    Dim sh As Object, last As Long, d As Variant, i As Long, j As Long, k As Long, raw As String, sTo As String, nMap As Long, mapTo() As String, why As String
    Dim nOpenRows As Long, nRcvRows As Long, nVis As Long, nGoods As Long, nDir As Long, nCanRows As Long, msg As String
    Dim shOrd As Object, dOrd As Variant, r As Long, kind As Long, sup As String, res As String, nOpenDone As Long, nPosDone As Long, nPosSkip As Long, nPosErr As Long, nVisDone As Long, nVisErr As Long
    Dim firstErr As String, bStop As Boolean, nDirDone As Long, nGoodsDone As Long, nWork As Long, nDone As Long, sCheck As String, sPoint As String, sE As String, nE As Long, bProg As Boolean
    Dim dC As Variant, dG As Variant, cc As Object, nVisSkip As Long
    On Error GoTo EH
    gSpelBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_SPEL) Then
        SpelApply = "ERR:в книге нет листа «" & SH_SPEL & "» — книга не обновлена до 1.3"
        gSpelBusy = False
        Exit Function
    End If
    why = WmsCore.PostingBlockReason()
    If why <> "" Then
        SpelApply = IIf(InStr(why, ":") > 0, "ERR:" & Mid(why, InStr(why, ":") + 1), "ERR:" & why)
        gSpelBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_SPEL)
    last = WmsOrders.LastRow(sh)
    If last < SPL_ROW0 Then
        SpelApply = "SKIP:в списке нет написаний"
        gSpelBusy = False
        Exit Function
    End If
    SpRead()
    ReDim mapTo(pCnt)
    d = sh.getCellRangeByPosition(0, SPL_ROW0, SPL_COLS - 1, last).getDataArray()
    For i = 0 To UBound(d)
        raw = CStr(d(i)(SPL_RAW_COL))
        sTo = Trim(CStr(d(i)(SPL_TO_COL)))
        If raw <> "" And sTo <> "" Then
            If sTo <> raw Then
                If WmsOrders.SpecialSupplier(sTo) <> "" Then
                    SpelApply = "ERR:«" & sTo & "» — не поставщик, а особый приход (" & SPECIAL_SUPPLIERS & "); впишите другое написание. Строка " & (SPL_ROW0 + i + 1) & " листа"
                    gSpelBusy = False
                    Exit Function
                End If
                If Len(sTo) > CAR_TEXT_MAX Then
                    SpelApply = "ERR:«" & Left(sTo, 30) & "…» длиннее " & CAR_TEXT_MAX & " знаков. Строка " & (SPL_ROW0 + i + 1) & " листа"
                    gSpelBusy = False
                    Exit Function
                End If
                For k = 1 To Len(sTo)
                    If Asc(Mid(sTo, k, 1)) < 32 Then
                        SpelApply = "ERR:в «" & sTo & "» есть знак перевода строки или табуляции. Строка " & (SPL_ROW0 + i + 1) & " листа"
                        gSpelBusy = False
                        Exit Function
                    End If
                Next k
                j = SpFind(raw)
                If j >= 0 Then
                    mapTo(j) = sTo
                    nMap = nMap + 1
                End If
            End If
        End If
    Next i
    If nMap = 0 Then
        SpelApply = "SKIP:в столбце «Привести к» ничего не вписано (или вписано то же написание) — менять нечего"
        gSpelBusy = False
        Exit Function
    End If
    ' a chain (A → B and B → C) is not followed: the storekeeper makes it one step
    For j = 0 To pCnt - 1
        If mapTo(j) <> "" Then
            k = SpFind(mapTo(j))
            If k >= 0 Then
                If mapTo(k) <> "" Then
                    SpelApply = "ERR:«" & pTxt(j) & "» → «" & mapTo(j) & "», а «" & mapTo(j) & "» тоже меняется на «" & mapTo(k) & "». Впишите для обоих одно и то же конечное написание"
                    gSpelBusy = False
                    Exit Function
                End If
            End If
        End If
    Next j
    ' the work, counted (the sheets are read again: the list may be older than the book)
    shOrd = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(shOrd)
    If last >= 1 Then
        dOrd = shOrd.getCellRangeByPosition(0, 1, OC_BLOCK, last).getDataArray()
        For i = 0 To UBound(dOrd)
            sup = CStr(dOrd(i)(OC_SUPPLIER))
            If Trim(sup) <> "" Then
                j = SpFind(sup)
                If j >= 0 Then
                    If mapTo(j) <> "" Then
                        kind = SupRowKind(dOrd(i))
                        If kind = 1 Then
                            nOpenRows = nOpenRows + 1
                        ElseIf kind = 2 Then
                            nRcvRows = nRcvRows + 1
                        ElseIf kind = 3 Then
                            nCanRows = nCanRows + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If
    For j = 0 To pCnt - 1
        If mapTo(j) <> "" Then
            nVis = nVis + pCar(j)
            nGoods = nGoods + pGoods(j)
            If pDir(j) > 0 Then nDir = nDir + 1
        End If
    Next j
    nWork = nOpenRows + nRcvRows + nVis + nGoods + nDir
    If nWork = 0 Then
        SpelApply = "SKIP:выбранных написаний в книге уже нет (кроме отменённых заказов: " & nCanRows & ") — менять нечего"
        gSpelBusy = False
        Exit Function
    End If
    If Not WmsUi.UiConfirm("Привести поставщиков к одному написанию?" & Chr(10) & Chr(10) & "Написаний к замене: " & nMap & "." & Chr(10) & "Будет изменено: строк «Заказов» ещё не принятых — " & nOpenRows _
        & ", строк уже принятых — " & nRcvRows & " (исправление позиции, каждое в журнале), визитов «Прихода авто» — " & nVis & ", товаров «Номенклатуры» — " & nGoods & ", строк справочника «Поставщики» — " & nDir & "." _
        & IIf(nCanRows > 0, Chr(10) & "Отменённые без прихода строки (" & nCanRows & ") не меняются.", "") & Chr(10) & Chr(10) _
        & "Количества, остатки и ЕИ не меняются — только название поставщика. Перед началом программа сделает копию книги. Продолжить?") Then
        SpelApply = "SKIP:не подтверждено — ничего не изменено"
        gSpelBusy = False
        Exit Function
    End If
    sPoint = WmsSafe.RestorePoint("перед приведением поставщиков к одному написанию")
    If Left(sPoint, 6) = "ОШИБКА" Or InStr(sPoint, "невозможна") > 0 Then
        SpelApply = "ERR:копию книги сделать не удалось, поэтому ничего не изменено. " & sPoint
        gSpelBusy = False
        Exit Function
    End If
    If nRcvRows + nVis > 15 Then
        WmsView.ProgressStart("Поставщики: привожу к одному написанию…", nRcvRows + nVis + 1)
        bProg = True
    End If
    ' 1. the visits of «Приход авто»: «Исправить» (CAR_FIX) with the reason
    If nVis > 0 And gDoc.Sheets.hasByName(SH_CARS) Then
        cc = gDoc.Sheets.getByName(SH_CARS)
        last = WmsOrders.LastRow(cc)
        If last >= CAR_FIRST Then
            dC = cc.getCellRangeByPosition(0, CAR_FIRST, CC_STATUS, last).getDataArray()
            dG = cc.getCellRangeByPosition(CC_SUP, CAR_FIRST, CC_SUP, last).getDataArray()
            gOpReason = "поставщик приведён к одному написанию"
            For i = 0 To UBound(dG)
                sup = CStr(dG(i)(0))
                If Trim(sup) <> "" And VarType(dC(i)(CC_NO)) = 5 And CStr(dC(i)(CC_STATUS)) <> CS_CANCEL Then
                    j = SpFind(sup)
                    If j >= 0 Then
                        If mapTo(j) <> "" Then
                            res = WmsCar.CarFix(CLng(dC(i)(CC_NO)), "", mapTo(j), "", "")
                            If Left(res, 3) = "OK:" Then
                                nVisDone = nVisDone + 1
                            ElseIf Left(res, 5) = "SKIP:" Or (Left(res, 4) = "ERR:" And InStr(res, "отменён") > 0) Then
                                nVisSkip = nVisSkip + 1
                            Else
                                nVisErr = nVisErr + 1
                                If firstErr = "" Then firstErr = "визит № " & dC(i)(CC_NO) & ": " & Mid(res, InStr(res, ":") + 1)
                                If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Then
                                    bStop = True
                                    Exit For
                                End If
                            End If
                            nDone = nDone + 1
                            If bProg Then
                                If (nDone Mod 5) = 0 Then WmsView.ProgressSet(nDone)
                            End If
                        End If
                    End If
                End If
            Next i
            gOpReason = ""
        End If
    End If
    ' 2. «Заказы»: a not received row — the new text in the cell as if typed; a received row — its position through «Исправить позицию»
    If Not bStop Then
        last = WmsOrders.LastRow(shOrd)
        If last >= 1 Then
            dOrd = shOrd.getCellRangeByPosition(0, 1, OC_BLOCK, last).getDataArray()
            For i = 0 To UBound(dOrd)
                sup = CStr(dOrd(i)(OC_SUPPLIER))
                If Trim(sup) <> "" Then
                    j = SpFind(sup)
                    If j >= 0 Then
                        If mapTo(j) <> "" Then
                            kind = SupRowKind(dOrd(i))
                            r = i + 1
                            If kind = 1 Then
                                If WmsOrders.OrderRowKind(r) = "OPEN" Then
                                    If shOrd.getCellByPosition(OC_SUPPLIER, r).getString() = sup Then
                                        WmsOrders.HandlerOff(True)
                                        shOrd.getCellByPosition(OC_SUPPLIER, r).setString(mapTo(j))
                                        WmsOrders.HandlerOff(False)
                                        nOpenDone = nOpenDone + 1
                                    End If
                                End If
                            ElseIf kind = 2 Then
                                ' (a position has several rows: the first fix changes them all, the next rows are seen already changed)
                                If shOrd.getCellByPosition(OC_SUPPLIER, r).getString() = sup Then
                                    res = PositionSupplier(r, mapTo(j))
                                    If Left(res, 3) = "OK:" Then
                                        nPosDone = nPosDone + 1
                                    ElseIf Left(res, 5) = "SKIP:" Then
                                        nPosSkip = nPosSkip + 1
                                        If firstErr = "" And InStr(res, "ничего не изменилось") = 0 Then firstErr = "строка " & (r + 1) & ": " & Mid(res, 6)
                                    Else
                                        nPosErr = nPosErr + 1
                                        If firstErr = "" Then firstErr = "строка " & (r + 1) & ": " & Mid(res, InStr(res, ":") + 1)
                                        If Left(res, 8) = "BLOCKED:" Or Left(res, 13) = "ERR-CRITICAL:" Or Left(res, 5) = "BUSY:" Then
                                            bStop = True
                                            Exit For
                                        End If
                                    End If
                                    nDone = nDone + 1
                                    If bProg Then
                                        If (nDone Mod 5) = 0 Then WmsView.ProgressSet(nDone)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End If
    ' 3. «Номенклатура» (the usual supplier of a goods) and the directory «Поставщики»
    If Not bStop Then
        If nGoods > 0 And gDoc.Sheets.hasByName(SH_GOODS) Then
            cc = gDoc.Sheets.getByName(SH_GOODS)
            last = WmsOrders.LastRow(cc)
            If last >= 1 Then
                dG = cc.getCellRangeByPosition(WmsSupply.GD_SUP, 1, WmsSupply.GD_SUP, last).getDataArray()
                For i = 0 To UBound(dG)
                    sup = CStr(dG(i)(0))
                    If Trim(sup) <> "" Then
                        j = SpFind(sup)
                        If j >= 0 Then
                            If mapTo(j) <> "" Then
                                cc.getCellByPosition(WmsSupply.GD_SUP, i + 1).setString(mapTo(j))
                                nGoodsDone = nGoodsDone + 1
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        For j = 0 To pCnt - 1
            If mapTo(j) <> "" And pDir(j) > 0 Then
                DirectoryMerge(pTxt(j), mapTo(j))
                nDirDone = nDirDone + 1
            End If
        Next j
    End If
    If bProg Then
        WmsView.ProgressEnd()
        bProg = False
    End If
    gSpelBusy = False
    SpelRefresh()
    gSpelBusy = True
    sCheck = ""
    If Not bStop Then
        sCheck = WmsDiagnostics.SelfCheck()
        sCheck = Left(sCheck, IIf(InStr(sCheck, Chr(10)) > 0, InStr(sCheck, Chr(10)) - 1, Len(sCheck)))
    End If
    gSpelBusy = False
    msg = "Поставщики приведены к одному написанию." & Chr(10) & "Строк «Заказов», ещё не принятых: изменено " & nOpenDone & "." & Chr(10) & "Принятых позиций исправлено: " & nPosDone _
        & IIf(nPosSkip > 0, ", не тронуто " & nPosSkip, "") & IIf(nPosErr > 0, ", отказ " & nPosErr, "") & "." & Chr(10) & "Визитов «Прихода авто» исправлено: " & nVisDone _
        & IIf(nVisSkip > 0, ", не тронуто " & nVisSkip, "") & IIf(nVisErr > 0, ", отказ " & nVisErr, "") & "." & Chr(10) & "Товаров «Номенклатуры»: " & nGoodsDone & "; строк справочника «Поставщики» объединено: " & nDirDone & "."
    If nCanRows > 0 Then msg = msg & Chr(10) & "Отменённые без прихода строки (" & nCanRows & ") не менялись."
    If firstErr <> "" Then msg = msg & Chr(10) & "Первое замечание: " & firstErr & "."
    If bStop Then msg = msg & Chr(10) & "Работа остановлена: книга просит внимания — перезапустите WMS."
    If sCheck <> "" Then msg = msg & Chr(10) & sCheck
    msg = msg & Chr(10) & "Копия книги до замены: " & sPoint
    SpelApply = "OK:" & msg
    gUiLastMsg = SpelApply
    Exit Function
EH:
    sE = Error$
    nE = Erl
    gOpReason = ""
    gPosFixOnly = ""
    WmsOrders.HandlerOff(False)
    If bProg Then WmsView.ProgressEnd()
    gSpelBusy = False
    SpelApply = "ERR:замена не завершена: " & sE & " (строка " & nE & "). Уже исправленное осталось в журнале; повторите — программа продолжит с того места."
End Function

' the sheet is opened from «Поставщики»
Sub NavSpel(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_SPEL, 1, 0)
End Sub
