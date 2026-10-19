' WmsFind — «Поиск»: one field for everything (PRIME 0.8, D-094). A word, a number, a name, an article, a supplier, a
' recipient, a place, a document — every working sheet is searched (by what the cells show); an EI (427, ЕИ-427,
' ЕИ-00000427) opens its card: who it came from, every movement with the running balance, and whether the balance of
' «Наличие» equals the sum of the movements. A double click on a found row opens its card or its row; «Перейти к строке»
' leads to the row of its sheet. Read only: nothing but this page is written.
Option Explicit

Public Const FD_INPUT_ROW = 3
Public Const FD_INPUT_COL = 2
Public Const FD_HEAD = 6
Public Const FD_INFO = 7
Public Const FD_TABLE_HEAD = 11
Public Const FD_TABLE = 12
Public Const FD_MAX = 500
Public Const FD_CARD_MAX = 3000
' hidden columns of the table: the sheet and the row of the source, the EI
Public Const FD_SRC_SHEET = 10
Public Const FD_SRC_ROW = 11
Public Const FD_SRC_EI = 12

Global gFdBusy As Boolean
Global gFdMode As String
Global gFdQuery As String
Global gFdSpecial As Boolean
' the EI of the card shown on the page (0 — no card)
Global gFdCardEI As Long
' 1.0: the regular expression of the place being searched ("" — a plain text search)
Global gFdPlaceRx As String
' 1.0: the number of EIs with a balance that «Подобрать ЕИ» found (it shows at most 15)
Global gPickAll As Long

Private Function FindSheet() As Object
    FindSheet = gDoc.Sheets.getByName(SH_FIND)
End Function

Private Sub ClearPage(sh As Object)
    sh.getCellRangeByPosition(1, FD_HEAD, 12, FD_TABLE + FD_CARD_MAX).clearContents(1 + 2 + 4)
    sh.getCellRangeByPosition(1, FD_TABLE, 8, FD_TABLE + FD_CARD_MAX).CharColor = RGB(38, 50, 56)
End Sub

' the columns F G H of the table: a list (quantity, unit, place) or a card (in, out, the running balance in bold)
Private Sub TableMode(sh As Object, bCard As Boolean)
    Dim rng As Object
    rng = sh.getCellRangeByPosition(5, FD_TABLE_HEAD, 5, FD_TABLE + FD_CARD_MAX)
    rng.HoriJustify = 3
    rng = sh.getCellRangeByPosition(6, FD_TABLE_HEAD, 7, FD_TABLE + FD_CARD_MAX)
    rng.HoriJustify = IIf(bCard, 3, 1)
    sh.getCellRangeByPosition(6, FD_TABLE, 7, FD_TABLE + FD_CARD_MAX).CharWeight = 100
    If bCard Then sh.getCellRangeByPosition(7, FD_TABLE, 7, FD_TABLE + FD_CARD_MAX).CharWeight = 150
End Sub

' ================================================================ buttons and events

Sub BtnFind(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = FindRun(Trim(FindSheet().getCellByPosition(FD_INPUT_COL, FD_INPUT_ROW).getString()))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the field changed (the storekeeper typed and pressed Enter): the search runs at once
Sub OnFindChange(Optional oTarget As Variant)
    Dim a As Object
    On Error GoTo EH
    WmsInit()
    If gFdBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    a = oTarget.getRangeAddress()
    If a.StartRow <= FD_INPUT_ROW And a.EndRow >= FD_INPUT_ROW And a.StartColumn <= FD_INPUT_COL And a.EndColumn >= FD_INPUT_COL Then
        FindRun(Trim(FindSheet().getCellByPosition(FD_INPUT_COL, FD_INPUT_ROW).getString()))
    End If
EH:
End Sub

' «Карточка ЕИ»: of the found row under the cursor, or of the EI typed in the field
Sub BtnFindCard(Optional oEvent As Variant)
    Dim r As Long, s As String, n As Long
    WmsInit()
    r = ViewRow(SH_FIND)
    If r >= FD_TABLE Then s = FindSheet().getCellByPosition(FD_SRC_EI, r).getString()
    If s = "" Then s = Trim(FindSheet().getCellByPosition(FD_INPUT_COL, FD_INPUT_ROW).getString())
    n = EiNo(s, CLng(SysNum(SK_NEXT_EI)))
    If n = 0 Then
        WmsUi.UiMessage("Карточка: впишите ЕИ в поле «Что найти» (например 427 или ЕИ-00000427) или встаньте на найденную строку с ЕИ")
        Exit Sub
    End If
    s = CardShow(n)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Перейти к строке»: the source row of the line under the cursor
Sub BtnFindGo(Optional oEvent As Variant)
    Dim r As Long
    WmsInit()
    On Error GoTo EH
    r = ViewRow(SH_FIND)
    If r >= FD_TABLE Then
        If FindSheet().getCellByPosition(FD_SRC_SHEET, r).getString() = "" Then r = -1
    End If
    If r < FD_TABLE Then
        WmsUi.UiMessage("Встаньте на строку найденного (или движение карточки) и нажмите «Перейти к строке» — откроется её лист на этой строке")
        Exit Sub
    End If
    GoSource(r)
    Exit Sub
EH:
    gUiLastMsg = "ERR:" & Error$ & " (строка " & Erl & ")"
End Sub

Private Sub GoSource(r As Long)
    Dim sh As Object
    sh = FindSheet()
    GoCell(sh.getCellByPosition(FD_SRC_SHEET, r).getString(), CLng(sh.getCellByPosition(FD_SRC_ROW, r).getValue()), 1)
    gUiLastMsg = "OK:" & sh.getCellByPosition(FD_SRC_SHEET, r).getString() & " " & (CLng(sh.getCellByPosition(FD_SRC_ROW, r).getValue()) + 1)
End Sub

' a double click on a line: a found row with an EI → its card; a movement of a card → its row. True: no editing of the cell
Function OnFindDoubleClick(Optional oTarget As Variant) As Boolean
    Dim r As Long, sh As Object, s As String
    On Error GoTo EH
    WmsInit()
    r = oTarget.getCellAddress().Row
    If r < FD_TABLE Then Exit Function
    sh = FindSheet()
    OnFindDoubleClick = True
    s = sh.getCellByPosition(FD_SRC_EI, r).getString()
    If gFdMode = "LIST" And s <> "" Then
        CardShow(EiNo(s, CLng(SysNum(SK_NEXT_EI))))
    ElseIf sh.getCellByPosition(FD_SRC_SHEET, r).getString() <> "" Then
        GoSource(r)
    End If
    Exit Function
EH:
End Function

' ================================================================ the search

' sQuery: an EI → its card; otherwise every working sheet is searched. OK:<found> / ERR:<why>
Function FindRun(sQuery As String) As String
    Dim n As Long, sh As Object
    WmsInit()
    If gFdBusy Then Exit Function
    If sQuery = "" Then
        sh = ViewOpen(SH_FIND)
        ClearPage(sh)
        sh.getCellByPosition(1, FD_HEAD).setString("Впишите, что найти, и нажмите Enter")
        ViewDone(sh)
        FindRun = "SKIP:пустой запрос"
        Exit Function
    End If
    gFdQuery = sQuery
    ' «текст*» — always the search of the text over the sheets (not the card of an EI, not «на руках»)
    If Right(sQuery, 1) = "*" And Len(sQuery) > 1 Then
        FindRun = FindList(Trim(Left(sQuery, Len(sQuery) - 1)))
        Exit Function
    End If
    n = EiNo(sQuery, CLng(SysNum(SK_NEXT_EI)))
    If n > 0 Then
        If WmsIssue.StockSheet().getCellByPosition(SC_EI, n).getString() <> "" Then
            FindRun = CardShow(n)
            Exit Function
        End If
    End If
    ' a recipient (the short or the full name of «Получатели», or a name of «Кому» of «Выдачи»): what is on his hands
    If RecipientName(sQuery) <> "" Then
        FindRun = HandsShow(RecipientName(sQuery))
        Exit Function
    End If
    FindRun = FindList(sQuery)
End Function

' the full name of the recipient sQuery (the whole short or full name, any case) or ""
Function RecipientName(sQuery As String) As String
    Dim sh As Object, last As Long, d As Variant, i As Long, q As String, desc As Object, found As Object
    q = LCase(Trim(sQuery))
    If q = "" Then Exit Function
    sh = gDoc.Sheets.getByName(SH_RCPT)
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        d = sh.getCellRangeByPosition(0, 1, 1, last).getDataArray()
        For i = 0 To UBound(d)
            If LCase(Trim(CStr(d(i)(0)))) = q Or LCase(Trim(CStr(d(i)(1)))) = q Then
                RecipientName = IIf(Trim(CStr(d(i)(1))) <> "", Trim(CStr(d(i)(1))), Trim(CStr(d(i)(0))))
                Exit Function
            End If
        Next i
    End If
    ' a name typed in «Кому» that is not in the list
    sh = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    desc = sh.createSearchDescriptor()
    desc.SearchString = sQuery
    desc.SearchCaseSensitive = False
    desc.SearchWords = False
    desc.SearchRegularExpression = False
    found = sh.getCellRangeByPosition(IC_WHO, 1, IC_WHO, last).findFirst(desc)
    If Not IsNull(found) Then
        If LCase(Trim(found.getString())) = q Then RecipientName = Trim(found.getString())
    End If
End Function

' «На руках»: the posted issues to sWho minus the returns of each of them; what is still on his hands first
Function HandsShow(sWho As String) As String
    Dim sh As Object, ish As Object, last As Long, d As Variant, i As Long, r0 As Long, r1 As Long, ret As Object, retQ() As Double, nR As Long
    Dim rows() As Variant, nOut As Long, k As Long, idx As Long, issued As Double, back As Double, nIss As Long, nHand As Long, desc As Object, found As Object
    Dim addrs As Variant, r As Long, q As Double, key() As Double, ord() As Long, tmp() As Variant, onHand As Boolean, k2 As Long
    gFdBusy = True
    On Error GoTo EH
    sh = ViewOpen(SH_FIND)
    ClearPage(sh)
    ' the returns of every issue (posted rows of «Возврат»)
    ret = New Collection
    last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_RETURNS))
    ReDim retQ(IIf(last > 0, last, 0))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = gDoc.Sheets.getByName(SH_RETURNS).getCellRangeByPosition(0, r0, RC_CTL, r1).getDataArray()
        For i = 0 To r1 - r0
            If FindPosted(d(i)(RC_CTL)) And VarType(d(i)(RC_QTY)) = 5 Then
                idx = CollGet(ret, CStr(d(i)(RC_ISSUE)))
                If idx < 0 Then
                    idx = nR
                    CollPut(ret, idx, CStr(d(i)(RC_ISSUE)))
                    nR = nR + 1
                End If
                retQ(idx) = retQ(idx) + d(i)(RC_QTY)
            End If
        Next i
    Next r0
    ' the issues to sWho: the whole cell of «Кому», any case; 1.0 — spaces, dots and «ё» do not part one person in two
    ' («Петров  И.И.» = «Петров И. И.»)
    ish = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(ish)
    ReDim rows(FD_CARD_MAX)
    ReDim key(FD_CARD_MAX)
    If last >= 1 Then
        desc = ish.createSearchDescriptor()
        desc.SearchString = WhoRegex(sWho)
        desc.SearchRegularExpression = True
        desc.SearchCaseSensitive = False
        found = ish.getCellRangeByPosition(IC_WHO, 1, IC_WHO, last).findAll(desc)
        If Not IsNull(found) Then
            addrs = found.getRangeAddresses()
            For k = 0 To UBound(addrs)
                For r = addrs(k).StartRow To addrs(k).EndRow
                    d = ish.getCellRangeByPosition(0, r, IC_LAST, r).getDataArray()(0)
                    If FindPosted(d(IC_CTL)) And VarType(d(IC_NO)) = 5 And VarType(d(IC_QTY)) = 5 Then
                        idx = CollGet(ret, CStr(d(IC_NO)))
                        q = 0
                        If idx >= 0 Then q = retQ(idx)
                        ' 1.0: every issue is counted; the list keeps up to FD_CARD_MAX, and what is still on the hands is never
                        ' pushed out by what came back (a returned row gives its place to one on the hands)
                        nIss = nIss + 1
                        issued = issued + d(IC_QTY)
                        back = back + q
                        onHand = (d(IC_QTY) - q > 0.0000005)
                        If onHand Then nHand = nHand + 1
                        k2 = -1
                        If nOut <= FD_CARD_MAX Then
                            k2 = nOut
                            nOut = nOut + 1
                        ElseIf onHand Then
                            k2 = HandsReturnedSlot(key, nOut)
                        End If
                        If k2 >= 0 Then
                            rows(k2) = Array(ViewDay(DateOf(d(IC_DATE))), "Выдача № " & d(IC_NO), CStr(d(IC_EI)), CStr(d(IC_NAME)), d(IC_QTY), IIf(q > 0, q, ""), _
                                WmsIssue.Round3(d(IC_QTY) - q), CStr(d(IC_UNIT)), "", SH_ISSUES, r, CStr(d(IC_EI)))
                            ' on the hands first, the newest first
                            key(k2) = IIf(onHand, 0, 100000) - DateOf(d(IC_DATE)) + r * 0.0000001
                        End If
                    End If
                Next r
            Next k
        End If
    End If
    ReDim ord(IIf(nOut > 0, nOut - 1, 0))
    For i = 0 To nOut - 1
        ord(i) = i
    Next i
    SortIdx(key, ord, nOut, False)
    gFdMode = "LIST"
    TableMode(sh, True)
    sh.getCellByPosition(1, FD_HEAD).setString("На руках у «" & sWho & "»: " & NumText(nHand) & " " & Plural(nHand, "выдача", "выдачи", "выдач") & " не возвращены")
    sh.getCellByPosition(1, FD_INFO).setString("Всего выдач " & NumText(nIss) & " · выдано " & WmsIssue.QtyText(issued) & " · возвращено " & WmsIssue.QtyText(back) _
        & " (количества разных единиц складываются только для счёта)" & IIf(nIss > nOut, " · в списке " & NumText(nOut) & " из " & NumText(nIss) _
        & ": всё, что на руках, и последние возвращённые", ""))
    sh.getCellByPosition(1, FD_INFO + 1).setString("Сначала — что ещё на руках, новые выше; затем возвращённое (серым). Двойной щелчок — карточка ЕИ. " _
        & "Искать «" & gFdQuery & "» как текст по всем листам — допишите звёздочку: «" & gFdQuery & "*».")
    sh.getCellRangeByPosition(1, FD_TABLE_HEAD, 8, FD_TABLE_HEAD).setDataArray(Array(Array("Дата", "Выдача", "ЕИ", "Наименование", "Выдано", "Возвращено", _
        "На руках", "Ед.")))
    If nOut > 0 Then
        ReDim tmp(nOut - 1)
        For i = 0 To nOut - 1
            tmp(i) = rows(ord(i))
        Next i
        sh.getCellRangeByPosition(1, FD_TABLE, 12, FD_TABLE + nOut - 1).setDataArray(tmp)
        If nHand < nOut Then sh.getCellRangeByPosition(1, FD_TABLE + nHand, 8, FD_TABLE + nOut - 1).CharColor = RGB(150, 150, 150)
    Else
        sh.getCellByPosition(2, FD_TABLE).setString("выдач этому получателю нет")
    End If
    ViewDone(sh)
    gFdBusy = False
    HandsShow = "OK:на руках у " & sWho & ": " & nHand & " из " & nIss & ", " & ViewMs() & " мс"
    Exit Function
EH:
    HandsShow = "ERR:«на руках» не построено: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    ViewDone(sh)
    gFdBusy = False
End Function

' 1.0: the regular expression of a recipient in «Кому»: the words of the name in that order, any spaces, dots or commas
' between and around them, «е» or «ё» («Петров И.И.» finds «петров  и. и.»)
Function WhoRegex(sWho As String) As String
    Dim w As Variant, i As Integer, r As String, x As String
    w = Split(WmsAsk.WhoKey(sWho), " ")
    For i = 0 To UBound(w)
        x = EscRegex(CStr(w(i)))
        x = Replace(x, "е", "[её]")
        r = r & IIf(r <> "", "[\s.,]+", "") & x
    Next i
    WhoRegex = "^[\s.,]*" & r & "[\s.,]*$"
End Function

' 1.0: the regular expression of place key k (5-12-3): its parts as people write them — any separator, leading zeros, «Р5»,
' «C12» in Latin — and nothing glued to it («15-12-3», «5-12-31» are other places)
' 1.0 (М6): what lies on place key now — the EIs of «Наличие» with a balance whose list of places holds it («пусто» when none)
Function PlaceNow(k As String) As String
    Dim st As Object, last As Long, desc As Object, found As Object, a As Variant, i As Long, r As Long, d As Variant, pls As Variant, j As Integer
    Dim s As String, n As Long
    st = WmsIssue.StockSheet()
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last >= 1 Then
        desc = st.createSearchDescriptor()
        desc.SearchString = PlaceRegex(k)
        desc.SearchRegularExpression = True
        desc.SearchCaseSensitive = False
        found = st.getCellRangeByPosition(SC_PLACE, 1, SC_PLACE, last).findAll(desc)
        If Not IsNull(found) Then
            a = found.getRangeAddresses()
            For i = 0 To UBound(a)
                For r = a(i).StartRow To a(i).EndRow
                    d = st.getCellRangeByPosition(0, r, SC_STATE, r).getDataArray()(0)
                    If VarType(d(SC_QTY)) = 5 Then
                        If d(SC_QTY) > 0.0000005 And CStr(d(SC_STATE)) <> EI_ST_STORNO Then
                            pls = PlaceList(CStr(d(SC_PLACE)))
                            For j = 0 To UBound(pls)
                                If pls(j) = k Then
                                    n = n + 1
                                    If n <= 8 Then s = s & IIf(s <> "", "; ", "") & d(SC_EI) & " " & d(SC_NAME) & " — " & WmsPack.QtyWithParts(r, CDbl(d(SC_QTY)), CStr(d(SC_UNIT)))
                                    Exit For
                                End If
                            Next j
                        End If
                    End If
                Next r
            Next i
        End If
    End If
    PlaceNow = IIf(n = 0, "пусто", NumText(n) & " ЕИ: " & s & IIf(n > 8, " …", ""))
End Function

' 1.0 (М5): «Показать на карте» — the goods of the card of «Поиск» (the same article, or the same name) on the map
Sub BtnShowOnMap(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = ShowCardOnMap()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Function ShowCardOnMap() As String
    Dim n As Long, d As Variant, desc As Object, found As Object, a As Variant, i As Long, r As Long, s As String, col As Integer, v As String, st As Object
    If gFdMode <> "CARD" Or gFdCardEI <= 0 Then
        ShowCardOnMap = "ERR:откройте карточку ЕИ на «" & SH_FIND & "» (номер ЕИ в поле поиска) — WMS покажет на карте все места этого товара"
        Exit Function
    End If
    n = gFdCardEI
    st = WmsIssue.StockSheet()
    d = st.getCellRangeByPosition(0, n, SC_ART, n).getDataArray()(0)
    If Trim(CStr(d(SC_ART))) <> "" Then
        col = SC_ART
        v = Trim(CStr(d(SC_ART)))
    Else
        col = SC_NAME
        v = Trim(CStr(d(SC_NAME)))
    End If
    desc = st.createSearchDescriptor()
    desc.SearchString = "^" & EscRegex(v) & "$"
    desc.SearchRegularExpression = True
    desc.SearchCaseSensitive = False
    found = st.getCellRangeByPosition(col, 1, col, CLng(SysNum(SK_NEXT_EI)) - 1).findAll(desc)
    s = CStr(n)
    If Not IsNull(found) Then
        a = found.getRangeAddresses()
        For i = 0 To UBound(a)
            For r = a(i).StartRow To a(i).EndRow
                If r <> n And st.getCellByPosition(SC_QTY, r).getValue() > 0.0000005 Then s = s & "," & r
            Next r
        Next i
    End If
    ShowCardOnMap = WmsMap.ShowOnMap(s, CStr(d(SC_NAME)))
End Function

Function PlaceRegex(k As String) As String
    Dim a As Variant, i As Integer, r As String, x As String
    a = Split(k, "-")
    For i = 0 To UBound(a)
        If IsNumeric(a(i)) Then
            x = "0*" & a(i)
            If i = 0 Then x = "(?:[РRP]\s*)?" & x
            If i = 1 Then x = "(?:[СSC]\s*)?" & x
            If i = 2 Then x = "(?:[ЯL]\s*)?" & x
        Else
            x = EscRegex(CStr(a(i)))
        End If
        r = r & IIf(r <> "", "\s*[-./,:_\\ ]\s*", "") & x
    Next i
    PlaceRegex = "(?:^|[^0-9])" & r & "(?:[^0-9]|$)"
End Function

' a returned row of the list of «на руках» (key ≥ 100000 − day) whose place a row still on the hands takes; -1 when none
Private Function HandsReturnedSlot(key() As Double, n As Long) As Long
    Dim i As Long
    HandsReturnedSlot = -1
    For i = n - 1 To 0 Step -1
        If key(i) > 50000 Then
            HandsReturnedSlot = i
            Exit Function
        End If
    Next i
End Function

' a text for a regular expression of the search (its special characters escaped)
Function EscRegex(s As String) As String
    Dim i As Integer, ch As String, o As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If InStr("\^$.|?*+()[]{}", ch) > 0 Then o = o & "\" & ch Else o = o & ch
    Next i
    EscRegex = o
End Function

' every working sheet searched for sQuery (what the cells show, any case, part of a cell)
Function FindList(sQuery As String) As String
    Dim sh As Object, specs As Variant, i As Integer, out() As Variant, nOut As Long, nAll As Long, totals As String, more As Boolean
    gFdBusy = True
    On Error GoTo EH
    sh = ViewOpen(SH_FIND)
    ClearPage(sh)
    ReDim out(FD_MAX)
    ' sheet | first row | last column
    specs = Array(Array(SH_STOCK, 1, SC_LAST), Array(SH_ORDERS, 1, OC_LAST), Array(SH_ISSUES, 1, IC_LAST), Array(SH_RETURNS, 1, RC_LAST), _
        Array(SH_SPECIAL, 1, XC_LAST), Array(SH_ADJUST, 1, AC_LAST), Array(SH_CARS, CAR_FIRST, CC_LAST), Array(SH_RCPT, 1, 1))
    ' 1.0: a place («1-2-3», «Р1-С2-3», «1.2.3») is found as a whole place: «1-2-3» does not find «11-2-3» or «1-2-31»
    gFdPlaceRx = ""
    If PlaceTextKey(sQuery) <> "" Then gFdPlaceRx = PlaceRegex(PlaceTextKey(sQuery))
    For i = 0 To UBound(specs)
        If gDoc.Sheets.hasByName(specs(i)(0)) Then
            SearchSheet(specs(i)(0), specs(i)(1), specs(i)(2), sQuery, out, nOut, nAll, totals)
        End If
    Next i
    gFdMode = "LIST"
    TableMode(sh, False)
    sh.getCellByPosition(1, FD_HEAD).setString(IIf(nAll = 0, "Ничего не найдено: «" & sQuery & "»", "Найдено «" & sQuery & "»: " & NumText(nAll) _
        & " " & Plural(nAll, "строка", "строки", "строк") & IIf(nAll > nOut, " (показаны первые " & nOut & ")", "")))
    sh.getCellByPosition(1, FD_INFO).setString(IIf(totals <> "", "В «" & SH_STOCK & "» по найденному: " & totals, ""))
    ' 1.0 (М6, М12): a place — what lies on it now; the rows below are its history (receipts, issues, moves through it)
    If gFdPlaceRx <> "" Then
        sh.getCellByPosition(1, FD_HEAD).setString("Место " & PlaceTextKey(sQuery) & ": сейчас — " & PlaceNow(PlaceTextKey(sQuery)) _
            & IIf(nAll > 0, " · история места — строки ниже (" & NumText(nAll) & ")", ""))
    End If
    sh.getCellByPosition(1, FD_INFO + 1).setString(IIf(nAll > 0, "Двойной щелчок по строке с ЕИ — карточка ЕИ; по другой строке — переход к ней. " _
        & "«Перейти к строке» — открыть строку на её листе.", "Ищется в «" & SH_STOCK & "», «" & SH_ORDERS & "», «" & SH_ISSUES & "», «" & SH_RETURNS _
        & "», «" & SH_SPECIAL & "», «" & SH_ADJUST & "», «" & SH_CARS & "», «" & SH_RCPT & "» — по тому, что видно в ячейках, без учёта регистра."))
    sh.getCellRangeByPosition(1, FD_TABLE_HEAD, 8, FD_TABLE_HEAD).setDataArray(Array(Array("Лист", "Строка", "ЕИ", "Наименование", "Количество", _
        "Ед.", "Место", "Подробно")))
    If nOut > 0 Then sh.getCellRangeByPosition(1, FD_TABLE, 12, FD_TABLE + nOut - 1).setDataArray(SubArr(out, nOut))
    ViewDone(sh)
    gFdBusy = False
    FindList = "OK:найдено " & nAll & ", показано " & nOut & ", " & ViewMs() & " мс"
    Exit Function
EH:
    FindList = "ERR:поиск не выполнен: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    ViewDone(sh)
    gFdBusy = False
End Function

Private Function SubArr(a As Variant, n As Long) As Variant
    Dim b() As Variant, i As Long
    ReDim b(n - 1)
    For i = 0 To n - 1
        b(i) = a(i)
    Next i
    SubArr = b
End Function

' the rows of one sheet with sQuery in a cell (native search, values as shown) → out (at most FD_MAX in all)
Private Sub SearchSheet(sName As String, r0 As Long, cLast As Integer, sQuery As String, out() As Variant, nOut As Long, nAll As Long, _
    totals As String)
    Dim sh As Object, last As Long, desc As Object, found As Object, addrs As Variant, k As Long, r As Long, seen() As Boolean, d As Variant
    Dim sum As Double, nSum As Long, unit As String, sameUnit As Boolean, nRead As Long
    sh = gDoc.Sheets.getByName(sName)
    last = WmsOrders.LastRow(sh)
    If last < r0 Then Exit Sub
    desc = sh.createSearchDescriptor()
    If gFdPlaceRx <> "" Then
        desc.SearchString = gFdPlaceRx
        desc.SearchRegularExpression = True
    Else
        desc.SearchString = sQuery
        desc.SearchRegularExpression = False
    End If
    desc.SearchCaseSensitive = False
    desc.SearchWords = False
    desc.setPropertyValue("SearchType", 1)
    found = sh.getCellRangeByPosition(0, r0, cLast, last).findAll(desc)
    If IsNull(found) Then Exit Sub
    addrs = found.getRangeAddresses()
    ' 1.0: a row found in several cells counts once — an array of the rows (a text of the rows seen grew with every row
    ' and stopped the count after ~28 000 rows: a frequent word was slow and «найдено» too small)
    ReDim seen(last)
    For k = 0 To UBound(addrs)
        For r = addrs(k).StartRow To addrs(k).EndRow
            If Not seen(r) Then
                seen(r) = True
                nAll = nAll + 1
                ' every row is counted; only the rows shown (and of «Наличие» up to 5000 — the sum) are read
                If nOut >= FD_MAX And (sName <> SH_STOCK Or nRead >= 5000) Then GoTo NEXT_ROW
                nRead = nRead + 1
                d = sh.getCellRangeByPosition(0, r, cLast, r).getDataArray()(0)
                ' (no short circuit in Basic: the columns of «Наличие» are read only for its rows)
                If sName = SH_STOCK Then
                    If VarType(d(SC_QTY)) = 5 Then
                        If CStr(d(SC_STATE)) <> EI_ST_STORNO Then
                            nSum = nSum + 1
                            If nSum = 1 Then
                                unit = CStr(d(SC_UNIT))
                                sameUnit = True
                            ElseIf CStr(d(SC_UNIT)) <> unit Then
                                sameUnit = False
                            End If
                            sum = sum + d(SC_QTY)
                        End If
                    End If
                End If
                If nOut < FD_MAX Then
                    out(nOut) = RowView(sName, r, d, addrs(k).StartColumn, sh)
                    nOut = nOut + 1
                End If
            End If
NEXT_ROW:
        Next r
    Next k
    ' the balance of the found EIs when they have one unit (otherwise the units differ and a sum means nothing)
    If nSum > 0 Then
        If sameUnit Then
            totals = "ЕИ " & NumText(nSum) & ", остаток " & WmsIssue.QtyText(sum) & " " & unit & IIf(nRead >= 5000, " (по первым 5000 найденным строкам)", "")
        Else
            totals = "ЕИ " & NumText(nSum) & " (единицы разные — общей суммы нет)"
        End If
    End If
End Sub

' one found row as the line of the table: sheet, row, EI, name, quantity, unit, place, details + the hidden source
Private Function RowView(sName As String, r As Long, d As Variant, cHit As Integer, sh As Object) As Variant
    Dim ei As String, nm As String, q As Variant, u As String, pl As String, det As String
    q = ""
    Select Case sName
    Case SH_STOCK
        ei = d(SC_EI): nm = d(SC_NAME): q = d(SC_QTY): u = d(SC_UNIT): pl = d(SC_PLACE)
        det = JoinParts(Array(d(SC_STATE), d(SC_SRC), d(SC_CAT)), " · ")
    Case SH_ORDERS
        ei = d(OC_EI): nm = d(OC_NAME): u = d(OC_UNIT): pl = d(OC_PLACE)
        q = IIf(CStr(d(OC_EI)) <> "", d(OC_FACT), d(OC_ORDQTY))
        det = "заказ " & d(OC_ORDER) & IIf(CStr(d(OC_SUPPLIER)) <> "", " · " & d(OC_SUPPLIER), "") & " · " & d(OC_STATUS) _
            & IIf(CStr(d(OC_EI)) = "" And CStr(d(OC_ORDQTY)) <> "", " (заказано)", "")
    Case SH_ISSUES
        ei = d(IC_EI): nm = d(IC_NAME): q = d(IC_QTY): u = d(IC_UNIT): pl = d(IC_PLACE)
        det = "выдача № " & d(IC_NO) & " · " & ViewDay(DateOf(d(IC_DATE))) & " → " & d(IC_WHO) & " · " & d(IC_CTL)
    Case SH_RETURNS
        ei = d(RC_EI): nm = d(RC_NAME): q = d(RC_QTY): u = d(RC_UNIT): pl = d(RC_PLACE)
        det = "возврат № " & d(RC_NO) & " по выдаче № " & d(RC_ISSUE) & " · " & ViewDay(DateOf(d(RC_DATE))) & " · от " & d(RC_WHO)
    Case SH_SPECIAL
        ei = d(XC_EI): nm = d(XC_NAME): q = d(XC_QTY): u = d(XC_UNIT): pl = d(XC_PLACE)
        det = CStr(d(XC_TYPE)) & " · " & ViewDay(DateOf(d(XC_DATE))) & IIf(CStr(d(XC_DOC)) <> "", " · " & d(XC_DOC), "") & " · " & d(XC_STATUS)
    Case SH_ADJUST
        ei = d(AC_EI): nm = d(AC_NAME): u = d(AC_UNIT): pl = d(AC_PLACE)
        q = IIf(CStr(d(AC_QTY)) <> "", d(AC_QTY), d(AC_DIFF))
        det = CStr(d(AC_KIND)) & " · " & ViewDay(DateOf(d(AC_DATE))) & IIf(CStr(d(AC_REASON)) <> "", " · " & d(AC_REASON), "")
    Case SH_CARS
        nm = "машина " & d(CC_PLATE)
        det = ViewDay(DateOf(d(CC_DATE))) & " · " & d(CC_SUP) & " · " & d(CC_STATUS)
    Case Else
        nm = d(0) & IIf(UBound(d) >= 1, " — " & d(1), "")
        det = "получатель"
    End Select
    RowView = Array(sName, r + 1, CStr(ei), CStr(nm), q, CStr(u), CStr(pl), JoinParts(Array(det, "найдено в «" & sh.getCellByPosition(cHit, IIf(sName = SH_CARS, CAR_HEAD_ROW, 0)).getString() & "»"), " · "), _
        "", sName, r, CStr(ei))
End Function

' a ready list on the page (the map shows its EIs not on the map here): rows in the form of RowView, n of them
Function ShowRows(sTitle As String, sInfo As String, rows As Variant, n As Long) As String
    Dim sh As Object
    gFdBusy = True
    On Error GoTo EH
    sh = ViewOpen(SH_FIND)
    ClearPage(sh)
    gFdMode = "LIST"
    TableMode(sh, False)
    sh.getCellByPosition(1, FD_HEAD).setString(sTitle)
    sh.getCellByPosition(1, FD_INFO).setString(sInfo)
    If n > 0 Then sh.getCellByPosition(1, FD_INFO + 1).setString("Двойной щелчок по строке — карточка ЕИ; «Перейти к строке» — строка в «" & SH_STOCK & "».")
    sh.getCellRangeByPosition(1, FD_TABLE_HEAD, 8, FD_TABLE_HEAD).setDataArray(Array(Array("Лист", "Строка", "ЕИ", "Наименование", "Количество", _
        "Ед.", "Место", "Подробно")))
    If n > 0 Then sh.getCellRangeByPosition(1, FD_TABLE, 12, FD_TABLE + n - 1).setDataArray(SubArr(rows, n))
    ViewDone(sh)
    gFdBusy = False
    ShowRows = "OK:строк " & n
    Exit Function
EH:
    ShowRows = "ERR:" & Error$
    On Error Resume Next
    ViewDone(sh)
    gFdBusy = False
End Function

' ================================================================ 1.0 (М8) the route of picking

Sub BtnPickRoute(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PickRoute()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the order of a place in a walk through the store: ряд, секция, ярус of its first place (a letter part by its letter); the rows
' without a known place go last
Private Function RouteKey(pls As Variant) As Double
    Dim a As Variant, i As Integer, x As Double, w As Variant, v As Double
    RouteKey = 99999999
    If UBound(pls) < 0 Then Exit Function
    a = Split(pls(0), "-")
    w = Array(100000, 100, 1)
    For i = 0 To UBound(a)
        If i > 2 Then Exit For
        If IsNumeric(a(i)) Then v = Val(a(i)) Else v = 900 + (AscW(Left(a(i), 1)) Mod 99)
        If v > IIf(i = 2, 99, 999) Then v = IIf(i = 2, 99, 999)
        x = x + v * w(i)
    Next i
    RouteKey = x
End Function

' the rows of «Выдачи» not posted yet (an EI written, no №) in the order of a walk through the store — by the place of the EI
' (ряд, секция, ярус; the rows without a place at the end): a picking list on «Поиск», ready to print (Ctrl+P). OK:строк N / ERR:…
Function PickRoute() As String
    Dim ish As Object, last As Long, r0 As Long, r1 As Long, d As Variant, i As Long, n As Long, rows() As Variant, key() As Double, ord() As Long
    Dim sEI As String, num As Long, canon As String, msg As String, st As Variant, pls As Variant, pl As String, tmp() As Variant, sh As Object
    Dim nm As String, u As String, q As Variant, nNo As Long, more As Boolean, rw As Variant
    gFdBusy = True
    On Error GoTo EH
    ish = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(ish)
    ReDim rows(FD_CARD_MAX)
    ReDim key(FD_CARD_MAX)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = ish.getCellRangeByPosition(0, r0, IC_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            If Trim(CStr(d(i)(IC_EI))) <> "" And VarType(d(i)(IC_NO)) <> 5 Then
                If n > FD_CARD_MAX Then
                    more = True
                    Exit For
                End If
                sEI = Trim(CStr(d(i)(IC_EI)))
                canon = sEI
                nm = CStr(d(i)(IC_NAME))
                u = CStr(d(i)(IC_UNIT))
                pl = ""
                If WmsIssue.NormalizeEI(sEI, num, canon, msg) Then
                    st = WmsIssue.StockData(num)
                    pl = Trim(CStr(st(SC_PLACE)))
                    nm = CStr(st(SC_NAME))
                    u = CStr(st(SC_UNIT))
                End If
                pls = PlaceList(pl)
                If UBound(pls) < 0 Then nNo = nNo + 1
                q = d(i)(IC_QTY)
                If VarType(q) <> 5 Then q = Trim(CStr(q))
                If Trim(CStr(d(i)(IC_PCT))) <> "" Then q = Trim(CStr(q) & " " & CStr(d(i)(IC_PCT)))
                rows(n) = Array(0, IIf(pl <> "", pl, "место не указано"), canon, nm, q, u, CStr(d(i)(IC_WHO)), "строка " & (r0 + i + 1), "", SH_ISSUES, r0 + i, canon)
                key(n) = RouteKey(pls) * 1000000 + r0 + i
                n = n + 1
            End If
        Next i
        If more Then Exit For
    Next r0
    sh = ViewOpen(SH_FIND)
    ClearPage(sh)
    gFdMode = "LIST"
    TableMode(sh, False)
    sh.getCellByPosition(1, FD_HEAD).setString("Маршрут сбора: " & NumText(n) & " " & Plural(n, "строка", "строки", "строк") & " «" & SH_ISSUES & "» ещё не проведены")
    sh.getCellByPosition(1, FD_INFO).setString("По порядку обхода склада: ряд, секция, ярус места ЕИ" & IIf(nNo > 0, "; без места на карте — в конце (" & nNo & ")", "") _
        & IIf(more, "; в списке первые " & NumText(FD_CARD_MAX + 1), "") & ". Собрали — проведите строки на листе «" & SH_ISSUES & "».")
    sh.getCellByPosition(1, FD_INFO + 1).setString("Напечатать: Ctrl+P (печатается этот лист). Двойной щелчок по строке — карточка ЕИ; «Перейти к строке» — строка выдачи.")
    sh.getCellRangeByPosition(1, FD_TABLE_HEAD, 8, FD_TABLE_HEAD).setDataArray(Array(Array("№", "Место", "ЕИ", "Наименование", "Количество", _
        "Ед.", "Кому", "Выдачи")))
    If n > 0 Then
        ReDim ord(n - 1)
        For i = 0 To n - 1
            ord(i) = i
        Next i
        SortIdx(key, ord, n, False)
        ReDim tmp(n - 1)
        For i = 0 To n - 1
            rw = rows(ord(i))
            rw(0) = i + 1
            tmp(i) = rw
        Next i
        sh.getCellRangeByPosition(1, FD_TABLE, 12, FD_TABLE + n - 1).setDataArray(tmp)
    Else
        sh.getCellByPosition(2, FD_TABLE).setString("непроведённых выдач нет — собирать нечего")
    End If
    ViewDone(sh)
    gFdBusy = False
    PickRoute = "OK:строк " & n
    gUiLastMsg = PickRoute
    Exit Function
EH:
    PickRoute = "ERR:маршрут не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = PickRoute
    On Error Resume Next
    ViewDone(sh)
    gFdBusy = False
End Function

' ================================================================ the card of an EI

' the card of EI n: who it came from, the movements with the running balance, the check against «Наличие». OK:… / ERR:…
Function CardShow(n As Long) As String
    Dim sh As Object, st As Variant, canon As String, mKey() As Double, mRow() As Variant, nM As Long, idx() As Long, i As Long
    Dim bal As Double, rows() As Variant, sumIn As Double, sumOut As Double, q As Variant, ok As Boolean, lg As Variant, rv As Variant
    Dim hasSrc As Boolean, hasPack As Boolean, pu As String, pn As Double, note As String
    gFdBusy = True
    On Error GoTo EH
    sh = ViewOpen(SH_FIND)
    ClearPage(sh)
    st = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
    canon = CStr(st(SC_EI))
    gFdMode = "CARD"
    gFdCardEI = IIf(canon <> "", n, 0)
    If canon = "" Then
        sh.getCellByPosition(1, FD_HEAD).setString("ЕИ-" & Format(n, "00000000") & " нет в «" & SH_STOCK & "»")
        ViewDone(sh)
        gFdBusy = False
        CardShow = "ERR:ЕИ-" & Format(n, "00000000") & " нет в «" & SH_STOCK & "»"
        Exit Function
    End If
    ReDim mKey(FD_CARD_MAX)
    ReDim mRow(FD_CARD_MAX)
    lg = WmsLegacy.LegSheet().getCellRangeByPosition(0, n, LG_LAST, n).getDataArray()(0)
    rv = WmsOrders.RcvSheet().getCellRangeByPosition(0, n, RV_LAST, n).getDataArray()(0)
    ' the transfer without an order: the balance it came with (no date: first)
    If CStr(lg(LG_KIND)) = LG_MIGRATE Then
        AddMove(mKey, mRow, nM, 0, 0, "", "Перенос из старой таблицы (без заказа)", CStr(st(SC_SRC)), "остаток на момент переноса", lg(LG_BAL0), 0, _
            "", -1, IIf(CDbl(Val(CStr(lg(LG_FIXES)))) > 0, "исправлено: " & lg(LG_ACT), ""))
    End If
    ' the receipt by an order (the row of «Заказы» in _RCV)
    If CStr(rv(RV_EI)) <> "" Then OrderReceipt(rv, lg, mKey, mRow, nM)
    ' every row of the EI on the sheets of movements
    gFdSpecial = False
    SheetMoves(SH_SPECIAL, XC_EI, XC_LAST, canon, mKey, mRow, nM)
    SheetMoves(SH_ISSUES, IC_EI, IC_LAST, canon, mKey, mRow, nM)
    SheetMoves(SH_RETURNS, RC_EI, RC_LAST, canon, mKey, mRow, nM)
    SheetMoves(SH_ADJUST, AC_EI, AC_LAST, canon, mKey, mRow, nM)
    ' 1.0 (М1): an EI split from another one starts with that split
    If SplitOrigin(canon, mKey, mRow, nM) Then gFdSpecial = True
    ' a source of the balance: a receipt by an order, a transfer, a special receipt (the synthetic EIs of a test book have none)
    hasSrc = (CStr(rv(RV_EI)) <> "" Or CStr(lg(LG_KIND)) <> "" Or gFdSpecial)
    ReDim idx(IIf(nM > 0, nM - 1, 0))
    For i = 0 To nM - 1
        idx(i) = i
    Next i
    SortIdx(mKey, idx, nM, False)
    ReDim rows(IIf(nM > 0, nM - 1, 0))
    hasPack = WmsPack.PackOf(n, pu, pn)
    For i = 0 To nM - 1
        q = mRow(idx(i))
        bal = WmsIssue.Round3(bal + q(4) - q(5))
        sumIn = sumIn + q(4)
        sumOut = sumOut + q(5)
        note = q(8)
        ' D-099: the movement in the parts of the packing too («1 кан»)
        If hasPack And (q(4) <> 0 Or q(5) <> 0) Then note = WmsPack.PartsText(WmsPack.PartsOf(IIf(q(4) <> 0, q(4), q(5)), pn), pu) & IIf(note <> "", " · " & note, "")
        rows(i) = Array(q(0), q(1), q(2), q(3), IIf(q(4) <> 0, q(4), ""), IIf(q(5) <> 0, q(5), ""), bal, note, "", q(6), q(7), "")
    Next i
    q = st(SC_QTY)
    ok = (VarType(q) = 5)
    If ok Then ok = (Abs(WmsIssue.Round3(q - bal)) < 0.0000005)
    sh.getCellByPosition(1, FD_HEAD).setString(canon & " · " & st(SC_NAME))
    sh.getCellByPosition(1, FD_INFO).setString(JoinParts(Array("Остаток " & WmsPack.QtyWithParts(n, IIf(VarType(q) = 5, q, 0), CStr(st(SC_UNIT))), _
        WmsPack.PackLine(n, CStr(st(SC_UNIT))), "место " & IIf(CStr(st(SC_PLACE)) <> "", st(SC_PLACE), "не указано"), st(SC_STATE), _
        IIf(CStr(st(SC_ART)) <> "", "артикул " & st(SC_ART), ""), st(SC_CAT)), "   ·   "))
    sh.getCellByPosition(1, FD_INFO + 1).setString(JoinParts(Array("Откуда: " & IIf(CStr(st(SC_SRC)) <> "", st(SC_SRC), "—") _
        & IIf(CStr(st(SC_STYPE)) <> "", " (" & st(SC_STYPE) & ")", ""), "пришло " & WmsIssue.QtyText(sumIn), "ушло " & WmsIssue.QtyText(sumOut), _
        "движений " & nM), "   ·   "))
    If Not hasSrc Then
        sh.getCellByPosition(1, FD_INFO + 2).setString("•  Первичного прихода в книге нет (ЕИ начальных данных) — сверка остатка с движениями не выполняется")
    ElseIf ok Then
        sh.getCellByPosition(1, FD_INFO + 2).setString("✓ Остаток сходится: сумма движений " & WmsIssue.QtyText(bal) & " = «" & SH_STOCK & "» " _
            & WmsIssue.QtyText(q) & " " & st(SC_UNIT))
    Else
        sh.getCellByPosition(1, FD_INFO + 2).setString("✗ Остаток не сходится: по движениям " & WmsIssue.QtyText(bal) & ", в «" & SH_STOCK & "» " & q _
            & " — пересчитайте ЕИ на месте («" & SH_ADJUST & "» → «Инвентаризация»), подробности — «" & SH_LEGREP & "»")
    End If
    TableMode(sh, True)
    sh.getCellRangeByPosition(1, FD_TABLE_HEAD, 8, FD_TABLE_HEAD).setDataArray(Array(Array("Дата", "Операция", "Документ / №", "Кто / откуда / причина", _
        "Пришло", "Ушло", "Остаток", "Примечание")))
    If nM > 0 Then sh.getCellRangeByPosition(1, FD_TABLE, 12, FD_TABLE + nM - 1).setDataArray(rows)
    If nM = 0 Then sh.getCellByPosition(2, FD_TABLE).setString("движений нет")
    For i = 0 To nM - 1
        If mRow(idx(i))(9) Then sh.getCellRangeByPosition(1, FD_TABLE + i, 8, FD_TABLE + i).CharColor = RGB(150, 150, 150)
    Next i
    sh.getCellByPosition(FD_SRC_EI, FD_HEAD).setString(canon)
    ViewDone(sh)
    gFdBusy = False
    CardShow = "OK:" & canon & ": движений " & nM & ", остаток по движениям " & WmsIssue.QtyText(bal) & IIf(ok, " = ", " ≠ ") & "«Наличие», " & ViewMs() & " мс"
    Exit Function
EH:
    CardShow = "ERR:карточка не построена: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    ViewDone(sh)
    gFdBusy = False
End Function

' a movement: key (date + order within the day), date text, operation, document, who, in, out, sheet, row, note; bGray — no effect
Private Sub AddMove(mKey() As Double, mRow() As Variant, nM As Long, dt As Double, rank As Integer, sDate As String, sOp As String, sDoc As String, _
    sWho As String, vIn As Variant, vOut As Variant, sSheet As String, r As Long, sNote As String, Optional bGray As Variant)
    Dim a As Double, b As Double, g As Boolean
    If nM > FD_CARD_MAX Then Exit Sub
    If VarType(vIn) = 5 Then a = vIn
    If VarType(vOut) = 5 Then b = vOut
    If Not IsMissing(bGray) Then g = bGray
    mKey(nM) = dt + rank * 0.01 + IIf(r > 0, r, 0) * 0.000000001
    mRow(nM) = Array(IIf(sDate <> "", sDate, ViewDay(dt)), sOp, sDoc, sWho, a, b, sSheet, IIf(r >= 0, r, ""), sNote, g)
    nM = nM + 1
End Sub

' the receipt of an order: its row of «Заказы»; a receipt of the transfer also shows what was used before the transfer
Private Sub OrderReceipt(rv As Variant, lg As Variant, mKey() As Double, mRow() As Variant, nM As Long)
    Dim r As Long, d As Variant, dt As Double, live As Boolean, isLeg As Boolean, pre As Double, doc As String
    r = CLng(rv(RV_ROW))
    d = WmsOrders.OrdersSheet().getCellRangeByPosition(0, r, OC_LAST, r).getDataArray()(0)
    dt = DateOf(d(OC_RDATE))
    live = (CStr(rv(RV_STATE)) = RV_LIVE)
    isLeg = (CStr(lg(LG_KIND)) = LG_RECEIPT Or CStr(lg(LG_KIND)) = LG_ADD)
    doc = CStr(d(OC_DOC)) & IIf(CStr(d(OC_INVOICE)) <> "", " · счёт " & d(OC_INVOICE), "")
    AddMove(mKey, mRow, nM, dt, 1, "", IIf(CStr(rv(RV_KIND)) = RV_ADD, "Ещё поступление", "Приход по заказу") & " № " & d(OC_ORDER) _
        & IIf(isLeg, " (перенос)", ""), doc, CStr(d(OC_SUPPLIER)), IIf(live, d(OC_FACT), 0), 0, SH_ORDERS, r, _
        IIf(live, IIf(isLeg And CDbl(Val(CStr(lg(LG_FIXES)))) > 0, "исправлено: " & lg(LG_ACT), ""), "приход удалён (сторно)"), Not live)
    If isLeg And live Then
        If VarType(lg(LG_PRE)) = 5 Then pre = lg(LG_PRE)
        If pre > 0.0000005 Then AddMove(mKey, mRow, nM, dt, 2, "", "Израсходовано до переноса", "старая таблица", "", 0, pre, SH_ORDERS, r, _
            "по старой таблице остаток на момент переноса " & WmsIssue.QtyText(CDbl(lg(LG_BAL0))))
    End If
End Sub

' the rows of the EI canon on a sheet of movements (found by the EI column, the whole cell)
Private Sub SheetMoves(sName As String, cEI As Integer, cLast As Integer, canon As String, mKey() As Double, mRow() As Variant, nM As Long)
    Dim sh As Object, last As Long, desc As Object, found As Object, addrs As Variant, k As Long, r As Long, d As Variant, dt As Double
    Dim live As Boolean, ctl As String
    If Not gDoc.Sheets.hasByName(sName) Then Exit Sub
    sh = gDoc.Sheets.getByName(sName)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    desc = sh.createSearchDescriptor()
    desc.SearchString = "^" & canon & "$"
    desc.SearchRegularExpression = True
    desc.SearchCaseSensitive = True
    found = sh.getCellRangeByPosition(cEI, 1, cEI, last).findAll(desc)
    If IsNull(found) Then Exit Sub
    addrs = found.getRangeAddresses()
    For k = 0 To UBound(addrs)
        For r = addrs(k).StartRow To addrs(k).EndRow
            d = sh.getCellRangeByPosition(0, r, cLast, r).getDataArray()(0)
            Select Case sName
            Case SH_ISSUES
                ctl = CStr(d(IC_CTL))
                If ctl <> "" And VarType(d(IC_NO)) = 5 Then
                    live = FindPosted(ctl)
                    dt = DateOf(d(IC_DATE))
                    AddMove(mKey, mRow, nM, dt, 5, "", "Выдача № " & d(IC_NO), CStr(d(IC_DOC)), "→ " & d(IC_WHO), 0, IIf(live, d(IC_QTY), 0), sName, r, _
                        IIf(live, CStr(d(IC_NOTE)), "удалена (сторно)"), Not live)
                End If
            Case SH_RETURNS
                ctl = CStr(d(RC_CTL))
                If ctl <> "" And VarType(d(RC_NO)) = 5 Then
                    live = FindPosted(ctl)
                    dt = DateOf(d(RC_DATE))
                    AddMove(mKey, mRow, nM, dt, 6, "", "Возврат № " & d(RC_NO), "по выдаче № " & d(RC_ISSUE), "← " & d(RC_WHO), IIf(live, d(RC_QTY), 0), 0, _
                        sName, r, IIf(live, CStr(d(RC_NOTE)), "удалён (сторно)"), Not live)
                End If
            Case SH_SPECIAL
                ctl = CStr(d(XC_STATUS))
                If ctl <> "" Then
                    live = FindPosted(ctl)
                    If live Then gFdSpecial = True
                    dt = DateOf(d(XC_DATE))
                    AddMove(mKey, mRow, nM, dt, 3, "", CStr(d(XC_TYPE)) & IIf(CStr(d(XC_EVENT)) <> "", " № " & d(XC_EVENT), ""), CStr(d(XC_DOC)), _
                        CStr(d(XC_WHO)), IIf(live, d(XC_QTY), 0), 0, sName, r, IIf(live, ctl, "удалено (сторно)"), Not live)
                End If
            Case SH_ADJUST
                ctl = CStr(d(AC_CTL))
                If ctl <> "" And VarType(d(AC_NO)) = 5 Then
                    live = FindPosted(ctl)
                    dt = DateOf(d(AC_DATE))
                    AdjMove(d, live, dt, sName, r, mKey, mRow, nM)
                End If
            End Select
        Next r
    Next k
End Sub

' 1.0 (М1): the split that made EI canon (_ADJ L = canon): its first movement. True when there is one
Private Function SplitOrigin(canon As String, mKey() As Double, mRow() As Variant, nM As Long) As Boolean
    Dim sh As Object, last As Long, desc As Object, found As Object, a As Variant, r As Long, d As Variant, live As Boolean
    If Not gDoc.Sheets.hasByName(SH_ADJ) Then Exit Function
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
    r = a(0).StartRow
    d = sh.getCellRangeByPosition(0, r, AJ_LAST, r).getDataArray()(0)
    If CStr(d(AJ_KIND)) <> "SPLIT" Then Exit Function
    live = (CStr(d(AJ_STATE)) = RV_LIVE)
    AddMove(mKey, mRow, nM, DateOf(d(AJ_DATE)), 1, "", "Отделён от " & d(AJ_EI) & " (корректировка № " & d(AJ_NO) & ")", "", "", _
        IIf(live And VarType(d(AJ_QTY)) = 5, -d(AJ_QTY), 0), 0, SH_ADJUST, -1, IIf(live, "место «" & d(AJ_TO) & "»", "разделение удалено (сторно)"), Not live)
    SplitOrigin = True
End Function

Private Sub AdjMove(d As Variant, live As Boolean, dt As Double, sName As String, r As Long, mKey() As Double, mRow() As Variant, nM As Long)
    Dim kind As String, ch As Double, note As String
    kind = CStr(d(AC_KIND))
    If VarType(d(AC_BEFORE)) = 5 And VarType(d(AC_AFTER)) = 5 Then ch = WmsIssue.Round3(d(AC_AFTER) - d(AC_BEFORE))
    If kind = "Перемещение" Then
        note = "место " & IIf(CStr(d(AC_FROM)) <> "", d(AC_FROM), "—") & " → " & d(AC_PLACE)
        ch = 0
    ElseIf kind = "Инвентаризация" Then
        note = "учётный " & d(AC_BOOK) & ", фактический " & d(AC_FACT)
    ElseIf kind = "Разделение" Then
        note = Mid(CStr(d(AC_CTL)), InStr(CStr(d(AC_CTL)), "·") + 2) & " · место «" & d(AC_PLACE) & "»"
    End If
    If Not live Then
        ch = 0
        note = "удалена (сторно)"
    End If
    AddMove(mKey, mRow, nM, dt, 7, "", kind & " № " & d(AC_NO), CStr(d(AC_BATCH)), CStr(d(AC_REASON)), IIf(ch > 0, ch, 0), IIf(ch < 0, -ch, 0), sName, r, _
        note, Not live Or (kind = "Перемещение"))
End Sub

' the non-empty parts joined by sep
Private Function JoinParts(parts As Variant, sep As String) As String
    Dim i As Integer, s As String
    For i = 0 To UBound(parts)
        If Trim(CStr(parts(i))) <> "" Then s = s & IIf(s <> "", sep, "") & Trim(CStr(parts(i)))
    Next i
    JoinParts = s
End Function

Private Function FindPosted(v As Variant) As Boolean
    FindPosted = (Left(CStr(v), Len(ST_POSTED)) = ST_POSTED)
End Function

' ================================================================ «Подобрать ЕИ» («Выдачи», «Корректировки»)
' The storekeeper knows the goods, not the number of the EI: a name, an article or a place → the EIs with a balance (the oldest
' first) → the chosen EI is written into the EI cell of the row under the cursor (an input cell of a row not yet posted).

Sub BtnPickEI(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PickEI()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Function PickEI() As String
    Dim sName As String, r As Long, sh As Object, cEI As Integer, cNo As Integer, v(0) As String, lines As Variant, opts() As String, c As Integer
    Dim s As String, n As Integer, k As Integer
    sName = gDoc.getCurrentController().getActiveSheet().Name
    If sName = SH_ISSUES Then
        cEI = IC_EI
        cNo = IC_NO
    ElseIf sName = SH_ADJUST Then
        cEI = AC_EI
        cNo = AC_NO
    Else
        PickEI = "ERR:«Подобрать ЕИ» работает на листах «" & SH_ISSUES & "» и «" & SH_ADJUST & "»"
        Exit Function
    End If
    r = ViewRow(sName)
    If r < 1 Then
        PickEI = "ERR:встаньте на строку, в которую нужен ЕИ, и нажмите «Подобрать ЕИ» снова"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(sName)
    If sh.getCellByPosition(cNo, r).getString() <> "" Then
        PickEI = "ERR:строка " & (r + 1) & " уже проведена — ЕИ в ней меняется кнопкой «Исправить»"
        Exit Function
    End If
    v(0) = ""
    If Not WmsPrimeUi.UiForm("Подобрать ЕИ", "Впишите наименование (или его часть), артикул или место — WMS покажет ЕИ с остатком, старые первыми.", _
        Array("Что ищем"), v, Array(""), "Найти") Then
        gUiLastMsg = "SKIP:окно закрыто"
        PickEI = gUiLastMsg
        Exit Function
    End If
    s = PickCandidates(Trim(v(0)))
    If Left(s, 4) = "ERR:" Then
        PickEI = s
        Exit Function
    End If
    lines = Split(s, Chr(10))
    n = UBound(lines) + 1
    ReDim opts(n - 1)
    For k = 0 To n - 1
        opts(k) = Mid(lines(k), InStr(lines(k), "|") + 1)
    Next k
    c = WmsPrimeUi.UiChoose("Подобрать ЕИ — «" & v(0) & "»", "ЕИ с остатком (старые первыми — их и выдавать первыми)" _
        & IIf(gPickAll > n, ". Показано " & n & " из " & gPickAll & " — чтобы найти нужный быстрее, уточните: добавьте слово, артикул или место", "") _
        & ". Выберите и нажмите «Выбрать»:", opts)
    If c < 0 Or c >= n Then
        gUiLastMsg = "SKIP:ЕИ не выбран"
        PickEI = gUiLastMsg
        Exit Function
    End If
    sh.getCellByPosition(cEI, r).setString(Left(lines(c), InStr(lines(c), "|") - 1))
    ' 1.0 (Р10): the row shows what it will do at once, and the cursor goes to the quantity
    If sName = SH_ISSUES Then
        WmsIssue.PreviewRow(r, True, False)
        GoCell(sName, r, IC_QTY)
    Else
        GoCell(sName, r, AC_QTY)
    End If
    gUiLastMsg = "OK:" & Left(lines(c), InStr(lines(c), "|") - 1) & " записан в строку " & (r + 1) & IIf(gPickAll > n, "; показано " & n & " из " & gPickAll, "")
    PickEI = gUiLastMsg
End Function

' EIs with a balance whose name, article or place holds sWhat: lines «EI|EI · name · balance unit · place», the oldest first, at most 15
Function PickCandidates(sWhat As String) As String
    Dim sh As Object, last As Long, desc As Object, found As Object, addrs As Variant, k As Long, r As Long, d As Variant, seen() As Integer, out As String
    Dim nOut As Integer, i As Long, words As Variant, w As Integer, nWords As Integer, nAll As Long, sKey As String
    If Trim(sWhat) = "" Then
        PickCandidates = "ERR:ничего не вписано — впишите наименование, артикул или место"
        Exit Function
    End If
    sh = WmsIssue.StockSheet()
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last < 1 Then
        PickCandidates = "ERR:в «" & SH_STOCK & "» ещё нет ЕИ"
        Exit Function
    End If
    ' 1.0 (Р10): the words in any order — «болт м8» finds «Болт оцинкованный М8х30»: a row counts when every word is in its
    ' name, article or place
    words = Split(Trim(Replace(sWhat, Chr(160), " ")), " ")
    ReDim seen(last)
    For w = 0 To UBound(words)
        If Trim(words(w)) <> "" Then
            nWords = nWords + 1
            desc = sh.createSearchDescriptor()
            desc.SearchString = Trim(words(w))
            desc.SearchCaseSensitive = False
            desc.SearchWords = False
            desc.SearchRegularExpression = False
            ' the name, the article (B, C) and the place (F); a row is counted once per word
            sKey = "|"
            For i = 0 To 1
                If i = 0 Then found = sh.getCellRangeByPosition(SC_NAME, 1, SC_ART, last).findAll(desc) Else found = sh.getCellRangeByPosition(SC_PLACE, 1, SC_PLACE, last).findAll(desc)
                If Not IsNull(found) Then
                    addrs = found.getRangeAddresses()
                    For k = 0 To UBound(addrs)
                        For r = addrs(k).StartRow To addrs(k).EndRow
                            If seen(r) = nWords - 1 Then seen(r) = nWords
                        Next r
                    Next k
                End If
            Next i
        End If
    Next w
    ' the oldest EI first (the row is the number of the EI); all the matching EIs with a balance are counted
    For r = 1 To last
        If seen(r) = nWords And nWords > 0 Then
            d = sh.getCellRangeByPosition(0, r, SC_STATE, r).getDataArray()(0)
            If VarType(d(SC_QTY)) = 5 Then
                If d(SC_QTY) > 0.0000005 And CStr(d(SC_STATE)) <> EI_ST_STORNO Then
                    nAll = nAll + 1
                    If nOut < 15 Then
                        out = out & IIf(out <> "", Chr(10), "") & d(SC_EI) & "|" & d(SC_EI) & " · " & d(SC_NAME) & " · " _
                            & WmsPack.QtyWithParts(r, CDbl(d(SC_QTY)), CStr(d(SC_UNIT))) & IIf(CStr(d(SC_PLACE)) <> "", " · " & d(SC_PLACE), "")
                        nOut = nOut + 1
                    End If
                End If
            End If
        End If
    Next r
    gPickAll = nAll
    If out = "" Then
        PickCandidates = "ERR:ЕИ с остатком по «" & sWhat & "» нет — проверьте написание или поищите на странице «" & SH_FIND & "»"
    Else
        PickCandidates = out
    End If
End Function

' the EI of the card shown on «Поиск» ("" — the page shows no card)
Function FindCardEI() As String
    If gFdMode = "CARD" And gFdCardEI > 0 Then FindCardEI = WmsIssue.EiCanon(gFdCardEI)
End Function

' ================================================================ test seam

Function TestFind(sQuery As String) As String
    TestFind = FindRun(sQuery)
End Function
