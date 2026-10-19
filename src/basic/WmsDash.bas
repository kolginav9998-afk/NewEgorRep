' WmsDash — «Дашборд»: the warehouse today on one page (PRIME 0.8, D-094). Six tiles (на складе, ожидается, пришло и
' выдано сегодня, машины, внимание), the positions whose expected date has come, the last operations of the journal, what is
' issued most in 30 days, what lies without movement, the last 14 days. Everything is counted from the sheets of the book
' when the page is shown (and the data changed) or by «Обновить»; nothing is written anywhere but this page.
Option Explicit

' the cells of the page (0-based; tools/build_ods.py build_dash draws the same layout)
Public Const DB_TITLE_ROW = 0
Public Const DB_TILE_CAP = 3
Public Const DB_TILE_VAL = 4
Public Const DB_TILE_SUB = 5
Public Const DB_LIST1_HEAD = 7
Public Const DB_LIST1 = 9
Public Const DB_LIST1_N = 12
Public Const DB_LIST2_HEAD = 22
Public Const DB_LIST2 = 24
Public Const DB_LIST2_N = 10
' 1.0: the stale EIs whose last movement is looked up (in groups of 50) to show the ten lying longest (the full list is a page)
Public Const DB_STALE_CAND = 100
Public Const DB_DAYS_HEAD = 35
Public Const DB_DAYS = 37
Public Const DB_DAYS_N = 14
' 1.0 (Ч5): the tails — an EI with a balance below one part of its packing (0,01 бут, an offcut of 0,3 м) or, without a
' packing, below 0,01 of the unit of its document
Public Const DB_TAILS_HEAD = 54
Public Const DB_TAILS = 56
Public Const DB_TAILS_N = 10
Public Const DB_FOOT = 67
' hidden service cells of column O: the state code, the LAST_SEQ, the day and the time (ms) of the last refresh
Public Const DB_SVC_COL = 14
Public Const DB_STALE_DAYS = 90
' a page counted longer than this is not counted again by itself when it is shown (a big book): «Обновить» counts it
Public Const DB_AUTO_MAX_MS = 5000
' 1.0 (А7): the rows of the list «Внимание» (as many as «Поиск» shows)
Public Const DB_AT_MAX = 500

Global gDbBusy As Boolean
' 1.0: the state line of the page while a save writes the file (DashSavedState)
Global gDbStateLine As String

Private Sub PutRow(sh As Object, r As Long, c0 As Integer, v As Variant)
    sh.getCellRangeByPosition(c0, r, c0 + UBound(v), r).setDataArray(Array(v))
End Sub

' «Обновить» on «Дашборд»
Sub BtnDashRefresh(Optional oEvent As Variant)
    Dim s As String
    s = DashRefresh()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the page was shown: counted again when an operation was posted since the last count or the day changed — by itself only
' while a count is quick; otherwise the line at the bottom says the figures are older than the last operations
Sub OnDashFocus(Optional oEvent As Variant)
    Dim sh As Object, ms As Double
    On Error GoTo EH
    WmsInit()
    If gDbBusy Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_DASH)
    If CStr(sh.getCellByPosition(DB_SVC_COL, 1).getValue()) = CStr(SysNum(SK_LAST_SEQ)) And sh.getCellByPosition(DB_SVC_COL, 2).getValue() = Today0() _
        And sh.getCellByPosition(DB_SVC_COL, 0).getString() = gState & "|" & gBlock Then Exit Sub
    ms = sh.getCellByPosition(DB_SVC_COL, 3).getValue()
    If ms > DB_AUTO_MAX_MS And sh.getCellByPosition(DB_SVC_COL, 2).getValue() = Today0() Then
        sh = ViewOpen(SH_DASH)
        If Left(sh.getCellByPosition(1, DB_FOOT).getString(), 1) <> "⚠" Then
            sh.getCellByPosition(1, DB_FOOT).setString("⚠ После подсчёта были новые операции — цифры могут отставать. Книга большая (подсчёт около " _
                & Format(ms / 1000, "0") & " с), поэтому страница не пересчитывается сама: нажмите «Обновить»")
        End If
        ViewDone(sh)
        Exit Sub
    End If
    DashRefresh()
EH:
End Sub

' 1.0 (З3): what is still expected of the position of row r of «Заказы» (ordered − received; the whole order before a receipt)
Private Function DashRest(r As Long) As Double
    Dim osh As Object, canon As String, n As Long, rv As Variant, od As Variant, q As Double, msg As String
    On Error GoTo EH
    osh = WmsOrders.OrdersSheet()
    canon = osh.getCellByPosition(OC_EI, r).getString()
    If canon = "" Then
        If WmsOrders.ParseQtyAs(osh.getCellByPosition(OC_ORDQTY, r), "заказано", q, msg) Then DashRest = q
    ElseIf WmsOrders.StrictEI(canon, n) Then
        If WmsOrders.RcvRegistered(n, rv) Then
            If WmsOrders.OlValid(CLng(rv(RV_OL)), od) Then
                If CStr(od(OD_CANCEL)) = "" Then DashRest = WmsIssue.Round3(od(OD_ORD) - od(OD_RCV))
            End If
        End If
    End If
    If DashRest < 0 Then DashRest = 0
EH:
End Function

' Counts and writes the page. OK:<what was counted, ms> or ERR:<why>. The sheets are counted by formulas of the Calc
' engine (WmsView.Eng): only the rows they find — the positions we wait for, the issues of 30 days — are read by Basic.
Function DashRefresh() As String
    Dim sh As Object, today As Double, nextEI As Long, last As Long, i As Long, k As Long, n As Long, s As String, a As Variant, nItems As Long
    Dim nEI As Long, nPos As Long, nReview As Long, nWait As Long, nLate As Long, nToday As Long, nNoDate As Long, wKey(DB_LIST1_N - 1) As Double, nW As Long
    Dim nNeed As Long, rest As Double
    Dim rcvToday As Long, splToday As Long, issToday As Long, retToday As Long, adjToday As Long, carsIn As Long, carsArr As Long, carsDep As Long
    Dim who As Object, nWho As Long, dayIss(13) As Long, dayRcv(13) As Long, dt As Double, key As Double, r As Long
    Dim top As Object, tName() As String, tCnt() As Double, tQty() As Double, tUnit() As String, nTop As Long, j As Long, nm As String
    Dim unposted As Long, attention As Long, recent() As Boolean, nSt As Long, stN(DB_STALE_CAND - 1) As Long, nStShown As Long, stDay As Variant
    Dim osh As Object, rows(), mx As Double, ord As Variant, tIdx() As Long, d As Variant, cond As String, dq As String, rg As String, rd As String
    Dim row As Variant, cD As Integer, cW As Integer, cN As Integer, cU As Integer, cQ As Integer, nList As Long, stKey() As Double, stOrd() As Long
    Dim tails As Variant, nTails As Long, tRows(), money As Variant
    If gDbBusy Then Exit Function
    gDbBusy = True
    On Error GoTo EH
    sh = ViewOpen(SH_DASH)
    gVwEngFail = ""
    today = Today0()
    nextEI = CLng(SysNum(SK_NEXT_EI))
    ' ---------------------------------------------------------------- «Наличие» (row = EI)
    last = nextEI - 1
    If last >= 1 Then
        rg = ColRng(SH_STOCK, SC_QTY, 1, last)
        rd = ColRng(SH_STOCK, SC_STATE, 1, last)
        a = Split(Eng("SUMPRODUCT((" & ColRng(SH_STOCK, SC_EI, 1, last) & "<>"""")*1)&""|""&SUMPRODUCT(ISNUMBER(" & rg & ")*(" & rg & ">0.0000005)*(" _
            & rd & "<>""" & EI_ST_STORNO & """))&""|""&SUMPRODUCT((" & rd & "=""" & EI_ST_REVIEW & """)*1)"), "|")
        nEI = ListNum(a, 0)
        nPos = ListNum(a, 1)
        nReview = ListNum(a, 2)
    End If
    ' ---------------------------------------------------------------- «Заказы»: the positions we wait for, the receipts
    osh = gDoc.Sheets.getByName(SH_ORDERS)
    last = WmsOrders.LastRow(osh)
    If last >= 1 Then
        ' a position still expected: a partly received one (its first receipt is on its own row) or an order row without a
        ' receipt — by its data, as the status of «Заказы» is computed (a row typed while the macros were off has no status yet)
        rg = ColRng(SH_ORDERS, OC_STATUS, 1, last)
        cond = "(((" & ColRng(SH_ORDERS, OC_EI, 1, last) & "<>"""")*((" & rg & "=""" & OS_WAITING & """)+(" & rg & "=""" & OS_PARTIAL & """)+(" & rg _
            & "=""" & OS_OVERDUE & """)+(" & rg & "=""" & OS_PARTIAL_OVERDUE & """))+(" & ColRng(SH_ORDERS, OC_EI, 1, last) & "="""")*(" & rg & "<>""" _
            & OS_CANCELLED & """)*(LEFT(" & ColRng(SH_ORDERS, OC_CTL, 1, last) & ";5)<>""КОПИЯ"")*(((" & ColRng(SH_ORDERS, OC_ORDER, 1, last) _
            & "<>"""")+(" & ColRng(SH_ORDERS, OC_NAME, 1, last) & "<>"""")+(" & ColRng(SH_ORDERS, OC_ORDQTY, 1, last) & "<>""""))>0))>0)"
        ' the expected date (a date, or a date typed as text): key = day × 2^20 + row
        rg = ColRng(SH_ORDERS, OC_EDATE, 1, last)
        dq = "IF(ISNUMBER(" & rg & ");INT(" & rg & ");IFERROR(DATEVALUE(" & rg & ");0))"
        a = EngList("TEXTJOIN("";"";1;IF(" & cond & ";" & dq & "*1048576+ROW(" & rg & ")-1;""""))", nItems)
        nList = DB_LIST1_N
        For i = 0 To nItems - 1
            key = CDbl(a(i))
            dt = Int(key / 1048576)
            nWait = nWait + 1
            If dt <= 0 Then
                nNoDate = nNoDate + 1
            Else
                If dt < today Then nLate = nLate + 1
                If dt = today Then nToday = nToday + 1
                If nW < nList Then
                    TopPush(wKey, nW, nList, key)
                ElseIf key < wKey(nList - 1) Then
                    TopPush(wKey, nW, nList, key)
                End If
            End If
        Next i
        ' the receipts of 14 days by day (today first)
        rg = ColRng(SH_ORDERS, OC_RDATE, 1, last)
        a = Split(Eng("TEXTJOIN("";"";0;FREQUENCY(IF((" & ColRng(SH_ORDERS, OC_EI, 1, last) & "<>"""")*" & DaysExpr(rg, today - 13, today) & ";" _
            & Format(today, "0") & "-INT(" & rg & ");99);ROW($A$1:$A$14)-1))"), ";")
        For k = 0 To 13
            dayRcv(k) = dayRcv(k) + ListNum(a, k)
        Next k
        rcvToday = ListNum(a, 0)
    End If
    ' ---------------------------------------------------------------- «Иной приход»
    If gDoc.Sheets.hasByName(SH_SPECIAL) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_SPECIAL))
        If last >= 1 Then
            rg = ColRng(SH_SPECIAL, XC_DATE, 1, last)
            a = Split(Eng("TEXTJOIN("";"";0;FREQUENCY(IF(" & PostedExpr(ColRng(SH_SPECIAL, XC_STATUS, 1, last)) & "*" & DaysExpr(rg, today - 13, today) & ";" _
                & Format(today, "0") & "-INT(" & rg & ");99);ROW($A$1:$A$14)-1))"), ";")
            For k = 0 To 13
                dayRcv(k) = dayRcv(k) + ListNum(a, k)
            Next k
            splToday = ListNum(a, 0)
        End If
    End If
    ' ---------------------------------------------------------------- «Выдачи» of 30 days: today, 14 days, the most issued
    top = New Collection
    who = New Collection
    ReDim tName(1000)
    ReDim tCnt(1000)
    ReDim tQty(1000)
    ReDim tUnit(1000)
    last = WmsOrders.LastRow(WmsIssue.IssuesSheet())
    If last >= 1 Then
        rg = ColRng(SH_ISSUES, IC_DATE, 1, last)
        a = EngList("TEXTJOIN("";"";1;IF(" & PostedExpr(ColRng(SH_ISSUES, IC_CTL, 1, last)) & "*ISNUMBER(" & ColRng(SH_ISSUES, IC_NO, 1, last) & ")*" _
            & DaysExpr(rg, today - 29, today) & ";ROW(" & rg & ")-1;""""))", nItems)
        d = RowsData(WmsIssue.IssuesSheet(), a, nItems, IC_WHO)
        ' the columns copied once: a name of another module costs 15–25 µs on each use inside a loop
        cD = IC_DATE
        cW = IC_WHO
        cN = IC_NAME
        cU = IC_UNIT
        cQ = IC_QTY
        For i = 0 To nItems - 1
            row = d(i)
            k = CLng(today - Int(row(cD)))
            If k <= 13 Then dayIss(k) = dayIss(k) + 1
            If k = 0 Then
                issToday = issToday + 1
                If CollPut(who, 1, LCase(Trim(CStr(row(cW))))) Then nWho = nWho + 1
            End If
            nm = Trim(CStr(row(cN)))
            If nm <> "" Then
                j = CollGet(top, LCase(nm))
                If j < 0 And nTop <= UBound(tName) Then
                    j = nTop
                    CollPut(top, j, LCase(nm))
                    tName(j) = nm
                    tUnit(j) = CStr(row(cU))
                    nTop = nTop + 1
                End If
                If j >= 0 Then
                    tCnt(j) = tCnt(j) + 1
                    If VarType(row(cQ)) = 5 Then tQty(j) = tQty(j) + row(cQ)
                End If
            End If
        Next i
    End If
    ' ---------------------------------------------------------------- «Возврат», «Корректировки» today
    last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_RETURNS))
    If last >= 1 Then retToday = ListNum(Array(Eng("SUMPRODUCT(" & PostedExpr(ColRng(SH_RETURNS, RC_CTL, 1, last)) & "*" _
        & DaysExpr(ColRng(SH_RETURNS, RC_DATE, 1, last), today, today) & ")")), 0)
    If gDoc.Sheets.hasByName(SH_ADJUST) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_ADJUST))
        If last >= 1 Then adjToday = ListNum(Array(Eng("SUMPRODUCT(" & PostedExpr(ColRng(SH_ADJUST, AC_CTL, 1, last)) & "*" _
            & DaysExpr(ColRng(SH_ADJUST, AC_DATE, 1, last), today, today) & ")")), 0)
    End If
    ' ---------------------------------------------------------------- «Приход авто»
    If gDoc.Sheets.hasByName(SH_CARS) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_CARS))
        If last >= CAR_FIRST Then
            rg = ColRng(SH_CARS, CC_STATUS, CAR_FIRST, last)
            rd = DaysExpr(ColRng(SH_CARS, CC_DATE, CAR_FIRST, last), today, today)
            a = Split(Eng("SUMPRODUCT((" & rg & "=""" & CS_OPEN & """)*1)&""|""&SUMPRODUCT(" & rd & "*((" & rg & "=""" & CS_OPEN & """)+(" & rg & "=""" _
                & CS_GONE & """)))&""|""&SUMPRODUCT(" & rd & "*(" & rg & "=""" & CS_GONE & """))"), "|")
            carsIn = ListNum(a, 0)
            carsArr = ListNum(a, 1)
            carsDep = ListNum(a, 2)
        End If
    End If
    ' ---------------------------------------------------------------- without movement 90+ days: EIs with a balance that did not
    ' move in 90 days (the oldest EIs first)
    ReDim recent(nextEI)
    MarkRecent(recent, nextEI, today - DB_STALE_DAYS + 1)
    last = nextEI - 1
    If last >= 1 Then
        rg = ColRng(SH_STOCK, SC_QTY, 1, last)
        a = EngList("TEXTJOIN("";"";1;IF(ISNUMBER(" & rg & ")*(" & rg & ">0.0000005)*(" & ColRng(SH_STOCK, SC_STATE, 1, last) & "<>""" & EI_ST_STORNO _
            & """);ROW(" & rg & ")-1;""""))", nItems)
        nList = DB_STALE_CAND
        For i = 0 To nItems - 1
            n = CLng(a(i))
            If Not recent(n) Then
                nSt = nSt + 1
                If nStShown < nList Then
                    stN(nStShown) = n
                    nStShown = nStShown + 1
                End If
            End If
        Next i
    End If
    ' ---------------------------------------------------------------- 1.0 (Ч5): the tails
    tails = TailEIs(nTails)
    ' ---------------------------------------------------------------- attention: unposted rows, «Требует разбора»
    unposted = WmsIssue.UnpostedCount()
    attention = unposted + nReview
    If gVwEngFail <> "" Then
        DashRefresh = "ERR:дашборд не обновлён: " & gVwEngFail
        ViewDone(sh)
        gDbBusy = False
        Exit Function
    End If
    ' ================================================================ the page
    sh.getCellRangeByPosition(1, DB_TILE_VAL, 12, DB_TILE_SUB).clearContents(1 + 2 + 4)
    sh.getCellByPosition(1, DB_TITLE_ROW).setString("Склад сегодня · " & Format(Now(), "DD.MM.YYYY"))
    If gState = "CLEAN" Then
        s = "●  Работа разрешена" & IIf(gOldNote <> "", " · ВНИМАНИЕ: в старой книге проводились операции — см. «" & SH_MAIN & "»", "")
    ElseIf gState = "" Then
        s = "●  WMS не запущена — откройте книгу с макросами"
    Else
        s = "●  ЗАБЛОКИРОВАНО: " & IIf(gBlockText <> "", gBlockText, gBlock) & " — см. «" & SH_MAIN & "»"
    End If
    sh.getCellByPosition(8, DB_TITLE_ROW).setString(s)
    sh.getCellByPosition(DB_SVC_COL, 0).setString(gState & "|" & gBlock)
    ' the tiles: value, line under it
    ' 1.0 (А2): the money of the stock by the price of the order (cached until the next operation)
    money = WmsLedger.StockMoney()
    PutTile(sh, 0, nPos, "на сумму " & WmsPrint.Money(money(0)) & IIf(nPos > money(1), " (без цены " & NumText(nPos - money(1)) & ")", "") _
        & " · всего ЕИ " & NumText(nEI) & " · требует разбора " & NumText(nReview))
    ' 1.0 (З10): the goods of «Номенклатура» below their minimum
    nNeed = 0
    If gDoc.Sheets.hasByName(WmsSupply.SH_GOODS) Then
        If Left(WmsSupply.NeedCompute(""), 3) = "OK:" Then nNeed = gNdN
    End If
    PutTile(sh, 1, nWait, "просрочено " & NumText(nLate) & " · сегодня " & NumText(nToday) & " · без даты " & NumText(nNoDate) _
        & IIf(nNeed > 0, " · заказать " & NumText(nNeed), ""))
    PutTile(sh, 2, rcvToday + splToday, "по заказам " & NumText(rcvToday) & " · иной приход " & NumText(splToday))
    PutTile(sh, 3, issToday, "получателей " & NumText(nWho) & " · возвратов " & NumText(retToday) & " · корректировок " & NumText(adjToday))
    PutTile(sh, 4, carsIn, "приехало сегодня " & NumText(carsArr) & " · уехало " & NumText(carsDep))
    PutTile(sh, 5, attention, IIf(attention = 0, "всё в порядке", "непроведённых " & NumText(unposted) & " · требует разбора " & NumText(nReview) _
        & " · двойной щелчок — список"))
    ' list 1: the positions by the expected date (the earliest first); 1.0 (З3): what is still expected of each, «и ещё N», the
    ' row of «Заказы» (the hidden column) for a double click
    sh.getCellRangeByPosition(1, DB_LIST1, 6, DB_LIST1 + DB_LIST1_N - 1).clearContents(1 + 2 + 4)
    sh.getCellRangeByPosition(DB_SVC_COL, DB_LIST1, DB_SVC_COL, DB_LIST1 + DB_LIST1_N - 1).clearContents(1 + 2 + 4)
    For i = 0 To nW - 1
        dt = Int(wKey(i) / 1048576)
        r = CLng(wKey(i) - dt * 1048576)
        ord = osh.getCellRangeByPosition(0, r, OC_EDATE, r).getDataArray()(0)
        k = CLng(today - dt)
        If k > 0 Then
            s = "просрочено " & k & " дн."
        ElseIf k = 0 Then
            s = "сегодня"
        Else
            s = "через " & (-k) & " дн."
        End If
        rest = DashRest(r)
        PutRow(sh, DB_LIST1 + i, 1, Array(CStr(ord(OC_ORDER)), CStr(ord(OC_NAME)) & IIf(rest > 0, " · осталось " & WmsIssue.QtyText(rest) _
            & IIf(CStr(ord(OC_UNIT)) <> "", " " & ord(OC_UNIT), ""), ""), "", CStr(ord(OC_SUPPLIER)), ViewDay(dt), s))
        sh.getCellByPosition(DB_SVC_COL, DB_LIST1 + i).setString(SH_ORDERS & "|" & r)
    Next i
    sh.getCellByPosition(1, DB_LIST1_HEAD).setString("Ожидается: срок прошёл и ближайшие поставки" & IIf(nWait > nW, " · и ещё " & NumText(nWait - nW) _
        & " — все на «" & WmsSupply.SH_SUPPLY & "»", ""))
    If nW = 0 Then sh.getCellByPosition(2, DB_LIST1).setString(IIf(nWait = 0, "ничего не ожидается", "ожидаемая дата (Q) не указана ни у одной позиции"))
    ' list 2: the last operations of the journal (1.0, А6: the EI of each for a double click — the hidden column right of DB_SVC_COL)
    sh.getCellRangeByPosition(7, DB_LIST1, 12, DB_LIST1 + DB_LIST1_N - 1).clearContents(1 + 2 + 4)
    sh.getCellRangeByPosition(DB_SVC_COL + 1, DB_LIST1, DB_SVC_COL + 1, DB_LIST2 + DB_LIST2_N - 1).clearContents(1 + 2 + 4)
    sh.getCellRangeByPosition(DB_SVC_COL, DB_LIST2, DB_SVC_COL, DB_LIST2 + DB_LIST2_N - 1).clearContents(1 + 2 + 4)
    LastOps(sh)
    ' list 3: the most issued in 30 days (by the number of issues)
    sh.getCellRangeByPosition(1, DB_LIST2, 12, DB_LIST2 + DB_LIST2_N - 1).clearContents(1 + 2 + 4)
    If nTop > 0 Then
        ReDim tIdx(nTop - 1)
        For i = 0 To nTop - 1
            tIdx(i) = i
        Next i
        SortIdx(tCnt, tIdx, nTop, True)
        For i = 0 To DB_LIST2_N - 1
            If i >= nTop Then Exit For
            j = tIdx(i)
            PutRow(sh, DB_LIST2 + i, 1, Array(tName(j), "", "", tCnt(j), tQty(j), tUnit(j)))
            sh.getCellByPosition(DB_SVC_COL, DB_LIST2 + i).setString("GOODS|" & tName(j))
        Next i
    Else
        sh.getCellByPosition(1, DB_LIST2).setString("за 30 дней выдач нет")
    End If
    ' list 4: what lies without movement (the oldest EIs first; the days since their last movement on the sheets)
    ' 1.0: the ten lying longest — by the day of their last movement (none at all first); it was the ten lowest EI numbers,
    ' after an upgrade always the transferred EIs, whatever they were
    stDay = StaleDays(stN, nStShown)
    ReDim stKey(IIf(nStShown > 0, nStShown - 1, 0))
    ReDim stOrd(IIf(nStShown > 0, nStShown - 1, 0))
    For i = 0 To nStShown - 1
        stKey(i) = stDay(i) + stN(i) * 0.000001
        stOrd(i) = i
    Next i
    SortIdx(stKey, stOrd, nStShown, False)
    For i = 0 To nStShown - 1
        If i >= DB_LIST2_N Then Exit For
        j = stOrd(i)
        n = stN(j)
        d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_QTY, n).getDataArray()(0)
        PutRow(sh, DB_LIST2 + i, 7, Array(CStr(d(SC_EI)), CStr(d(SC_NAME)), "", "", d(SC_QTY), IIf(stDay(j) > 0, CLng(today - stDay(j)), "нет движений")))
        sh.getCellByPosition(DB_SVC_COL + 1, DB_LIST2 + i).setString("EI|" & d(SC_EI))
    Next i
    If nSt = 0 Then sh.getCellByPosition(8, DB_LIST2).setString("ЕИ без движения " & DB_STALE_DAYS & "+ дней нет")
    sh.getCellByPosition(7, DB_LIST2_HEAD).setString("Лежит без движения " & DB_STALE_DAYS & "+ дней" & IIf(nSt > 0, " — ЕИ: " & NumText(nSt) _
        & ", дольше всех лежащие первыми" & IIf(nSt > DB_STALE_CAND, " (из первых " & DB_STALE_CAND & ")", ""), ""))
    ' 1.0 (Ч5): the tails, the smallest first
    sh.getCellRangeByPosition(1, DB_TAILS, 12, DB_TAILS + DB_TAILS_N - 1).clearContents(1 + 2 + 4)
    For i = 0 To nTails - 1
        If i >= DB_TAILS_N Then Exit For
        n = tails(i)
        d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_PLACE, n).getDataArray()(0)
        PutRow(sh, DB_TAILS + i, 1, Array(CStr(d(SC_EI)), CStr(d(SC_NAME)), "", "", "", WmsPack.QtyWithParts(n, CDbl(d(SC_QTY)), CStr(d(SC_UNIT))), "", "", CStr(d(SC_PLACE))))
    Next i
    If nTails = 0 Then sh.getCellByPosition(2, DB_TAILS).setString("хвостов нет")
    sh.getCellByPosition(1, DB_TAILS_HEAD).setString("Хвосты — остаток меньше одной части" & IIf(nTails > 0, ": ЕИ " & NumText(nTails) _
        & IIf(nTails > DB_TAILS_N, " (показаны первые " & DB_TAILS_N & ")", "") & " — списать: кнопка «Списать хвосты…»", ""))
    ' the last 14 days: today first
    mx = 1
    For k = 0 To 13
        If dayIss(k) > mx Then mx = dayIss(k)
        If dayRcv(k) > mx Then mx = dayRcv(k)
    Next k
    ReDim rows(13)
    For k = 0 To 13
        dt = today - k
        rows(k) = Array(Format(CDate(dt), "DD.MM") & " " & Choose(WeekDay(CDate(dt)), "вс", "пн", "вт", "ср", "чт", "пт", "сб"), _
            IIf(dayIss(k) > 0, dayIss(k), ""), Bar(dayIss(k), mx, 24), "", "", "", "", IIf(dayRcv(k) > 0, dayRcv(k), ""), Bar(dayRcv(k), mx, 20), "", "", "")
    Next k
    sh.getCellRangeByPosition(1, DB_DAYS, 12, DB_DAYS + 13).setDataArray(rows)
    sh.getCellByPosition(DB_SVC_COL, 1).setValue(SysNum(SK_LAST_SEQ))
    sh.getCellByPosition(DB_SVC_COL, 2).setValue(today)
    sh.getCellByPosition(DB_SVC_COL, 3).setValue(ViewMs())
    sh.getCellByPosition(1, DB_FOOT).setString("Обновлено " & Format(Now(), "DD.MM.YYYY HH:MM") & " · считается по листам книги при " _
        & "открытии страницы и после каждой операции; страница ничего не изменяет")
    ViewDone(sh)
    gDbBusy = False
    DashRefresh = "OK:ЕИ " & nEI & ", с остатком " & nPos & ", ждём " & nWait & ", выдач сегодня " & issToday & ", внимание " & attention & ", " & ViewMs() & " мс"
    Exit Function
EH:
    DashRefresh = "ERR:дашборд не обновлён: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    ViewDone(sh)
    gDbBusy = False
End Function

' 1.0 (Ч5): the EIs of a tail, the smallest balance first: with a packing — more than 0 and less than one part; without —
' more than 0 and less than 0,01 of the unit of the document (what rounding leaves: 0,001 шт). A cancelled EI is not one
Function TailEIs(ByRef nItems As Long) As Variant
    Dim last As Long, rg As String, rp As String, a As Variant, i As Long, keys() As Double, idx() As Long, out() As Long, d As Variant, n As Long
    nItems = 0
    TailEIs = Array()
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last < 1 Then Exit Function
    rg = ColRng(SH_STOCK, SC_QTY, 1, last)
    If gDoc.Sheets.hasByName(SH_PACK) Then
        rp = ColRng(SH_PACK, PK_N, 1, last)
        rp = "IF((" & ColRng(SH_PACK, PK_EI, 1, last) & "=" & ColRng(SH_STOCK, SC_EI, 1, last) & ")*(" & rp & ">0);" & rg & "*" & rp & "<0.9999999;" _
            & rg & "<0.0099999)"
    Else
        rp = rg & "<0.0099999"
    End If
    a = EngList("TEXTJOIN("";"";1;IF(ISNUMBER(" & rg & ")*(" & rg & ">0.0000005)*(" & ColRng(SH_STOCK, SC_STATE, 1, last) & "<>""" & EI_ST_STORNO _
        & """)*(" & rp & ");ROW(" & rg & ")-1;""""))", nItems)
    If nItems = 0 Then Exit Function
    ReDim keys(nItems - 1)
    ReDim idx(nItems - 1)
    ReDim out(nItems - 1)
    For i = 0 To nItems - 1
        n = CLng(a(i))
        d = WmsIssue.StockSheet().getCellByPosition(SC_QTY, n).getValue()
        keys(i) = d + n * 0.0000000001
        idx(i) = i
    Next i
    SortIdx(keys, idx, nItems, False)
    For i = 0 To nItems - 1
        out(i) = CLng(a(idx(i)))
    Next i
    TailEIs = out
End Function

' «Списать хвосты…»: every tail is written off as a whole («всё») with one reason — rows of «Корректировки», posted one by one
Sub BtnTails(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = WriteOffTails()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Function WriteOffTails() As String
    Dim tails As Variant, nT As Long, i As Long, n As Long, d As Variant, sList As String, v(0) As String, ash As Object, r As Long, res As String
    Dim nOk As Long, nErr As Long, sErr As String
    tails = TailEIs(nT)
    If nT = 0 Then
        WriteOffTails = "ERR:хвостов нет — у всех ЕИ остаток не меньше одной части"
        Exit Function
    End If
    For i = 0 To nT - 1
        If i >= 8 Then
            sList = sList & Chr(10) & "… и ещё " & (nT - 8)
            Exit For
        End If
        n = tails(i)
        d = WmsIssue.StockData(n)
        sList = sList & Chr(10) & WmsIssue.EiCanon(n) & " · " & d(SC_NAME) & " · " & WmsPack.QtyWithParts(n, CDbl(d(SC_QTY)), CStr(d(SC_UNIT)))
    Next i
    v(0) = "хвост: остаток меньше одной части"
    If Not WmsPrimeUi.UiForm("Списать хвосты", "Списать весь остаток у " & nT & " ЕИ (на листе «" & SH_ADJUST & "» появятся строки «Списание … всё»):" _
        & sList, Array("Причина"), v, Array(""), "Списать") Then
        WriteOffTails = "SKIP:окно закрыто"
        gUiLastMsg = WriteOffTails
        Exit Function
    End If
    If Trim(v(0)) = "" Then
        WriteOffTails = "ERR:укажите причину списания"
        Exit Function
    End If
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    For i = 0 To nT - 1
        n = tails(i)
        r = WmsOrders.LastRow(ash) + 1
        If r < 1 Then r = 1
        ash.getCellByPosition(AC_KIND, r).setString("Списание")
        ash.getCellByPosition(AC_EI, r).setString(WmsIssue.EiCanon(n))
        ash.getCellByPosition(AC_QTY, r).setString("всё")
        ash.getCellByPosition(AC_DATE, r).setValue(CDbl(Date()))
        ash.getCellByPosition(AC_REASON, r).setString(Trim(v(0)))
        res = WmsAdjust.AdjustPostRow(r)
        If Left(res, 3) = "OK:" Then
            nOk = nOk + 1
        Else
            nErr = nErr + 1
            sErr = sErr & Chr(10) & WmsIssue.EiCanon(n) & ": " & Mid(res, InStr(res, ":") + 1)
            If Left(res, 4) <> "ERR:" And Left(res, 5) <> "SKIP:" Then Exit For
        End If
    Next i
    WriteOffTails = "OK=" & nOk & ";ERR=" & nErr
    If nErr > 0 Then WmsUi.UiMessage("Списано хвостов: " & nOk & " из " & nT & "; не списаны:" & sErr)
    gUiLastMsg = WriteOffTails & sErr
    DashRefresh()
End Function

Private Sub PutTile(sh As Object, t As Integer, v As Double, sSub As String)
    sh.getCellByPosition(1 + 2 * t, DB_TILE_VAL).setValue(v)
    sh.getCellByPosition(1 + 2 * t, DB_TILE_SUB).setString(sSub)
End Sub

' the last operations (the end of the newest journal file): time, operation, what, quantity
' the day of the last movement of each of n EIs (0 — none), looked up 50 at a time (LastMoveDays builds one formula of them)
Private Function StaleDays(eis() As Long, n As Long) As Variant
    Dim out() As Double, part() As Long, i As Long, j As Long, k As Long, m As Long, dd As Variant
    ReDim out(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1 Step 50
        m = n - i
        If m > 50 Then m = 50
        ReDim part(m - 1)
        For j = 0 To m - 1
            part(j) = eis(i + j)
        Next j
        dd = LastMoveDays(part, m)
        For k = 0 To m - 1
            out(i + k) = dd(k)
        Next k
    Next i
    StaleDays = out
End Function

Private Sub LastOps(sh As Object)
    Dim files As Variant, url As String, txt As String, a As Variant, i As Long, k As Integer, seq As Long, ts As String, inst As String
    Dim sType As String, fields As Variant, what As String, q As String, f As Integer, sz As Double, off As Double, xs As Object, tis As Object
    On Error GoTo EH
    files = WmsJournal.ListFiles()
    For f = UBound(files) To 0 Step -1
        url = gJDir & files(f)
        sz = gSfa.getSize(url)
        off = sz - 24000
        If off < 0 Then off = 0
        xs = gSfa.openFileRead(url)
        xs.seek(off)
        tis = CreateUnoService("com.sun.star.io.TextInputStream")
        tis.setInputStream(xs)
        tis.setEncoding("UTF-8")
        txt = tis.readString(Array(), False)
        tis.closeInput()
        a = Split(Replace(txt, Chr(13), ""), Chr(10))
        For i = UBound(a) To 0 Step -1
            If k >= DB_LIST1_N Then Exit For
            If WmsJournal.ParseLine(a(i), seq, ts, inst, sType, fields) Then
                If sType <> "LEGACY_BEGIN" And sType <> "LEGACY_END" Then
                    what = JField(fields, "EI")
                    If what = "" Then what = JField(fields, "PLATE")
                    If what = "" And JField(fields, "ORDER") <> "" Then what = "заказ " & JField(fields, "ORDER")
                    If JField(fields, "NAME") <> "" Then what = what & IIf(what <> "", " · ", "") & JField(fields, "NAME")
                    If JField(fields, "WHO") <> "" Then what = what & " → " & JField(fields, "WHO")
                    ' 1.0: the quantity as people write it — the decimal comma («0,25», the journal keeps «0.25»), the parts
                    ' it was given in («1 кан = 0,25 шт»); a packing — «в 1 шт — 4 кан»
                    q = JField(fields, "QTY")
                    If q = "" Then q = JField(fields, "DIFF")
                    If q <> "" Then q = IIf(Left(q, 1) = "-", "−", "") & WmsIssue.QtyText(Abs(Val(q)))
                    If q <> "" And JField(fields, "UNIT") <> "" Then q = q & " " & JField(fields, "UNIT")
                    If JField(fields, "PARTS") <> "" Then q = WmsIssue.QtyText(Val(JField(fields, "PARTS"))) & " " & JField(fields, "PART_UNIT") & " = " & q
                    If sType = "EI_PACK" Then
                        If JField(fields, "N") <> "" And Val(JField(fields, "N")) > 0 Then
                            q = "в 1 " & JField(fields, "DOC_UNIT") & " — " & WmsIssue.QtyText(Val(JField(fields, "N"))) & " " & JField(fields, "UNIT")
                        Else
                            q = "фасовка снята"
                        End If
                    End If
                    sh.getCellRangeByPosition(7, DB_LIST1 + k, 12, DB_LIST1 + k).setDataArray(Array(Array( _
                        Mid(ts, 9, 2) & "." & Mid(ts, 6, 2) & " " & Mid(ts, 12, 5), OpName(sType), "", what, "", q)))
                    If JField(fields, "EI") <> "" Then sh.getCellByPosition(DB_SVC_COL + 1, DB_LIST1 + k).setString("EI|" & JField(fields, "EI"))
                    k = k + 1
                End If
            End If
        Next i
        If k >= DB_LIST1_N Or off > 0 Then Exit For
    Next f
    If k = 0 Then sh.getCellByPosition(8, DB_LIST1).setString("операций ещё не было")
    Exit Sub
EH:
    sh.getCellByPosition(8, DB_LIST1).setString("журнал не прочитан: " & Error$)
End Sub

' ================================================================ 1.0 (А6, А7): the transitions of the page

' a double click on the page: the tile «Внимание» — the list of what needs action (А7); a row of a list — its order, its EI
' card or the goods on «Поиск» (А6). True: the cell is not edited
Function OnDashDoubleClick(Optional oTarget As Variant) As Boolean
    Dim r As Long, c As Long, sh As Object, s As String
    On Error GoTo EH
    WmsInit()
    r = oTarget.getCellAddress().Row
    c = oTarget.getCellAddress().Column
    sh = gDoc.Sheets.getByName(SH_DASH)
    If r >= DB_TILE_CAP And r <= DB_TILE_SUB Then
        OnDashDoubleClick = True
        If c >= 11 And c <= 12 Then
            AttentionShow()
            WmsView.GoCell(SH_FIND, FD_TABLE, 1)
        ElseIf c >= 3 And c <= 4 Then
            WmsSupply.NavSupply()
        End If
        Exit Function
    End If
    If (r >= DB_LIST1 And r < DB_LIST1 + DB_LIST1_N) Or (r >= DB_LIST2 And r < DB_LIST2 + DB_LIST2_N) Then
        s = sh.getCellByPosition(IIf(c <= 6, DB_SVC_COL, DB_SVC_COL + 1), r).getString()
    End If
    If s = "" Then Exit Function
    OnDashDoubleClick = True
    DashGo(s)
    Exit Function
EH:
End Function

' «Заказы|r» — the row; «EI|ЕИ-…» — the card on «Поиск»; «GOODS|name» — the search of the name
Sub DashGo(s As String)
    Dim a As Variant, n As Long
    a = Split(s, "|")
    If UBound(a) < 1 Then Exit Sub
    Select Case a(0)
    Case "EI"
        n = EiNo(CStr(a(1)), CLng(SysNum(SK_NEXT_EI)))
        If n > 0 Then
            WmsFind.CardShow(n)
            WmsView.GoCell(SH_FIND, FD_TABLE, 1)
        End If
    Case "GOODS"
        WmsFind.FindRun(CStr(a(1)) & "*")
        WmsView.GoCell(SH_FIND, FD_TABLE, 1)
    Case Else
        If gDoc.Sheets.hasByName(a(0)) And IsNumeric(a(1)) Then WmsView.GoCell(CStr(a(0)), CLng(a(1)), 1)
    End Select
    gUiLastMsg = "OK:" & s
End Sub

' «Все…» under the lists of the page
Sub BtnAllWait(Optional oEvent As Variant)
    WmsInit()
    WmsSupply.NavSupply()
End Sub

Sub BtnAllOps(Optional oEvent As Variant)
    WmsInit()
    WmsAudit.HistoryShow("", Today0() - 6, Today0(), "всё")
    WmsView.GoCell(SH_FIND, FD_TABLE, 1)
End Sub

Sub BtnAllIssued(Optional oEvent As Variant)
    WmsInit()
    WmsLedger.Consumers(Today0() - 29, Today0())
End Sub

Sub BtnAllStale(Optional oEvent As Variant)
    WmsInit()
    WmsLedger.DeadStock(DB_STALE_DAYS, False)
End Sub

' 1.0 (А7): what needs the storekeeper, as a list on «Поиск» with the transitions (a double click — the row; an EI — its card): the
' vehicles on the territory since an earlier day, the rows typed and not posted on every sheet, the EIs «Требует разбора», the
' possible duplicates of the receipts. OK:<rows>
Function AttentionShow() As String
    Dim rows() As Variant, n As Long, today As Double, sh As Object, last As Long, d As Variant, i As Long, items As Variant, nItems As Long, r As Long
    Dim nCar As Long, nOpen As Long, nRev As Long, nDup As Long, k As Long
    WmsInit()
    On Error GoTo EH
    today = Today0()
    ReDim rows(DB_AT_MAX)
    If gDoc.Sheets.hasByName(SH_CARS) Then
        sh = gDoc.Sheets.getByName(SH_CARS)
        last = WmsOrders.LastRow(sh)
        If last >= CAR_FIRST Then
            d = sh.getCellRangeByPosition(0, CAR_FIRST, CC_STATUS, last).getDataArray()
            For i = 0 To UBound(d)
                If CStr(d(i)(CC_STATUS)) = CS_OPEN And DateOf(d(i)(CC_DATE)) > 0 And DateOf(d(i)(CC_DATE)) < today Then
                    AtAdd(rows, n, Array("Машина с прошлого дня", SH_CARS & ", строка " & (CAR_FIRST + i + 1), ViewDay(DateOf(d(i)(CC_DATE))), CStr(d(i)(CC_PLATE)), _
                        CStr(d(i)(CC_SUP)), "всё ещё «" & CS_OPEN & "» — отметьте «Уехал» или «Отменить визит»", "", "", "", SH_CARS, CAR_FIRST + i, ""))
                    nCar = nCar + 1
                End If
            Next i
        End If
    End If
    nOpen = AtUnposted(rows, n, SH_ISSUES, IC_NO, IC_EI, IC_NAME, IC_DATE, "Выдача не проведена")
    nOpen = nOpen + AtUnposted(rows, n, SH_RETURNS, RC_NO, RC_EI, RC_NAME, RC_DATE, "Возврат не проведён")
    nOpen = nOpen + AtUnposted(rows, n, SH_ADJUST, AC_NO, AC_EI, AC_NAME, AC_DATE, "Корректировка не проведена")
    nOpen = nOpen + AtUnposted(rows, n, SH_SPECIAL, XC_NO, XC_QTY, XC_NAME, XC_DATE, "Иной приход не проведён")
    nOpen = nOpen + AtUnposted(rows, n, SH_ORDERS, OC_EI, OC_FACT, OC_NAME, OC_RDATE, "Приход не проведён")
    ' the EIs «Требует разбора»
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last >= 1 Then
        items = EngList("TEXTJOIN("";"";1;IF(" & ColRng(SH_STOCK, SC_STATE, 1, last) & "=""" & EI_ST_REVIEW & """;ROW(" & ColRng(SH_STOCK, SC_STATE, 1, last) _
            & ")-1;""""))", nItems)
        For i = 0 To nItems - 1
            r = CLng(items(i))
            d = WmsIssue.StockSheet().getCellRangeByPosition(0, r, SC_LAST, r).getDataArray()(0)
            AtAdd(rows, n, Array("ЕИ требует разбора", SH_STOCK & ", строка " & (r + 1), "", CStr(d(SC_EI)), CStr(d(SC_NAME)), "откуда: " & d(SC_SRC) _
                & " — разберите на «" & SH_SPECIAL & "» («Разобрать»)", "", "", "", SH_STOCK, r, CStr(d(SC_EI))))
            nRev = nRev + 1
        Next i
    End If
    ' the possible duplicates of the receipts
    sh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        items = EngList("TEXTJOIN("";"";1;IF(LEFT(" & ColRng(SH_ORDERS, OC_DUP, 1, last) & ";15)=""Возможный дубль"";ROW(" & ColRng(SH_ORDERS, OC_DUP, 1, last) _
            & ")-1;""""))", nItems)
        For i = 0 To nItems - 1
            r = CLng(items(i))
            d = sh.getCellRangeByPosition(0, r, OC_DUP, r).getDataArray()(0)
            AtAdd(rows, n, Array("Возможный дубль прихода", SH_ORDERS & ", строка " & (r + 1), ViewDay(DateOf(d(OC_RDATE))), CStr(d(OC_EI)), CStr(d(OC_NAME)), _
                CStr(d(OC_DUP)), "", "", "", SH_ORDERS, r, ""))
            nDup = nDup + 1
        Next i
    End If
    AttentionShow = WmsFind.ShowTable("Внимание: что требует действия", IIf(n = 0, "Всё в порядке: ничего не ждёт.", "Двойной щелчок по строке — перейти к ней " _
        & "(у ЕИ — карточка)."), "машины с прошлых дней " & nCar & " · не проведено строк " & nOpen & " · требует разбора " & nRev & " · возможных дублей " & nDup _
        & IIf(n > DB_AT_MAX, " · показаны первые " & DB_AT_MAX, ""), Array("Что", "Где", "Дата", "ЕИ / машина", "Наименование / поставщик", "Подробно", "", ""), rows, _
        IIf(n > DB_AT_MAX, DB_AT_MAX, n), False)
    If Left(AttentionShow, 3) = "OK:" Then AttentionShow = "OK:" & n
    gUiLastMsg = AttentionShow
    Exit Function
EH:
    AttentionShow = "ERR:список не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = AttentionShow
End Function

Private Sub AtAdd(rows() As Variant, n As Long, row As Variant)
    If n < DB_AT_MAX Then rows(n) = row
    n = n + 1
End Sub

' the rows of a sheet with an input (cTest) and no number (cKey — the number, or the EI on «Заказы»): fast through the empty
' cells of the key column. The count
Private Function AtUnposted(rows() As Variant, n As Long, sName As String, cKey As Integer, cTest As Integer, cName As Integer, cDate As Integer, _
    sWhat As String) As Long
    Dim sh As Object, last As Long, blocks As Variant, got As Variant, i As Long, j As Long, r As Long, flags As Long, d As Variant, k As Long, ctl As String
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(sName) Then Exit Function
    sh = gDoc.Sheets.getByName(sName)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA
    blocks = sh.getCellRangeByPosition(cKey, 1, cKey, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= 200 Then Exit For
        got = sh.getCellRangeByPosition(cTest, blocks(i).StartRow, cTest, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            For r = got(j).StartRow To got(j).EndRow
                d = sh.getCellRangeByPosition(0, r, cName, r).getDataArray()(0)
                ctl = ""
                If sName = SH_ORDERS Then ctl = sh.getCellByPosition(OC_STATUS, r).getString()
                If ctl <> OS_CANCELLED Then
                    AtAdd(rows, n, Array(sWhat, sName & ", строка " & (r + 1), ViewDay(DateOf(sh.getCellByPosition(cDate, r).getValue())), _
                        sh.getCellByPosition(cTest, r).getString(), CStr(d(cName)), "строка заполнена, но не проведена — «Почему?» на ней", "", "", "", sName, r, ""))
                    k = k + 1
                End If
                If k > DB_AT_MAX Then Exit For
            Next r
        Next j
    Next i
    AtUnposted = k
    Exit Function
EH:
    AtUnposted = k
End Function

' 1.0: the state line of the page in the saved file — «WMS не запущена» for someone who opens it with the macros off;
' after the save the line of this session again (WmsUi.UiSavedState)
Sub DashSavedState(bSaving As Boolean)
    Dim sh As Object, c As Object
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_DASH) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_DASH)
    c = sh.getCellByPosition(8, DB_TITLE_ROW)
    If Not bSaving And gDbStateLine = "" Then Exit Sub
    ' the cell is unlocked (tools/build_ods.py layout_10): no change of the protection while the book is being saved — it
    ' would put the writes into the Undo stack
    If c.CellProtection.IsLocked And sh.isProtected() Then Exit Sub
    If bSaving Then
        gDbStateLine = c.getString()
        c.setString("●  WMS не запущена — откройте книгу с включёнными макросами (см. «" & SH_MAIN & "»)")
    Else
        c.setString(gDbStateLine)
        gDbStateLine = ""
    End If
    Exit Sub
EH:
End Sub
