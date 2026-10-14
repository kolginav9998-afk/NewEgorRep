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
Public Const DB_DAYS_HEAD = 35
Public Const DB_DAYS = 37
Public Const DB_DAYS_N = 14
Public Const DB_FOOT = 52
' hidden service cells of column O: the state code, the LAST_SEQ, the day and the time (ms) of the last refresh
Public Const DB_SVC_COL = 14
Public Const DB_STALE_DAYS = 90
' a page counted longer than this is not counted again by itself when it is shown (a big book): «Обновить» counts it
Public Const DB_AUTO_MAX_MS = 5000

Global gDbBusy As Boolean

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

' Counts and writes the page. OK:<what was counted, ms> or ERR:<why>. The sheets are counted by formulas of the Calc
' engine (WmsView.Eng): only the rows they find — the positions we wait for, the issues of 30 days — are read by Basic.
Function DashRefresh() As String
    Dim sh As Object, today As Double, nextEI As Long, last As Long, i As Long, k As Long, n As Long, s As String, a As Variant, nItems As Long
    Dim nEI As Long, nPos As Long, nReview As Long, nWait As Long, nLate As Long, nToday As Long, nNoDate As Long, wKey(DB_LIST1_N - 1) As Double, nW As Long
    Dim rcvToday As Long, splToday As Long, issToday As Long, retToday As Long, adjToday As Long, carsIn As Long, carsArr As Long, carsDep As Long
    Dim who As Object, nWho As Long, dayIss(13) As Long, dayRcv(13) As Long, dt As Double, key As Double, r As Long
    Dim top As Object, tName() As String, tCnt() As Double, tQty() As Double, tUnit() As String, nTop As Long, j As Long, nm As String
    Dim unposted As Long, attention As Long, recent() As Boolean, nSt As Long, stN(DB_LIST2_N - 1) As Long, nStShown As Long, stDay As Variant
    Dim osh As Object, rows(), mx As Double, ord As Variant, tIdx() As Long, d As Variant, cond As String, dq As String, rg As String, rd As String
    Dim row As Variant, cD As Integer, cW As Integer, cN As Integer, cU As Integer, cQ As Integer, nList As Long
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
        nList = DB_LIST2_N
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
        s = "●  Работа разрешена"
    ElseIf gState = "" Then
        s = "●  WMS не запущена — откройте книгу с макросами"
    Else
        s = "●  ЗАБЛОКИРОВАНО: " & IIf(gBlockText <> "", gBlockText, gBlock) & " — см. «" & SH_MAIN & "»"
    End If
    sh.getCellByPosition(8, DB_TITLE_ROW).setString(s)
    sh.getCellByPosition(DB_SVC_COL, 0).setString(gState & "|" & gBlock)
    ' the tiles: value, line under it
    PutTile(sh, 0, nPos, "всего ЕИ " & NumText(nEI) & " · требует разбора " & NumText(nReview))
    PutTile(sh, 1, nWait, "просрочено " & NumText(nLate) & " · сегодня " & NumText(nToday) & " · без даты " & NumText(nNoDate))
    PutTile(sh, 2, rcvToday + splToday, "по заказам " & NumText(rcvToday) & " · иной приход " & NumText(splToday))
    PutTile(sh, 3, issToday, "получателей " & NumText(nWho) & " · возвратов " & NumText(retToday) & " · корректировок " & NumText(adjToday))
    PutTile(sh, 4, carsIn, "приехало сегодня " & NumText(carsArr) & " · уехало " & NumText(carsDep))
    PutTile(sh, 5, attention, IIf(attention = 0, "всё в порядке", "непроведённых " & NumText(unposted) & " · требует разбора " & NumText(nReview)))
    ' list 1: the positions by the expected date (the earliest first)
    sh.getCellRangeByPosition(1, DB_LIST1, 6, DB_LIST1 + DB_LIST1_N - 1).clearContents(1 + 2 + 4)
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
        PutRow(sh, DB_LIST1 + i, 1, Array(CStr(ord(OC_ORDER)), CStr(ord(OC_NAME)), "", CStr(ord(OC_SUPPLIER)), ViewDay(dt), s))
    Next i
    If nW = 0 Then sh.getCellByPosition(2, DB_LIST1).setString(IIf(nWait = 0, "ничего не ожидается", "ожидаемая дата (Q) не указана ни у одной позиции"))
    ' list 2: the last operations of the journal
    sh.getCellRangeByPosition(7, DB_LIST1, 12, DB_LIST1 + DB_LIST1_N - 1).clearContents(1 + 2 + 4)
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
        Next i
    Else
        sh.getCellByPosition(1, DB_LIST2).setString("за 30 дней выдач нет")
    End If
    ' list 4: what lies without movement (the oldest EIs first; the days since their last movement on the sheets)
    stDay = LastMoveDays(stN, nStShown)
    For i = 0 To nStShown - 1
        n = stN(i)
        d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_QTY, n).getDataArray()(0)
        PutRow(sh, DB_LIST2 + i, 7, Array(CStr(d(SC_EI)), CStr(d(SC_NAME)), "", "", d(SC_QTY), IIf(stDay(i) > 0, CLng(today - stDay(i)), "нет движений")))
    Next i
    If nSt = 0 Then sh.getCellByPosition(8, DB_LIST2).setString("ЕИ без движения " & DB_STALE_DAYS & "+ дней нет")
    sh.getCellByPosition(7, DB_LIST2_HEAD).setString("Лежит без движения " & DB_STALE_DAYS & "+ дней" & IIf(nSt > 0, " — ЕИ: " & NumText(nSt) & ", старые первыми", ""))
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
    sh.getCellByPosition(1, DB_FOOT).setString("Обновлено " & Format(Now(), "DD.MM.YYYY HH:MM") & " за " & ViewMs() & " мс · считается по листам книги при " _
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

Private Sub PutTile(sh As Object, t As Integer, v As Double, sSub As String)
    sh.getCellByPosition(1 + 2 * t, DB_TILE_VAL).setValue(v)
    sh.getCellByPosition(1 + 2 * t, DB_TILE_SUB).setString(sSub)
End Sub

' the last operations (the end of the newest journal file): time, operation, what, quantity
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
                    q = JField(fields, "QTY")
                    If q = "" Then q = JField(fields, "DIFF")
                    If q <> "" And JField(fields, "UNIT") <> "" Then q = q & " " & JField(fields, "UNIT")
                    sh.getCellRangeByPosition(7, DB_LIST1 + k, 12, DB_LIST1 + k).setDataArray(Array(Array( _
                        Mid(ts, 9, 2) & "." & Mid(ts, 6, 2) & " " & Mid(ts, 12, 5), OpName(sType), "", what, "", q)))
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
