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
' hidden service cells of column O: the state code, the LAST_SEQ and the day of the last refresh
Public Const DB_SVC_COL = 14
Public Const DB_STALE_DAYS = 90

Global gDbBusy As Boolean

Private Sub PutRow(sh As Object, r As Long, c0 As Integer, v As Variant)
    sh.getCellRangeByPosition(c0, r, c0 + UBound(v), r).setDataArray(Array(v))
End Sub

Private Function WaitingStatus(s As String) As Boolean
    WaitingStatus = (s = OS_WAITING Or s = OS_PARTIAL Or s = OS_OVERDUE Or s = OS_PARTIAL_OVERDUE)
End Function

Private Function DashPosted(v As Variant) As Boolean
    DashPosted = (Left(CStr(v), Len(ST_POSTED)) = ST_POSTED)
End Function

' «Обновить» on «Дашборд»
Sub BtnDashRefresh(Optional oEvent As Variant)
    Dim s As String
    s = DashRefresh()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the page was shown: counted again when an operation was posted since the last count or the day changed
Sub OnDashFocus(Optional oEvent As Variant)
    Dim sh As Object
    On Error GoTo EH
    WmsInit()
    If gDbBusy Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_DASH)
    If CStr(sh.getCellByPosition(DB_SVC_COL, 1).getValue()) = CStr(SysNum(SK_LAST_SEQ)) And sh.getCellByPosition(DB_SVC_COL, 2).getValue() = Today0() _
        And sh.getCellByPosition(DB_SVC_COL, 0).getString() = gState & "|" & gBlock Then Exit Sub
    DashRefresh()
EH:
End Sub

' Counts and writes the page. OK:<what was counted, ms> or ERR:<why>
Function DashRefresh() As String
    Dim sh As Object, today As Double, nextEI As Long, last As Long, r0 As Long, r1 As Long, i As Long, d As Variant, n As Long
    Dim qty() As Double, lastMv() As Double, isLive() As Boolean, nEI As Long, nPos As Long, nReview As Long
    Dim nWait As Long, nLate As Long, nToday As Long, nNoDate As Long, wKey() As Double, wRow() As Long, wIdx() As Long, nW As Long
    Dim rcvToday As Long, splToday As Long, issToday As Long, retToday As Long, adjToday As Long, carsIn As Long, carsArr As Long, carsDep As Long
    Dim whoToday As String, nWho As Long, dayIss(13) As Long, dayRcv(13) As Long, k As Long, dt As Double, s As String, ok As Boolean
    Dim top As Object, tName() As String, tCnt() As Double, tQty() As Double, tUnit() As String, nTop As Long, key As String, j As Long
    Dim unposted As Long, balBad As Long, attention As Long, stKey() As Double, stIdx() As Long, nSt As Long
    Dim osh As Object, rows(), mx As Double, ord As Variant, tIdx() As Long
    If gDbBusy Then Exit Function
    gDbBusy = True
    On Error GoTo EH
    sh = ViewOpen(SH_DASH)
    today = Today0()
    nextEI = CLng(SysNum(SK_NEXT_EI))
    ReDim qty(nextEI)
    ReDim lastMv(nextEI)
    ReDim isLive(nextEI)
    ' ---------------------------------------------------------------- «Наличие» (row = EI)
    last = nextEI - 1
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = WmsIssue.StockSheet().getCellRangeByPosition(0, r0, SC_STATE, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(d(i)(SC_EI)) <> "" Then
                n = r0 + i
                nEI = nEI + 1
                If CStr(d(i)(SC_STATE)) <> EI_ST_STORNO Then
                    isLive(n) = True
                    If VarType(d(i)(SC_QTY)) = 5 Then qty(n) = d(i)(SC_QTY)
                    If qty(n) > 0.0000005 Then nPos = nPos + 1
                End If
                If CStr(d(i)(SC_STATE)) = EI_ST_REVIEW Then nReview = nReview + 1
            End If
        Next i
    Next r0
    ' ---------------------------------------------------------------- «Заказы»: receipts and the positions we wait for
    osh = gDoc.Sheets.getByName(SH_ORDERS)
    last = WmsOrders.LastRow(osh)
    ReDim wKey(IIf(last > 0, last, 0))
    ReDim wRow(IIf(last > 0, last, 0))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = osh.getCellRangeByPosition(0, r0, OC_CTL, r1).getDataArray()
        For i = 0 To r1 - r0
            s = CStr(d(i)(OC_EI))
            If s <> "" Then
                n = EiNo(s, nextEI)
                dt = DateOf(d(i)(OC_RDATE))
                If n > 0 And dt > lastMv(n) Then lastMv(n) = dt
                If dt = today Then rcvToday = rcvToday + 1
                k = CLng(today - dt)
                If dt > 0 And k >= 0 And k <= 13 Then dayRcv(k) = dayRcv(k) + 1
            End If
            ' a position still expected: a partly received one (its first receipt is on its own row) or an order row without a
            ' receipt — by its data, as the status of «Заказы» is computed (a row typed while the macros were off has no status yet)
            ok = False
            If s <> "" Then
                ok = WaitingStatus(CStr(d(i)(OC_STATUS)))
            ElseIf CStr(d(i)(OC_STATUS)) <> OS_CANCELLED And Left(CStr(d(i)(OC_CTL)), 5) <> "КОПИЯ" Then
                ok = (CStr(d(i)(OC_ORDER)) <> "" Or CStr(d(i)(OC_NAME)) <> "" Or CStr(d(i)(OC_ORDQTY)) <> "")
            End If
            If ok Then
                nWait = nWait + 1
                dt = DateOf(d(i)(OC_EDATE))
                If dt <= 0 Then
                    nNoDate = nNoDate + 1
                Else
                    If dt < today Then nLate = nLate + 1
                    If dt = today Then nToday = nToday + 1
                    wKey(nW) = dt
                    wRow(nW) = r0 + i
                    nW = nW + 1
                End If
            End If
        Next i
    Next r0
    ' ---------------------------------------------------------------- «Иной приход»
    If gDoc.Sheets.hasByName(SH_SPECIAL) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_SPECIAL))
        For r0 = 1 To last Step CHECK_CHUNK_ROWS
            r1 = r0 + CHECK_CHUNK_ROWS - 1
            If r1 > last Then r1 = last
            d = gDoc.Sheets.getByName(SH_SPECIAL).getCellRangeByPosition(0, r0, XC_STATUS, r1).getDataArray()
            For i = 0 To r1 - r0
                If DashPosted(d(i)(XC_STATUS)) Then
                    n = EiNo(CStr(d(i)(XC_EI)), nextEI)
                    dt = DateOf(d(i)(XC_DATE))
                    If n > 0 And dt > lastMv(n) Then lastMv(n) = dt
                    If dt = today Then splToday = splToday + 1
                    k = CLng(today - dt)
                    If dt > 0 And k >= 0 And k <= 13 Then dayRcv(k) = dayRcv(k) + 1
                End If
            Next i
        Next r0
    End If
    ' ---------------------------------------------------------------- «Выдачи»: today, 14 days, the most issued in 30 days
    top = New Collection
    ReDim tName(1000)
    ReDim tCnt(1000)
    ReDim tQty(1000)
    ReDim tUnit(1000)
    last = WmsOrders.LastRow(WmsIssue.IssuesSheet())
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = WmsIssue.IssuesSheet().getCellRangeByPosition(0, r0, IC_CTL, r1).getDataArray()
        For i = 0 To r1 - r0
            If DashPosted(d(i)(IC_CTL)) And VarType(d(i)(IC_NO)) = 5 Then
                n = EiNo(CStr(d(i)(IC_EI)), nextEI)
                dt = DateOf(d(i)(IC_DATE))
                If n > 0 And dt > lastMv(n) Then lastMv(n) = dt
                If dt = today Then
                    issToday = issToday + 1
                    s = "|" & LCase(Trim(CStr(d(i)(IC_WHO)))) & "|"
                    If InStr(whoToday, s) = 0 Then
                        whoToday = whoToday & s
                        nWho = nWho + 1
                    End If
                End If
                k = CLng(today - dt)
                If dt > 0 And k >= 0 And k <= 13 Then dayIss(k) = dayIss(k) + 1
                If dt > 0 And k >= 0 And k < 30 And Trim(CStr(d(i)(IC_NAME))) <> "" Then
                    key = LCase(Trim(CStr(d(i)(IC_NAME))))
                    j = CollGet(top, key)
                    If j < 0 And nTop <= UBound(tName) Then
                        j = nTop
                        CollPut(top, j, key)
                        tName(j) = Trim(CStr(d(i)(IC_NAME)))
                        tUnit(j) = CStr(d(i)(IC_UNIT))
                        nTop = nTop + 1
                    End If
                    If j >= 0 Then
                        tCnt(j) = tCnt(j) + 1
                        If VarType(d(i)(IC_QTY)) = 5 Then tQty(j) = tQty(j) + d(i)(IC_QTY)
                    End If
                End If
            End If
        Next i
    Next r0
    ' ---------------------------------------------------------------- «Возврат», «Корректировки»
    last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_RETURNS))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = gDoc.Sheets.getByName(SH_RETURNS).getCellRangeByPosition(0, r0, RC_CTL, r1).getDataArray()
        For i = 0 To r1 - r0
            If DashPosted(d(i)(RC_CTL)) Then
                n = EiNo(CStr(d(i)(RC_EI)), nextEI)
                dt = DateOf(d(i)(RC_DATE))
                If n > 0 And dt > lastMv(n) Then lastMv(n) = dt
                If dt = today Then retToday = retToday + 1
            End If
        Next i
    Next r0
    If gDoc.Sheets.hasByName(SH_ADJUST) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_ADJUST))
        For r0 = 1 To last Step CHECK_CHUNK_ROWS
            r1 = r0 + CHECK_CHUNK_ROWS - 1
            If r1 > last Then r1 = last
            d = gDoc.Sheets.getByName(SH_ADJUST).getCellRangeByPosition(0, r0, AC_CTL, r1).getDataArray()
            For i = 0 To r1 - r0
                If DashPosted(d(i)(AC_CTL)) Then
                    n = EiNo(CStr(d(i)(AC_EI)), nextEI)
                    dt = DateOf(d(i)(AC_DATE))
                    If n > 0 And dt > lastMv(n) Then lastMv(n) = dt
                    If dt = today Then adjToday = adjToday + 1
                End If
            Next i
        Next r0
    End If
    ' ---------------------------------------------------------------- «Приход авто»
    If gDoc.Sheets.hasByName(SH_CARS) Then
        last = WmsOrders.LastRow(gDoc.Sheets.getByName(SH_CARS))
        If last >= CAR_FIRST Then
            d = gDoc.Sheets.getByName(SH_CARS).getCellRangeByPosition(0, CAR_FIRST, CC_STATUS, last).getDataArray()
            For i = 0 To UBound(d)
                s = CStr(d(i)(CC_STATUS))
                If s = CS_OPEN Then carsIn = carsIn + 1
                If DateOf(d(i)(CC_DATE)) = today And (s = CS_OPEN Or s = CS_GONE) Then
                    carsArr = carsArr + 1
                    If s = CS_GONE Then carsDep = carsDep + 1
                End If
            Next i
        End If
    End If
    ' ---------------------------------------------------------------- attention: unposted rows, «Требует разбора», balance ≠ movements
    unposted = WmsIssue.UnpostedCount()
    WmsLegacy.BalanceCheck()
    balBad = gLgBad
    attention = unposted + nReview + balBad
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
    PutTile(sh, 5, attention, IIf(attention = 0, "всё в порядке", "непроведённых " & NumText(unposted) & " · разбор " & NumText(nReview) _
        & " · остаток ≠ движениям " & NumText(balBad)))
    ' list 1: the positions by the expected date (the earliest first)
    ReDim wIdx(IIf(nW > 0, nW - 1, 0))
    For i = 0 To nW - 1
        wIdx(i) = i
    Next i
    SortIdx(wKey, wIdx, nW, False)
    sh.getCellRangeByPosition(1, DB_LIST1, 6, DB_LIST1 + DB_LIST1_N - 1).clearContents(1 + 2 + 4)
    For i = 0 To DB_LIST1_N - 1
        If i >= nW Then Exit For
        r0 = wRow(wIdx(i))
        ord = osh.getCellRangeByPosition(0, r0, OC_EDATE, r0).getDataArray()(0)
        dt = wKey(wIdx(i))
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
    ' list 4: what lies without movement (a balance, the last movement longest ago; 90 days and more)
    ReDim stKey(nextEI)
    ReDim stIdx(nextEI)
    For n = 1 To nextEI - 1
        If isLive(n) And qty(n) > 0.0000005 And lastMv(n) > 0 And today - lastMv(n) >= DB_STALE_DAYS Then
            stKey(n) = lastMv(n)
            stIdx(nSt) = n
            nSt = nSt + 1
        End If
    Next n
    SortIdx(stKey, stIdx, nSt, False)
    For i = 0 To DB_LIST2_N - 1
        If i >= nSt Then Exit For
        n = stIdx(i)
        d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_UNIT, n).getDataArray()(0)
        PutRow(sh, DB_LIST2 + i, 7, Array(CStr(d(SC_EI)), CStr(d(SC_NAME)), "", "", qty(n), CLng(today - lastMv(n))))
    Next i
    If nSt = 0 Then sh.getCellByPosition(8, DB_LIST2).setString("ЕИ без движения " & DB_STALE_DAYS & "+ дней нет")
    sh.getCellByPosition(7, DB_LIST2_HEAD).setString("Лежит без движения " & DB_STALE_DAYS & "+ дней" & IIf(nSt > 0, " — ЕИ: " & NumText(nSt), ""))
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
