' WmsOps — 1.1: «Операции дня» — the journal of WMS in words. The journal is the main protection of the data, but a line of it is
' not for reading (J1;3049;…;SP_REFILL;EI=…;QTY=…). The sheet shows the operations of a date (or of the days from … to …) one per row: the time, what was
' done, the EI and its name, how much, where to or to whom, the document, the rest of the fields; and the tail of the journal — the operations that are
' not in the book yet (before «Восстановить» or «Отложить хвост журнала» one sees what they are). Reads the files of the journal only; works in a blocked
' book too; posts nothing. Library WmsPlus.
Option Explicit

Private Const SH_OPS = "Операции дня"
Private Const OP_FROM_ROW = 1
Private Const OP_STAT_ROW = 2
Private Const OP_HEAD = 3
Private Const OP_ROW0 = 4
Private Const OP_COLS = 10
Private Const OP_MAX = 5000

Global gOpsBusy As Boolean

' the name of an operation as the storekeeper says it
Function OpsName(sType As String, sMode As String, sSrc As String) As String
    Dim k As Long, codes As Variant, names As Variant
    Select Case sType
    Case "ISSUE": OpsName = "Выдача"
    Case "ISSUE_FIX": OpsName = "Выдача — исправление"
    Case "ISSUE_DEL": OpsName = "Выдача — удалена (сторно)"
    Case "ISSUE_DOC": OpsName = "Номер документа выдачи записан"
    Case "RECEIPT": OpsName = "Приход по заказу"
    Case "RECEIPT_ADD": OpsName = "Приход по заказу (ещё поступление)"
    Case "RECEIPT_FIX": OpsName = "Приход по заказу — исправление"
    Case "RECEIPT_DEL": OpsName = "Приход по заказу — удалён (сторно)"
    Case "RETURN": OpsName = "Возврат"
    Case "RETURN_FIX": OpsName = "Возврат — исправление"
    Case "RETURN_DEL": OpsName = "Возврат — удалён (сторно)"
    Case "SP_RECEIPT", "SP_REFILL", "SP_FIX", "SP_DEL", "SP_IDENTIFY"
        codes = Split(SRC_CODES, "|")
        names = Split(SRC_NAMES, "|")
        For k = 0 To UBound(codes)
            If codes(k) = sSrc Then OpsName = "Иной приход — " & LCase(names(k))
        Next k
        If OpsName = "" Then OpsName = "Иной приход"
        Select Case sType
        Case "SP_REFILL": OpsName = OpsName & " (пополнение ЕИ)"
        Case "SP_FIX": OpsName = OpsName & " — исправление"
        Case "SP_DEL": OpsName = OpsName & " — удалён (сторно)"
        Case "SP_IDENTIFY": OpsName = "Иной приход — деталь опознана"
        End Select
    Case "MOVE": OpsName = "Перемещение ЕИ"
    Case "WRITE_OFF": OpsName = "Списание"
    Case "INV_ADJ": OpsName = "Инвентаризация (корректировка остатка)"
    Case "SPLIT": OpsName = "Разделение ЕИ"
    Case "SUP_RET": OpsName = "Возврат поставщику"
    Case "KIT", "KIT_BUILD": OpsName = "Сборка комплекта"
    Case "KIT_DEL": OpsName = "Комплект разобран"
    Case "ASM_BUILD": OpsName = "Сборка детали"
    Case "ASM_DEL": OpsName = "Сборка детали отменена"
    Case "MOVE_FIX", "WRITE_OFF_FIX", "INV_ADJ_FIX", "SPLIT_FIX", "SUP_RET_FIX", "KIT_FIX": OpsName = "Корректировка — исправление"
    Case "MOVE_DEL", "WRITE_OFF_DEL", "INV_ADJ_DEL", "SPLIT_DEL", "SUP_RET_DEL": OpsName = "Корректировка — удалена (сторно)"
    Case "EI_PROP": OpsName = "Свойства ЕИ (серийный номер, срок, примечание)"
    Case "EI_PACK": OpsName = "Фасовка ЕИ"
    Case "CAR_ARRIVE": OpsName = "Машина приехала"
    Case "CAR_DEPART": OpsName = "Машина уехала"
    Case "CAR_FIX": OpsName = "Визит машины — исправление"
    Case "CAR_CANCEL": OpsName = "Визит машины — отменён"
    Case "ORDER_CANCEL", "ORDER_CANCEL_REST": OpsName = "Позиция заказа отменена"
    Case "ORDER_UNCANCEL": OpsName = "Отмена позиции заказа снята"
    Case "POSITION_FIX": OpsName = "Позиция заказа исправлена"
    Case "MIGRATE", "MIGRATE_FIX": OpsName = "Перенос из старой таблицы"
    Case "LEGACY_RECEIPT", "LEGACY_RECEIPT_ADD", "LEGACY_FIX": OpsName = "Перенос из старой таблицы (заказы)"
    Case "MONTH_CLOSE": OpsName = "Месяц закрыт"
    Case "MONTH_OPEN": OpsName = "Месяц открыт"
    Case "ABANDON": OpsName = "Отложенный хвост журнала"
    Case "START": OpsName = "Начало журнала"
    Case Else: OpsName = sType
    End Select
    If sMode = "ADD" And sType = "SP_RECEIPT" Then OpsName = OpsName & " (пополнение)"
End Function

' «2026-10-28T10:27:01/111» → «28.10.2026 10:27»
Private Function TsText(ts As String) As String
    If Len(ts) >= 16 And Mid(ts, 5, 1) = "-" Then
        TsText = Mid(ts, 9, 2) & "." & Mid(ts, 6, 2) & "." & Left(ts, 4) & " " & Mid(ts, 12, 5)
    Else
        TsText = ts
    End If
End Function

Private Function Skipped(sKey As String) As Boolean
    Skipped = (sKey = "CH" Or sKey = "INST" Or sKey = "TX")
End Function

' the fields that were not shown in their own columns, «КЛЮЧ=значение; …»
Private Function Rest(fields As Variant) As String
    Dim i As Long, p As Long, k As String, s As String
    For i = 0 To UBound(fields)
        p = InStr(fields(i), "=")
        If p > 1 Then
            k = Left(fields(i), p - 1)
            If Not Skipped(k) And k <> "EI" And k <> "QTY" And k <> "UNIT" And k <> "PLACE" And k <> "WHO" And k <> "DOC" Then
                s = s & IIf(s <> "", "; ", "") & k & "=" & WmsCore.Unesc(Mid(fields(i), p + 1))
            End If
        End If
    Next i
    If Len(s) > 300 Then s = Left(s, 300) & "…"
    Rest = s
End Function

Private Function JournalLines(sUrl As String) As Variant
    Dim xs As Object, tis As Object, txt As String
    If Not gSfa.exists(sUrl) Then
        JournalLines = Array()
        Exit Function
    End If
    xs = gSfa.openFileRead(sUrl)
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(xs)
    tis.setEncoding("UTF-8")
    txt = tis.readString(Array(), False)
    tis.closeInput()
    JournalLines = Split(Replace(txt, Chr(13), ""), Chr(10))
End Function

' the operation a line says as the row of the sheet: (№, time, what, EI, name, quantity, unit, place/whom, document, details)
Private Function OpRow(seq As Long, ts As String, sType As String, fields As Variant, names As Variant, nextEI As Long) As Variant
    Dim ei As String, n As Long, nm As String, q As String, sWho As String, sDate As String, doc As String, mode As String, src As String
    ei = WmsView.JField(fields, "EI")
    ' 1.2.0: the build of a part names the permanent EI of the part (NEW_EI) and the number of pieces (N)
    If sType = "ASM_BUILD" Or sType = "ASM_DEL" Then ei = WmsView.JField(fields, "NEW_EI")
    If ei <> "" Then
        n = WmsView.EiFast(ei, nextEI)
        If n >= 1 And n <= UBound(names) Then nm = CStr(names(n))
    End If
    q = WmsView.JField(fields, "QTY")
    If sType = "ASM_BUILD" Or sType = "ASM_DEL" Then q = WmsView.JField(fields, "N")
    sWho = WmsView.JField(fields, "WHO")
    If sWho = "" Then sWho = WmsView.JField(fields, "PLACE")
    If sType = "ISSUE" Or sType = "ISSUE_FIX" Then
        sDate = WmsView.JField(fields, "DATE")
        If sDate <> "" And Len(sDate) >= 10 Then
            If Left(sDate, 10) <> Left(ts, 10) Then sWho = sWho & " (дата выдачи " & Mid(sDate, 9, 2) & "." & Mid(sDate, 6, 2) & "." & Left(sDate, 4) & ")"
        End If
    End If
    doc = WmsView.JField(fields, "DOC")
    mode = WmsView.JField(fields, "MODE")
    src = WmsView.JField(fields, "SRC")
    OpRow = Array(seq, TsText(ts), OpsName(sType, mode, src), ei, nm, IIf(q <> "", Val(q), ""), WmsView.JField(fields, "UNIT"), sWho, doc, Rest(fields))
End Function

' the names of the EIs (index = the number of the EI)
Private Function EiNames(ByRef nextEI As Long) As Variant
    Dim d As Variant, i As Long, st As Object, cur As Object, last As Long, out() As String
    nextEI = CLng(SysNum(SK_NEXT_EI))
    st = WmsIssue.StockSheet()
    cur = st.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    If last < 1 Then
        EiNames = Array("")
        Exit Function
    End If
    d = st.getCellRangeByPosition(SC_NAME, 0, SC_NAME, last).getDataArray()
    ReDim out(last)
    For i = 1 To last
        out(i) = CStr(d(i)(0))
    Next i
    EiNames = out
End Function

' the date in a cell: a number (a date) or a text dd.mm.yyyy; 0 — empty or not a date
Private Function CellDate(sh As Object, c As Integer, r As Integer) As Double
    Dim cell As Object, t As String
    cell = sh.getCellByPosition(c, r)
    If cell.getType() = com.sun.star.table.CellContentType.VALUE Then
        CellDate = Int(cell.getValue())
        Exit Function
    End If
    t = Trim(cell.getString())
    If Len(t) = 10 And Mid(t, 3, 1) = "." And Mid(t, 6, 1) = "." Then
        If IsNumeric(Left(t, 2)) And IsNumeric(Mid(t, 4, 2)) And IsNumeric(Right(t, 4)) Then CellDate = CDbl(DateSerial(Val(Right(t, 4)), Val(Mid(t, 4, 2)), Val(Left(t, 2))))
    End If
End Function

Sub BtnOpsShow(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = OpsShowDates()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub BtnOpsTail(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = OpsShowTail()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub BtnOpsYesterday(Optional oEvent As Variant)
    Dim sh As Object, d As Date, s As String
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_OPS) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_OPS)
    d = WmsOrders.Today() - 1
    Do While Weekday(d, 2) > 5
        d = d - 1
    Loop
    sh.getCellByPosition(1, OP_FROM_ROW).setString(Format(d, "DD.MM.YYYY"))
    sh.getCellByPosition(3, OP_FROM_ROW).setString("")
    s = OpsShowDates()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the sheet is opened empty: today's operations
Sub OnOpsFocus(Optional oEvent As Variant)
    Dim sh As Object
    If gOpsBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    sh = gDoc.Sheets.getByName(SH_OPS)
    If sh.getCellByPosition(0, OP_ROW0).getString() = "" And sh.getCellByPosition(0, OP_STAT_ROW).getString() = "" Then OpsShowDates()
    Exit Sub
EH:
End Sub

Sub OnOpsChange(Optional oTarget As Variant)
    Dim ra As Object
    If gOpsBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> OP_FROM_ROW Or (ra.Column <> 1 And ra.Column <> 3) Then Exit Sub
    WmsInit()
    OpsShowDates()
    Exit Sub
EH:
End Sub

' the operations of the days «с» … «по» of the sheet (empty — today; «по» empty — the same day). OK:<rows> / ERR:…
Function OpsShowDates() As String
    Dim sh As Object, d0 As Double, d1 As Double
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_OPS) Then
        OpsShowDates = "ERR:в книге нет листа «" & SH_OPS & "» — книга не обновлена до 1.1"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_OPS)
    d0 = CellDate(sh, 1, OP_FROM_ROW)
    d1 = CellDate(sh, 3, OP_FROM_ROW)
    If d0 = 0 Then d0 = CDbl(WmsOrders.Today())
    If d1 = 0 Then d1 = d0
    If d1 < d0 Then
        OpsShowDates = "ERR:дата «по» раньше даты «с»"
        Exit Function
    End If
    OpsShowDates = OpsLoad(d0, d1, False)
    Exit Function
EH:
    OpsShowDates = "ERR:операции не показаны: " & Error$ & " (строка " & Erl & ")"
End Function

Function OpsShowTail() As String
    On Error GoTo EH
    OpsShowTail = OpsLoad(0, 0, True)
    Exit Function
EH:
    OpsShowTail = "ERR:хвост журнала не показан: " & Error$ & " (строка " & Erl & ")"
End Function

' the rows of the sheet: the operations of the days d0 … d1, or (bTail) the ones after the last operation of the book
Function OpsLoad(d0 As Double, d1 As Double, bTail As Boolean) As String
    Dim sh As Object, lines As Variant, names As Variant, nextEI As Long, i As Long, seq As Long, ts As String, inst As String, sType As String, fields As Variant
    Dim out() As Variant, n As Long, last As Long, day As String, lo As String, hi As String, dayD As Double, files As Variant, f As Long, mName As String
    Dim seqBook As Long, more As Boolean, cur As Object, e As Long, d As Date
    gOpsBusy = True
    On Error GoTo EH
    sh = gDoc.Sheets.getByName(SH_OPS)
    names = EiNames(nextEI)
    seqBook = CLng(SysNum(SK_LAST_SEQ))
    ReDim out(OP_MAX - 1)
    If bTail Then
        files = Array(WmsJournal.JournalFileName(DateAdd("m", -1, Now())), WmsJournal.JournalFileName(Now()))
    Else
        ' the files of the months of the days
        d = CDate(d0)
        mName = ""
        files = Array()
        Do While d <= CDate(d1)
            If InStr(mName, "|" & WmsJournal.JournalFileName(d) & "|") = 0 Then
                mName = mName & "|" & WmsJournal.JournalFileName(d) & "|"
                ReDim Preserve files(UBound(files) + 1)
                files(UBound(files)) = WmsJournal.JournalFileName(d)
            End If
            d = DateAdd("d", 1, d)
            If UBound(files) > 14 Then Exit Do
        Loop
    End If
    lo = Format(CDate(d0), "YYYY-MM-DD")
    hi = Format(CDate(d1), "YYYY-MM-DD")
    For f = 0 To UBound(files)
        lines = JournalLines(gJDir & files(f))
        For i = 0 To UBound(lines)
            If lines(i) <> "" Then
                If WmsJournal.ParseLine(CStr(lines(i)), seq, ts, inst, sType, fields) Then
                    day = Left(ts, 10)
                    If (bTail And seq > seqBook) Or (Not bTail And day >= lo And day <= hi) Then
                        If n < OP_MAX Then
                            out(n) = OpRow(seq, ts, sType, fields, names, nextEI)
                            n = n + 1
                        Else
                            more = True
                        End If
                    End If
                End If
            End If
        Next i
    Next f
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    e = cur.getRangeAddress().EndRow
    If e >= OP_ROW0 Then sh.getCellRangeByPosition(0, OP_ROW0, OP_COLS - 1, e).clearContents(1 + 2 + 4 + 16)
    If n > 0 Then
        ReDim Preserve out(n - 1)
        sh.getCellRangeByPosition(0, OP_ROW0, OP_COLS - 1, OP_ROW0 + n - 1).setDataArray(out)
    End If
    If bTail Then
        sh.getCellByPosition(0, OP_STAT_ROW).setString("Хвост журнала: операций " & n & ", которых нет в книге (книга дошла до № " & seqBook & ")" _
            & IIf(n = 0, " — книга и журнал совпадают", ". Это то, что сделает «Восстановить» (или отложит «Отложить хвост журнала»)."))
    Else
        sh.getCellByPosition(0, OP_STAT_ROW).setString("Операций за " & Format(CDate(d0), "DD.MM.YYYY") & IIf(d1 > d0, " — " & Format(CDate(d1), "DD.MM.YYYY"), "") & ": " & n _
            & IIf(more, " (показаны первые " & OP_MAX & ")", "") & ". Журнал только читается.")
    End If
    sh.getCellRangeByPosition(0, OP_ROW0, 0, OP_ROW0 + IIf(n > 0, n - 1, 0)).NumberFormat = 0
    gOpsBusy = False
    OpsLoad = "OK:" & n
    gUiLastMsg = OpsLoad
    Exit Function
EH:
    gOpsBusy = False
    OpsLoad = "ERR:операции не показаны: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = OpsLoad
End Function

' «Главная» → «Операции дня»: the sheet opened (today's operations; in a book with a tail of the journal — the tail)
Sub BtnOpsOpen(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_OPS) Then
        WmsUi.UiMessage("В этой книге нет листа «" & SH_OPS & "» — книга не обновлена до 1.1.")
        Exit Sub
    End If
    gDoc.getCurrentController().setActiveSheet(gDoc.Sheets.getByName(SH_OPS))
    If gBlock = "TAIL" Then s = OpsShowTail() Else s = OpsShowDates()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub
