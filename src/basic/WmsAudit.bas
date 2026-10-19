' WmsAudit — 1.0, этап G «Контроль и документы» (выбор кладовщика): «История изменений» по журналу (К2), остаток на дату (К7),
' оборотная ведомость (К8), выгрузка для бухгалтерии (К9), сверка с 1С (К10), закрытие месяца (К13), пересчёт места (К5).
' Всё читает книгу и журнал и ничего не проводит само (кроме закрытия месяца — это операция ядра). Library WmsPlus.
Option Explicit

' the sheets whose writes the history shows as «было → стало» (the service sheets are left out)
Private Const AU_SHEETS = "|Заказы|Выдачи|Возврат|Иной приход|Корректировки|Приход авто|Наличие|"
Public Const AU_MAX = 3000

Global gAuHead As Object
Global gAuHeadSheets As String
Global gAuHeadTxt() As String
Global gAuHeadN As Long

' ================================================================ К2 «История изменений»

Sub BtnHistory(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = HistoryWindow()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the window of the filters: an EI or a word, the period, the kind of the operations
Function HistoryWindow() As String
    Dim v(3) As String, d0 As Double, d1 As Double, msg As String, sEI As String
    sEI = WmsDaily.CursorEI()
    If sEI <> "" Then v(0) = sEI
    v(1) = Format(WmsOrders.Today() - 30, "DD.MM.YYYY")
    v(2) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    v(3) = "всё"
    If Not WmsPrimeUi.UiForm("История изменений", "Журнал WMS по-русски: когда, что сделано, что было → что стало, причина. Пустое первое поле — все " _
        & "операции периода.", Array("ЕИ или слово", "С даты", "По дату", "Что показывать"), v, _
        Array("", "", "", "всё|приходы|выдачи и возвраты|корректировки|исправления и удаления|машины"), "Показать") Then
        HistoryWindow = "SKIP:окно закрыто"
        gUiLastMsg = HistoryWindow
        Exit Function
    End If
    If Not WmsOrders.ParseDateTextAs(Trim(v(1)), "с даты", d0, msg) Then
        HistoryWindow = "ERR:" & msg
    ElseIf Not WmsOrders.ParseDateTextAs(Trim(v(2)), "по дату", d1, msg) Then
        HistoryWindow = "ERR:" & msg
    Else
        HistoryWindow = HistoryShow(Trim(v(0)), d0, d1, Trim(v(3)))
    End If
    gUiLastMsg = HistoryWindow
End Function

' the operations of the journal posted from d0 to d1 (the newest first): sWhat — an EI (all its lines) or a word anywhere in the
' line; sKind — the group of the operations. On «Поиск». OK:<n> / ERR:…
Function HistoryShow(sWhat As String, d0 As Double, d1 As Double, sKind As String) As String
    Dim files As Variant, f As Integer, url As String, txt As String, a As Variant, i As Long, seq As Long, ts As String, inst As String
    Dim sType As String, fields As Variant, rows() As Variant, n As Long, dt As Double, canon As String, en As Long, msg As String, key As String
    Dim m0 As String, m1 As String, fm As String, more As Boolean, nAll As Long, w As String
    On Error GoTo EH
    If sWhat <> "" Then
        If WmsIssue.NormalizeEI(sWhat, en, canon, msg) Then key = canon
    End If
    m0 = Format(d0, "YYYY-MM")
    m1 = Format(d1, "YYYY-MM")
    files = WmsJournal.ListFiles()
    ReDim rows(AU_MAX)
    gAuHeadSheets = "|"
    gAuHead = New Collection
    gAuHeadN = 0
    ReDim gAuHeadTxt(300)
    For f = UBound(files) To 0 Step -1
        fm = Mid(files(f), Len("WMS_journal_") + 1, 7)
        If fm >= m0 And fm <= m1 Then
            url = gJDir & files(f)
            txt = ReadText(url)
            a = Split(Replace(txt, Chr(13), ""), Chr(10))
            For i = UBound(a) To 0 Step -1
                If WmsJournal.ParseLine(a(i), seq, ts, inst, sType, fields) Then
                    dt = DateSerial(CInt(Left(ts, 4)), CInt(Mid(ts, 6, 2)), CInt(Mid(ts, 9, 2)))
                    If dt >= d0 And dt <= d1 And KindIn(sType, sKind) And sType <> "LEGACY_BEGIN" And sType <> "LEGACY_END" Then
                        If LineMatches(a(i), fields, key, sWhat) Then
                            nAll = nAll + 1
                            If n <= AU_MAX Then
                                rows(n) = HistoryRow(ts, seq, sType, fields)
                                n = n + 1
                            Else
                                more = True
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next f
    w = IIf(key <> "", key, IIf(sWhat <> "", "«" & sWhat & "»", "все операции"))
    HistoryShow = WmsFind.ShowTable("История изменений: " & w & " · " & Format(d0, "DD.MM.YYYY") & " — " & Format(d1, "DD.MM.YYYY") & " · " _
        & WmsView.NumText(nAll) & " " & WmsView.Plural(nAll, "операция", "операции", "операций"), "Новые сверху. «Было → стало» — изменённые ячейки " _
        & "листов (исправления, сторно); причина — как её записали при «Удалить» / «Исправить» и в корректировках." & IIf(more, " Показаны " & AU_MAX + 1 _
        & " последних — сузьте период или впишите ЕИ.", ""), "Двойной щелчок по строке с ЕИ — его карточка. Вид — «" & IIf(sKind <> "", sKind, "всё") & "».", _
        Array("Когда", "Операция", "ЕИ", "Что", "Количество", "№", "", "Было → стало, причина"), rows, n, False)
    If Left(HistoryShow, 3) = "OK:" Then HistoryShow = "OK:" & nAll
    WmsView.GoCell(SH_FIND, FD_TABLE, 1)
    gUiLastMsg = HistoryShow
    Exit Function
EH:
    HistoryShow = "ERR:история не построена: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = HistoryShow
End Function

' the whole text of a journal file
Function ReadText(url As String) As String
    Dim xs As Object, tis As Object
    xs = gSfa.openFileRead(url)
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(xs)
    tis.setEncoding("UTF-8")
    ReadText = tis.readString(Array(), False)
    tis.closeInput()
End Function

' the group of the operations of the window
Private Function KindIn(sType As String, sKind As String) As Boolean
    Select Case LCase(sKind)
    Case "", "всё"
        KindIn = True
    Case "приходы"
        KindIn = (Left(sType, 7) = "RECEIPT" Or Left(sType, 3) = "SP_" Or Left(sType, 5) = "ORDER" Or sType = "MIGRATE")
    Case "выдачи и возвраты"
        KindIn = (Left(sType, 5) = "ISSUE" Or Left(sType, 6) = "RETURN")
    Case "корректировки"
        KindIn = (InStr("|MOVE|WRITE_OFF|INV_ADJ|SPLIT|SUP_RET|", "|" & Replace(Replace(sType, "_FIX", ""), "_DEL", "") & "|") > 0)
    Case "исправления и удаления"
        KindIn = (Right(sType, 4) = "_FIX" Or Right(sType, 4) = "_DEL" Or sType = "CAR_CANCEL" Or sType = "ORDER_CANCEL")
    Case "машины"
        KindIn = (Left(sType, 4) = "CAR_")
    Case Else
        KindIn = True
    End Select
End Function

' a line of the journal names the EI key (EI, NEW_EI, OLD_EI) or holds the word sWhat
Private Function LineMatches(ln As String, fields As Variant, key As String, sWhat As String) As Boolean
    If key <> "" Then
        LineMatches = (WmsView.JField(fields, "EI") = key Or WmsView.JField(fields, "NEW_EI") = key Or WmsView.JField(fields, "OLD_EI") = key)
    ElseIf sWhat <> "" Then
        LineMatches = (InStr(1, WmsCore.Unesc(ln), sWhat, 1) > 0)
    Else
        LineMatches = True
    End If
End Function

' one line of the history: when, what, the EI, what for whom, the quantity, the number, «было → стало» and the reason
Private Function HistoryRow(ts As String, seq As Long, sType As String, fields As Variant) As Variant
    Dim ei As String, what As String, q As String, no As String, ch As String, why As String
    ei = WmsView.JField(fields, "EI")
    If ei = "" Then ei = WmsView.JField(fields, "NEW_EI")
    what = WmsView.JField(fields, "NAME")
    If what = "" And WmsView.JField(fields, "PLATE") <> "" Then what = WmsView.JField(fields, "PLATE") & " · " & WmsView.JField(fields, "SUP")
    If WmsView.JField(fields, "WHO") <> "" Then what = what & IIf(what <> "", " → ", "→ ") & WmsView.JField(fields, "WHO")
    If WmsView.JField(fields, "SUPPLIER") <> "" Then what = what & IIf(what <> "", " · ", "") & "поставщику " & WmsView.JField(fields, "SUPPLIER")
    If WmsView.JField(fields, "TO") <> "" Then what = what & IIf(what <> "", " · ", "") & "место «" & WmsView.JField(fields, "TO") & "»"
    q = WmsView.JField(fields, "QTY")
    If q = "" Then q = WmsView.JField(fields, "DIFF")
    If q <> "" Then q = IIf(Left(q, 1) = "-", "−", "") & WmsIssue.QtyText(Abs(Val(q)))
    If q <> "" And WmsView.JField(fields, "UNIT") <> "" Then q = q & " " & WmsView.JField(fields, "UNIT")
    no = WmsView.JField(fields, "NO")
    If no = "" Then no = WmsView.JField(fields, "ADJ")
    If no = "" Then no = WmsView.JField(fields, "RET")
    If no = "" Then no = WmsView.JField(fields, "CAR")
    If no = "" Then no = WmsView.JField(fields, "SPL")
    If no <> "" Then no = "№ " & no
    ' «было → стало»: the writes of the operation to the cells of the working sheets — for a correction and a storno
    If Right(sType, 4) = "_FIX" Or Right(sType, 4) = "_DEL" Or sType = "CAR_CANCEL" Or sType = "LEGACY_FIX" Or sType = "POSITION_FIX" Then ch = Changes(fields)
    why = WmsView.JField(fields, "WHY")
    If why = "" And (Right(sType, 4) = "_FIX" Or Right(sType, 4) = "_DEL") Then why = WmsView.JField(fields, "REASON")
    If why <> "" Then ch = ch & IIf(ch <> "", " · ", "") & "причина: " & why
    HistoryRow = Array(Mid(ts, 9, 2) & "." & Mid(ts, 6, 2) & "." & Left(ts, 4) & " " & Mid(ts, 12, 5), WmsView.OpName(sType), ei, what, q, no, "", ch, _
        "", "", "", ei)
End Function

' the changed cells of the working sheets of a journal line: «Количество: 2 → 3; Место: 1-1-1 → 1-2-1» (the derived writes and
' the protection are left out; at most 6)
Function Changes(fields As Variant) As String
    Dim i As Long, w As Variant, sName As String, c As Long, hd As String, b As String, a As String, k As Integer
    For i = 0 To UBound(fields)
        If Left(fields(i), 4) = "W=V|" Then
            w = Split(Mid(fields(i), 3), "|")
            If UBound(w) >= 5 Then
                sName = w(1)
                If InStr(AU_SHEETS, "|" & sName & "|") > 0 Then
                    c = CLng(Val(w(3)))
                    hd = HeadOf(sName, c)
                    b = CellText(CStr(w(4)), hd)
                    a = CellText(CStr(w(5)), hd)
                    If b <> a And k < 6 Then
                        Changes = Changes & IIf(Changes <> "", "; ", "") & hd & ": " & IIf(b <> "", b, "—") & " → " & IIf(a <> "", a, "—")
                        k = k + 1
                    End If
                End If
            End If
        End If
    Next i
End Function

' the header of column c of sheet sName (read once per sheet)
Private Function HeadOf(sName As String, c As Long) As String
    Dim sh As Object, hr As Long, d As Variant, j As Long, idx As Long
    If InStr(gAuHeadSheets, "|" & sName & "|") = 0 Then
        gAuHeadSheets = gAuHeadSheets & sName & "|"
        If gDoc.Sheets.hasByName(sName) Then
            sh = gDoc.Sheets.getByName(sName)
            hr = IIf(sName = SH_CARS, CAR_HEAD_ROW, 0)
            d = sh.getCellRangeByPosition(0, hr, 30, hr).getDataArray()(0)
            For j = 0 To UBound(d)
                If gAuHeadN > UBound(gAuHeadTxt) Then
                    ReDim Preserve gAuHeadTxt(2 * gAuHeadN)
                End If
                CollPut(gAuHead, gAuHeadN, sName & "|" & j)
                gAuHeadTxt(gAuHeadN) = CStr(d(j))
                gAuHeadN = gAuHeadN + 1
            Next j
        End If
    End If
    idx = CollGet(gAuHead, sName & "|" & c)
    If idx >= 0 Then HeadOf = gAuHeadTxt(idx)
    If HeadOf = "" Then HeadOf = "колонка " & WmsView.ColLetters(c)
End Function

' the value of an encoded cell of the journal as people read it (a date column — a date)
Private Function CellText(enc As String, hd As String) As String
    Dim x As Double
    Select Case Left(enc, 1)
    Case "E"
        CellText = ""
    Case "N"
        x = Val(Mid(enc, 2))
        If InStr(LCase(hd), "дата") > 0 And x > 30000 And x < 80000 Then
            CellText = Format(CDate(x), "DD.MM.YYYY")
        Else
            CellText = WmsIssue.QtyText(x)
        End If
    Case "S"
        CellText = WmsCore.Unesc(Mid(enc, 2))
    Case Else
        CellText = WmsCore.Unesc(Mid(enc, 2))
    End Select
End Function
