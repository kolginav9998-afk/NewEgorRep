' WmsExport — контракт WMS и внешних инструментов (задание «FINAL WMS MARATHON», §3; docs/EXPORT_CONTRACT.md).
'  - «Экспорт для инструментов»: неизменяемый снимок вне книги — папка WMS_Export/WMS_SNAPSHOT_<дата-время>_seq<N>/:
'    manifest.csv (формат, версии, экземпляр, время, LAST_SEQ, позиция журнала, счётчики, защиты листов и структуры
'    книги — для WMS_DOCTOR, файлы: строк, байт, SHA-256),
'    таблицы листов (UTF-8, «;», первая строка — английские имена колонок, строка k файла = строка k листа, даты
'    ГГГГ-ММ-ДД, числа с точкой), service/ (служебные таблицы — для WMS_DOCTOR), journal/ (копия журнала). Листы копируются
'    во временную скрытую книгу средствами LibreOffice (importSheet), там приводятся форматы, запись — фильтром CSV:
'    без построчного цикла Basic. Книга не меняется, операция в журнал не пишется. Папка появляется под своим именем
'    только целиком.
'  - «Загрузить пакет»: файл инструмента (WMS-BATCH-1) → проверка целиком → строки ввода листа «Корректировки»,
'    «Иной приход» или «Заказы» (только колонки ввода, служебные листы не трогаются) → по подтверждению — обычное
'    проведение этих строк (каждая строка — своя операция со всеми проверками). Один пакет загружается один раз.
Option Explicit

Public Const SNAPSHOT_FORMAT = "WMS-SNAPSHOT-1"
Public Const BATCH_FORMAT = "WMS-BATCH-1"
Public Const EXPORT_DIR = "WMS_Export"
Public Const BATCH_MAX_ROWS = 5000

' test seam: the file of the next «Загрузить пакет» (the file window is not shown), the answer to «Провести сейчас?»
Global gBatchFileSet As Boolean
Global gBatchFile As String
' the folder of the last snapshot (the test seam reads it)
Global gLastSnapshot As String

' ================================================================ CSV

' one field of the contract: text as is (quoted when it holds ; " or a line break), numbers with a dot, dates ISO
Function CsvField(v As Variant, bDate As Boolean) As String
    Dim s As String
    Select Case VarType(v)
    Case 0, 1
        CsvField = ""
    Case 2, 3, 4, 5, 6
        If bDate And v >= 36526 And v < 73051 And v = Int(v) Then
            CsvField = Format(CDate(v), "YYYY-MM-DD")
        Else
            CsvField = NumStr(CDbl(v))
        End If
    Case Else
        s = CStr(v)
        If InStr(s, ";") > 0 Or InStr(s, """") > 0 Or InStr(s, Chr(10)) > 0 Or InStr(s, Chr(13)) > 0 Then
            CsvField = """" & Replace(s, """", """""") & """"
        Else
            CsvField = s
        End If
    End Select
End Function

' the fields of one CSV line (quotes and doubled quotes as written by CsvField; a line break inside quotes is not
' expected in a batch line: the reader splits lines first)
Function CsvSplit(ByVal sLine As String) As Variant
    Dim aOut() As String, n As Long, i As Long, sCh As String, sFld As String, inQ As Boolean, q As String, k As Long, res() As String
    q = Chr(34)
    ' at most one field per character: no ReDim Preserve inside the loop
    ReDim aOut(Len(sLine) + 1) As String
    i = 1
    Do While i <= Len(sLine)
        sCh = Mid(sLine, i, 1)
        If inQ And sCh = q And Mid(sLine, i + 1, 1) = q Then
            sFld = sFld & q
            i = i + 1
        ElseIf inQ And sCh = q Then
            inQ = False
        ElseIf inQ Then
            sFld = sFld & sCh
        ElseIf sCh = q Then
            inQ = True
        ElseIf sCh = ";" Then
            aOut(n) = sFld
            n = n + 1
            sFld = ""
        Else
            sFld = sFld & sCh
        End If
        i = i + 1
    Loop
    aOut(n) = sFld
    ReDim res(n) As String
    For k = 0 To n
        res(k) = aOut(k)
    Next k
    CsvSplit = res
End Function

' SHA-256 of a file (hex, lower case) through the digest of LibreOffice (NSS); "" when not available
Function Sha256File(url As String, ByRef nBytes As Double) As String
    Dim nss As Object, dc As Object, inp As Object, buf() As Byte, n As Long, h As Variant, i As Long, s As String
    On Error GoTo EH
    nBytes = 0
    nss = CreateUnoService("com.sun.star.xml.crypto.NSSInitializer")
    dc = nss.getDigestContext(com.sun.star.xml.crypto.DigestID.SHA256, Array())
    inp = CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(url)
    Do
        n = inp.readBytes(buf, 1048576)
        If n <= 0 Then Exit Do
        nBytes = nBytes + n
        If n < 1048576 Then ReDim Preserve buf(n - 1)
        dc.updateDigest(buf)
    Loop While n = 1048576
    inp.closeInput()
    h = dc.finalizeDigestAndDispose()
    For i = 0 To UBound(h)
        s = s & Right("0" & LCase(Hex(h(i) And 255)), 2)
    Next i
    Sha256File = s
    Exit Function
EH:
    Sha256File = ""
End Function

' ================================================================ «Экспорт для инструментов»

' the tables of the snapshot: sheet, file (without .csv), column keys (| separated; "" — the header row of the book is
' kept: the service tables), date columns (0-based, comma separated)
Private Function TableSpecs() As Variant
    TableSpecs = Array( _
        Array(SH_STOCK, "stock", "ei|name|article|unit|qty|place|category|state|source|source_type", ""), _
        Array(SH_ORDERS, "orders", "order_no|name|doc_no|invoice_no|article|qty_fact|qty_doc|qty_ordered|unit|price|amount|supplier|seller|" _
            & "date_received|date_doc|date_order|date_expected|buyer|category|assignee|place|ei|status|stock|control|note|lead_days|possible_dup", _
            "13,14,15,16"), _
        Array(SH_SPECIAL, "special", "line_no|type|event|name|article|qty|unit|date|place|category|from_who|doc|old_mark|ei|bal_before|" _
            & "bal_after|status|control|note", "7"), _
        Array(SH_ISSUES, "issues", "issue_no|doc_no|name|article|qty|qty_pct|unit|date|recipient|place|category|ei|returned_legacy|" _
            & "returned_pct_legacy|note|bal_before|bal_after|status", "7"), _
        Array(SH_RETURNS, "returns", "return_no|issue_no|ei|name|article|qty|unit|date|from_who|place|category|bal_before|bal_after|" _
            & "status|note", "7"), _
        Array(SH_ADJUST, "adjustments", "adj_no|kind|ei|name|article|unit|qty|fact|book|place_to|date|reason|place_from|bal_before|" _
            & "bal_after|diff|status|batch|note", "10"), _
        Array(SH_RCPT, "recipients", "short|full", ""), _
        Array(SH_ORD, "service/_ORD", "", ""), Array(SH_RCV, "service/_RCV", "", ""), Array(SH_SPR, "service/_SPR", "", "8"), _
        Array(SH_ART, "service/_ART", "", ""), Array(SH_RET, "service/_RET", "", ""), Array(SH_ISS, "service/_ISS", "", ""), _
        Array(SH_ADJ, "service/_ADJ", "", "10"), Array(SYS_SHEET, "service/_SYS", "", ""))
End Function

Private Function PV(sName As String, v As Variant) As com.sun.star.beans.PropertyValue
    Dim x As New com.sun.star.beans.PropertyValue
    x.Name = sName
    x.Value = v
    PV = x
End Function

' Writes the snapshot. OK:<folder> — <tables and rows>; <ms>, ERR:<why>, BUSY:… Nothing in the book changes.
Function ExportSnapshot() As String
    Dim sfa As Object, sBase As String, sName As String, tmp As String, fin As String, specs As Variant, i As Integer, t0 As Long
    Dim files() As String, rows() As Long, nf As Integer, man As String, h As String, nb As Double, k As Integer, jf As Variant
    Dim summary As String, td As Object, sh As Object, loc As New com.sun.star.lang.Locale, nfs As Object, kGen As Long, kDate As Long
    Dim keys As Variant, dates As Variant, j As Integer, last As Long, lastCol As Long, stage() As String, cur As Object
    WmsInit()
    If gBusy Then
        ExportSnapshot = "BUSY:операция уже выполняется"
        Exit Function
    End If
    If gState <> "CLEAN" Then
        ExportSnapshot = "ERR:WMS заблокирована (" & StateLine() & ") — снимок делается только из рабочего состояния"
        Exit Function
    End If
    On Error GoTo EH
    t0 = GetSystemTicks()
    gBusy = True
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    sBase = gBaseDir & EXPORT_DIR & "/"
    If Not sfa.exists(sBase) Then sfa.createFolder(sBase)
    sName = "WMS_SNAPSHOT_" & Format(Now(), "YYYYMMDD-HHMMSS") & "_seq" & CLng(SysNum(SK_LAST_SEQ))
    fin = sBase & sName & "/"
    k = 1
    Do While sfa.exists(fin)
        k = k + 1
        fin = sBase & sName & "_" & k & "/"
    Loop
    If k > 1 Then sName = sName & "_" & k
    tmp = sBase & "." & sName & ".part/"
    If sfa.exists(tmp) Then sfa.kill(tmp)
    sfa.createFolder(tmp)
    sfa.createFolder(tmp & "service/")
    sfa.createFolder(tmp & "journal/")
    specs = TableSpecs()
    jf = WmsJournal.ListFiles()
    ReDim files(UBound(specs) + UBound(jf) + 2)
    ReDim rows(UBound(specs) + UBound(jf) + 2)
    ReDim stage(UBound(specs))
    ' 1. the staging book: every table copied by LibreOffice itself, one sheet per file (the sheet names are s0, s1 …)
    td = StarDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array(PV("Hidden", True)))
    nfs = td.getNumberFormats()
    loc.Language = "en"
    loc.Country = "US"
    kGen = nfs.getStandardFormat(com.sun.star.util.NumberFormat.NUMBER, loc)
    kDate = nfs.queryKey("YYYY-MM-DD", loc, False)
    If kDate < 0 Then kDate = nfs.addNew("YYYY-MM-DD", loc)
    For i = 0 To UBound(specs)
        td.Sheets.importSheet(gDoc, specs(i)(0), td.Sheets.getCount())
        sh = td.Sheets.getByIndex(td.Sheets.getCount() - 1)
        If sh.isProtected() Then sh.unprotect(PROTECT_PWD)
        stage(i) = "s" & i
        sh.Name = stage(i)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        last = cur.getRangeAddress().EndRow
        lastCol = cur.getRangeAddress().EndColumn
        If specs(i)(2) <> "" Then
            keys = Split(specs(i)(2), "|")
            lastCol = UBound(keys)
            sh.getCellRangeByPosition(0, 0, lastCol, 0).setDataArray(Array(keys))
        End If
        ' numbers with a dot, whole precision; dates ISO (the contract does not depend on the locale of the PC)
        sh.getCellRangeByPosition(0, 0, lastCol, IIf(last < 1, 1, last)).NumberFormat = kGen
        If specs(i)(3) <> "" Then
            dates = Split(specs(i)(3), ",")
            For j = 0 To UBound(dates)
                sh.getCellRangeByPosition(CInt(dates(j)), 1, CInt(dates(j)), IIf(last < 1, 1, last)).NumberFormat = kDate
            Next j
        End If
        ' a table of the contract has exactly its columns (nothing to the right of them)
        If specs(i)(2) <> "" And cur.getRangeAddress().EndColumn > lastCol Then
            sh.getCellRangeByPosition(lastCol + 1, 0, cur.getRangeAddress().EndColumn, IIf(last < 1, 1, last)).clearContents(1023)
        End If
        rows(nf) = last
        files(nf) = specs(i)(1) & ".csv"
        nf = nf + 1
    Next i
    td.Sheets.removeByName(td.Sheets.getByIndex(0).Name)
    ' 2. one CSV file per sheet, written by the CSV filter of LibreOffice (UTF-8, «;», «"»; cell contents as formatted)
    td.storeToURL(tmp & "t.csv", Array(PV("FilterName", "Text - txt - csv (StarCalc)"), PV("FilterOptions", "59,34,76,1,,1033,false,true,true,false,false,-1")))
    td.close(True)
    td = Nothing
    For i = 0 To UBound(specs)
        sfa.move(tmp & "t-" & stage(i) & ".csv", tmp & files(i))
    Next i
    ' 3. the journal (every file; the book state ends at LAST_SEQ of the manifest)
    For i = 0 To UBound(jf)
        sfa.copy(gJDir & jf(i), tmp & "journal/" & jf(i))
        files(nf) = "journal/" & jf(i)
        rows(nf) = -1
        nf = nf + 1
    Next i
    man = "key;value" & Chr(10) & "format;" & SNAPSHOT_FORMAT & Chr(10) & "product_version;" & WMS_PRODUCT_VERSION & Chr(10) _
        & "core_version;" & WMS_CORE_VERSION & Chr(10) & "schema;" & SysStr(SK_SCHEMA) & Chr(10) & "instance_id;" & SysStr(SK_INSTANCE) & Chr(10) _
        & "mode;" & SysStr(SK_MODE) & Chr(10) & "created;" & TS() & Chr(10) & "book;" & CsvField(ConvertFromURL(gDoc.getURL()), False) & Chr(10) _
        & "last_seq;" & SysStr(SK_LAST_SEQ) & Chr(10) & "journal_pos;" & CsvField(SysStr(SK_JPOS), False) & Chr(10) _
        & "next_ei;" & SysStr(SK_NEXT_EI) & Chr(10) & "next_issue;" & SysStr(SK_NEXT_NO) & Chr(10) & "next_return;" & SysStr(SK_NEXT_RET) & Chr(10) _
        & "next_line;" & SysStr(SK_NEXT_SPL) & Chr(10) & "next_adj;" & SysStr(SK_NEXT_ADJ) & Chr(10) & "book_modified;" & IIf(gDoc.isModified(), "1", "0") & Chr(10)
    ' the protections of the book (WMS_DOCTOR): every sheet — protected, visible; the structure of the book
    For i = 0 To gDoc.Sheets.getCount() - 1
        sh = gDoc.Sheets.getByIndex(i)
        man = man & "sheet;" & CsvField(sh.Name, False) & ";" & IIf(sh.isProtected(), "1", "0") & ";" & IIf(sh.IsVisible, "1", "0") & Chr(10)
    Next i
    man = man & "structure;" & IIf(gDoc.isProtected(), "1", "0") & Chr(10)
    For i = 0 To nf - 1
        h = Sha256File(tmp & files(i), nb)
        If h = "" Then
            ExportSnapshot = "ERR:не удалось посчитать SHA-256 файла " & files(i) & " — снимок не создан"
            GoTo FAIL
        End If
        man = man & "file;" & files(i) & ";" & rows(i) & ";" & NumStr(nb) & ";" & h & Chr(10)
        If rows(i) >= 0 And InStr(files(i), "/") = 0 Then summary = summary & IIf(summary <> "", ", ", "") & files(i) & " " & rows(i)
    Next i
    WriteTextFile(tmp & "manifest.csv", man)
    sfa.move(tmp, fin)
    gLastSnapshot = fin
    gBusy = False
    ExportSnapshot = "OK:" & ConvertFromURL(fin) & " — " & summary & "; " & (GetSystemTicks() - t0) & " мс"
    Exit Function
EH:
    ExportSnapshot = "ERR:снимок не создан: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
FAIL:
    On Error Resume Next
    If Not IsNull(td) And Not IsEmpty(td) Then td.close(True)
    If tmp <> "" Then
        If sfa.exists(tmp) Then sfa.kill(tmp)
    End If
    gBusy = False
End Function

' ================================================================ «Загрузить пакет»

' the target sheets of a batch and their input columns: name, key (| separated), 0-based column (| separated)
Private Function BatchTargets() As Variant
    BatchTargets = Array( _
        Array(SH_ADJUST, "kind|ei|qty|fact|book|place|date|reason|note", AC_KIND & "|" & AC_EI & "|" & AC_QTY & "|" & AC_FACT & "|" & AC_BOOK _
            & "|" & AC_PLACE & "|" & AC_DATE & "|" & AC_REASON & "|" & AC_NOTE), _
        Array(SH_SPECIAL, "type|event|name|article|qty|unit|date|place|category|from_who|doc|old_mark|note", XC_TYPE & "|" & XC_EVENT & "|" _
            & XC_NAME & "|" & XC_ART & "|" & XC_QTY & "|" & XC_UNIT & "|" & XC_DATE & "|" & XC_PLACE & "|" & XC_CAT & "|" & XC_WHO & "|" _
            & XC_DOC & "|" & XC_MARK & "|" & XC_NOTE), _
        Array(SH_ORDERS, "order_no|name|doc_no|invoice_no|article|qty_fact|qty_doc|qty_ordered|unit|price|amount|supplier|seller|" _
            & "date_received|date_doc|date_order|date_expected|buyer|category|assignee|place|note|lead_days", _
            "0|1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|" & OC_NOTE & "|" & (OC_NOTE + 1)))
End Function

' the note column where the batch mark «[пакет <id>]» goes (Корректировки: its own column R «Пакет»)
Private Function MarkColumn(target As String) As Integer
    Select Case target
    Case SH_ADJUST
        MarkColumn = AC_BATCH
    Case SH_SPECIAL
        MarkColumn = XC_NOTE
    Case Else
        MarkColumn = OC_NOTE
    End Select
End Function

' Reads, checks and loads a batch file into open rows of its target sheet. OK:<target>|<first row>|<last row>|<rows>|<id>
' (1-based rows), ERR:<why> (nothing loaded). A batch is checked as a whole: one bad line — nothing is loaded.
Function LoadBatch(url As String) As String
    Dim txt As String, lines As Variant, i As Long, ln As String, meta As String, id As String, target As String, src As String, sShared As String
    Dim hdr As Variant, keys As Variant, cols As Variant, tg As Variant, t As Integer, map() As Integer, nData As Long, data() As Variant
    Dim f As Variant, j As Integer, k As Integer, sh As Object, last As Long, r As Long, markCol As Integer, cnt As Long, why As String
    Dim fmt As Boolean, hasHdr As Boolean, c As Integer, ok As Boolean, s As String, wasProt As Boolean, um As Object, inCtx As Boolean
    WmsInit()
    If gBusy Then
        LoadBatch = "BUSY:операция уже выполняется"
        Exit Function
    End If
    If gState <> "CLEAN" Then
        LoadBatch = "ERR:WMS заблокирована (" & StateLine() & ") — пакет не загружается"
        Exit Function
    End If
    On Error GoTo EH
    txt = ReadTextFile(url)
    If Left(txt, 1) = Chr(65279) Then txt = Mid(txt, 2)
    lines = Split(Replace(txt, Chr(13), ""), Chr(10))
    ReDim data(IIf(UBound(lines) < 0, 0, UBound(lines)))
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Trim(ln) = "" Then
        ElseIf Left(ln, 1) = "#" Then
            f = Split(Mid(ln, 2), ";")
            Select Case LCase(Trim(f(0)))
            Case LCase(BATCH_FORMAT)
                fmt = True
            Case "id"
                If UBound(f) >= 1 Then id = Trim(f(1))
            Case "target"
                If UBound(f) >= 1 Then target = Trim(f(1))
            Case "source"
                If UBound(f) >= 1 Then src = Trim(f(1))
            Case "shared"
                If UBound(f) >= 1 Then sShared = LCase(Trim(f(1)))
            End Select
        ElseIf Not hasHdr Then
            hdr = CsvSplit(ln)
            hasHdr = True
        Else
            If nData >= BATCH_MAX_ROWS Then
                LoadBatch = "ERR:в пакете больше " & BATCH_MAX_ROWS & " строк — разделите его"
                Exit Function
            End If
            data(nData) = CsvSplit(ln)
            nData = nData + 1
        End If
    Next i
    If Not fmt Then
        LoadBatch = "ERR:это не пакет WMS: нет строки «#" & BATCH_FORMAT & "»"
        Exit Function
    End If
    If id = "" Or Len(id) > 60 Then
        LoadBatch = "ERR:у пакета нет идентификатора (#id) или он длиннее 60 символов"
        Exit Function
    End If
    For i = 1 To Len(id)
        s = Mid(id, i, 1)
        If Not ((s >= "0" And s <= "9") Or (s >= "A" And s <= "Z") Or (s >= "a" And s <= "z") Or s = "-" Or s = "_" Or s = ".") Then
            LoadBatch = "ERR:идентификатор пакета «" & id & "»: допустимы латинские буквы, цифры, «-», «_», «.»"
            Exit Function
        End If
    Next i
    tg = BatchTargets()
    t = -1
    For i = 0 To UBound(tg)
        If target = tg(i)(0) Then t = i
    Next i
    If t < 0 Then
        LoadBatch = "ERR:лист пакета (#target) «" & target & "» — пакет загружается только на «" & SH_ADJUST & "», «" & SH_SPECIAL & "» или «" & SH_ORDERS & "»"
        Exit Function
    End If
    If Not hasHdr Or nData = 0 Then
        LoadBatch = "ERR:в пакете нет строк"
        Exit Function
    End If
    keys = Split(tg(t)(1), "|")
    cols = Split(tg(t)(2), "|")
    ReDim map(UBound(hdr))
    For j = 0 To UBound(hdr)
        map(j) = -1
        For k = 0 To UBound(keys)
            If LCase(Trim(hdr(j))) = keys(k) Then map(j) = CInt(cols(k))
        Next k
        If map(j) < 0 Then
            LoadBatch = "ERR:колонка пакета «" & hdr(j) & "» не является колонкой ввода листа «" & target & "» (допустимы: " & Replace(tg(t)(1), "|", ", ") & ")"
            Exit Function
        End If
        For k = 0 To j - 1
            If map(k) = map(j) Then
                LoadBatch = "ERR:колонка пакета «" & hdr(j) & "» повторяется"
                Exit Function
            End If
        Next k
    Next j
    For i = 0 To nData - 1
        If UBound(data(i)) <> UBound(hdr) Then
            LoadBatch = "ERR:строка " & (i + 1) & " пакета: полей " & (UBound(data(i)) + 1) & ", а колонок " & (UBound(hdr) + 1)
            Exit Function
        End If
        ok = False
        For j = 0 To UBound(hdr)
            If Trim(data(i)(j)) <> "" Then ok = True
            If Left(data(i)(j), 1) = "=" Then
                LoadBatch = "ERR:строка " & (i + 1) & " пакета: формулы не допускаются («" & Left(data(i)(j), 30) & "»)"
                Exit Function
            End If
        Next j
        If Not ok Then
            LoadBatch = "ERR:строка " & (i + 1) & " пакета пуста"
            Exit Function
        End If
    Next i
    ' one batch — once: its mark is searched in the mark column of the target (a formula of the Calc engine)
    sh = gDoc.Sheets.getByName(target)
    markCol = MarkColumn(target)
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        s = "$'" & target & "'." & ColName(markCol) & "$2:" & ColName(markCol) & "$" & (last + 1)
        If target = SH_ADJUST Then
            s = WmsOrders.ScratchEval("COUNTIF(" & s & ";""" & id & """)", ok)
        Else
            s = WmsOrders.ScratchEval("COUNTIF(" & s & ";""*[пакет " & id & "]*"")", ok)
        End If
        If Not ok Or Not IsDigits(s) Then
            LoadBatch = "ERR:не удалось проверить, загружался ли пакет " & id & " (формула движка Calc) — пакет не загружен"
            Exit Function
        End If
        If CLng(s) > 0 Then
            LoadBatch = "ERR:пакет " & id & " уже загружен на лист «" & target & "» (строк с его отметкой: " & s & ") — повторно не загружается"
            Exit Function
        End If
    End If
    ' the rows below the last used row: only the input columns are written (the change handler is off, the rows are
    ' checked and posted by the usual code afterwards); one undo step removes the whole loaded block
    r = last + 1
    gBusy = True
    WmsAdjust.HandlerOff(True)
    WmsSpecial.SpecialHandlerOff(True)
    WmsOrders.HandlerOff(True)
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: загрузка пакета " & id)
    inCtx = True
    For i = 0 To nData - 1
        For j = 0 To UBound(hdr)
            s = data(i)(j)
            ' a date of the contract (ГГГГ-ММ-ДД) is entered as the user types it (дд.мм.гггг)
            If Left(LCase(Trim(hdr(j))), 4) = "date" And Len(s) = 10 And Mid(s, 5, 1) = "-" And Mid(s, 8, 1) = "-" Then
                s = Mid(s, 9, 2) & "." & Mid(s, 6, 2) & "." & Left(s, 4)
            End If
            If s <> "" Then sh.getCellByPosition(map(j), r + i).setString(s)
        Next j
        If target = SH_ADJUST Then
            sh.getCellByPosition(AC_BATCH, r + i).setString(id)
        Else
            s = sh.getCellByPosition(markCol, r + i).getString()
            sh.getCellByPosition(markCol, r + i).setString(Trim(s & " [пакет " & id & "]"))
        End If
    Next i
    um.leaveUndoContext()
    inCtx = False
    WmsAdjust.HandlerOff(False)
    WmsSpecial.SpecialHandlerOff(False)
    WmsOrders.HandlerOff(False)
    gBusy = False
    LoadBatch = "OK:" & target & "|" & (r + 1) & "|" & (r + nData) & "|" & nData & "|" & id & "|" & IIf(sShared = "yes" Or sShared = "1" Or sShared = "да", "1", "0")
    Exit Function
EH:
    LoadBatch = "ERR:пакет не загружен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    WmsAdjust.HandlerOff(False)
    WmsSpecial.SpecialHandlerOff(False)
    WmsOrders.HandlerOff(False)
    gBusy = False
End Function

' the loaded rows (0-based r0..r1) of the target posted by the usual code of that sheet (each row its own operation)
Function PostBatchRows(target As String, r0 As Long, r1 As Long, bShared As Boolean) As String
    Select Case target
    Case SH_ADJUST
        PostBatchRows = WmsAdjust.AdjustPostRange(r0, r1)
    Case SH_SPECIAL
        PostBatchRows = WmsSpecial.SpecialPostRange(r0, r1, bShared, False)
    Case Else
        PostBatchRows = WmsReceipt.ReceiptPostRange(r0, r1)
    End Select
End Function

Private Function ColName(c As Integer) As String
    If c < 26 Then
        ColName = "$" & Chr(65 + c)
    Else
        ColName = "$" & Chr(64 + c \ 26) & Chr(65 + c Mod 26)
    End If
End Function

' ================================================================ buttons of «Главная»

Sub BtnExport(Optional oEvent As Variant)
    Dim res As String
    WmsInit()
    res = ExportSnapshot()
    gUiLastMsg = res
    If Left(res, 3) = "OK:" Then
        UiInfo("Снимок для инструментов создан:" & Chr(10) & Mid(res, 4))
        UiRefresh("«Экспорт для инструментов»: " & Mid(res, 4))
    Else
        UiMessage("Экспорт не выполнен: " & Mid(res, InStr(res, ":") + 1))
    End If
    gUiLastMsg = res
End Sub

Sub BtnLoadBatch(Optional oEvent As Variant)
    Dim url As String, res As String, p As Variant, fp As Object, target As String, r0 As Long, r1 As Long, n As Long, a As Variant
    Dim parts As Variant, msg As String
    WmsInit()
    If gBatchFileSet Then
        url = gBatchFile
        gBatchFileSet = False
    ElseIf gUiAuto <> 0 Then
        gUiLastMsg = "SKIP:файл пакета не выбран"
        Exit Sub
    Else
        fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
        fp.setTitle("Пакет инструмента WMS (" & BATCH_FORMAT & ")")
        fp.appendFilter("Пакет WMS (*.csv)", "*.csv")
        If gBaseDir <> "" Then fp.setDisplayDirectory(gBaseDir)
        If fp.execute() <> 1 Then
            gUiLastMsg = "SKIP:файл пакета не выбран"
            Exit Sub
        End If
        url = fp.getSelectedFiles()(0)
    End If
    res = LoadBatch(url)
    If Left(res, 3) <> "OK:" Then
        UiMessage("Пакет не загружен: " & Mid(res, InStr(res, ":") + 1))
        gUiLastMsg = res
        Exit Sub
    End If
    p = Split(Mid(res, 4), "|")
    target = p(0)
    r0 = CLng(p(1)) - 1
    r1 = CLng(p(2)) - 1
    n = CLng(p(3))
    UiRefresh("«Загрузить пакет»: " & p(4) & " — строк " & n & " на лист «" & target & "» (строки " & p(1) & "–" & p(2) & ")")
    If Not UiConfirm("Пакет " & p(4) & " загружен на лист «" & target & "»: строк " & n & " (строки " & p(1) & "–" & p(2) & ")." & Chr(10) & Chr(10) _
        & "Провести эти строки сейчас? Каждая строка проводится своей операцией со всеми обычными проверками; строка с ошибкой не " _
        & "проводится (причина — в «Контроль»)." & Chr(10) & "«Нет» — строки останутся непроведёнными для просмотра.") Then
        gUiLastMsg = "OK:загружено " & n & ", не проведено"
        Exit Sub
    End If
    parts = Split(PostBatchRows(target, r0, r1, p(5) = "1"), Chr(10))
    a = Split(parts(0), ";")
    msg = "Пакет " & p(4) & ": проведено " & Val(Mid(a(0), 4)) & ", отклонено с ошибкой " & Val(Mid(a(1), 5)) & ", пропущено " & Val(Mid(a(2), 6)) & "."
    If parts(1) <> "" Then msg = msg & Chr(10) & "Первая ошибка: " & parts(1)
    If Val(Mid(a(3), 6)) > 0 Then msg = msg & Chr(10) & "ОБРАБОТКА ОСТАНОВЛЕНА — системная ошибка WMS: " & parts(2)
    UiRefresh("«Загрузить пакет»: " & msg)
    If Val(Mid(a(1), 5)) > 0 Or Val(Mid(a(3), 6)) > 0 Then UiMessage(msg) Else UiInfo(msg)
    gUiLastMsg = Join(parts, Chr(10))
End Sub

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestBatchFile(sPath As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestBatchFile = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gBatchFile = ConvertToURL(sPath)
    gBatchFileSet = True
    TestBatchFile = "OK"
End Function

Function TestLastSnapshot() As String
    TestLastSnapshot = ConvertFromURL(gLastSnapshot)
End Function
