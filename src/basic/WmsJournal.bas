' WmsJournal — external append-only journal (spec §20, D-027, D-029).
' Line: J1;seq;timestamp;instance;type;fields...;END;length   (length = characters before ";END;")
' UTF-8, LF, one file per month (WMS_journal_YYYY-MM.csv), seq strictly continuous.
' The saved position in the book is "file|offset|seq" and is trusted only after checking the file content.
Option Explicit

' results of the last ReadJournal
Global gJR_Error As String
Global gJR_Full As Boolean
Global gJR_Note As String
Global gJR_N As Long
Global gJR_Seq() As Long
Global gJR_Type() As String
Global gJR_Ts() As String
Global gJR_Fields() As Variant
Global gJR_MaxSeq As Long
Global gJR_MaxFile As String
Global gJR_HasLast As Boolean
Global gJR_FirstSeq As Long
Global gJR_FirstType As String
Global gJR_Damaged As Long
Global gJR_Torn As Boolean
Global gJR_Foreign As Long
Global gJR_Order As String
Global gJR_Lines As Long
Global gJR_Files As Long
' 1.0: the mark («time/length») of the journal entry with the book's last seq, as read now ("" — not read)
Global gJR_LastMark As String

Function JournalFileName(d As Date) As String
    JournalFileName = JOURNAL_PREFIX & Format(d, "YYYY-MM") & JOURNAL_EXT
End Function

' journal file names (not URLs), oldest first
Function ListFiles() As Variant
    Dim urls As Variant, names() As String, n As Long, i As Long, j As Long, nm As String, t As String
    ListFiles = Array()
    If gJDir = "" Then Exit Function
    If Not gSfa.exists(gJDir) Then Exit Function
    urls = gSfa.getFolderContents(gJDir, False)
    n = -1
    For i = 0 To UBound(urls)
        nm = WmsCore.FileNameOf(urls(i))
        If Left(nm, Len(JOURNAL_PREFIX)) = JOURNAL_PREFIX And Right(nm, Len(JOURNAL_EXT)) = JOURNAL_EXT Then
            n = n + 1
            ReDim Preserve names(n)
            names(n) = nm
        End If
    Next i
    If n < 0 Then Exit Function
    For i = 1 To n
        t = names(i)
        j = i - 1
        Do While j >= 0
            If names(j) <= t Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = t
    Next i
    ListFiles = names
End Function

' fields must already be escaped (WmsCore.PlanJournalFields / PlanField do it)
Function BuildLine(seq As Long, sType As String, fields As Variant) As String
    Dim body As String, i As Long
    body = "J1;" & seq & ";" & WmsCore.TS() & ";" & SysStr(SK_INSTANCE) & ";" & sType
    For i = 0 To UBound(fields)
        body = body & ";" & fields(i)
    Next i
    BuildLine = body & ";END;" & Len(body)
End Function

' 1.0: the mark of an entry — its time and length («2026-10-15T09:30:12/245»). The book keeps the mark of its last entry
' in the journal position: an entry with the same seq but another mark means the book and the journal belong to two
' copies of the WMS that went separate ways (a copy worked on elsewhere and put back)
Function LineMark(ByVal ln As String) As String
    Dim p As Long, a As Variant
    ln = StripEol(ln)
    If Left(ln, 1) = Chr(10) Then ln = Mid(ln, 2)
    p = EndPos(ln)
    If p = 0 Or Left(ln, 3) <> "J1;" Then Exit Function
    a = Split(Left(ln, p - 1), ";", 4)
    If UBound(a) < 3 Then Exit Function
    LineMark = a(2) & "/" & Mid(ln, p + 5)
End Function

' position of ";END;" (it can only occur once, near the end: field values are escaped)
Private Function EndPos(ln As String) As Long
    Dim st As Long
    st = Len(ln) - 20
    If st < 1 Then st = 1
    EndPos = InStr(st, ln, ";END;")
End Function

Private Function StripEol(ByVal ln As String) As String
    Do While Len(ln) > 0
        If Right(ln, 1) = Chr(10) Or Right(ln, 1) = Chr(13) Then ln = Left(ln, Len(ln) - 1) Else Exit Do
    Loop
    StripEol = ln
End Function

' header only (fast): False for a damaged line
Function ParseHeader(ByVal ln As String, ByRef seq As Long, ByRef inst As String, ByRef sType As String) As Boolean
    Dim p As Long, body As String, sLen As String, p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long, s As String
    ln = StripEol(ln)
    If Left(ln, 3) <> "J1;" Then Exit Function
    p = EndPos(ln)
    If p = 0 Then Exit Function
    body = Left(ln, p - 1)
    sLen = Mid(ln, p + 5)
    If Not WmsCore.IsDigits(sLen) Or Len(sLen) > 9 Then Exit Function
    If CLng(sLen) <> Len(body) Then Exit Function
    p1 = 3
    p2 = InStr(p1 + 1, body, ";")
    If p2 = 0 Then Exit Function
    p3 = InStr(p2 + 1, body, ";")
    If p3 = 0 Then Exit Function
    p4 = InStr(p3 + 1, body, ";")
    If p4 = 0 Then Exit Function
    p5 = InStr(p4 + 1, body, ";")
    If p5 = 0 Then p5 = Len(body) + 1
    s = Mid(body, p1 + 1, p2 - p1 - 1)
    If Not WmsCore.IsDigits(s) Or Len(s) > 9 Then Exit Function
    seq = CLng(s)
    inst = Mid(body, p3 + 1, p4 - p3 - 1)
    sType = Mid(body, p4 + 1, p5 - p4 - 1)
    ParseHeader = True
End Function

' full parse; fields = array of "name=value" strings (still escaped)
Function ParseLine(ByVal ln As String, ByRef seq As Long, ByRef ts As String, ByRef inst As String, ByRef sType As String, ByRef fields As Variant) As Boolean
    Dim a As Variant, f() As String, i As Long, p As Long, body As String
    If Not ParseHeader(ln, seq, inst, sType) Then Exit Function
    ln = StripEol(ln)
    p = EndPos(ln)
    body = Left(ln, p - 1)
    a = Split(body, ";")
    ts = a(2)
    If UBound(a) >= 5 Then
        ReDim f(UBound(a) - 5)
        For i = 5 To UBound(a)
            f(i - 5) = a(i)
        Next i
        fields = f
    Else
        fields = Array()
    End If
    ParseLine = True
End Function

' ---------------------------------------------------------------- append (the commit point)
' 1 = appended, 0 = not appended (the caller rolls back), 2 = cannot tell (the caller blocks; the next start decides)

Function AppendLine(sLine As String, seq As Long, ByRef why As String) As Integer
    Dim fn As String, url As String, sz As Double, expect As Double, prefix As String, xs As Object, tos As Object, nsz As Double
    Dim txt As String, cut As Long, z As Integer
    fn = JournalFileName(Now())
    ' never a file that sorts before the one already in use (the computer's clock went back): continue the newest one
    If fn < gJFile Then fn = gJFile
    url = gJDir & fn
    On Error GoTo EH
    If gJDir = "" Then
        why = "книга не сохранена в файл — журнал недоступен"
        AppendLine = 0
        Exit Function
    End If
    If Not gSfa.exists(gJDir) Then gSfa.createFolder(gJDir)
    If gSfa.exists(url) Then sz = gSfa.getSize(url) Else sz = 0
    If fn = gJFile Then expect = gJSize Else expect = 0
    If sz <> expect Then
        why = "журнал «" & fn & "» изменён другим процессом (ожидалось " & NumStr(expect) & " байт, найдено " & NumStr(sz) & ")"
        WmsCore.SetBlocked("JOURNAL_FOREIGN_WRITE", "журнал изменён другим процессом — проведение остановлено")
        AppendLine = 0
        Exit Function
    End If
    prefix = ""
    If sz > 0 Then
        ' a torn last line (crash during a previous append) is terminated, never continued
        If LastByte(url) <> 10 Then prefix = Chr(10)
    End If
    xs = gSfa.openFileReadWrite(url)
    xs.seek(xs.getLength())
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(xs.getOutputStream())
    tos.setEncoding("UTF-8")
    txt = prefix & sLine & Chr(10)
    cut = WmsCore.FaultCut(Len(txt))
    If cut > 0 Then
        ' test seam (TEST books only): the write stops part-way, then fails
        tos.writeString(Left(txt, cut))
        tos.closeOutput()
        z = 0
        z = 1 / z
    End If
    tos.writeString(txt)
    tos.closeOutput()
    nsz = gSfa.getSize(url)
    gJFile = fn
    gJSize = nsz
    WmsCore.SysPutStr(SK_JPOS, fn & "|" & NumStr(nsz) & "|" & seq & "|" & LineMark(sLine))
    AppendLine = 1
    Exit Function
EH:
    why = "ошибка записи журнала: " & Error$
    AppendLine = VerifyAfterError(url, fn, sz, prefix & sLine & Chr(10), seq, LineMark(sLine))
End Function

Private Function VerifyAfterError(url As String, fn As String, sz As Double, expected As String, seq As Long, sMark As String) As Integer
    Dim nsz As Double, got As String
    On Error GoTo UNKNOWN
    If Not gSfa.exists(url) Then
        VerifyAfterError = 0
        Exit Function
    End If
    nsz = gSfa.getSize(url)
    If nsz = sz Then
        VerifyAfterError = 0
        Exit Function
    End If
    got = ReadFrom(url, sz)
    gJFile = fn
    gJSize = nsz
    If got = expected Then
        WmsCore.SysPutStr(SK_JPOS, fn & "|" & NumStr(nsz) & "|" & seq & "|" & sMark)
        VerifyAfterError = 1
    ElseIf got = Left(expected, Len(expected) - 1) Then
        ' the whole line without its LF: every reader accepts it, so the operation is committed
        ' (the next append terminates the line; the saved position stays behind and is re-established by reading)
        VerifyAfterError = 1
    ElseIf Left(expected, Len(got)) = got Then
        ' a part of the line only: the torn line is terminated before the next append and reported as damaged
        VerifyAfterError = 0
    Else
        ' bytes that are not ours: cannot tell
        VerifyAfterError = 2
    End If
    Exit Function
UNKNOWN:
    VerifyAfterError = 2
End Function

Private Function LastByte(url As String) As Integer
    Dim xs As Object, buf(), n As Long, ln As Double
    LastByte = -1
    ln = gSfa.getSize(url)
    If ln <= 0 Then Exit Function
    xs = gSfa.openFileRead(url)
    xs.seek(ln - 1)
    n = xs.readBytes(buf, 1)
    xs.closeInput()
    If n = 1 Then LastByte = buf(0) Else LastByte = -1
End Function

' text of the file from byte offset off to the end
Private Function ReadFrom(url As String, off As Double) As String
    Dim xs As Object, tis As Object
    xs = gSfa.openFileRead(url)
    xs.seek(off)
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(xs)
    tis.setEncoding("UTF-8")
    ReadFrom = tis.readString(Array(), False)
    tis.closeInput()
End Function

' ---------------------------------------------------------------- saved position (spec §20)

' «файл|смещение|seq» (up to 0.8) or «файл|смещение|seq|отметка» (1.0, LineMark of the entry seq)
Function PosFile() As String
    Dim a As Variant
    a = Split(SysStr(SK_JPOS), "|")
    If UBound(a) = 2 Or UBound(a) = 3 Then PosFile = a(0) Else PosFile = ""
End Function

Function PosOffset() As Double
    Dim a As Variant
    a = Split(SysStr(SK_JPOS), "|")
    If UBound(a) = 2 Or UBound(a) = 3 Then
        If WmsCore.IsDigits(CStr(a(1))) Then PosOffset = Val(a(1)) Else PosOffset = -1
    Else
        PosOffset = -1
    End If
End Function

Function PosSeq() As Long
    Dim a As Variant
    a = Split(SysStr(SK_JPOS), "|")
    PosSeq = -1
    If UBound(a) = 2 Or UBound(a) = 3 Then
        If WmsCore.IsDigits(CStr(a(2))) Then PosSeq = CLng(a(2))
    End If
End Function

' the mark kept with the position ("" — a position saved before 1.0)
Function PosMark() As String
    Dim a As Variant
    a = Split(SysStr(SK_JPOS), "|")
    If UBound(a) = 3 Then PosMark = a(3)
End Function

' the journal line that ends exactly at byte off must be entry seq ("J1;<seq>;"), reading at most 4 KB
Function PosValid(url As String, off As Double, seq As Long) As Boolean
    Dim sMark As String
    PosValid = PosCheck(url, off, seq, sMark)
End Function

' PosValid that also gives the mark of that line (LineMark; its time and length are plain ASCII, read from the bytes)
Function PosCheck(url As String, off As Double, seq As Long, ByRef sMark As String) As Boolean
    Dim xs As Object, st As Double, n As Long, k As Long, j As Long, want As String, buf(), e As Long, sTs As String, sLen As String
    sMark = ""
    On Error GoTo EH
    If seq < 1 Or off <= 0 Then Exit Function
    If Not gSfa.exists(url) Then Exit Function
    If off > gSfa.getSize(url) Then Exit Function
    st = off - JPOS_VERIFY_WINDOW
    If st < 0 Then st = 0
    xs = gSfa.openFileRead(url)
    xs.seek(st)
    n = xs.readBytes(buf, CLng(off - st))
    xs.closeInput()
    If n <> CLng(off - st) Then Exit Function
    If buf(n - 1) <> 10 Then Exit Function
    k = n - 2
    Do While k >= 0
        If buf(k) = 10 Then Exit Do
        k = k - 1
    Loop
    If k < 0 And st > 0 Then Exit Function
    want = "J1;" & seq & ";"
    For j = 1 To Len(want)
        If k + j > n - 1 Then Exit Function
        If buf(k + j) <> Asc(Mid(want, j, 1)) Then Exit Function
    Next j
    ' the time: 19 characters after the seq; the length: the digits after the last ";" of the line
    For j = k + Len(want) + 1 To k + Len(want) + 19
        If j > n - 2 Then Exit For
        sTs = sTs & Chr(buf(j))
    Next j
    e = n - 2
    If e >= 0 Then
        If buf(e) = 13 Then e = e - 1
    End If
    Do While e > k
        If buf(e) < 48 Or buf(e) > 57 Then Exit Do
        sLen = Chr(buf(e)) & sLen
        e = e - 1
    Loop
    If Len(sTs) = 19 And sLen <> "" Then sMark = sTs & "/" & sLen
    PosCheck = True
    Exit Function
EH:
    PosCheck = False
End Function

' ---------------------------------------------------------------- reading

Private Sub ResetResults()
    gJR_Error = ""
    gJR_Full = False
    gJR_Note = ""
    gJR_N = 0
    ReDim gJR_Seq(15)
    ReDim gJR_Type(15)
    ReDim gJR_Ts(15)
    ReDim gJR_Fields(15)
    gJR_MaxSeq = 0
    gJR_MaxFile = ""
    gJR_HasLast = False
    gJR_FirstSeq = 0
    gJR_FirstType = ""
    gJR_Damaged = 0
    gJR_Torn = False
    gJR_Foreign = 0
    gJR_Order = ""
    gJR_Lines = 0
    gJR_Files = 0
    gJR_LastMark = ""
End Sub

Private Sub AddTail(seq As Long, sType As String, ts As String, fields As Variant)
    If gJR_N > UBound(gJR_Seq) Then
        ReDim Preserve gJR_Seq(2 * gJR_N + 1)
        ReDim Preserve gJR_Type(2 * gJR_N + 1)
        ReDim Preserve gJR_Ts(2 * gJR_N + 1)
        ReDim Preserve gJR_Fields(2 * gJR_N + 1)
    End If
    gJR_Seq(gJR_N) = seq
    gJR_Type(gJR_N) = sType
    gJR_Ts(gJR_N) = ts
    gJR_Fields(gJR_N) = fields
    gJR_N = gJR_N + 1
End Sub

' Reads entries after lastSeq, choosing the cheapest read that is still safe:
'   1. the saved position is confirmed by the file content → only the part after it;
'   2. otherwise the file named in the saved position, from its start, and the later files — accepted only when they
'      hold the book's last operation and the sequence in them is continuous (entries are appended to files in name
'      order, see AppendLine, so nothing after the book's last operation can be in an earlier file);
'   3. otherwise every journal file.
' bBaseline: take the size of the file the next append goes to as the ownership baseline (startup and recovery only —
' diagnostics must not launder a foreign write).
Sub ReadJournal(lastSeq As Long, bForceFull As Boolean, bBaseline As Boolean)
    Dim files As Variant, i As Long, idx As Long, pf As String, po As Double, ps As Long, myInst As String, cur As String
    Dim why As String, clockNote As String
    ResetResults()
    On Error GoTo EH
    If gJDir = "" Then
        gJR_Error = "книга не сохранена в файл"
        Exit Sub
    End If
    myInst = SysStr(SK_INSTANCE)
    files = ListFiles()
    If bBaseline Then
        cur = JournalFileName(Now())
        If UBound(files) >= 0 Then
            If files(UBound(files)) > cur Then
                clockNote = "дата компьютера (" & Format(Now(), "YYYY-MM-DD") & ") раньше последнего файла журнала «" _
                    & files(UBound(files)) & "» — записи продолжаются в нём; проверьте часы компьютера"
                cur = files(UBound(files))
            End If
        End If
        gJFile = cur
        gJSize = 0
        If gSfa.exists(gJDir & cur) Then gJSize = gSfa.getSize(gJDir & cur)
    End If
    If UBound(files) < 0 Then
        gJR_HasLast = (lastSeq = 0)
        gJR_Full = True
        gJR_Note = clockNote
        Exit Sub
    End If
    idx = -1
    If Not bForceFull Then
        pf = PosFile()
        po = PosOffset()
        ps = PosSeq()
        For i = 0 To UBound(files)
            If files(i) = pf Then idx = i
        Next i
        If SysStr(SK_JPOS) = "" Then
            If lastSeq > 0 Then why = "позиция не сохранена"
        ElseIf pf = "" Or po < 0 Or ps < 0 Then
            why = "значение «" & SysStr(SK_JPOS) & "» повреждено"
        ElseIf idx < 0 Then
            why = "файла «" & pf & "» нет"
        ElseIf ps <> lastSeq Then
            why = "в ней seq " & ps & ", а в книге " & lastSeq
        ElseIf PosCheck(gJDir & pf, po, ps, gJR_LastMark) Then
            gJR_HasLast = True
            gJR_MaxSeq = lastSeq
            ReadRange(files, idx, po, lastSeq, lastSeq, myInst)
            gJR_Note = clockNote
            Exit Sub
        Else
            why = "файл изменён: окончания строк, обрезка или правка"
        End If
        If why <> "" Then why = "позиция журнала не подтвердилась (" & why & ")"
        If idx >= 0 And lastSeq > 0 Then
            ReadRange(files, idx, 0, -1, lastSeq, myInst)
            If gJR_HasLast And gJR_Order = "" Then
                gJR_Note = JoinNotes(clockNote, why & " — прочитан файл " & pf & IIf(idx < UBound(files), " и последующие", ""))
                Exit Sub
            End If
            ResetResults()
        End If
    End If
    gJR_Full = True
    ReadRange(files, 0, 0, -1, lastSeq, myInst)
    If why <> "" Then gJR_Note = JoinNotes(clockNote, why & " — журнал прочитан целиком") Else gJR_Note = clockNote
    Exit Sub
EH:
    gJR_Error = "ошибка чтения журнала: " & Error$
End Sub

Private Function JoinNotes(a As String, b As String) As String
    If a = "" Then
        JoinNotes = b
    ElseIf b = "" Then
        JoinNotes = a
    Else
        JoinNotes = a & "; " & b
    End If
End Function

' Reads files(startIdx..), the first one from byte offset startOff, one call per file (then split into lines).
' prevSeq: -1 = the first entry starts the sequence; otherwise the seq that the next entry must follow.
Private Sub ReadRange(files As Variant, startIdx As Long, startOff As Double, ByVal prevSeq As Long, lastSeq As Long, myInst As String)
    Dim i As Long, j As Long, fn As String, url As String, txt As String, a As Variant, ln As String, isLastFile As Boolean
    Dim seq As Long, inst As String, sType As String, ts As String, fields As Variant
    For i = startIdx To UBound(files)
        fn = files(i)
        url = gJDir & fn
        gJR_Files = gJR_Files + 1
        isLastFile = (i = UBound(files))
        If i = startIdx Then
            txt = ReadFrom(url, startOff)
        Else
            txt = ReadFrom(url, 0)
        End If
        a = Split(txt, Chr(10))
        For j = 0 To UBound(a)
            ln = StripEol(a(j))
            If ln <> "" Then
                gJR_Lines = gJR_Lines + 1
                If Not ParseHeader(ln, seq, inst, sType) Then
                    ' an unterminated last line of the last file is a torn append; anything else is damage
                    If isLastFile And j = UBound(a) Then
                        gJR_Torn = True
                    Else
                        gJR_Damaged = gJR_Damaged + 1
                    End If
                ElseIf inst <> myInst Then
                    gJR_Foreign = gJR_Foreign + 1
                Else
                    If prevSeq = -1 Then
                        gJR_FirstSeq = seq
                        gJR_FirstType = sType
                    ElseIf seq <> prevSeq + 1 And gJR_Order = "" Then
                        If seq <= prevSeq Then
                            gJR_Order = "seq " & seq & " после " & prevSeq & " (повтор или нарушен порядок) в файле " & fn
                        Else
                            gJR_Order = "разрыв: после seq " & prevSeq & " идёт " & seq & " в файле " & fn
                        End If
                    End If
                    prevSeq = seq
                    If seq > gJR_MaxSeq Then
                        gJR_MaxSeq = seq
                        gJR_MaxFile = fn
                    End If
                    If seq = lastSeq Then
                        gJR_HasLast = True
                        gJR_LastMark = LineMark(ln)
                    End If
                    If seq > lastSeq Then
                        If Not ParseLine(ln, seq, ts, inst, sType, fields) Then
                            gJR_Damaged = gJR_Damaged + 1
                        Else
                            AddTail(seq, sType, ts, fields)
                        End If
                    End If
                End If
            End If
        Next j
    Next i
End Sub

' 1.0: «Сделать рабочим файлом» with a new journal while the journal is behind the book or belongs to another copy: the
' journal files are moved into WMS_Journal/старый_журнал_<дата_время> — copied, checked by size, and only then removed
' (nothing is deleted). sDir — the folder (for the message). "" — done; otherwise why not, and no file was removed
Function ArchiveFiles(ByRef sDir As String) As String
    Dim files As Variant, i As Long, dst As String
    On Error GoTo EH
    files = ListFiles()
    dst = ConvertToURL(ConvertFromURL(gJDir) & "старый_журнал_" & Format(Now(), "YYYY-MM-DD") & "_" & Format(Now(), "HHMMSS")) & "/"
    If gSfa.exists(dst) Then dst = Left(dst, Len(dst) - 1) & "_" & WmsCore.NewId(4) & "/"
    gSfa.createFolder(dst)
    For i = 0 To UBound(files)
        gSfa.copy(gJDir & files(i), dst & files(i))
        If gSfa.getSize(dst & files(i)) <> gSfa.getSize(gJDir & files(i)) Then
            ArchiveFiles = "копия файла журнала «" & files(i) & "» неполная — проверьте место на диске; журнал не тронут"
            Exit Function
        End If
    Next i
    For i = 0 To UBound(files)
        gSfa.kill(gJDir & files(i))
    Next i
    sDir = ConvertFromURL(dst)
    Exit Function
EH:
    ArchiveFiles = "не удалось перенести журнал: " & Error$
End Function

' after «Восстановить» has replayed the whole tail read by ReadJournal: the saved position moves to its last entry
' (only when that entry ends its file with a LF; otherwise the next start re-establishes it by reading)
Sub PinPosition(seq As Long)
    Dim url As String, sz As Double, sMark As String
    On Error GoTo EH
    If gJR_MaxFile = "" Or seq <> gJR_MaxSeq Then Exit Sub
    url = gJDir & gJR_MaxFile
    sz = gSfa.getSize(url)
    If PosCheck(url, sz, seq, sMark) Then WmsCore.SysPutStr(SK_JPOS, gJR_MaxFile & "|" & NumStr(sz) & "|" & seq & IIf(sMark <> "", "|" & sMark, ""))
    Exit Sub
EH:
End Sub

' first-entry rule of a full read: the history may start later than seq 1 only with a START record or when older
' files were removed below lastSeq; entries above lastSeq may never be missing
Function StartProblem(lastSeq As Long) As String
    StartProblem = ""
    If Not gJR_Full Then Exit Function
    If gJR_FirstSeq <= 1 Then Exit Function
    If gJR_FirstType = "START" Then Exit Function
    If gJR_FirstSeq - 1 > lastSeq Then
        StartProblem = "журнал начинается с seq " & gJR_FirstSeq & ", а в книге последняя операция " & lastSeq & " — записи между ними отсутствуют"
    End If
End Function
