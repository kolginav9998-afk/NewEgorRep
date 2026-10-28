' WmsSafe — 1.0 (этап J): надёжность. «Копия на флешку» с проверкой и напоминание о ней (Н1); мастер «Вернуться к копии»
' (Н2); зеркало журнала — настройка, состояние, восстановление журнала из зеркала (Н3); контрольная цепочка журнала —
' проверка (Н4); ежедневная автопроверка (Н5); точки восстановления (Н6); разбор отложенных операций (Н7); сторож часов
' (Н8) и диска (Н9). Состояние этих функций (дата копии на флешку, папка зеркала, докуда проверена цепочка, итог проверки
' дня) — файл WMS_State/wms_state.txt рядом с книгой: это не учёт, книга от него не меняется. Library WmsPlus.
Option Explicit

' the folder of the state of these functions next to the book (not the journal, not the copies)
Public Const SAFE_DIR = "WMS_State"
Public Const SAFE_STATE = "wms_state.txt"
' the copies of the book made by 1.0 — «имя;операций в книге;вид;пометка;когда» (WmsBackup.BackupNow adds a line)
Public Const SAFE_COPIES = "wms_copies.txt"
Public Const FLASH_DIR = "WMS_FLASH"
Public Const FLASH_KEEP = 8
Public Const FLASH_REMIND_DAYS = 7
' Н2: the order of a return to a copy (written by the working book, carried out by the copy) and the order after it (the
' restored book replays the journal and checks itself)
Public Const RESTORE_ORDER = "wms_restore_order.txt"
Public Const AFTER_RESTORE = "wms_after_restore.txt"
' Н9: below this free space (MB) no copy is made; the warning starts at three times the size of the book (at least 1 GB)
Public Const DISK_MIN_MB = 100
Public Const DISK_WARN_MB = 1024

Global gSfState As String
Global gSfStateOk As Boolean
Global gSfStateDir As String
' test seams: the folder «Копия на флешку» takes without the window; the free space reported (MB, 0 — measured)
Global gSfFlashTest As String
Global gSfDiskTest As Double
' Н2: the async step of a return to a copy (the copy carries it out after the working book's macro ended)
Global gSfCb As Object
Global gSfCbKind As String

' ================================================================ the state file

' the folder WMS_State of the book (URL with «/»; "" — the book is not in a file)
Function SafeDir() As String
    If gBaseDir <> "" Then SafeDir = gBaseDir & SAFE_DIR & "/"
End Function

Private Function StateUrl() As String
    StateUrl = SafeDir() & SAFE_STATE
End Function

Private Sub StLoad()
    gSfState = ""
    gSfStateDir = gBaseDir
    gSfStateOk = True
    On Error GoTo EH
    If gBaseDir = "" Then Exit Sub
    If gSfa.exists(StateUrl()) Then gSfState = Replace(WmsAudit.ReadText(StateUrl()), Chr(13), "")
EH:
End Sub

' the value of a key of the state file ("" — none)
Function StGet(sKey As String) As String
    Dim a As Variant, i As Long
    If Not gSfStateOk Or gSfStateDir <> gBaseDir Then StLoad()
    a = Split(gSfState, Chr(10))
    For i = 0 To UBound(a)
        If Left(a(i), Len(sKey) + 1) = sKey & "=" Then
            StGet = Mid(a(i), Len(sKey) + 2)
            Exit Function
        End If
    Next i
End Function

Sub StSet(sKey As String, ByVal sVal As String)
    Dim a As Variant, i As Long, s As String, done As Boolean
    If Not gSfStateOk Or gSfStateDir <> gBaseDir Then StLoad()
    sVal = Replace(Replace(sVal, Chr(10), " "), Chr(13), " ")
    a = Split(gSfState, Chr(10))
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            If Left(a(i), Len(sKey) + 1) = sKey & "=" Then
                If Not done Then s = s & sKey & "=" & sVal & Chr(10)
                done = True
            Else
                s = s & a(i) & Chr(10)
            End If
        End If
    Next i
    If Not done Then s = s & sKey & "=" & sVal & Chr(10)
    gSfState = s
    WriteText(StateUrl(), s)
End Sub

' a small text file written whole (through a temporary file: a failed write leaves the old one)
Function WriteText(url As String, s As String) As Boolean
    Dim tmp As String, xs As Object, tos As Object, sDir As String
    On Error GoTo EH
    sDir = FolderOf(url)
    If Not gSfa.exists(sDir) Then gSfa.createFolder(sDir)
    tmp = url & ".tmp"
    If gSfa.exists(tmp) Then gSfa.kill(tmp)
    xs = gSfa.openFileWrite(tmp)
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(xs.getOutputStream())
    tos.setEncoding("UTF-8")
    tos.writeString(s)
    tos.closeOutput()
    If gSfa.exists(url) Then gSfa.kill(url)
    gSfa.move(tmp, url)
    WriteText = True
EH:
End Function

' ================================================================ Н9 the disk, Н8 the clock

' the free space of the disk of folder sDirUrl in MB; -1 — not known (the check does not stop the work then)
Function DiskFreeMb(sDirUrl As String) As Double
    Dim p As String, f As String, fso As Object, drv As Object, s As String, a As Variant, ln As Variant, i As Integer, w As Integer
    DiskFreeMb = -1
    If gSfDiskTest > 0 Then
        DiskFreeMb = gSfDiskTest
        Exit Function
    End If
    On Error GoTo EH
    p = ConvertFromURL(sDirUrl)
    If GetGuiType() = 1 Then
        fso = CreateObject("Scripting.FileSystemObject")
        drv = fso.GetDrive(fso.GetDriveName(p))
        DiskFreeMb = Int(drv.AvailableSpace / 1048576)
    Else
        If SafeDir() = "" Then Exit Function
        If Not gSfa.exists(SafeDir()) Then gSfa.createFolder(SafeDir())
        f = SafeDir() & "df_" & NewId(6) & ".txt"
        Shell("/bin/sh", 0, "-c ""df -Pk '" & Replace(p, "'", "'\''") & "' > '" & Replace(ConvertFromURL(f), "'", "'\''") & "' 2>/dev/null""", True)
        If Not gSfa.exists(f) Then Exit Function
        s = WmsAudit.ReadText(f)
        gSfa.kill(f)
        ln = Split(s, Chr(10))
        If UBound(ln) < 1 Then Exit Function
        a = Split(Trim(ln(1)), " ")
        ' «Filesystem 1024-blocks Used Available Capacity Mounted»: the fourth word (empty words of repeated spaces skipped)
        For i = 0 To UBound(a)
            If a(i) <> "" Then
                w = w + 1
                If w = 4 Then
                    If IsNumeric(a(i)) Then DiskFreeMb = Int(CDbl(a(i)) / 1024)
                    Exit For
                End If
            End If
        Next i
    End If
    Exit Function
EH:
    DiskFreeMb = -1
End Function

' the size of the book file in MB
Function BookMb() As Double
    On Error GoTo EH
    BookMb = gSfa.getSize(gDoc.getURL()) / 1048576
EH:
End Function

' Н9: "" — a copy of about nMb can be made on the disk of sDirUrl; otherwise why not (the copy is not made then)
Function DiskRefuse(sDirUrl As String, nMb As Double) As String
    Dim free As Double
    free = DiskFreeMb(sDirUrl)
    If free < 0 Then Exit Function
    If free < nMb + DISK_MIN_MB Then DiskRefuse = "на диске " & ConvertFromURL(sDirUrl) & " свободно " & Format(free, "0") & " МБ, а копии нужно " _
        & Format(nMb, "0") & " МБ и ещё " & DISK_MIN_MB & " МБ запаса — освободите место (удалите ненужные файлы, перенесите старые копии на другой носитель)"
End Function

' Н9: the warning of the free space of the disk of the book ("" — enough or not known)
Function DiskNote() As String
    Dim free As Double, need As Double
    free = DiskFreeMb(gBaseDir)
    If free < 0 Then Exit Function
    need = 3 * BookMb()
    If need < DISK_WARN_MB Then need = DISK_WARN_MB
    If free < need Then DiskNote = "ВНИМАНИЕ: на диске с книгой свободно " & Format(free, "0") & " МБ — мало для копий и журнала (нужно хотя бы " _
        & Format(need, "0") & " МБ). Освободите место; при нехватке WMS копию не делает"
End Function

' Н8: the clock of the computer behind the last operation of the journal (more than an hour) — the dates of the movements
' and of the copies come from it. "" or the warning
Function ClockNote() As String
    Dim m As String, t As Double, d As Double
    On Error GoTo EH
    m = WmsJournal.PosMark()
    If Len(m) < 19 Then Exit Function
    t = CDbl(DateSerial(Val(Left(m, 4)), Val(Mid(m, 6, 2)), Val(Mid(m, 9, 2)))) + CDbl(TimeSerial(Val(Mid(m, 12, 2)), Val(Mid(m, 15, 2)), Val(Mid(m, 18, 2))))
    d = t - CDbl(Now())
    If d > 1 / 24 Then
        ClockNote = "ВНИМАНИЕ: часы компьютера отстают" & IIf(d >= 1, " на " & Int(d) & " дн.", "") & ": сейчас " & Format(Now(), "DD.MM.YYYY HH:MM") _
            & ", а последняя операция журнала записана " & Format(t, "DD.MM.YYYY HH:MM") & ". Проверьте дату и время на компьютере — по ним ставятся " _
            & "даты движений и копий"
    End If
EH:
End Function

' ================================================================ the copies of the book (Н2, Н6)

' WmsBackup.BackupNow: a whole copy was made — its line in the list of the copies (the copies of 1.0 can be returned to)
Sub CopyIndexAdd(sName As String, kind As String, sLabel As String)
    Dim s As String, url As String, a As Variant, i As Long, out As String, nm As String
    On Error GoTo EH
    url = SafeDir() & SAFE_COPIES
    If gSfa.exists(url) Then s = Replace(WmsAudit.ReadText(url), Chr(13), "")
    ' the lines of the copies removed by the rotation are dropped
    a = Split(s, Chr(10))
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            nm = Split(a(i), ";")(0)
            If gSfa.exists(gBDir & nm) And nm <> sName Then out = out & a(i) & Chr(10)
        End If
    Next i
    out = out & sName & ";" & NumStr(SysNum(SK_SAVE_SEQ)) & ";" & kind & ";" & Replace(sLabel, ";", ",") & ";" & Format(Now(), "YYYY-MM-DD HH:MM:SS") & Chr(10)
    WriteText(url, out)
EH:
End Sub

' the copies that can be returned to, newest first: rows Array(name, operations, kind, label, when, size MB)
Function CopyList() As Variant
    Dim s As String, a As Variant, i As Long, n As Long, out() As Variant, f As Variant, j As Long, url As String
    CopyList = Array()
    On Error GoTo EH
    url = SafeDir() & SAFE_COPIES
    If Not gSfa.exists(url) Then Exit Function
    s = Replace(WmsAudit.ReadText(url), Chr(13), "")
    a = Split(s, Chr(10))
    ReDim out(UBound(a))
    For i = UBound(a) To 0 Step -1
        If a(i) <> "" Then
            f = Split(a(i), ";")
            If UBound(f) >= 4 Then
                If gSfa.exists(gBDir & f(0)) Then
                    If gSfa.getSize(gBDir & f(0)) > 0 Then
                        out(n) = Array(f(0), f(1), f(2), f(3), f(4), gSfa.getSize(gBDir & f(0)) / 1048576)
                        n = n + 1
                    End If
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve out(n - 1)
    CopyList = out
EH:
End Function

' Н6: a restore point before a risky action — a copy of the book with an understandable mark («перед загрузкой топологии»)
Function RestorePoint(sLabel As String) As String
    gBkLabel = sLabel
    RestorePoint = WmsBackup.BackupNow("point")
    gBkLabel = ""
End Function

' ================================================================ Н1 the copy to a flash drive

Sub BtnFlashCopy(Optional oEvent As Variant)
    Dim s As String, dest As String
    WmsInit()
    dest = PickFolder("Копия на флешку: выберите флешку или папку на другом диске")
    If dest = "" Then Exit Sub
    s = FlashCopy(dest)
    If Left(s, 3) = "OK:" Then
        WmsOrdersUi.UiInfo(Mid(s, 4))
    Else
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    End If
End Sub

' a folder chosen by the storekeeper (URL, "" — cancelled)
Function PickFolder(sTitle As String) As String
    Dim fp As Object
    If gSfFlashTest <> "" Then
        PickFolder = ConvertToURL(gSfFlashTest)
        Exit Function
    End If
    If gUiAuto <> 0 Then Exit Function
    fp = CreateUnoService("com.sun.star.ui.dialogs.FolderPicker")
    fp.setTitle(sTitle)
    If fp.execute() = 1 Then PickFolder = fp.getDirectory()
End Function

' the book (as last saved) and the whole journal into WMS_FLASH/<книга>_<дата>/ of folder dest: written as «….part», every
' file checked by its size, MANIFEST.txt, then the folder gets its name; FLASH_KEEP newest copies of this book are kept there.
' OK:… / ERR:…
Function FlashCopy(ByVal dest As String) As String
    Dim sBase As String, stamp As String, root As String, part As String, fin As String, src() As String, rel() As String, n As Long, i As Long
    Dim jf As Variant, total As Double, nSrc As Double, nDst As Double, man As String, why As String, t0 As Long, sz() As Double
    On Error GoTo EH
    t0 = GetSystemTicks()
    If Not gDoc.hasLocation() Then
        FlashCopy = "ERR:книга не сохранена в файл"
        Exit Function
    End If
    If Right(dest, 1) <> "/" Then dest = dest & "/"
    If Left(LCase(dest), Len(gBaseDir)) = LCase(gBaseDir) Then
        FlashCopy = "ERR:копию нельзя класть в папку WMS: при поломке диска пропадут обе. Выберите флешку или папку на другом диске"
        Exit Function
    End If
    If Not gSfa.exists(dest) Then
        FlashCopy = "ERR:папка " & ConvertFromURL(dest) & " не найдена — флешка вставлена?"
        Exit Function
    End If
    jf = WmsJournal.ListFiles()
    ReDim src(UBound(jf) + 1)
    ReDim rel(UBound(jf) + 1)
    ReDim sz(UBound(jf) + 1)
    src(0) = gDoc.getURL()
    rel(0) = FileNameOf(gDoc.getURL())
    For i = 0 To UBound(jf)
        src(i + 1) = gJDir & jf(i)
        rel(i + 1) = JOURNAL_DIR & "/" & jf(i)
    Next i
    n = UBound(jf) + 2
    For i = 0 To n - 1
        sz(i) = gSfa.getSize(src(i))
        total = total + sz(i)
    Next i
    why = DiskRefuse(dest, total / 1048576)
    If why <> "" Then
        FlashCopy = "ERR:копия не сделана: " & why
        Exit Function
    End If
    sBase = WmsBackup.BookBaseName()
    root = dest & FLASH_DIR & "/"
    If Not gSfa.exists(root) Then gSfa.createFolder(root)
    stamp = Format(Now(), "YYYY-MM-DD") & "_" & Format(Now(), "HHMMSS")
    If gSfa.exists(root & sBase & "_" & stamp & "/") Then
        ' two copies in one second: the second one a second later
        Wait 1100
        stamp = Format(Now(), "YYYY-MM-DD") & "_" & Format(Now(), "HHMMSS")
    End If
    fin = root & sBase & "_" & stamp & "/"
    part = root & sBase & "_" & stamp & ".part/"
    If gSfa.exists(part) Then gSfa.kill(part)
    gSfa.createFolder(part)
    gSfa.createFolder(part & JOURNAL_DIR)
    For i = 0 To n - 1
        gSfa.copy(src(i), part & rel(i))
        nDst = gSfa.getSize(part & rel(i))
        If nDst <> sz(i) Then
            gSfa.kill(part)
            FlashCopy = "ERR:копия не сделана: файл " & rel(i) & " записан не полностью (" & nDst & " из " & sz(i) & " байт) — носитель полон или " _
                & "неисправен; прежние копии на нём не тронуты"
            Exit Function
        End If
        man = man & rel(i) & ";" & NumStr(sz(i)) & Chr(10)
    Next i
    WriteText(part & "MANIFEST.txt", "WMS — копия на носитель " & Format(Now(), "DD.MM.YYYY HH:MM") & ", книга " & rel(0) & ", операций в журнале по " _
        & NumStr(SysNum(SK_LAST_SEQ)) & Chr(10) & man)
    gSfa.move(part, fin)
    FlashRotate(root, sBase)
    StSet("FLASH", Format(Now(), "YYYY-MM-DD HH:MM") & "|" & ConvertFromURL(fin))
    FlashCopy = "OK:Копия на носитель сделана и проверена: " & ConvertFromURL(fin) & Chr(10) & "Книга (на момент последнего сохранения) и журнал — " _
        & n & " файл(ов), " & Format(total / 1048576, "0.0") & " МБ, " & Format((GetSystemTicks() - t0) / 1000, "0.0") & " с." _
        & IIf(gDoc.isModified(), Chr(10) & "Операции после последнего сохранения книги — в журнале, он в копии полный.", "")
    gUiLastMsg = FlashCopy
    Exit Function
EH:
    FlashCopy = "ERR:копия не сделана: " & Error$
    gUiLastMsg = FlashCopy
    On Error Resume Next
    If part <> "" Then
        If gSfa.exists(part) Then gSfa.kill(part)
    End If
End Function

' the FLASH_KEEP newest copies of this book in folder root are kept (the folders sort by their date)
Private Sub FlashRotate(root As String, sBase As String)
    Dim f As Variant, names() As String, n As Long, i As Long, j As Long, t As String, nm As String
    On Error Resume Next
    f = gSfa.getFolderContents(root, True)
    n = -1
    For i = 0 To UBound(f)
        nm = FileNameOf(f(i))
        If Right(nm, 1) = "/" Then nm = Left(nm, Len(nm) - 1)
        If gSfa.isFolder(f(i)) And Left(nm, Len(sBase) + 1) = sBase & "_" And Right(nm, 5) <> ".part" And Len(nm) = Len(sBase) + 18 Then
            n = n + 1
            ReDim Preserve names(n)
            names(n) = nm
        End If
    Next i
    If n + 1 <= FLASH_KEEP Then Exit Sub
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
    For i = 0 To n - FLASH_KEEP
        gSfa.kill(root & names(i) & "/")
    Next i
End Sub

' Н1: the reminder — no copy on a medium for FLASH_REMIND_DAYS days ("" — not due)
Function FlashNote() As String
    Dim s As String, d As Double, seen As String
    On Error GoTo EH
    If SysNum(SK_LAST_SEQ) <= 0 Then Exit Function
    s = StGet("FLASH")
    If s = "" Then
        ' never: counted from the day WMS first saw no copy
        seen = StGet("FLASH_NONE_SINCE")
        If seen = "" Then
            StSet("FLASH_NONE_SINCE", Format(Now(), "YYYY-MM-DD"))
            Exit Function
        End If
        d = CDbl(Now()) - CDbl(DateSerial(Val(Left(seen, 4)), Val(Mid(seen, 6, 2)), Val(Mid(seen, 9, 2))))
        If d >= FLASH_REMIND_DAYS Then FlashNote = "Копии на флешке (другом носителе) ещё не было — нажмите «Копия на флешку…» на «Главной»: при поломке " _
            & "компьютера книга и журнал останутся"
        Exit Function
    End If
    d = CDbl(Now()) - CDbl(DateSerial(Val(Left(s, 4)), Val(Mid(s, 6, 2)), Val(Mid(s, 9, 2))))
    If d >= FLASH_REMIND_DAYS Then FlashNote = "Копии на флешке нет " & Int(d) & " дн. (последняя " & Mid(s, 9, 2) & "." & Mid(s, 6, 2) & "." & Left(s, 4) _
        & ") — нажмите «Копия на флешку…» на «Главной»"
EH:
End Function

' ================================================================ at the start

' the notes of the start of the WMS: the clock, the disk, the copy on a medium, the mirror of the journal, the check of the day
Function StartupNotes() As String
    Dim s As String, x As String
    On Error GoTo EH
    ' Н2: the book a return to a copy just put in place — the journal and the check run after the start
    If AfterRestoreMine() Then s = "возврат к копии: книга будет доведена по журналу и проверена"
    ' П4, П12: the role of the copy said first
    If WmsHelp.BookRole() = WmsHelp.ROLE_TRAINING Then s = "УЧЕБНАЯ WMS: здесь можно пробовать всё, рабочая книга не меняется" & IIf(s <> "", " | " & s, "")
    If WmsHelp.BookRole() = WmsHelp.ROLE_ANON Then s = "ОБЕЗЛИЧЕННАЯ копия WMS (для разработчика)" & IIf(s <> "", " | " & s, "")
    x = ClockNote()
    If x <> "" Then s = s & IIf(s <> "", " | ", "") & x
    x = DiskNote()
    If x <> "" Then s = s & IIf(s <> "", " | ", "") & x
    If gState = "CLEAN" Or gBlock = "TAIL" Then
        ' Н3: the mirror followed (the files it lacks copied)
        x = MirrorStart()
        If x <> "" Then s = s & IIf(s <> "", " | ", "") & x
    End If
    ' 1.0.1: the grid over the coloured cells (a setting of LibreOffice, see WmsCare.GridOnColored)
    WmsCare.GridOnColored()
    If gState = "CLEAN" Then
        x = FlashNote()
        If x <> "" Then s = s & IIf(s <> "", " | ", "") & x
        ' Н5: the check of the day (at its first start); С3, С4: the growth of the book and its advice
        ' (П2 «Что нового» is shown by WmsCore.WmsStartup after the pages: a window opened here would wait over the old panel)
        If Left(StGet("DAILY"), 10) = Format(Now(), "YYYY-MM-DD") Then
            ' Т8: LibreOffice updated since the check of the day — its formulas checked now
            x = CalcAtStart()
            If x <> "" Then s = s & IIf(s <> "", " | ", "") & x
        Else
            DailyCheck(False)
            s = s & IIf(s <> "", " | ", "") & DailyLine()
            WmsCare.GrowthMark()
            x = WmsCare.GrowthNote()
            If x <> "" Then s = s & " | " & x
        End If
    End If
    StartupNotes = s
    Exit Function
EH:
    StartupNotes = s
End Function

' ================================================================ Н2 «Вернуться к копии»
' The working book cannot replace its own open file. It writes the order (WMS_State/wms_restore_order.txt), gives up the lock
' and opens the copy with the macros: the copy sees the order at its start and, after the macro of the working book ended
' (an asynchronous call), closes the working book, puts its file aside into WMS_Backups («…_отложена.ods»), copies itself in
' its place (each step checked by the size of the file; a failed step puts the working book back) and opens it. The
' restored book sees the order after the return: it closes the copy, replays the journal («Восстановить») and checks itself.

Sub BtnReturnToCopy(Optional oEvent As Variant)
    Dim lst As Variant, opts() As String, i As Long, k As Integer, s As String
    WmsInit()
    ' a copy opened right from WMS_Backups: it can be made the working book itself
    If IsCopyBook() Then
        s = CopyToWork(True)
        If Left(s, 3) <> "OK:" Then WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
        Exit Sub
    End If
    lst = CopyList()
    If UBound(lst) < 0 Then
        WmsUi.UiMessage("Копий, к которым можно вернуться, пока нет. Их делает WMS 1.0: ежедневную — при первом запуске дня, «Резервная копия» — по " _
            & "кнопке, точки восстановления — перед загрузкой топологии, «Отложить хвост» и коррекцией переноса.")
        Exit Sub
    End If
    ReDim opts(UBound(lst))
    For i = 0 To UBound(lst)
        opts(i) = CopyLabel(lst(i))
    Next i
    k = WmsPrimeUi.UiChoose("Вернуться к копии", "Выберите копию книги (новые сверху). Операции после неё WMS доведёт по журналу сама.", opts)
    If k < 0 Then Exit Sub
    s = ReturnToCopy(CStr(lst(k)(0)), True)
    If Left(s, 3) <> "OK:" And Left(s, 5) <> "SKIP:" Then WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
End Sub

' «20.10.2026 09:15 · ежедневная · операций в копии 1234 · 12,3 МБ»
Function CopyLabel(x As Variant) As String
    Dim t As String
    t = CStr(x(4))
    CopyLabel = Mid(t, 9, 2) & "." & Mid(t, 6, 2) & "." & Left(t, 4) & " " & Mid(t, 12, 5) & " · " & KindText(CStr(x(2)), CStr(x(3))) _
        & " · операций в копии " & x(1) & " · " & Format(x(5), "0.0") & " МБ"
End Function

Private Function KindText(kind As String, sLabel As String) As String
    Select Case kind
    Case "daily": KindText = "ежедневная"
    Case "manual": KindText = "ручная"
    Case "preupgrade": KindText = "перед обновлением"
    Case "premigration": KindText = "перед переносом"
    Case "point": KindText = "точка восстановления" & IIf(sLabel <> "", " " & sLabel, "")
    Case "good": KindText = "последняя без ошибок"
    Case Else: KindText = kind
    End Select
End Function

' the book is a copy opened from WMS_Backups
Function IsCopyBook() As Boolean
    IsCopyBook = (InStr(gDoc.getURL(), "/" & BACKUP_DIR & "/") > 0)
End Function

' the return of the working book to copy sName of WMS_Backups (bAsk — the question with what will happen). OK:… / SKIP / ERR
Function ReturnToCopy(sName As String, bAsk As Boolean) As String
    Dim copyUrl As String, k As Integer, why As String, x As Variant, lst As Variant, i As Long, sInfo As String
    WmsInit()
    On Error GoTo EH
    If gBDir = "" Or Not gDoc.hasLocation() Then
        ReturnToCopy = "ERR:книга не сохранена в файл"
        Exit Function
    End If
    copyUrl = gBDir & sName
    If Not gSfa.exists(copyUrl) Then
        ReturnToCopy = "ERR:копии «" & sName & "» нет в папке " & ConvertFromURL(gBDir)
        Exit Function
    End If
    why = DiskRefuse(gBaseDir, BookMb() + gSfa.getSize(copyUrl) / 1048576)
    If why <> "" Then
        ReturnToCopy = "ERR:возврат невозможен: " & why
        Exit Function
    End If
    lst = CopyList()
    For i = 0 To UBound(lst)
        If lst(i)(0) = sName Then sInfo = CopyLabel(lst(i))
    Next i
    If bAsk Then
        k = WmsPrimeUi.UiButtons("Вернуться к копии", "Книга вернётся к копии: " & IIf(sInfo <> "", sInfo, sName) & "." & Chr(10) & Chr(10) _
            & "1. Эта книга закроется и будет отложена в папку «" & BACKUP_DIR & "» (её можно будет открыть и посмотреть)." & Chr(10) _
            & "2. Копия станет рабочей книгой; WMS сама доведёт её по журналу — операции после копии вернутся — и проверит." & Chr(10) & Chr(10) _
            & "Журнал не меняется. Не переносятся только непроведённые строки и правки после последнего сохранения книги.", _
            Array("Вернуться к копии", "Отмена"), Array(1, 0))
        If k <> 1 Then
            ReturnToCopy = "SKIP:отменено"
            Exit Function
        End If
    End If
    ReturnToCopy = OrderAndOpen(copyUrl, gDoc.getURL())
    Exit Function
EH:
    ReturnToCopy = "ERR:возврат к копии не начат: " & Error$
    gUiLastMsg = ReturnToCopy
End Function

' the order written, the lock given up (this book does not post any more), the copy opened with the macros
Private Function OrderAndOpen(copyUrl As String, target As String) As String
    Dim args(0) As New com.sun.star.beans.PropertyValue, doc As Object
    If Not WriteText(FolderOf(target) & SAFE_DIR & "/" & RESTORE_ORDER, "COPY=" & copyUrl & Chr(10) & "TARGET=" & target & Chr(10) & "TIME=" _
        & Format(Now(), "YYYY-MM-DD HH:MM:SS") & Chr(10) & "AUTO=" & gUiAuto & Chr(10)) Then
        OrderAndOpen = "ERR:не удалось записать приказ возврата в " & ConvertFromURL(FolderOf(target) & SAFE_DIR)
        Exit Function
    End If
    SetBlocked("RESTORING", "идёт возврат к копии — эта книга будет закрыта")
    WmsLock.ReleaseIfOwner()
    args(0).Name = "MacroExecutionMode"
    args(0).Value = 4
    doc = StarDesktop.loadComponentFromURL(copyUrl, "_blank", 0, args())
    OrderAndOpen = "OK:возврат к копии начат"
    gUiLastMsg = OrderAndOpen
End Function

' a copy opened right from WMS_Backups made the working book (the working one put aside), bAsk — the question
Function CopyToWork(bAsk As Boolean) As String
    Dim target As String, k As Integer
    WmsInit()
    target = SysStr(SK_REG_URL)
    If target = "" Then
        CopyToWork = "ERR:в копии не записан рабочий файл WMS — скопируйте её на место рабочей книги вручную (руководство, «Восстановление из копии»)"
        Exit Function
    End If
    If bAsk Then
        k = WmsPrimeUi.UiButtons("Вернуться к копии", "Сделать эту копию рабочей книгой (" & ConvertFromURL(target) & ")?" & Chr(10) & "Рабочая книга будет отложена " _
            & "в «" & BACKUP_DIR & "» (если она открыта — закроется), копия встанет на её место, WMS доведёт её по журналу и проверит.", _
            Array("Сделать рабочей", "Отмена"), Array(1, 0))
        If k <> 1 Then
            CopyToWork = "SKIP:отменено"
            Exit Function
        End If
    End If
    If Not WriteText(FolderOf(target) & SAFE_DIR & "/" & RESTORE_ORDER, "COPY=" & gDoc.getURL() & Chr(10) & "TARGET=" & target & Chr(10) & "TIME=" _
        & Format(Now(), "YYYY-MM-DD HH:MM:SS") & Chr(10) & "AUTO=" & gUiAuto & Chr(10)) Then
        CopyToWork = "ERR:не удалось записать приказ возврата"
        Exit Function
    End If
    Later("RESTORE")
    CopyToWork = "OK:возврат начат"
End Function

' the value of key k of an order text
Private Function OrderVal(s As String, k As String) As String
    Dim a As Variant, i As Long
    a = Split(Replace(s, Chr(13), ""), Chr(10))
    For i = 0 To UBound(a)
        If Left(a(i), Len(k) + 1) = k & "=" Then OrderVal = Mid(a(i), Len(k) + 2)
    Next i
End Function

' an order is fresh for 10 minutes (a copy opened days later never acts on an old one: it is removed)
Private Function OrderFresh(s As String) As Boolean
    Dim t As String, d As Double
    t = OrderVal(s, "TIME")
    If Len(t) < 19 Then Exit Function
    d = CDbl(DateSerial(Val(Left(t, 4)), Val(Mid(t, 6, 2)), Val(Mid(t, 9, 2)))) + CDbl(TimeSerial(Val(Mid(t, 12, 2)), Val(Mid(t, 15, 2)), Val(Mid(t, 18, 2))))
    OrderFresh = (Abs(CDbl(Now()) - d) < 10 / 1440)
End Function

' WmsRecovery.Analyze: this book is the copy a fresh order returns to — the return goes on after the start (True)
Function RestoreOrderMine() As Boolean
    Dim url As String, s As String, parent As String
    On Error GoTo EH
    If Not IsCopyBook() Then Exit Function
    parent = Left(gBaseDir, Len(gBaseDir) - Len(BACKUP_DIR) - 1)
    url = parent & SAFE_DIR & "/" & RESTORE_ORDER
    If Not gSfa.exists(url) Then Exit Function
    s = WmsAudit.ReadText(url)
    If OrderVal(s, "COPY") <> gDoc.getURL() Then Exit Function
    If Not OrderFresh(s) Then
        gSfa.kill(url)
        Exit Function
    End If
    Later("RESTORE")
    RestoreOrderMine = True
EH:
End Function

' the next step runs when the current macro (and the load of this book) ended
Private Sub Later(kind As String)
    gSfCbKind = kind
    gSfCb = CreateUnoListener("SfCb_", "com.sun.star.awt.XCallback")
    CreateUnoService("com.sun.star.awt.AsyncCallback").addCallback(gSfCb, Array())
End Sub

Sub SfCb_notify(aData As Variant)
    Dim k As String
    k = gSfCbKind
    gSfCbKind = ""
    If k = "RESTORE" Then
        RestoreRun()
    ElseIf k = "AFTER" Then
        AfterRestoreRun()
    End If
End Sub

' the copy carries the order out
Sub RestoreRun()
    Dim orderUrl As String, s As String, target As String, copyUrl As String, aside As String, nT As Double, nC As Double, msg As String
    Dim args(0) As New com.sun.star.beans.PropertyValue, en As Object, comp As Object, stamp As String, bk As String
    WmsInit()
    On Error GoTo EH
    copyUrl = gDoc.getURL()
    bk = FolderOf(copyUrl)
    orderUrl = Left(bk, Len(bk) - Len(BACKUP_DIR) - 1) & SAFE_DIR & "/" & RESTORE_ORDER
    If Not gSfa.exists(orderUrl) Then Exit Sub
    s = WmsAudit.ReadText(orderUrl)
    gSfa.kill(orderUrl)
    ' the windows as in the working book (a test answers without them)
    If OrderVal(s, "AUTO") <> "" Then gUiAuto = Val(OrderVal(s, "AUTO"))
    target = OrderVal(s, "TARGET")
    If target = "" Or OrderVal(s, "COPY") <> copyUrl Then Exit Sub
    ' 1. the working book closed if it is open here (its operations are in the journal)
    en = StarDesktop.getComponents().createEnumeration()
    Do While en.hasMoreElements()
        comp = en.nextElement()
        If HasUnoInterfaces(comp, "com.sun.star.frame.XModel") Then
            If comp.getURL() = target Then
                comp.setModified(False)
                comp.close(True)
            End If
        End If
    Loop
    ' 2. the working book put aside — a copy first: its file stays until the copy is in place
    stamp = Format(Now(), "YYYY-MM-DD") & "_" & Format(Now(), "HHMMSS")
    If gSfa.exists(target) Then
        nT = gSfa.getSize(target)
        aside = bk & BaseName(target) & "_" & stamp & "_отложена.ods"
        gSfa.copy(target, aside)
        If gSfa.getSize(aside) <> nT Then
            msg = "рабочая книга не отложена (копия неполная) — ничего не изменено"
            GoTo FAIL
        End If
        gSfa.kill(target)
    End If
    ' 3. the copy in its place
    nC = gSfa.getSize(copyUrl)
    gSfa.copy(copyUrl, target)
    If gSfa.getSize(target) <> nC Then
        msg = "копия не встала на место рабочей книги (записана не полностью)"
        If aside <> "" Then
            gSfa.kill(target)
            gSfa.copy(aside, target)
            msg = msg & "; рабочая книга возвращена на место"
        End If
        GoTo FAIL
    End If
    ' 4. the order after the return, 5. the restored book opened
    WriteText(FolderOf(target) & SAFE_DIR & "/" & AFTER_RESTORE, "COPY=" & copyUrl & Chr(10) & "ASIDE=" & aside & Chr(10) & "TIME=" _
        & Format(Now(), "YYYY-MM-DD HH:MM:SS") & Chr(10) & "AUTO=" & gUiAuto & Chr(10))
    args(0).Name = "MacroExecutionMode"
    args(0).Value = 4
    StarDesktop.loadComponentFromURL(target, "_blank", 0, args())
    Exit Sub
EH:
    msg = Error$
FAIL:
    gUiLastMsg = "ERR:возврат к копии не выполнен: " & msg
    WmsUi.UiMessage("Возврат к копии не выполнен: " & msg & Chr(10) & "Рабочая книга: " & ConvertFromURL(target) & IIf(aside <> "", Chr(10) _
        & "Отложенная: " & ConvertFromURL(aside), ""))
End Sub

' the file name without «.ods»
Private Function BaseName(url As String) As String
    Dim f As String
    f = FileNameOf(url)
    If LCase(Right(f, 4)) = ".ods" Then f = Left(f, Len(f) - 4)
    BaseName = f
End Function

' StartupNotes of the restored book: the order after the return found — the rest runs after the start
Private Function AfterRestoreMine() As Boolean
    Dim url As String, s As String
    On Error GoTo EH
    url = SafeDir() & AFTER_RESTORE
    If Not gSfa.exists(url) Then Exit Function
    s = WmsAudit.ReadText(url)
    If Not OrderFresh(s) Then
        gSfa.kill(url)
        Exit Function
    End If
    Later("AFTER")
    AfterRestoreMine = True
EH:
End Function

' the restored book: the copy closed, the journal replayed, the self-check — and the result said
Sub AfterRestoreRun()
    Dim url As String, s As String, copyUrl As String, aside As String, en As Object, comp As Object, r1 As String, r2 As String, txt As String
    WmsInit()
    On Error GoTo EH
    url = SafeDir() & AFTER_RESTORE
    If Not gSfa.exists(url) Then Exit Sub
    s = WmsAudit.ReadText(url)
    gSfa.kill(url)
    If OrderVal(s, "AUTO") <> "" Then gUiAuto = Val(OrderVal(s, "AUTO"))
    copyUrl = OrderVal(s, "COPY")
    aside = OrderVal(s, "ASIDE")
    en = StarDesktop.getComponents().createEnumeration()
    Do While en.hasMoreElements()
        comp = en.nextElement()
        If HasUnoInterfaces(comp, "com.sun.star.frame.XModel") Then
            If comp.getURL() = copyUrl Then comp.close(True)
        End If
    Loop
    If gBlock = "TAIL" Then r1 = WmsCore.ActionRecover()
    r2 = WmsCore.ActionSelfCheck()
    StSet("RESTORED", Format(Now(), "YYYY-MM-DD HH:MM") & "|" & FileNameOf(copyUrl) & "|" & ConvertFromURL(aside))
    txt = "Книга возвращена к копии «" & FileNameOf(copyUrl) & "»." & Chr(10) & IIf(r1 <> "", "По журналу: " & Left(r1, InStr(r1 & " ||", " ||") - 1) & Chr(10), "") _
        & "Состояние: " & IIf(gState = "CLEAN", "работа разрешена", "ЗАБЛОКИРОВАНО — " & gBlockText) & Chr(10) & Left(r2, 300) & Chr(10) _
        & IIf(aside <> "", "Прежняя книга отложена: " & ConvertFromURL(aside) & Chr(10), "") & "Сохраните книгу (Ctrl+S)."
    gUiLastMsg = IIf(gState = "CLEAN", "OK:", "ERR:") & txt
    WmsUi.UiMessage(txt)
    Exit Sub
EH:
    gUiLastMsg = "ERR:после возврата к копии: " & Error$
    WmsUi.UiMessage("После возврата к копии: " & Error$ & ". Нажмите «Восстановить», затем «Самопроверка».")
End Sub

' test seam: the time (ms) of n SHA-256 checksums (NSS of LibreOffice) of a line of k characters; -1 — not available
Function TestShaMs(n As Long, k As Long) As String
    Dim s As String, i As Long, t0 As Long, h As String
    s = String(k, "Ж")
    t0 = GetSystemTicks()
    For i = 1 To n
        h = WmsJournal.LineSha(s)
    Next i
    TestShaMs = (GetSystemTicks() - t0) & "|" & h
End Function

' ================================================================ Н3 the mirror of the journal

' the folder of the mirror of this book inside folder sDir (URL)
Function MirrorFolder(ByVal sDir As String) As String
    If Right(sDir, 1) <> "/" Then sDir = sDir & "/"
    MirrorFolder = sDir & "WMS_Зеркало_журнала_" & WmsBackup.BookBaseName() & "/"
End Function

Sub BtnMirror(Optional oEvent As Variant)
    Dim cur As String, k As Integer, d As String, s As String
    WmsInit()
    cur = StGet("MIRROR")
    k = WmsPrimeUi.UiButtons("Зеркало журнала", IIf(cur <> "", "Сейчас зеркало журнала: " & ConvertFromURL(MirrorFolder(cur)) & "." & Chr(10) _
        & IIf(gJMirrorErr <> "", "Последняя неполадка: " & gJMirrorErr & Chr(10), "Зеркало в порядке." & Chr(10)), "Зеркала журнала нет." & Chr(10)) & Chr(10) _
        & "Зеркало — вторая копия журнала: каждая проведённая операция пишется и туда. Выберите другой диск или флешку, которая всегда " _
        & "вставлена. Испорченный журнал WMS восстановит из зеркала («Журнал из зеркала»).", _
        IIf(cur <> "", Array("Выбрать другую папку…", "Выключить зеркало", "Закрыть"), Array("Выбрать папку…", "Закрыть")), _
        IIf(cur <> "", Array(1, 2, 0), Array(1, 0)))
    If k = 1 Then
        d = PickFolder("Папка зеркала журнала (другой диск или флешка)")
        If d = "" Then Exit Sub
        s = MirrorSet(d)
    ElseIf k = 2 Then
        s = MirrorSet("")
    Else
        Exit Sub
    End If
    If Left(s, 3) = "OK:" Then WmsOrdersUi.UiInfo(Mid(s, 4)) Else WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
End Sub

' the mirror turned on into folder sDir (URL; "" — off): every journal file copied there and checked. OK:… / ERR:…
Function MirrorSet(sDir As String) As String
    Dim m As String, n As Long
    WmsInit()
    On Error GoTo EH
    If sDir = "" Then
        StSet("MIRROR", "")
        gJMirror = ""
        gJMirrorErr = ""
        MirrorSet = "OK:Зеркало журнала выключено. Файлы зеркала не удалены."
        Exit Function
    End If
    If Right(sDir, 1) <> "/" Then sDir = sDir & "/"
    If Left(LCase(sDir), Len(gBaseDir)) = LCase(gBaseDir) Then
        MirrorSet = "ERR:зеркало в папке WMS бесполезно — при поломке диска пропадёт вместе с журналом. Выберите другой диск или флешку"
        Exit Function
    End If
    If Not gSfa.exists(sDir) Then
        MirrorSet = "ERR:папка " & ConvertFromURL(sDir) & " не найдена"
        Exit Function
    End If
    m = MirrorFolder(sDir)
    gJMirror = m
    n = MirrorSync()
    If gJMirrorErr <> "" Then
        gJMirror = ""
        MirrorSet = "ERR:зеркало не включено: " & gJMirrorErr
        Exit Function
    End If
    StSet("MIRROR", sDir)
    MirrorSet = "OK:Зеркало журнала включено: " & ConvertFromURL(m) & Chr(10) & "Файлов журнала: " & n & ", все проверены по размеру. Дальше каждая " _
        & "операция пишется и туда."
    gUiLastMsg = MirrorSet
    Exit Function
EH:
    gJMirror = ""
    MirrorSet = "ERR:зеркало не включено: " & Error$
    gUiLastMsg = MirrorSet
End Function

' every journal file into the mirror where it is missing or shorter (a longer one is never overwritten here: it may be the
' whole one — «Журнал из зеркала» decides). The number of the journal files; trouble — gJMirrorErr
Function MirrorSync() As Long
    Dim files As Variant, i As Long, src As String, dst As String, nS As Double, nD As Double, longer As String
    On Error GoTo EH
    If gJMirror = "" Then Exit Function
    If Not gSfa.exists(gJMirror) Then gSfa.createFolder(gJMirror)
    files = WmsJournal.ListFiles()
    For i = 0 To UBound(files)
        src = gJDir & files(i)
        dst = gJMirror & files(i)
        nS = gSfa.getSize(src)
        nD = -1
        If gSfa.exists(dst) Then nD = gSfa.getSize(dst)
        If nD > nS Then
            longer = longer & IIf(longer <> "", ", ", "") & files(i)
        ElseIf nD < nS Then
            If nD >= 0 Then gSfa.kill(dst)
            gSfa.copy(src, dst)
            If gSfa.getSize(dst) <> nS Then
                gJMirrorErr = Format(Now(), "DD.MM.YYYY HH:MM") & " — «" & files(i) & "» в зеркало записан не полностью"
                MirrorSync = i
                Exit Function
            End If
        End If
    Next i
    MirrorSync = UBound(files) + 1
    If longer <> "" Then
        gJMirrorErr = "в зеркале больше записей, чем в журнале (" & longer & ") — журнал мог потерять конец: «Журнал из зеркала»"
    Else
        gJMirrorErr = ""
    End If
    Exit Function
EH:
    gJMirrorErr = Format(Now(), "DD.MM.YYYY HH:MM") & " — " & Error$
End Function

' the start: the mirror of the state file followed; "" or its trouble
Function MirrorStart() As String
    Dim m As String
    m = StGet("MIRROR")
    gJMirror = ""
    If m = "" Then Exit Function
    gJMirror = MirrorFolder(m)
    MirrorSync()
    If gJMirrorErr <> "" Then MirrorStart = "ВНИМАНИЕ: зеркало журнала (" & ConvertFromURL(gJMirror) & "): " & gJMirrorErr
End Function

Sub BtnMirrorRestore(Optional oEvent As Variant)
    Dim s As String, k As Integer
    WmsInit()
    If StGet("MIRROR") = "" Then
        WmsUi.UiMessage("Зеркало журнала не включено — восстанавливать не из чего. Включите его кнопкой «Зеркало журнала…».")
        Exit Sub
    End If
    k = WmsPrimeUi.UiButtons("Журнал из зеркала", "WMS сравнит каждый файл журнала с зеркалом. Файл, в котором не хватает записей или есть испорченные строки, " _
        & "будет заменён целым из зеркала; прежний файл — отложен в «" & BACKUP_DIR & "». Затем WMS заново проверит книгу и журнал.", _
        Array("Сравнить и восстановить", "Отмена"), Array(1, 0))
    If k <> 1 Then Exit Sub
    s = MirrorRestore()
    WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
End Sub

' each journal file compared with the mirror: missing, shorter (the mirror goes on after it) or with damaged lines while
' the mirror has fewer — replaced by the mirror (the old file put aside into WMS_Backups); then the start runs again
Function MirrorRestore() As String
    Dim mf As String, files As Variant, i As Long, fn As String, j As String, mm As String, jt As String, mt As String, n As Long, nSkip As Long
    Dim aside As String, stamp As String, sDone As String, sConf As String, dj As Long, dm As Long, res As String
    WmsInit()
    On Error GoTo EH
    If StGet("MIRROR") = "" Then
        MirrorRestore = "ERR:зеркало журнала не включено"
        Exit Function
    End If
    mf = MirrorFolder(StGet("MIRROR"))
    If Not gSfa.exists(mf) Then
        MirrorRestore = "ERR:папка зеркала " & ConvertFromURL(mf) & " недоступна — флешка вставлена?"
        Exit Function
    End If
    stamp = Format(Now(), "YYYY-MM-DD") & "_" & Format(Now(), "HHMMSS")
    files = gSfa.getFolderContents(mf, False)
    For i = 0 To UBound(files)
        fn = FileNameOf(files(i))
        If Left(fn, Len(JOURNAL_PREFIX)) = JOURNAL_PREFIX And Right(fn, Len(JOURNAL_EXT)) = JOURNAL_EXT Then
            j = gJDir & fn
            mm = mf & fn
            If Not gSfa.exists(j) Then
                gSfa.copy(mm, j)
                n = n + 1
                sDone = sDone & Chr(10) & "• " & fn & ": не было — взят из зеркала"
            Else
                jt = WmsAudit.ReadText(j)
                mt = WmsAudit.ReadText(mm)
                If jt = mt Or Left(jt, Len(mt)) = mt Then
                    nSkip = nSkip + 1
                Else
                    dj = DamagedLines(jt)
                    dm = DamagedLines(mt)
                    If Left(mt, Len(jt)) = jt Or dm < dj Then
                        aside = gBDir & fn & ".отложен_" & stamp
                        If Not gSfa.exists(gBDir) Then gSfa.createFolder(gBDir)
                        gSfa.copy(j, aside)
                        gSfa.kill(j)
                        gSfa.copy(mm, j)
                        n = n + 1
                        sDone = sDone & Chr(10) & "• " & fn & ": " & IIf(Left(mt, Len(jt)) = jt, "в журнале не хватало записей", "испорченных строк " & dj _
                            & " (в зеркале " & dm & ")") & " — взят из зеркала, прежний отложен"
                    Else
                        sConf = sConf & Chr(10) & "• " & fn & ": журнал и зеркало различаются, а испорченных строк в зеркале не меньше — не тронут"
                    End If
                End If
            End If
        End If
    Next i
    gJPrevKnown = False
    res = WmsCore.WmsStartup()
    MirrorRestore = IIf(sConf = "", "OK:", "ERR:") & "Журнал из зеркала: заменено файлов " & n & ", совпадают " & nSkip & "." & sDone & sConf & Chr(10) & Chr(10) _
        & "Состояние WMS: " & IIf(gState = "CLEAN", "работа разрешена", "ЗАБЛОКИРОВАНО — " & gBlockText)
    gUiLastMsg = MirrorRestore
    Exit Function
EH:
    MirrorRestore = "ERR:журнал из зеркала не восстановлен: " & Error$
    gUiLastMsg = MirrorRestore
End Function

' the lines of a journal text that are not whole entries
Function DamagedLines(t As String) As Long
    Dim a As Variant, i As Long, seq As Long, inst As String, sType As String, n As Long
    a = Split(Replace(t, Chr(13), ""), Chr(10))
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            If Not WmsJournal.ParseHeader(CStr(a(i)), seq, inst, sType) Then n = n + 1
        End If
    Next i
    DamagedLines = n
End Function

' ================================================================ Н4 the chain of the journal

' the lines written since the last check checked by the chain: the field CH of each equals the checksum of the line before.
' "" or the trouble (the entries named); nLines — the lines checked. Where the check ended is kept (CHAIN=file|offset|hash)
Function ChainCheck(ByRef nLines As Long) As String
    Dim st As String, a As Variant, files As Variant, idx As Long, i As Long, k As Long, off As Double, sz As Double, prev As String
    Dim txt As String, ln As Variant, s As String, p As Long, q As Long, ch As String, seq As Long, inst As String, sType As String
    Dim lastFile As String, lastOff As Double, nBad As Long, h As String, url As String
    nLines = 0
    On Error GoTo EH
    files = WmsJournal.ListFiles()
    If UBound(files) < 0 Then Exit Function
    If WmsJournal.LineHash("x") = "" Then
        ChainCheck = "контрольная цепочка на этом ПК не проверяется (нет модуля NSS LibreOffice)"
        Exit Function
    End If
    st = StGet("CHAIN")
    a = Split(st, "|")
    idx = -1
    If UBound(a) >= 2 Then
        For i = 0 To UBound(files)
            If files(i) = a(0) Then idx = i
        Next i
    End If
    If idx < 0 Then
        ' the first check (or its file is gone): the chain is followed from the end of the journal now on
        lastFile = files(UBound(files))
        url = gJDir & lastFile
        StSet("CHAIN", lastFile & "|" & NumStr(gSfa.getSize(url)) & "|" & WmsJournal.LineHash(WmsJournal.LastLineOf(url)))
        Exit Function
    End If
    prev = a(2)
    For i = idx To UBound(files)
        url = gJDir & files(i)
        sz = gSfa.getSize(url)
        off = 0
        If i = idx Then off = Val(a(1))
        If off > sz Then
            s = s & IIf(s <> "", "; ", "") & "файл «" & files(i) & "» стал короче, чем при прошлой проверке (" & NumStr(off) & " → " & NumStr(sz) & " байт)"
            off = 0
            prev = ""
        End If
        lastFile = files(i)
        lastOff = off
        If off < sz Then
            txt = WmsJournal.ReadFrom(url, off)
            ' only whole lines: a line still being cut by a crash waits for the next check
            If Right(txt, 1) = Chr(10) Then
                ln = Split(Left(txt, Len(txt) - 1), Chr(10))
                For k = 0 To UBound(ln)
                    If Right(ln(k), 1) = Chr(13) Then ln(k) = Left(ln(k), Len(ln(k)) - 1)
                    If WmsJournal.ParseHeader(CStr(ln(k)), seq, inst, sType) Then
                        nLines = nLines + 1
                        p = InStr(IIf(Len(ln(k)) > 60, Len(ln(k)) - 60, 1), ln(k), ";CH=")
                        If p > 0 And prev <> "" Then
                            q = InStr(p + 4, ln(k), ";")
                            ch = Mid(ln(k), p + 4, q - p - 4)
                            If ch <> prev Then
                                nBad = nBad + 1
                                If nBad <= 5 Then s = s & IIf(s <> "", "; ", "") & "операция журнала № " & seq & " (" & files(i) & ") не продолжает цепочку — строка " _
                                    & "перед ней изменена, удалена или вставлена не WMS"
                            End If
                        End If
                        prev = WmsJournal.LineHash(CStr(ln(k)))
                    Else
                        prev = ""
                    End If
                Next k
                lastOff = sz
            End If
        End If
    Next i
    StSet("CHAIN", lastFile & "|" & NumStr(lastOff) & "|" & prev)
    If nBad > 5 Then s = s & "; и ещё " & (nBad - 5)
    ChainCheck = s
    Exit Function
EH:
    ChainCheck = "цепочка журнала не проверена: " & Error$
End Function

' ================================================================ Н5 the check of the day

' at the first start of a day (or bForce): «остаток = движения», the packing, the keys of the sheets, the chain of the
' journal, the mirror. The result kept: DAILY=date|OK or FAIL|text|ms; returned as is
Function DailyCheck(bForce As Boolean) As String
    Dim today As String, t0 As Long, p As String, probs As String, n As Long, res As String
    WmsInit()
    today = Format(Now(), "YYYY-MM-DD")
    If Not bForce And Left(StGet("DAILY"), 10) = today Then
        DailyCheck = StGet("DAILY")
        Exit Function
    End If
    On Error GoTo EH
    t0 = GetSystemTicks()
    WmsView.ProgressStart("Проверка дня: остаток = движения, фасовка, ключи, журнал, формулы Calc", 1)
    p = WmsLegacy.BalanceCheck()
    If p <> "" Then probs = "остаток не равен движениям: " & p
    p = WmsPack.PackProblems()
    If p <> "" Then probs = probs & IIf(probs <> "", " · ", "") & "фасовка: " & p
    p = WmsDiagnostics.FullKeyCheck(False)
    If Left(p, 6) = "ОШИБКА" Then probs = probs & IIf(probs <> "", " · ", "") & p
    p = ChainCheck(n)
    If p <> "" And Left(p, 20) <> "контрольная цепочка " Then probs = probs & IIf(probs <> "", " · ", "") & "журнал: " & p
    If gJMirrorErr <> "" Then probs = probs & IIf(probs <> "", " · ", "") & "зеркало журнала: " & gJMirrorErr
    ' Т8: Calc counts as WMS expects
    p = CalcCheck()
    If p <> "" Then
        probs = probs & IIf(probs <> "", " · ", "") & p
    Else
        StSet("LOVER", WmsStatus.LoVersion())
    End If
    res = today & "|" & IIf(probs = "", "OK", "FAIL") & "|" & IIf(probs = "", "остаток = движения, фасовка, ключи листов, цепочка журнала (" & n _
        & " новых строк), формулы Calc — без замечаний", probs) & "|" & (GetSystemTicks() - t0)
    StSet("DAILY", res)
    DailyCheck = res
    WmsView.ProgressEnd()
    Exit Function
EH:
    WmsView.ProgressEnd()
    res = today & "|FAIL|проверка дня не выполнена: " & Error$ & "|" & (GetSystemTicks() - t0)
    StSet("DAILY", res)
    DailyCheck = res
End Function

' the line of the check of the day for the pages: «Проверка дня 20.10: ✓ без замечаний» / «✗ …»; "" — not made today
Function DailyLine() As String
    Dim a As Variant
    a = Split(StGet("DAILY"), "|")
    If UBound(a) < 2 Then Exit Function
    DailyLine = "Проверка дня " & Mid(a(0), 9, 2) & "." & Mid(a(0), 6, 2) & ": " & IIf(a(1) = "OK", "✓ " & a(2), "✗ " & a(2))
End Function

Sub BtnDailyCheck(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = DailyCheck(True)
    WmsUi.UiMessage(DailyLine())
End Sub

' ================================================================ Т8 the check of Calc, Т7 the restore drill, Т6 the health

' Т8: formulas with a known answer, through the same engine the pages and the reports count with (an array formula on _IDX):
' the sum of products, the search in a list, the rounding, the text, the error caught, the dates, the remainder, the maximum
Private Const SF_CALC_BODY = "TEXTJOIN(""|"";0;SUMPRODUCT({1;2;3}*{4;5;6});MATCH(""b"";{""a"";""b"";""c""};0);ROUND(2.5;0);LEFT(""ЕИ-427"";2);LEN(""ЕИ-427"");IFERROR(1/0;""x"");INDEX({5;6;7};2);DATE(2026;10;20)-DATE(2026;1;1);MOD(7;3);SUMPRODUCT(ISNUMBER({1;""a"";3})*1);MAX({4;9;2});ROUND((0.1+0.2)*10;0))"
Private Const SF_CALC_WANT = "32|2|3|ЕИ|6|x|6|292|1|2|9|3"
Public Const DRILL_DAYS = 31
' test seam (TestCalcWant): the answer the check of Calc expects, replaced
Global gSfCalcWant As String

' "" — Calc counts as WMS expects; otherwise what is wrong (the posting does not depend on it: only the pages and the reports)
Function CalcCheck() As String
    Dim s As String, ok As Boolean, want As String
    On Error GoTo EH
    want = SF_CALC_WANT
    If gSfCalcWant <> "" Then want = gSfCalcWant
    s = WmsOrders.ScratchEval(SF_CALC_BODY, ok)
    If Not ok Then
        CalcCheck = "формулы Calc не вычисляются (LibreOffice " & WmsStatus.LoVersion() & ") — страницы и отчёты не построятся; проведение не затронуто"
    ElseIf s <> want Then
        CalcCheck = "Calc (LibreOffice " & WmsStatus.LoVersion() & ") считает не так, как ждёт WMS: «" & s & "» вместо «" & want _
            & "» — страницы и отчёты могут ошибаться; проведение не затронуто. Верните прежнюю версию LibreOffice"
    End If
    Exit Function
EH:
    CalcCheck = "формулы Calc не проверены: " & Error$
End Function

' at a start: a new version of LibreOffice checked at once (not at the next day); "" — nothing to say
Function CalcAtStart() As String
    Dim v As String
    v = WmsStatus.LoVersion()
    If StGet("LOVER") = v Then Exit Function
    CalcAtStart = CalcCheck()
    If CalcAtStart = "" Then StSet("LOVER", v)
End Function

Sub BtnRestoreDrill(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    If gUiAuto = 0 Then
        If WmsPrimeUi.UiButtons("Учебное восстановление", "Самая свежая копия книги откроется отдельно (рабочая книга не меняется), WMS проверит, " _
            & "что она читается, и что журнал от неё до последней операции непрерывен: при беде «Вернуться к копии» вернёт всё." & Chr(10) _
            & "Большая книга — до минуты.", Array("Проверить", "Отмена"), Array(1, 0)) <> 1 Then Exit Sub
    End If
    s = RestoreDrill()
    WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    ServiceRefresh()
End Sub

' Т7: the newest copy of the book is unpacked apart (WMS_State/drill, opened hidden, read-only, its macros off) and read: its
' last operation, its EIs; the journal from it to the last operation of the book read in full: continuous, of this book, not
' damaged. «OK:…» — a return to this copy would bring back every operation; «ERR:…» — what stands in the way
Function RestoreDrill() As String
    Dim lst As Variant, src As String, dst As String, ddir As String, oDoc As Object, props(2) As New com.sun.star.beans.PropertyValue
    Dim cSeq As Long, cSchema As String, cEI As Double, last As Long, t0 As Long, s As String, sh As Object, c As Object, res As String
    WmsInit()
    t0 = GetSystemTicks()
    On Error GoTo EH
    lst = CopyList()
    If UBound(lst) < 0 Then
        res = "ERR:в «" & BACKUP_DIR & "» нет ни одной копии книги — сохраните книгу (копия делается при сохранении) или «Копия на флешку…»"
        GoTo DONE
    End If
    WmsView.ProgressStart("Учебное восстановление: копия " & CopyLabel(lst(0)), 3)
    src = gBDir & lst(0)(0)
    ddir = SafeDir() & "drill/"
    If gSfa.exists(ddir) Then gSfa.kill(ddir)
    gSfa.createFolder(ddir)
    dst = ddir & "drill.ods"
    gSfa.copy(src, dst)
    If gSfa.getSize(dst) <> gSfa.getSize(src) Then
        res = "ERR:копия «" & lst(0)(0) & "» не скопировалась целиком (мало места на диске?)"
        GoTo DONE
    End If
    WmsView.ProgressSet(1)
    props(0).Name = "Hidden"
    props(0).Value = True
    props(1).Name = "ReadOnly"
    props(1).Value = True
    props(2).Name = "MacroExecutionMode"
    props(2).Value = com.sun.star.document.MacroExecMode.NEVER_EXECUTE
    oDoc = StarDesktop.loadComponentFromURL(dst, "_blank", 0, props())
    If IsNull(oDoc) Then
        res = "ERR:копия «" & lst(0)(0) & "» не открывается — она повреждена; возьмите более раннюю («Вернуться к копии…»)"
        GoTo DONE
    End If
    If Not oDoc.Sheets.hasByName(SYS_SHEET) Or Not oDoc.Sheets.hasByName(SH_STOCK) Then
        res = "ERR:в копии «" & lst(0)(0) & "» нет служебного листа или «" & SH_STOCK & "» — это не книга WMS"
        GoTo DONE
    End If
    sh = oDoc.Sheets.getByName(SYS_SHEET)
    cSchema = sh.getCellByPosition(1, SK_SCHEMA).getString()
    c = sh.getCellByPosition(1, SK_LAST_SEQ)
    If c.getType() <> com.sun.star.table.CellContentType.VALUE Then
        res = "ERR:в копии «" & lst(0)(0) & "» не читается номер последней операции"
        GoTo DONE
    End If
    cSeq = c.getValue()
    c = sh.getCellByPosition(1, SK_NEXT_EI)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then cEI = c.getValue() - 1
    If sh.getCellByPosition(1, SK_INSTANCE).getString() <> SysStr(SK_INSTANCE) Then
        res = "ERR:копия «" & lst(0)(0) & "» — от другой книги WMS"
        GoTo DONE
    End If
    oDoc.close(True)
    oDoc = Nothing
    WmsView.ProgressSet(2)
    last = SysNum(SK_LAST_SEQ)
    If cSeq > last Then
        res = "ERR:в копии операций больше (" & cSeq & "), чем в книге (" & last & ") — книга старше своей копии, проверьте «Проверка перед работой»"
        GoTo DONE
    End If
    WmsJournal.ReadJournal(cSeq, True, False)
    If gJR_Error <> "" Then
        s = gJR_Error
    ElseIf gJR_Order <> "" Then
        s = gJR_Order
    ElseIf WmsJournal.StartProblem(cSeq) <> "" Then
        s = WmsJournal.StartProblem(cSeq)
    ElseIf cSeq > 0 And Not gJR_HasLast Then
        s = "в журнале нет операции " & cSeq & ", на которой сделана копия"
    ElseIf gJR_N <> last - cSeq Or (gJR_N > 0 And gJR_MaxSeq <> last) Then
        s = "в журнале после копии " & gJR_N & " операций, а нужно " & (last - cSeq) & " (до " & last & ")"
    ElseIf gJR_Damaged > 0 Then
        s = "повреждённых строк журнала: " & gJR_Damaged & " — «Журнал из зеркала»"
    End If
    If s <> "" Then
        res = "ERR:от копии «" & lst(0)(0) & "» восстановить всё нельзя: " & s
    Else
        res = "OK:восстанавливается. Копия " & CopyLabel(lst(0)) & " открывается и читается (ЕИ: " & NumStr(cEI) _
            & IIf(cSchema = SysStr(SK_SCHEMA), "", ", схема " & cSchema) & "); журнал после неё непрерывен: операций " & NumStr(gJR_N) _
            & ", до последней № " & NumStr(last) & ", без разрывов и чужих строк. При беде «Вернуться к копии…» вернёт книгу к ней и доведёт по журналу."
    End If
DONE:
    On Error Resume Next
    If Not IsNull(oDoc) And Not IsEmpty(oDoc) Then oDoc.close(True)
    If ddir <> "" Then
        If gSfa.exists(ddir) Then gSfa.kill(ddir)
    End If
    WmsView.ProgressEnd()
    StSet("DRILL", Format(Now(), "YYYY-MM-DD HH:MM") & "|" & Left(res, InStr(res, ":") - 1) & "|" & Mid(res, InStr(res, ":") + 1) & "|" & (GetSystemTicks() - t0))
    RestoreDrill = res
    Exit Function
EH:
    res = "ERR:учебное восстановление не выполнено: " & Error$
    Resume DONE
End Function

' the line of the last drill: «20.10.2026: ✓ восстанавливается…» / «✗ …» / "" — never made
Function DrillLine() As String
    Dim a As Variant
    a = Split(StGet("DRILL"), "|")
    If UBound(a) < 2 Then Exit Function
    DrillLine = IIf(a(1) = "OK", "✓ ", "✗ ") & Mid(a(0), 9, 2) & "." & Mid(a(0), 6, 2) & "." & Left(a(0), 4) & ": " & a(2)
End Function

' the days since the last drill (999 — never)
Function DrillAge() As Double
    Dim s As String
    s = StGet("DRILL")
    DrillAge = 999
    If Len(s) < 10 Then Exit Function
    DrillAge = CDbl(Now()) - CDbl(DateSerial(Val(Left(s, 4)), Val(Mid(s, 6, 2)), Val(Mid(s, 9, 2))))
End Function

' Т6: one light for the whole WMS — 0 green, 1 yellow (do it soon), 2 red (do it now); sText — what and what to do
Function Health(ByRef sText As String) As Integer
    Dim red As String, yel As String, a As Variant, x As String, free As Double
    On Error GoTo EH
    If gState <> "CLEAN" Then red = "работа остановлена (" & gState & IIf(gBlock <> "", " " & gBlock, "") & ") — «Проверка перед работой», затем «Восстановить»"
    a = Split(StGet("DAILY"), "|")
    If UBound(a) < 2 Then
        yel = "проверка дня не выполнялась — «Проверить сейчас»"
    ElseIf a(1) <> "OK" Then
        red = red & IIf(red <> "", " · ", "") & "проверка дня: " & a(2) & " — «Проверить сейчас», затем «Проверка перед работой»"
    ElseIf Left(a(0), 10) <> Format(Now(), "YYYY-MM-DD") Then
        yel = "проверка дня сегодня не выполнялась — «Проверить сейчас»"
    End If
    free = DiskFreeMb(gBaseDir)
    If free >= 0 And free < DISK_WARN_MB Then red = red & IIf(red <> "", " · ", "") & "на диске свободно " & Format(free, "0") & " МБ — освободите место"
    If gJMirror <> "" And gJMirrorErr <> "" Then red = red & IIf(red <> "", " · ", "") & "зеркало журнала не пишется: " & gJMirrorErr
    x = FlashNote()
    If x <> "" Then yel = yel & IIf(yel <> "", " · ", "") & x
    If UBound(CopyList()) < 0 Then yel = yel & IIf(yel <> "", " · ", "") & "копий книги в «" & BACKUP_DIR & "» нет — сохраните книгу"
    If DrillAge() >= DRILL_DAYS Then
        yel = yel & IIf(yel <> "", " · ", "") & "учебное восстановление " & IIf(DrillAge() >= 999, "ещё не делалось", "было " & Int(DrillAge()) & " дн. назад") _
            & " — раз в месяц: «Учебное восстановление»"
    ElseIf Left(StGet("DRILL"), 1) <> "" And InStr(StGet("DRILL"), "|OK|") = 0 Then
        red = red & IIf(red <> "", " · ", "") & "учебное восстановление не прошло — " & Mid(DrillLine(), InStr(DrillLine(), ":") + 2)
    End If
    If gMpN > 0 And gMpNOff > 0 Then
        yel = yel & IIf(yel <> "", " · ", "") & "ЕИ с остатком не на карте: " & gMpNOff & " — «Карта склада» → «Не на карте»"
    End If
EH:
    If red <> "" Then
        Health = 2
        sText = "✗ Нужно сейчас: " & red & IIf(yel <> "", " · Скоро: " & yel, "")
    ElseIf yel <> "" Then
        Health = 1
        sText = "! Скоро: " & yel
    Else
        Health = 0
        sText = "✓ Всё в порядке: проверки, копии, журнал, диск"
    End If
End Function

' ================================================================ Н7 the operations put aside («Отложить хвост»)

Global gSfAbN As Long
Global gSfAbSeq() As Long
Global gSfAbType() As String
Global gSfAbFields() As Variant

' the operations of every ABANDON of the journal (the entries FROM…TO were never brought into the book); their number
Function AbandonedRead() As Long
    Dim files As Variant, i As Long, k As Long, txt As String, ln As Variant, seq As Long, ts As String, inst As String, sType As String
    Dim fields As Variant, nR As Long, rFrom() As Long, rTo() As Long, j As Long, f As Variant, allSeq() As Long, allType() As String
    Dim allTs() As String, allF() As Variant, nAll As Long, cap As Long
    gSfAbN = 0
    files = WmsJournal.ListFiles()
    cap = 1000
    ReDim allSeq(cap)
    ReDim allType(cap)
    ReDim allTs(cap)
    ReDim allF(cap)
    ' every entry once (in a ring of the newest 1000 before each ABANDON is enough: a tail put aside is a few operations)
    For i = 0 To UBound(files)
        txt = WmsAudit.ReadText(gJDir & files(i))
        ln = Split(Replace(txt, Chr(13), ""), Chr(10))
        For k = 0 To UBound(ln)
            If ln(k) <> "" Then
                If WmsJournal.ParseLine(CStr(ln(k)), seq, ts, inst, sType, fields) Then
                    If sType = "ABANDON" Then
                        For j = 0 To UBound(fields)
                            f = Split(fields(j), "=")
                            If f(0) = "FROM" Then
                                nR = nR + 1
                                ReDim Preserve rFrom(nR - 1)
                                ReDim Preserve rTo(nR - 1)
                                rFrom(nR - 1) = Val(f(1))
                            ElseIf f(0) = "TO" Then
                                rTo(nR - 1) = Val(f(1))
                            End If
                        Next j
                        ' the entries of the range seen so far (they stand before the ABANDON)
                        For j = 0 To cap
                            If allSeq(j) >= rFrom(nR - 1) And allSeq(j) <= rTo(nR - 1) And allSeq(j) > 0 Then AbAdd(allSeq(j), allType(j), allTs(j), allF(j))
                        Next j
                    ElseIf sType <> "START" Then
                        allSeq(nAll Mod (cap + 1)) = seq
                        allType(nAll Mod (cap + 1)) = sType
                        allTs(nAll Mod (cap + 1)) = ts
                        allF(nAll Mod (cap + 1)) = fields
                        nAll = nAll + 1
                    End If
                End If
            End If
        Next k
    Next i
    AbandonedRead = gSfAbN
End Function

Private Sub AbAdd(seq As Long, sType As String, ts As String, fields As Variant)
    Dim i As Long
    For i = 0 To gSfAbN - 1
        If gSfAbSeq(i) = seq Then Exit Sub
    Next i
    ReDim Preserve gSfAbSeq(gSfAbN)
    ReDim Preserve gSfAbType(gSfAbN)
    ReDim Preserve gSfAbFields(gSfAbN)
    gSfAbSeq(gSfAbN) = seq
    gSfAbType(gSfAbN) = sType & "|" & ts
    gSfAbFields(gSfAbN) = fields
    gSfAbN = gSfAbN + 1
End Sub

' the value of field k of entry i of the operations put aside
Function AbField(i As Long, k As String) As String
    Dim f As Variant, j As Long
    f = gSfAbFields(i)
    For j = 0 To UBound(f)
        If Left(f(j), Len(k) + 1) = k & "=" Then
            AbField = Unesc(Mid(f(j), Len(k) + 2))
            Exit Function
        End If
    Next j
End Function

Sub BtnAbandoned(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = AbandonedShow()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsView.GoCell(SH_FIND, FD_TABLE, 1)
End Sub

' «Отложенные операции» on «Поиск»: what «Отложить хвост» left out of the book, the newest first; a selected issue goes back to
' «Выдачи» by «Вернуть в работу»
Function AbandonedShow() As String
    Dim n As Long, i As Long, t() As Variant, x As Variant, k As Long
    On Error GoTo EH
    n = AbandonedRead()
    If n = 0 Then
        AbandonedShow = WmsFind.ShowTable("Отложенных операций нет", "Операции попадают сюда, когда хвост журнала откладывают («Отложить хвост журнала»): " _
            & "их нет в книге, их проводят заново.", "", Array("№ журнала", "Когда", "Операция", "ЕИ", "Количество", "Кто / место", "Подробно", ""), Array(), 0, False)
        Exit Function
    End If
    ReDim t(n - 1)
    For i = 0 To n - 1
        k = n - 1 - i
        x = Split(gSfAbType(k), "|")
        t(i) = Array(gSfAbSeq(k), Replace(x(1), "T", " "), WmsView.OpName(CStr(x(0))), AbField(k, "EI"), AbField(k, "QTY") & IIf(AbField(k, "UNIT") <> "", " " _
            & AbField(k, "UNIT"), ""), IIf(AbField(k, "WHO") <> "", AbField(k, "WHO"), AbField(k, "PLACE")), IIf(AbField(k, "DATE") <> "", "дата " _
            & AbField(k, "DATE"), ""), IIf(x(0) = "ISSUE", "можно вернуть", "провести вручную"), "", "", gSfAbSeq(k), "")
    Next i
    AbandonedShow = WmsFind.ShowTable("Отложенные операции: " & n, "Этих операций нет в книге: хвост журнала был отложен. Проверьте по бумагам и проведите " _
        & "заново. Выдачи: выделите строки и нажмите «Вернуть в работу» — они встанут непроведёнными строками в «" & SH_ISSUES & "».", "Новые сверху.", _
        Array("№ журнала", "Когда", "Операция", "ЕИ", "Количество", "Кто / место", "Подробно", "Как вернуть"), t, n, False)
    Exit Function
EH:
    AbandonedShow = "ERR:список не построен: " & Error$
End Function

Sub BtnAbandonedRedo(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = AbandonedRedo(WmsPrint.SelRows(SH_FIND))
    If Left(s, 3) = "OK:" Then WmsOrdersUi.UiInfo(Mid(s, 4)) Else WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
End Sub

' the issues among rows of «Поиск» (the list of the operations put aside) back as unposted rows of «Выдачи». OK:… / ERR:…
Function AbandonedRedo(rows As Variant) As String
    Dim fs As Object, ish As Object, i As Long, j As Long, seq As Long, r As Long, n As Long, nOther As Long, d As Double, dd As String
    WmsInit()
    On Error GoTo EH
    If gSfAbN = 0 Then AbandonedRead()
    fs = gDoc.Sheets.getByName(SH_FIND)
    ish = WmsIssue.IssuesSheet()
    r = WmsOrders.LastRow(ish) + 1
    For i = 0 To UBound(rows)
        seq = CLng(fs.getCellByPosition(FD_SRC_ROW, rows(i)).getValue())
        For j = 0 To gSfAbN - 1
            If gSfAbSeq(j) = seq And seq > 0 Then
                If Left(gSfAbType(j), 6) = "ISSUE|" Then
                    ish.getCellByPosition(IC_EI, r).setString(AbField(j, "EI"))
                    ish.getCellByPosition(IC_QTY, r).setString(Replace(AbField(j, "QTY"), ".", ","))
                    dd = AbField(j, "DATE")
                    If Len(dd) = 10 Then ish.getCellByPosition(IC_DATE, r).setValue(CDbl(DateSerial(Val(Left(dd, 4)), Val(Mid(dd, 6, 2)), Val(Mid(dd, 9, 2)))))
                    ish.getCellByPosition(IC_WHO, r).setString(AbField(j, "WHO"))
                    If AbField(j, "DOC") <> "" Then ish.getCellByPosition(IC_DOC, r).setString(AbField(j, "DOC"))
                    ish.getCellByPosition(IC_NOTE, r).setString("отложенная операция журнала № " & seq)
                    r = r + 1
                    n = n + 1
                Else
                    nOther = nOther + 1
                End If
            End If
        Next j
    Next i
    If n = 0 Then
        AbandonedRedo = "ERR:среди выделенных строк нет выдач" & IIf(nOther > 0, " (остальные операции проведите вручную по их описанию)", " — выделите строки списка отложенных операций")
        Exit Function
    End If
    WmsView.GoCell(SH_ISSUES, r - n, IC_EI)
    AbandonedRedo = "OK:Выдач возвращено в «" & SH_ISSUES & "»: " & n & " (строки " & (r - n + 1) & "…" & r & ", не проведены). Проверьте их и нажмите «Провести»." _
        & IIf(nOther > 0, Chr(10) & "Других операций среди выделенных: " & nOther & " — проведите их вручную.", "")
    gUiLastMsg = AbandonedRedo
    Exit Function
EH:
    AbandonedRedo = "ERR:не возвращено: " & Error$
    gUiLastMsg = AbandonedRedo
End Function

' ================================================================ the page «Обслуживание»

Public Const SH_SERVICE = "Обслуживание"
Public Const SV_STATE_ROW = 3

Sub NavService(Optional oEvent As Variant)
    WmsInit()
    ServiceRefresh()
    WmsView.GoCell(SH_SERVICE, 0, 0)
End Sub

Sub OnServiceFocus(Optional oEvent As Variant)
    On Error Resume Next
    ServiceRefresh()
End Sub

' the state lines of the page: the copy on a medium, the mirror, the check of the day, the copies, the disk, the book
Function ServiceRefresh() As String
    Dim sh As Object, v(8) As String, s As String, a As Variant, lst As Variant, free As Double, files As Variant, i As Long, jMb As Double
    Dim bMb As Double, d As Double, nB As Long
    WmsInit()
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_SERVICE) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_SERVICE)
    s = StGet("FLASH")
    If s = "" Then
        v(0) = "✗ ещё не было — «Копия на флешку…»"
    Else
        a = Split(s, "|")
        d = CDbl(Now()) - CDbl(DateSerial(Val(Left(a(0), 4)), Val(Mid(a(0), 6, 2)), Val(Mid(a(0), 9, 2))))
        v(0) = IIf(d < FLASH_REMIND_DAYS, "✓ ", "✗ ") & Mid(a(0), 9, 2) & "." & Mid(a(0), 6, 2) & "." & Left(a(0), 4) & " " & Mid(a(0), 12, 5) _
            & IIf(UBound(a) >= 1, " — " & a(1), "") & IIf(d >= FLASH_REMIND_DAYS, " (" & Int(d) & " дн. назад — пора сделать новую)", "")
    End If
    If StGet("MIRROR") = "" Then
        v(1) = "выключено — «Зеркало журнала…»"
    Else
        v(1) = IIf(gJMirrorErr = "", "✓ ", "✗ ") & ConvertFromURL(MirrorFolder(StGet("MIRROR"))) & IIf(gJMirrorErr <> "", " — " & gJMirrorErr, "")
    End If
    v(2) = DailyLine()
    If v(2) = "" Then v(2) = "сегодня не выполнялась — «Проверить сейчас»" Else v(2) = Mid(v(2), InStr(v(2), ":") + 2)
    lst = CopyList()
    If gSfa.exists(gBDir) Then
        files = gSfa.getFolderContents(gBDir, False)
        For i = 0 To UBound(files)
            If LCase(Right(files(i), 4)) = ".ods" Then nB = nB + 1
        Next i
    End If
    v(3) = nB & " в «" & BACKUP_DIR & "»" & IIf(UBound(lst) >= 0, ", последняя: " & CopyLabel(lst(0)), "")
    free = DiskFreeMb(gBaseDir)
    v(4) = IIf(free < 0, "не определено", IIf(free >= DISK_WARN_MB, "✓ ", "✗ ") & Format(free / 1024, "0.0") & " ГБ")
    bMb = BookMb()
    files = WmsJournal.ListFiles()
    For i = 0 To UBound(files)
        jMb = jMb + gSfa.getSize(gJDir & files(i)) / 1048576
    Next i
    v(5) = "книга " & Format(bMb, "0.0") & " МБ, журнал " & Format(jMb, "0.0") & " МБ (" & (UBound(files) + 1) & " мес.), операций " & NumStr(SysNum(SK_LAST_SEQ))
    ' С3, С4, С2: the growth, the times, the speed of this computer
    v(6) = WmsCare.GrowthText()
    v(7) = WmsCare.TimingText()
    a = Split(StGet("SPEED"), "|")
    If UBound(a) >= 1 Then
        v(8) = "держит до ~" & WmsView.NumText(Val(a(1))) & " движений (замер " & Mid(a(0), 9, 2) & "." & Mid(a(0), 6, 2) & "." & Left(a(0), 4) & ")"
    Else
        v(8) = "не замерялась — «Замер скорости ПК»"
    End If
    WmsView.ViewOpen(SH_SERVICE)
    For i = 0 To 8
        sh.getCellByPosition(2, SV_STATE_ROW + i).setString(v(i))
    Next i
    ' Т6: the light of the whole WMS above the lines; Т7: the last drill
    Health(s)
    sh.getCellByPosition(1, SV_STATE_ROW - 1).setString("Здоровье WMS")
    sh.getCellByPosition(2, SV_STATE_ROW - 1).setString(s)
    s = DrillLine()
    If s = "" Then s = "✗ ещё не делалось — «Учебное восстановление» (раз в месяц)"
    If DrillAge() >= DRILL_DAYS And Left(s, 1) = "✓" Then s = "✗" & Mid(s, 2) & " (" & Int(DrillAge()) & " дн. назад — пора снова)"
    sh.getCellByPosition(1, SV_STATE_ROW + 9).setString("Учебное восстановление")
    sh.getCellByPosition(2, SV_STATE_ROW + 9).setString(s)
    WmsView.ViewDone(sh)
    ServiceRefresh = "OK"
    Exit Function
EH:
    ServiceRefresh = "ERR:" & Error$
    On Error Resume Next
    WmsView.ViewDone(sh)
End Function

' the troubles of the safety of the data for «Внимание» of «Дашборд»: rows Array(what, detail); their number
Function SafeAttention(ByRef rows As Variant) As Long
    Dim out(2) As Variant, n As Long, a As Variant, s As String
    rows = Array()
    On Error GoTo EH
    a = Split(StGet("DAILY"), "|")
    If UBound(a) >= 2 Then
        If a(1) <> "OK" Then
            out(n) = Array("Проверка дня: замечания", a(2))
            n = n + 1
        End If
    End If
    s = FlashNote()
    If s <> "" Then
        out(n) = Array("Нет свежей копии на флешке", s)
        n = n + 1
    End If
    If gJMirror <> "" And gJMirrorErr <> "" Then
        out(n) = Array("Зеркало журнала не пишется", gJMirrorErr)
        n = n + 1
    End If
EH:
    If n > 0 Then
        ReDim Preserve out(n - 1)
        rows = out
    End If
    SafeAttention = n
End Function

' ================================================================ test seams (TEST books only)

' the folder «Копия на флешку» and «Зеркало журнала» take without the window ("" — the window), the free space reported (MB, 0 — measured)
Function TestSafe(sFolder As String, nDiskMb As Double) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSafe = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gSfFlashTest = sFolder
    gSfDiskTest = nDiskMb
    TestSafe = "OK"
End Function

' the mark («время/длина») of the saved journal position replaced (the clock watchdog); the old value of the position returned
Function TestJposMark(sMark As String) As String
    Dim a As Variant, old As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then Exit Function
    old = SysStr(SK_JPOS)
    a = Split(old, "|")
    If UBound(a) >= 2 Then SysPutStr(SK_JPOS, a(0) & "|" & a(1) & "|" & a(2) & "|" & sMark)
    TestJposMark = old
End Function

Function TestJposSet(s As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then Exit Function
    SysPutStr(SK_JPOS, s)
    TestJposSet = "OK"
End Function

' test seam: the rows of SafeAttention
Function TestSafeAttention() As Variant
    Dim rows As Variant
    SafeAttention(rows)
    TestSafeAttention = rows
End Function

' Т8: the answer the check of Calc expects replaced ("" — the real one), to see a Calc that counts otherwise
Function TestCalcWant(s As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestCalcWant = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gSfCalcWant = s
    TestCalcWant = "OK"
End Function
