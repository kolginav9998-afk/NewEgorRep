' WmsBackup — copies of the working book (spec §26, D-003).
' A backup is the file as it was last saved. It never replaces the journal: after a backup is copied back over the
' working book, "Восстановить" replays the journal tail on top of it.
Option Explicit

' 1.0 (Н6): the mark of the next copy («перед загрузкой топологии»), set by WmsSafe.RestorePoint
Global gBkLabel As String
Public Const BACKUP_KEEP_POINT = 10

Function BookBaseName() As String
    Dim fn As String
    fn = FileNameOf(gDoc.getURL())
    If LCase(Right(fn, 4)) = ".ods" Then fn = Left(fn, Len(fn) - 4)
    BookBaseName = fn
End Function

Function KeepFor(kind As String) As Integer
    Select Case kind
    Case "daily"
        KeepFor = BACKUP_KEEP_DAILY
    Case "preupgrade"
        KeepFor = BACKUP_KEEP_UPGRADE
    Case "premigration"
        KeepFor = BACKUP_KEEP_MIGRATION
    Case "point"
        KeepFor = BACKUP_KEEP_POINT
    Case Else
        KeepFor = BACKUP_KEEP_MANUAL
    End Select
End Function

' 1.0: the kind as the storekeeper reads it (the name of the file keeps the code)
Private Function KindName(kind As String) As String
    Select Case kind
    Case "daily": KindName = "ежедневная"
    Case "preupgrade": KindName = "перед обновлением"
    Case "premigration": KindName = "перед переносом"
    Case "point": KindName = "точка восстановления" & IIf(gBkLabel <> "", " " & gBkLabel, "")
    Case Else: KindName = "ручная"
    End Select
End Function

' kind: daily | manual | preupgrade | premigration | point (1.0, Н6)
Function BackupNow(kind As String) As String
    Dim src As String, dst As String, d As Date, t0 As Long, stamp As String, nSrc As Double, nDst As Double, sClock As String, why As String
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        BackupNow = "резервная копия невозможна: книга не сохранена в файл"
        Exit Function
    End If
    If Not gSfa.exists(gBDir) Then gSfa.createFolder(gBDir)
    ' 1.0 (Н9): not enough free space — no copy (a copy cut short is worse than none; the older ones stay)
    why = WmsSafe.DiskRefuse(gBDir, gSfa.getSize(gDoc.getURL()) / 1048576)
    If why <> "" Then
        BackupNow = "ОШИБКА резервной копии: копия не сделана — " & why
        Exit Function
    End If
    t0 = GetSystemTicks()
    d = Now()
    stamp = Format(d, "YYYY-MM-DD") & "_" & Format(d, "HHMMSS")
    sClock = ClockProblem(stamp)
    src = gDoc.getURL()
    dst = gBDir & BookBaseName() & "_" & stamp & "_" & kind & ".ods"
    If gSfa.exists(dst) Then dst = gBDir & BookBaseName() & "_" & stamp & "_" & kind & "_" & NewId(4) & ".ods"
    gSfa.copy(src, dst)
    ' 1.0: a copy cut short (the disk full, the folder on a network that dropped) is not a copy: it is removed, and the
    ' older copies are kept — the rotation runs only after a complete one
    nSrc = gSfa.getSize(src)
    nDst = gSfa.getSize(dst)
    If nDst <> nSrc Or nDst <= 0 Then
        On Error Resume Next
        gSfa.kill(dst)
        BackupNow = "ОШИБКА резервной копии: копия получилась неполной (" & nDst & " из " & nSrc & " байт) и удалена; старые копии не тронуты. " _
            & "Проверьте место на диске и папку " & ConvertFromURL(gBDir)
        Exit Function
    End If
    Rotate(kind, KeepFor(kind), FileNameOf(dst))
    ' 1.0 (Н2): the list of the copies a return can be made to (with the number of the operations in the copy)
    WmsSafe.CopyIndexAdd(FileNameOf(dst), kind, gBkLabel)
    BackupNow = "резервная копия «" & KindName(kind) & "» (последнее сохранённое состояние): " & FileNameOf(dst) & ", " & (GetSystemTicks() - t0) & " мс" & sClock
    Exit Function
EH:
    BackupNow = "ОШИБКА резервной копии: " & Error$
End Function

' the date and time in the name of a copy of this book («YYYY-MM-DD_HHMMSS»), "" — not a copy of this book
Private Function StampOf(nm As String) As String
    Dim sPrefix As String, t As String
    sPrefix = BookBaseName() & "_"
    If Left(nm, Len(sPrefix)) <> sPrefix Or LCase(Right(nm, 4)) <> ".ods" Then Exit Function
    t = Mid(nm, Len(sPrefix) + 1, 17)
    If Len(t) = 17 And Mid(t, 5, 1) = "-" And Mid(t, 8, 1) = "-" And Mid(t, 11, 1) = "_" And IsNumeric(Left(t, 4)) Then StampOf = t
End Function

' 1.0: the computer clock behind the newest copy (a dead clock battery, a wrong year): the copies are ordered by the date
' in their names, so a wrong date is said aloud — the dates of the movements the WMS checks come from the same clock
Private Function ClockProblem(stamp As String) As String
    Dim files As Variant, i As Long, t As String, tMax As String
    On Error GoTo EH
    If Not gSfa.exists(gBDir) Then Exit Function
    files = gSfa.getFolderContents(gBDir, False)
    For i = 0 To UBound(files)
        t = StampOf(FileNameOf(files(i)))
        If t > tMax Then tMax = t
    Next i
    If tMax > stamp Then ClockProblem = "; ВНИМАНИЕ: часы компьютера показывают " & Format(Now(), "DD.MM.YYYY HH:MM") & ", а последняя копия сделана " _
        & Mid(tMax, 9, 2) & "." & Mid(tMax, 6, 2) & "." & Left(tMax, 4) & " — проверьте дату и время на компьютере"
    Exit Function
EH:
End Function

' one automatic copy per day, made by the lock owner at the first start of the day. 1.0: only a copy of this book counts,
' and only a whole one (an empty file left by a failed copy does not stop today's copy)
Function DailyIfNeeded() As String
    Dim files As Variant, i As Long, today As String, nm As String
    On Error GoTo EH
    today = Format(Now(), "YYYY-MM-DD")
    If gSfa.exists(gBDir) Then
        files = gSfa.getFolderContents(gBDir, False)
        For i = 0 To UBound(files)
            nm = FileNameOf(files(i))
            If Left(StampOf(nm), 10) = today And IsKind(nm, "daily") Then
                If gSfa.getSize(files(i)) > 0 Then
                    DailyIfNeeded = "ежедневная копия за сегодня уже есть"
                    Exit Function
                End If
            End If
        Next i
    End If
    DailyIfNeeded = BackupNow("daily")
    Exit Function
EH:
    DailyIfNeeded = "ОШИБКА ежедневной копии: " & Error$
End Function

' hooks for the future upgrade and migration functions: they must call these before changing anything
Function BeforeUpgrade() As String
    BeforeUpgrade = BackupNow("preupgrade")
End Function

Function BeforeMigration() As String
    BeforeMigration = BackupNow("premigration")
End Function

Private Function IsKind(nm As String, kind As String) As Boolean
    IsKind = (Right(nm, Len(kind) + 5) = "_" & kind & ".ods") Or (InStr(nm, "_" & kind & "_") > 0 And Right(nm, 4) = ".ods")
End Function

' keep the newest `keep` copies of this kind (names sort chronologically); other kinds are never touched. sNew — the copy
' just made: it is always kept (with the clock set back it would sort as the oldest and be removed at once)
Sub Rotate(kind As String, keep As Integer, Optional sNew As Variant)
    Dim files As Variant, names() As String, n As Long, i As Long, j As Long, t As String, nm As String, sKeep As String, nKeep As Integer
    On Error Resume Next
    If Not gSfa.exists(gBDir) Then Exit Sub
    If Not IsMissing(sNew) Then sKeep = sNew
    nKeep = keep
    If sKeep <> "" Then nKeep = keep - 1
    files = gSfa.getFolderContents(gBDir, False)
    n = -1
    For i = 0 To UBound(files)
        nm = FileNameOf(files(i))
        If Left(nm, Len(BookBaseName()) + 1) = BookBaseName() & "_" And IsKind(nm, kind) And nm <> sKeep Then
            n = n + 1
            ReDim Preserve names(n)
            names(n) = nm
        End If
    Next i
    If n + 1 <= nKeep Then Exit Sub
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
    For i = 0 To n - nKeep
        gSfa.kill(gBDir & names(i))
    Next i
End Sub
