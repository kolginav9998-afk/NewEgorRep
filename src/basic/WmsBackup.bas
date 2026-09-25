' WmsBackup — copies of the working book (spec §26, D-003).
' A backup is the file as it was last saved. It never replaces the journal: after a backup is copied back over the
' working book, "Восстановить" replays the journal tail on top of it.
Option Explicit

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
    Case Else
        KeepFor = BACKUP_KEEP_MANUAL
    End Select
End Function

' kind: daily | manual | preupgrade | premigration
Function BackupNow(kind As String) As String
    Dim src As String, dst As String, d As Date, t0 As Long, stamp As String
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        BackupNow = "резервная копия невозможна: книга не сохранена в файл"
        Exit Function
    End If
    If Not gSfa.exists(gBDir) Then gSfa.createFolder(gBDir)
    t0 = GetSystemTicks()
    d = Now()
    stamp = Format(d, "YYYY-MM-DD") & "_" & Format(d, "HHMMSS")
    src = gDoc.getURL()
    dst = gBDir & BookBaseName() & "_" & stamp & "_" & kind & ".ods"
    If gSfa.exists(dst) Then dst = gBDir & BookBaseName() & "_" & stamp & "_" & kind & "_" & NewId(4) & ".ods"
    gSfa.copy(src, dst)
    Rotate(kind, KeepFor(kind))
    BackupNow = "резервная копия " & kind & " (последнее сохранённое состояние): " & FileNameOf(dst) & ", " & (GetSystemTicks() - t0) & " мс"
    Exit Function
EH:
    BackupNow = "ОШИБКА резервной копии: " & Error$
End Function

' one automatic copy per day, made by the lock owner at the first start of the day
Function DailyIfNeeded() As String
    Dim files As Variant, i As Long, today As String, nm As String
    On Error GoTo EH
    today = "_" & Format(Now(), "YYYY-MM-DD") & "_"
    If gSfa.exists(gBDir) Then
        files = gSfa.getFolderContents(gBDir, False)
        For i = 0 To UBound(files)
            nm = FileNameOf(files(i))
            If InStr(nm, today) > 0 And IsKind(nm, "daily") Then
                DailyIfNeeded = "ежедневная копия за сегодня уже есть"
                Exit Function
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

' keep the newest `keep` copies of this kind (names sort chronologically); other kinds are never touched
Sub Rotate(kind As String, keep As Integer)
    Dim files As Variant, names() As String, n As Long, i As Long, j As Long, t As String, nm As String
    On Error Resume Next
    If Not gSfa.exists(gBDir) Then Exit Sub
    files = gSfa.getFolderContents(gBDir, False)
    n = -1
    For i = 0 To UBound(files)
        nm = FileNameOf(files(i))
        If Left(nm, Len(BookBaseName()) + 1) = BookBaseName() & "_" And IsKind(nm, kind) Then
            n = n + 1
            ReDim Preserve names(n)
            names(n) = nm
        End If
    Next i
    If n + 1 <= keep Then Exit Sub
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
    For i = 0 To n - keep
        gSfa.kill(gBDir & names(i))
    Next i
End Sub
