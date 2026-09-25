' WmsLock — the WMS lock (one working session per book) and the registered working path (spec §22).
' The lock file lives in WMS_Journal and names its owner session; it is released only by its owner, or by the
' explicit action "Снять блокировку WMS" after a crash.
Option Explicit

Function LockUrl() As String
    LockUrl = gJDir & LOCK_FILE
End Function

Private Function OurLockText() As String
    OurLockText = "WMSLOCK1;session=" & gSession & ";instance=" & SysStr(SK_INSTANCE) & ";opened=" & TS() _
        & ";user=" & Environ("USER") & Environ("USERNAME") & ";host=" & Environ("HOSTNAME") & Environ("COMPUTERNAME") & ";"
End Function

Function IsOurs(sLock As String) As Boolean
    IsOurs = (InStr(sLock, ";session=" & gSession & ";") > 0)
End Function

Private Function FieldOf(sLock As String, sName As String) As String
    Dim p As Long, q As Long
    p = InStr(sLock, ";" & sName & "=")
    If p = 0 Then Exit Function
    p = p + Len(sName) + 2
    q = InStr(p, sLock, ";")
    If q = 0 Then q = Len(sLock) + 1
    FieldOf = Mid(sLock, p, q - p)
End Function

Function DescribeLock(sLock As String) As String
    Dim s As String
    s = "открыта " & Replace(FieldOf(sLock, "opened"), "T", " ")
    If FieldOf(sLock, "user") <> "" Then s = s & ", пользователь " & FieldOf(sLock, "user")
    If FieldOf(sLock, "host") <> "" Then s = s & ", компьютер " & FieldOf(sLock, "host")
    DescribeLock = s
End Function

' "" when this session owns the lock (created now or already ours); otherwise the reason another session holds it
Function Acquire() As String
    Dim url As String, s As String
    On Error GoTo EH
    url = LockUrl()
    If Not gSfa.exists(gJDir) Then gSfa.createFolder(gJDir)
    If gSfa.exists(url) Then
        s = ReadTextFile(url)
        If IsOurs(s) Then
            gLockOwned = True
            Acquire = ""
        Else
            gLockOwned = False
            Acquire = "WMS уже открыта (" & DescribeLock(s) & ") или была закрыта аварийно. Если другой WMS точно нет — «Снять блокировку WMS»"
        End If
        Exit Function
    End If
    WriteTextFile(url, OurLockText() & Chr(10))
    ' two sessions starting at the same moment: the last writer wins, the other one sees a foreign lock
    s = ReadTextFile(url)
    gLockOwned = IsOurs(s)
    If gLockOwned Then
        Acquire = ""
    Else
        Acquire = "WMS одновременно запускается в другом окне (" & DescribeLock(s) & ")"
    End If
    Exit Function
EH:
    gLockOwned = False
    Acquire = "не удалось создать блокировку WMS: " & Error$
End Function

' before every posting: the lock must still exist and be ours (another session may have released it by force)
Function VerifyOwner() As String
    Dim url As String, s As String
    On Error GoTo EH
    If Not gLockOwned Then
        VerifyOwner = "этот экземпляр WMS не владеет блокировкой"
        Exit Function
    End If
    url = LockUrl()
    If Not gSfa.exists(url) Then
        gLockOwned = False
        VerifyOwner = "блокировка WMS снята другим экземпляром — проведение остановлено, перезапустите WMS"
        Exit Function
    End If
    s = ReadTextFile(url)
    If Not IsOurs(s) Then
        gLockOwned = False
        VerifyOwner = "блокировку WMS захватил другой экземпляр (" & DescribeLock(s) & ") — проведение остановлено"
        Exit Function
    End If
    VerifyOwner = ""
    Exit Function
EH:
    VerifyOwner = "не удалось проверить блокировку WMS: " & Error$
End Function

' clean close: delete the lock only when it is ours
Sub ReleaseIfOwner()
    Dim url As String
    On Error Resume Next
    If Not gLockOwned Then Exit Sub
    url = LockUrl()
    If gSfa.exists(url) Then
        If IsOurs(ReadTextFile(url)) Then gSfa.kill(url)
    End If
    gLockOwned = False
End Sub

' explicit action "Снять блокировку WMS" (the user confirms that no other WMS is running)
Function ForceRelease() As String
    Dim url As String, s As String
    On Error GoTo EH
    url = LockUrl()
    If url <> LOCK_FILE And gSfa.exists(url) Then
        s = ReadTextFile(url)
        gSfa.kill(url)
        ForceRelease = "блокировка WMS снята (" & DescribeLock(s) & ")"
    Else
        ForceRelease = "блокировки WMS не было"
    End If
    gLockOwned = False
    Exit Function
EH:
    ForceRelease = "не удалось снять блокировку WMS: " & Error$
End Function

' ---------------------------------------------------------------- registered working path
' A copy or backup opened directly must not post (spec §22).
' A book without any operation yet (LAST_SEQ = 0) registers itself on its first start — also when that first session
' was never saved; the journal checks that follow decide the rest (a tail found next to it is replayed by «Восстановить»,
' a journal of another WMS blocks). Never inside WMS_Backups: a backup opened there must stay a backup.
Function RegisteredProblem() As String
    Dim reg As String, url As String, inBackups As Boolean
    url = gDoc.getURL()
    reg = SysStr(SK_REG_URL)
    If reg = url Then
        RegisteredProblem = ""
        Exit Function
    End If
    inBackups = (InStr(url, "/" & BACKUP_DIR & "/") > 0)
    If reg = "" And Not inBackups Then
        If SysNum(SK_LAST_SEQ) = 0 Then
            SysPutStr(SK_REG_URL, url)
            AddNote("новая WMS: рабочий файл зарегистрирован")
            RegisteredProblem = ""
            Exit Function
        End If
        RegisteredProblem = "рабочий путь WMS не зарегистрирован — если это рабочая книга, выполните «Сделать рабочим файлом»"
        Exit Function
    End If
    If inBackups Then
        RegisteredProblem = "открыта резервная копия «" & FileNameOf(url) & "» — проводить в ней нельзя. Чтобы вернуться к ней, " _
            & "скопируйте её на место рабочей книги и выполните «Восстановить» (журнал доведёт книгу до актуального состояния)"
    Else
        RegisteredProblem = "это не рабочий файл WMS (рабочий: " & ConvertFromURL(reg) & "). Если WMS перенесена сюда намеренно — «Сделать рабочим файлом»"
    End If
End Function
