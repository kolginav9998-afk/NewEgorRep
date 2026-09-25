' WmsRecovery — fail-closed startup analysis and the explicit recovery actions (spec §18, §21, D-027, D-028).
' Startup states: CLEAN, or BLOCKED with one of NO_FILE, READ_ONLY, NOT_ODS, SYS_CORRUPT, NOT_REGISTERED (copy or old
' backup opened directly), LOCKED (stale lock / second instance), MARKER_CORRUPT, JOURNAL_UNREADABLE, JOURNAL_FOREIGN,
' JOURNAL_GAP, JOURNAL_BEHIND (journal mismatch), TAIL (unsaved journal tail), SAVED_WITHOUT_WMS, STARTUP_ERROR.
' Actions: «Восстановить» (Recover), «Отложить хвост журнала» (AbandonTail), «Снять блокировку WMS»
' (WmsLock.ForceRelease), «Сделать рабочим файлом» (RegisterHere).
Option Explicit

Sub Analyze()
    Dim p As String, st As String, lastSeq As Long, bi As String, ec As Double
    gState = "BLOCKED"
    ' 1. environment
    If Not gDoc.hasLocation() Then
        SetBlocked("NO_FILE", "книга не сохранена в файл")
        Exit Sub
    End If
    If gDoc.isReadonly() Then
        SetBlocked("READ_ONLY", "книга открыта только для чтения")
        Exit Sub
    End If
    If Not IsOdsDocument() Then
        SetBlocked("NOT_ODS", "рабочая книга WMS должна быть в формате ODS, открыт формат «" & DocFilterName() & "»")
        Exit Sub
    End If
    ' 2. service data
    p = SysLayoutProblem()
    If p = "" Then p = WmsIssue.SheetsProblem()
    If p = "" Then p = WmsOrders.SheetsProblem()
    If p <> "" Then
        SetBlocked("SYS_CORRUPT", p)
        Exit Sub
    End If
    AddNote("ядро " & WMS_CORE_VERSION & IIf(SysStr(SK_MODE) = "TEST", " (ТЕСТОВАЯ книга)", ""))
    ' 3. working path: a copy or a backup opened directly touches neither the lock, nor the journal, nor backups
    p = WmsLock.RegisteredProblem()
    If p <> "" Then
        SetBlocked("NOT_REGISTERED", p)
        Exit Sub
    End If
    ' 4. lock
    p = WmsLock.Acquire()
    If p <> "" Then
        SetBlocked("LOCKED", p)
        Exit Sub
    End If
    ' 5. this session owns the WMS from here on
    AddNote(WmsBackup.DailyIfNeeded())
    AddNote(WmsConfig.AutoInputDisable())
    ' 6. transaction marker: an operation that did not reach COMMITTED is undone from its before-image
    st = SysStr(SK_TX_STATE)
    If st = TX_STARTED Then
        bi = SysStr(SK_TX_BI)
        p = BeforeImageProblem(bi)
        If p <> "" Then
            SetBlocked("MARKER_CORRUPT", "маркер незавершённой операции повреждён (" & p & "). Восстановите книгу из резервной копии и выполните «Восстановить»")
            Exit Sub
        End If
        ' the operation in progress always follows the book's last one; anything else means LAST_SEQ moved without
        ' COMMITTED (rolling the cells back would then silently drop an operation the book counts as applied)
        If SysNum(SK_TX_SEQ) <> SysNum(SK_LAST_SEQ) + 1 Then
            SetBlocked("MARKER_CORRUPT", "маркер незавершённой операции seq " & SysStr(SK_TX_SEQ) & " не согласован с LAST_SEQ " _
                & SysStr(SK_LAST_SEQ) & ". Восстановите книгу из резервной копии и выполните «Восстановить»")
            Exit Sub
        End If
        If Not RestoreBeforeImage(bi) Then
            SetBlocked("MARKER_CORRUPT", "не удалось откатить незавершённую операцию seq " & SysStr(SK_TX_SEQ))
            Exit Sub
        End If
        AddNote("незавершённая операция seq " & SysStr(SK_TX_SEQ) & " (" & SysStr(SK_TX_TYPE) & ") отменена по снимку")
        ClearMarker()
    ElseIf st = TX_COMMITTED Or st = TX_NONE Then
        ' a completion interrupted after COMMITTED leaves the before-image (and TX_SEQ) behind: tidy up
        If SysStr(SK_TX_BI) <> "" Or SysNum(SK_TX_SEQ) <> SysNum(SK_LAST_SEQ) Then
            ClearMarker()
            AddNote("маркер операции приведён в порядок: завершение предыдущей операции было прервано")
        End If
    Else
        SetBlocked("MARKER_CORRUPT", "неизвестное состояние маркера операции «" & st & "»")
        Exit Sub
    End If
    ' 7. journal against the book
    lastSeq = CLng(SysNum(SK_LAST_SEQ))
    WmsJournal.ReadJournal(lastSeq, False, True)
    If gJR_Error <> "" Then
        SetBlocked("JOURNAL_UNREADABLE", gJR_Error)
        Exit Sub
    End If
    If gJR_Note <> "" Then AddNote(gJR_Note)
    If gJR_Foreign > 0 Then
        SetBlocked("JOURNAL_FOREIGN", "в журнале " & gJR_Foreign & " строк другой WMS (другой INSTANCE_ID) — журналы смешаны")
        Exit Sub
    End If
    If gJR_Order <> "" Then
        SetBlocked("JOURNAL_GAP", "нарушена непрерывность журнала: " & gJR_Order)
        Exit Sub
    End If
    p = WmsJournal.StartProblem(lastSeq)
    If p <> "" Then
        SetBlocked("JOURNAL_GAP", p)
        Exit Sub
    End If
    If lastSeq > 0 And Not gJR_HasLast Then
        SetBlocked("JOURNAL_BEHIND", "в книге последняя операция seq " & lastSeq & ", а в журнале " _
            & IIf(gJR_MaxSeq > 0, "последняя запись seq " & gJR_MaxSeq, "записей нет") _
            & " — журнал неполный или от другой копии. Если WMS перенесена без журнала — «Сделать рабочим файлом» с новым журналом")
        Exit Sub
    End If
    If gJR_Damaged > 0 Then AddNote("повреждённых строк журнала: " & gJR_Damaged & " (непрерывность seq не нарушена)")
    If gJR_Torn Then AddNote("последняя строка журнала оборвана сбоем записи (эта операция не была зафиксирована)")
    If gJR_N > 0 Then
        SetBlocked("TAIL", "в журнале " & gJR_N & " операций, которых нет в книге (seq " & gJR_Seq(0) & ".." & gJR_Seq(gJR_N - 1) _
            & "): выполните «Восстановить» или «Отложить хвост журнала»")
        Exit Sub
    End If
    ' 8. saved without WMS (macros off, a Basic compile error, another program) — spec §30
    ec = gDoc.getDocumentProperties().EditingCycles
    If SysNum(SK_SAVE_STAMP) <> ec Then
        AddNote("книга сохранялась без WMS (макросы были отключены?)")
        p = WmsDiagnostics.FullKeyCheck(True)
        AddNote(p)
        If Left(p, 6) = "ОШИБКА" Then
            SetBlocked("SAVED_WITHOUT_WMS", p)
            Exit Sub
        End If
        SysPutNum(SK_SAVE_STAMP, ec)
    End If
    ' written only when it changes: an unchanged book must not look modified after a start (LibreOffice would ask to save)
    If SysStr(SK_CORE_VERSION) <> WMS_CORE_VERSION Then SysPutStr(SK_CORE_VERSION, WMS_CORE_VERSION)
    gState = "CLEAN"
End Sub

' the tail must be re-read and re-checked by every action; "" when it can be processed
Private Function TailProblem(lastSeq As Long) As String
    WmsJournal.ReadJournal(lastSeq, False, True)
    If gJR_Error <> "" Then
        TailProblem = gJR_Error
    ElseIf gJR_Foreign > 0 Then
        TailProblem = "в журнале строки другой WMS"
    ElseIf gJR_Order <> "" Then
        TailProblem = "нарушена непрерывность журнала: " & gJR_Order
    ElseIf lastSeq > 0 And Not gJR_HasLast Then
        TailProblem = "журнал не содержит последнюю операцию книги (seq " & lastSeq & ")"
    ElseIf gJR_N = 0 Then
        TailProblem = "в журнале нет операций после seq " & lastSeq
    Else
        TailProblem = ""
    End If
End Function

Private Function FieldValue(fields As Variant, sName As String) As String
    Dim i As Long
    FieldValue = ""
    For i = 0 To UBound(fields)
        If Left(fields(i), Len(sName) + 1) = sName & "=" Then
            FieldValue = Unesc(Mid(fields(i), Len(sName) + 2))
            Exit Function
        End If
    Next i
End Function

Private Function IsAbandoned(seq As Long) As Boolean
    Dim i As Long
    For i = 0 To gJR_N - 1
        If gJR_Type(i) = "ABANDON" Then
            If seq >= Val(FieldValue(gJR_Fields(i), "FROM")) And seq <= Val(FieldValue(gJR_Fields(i), "TO")) Then
                IsAbandoned = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ColLetter(c As Integer) As String
    If c < 26 Then
        ColLetter = Chr(65 + c)
    Else
        ColLetter = Chr(64 + c \ 26) & Chr(65 + c Mod 26)
    End If
End Function

' APPLY — the book is in the state before the operation (inputs may be missing and are restored);
' ALREADY — every write is already in the book; otherwise CONFLICT:<details>
Private Function ClassifyReplay() As String
    Dim i As Long, cur As String, nApply As Long, nAlready As Long, oCell As Object
    For i = 0 To gPlanN - 1
        ' L (cell protection) and D (derived display values, refreshed by WMS outside operations too) are not evidence
        If gPW_Kind(i) <> "L" And gPW_Kind(i) <> "D" Then
            oCell = SheetByName(gPW_Sheet(i)).getCellByPosition(CInt(gPW_Col(i)), gPW_Row(i))
            cur = EncCell(oCell)
            If gPW_Kind(i) = "V" And SameEnc(gPW_Before(i), gPW_After(i)) Then
                ' a write that leaves the cell as it was tells nothing about whether the operation ran (e.g. a preview
                ' already showing the registry value); only a third value is a conflict
                If Not (SameEnc(cur, gPW_After(i)) Or (gPW_Restorable(i) And cur = "E")) Then
                    ClassifyReplay = "CONFLICT:" & gPW_Sheet(i) & "!" & ColLetter(CInt(gPW_Col(i))) & (gPW_Row(i) + 1) _
                        & " — в книге «" & Left(cur, 40) & "», ожидалось «" & Left(gPW_After(i), 40) & "»"
                    Exit Function
                End If
            ElseIf gPW_Kind(i) = "V" Then
                If SameEnc(cur, gPW_After(i)) Then
                    nAlready = nAlready + 1
                ElseIf SameEnc(cur, gPW_Before(i)) Or (gPW_Restorable(i) And cur = "E") Then
                    nApply = nApply + 1
                Else
                    ClassifyReplay = "CONFLICT:" & gPW_Sheet(i) & "!" & ColLetter(CInt(gPW_Col(i))) & (gPW_Row(i) + 1) _
                        & " — в книге «" & Left(cur, 40) & "», ожидалось «" & Left(gPW_Before(i), 40) & "» или «" & Left(gPW_After(i), 40) & "»"
                    Exit Function
                End If
            ElseIf gPW_Kind(i) = "I" Then
                If cur = "E" Then
                    nApply = nApply + 1
                ElseIf Not SameEnc(cur, gPW_After(i)) Then
                    ClassifyReplay = "CONFLICT:" & gPW_Sheet(i) & "!" & ColLetter(CInt(gPW_Col(i))) & (gPW_Row(i) + 1) _
                        & " — введено «" & Left(cur, 40) & "», а в журнале «" & Left(gPW_After(i), 40) & "»"
                    Exit Function
                End If
            End If
        End If
    Next i
    If nApply > 0 And nAlready > 0 Then
        ClassifyReplay = "CONFLICT:операция присутствует в книге частично"
    ElseIf nApply = 0 Then
        ClassifyReplay = "ALREADY"
    Else
        ClassifyReplay = "APPLY"
    End If
End Function

' «Восстановить»: replays the journal tail strictly in seq order through ApplyOperation; stops at the first problem
Function Recover() As String
    Dim lastSeq As Long, i As Long, nOk As Long, nSkip As Long, nAlready As Long, seq As Long, p As String, r As String, cls As String
    WmsInit()
    If gBlock <> "TAIL" Then
        Recover = "«Восстановить» не требуется: " & StateLine()
        Exit Function
    End If
    lastSeq = CLng(SysNum(SK_LAST_SEQ))
    p = TailProblem(lastSeq)
    If p <> "" Then
        Recover = "восстановление невозможно: " & p
        WmsStartup()
        Exit Function
    End If
    gState = "RECOVERING"
    For i = 0 To gJR_N - 1
        seq = gJR_Seq(i)
        If seq <> CLng(SysNum(SK_LAST_SEQ)) + 1 Then
            r = "ОСТАНОВКА: ожидалась seq " & (CLng(SysNum(SK_LAST_SEQ)) + 1) & ", в журнале " & seq
            GoTo DONE
        End If
        If gJR_Type(i) = "ABANDON" Or gJR_Type(i) = "START" Or IsAbandoned(seq) Then
            If gJR_Type(i) = "ABANDON" Then ApplyAbandonCounters(gJR_Fields(i))
            CommitSeq(seq)
            nSkip = nSkip + 1
        Else
            p = PlanFromJournalFields(gJR_Fields(i), gJR_Type(i))
            If p <> "" Then
                r = "ОСТАНОВКА на seq " & seq & ": " & p
                GoTo DONE
            End If
            cls = ClassifyReplay()
            If cls = "ALREADY" Then
                CommitSeq(seq)
                nAlready = nAlready + 1
            ElseIf cls = "APPLY" Then
                r = ApplyOperation(1, seq)
                If Left(r, 3) <> "OK:" Then
                    r = "ОСТАНОВКА на seq " & seq & ": " & r
                    GoTo DONE
                End If
                nOk = nOk + 1
            Else
                r = "ОСТАНОВКА на seq " & seq & ": " & Mid(cls, 10) & " — книга отличается от состояния, в котором операция была проведена"
                GoTo DONE
            End If
        End If
    Next i
    r = ""
    ' the whole tail is in the book: the saved journal position follows, so the next start reads only what is new
    WmsJournal.PinPosition(CLng(SysNum(SK_LAST_SEQ)))
DONE:
    gState = ""
    WmsStartup()
    Recover = "восстановлено операций: " & nOk & IIf(nAlready > 0, ", уже были в книге: " & nAlready, "") _
        & IIf(nSkip > 0, ", пропущено отложенных/служебных: " & nSkip, "") & IIf(r <> "", "; " & r, "")
End Function

' «Отложить хвост журнала»: the tail is not applied; an ABANDON record (append-only) closes it and the book continues
' after it. The abandoned operations stay in the journal for manual reconciliation.
Function AbandonTail() As String
    Dim lastSeq As Long, fromSeq As Long, toSeq As Long, p As String, sLine As String, why As String, rc As Integer
    Dim i As Long, lst As String, maxEI As Double, maxNo As Double, nextEI As Double, nextNo As Double
    WmsInit()
    If gBlock <> "TAIL" Then
        AbandonTail = "«Отложить хвост журнала» не требуется: " & StateLine()
        Exit Function
    End If
    lastSeq = CLng(SysNum(SK_LAST_SEQ))
    p = TailProblem(lastSeq)
    If p <> "" Then
        AbandonTail = "нельзя отложить хвост: " & p
        WmsStartup()
        Exit Function
    End If
    fromSeq = lastSeq + 1
    toSeq = gJR_Seq(gJR_N - 1)
    For i = 0 To gJR_N - 1
        If i < 20 Then lst = lst & IIf(lst <> "", ", ", "") & gJR_Seq(i) & ":" & gJR_Type(i)
        TailCounters(gJR_Fields(i), maxEI, maxNo)
    Next i
    ' numbers the abandoned operations handed out (an EI label, an issue №) are never issued again
    nextEI = SysNum(SK_NEXT_EI)
    If maxEI + 1 > nextEI Then nextEI = maxEI + 1
    nextNo = SysNum(SK_NEXT_NO)
    If maxNo + 1 > nextNo Then nextNo = maxNo + 1
    sLine = WmsJournal.BuildLine(toSeq + 1, "ABANDON", Array("FROM=" & fromSeq, "TO=" & toSeq, "NEXT_EI=" & NumStr(nextEI), _
        "NEXT_NO=" & NumStr(nextNo), "REASON=" & Esc("решение пользователя: хвост журнала не применён")))
    rc = WmsJournal.AppendLine(sLine, toSeq + 1, why)
    If rc <> 1 Then
        AbandonTail = "ОШИБКА: запись в журнал не удалась: " & why
        gState = ""
        WmsStartup()
        Exit Function
    End If
    ApplyAbandonCounters(Array("NEXT_EI=" & NumStr(nextEI), "NEXT_NO=" & NumStr(nextNo)))
    CommitSeq(toSeq + 1)
    gState = ""
    WmsStartup()
    AbandonTail = "отложено операций: " & (toSeq - fromSeq + 1) & " (seq " & fromSeq & ".." & toSeq & "), они остаются в журнале для ручной сверки: " & lst
End Function

' largest EI and issue № named by the fields of one journal entry (0 when none)
Private Sub TailCounters(fields As Variant, ByRef maxEI As Double, ByRef maxNo As Double)
    Dim v As String, n As Long, canon As String, msg As String
    v = FieldValue(fields, "EI")
    If v <> "" Then
        If WmsIssue.NormalizeEI(v, n, canon, msg) Then
            If n > maxEI Then maxEI = n
        End If
    End If
    v = FieldValue(fields, "NO")
    If IsDigits(v) And Len(v) <= 9 Then
        If CDbl(v) > maxNo Then maxNo = CDbl(v)
    End If
End Sub

' an ABANDON record raises NEXT_EI / NEXT_NO past the numbers of the abandoned operations (never lowers them); repeated
' application is harmless, so a crash between the journal line and the book is covered by «Восстановить»
Private Sub ApplyAbandonCounters(fields As Variant)
    Dim v As String
    v = FieldValue(fields, "NEXT_EI")
    If IsDigits(v) And Len(v) <= 9 Then
        If CDbl(v) > SysNum(SK_NEXT_EI) Then SysPutNum(SK_NEXT_EI, CDbl(v))
    End If
    v = FieldValue(fields, "NEXT_NO")
    If IsDigits(v) And Len(v) <= 9 Then
        If CDbl(v) > SysNum(SK_NEXT_NO) Then SysPutNum(SK_NEXT_NO, CDbl(v))
    End If
End Sub

' «Сделать рабочим файлом»: this file becomes the working WMS (after a deliberate move or restore).
' bStartNewJournal: only when the folder has no journal at all — a START record continues the numbering.
Function RegisterHere(bStartNewJournal As Boolean) As String
    Dim p As String, s As String, lastSeq As Long, sLine As String, why As String, rc As Integer
    WmsInit()
    If Not gDoc.hasLocation() Then
        RegisterHere = "нельзя: книга не сохранена в файл"
        Exit Function
    End If
    If gDoc.isReadonly() Or Not IsOdsDocument() Then
        RegisterHere = "нельзя: книга только для чтения или не в формате ODS"
        Exit Function
    End If
    p = SysLayoutProblem()
    If p <> "" Then
        RegisterHere = "нельзя: " & p
        Exit Function
    End If
    SysPutStr(SK_REG_URL, gDoc.getURL())
    s = "рабочий файл WMS: " & ConvertFromURL(gDoc.getURL())
    If bStartNewJournal Then
        lastSeq = CLng(SysNum(SK_LAST_SEQ))
        If UBound(WmsJournal.ListFiles()) >= 0 Then
            s = s & "; новый журнал НЕ начат: в папке уже есть журнал"
        Else
            p = WmsLock.Acquire()
            If p <> "" Then
                RegisterHere = s & "; новый журнал не начат: " & p
                gState = ""
                WmsStartup()
                Exit Function
            End If
            gJFile = WmsJournal.JournalFileName(Now())
            gJSize = 0
            sLine = WmsJournal.BuildLine(lastSeq + 1, "START", Array("BASE=" & lastSeq, "REASON=" & Esc("новый журнал после переноса книги")))
            rc = WmsJournal.AppendLine(sLine, lastSeq + 1, why)
            If rc <> 1 Then
                RegisterHere = s & "; ОШИБКА записи журнала: " & why
                gState = ""
                WmsStartup()
                Exit Function
            End If
            CommitSeq(lastSeq + 1)
            s = s & "; начат новый журнал с seq " & (lastSeq + 1)
        End If
    End If
    gState = ""
    WmsStartup()
    RegisterHere = s
End Function
