' WmsCore — shared state, _SYS access, document events, posting guard, the single write path ApplyOperation
' (two-phase, printable before-image, journal as commit point), Undo wrapper, locale-independent quantity parser.
' MASTER SPEC v0.3: §12 quantity, §17 Undo, §18 two-phase, §19 before-image, §21 fail-closed.
Option Explicit

Global gInit As Boolean
Global gDoc As Object
Global gSysSh As Object
Global gSfa As Object
Global gBaseDir As String
Global gJDir As String
Global gBDir As String
Global gSession As String
' "" = startup not run yet, CLEAN = posting allowed, BLOCKED, RECOVERING (replay in progress)
Global gState As String
Global gBlock As String
Global gBlockText As String
Global gReport As String
Global gBusy As Boolean
' journal ownership: the file this session writes to and the size it expects (spec §20)
Global gJFile As String
Global gJSize As Double
Global gLockOwned As Boolean
' test seam: honoured only when _SYS MODE = TEST (see FaultPoint, FaultCut)
Global gFaultPoint As Integer
Global gFaultMode As Integer
Global gHardStop As Boolean
' operation plan under construction (one at a time; Basic runs macros sequentially)
Global gPlanType As String
Global gPlanKey As String
Global gPlanFieldN As Long
Global gPlanFields() As String
Global gPlanN As Long
Global gPW_Kind() As String
Global gPW_Sheet() As String
Global gPW_Row() As Long
Global gPW_Col() As String
Global gPW_Before() As String
Global gPW_After() As String
Global gPW_Restorable() As Boolean

Private mUndoLocked As Boolean

' ================================================================ init and utilities

Sub WmsInit()
    If gInit Then Exit Sub
    Randomize
    gDoc = ThisComponent
    gSfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    gSession = NewId(12)
    gBaseDir = ""
    gJDir = ""
    gBDir = ""
    If gDoc.hasLocation() Then
        gBaseDir = FolderOf(gDoc.getURL())
        gJDir = gBaseDir & JOURNAL_DIR & "/"
        gBDir = gBaseDir & BACKUP_DIR & "/"
    End If
    If gDoc.Sheets.hasByName(SYS_SHEET) Then gSysSh = gDoc.Sheets.getByName(SYS_SHEET)
    gState = ""
    gBlock = ""
    gBlockText = ""
    gReport = ""
    gJFile = ""
    gJSize = -1
    gInit = True
End Sub

Function NewId(n As Integer) As String
    Dim i As Integer, s As String
    For i = 1 To n
        s = s & Mid("0123456789ABCDEF", Int(Rnd() * 16) + 1, 1)
    Next i
    NewId = s
End Function

Function TS() As String
    Dim d As Date
    d = Now()
    TS = Format(d, "YYYY-MM-DD") & "T" & Format(d, "HH:MM:SS")
End Function

' locale-independent number text ("1.5", never "1,5")
Function NumStr(x As Double) As String
    NumStr = Trim(Str(x))
End Function

Function IsDigits(s As String) As Boolean
    Dim i As Long, c As String
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    IsDigits = True
End Function

Function FolderOf(u As String) As String
    Dim i As Long
    For i = Len(u) To 1 Step -1
        If Mid(u, i, 1) = "/" Then
            FolderOf = Left(u, i)
            Exit Function
        End If
    Next i
    FolderOf = ""
End Function

Function FileNameOf(u As String) As String
    FileNameOf = Mid(u, Len(FolderOf(u)) + 1)
End Function

' Printable escaping for service data (journal lines, before-image): % ; | = ~ and control characters become %XX.
Function Esc(s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = Asc(ch)
        If code < 32 Or code = 127 Or ch = "%" Or ch = ";" Or ch = "|" Or ch = "=" Or ch = "~" Then
            out = out & "%" & Right("0" & Hex(code), 2)
        Else
            out = out & ch
        End If
    Next i
    Esc = out
End Function

Function Unesc(s As String) As String
    Dim i As Long, ch As String, out As String
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch = "%" And i + 2 <= Len(s) Then
            out = out & Chr(CLng("&H" & Mid(s, i + 1, 2)))
            i = i + 3
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    Unesc = out
End Function

Function ReadTextFile(url As String) As String
    Dim tis As Object
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(gSfaOrNew().openFileRead(url))
    tis.setEncoding("UTF-8")
    ReadTextFile = tis.readString(Array(), False)
    tis.closeInput()
End Function

Sub WriteTextFile(url As String, s As String)
    Dim tos As Object, sfa As Object
    sfa = gSfaOrNew()
    If sfa.exists(url) Then sfa.kill(url)
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(sfa.openFileWrite(url))
    tos.setEncoding("UTF-8")
    tos.writeString(s)
    tos.closeOutput()
End Sub

Private Function gSfaOrNew() As Object
    If IsNull(gSfa) Or IsEmpty(gSfa) Then gSfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    gSfaOrNew = gSfa
End Function

Function DocFilterName() As String
    Dim a As Variant, i As Integer
    a = gDoc.getArgs()
    For i = 0 To UBound(a)
        If a(i).Name = "FilterName" Then DocFilterName = a(i).Value
    Next i
End Function

' the working book must be an .ods file loaded by the ODS filter (spec §29): other formats lose macros or protection
Function IsOdsDocument() As Boolean
    IsOdsDocument = (DocFilterName() = "calc8") And (LCase(Right(gDoc.getURL(), 4)) = ".ods")
End Function

Function SheetByName(sName As String) As Object
    SheetByName = gDoc.Sheets.getByName(sName)
End Function

' ================================================================ _SYS access

Function SysStr(k As Integer) As String
    SysStr = gSysSh.getCellByPosition(1, k).getString()
End Function

' numeric value of a _SYS key; -1 when the cell does not hold a number
Function SysNum(k As Integer) As Double
    Dim c As Object
    c = gSysSh.getCellByPosition(1, k)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then
        SysNum = c.getValue()
    Else
        SysNum = -1
    End If
End Function

Sub SysPutStr(k As Integer, s As String)
    If s = "" Then
        ' only TX_BEFORE_IMAGE is emptied at run time; a leftover image under a COMMITTED marker is ignored at start and
        ' reported by the self-check, so a failed clear cannot corrupt data
        ClearCell(gSysSh.getCellByPosition(1, k))
    Else
        gSysSh.getCellByPosition(1, k).setString(s)
    End If
End Sub

' Empties a cell and confirms it. setString("")/setFormula("") leave the old content, and clearContents is silently
' refused for a locked cell of a protected sheet — then the sheet is unprotected for the moment of clearing.
Function ClearCell(oCell As Object) As Boolean
    Dim sh As Object, wasProt As Boolean, flags As Long
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    On Error Resume Next
    oCell.clearContents(flags)
    On Error GoTo 0
    If oCell.getType() = com.sun.star.table.CellContentType.EMPTY Then
        ClearCell = True
        Exit Function
    End If
    sh = oCell.getSpreadsheet()
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    oCell.clearContents(flags)
    If wasProt Then sh.protect(PROTECT_PWD)
    ClearCell = (oCell.getType() = com.sun.star.table.CellContentType.EMPTY)
End Function

Sub SysPutNum(k As Integer, x As Double)
    gSysSh.getCellByPosition(1, k).setValue(x)
End Sub

' "" when _SYS has the expected layout, otherwise a description (spec: corrupt service data → fail-closed)
Function SysLayoutProblem() As String
    Dim names As Variant, i As Integer, k As Variant, x As Double
    If IsNull(gSysSh) Or IsEmpty(gSysSh) Then
        SysLayoutProblem = "служебный лист " & SYS_SHEET & " отсутствует"
        Exit Function
    End If
    names = SysKeyNames()
    For i = 0 To UBound(names)
        If gSysSh.getCellByPosition(0, i).getString() <> names(i) Then
            SysLayoutProblem = SYS_SHEET & ": строка " & (i + 1) & " — ожидался ключ " & names(i) & ", найдено «" & gSysSh.getCellByPosition(0, i).getString() & "»"
            Exit Function
        End If
    Next i
    If SysStr(SK_SCHEMA) <> WMS_SYS_SCHEMA Then
        SysLayoutProblem = SYS_SHEET & ": схема «" & SysStr(SK_SCHEMA) & "», ядро ожидает «" & WMS_SYS_SCHEMA & "»"
        Exit Function
    End If
    If Len(SysStr(SK_INSTANCE)) < 8 Then
        SysLayoutProblem = SYS_SHEET & ": пустой или неверный INSTANCE_ID"
        Exit Function
    End If
    For Each k In SysNumericKeys()
        x = SysNum(k)
        If x < 0 Or x <> Int(x) Then
            SysLayoutProblem = SYS_SHEET & ": " & names(k) & " должен быть целым числом ≥ 0, найдено «" & SysStr(k) & "»"
            Exit Function
        End If
    Next k
    If SysNum(SK_NEXT_EI) < 1 Or SysNum(SK_NEXT_NO) < 1 Or SysNum(SK_NEXT_RET) < 1 Or SysNum(SK_MAX_QTY) < 1 Then
        SysLayoutProblem = SYS_SHEET & ": счётчики NEXT_* и MAX_QTY должны быть ≥ 1"
        Exit Function
    End If
    SysLayoutProblem = ""
End Function

' ================================================================ state, report, blocking

Sub AddNote(s As String)
    If s = "" Then Exit Sub
    If gReport = "" Then gReport = s Else gReport = gReport & " | " & s
End Sub

Sub SetBlocked(code As String, txt As String)
    gState = "BLOCKED"
    If gBlock = "" Then
        gBlock = code
        gBlockText = txt
    End If
    AddNote("ЗАБЛОКИРОВАНО [" & code & "]: " & txt)
End Sub

Function StartupReport() As String
    WmsInit()
    If gState = "" Then WmsStartup()
    StartupReport = StateLine() & " | " & gReport
End Function

Function StateLine() As String
    If gState = "CLEAN" Then
        StateLine = "СОСТОЯНИЕ: работа разрешена"
    Else
        StateLine = "СОСТОЯНИЕ: " & gState & IIf(gBlock <> "", " [" & gBlock & "] " & gBlockText, "")
    End If
End Function

Function StateDump() As String
    Dim s As String, um As Object
    WmsInit()
    s = "STATE=" & gState & ";BLOCK=" & gBlock
    If Not (IsNull(gSysSh) Or IsEmpty(gSysSh)) Then
        s = s & ";LAST_SEQ=" & SysStr(SK_LAST_SEQ) & ";NEXT_NO=" & SysStr(SK_NEXT_NO) & ";TX=" & SysStr(SK_TX_STATE) & ":" & SysStr(SK_TX_SEQ) _
            & ";JPOS=" & SysStr(SK_JPOS) & ";REG=" & IIf(SysStr(SK_REG_URL) = gDoc.getURL(), "own", "other")
    End If
    um = gDoc.getUndoManager()
    s = s & ";UNDO_LOCKED=" & um.isLocked() & ";UNDO_CAN=" & um.isUndoPossible() & ";SESSION=" & gSession
    StateDump = s
End Function

' Every accounting write goes through this guard: startup must have completed cleanly, the WMS lock must still be ours.
Function PostingBlockReason() As String
    Dim why As String
    WmsInit()
    If gState = "" Then WmsStartup()
    If gState <> "CLEAN" Then
        PostingBlockReason = "BLOCKED:" & IIf(gBlockText <> "", gBlockText, "WMS не готова (" & gState & ")")
        Exit Function
    End If
    If gBusy Then
        PostingBlockReason = "BUSY:операция уже выполняется"
        Exit Function
    End If
    why = WmsLock.VerifyOwner()
    If why <> "" Then
        SetBlocked("LOCK_LOST", why)
        PostingBlockReason = "BLOCKED:" & why
        Exit Function
    End If
    PostingBlockReason = ""
End Function

' ================================================================ document events (bound by tools/build_ods.py)

Sub OnDocLoad(Optional oEvent As Variant)
    WmsStartup()
End Sub

Function WmsStartup() As String
    Dim t0 As Long
    WmsInit()
    t0 = GetSystemTicks()
    gState = ""
    gBlock = ""
    gBlockText = ""
    gReport = ""
    On Error GoTo EH
    ' the analysis writes to the book (registration, marker rollback, copy marks): none of it may be undone by Ctrl+Z
    UndoBegin()
    WmsRecovery.Analyze()
    UndoEnd()
    AddNote("запуск WMS " & (GetSystemTicks() - t0) & " мс")
    WmsStartup = StateLine() & " | " & gReport
    Exit Function
EH:
    UndoEnd()
    ' unexpected error anywhere in startup: posting stays forbidden (D-028)
    gState = "BLOCKED"
    If gBlock = "" Then
        gBlock = "STARTUP_ERROR"
        gBlockText = "ошибка запуска WMS: " & Error$ & " (строка " & Erl & ")"
    End If
    AddNote("ЗАБЛОКИРОВАНО [STARTUP_ERROR]: " & Error$)
    WmsStartup = StateLine() & " | " & gReport
End Function

' user closes the window: release our lock, give AutoInput back
Sub OnDocPrepareUnload(Optional oEvent As Variant)
    On Error Resume Next
    If Not gInit Then Exit Sub
    WmsLock.ReleaseIfOwner()
    WmsConfig.AutoInputRestore()
End Sub

' "Save": the saved file carries the revision it will have after this save; a mismatch at the next start means
' the book was saved without WMS (macros disabled, compile error, other program) — spec §30
Sub OnDocSave(Optional oEvent As Variant)
    On Error Resume Next
    WmsInit()
    If IsNull(gSysSh) Or IsEmpty(gSysSh) Then Exit Sub
    SysPutNum(SK_SAVE_STAMP, gDoc.getDocumentProperties().EditingCycles + 1)
    SysPutNum(SK_SAVE_SEQ, SysNum(SK_LAST_SEQ))
End Sub

' ================================================================ explicit user actions (spec §21)

Function ActionRecover() As String
    WmsInit()
    ActionRecover = WmsRecovery.Recover() & " || " & StartupReport()
End Function

Function ActionAbandonTail() As String
    WmsInit()
    ActionAbandonTail = WmsRecovery.AbandonTail() & " || " & StartupReport()
End Function

Function ActionForceUnlock() As String
    WmsInit()
    ActionForceUnlock = WmsLock.ForceRelease() & " || " & WmsStartup()
End Function

Function ActionRegisterHere(Optional bStartNewJournal As Variant) As String
    Dim b As Boolean
    WmsInit()
    If Not IsMissing(bStartNewJournal) Then b = bStartNewJournal
    ActionRegisterHere = WmsRecovery.RegisterHere(b) & " || " & StartupReport()
End Function

Function ActionBackupNow() As String
    WmsInit()
    If gState = "" Then WmsStartup()
    ActionBackupNow = WmsBackup.BackupNow("manual")
End Function

Function ActionSelfCheck() As String
    WmsInit()
    If gState = "" Then WmsStartup()
    ActionSelfCheck = WmsDiagnostics.SelfCheck()
End Function

' ================================================================ value encoding (printable, locale-independent)
' E = empty, N<number>, S<escaped text>, F<escaped formula>

Function EncCell(oCell As Object) As String
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        EncCell = "E"
    Case com.sun.star.table.CellContentType.VALUE
        EncCell = "N" & NumStr(oCell.getValue())
    Case com.sun.star.table.CellContentType.TEXT
        EncCell = "S" & Esc(oCell.getString())
    Case Else
        EncCell = "F" & Esc(oCell.getFormula())
    End Select
End Function

Function EncVar(v As Variant) As String
    Select Case VarType(v)
    Case 0, 1
        EncVar = "E"
    Case 2, 3, 4, 5, 6
        EncVar = "N" & NumStr(CDbl(v))
    Case Else
        If CStr(v) = "" Then EncVar = "E" Else EncVar = "S" & Esc(CStr(v))
    End Select
End Function

' False for an unknown encoding (never guessed)
Function DecToCell(oCell As Object, enc As String) As Boolean
    Select Case Left(enc, 1)
    Case "E"
        If Not ClearCell(oCell) Then Exit Function
    Case "N"
        oCell.setValue(Val(Mid(enc, 2)))
    Case "S"
        oCell.setString(Unesc(Mid(enc, 2)))
    Case "F"
        oCell.setFormula(Unesc(Mid(enc, 2)))
    Case Else
        Exit Function
    End Select
    DecToCell = True
End Function

Function SameEnc(a As String, b As String) As Boolean
    Dim x As Double, y As Double
    If Left(a, 1) = "N" And Left(b, 1) = "N" Then
        x = Val(Mid(a, 2))
        y = Val(Mid(b, 2))
        SameEnc = (Abs(x - y) <= 0.000000001 * IIf(Abs(x) > 1, Abs(x), 1))
    Else
        SameEnc = (a = b)
    End If
End Function

' ================================================================ plan builder
' An operation module (Phase 2+: issue, receipt, return; tests: WmsTestOps) validates and calculates without
' touching the document, describes every write with PlanSetValue/PlanInput/PlanLock, then calls ApplyOperation.
' Operations only write cells and cell locks; they never insert or delete rows (new rows are appended).

Sub PlanBegin(sType As String, sKey As String)
    gPlanType = sType
    gPlanKey = sKey
    gPlanFieldN = 0
    gPlanN = 0
    ReDim gPlanFields(15)
    ReDim gPW_Kind(31)
    ReDim gPW_Sheet(31)
    ReDim gPW_Row(31)
    ReDim gPW_Col(31)
    ReDim gPW_Before(31)
    ReDim gPW_After(31)
    ReDim gPW_Restorable(31)
End Sub

' logical field of the journal record (audit; e.g. EI, quantity, balance before/after)
Sub PlanField(sName As String, v As Variant)
    Dim sv As String
    ' block form on purpose: after a one-line "If ... Then ReDim ..." LibreOffice Basic treats the next line as part of the If
    If gPlanFieldN > UBound(gPlanFields) Then
        ReDim Preserve gPlanFields(2 * gPlanFieldN + 1)
    End If
    If VarType(v) >= 2 And VarType(v) <= 6 Then sv = NumStr(CDbl(v)) Else sv = CStr(v)
    gPlanFields(gPlanFieldN) = sName & "=" & Esc(sv)
    gPlanFieldN = gPlanFieldN + 1
End Sub

Private Sub PlanAdd(sKind As String, sSheet As String, r As Long, sCol As String, sBefore As String, sAfter As String, bRestorable As Boolean)
    Dim m As Long
    If gPlanN > UBound(gPW_Kind) Then
        m = 2 * gPlanN + 1
        ReDim Preserve gPW_Kind(m)
        ReDim Preserve gPW_Sheet(m)
        ReDim Preserve gPW_Row(m)
        ReDim Preserve gPW_Col(m)
        ReDim Preserve gPW_Before(m)
        ReDim Preserve gPW_After(m)
        ReDim Preserve gPW_Restorable(m)
    End If
    gPW_Kind(gPlanN) = sKind
    gPW_Sheet(gPlanN) = sSheet
    gPW_Row(gPlanN) = r
    gPW_Col(gPlanN) = sCol
    gPW_Before(gPlanN) = sBefore
    gPW_After(gPlanN) = sAfter
    gPW_Restorable(gPlanN) = bRestorable
    gPlanN = gPlanN + 1
End Sub

' value written by WMS; bRestorable = the old content was user input that may be missing after a crash (e.g. the
' quantity text replaced by its numeric value), so an empty cell is also an acceptable precondition on replay
Sub PlanSetValue(sSheet As String, r As Long, c As Integer, vAfter As Variant, bRestorable As Boolean)
    PlanAdd("V", sSheet, r, CStr(c), EncCell(SheetByName(sSheet).getCellByPosition(c, r)), EncVar(vAfter), bRestorable)
End Sub

' user input that the operation relies on; not modified now, restored from the journal if lost in a crash
Sub PlanInput(sSheet As String, r As Long, c As Integer)
    Dim e As String
    e = EncCell(SheetByName(sSheet).getCellByPosition(c, r))
    PlanAdd("I", sSheet, r, CStr(c), e, e, True)
End Sub

' cell protection of columns c1..c2 in row r
Sub PlanLock(sSheet As String, r As Long, c1 As Integer, c2 As Integer, bLocked As Boolean)
    PlanAdd("L", sSheet, r, c1 & ":" & c2, LockBits(sSheet, r, c1, c2), String(c2 - c1 + 1, IIf(bLocked, "1", "0")), False)
End Sub

Function PlanCount() As Long
    PlanCount = gPlanN
End Function

' journal fields: logical fields, then one W record per write (kind|sheet|row|col|before|after|restorable)
Function PlanJournalFields() As Variant
    Dim a() As String, i As Long, n As Long
    n = gPlanFieldN + gPlanN
    If n = 0 Then
        PlanJournalFields = Array()
        Exit Function
    End If
    ReDim a(n - 1)
    For i = 0 To gPlanFieldN - 1
        a(i) = gPlanFields(i)
    Next i
    For i = 0 To gPlanN - 1
        a(gPlanFieldN + i) = "W=" & gPW_Kind(i) & "|" & Esc(gPW_Sheet(i)) & "|" & gPW_Row(i) & "|" & gPW_Col(i) & "|" _
            & gPW_Before(i) & "|" & gPW_After(i) & "|" & IIf(gPW_Restorable(i), "1", "0")
    Next i
    PlanJournalFields = a
End Function

' rebuild a plan from the W records of a journal entry (replay)
Function PlanFromJournalFields(fields As Variant, sType As String) As String
    Dim i As Long, p As Variant, f As String
    PlanBegin(sType, "")
    For i = 0 To UBound(fields)
        f = fields(i)
        If Left(f, 2) = "W=" Then
            p = Split(Mid(f, 3), "|")
            If UBound(p) <> 6 Then
                PlanFromJournalFields = "повреждена запись изменения: " & Left(f, 60)
                Exit Function
            End If
            If Not IsDigits(CStr(p(2))) Then
                PlanFromJournalFields = "неверная строка в записи изменения: " & Left(f, 60)
                Exit Function
            End If
            PlanAdd(CStr(p(0)), Unesc(CStr(p(1))), CLng(p(2)), CStr(p(3)), CStr(p(4)), CStr(p(5)), (p(6) = "1"))
        End If
    Next i
    PlanFromJournalFields = ""
End Function

' ================================================================ cell locks

Function LockBits(sSheet As String, r As Long, c1 As Integer, c2 As Integer) As String
    Dim sh As Object, c As Integer, s As String
    sh = SheetByName(sSheet)
    For c = c1 To c2
        If sh.getCellByPosition(c, r).CellProtection.IsLocked Then s = s & "1" Else s = s & "0"
    Next c
    LockBits = s
End Function

' "" on success, otherwise the error (the sheet is protected again in every case)
Function ApplyLockBits(sSheet As String, r As Long, c1 As Integer, c2 As Integer, bits As String) As String
    Dim sh As Object, c As Integer, p As Object, wasProt As Boolean
    On Error GoTo EH
    sh = SheetByName(sSheet)
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    For c = c1 To c2
        p = sh.getCellByPosition(c, r).CellProtection
        p.IsLocked = (Mid(bits, c - c1 + 1, 1) = "1")
        sh.getCellByPosition(c, r).CellProtection = p
    Next c
    If wasProt Then sh.protect(PROTECT_PWD)
    ApplyLockBits = ""
    Exit Function
EH:
    ApplyLockBits = "не удалось изменить защиту ячеек «" & sSheet & "»: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
End Function

Private Sub ColRange(sCol As String, ByRef c1 As Integer, ByRef c2 As Integer)
    Dim p As Long
    p = InStr(sCol, ":")
    If p > 0 Then
        c1 = CInt(Left(sCol, p - 1))
        c2 = CInt(Mid(sCol, p + 1))
    Else
        c1 = CInt(sCol)
        c2 = c1
    End If
End Sub

' ================================================================ Undo wrapper (spec §17)
' The lock keeps WMS writes out of the Undo stack while WMS works, but Calc lifts it internally (sheet protection and
' cell attribute changes re-enable Undo). The guarantee is therefore the clearing at the end: after every operation,
' replay and startup analysis — on every path — nothing WMS wrote can be undone with Ctrl+Z.

Sub UndoBegin()
    Dim um As Object
    If mUndoLocked Then Exit Sub
    um = gDoc.getUndoManager()
    um.lock()
    mUndoLocked = True
End Sub

Sub UndoEnd()
    Dim um As Object
    On Error Resume Next
    um = gDoc.getUndoManager()
    If mUndoLocked Then
        um.unlock()
        mUndoLocked = False
    End If
    um.clear()
End Sub

' ================================================================ before-image (spec §19)
' "BI1" then records separated by "~": V|sheet|row|col|value  or  L|sheet|row|c1:c2|bits  — printable only.

Function BuildBeforeImage() As String
    Dim i As Long, s As String, c1 As Integer, c2 As Integer
    s = "BI1"
    For i = 0 To gPlanN - 1
        If gPW_Kind(i) = "L" Then
            ColRange(gPW_Col(i), c1, c2)
            s = s & "~L|" & Esc(gPW_Sheet(i)) & "|" & gPW_Row(i) & "|" & gPW_Col(i) & "|" & LockBits(gPW_Sheet(i), gPW_Row(i), c1, c2)
        Else
            s = s & "~V|" & Esc(gPW_Sheet(i)) & "|" & gPW_Row(i) & "|" & gPW_Col(i) & "|" _
                & EncCell(SheetByName(gPW_Sheet(i)).getCellByPosition(CInt(gPW_Col(i)), gPW_Row(i)))
        End If
    Next i
    BuildBeforeImage = s
End Function

' "" when the before-image can be applied (all sheets exist, records well formed), otherwise the problem
Function BeforeImageProblem(bi As String) As String
    Dim recs As Variant, i As Long, p As Variant
    If Left(bi, 3) <> "BI1" Then
        BeforeImageProblem = "неизвестный формат снимка «" & Left(bi, 12) & "»"
        Exit Function
    End If
    recs = Split(bi, "~")
    For i = 1 To UBound(recs)
        p = Split(recs(i), "|")
        If UBound(p) <> 4 Then
            BeforeImageProblem = "повреждена запись снимка №" & i
            Exit Function
        End If
        If (p(0) <> "V" And p(0) <> "L") Or Not IsDigits(CStr(p(2))) Then
            BeforeImageProblem = "повреждена запись снимка №" & i
            Exit Function
        End If
        If Not gDoc.Sheets.hasByName(Unesc(CStr(p(1)))) Then
            BeforeImageProblem = "снимок ссылается на отсутствующий лист «" & Unesc(CStr(p(1))) & "»"
            Exit Function
        End If
    Next i
    BeforeImageProblem = ""
End Function

' restore in reverse order; True when every record was restored
Function RestoreBeforeImage(bi As String) As Boolean
    Dim recs As Variant, i As Long, p As Variant, c1 As Integer, c2 As Integer, nErr As Long
    On Error GoTo EH
    recs = Split(bi, "~")
    For i = UBound(recs) To 1 Step -1
        p = Split(recs(i), "|")
        If p(0) = "L" Then
            ColRange(CStr(p(3)), c1, c2)
            If ApplyLockBits(Unesc(CStr(p(1))), CLng(p(2)), c1, c2, CStr(p(4))) <> "" Then nErr = nErr + 1
        Else
            If Not DecToCell(SheetByName(Unesc(CStr(p(1)))).getCellByPosition(CInt(p(3)), CLng(p(2))), CStr(p(4))) Then nErr = nErr + 1
        End If
NextRec:
    Next i
    RestoreBeforeImage = (nErr = 0)
    Exit Function
EH:
    nErr = nErr + 1
    Resume NextRec
End Function

' ================================================================ ApplyOperation — the single write path (spec §18)
' mode 0: new operation (seq = LAST_SEQ + 1, the journal line is the commit point)
' mode 1: replay of journal entry replaySeq during recovery (the entry already exists, nothing is appended)
' Before the journal line: any error → full rollback from the before-image. After it: the operation is finished.

Function ApplyOperation(mode As Integer, replaySeq As Long) As String
    Dim seq As Long, why As String, sLine As String, bi As String, em As String
    Dim jDone As Boolean, started As Boolean, rc As Integer
    WmsInit()
    If mode = 0 Then
        why = PostingBlockReason()
        If why <> "" Then
            ApplyOperation = why
            Exit Function
        End If
        seq = CLng(SysNum(SK_LAST_SEQ)) + 1
    Else
        If gState <> "RECOVERING" Then
            ApplyOperation = "BLOCKED:повтор операции разрешён только при восстановлении"
            Exit Function
        End If
        seq = replaySeq
    End If
    gBusy = True
    gHardStop = False
    UndoBegin()
    On Error GoTo RB
    ' phase 2.1: STARTED + before-image of everything the operation will change
    bi = BuildBeforeImage()
    SysPutStr(SK_TX_STATE, TX_STARTED)
    SysPutNum(SK_TX_SEQ, seq)
    SysPutStr(SK_TX_TYPE, gPlanType)
    SysPutStr(SK_TX_TIME, TS())
    SysPutStr(SK_TX_BI, bi)
    started = True
    FaultPoint(1)
    ' phase 2.2: mutations
    why = ApplyWrites(mode)
    If why <> "" Then GoTo WRITE_FAILED
    ' phase 2.3: commit point
    If mode = 0 Then
        sLine = WmsJournal.BuildLine(seq, gPlanType, PlanJournalFields())
        rc = WmsJournal.AppendLine(sLine, seq, why)
        If rc = 0 Then GoTo JOURNAL_REFUSED
        If rc = 2 Then GoTo JOURNAL_UNCERTAIN
    End If
    jDone = True
    FaultPoint(3)
    ' phase 2.4: COMMITTED + LAST_SEQ
    CommitSeq(seq)
    UndoEnd()
    gBusy = False
    ApplyOperation = "OK:" & seq
    Exit Function

WRITE_FAILED:
JOURNAL_REFUSED:
    ' a write failed or the journal did not take the record: the operation never happened
    If RestoreBeforeImage(bi) Then
        ClearMarker()
    Else
        SetBlocked("ROLLBACK_FAILED", "откат операции не удался — перезапустите WMS")
    End If
    UndoEnd()
    gBusy = False
    ApplyOperation = "ERR-RB:" & why
    Exit Function

JOURNAL_UNCERTAIN:
    ' cannot tell whether the line reached the disk: leave the marker for the next start to decide
    SetBlocked("JOURNAL_UNCERTAIN", "неизвестно, записана ли операция в журнал: " & why & " — перезапустите WMS")
    UndoEnd()
    gBusy = False
    ApplyOperation = "ERR-CRITICAL:" & why
    Exit Function

RB:
    em = Error$
    If gHardStop Then
        ' test seam: the book was stored mid-operation and the process is about to be killed
        gBusy = False
        ApplyOperation = "CRASH-SIM"
        Exit Function
    End If
    If jDone Then
        ApplyOperation = FinishAfterJournal(seq, em)
        Exit Function
    End If
    If started Then
        If RestoreBeforeImage(bi) Then
            ClearMarker()
        Else
            SetBlocked("ROLLBACK_FAILED", "откат операции не удался — перезапустите WMS")
        End If
    End If
    UndoEnd()
    gBusy = False
    ApplyOperation = "ERR-RB:" & em
End Function

' "" when every write was applied; runtime errors propagate to ApplyOperation's handler
Private Function ApplyWrites(mode As Integer) As String
    Dim i As Long, c1 As Integer, c2 As Integer, oCell As Object, half As Long, why As String
    half = gPlanN \ 2
    For i = 0 To gPlanN - 1
        Select Case gPW_Kind(i)
        Case "V"
            If Not DecToCell(SheetByName(gPW_Sheet(i)).getCellByPosition(CInt(gPW_Col(i)), gPW_Row(i)), gPW_After(i)) Then
                ApplyWrites = "неизвестная кодировка значения в плане операции"
                Exit Function
            End If
        Case "I"
            If mode = 1 Then
                oCell = SheetByName(gPW_Sheet(i)).getCellByPosition(CInt(gPW_Col(i)), gPW_Row(i))
                If oCell.getType() = com.sun.star.table.CellContentType.EMPTY Then
                    If Not DecToCell(oCell, gPW_After(i)) Then
                        ApplyWrites = "неизвестная кодировка значения в плане операции"
                        Exit Function
                    End If
                End If
            End If
        Case "L"
            ColRange(gPW_Col(i), c1, c2)
            why = ApplyLockBits(gPW_Sheet(i), gPW_Row(i), c1, c2, gPW_After(i))
            If why <> "" Then
                ApplyWrites = why
                Exit Function
            End If
        Case Else
            ApplyWrites = "неизвестный вид изменения «" & gPW_Kind(i) & "»"
            Exit Function
        End Select
        If i = half Then FaultPoint(2)
    Next i
    ApplyWrites = ""
End Function

Private Function FinishAfterJournal(seq As Long, em As String) As String
    On Error GoTo EH
    CommitSeq(seq)
    UndoEnd()
    gBusy = False
    FinishAfterJournal = "OK:" & seq & " (операция уже в журнале; завершена после сбоя: " & em & ")"
    Exit Function
EH:
    ' the next start rolls back by the marker and replays the journal entry
    SetBlocked("COMMIT_FAILED", "операция seq " & seq & " записана в журнал, но не завершена в книге — перезапустите WMS")
    UndoEnd()
    gBusy = False
    FinishAfterJournal = "ERR-CRITICAL:операция seq " & seq & " в журнале, завершение не удалось: " & em
End Function

' COMMITTED is written before LAST_SEQ, so a completion interrupted part-way (error, save, crash) never leaves LAST_SEQ
' ahead of a STARTED marker — the next start would roll the cells back by the before-image and still count the
' operation as applied. Every prefix of these writes is safe: STARTED → rolled back, then replayed from the journal;
' COMMITTED without LAST_SEQ → replay finds the operation already in the book; the rest is tidied up at the next start.
Sub CommitSeq(seq As Long)
    FaultPoint(6)
    SysPutStr(SK_TX_STATE, TX_COMMITTED)
    FaultPoint(7)
    SysPutNum(SK_LAST_SEQ, seq)
    FaultPoint(8)
    SysPutNum(SK_TX_SEQ, seq)
    SysPutStr(SK_TX_BI, "")
End Sub

Sub ClearMarker()
    If SysNum(SK_LAST_SEQ) > 0 Then SysPutStr(SK_TX_STATE, TX_COMMITTED) Else SysPutStr(SK_TX_STATE, TX_NONE)
    SysPutNum(SK_TX_SEQ, SysNum(SK_LAST_SEQ))
    SysPutStr(SK_TX_BI, "")
End Sub

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestSetFault(point As Integer, fmode As Integer) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSetFault = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gFaultPoint = point
    gFaultMode = fmode
    TestSetFault = "OK"
End Function

' points: 1 after STARTED, 2 in the middle of the writes, 3 after the journal append, 6/7/8 inside the completion
' (CommitSeq: before any write / after COMMITTED / after LAST_SEQ) — see FaultPoint;
' 4 and 5 inside the append: the line is written only in part (4: half, 5: all but its final LF), then the write fails.
' modes: 1 error, 2 store the book and stop (crash simulation), 3 error that stays armed (the retry fails as well)
Function FaultCut(n As Long) As Long
    FaultCut = 0
    If gFaultPoint <> 4 And gFaultPoint <> 5 Then Exit Function
    If SysStr(SK_MODE) <> "TEST" Then Exit Function
    If gFaultPoint = 4 Then
        FaultCut = n \ 2
    Else
        FaultCut = n - 1
    End If
    gFaultPoint = 0
End Function

Private Sub FaultPoint(n As Integer)
    Dim fm As Integer, z As Integer
    If gFaultPoint <> n Then Exit Sub
    If SysStr(SK_MODE) <> "TEST" Then Exit Sub
    fm = gFaultMode
    If fm <> 3 Then gFaultPoint = 0
    If fm = 2 Then
        ' crash simulation: the in-flight state reaches the disk, then the caller stops without cleanup
        gHardStop = True
        gDoc.store()
    End If
    z = 0
    z = 1 / z
End Sub

' ================================================================ locale-independent quantity parser (spec §12, D-030)
' Status: 0 ok, 1 empty, 2 invalid, 3 ambiguous grouping, 4 too many decimals, 5 not positive, 6 date-like
' Accepted: digits, at most one "," or "." and up to 3 significant decimals ("1,5", "1.5", "0,125", "10,75").
' Ambiguous: exactly 3 digits after the separator and 1–3 digits before it ("1,500", "12.345") — thousands grouping
' or decimals? Rejected; an unambiguous form is "1,5"/"1500" or a fourth trailing zero ("2,1250" = 2,125).

Function ParseQtyText(ByVal s As String, ByRef value As Double, ByRef msg As String) As Integer
    Dim t As String, i As Long, ch As String, ip As String, fp As String, nSep As Integer, code As Long
    value = 0
    t = s
    Do While Len(t) > 0
        code = Asc(Left(t, 1))
        If code = 32 Or code = 9 Or code = 160 Or code = 8239 Then t = Mid(t, 2) Else Exit Do
    Loop
    Do While Len(t) > 0
        code = Asc(Right(t, 1))
        If code = 32 Or code = 9 Or code = 160 Or code = 8239 Then t = Left(t, Len(t) - 1) Else Exit Do
    Loop
    If t = "" Then
        msg = "не указано количество"
        ParseQtyText = 1
        Exit Function
    End If
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        code = Asc(ch)
        If ch >= "0" And ch <= "9" Then
            If nSep = 0 Then ip = ip & ch Else fp = fp & ch
        ElseIf ch = "," Or ch = "." Then
            nSep = nSep + 1
            If nSep > 1 Then
                msg = "«" & t & "»: несколько разделителей — похоже на дату или разряды; введите число, например 1,5"
                ParseQtyText = 6
                Exit Function
            End If
        ElseIf ch = "-" Or code = 8722 Then
            msg = "«" & t & "»: отрицательное количество не допускается"
            ParseQtyText = 2
            Exit Function
        ElseIf ch = "e" Or ch = "E" Then
            msg = "«" & t & "»: экспоненциальная запись не допускается"
            ParseQtyText = 2
            Exit Function
        ElseIf ch = "/" Or code = 8260 Or (code >= 188 And code <= 190) Or (code >= 8528 And code <= 8542) Then
            msg = "«" & t & "»: дробь не допускается — введите десятичное число, например 0,5"
            ParseQtyText = 2
            Exit Function
        ElseIf code = 32 Or code = 160 Or code = 8239 Or ch = "'" Then
            msg = "«" & t & "»: пробелы и разделители разрядов не допускаются — введите 1000"
            ParseQtyText = 2
            Exit Function
        Else
            msg = "«" & t & "»: недопустимый символ «" & ch & "»"
            ParseQtyText = 2
            Exit Function
        End If
    Next i
    If ip = "" Then
        msg = "«" & t & "»: нет целой части — введите, например, 0,5"
        ParseQtyText = 2
        Exit Function
    End If
    If nSep = 1 And fp = "" Then
        msg = "«" & t & "»: после разделителя нет цифр"
        ParseQtyText = 2
        Exit Function
    End If
    If Len(ip) > 1 And Left(ip, 1) = "0" Then
        msg = "«" & t & "»: лишние нули в начале числа"
        ParseQtyText = 2
        Exit Function
    End If
    If Len(ip) > QTY_MAX_INT_DIGITS Then
        msg = "«" & t & "»: слишком большое число"
        ParseQtyText = 2
        Exit Function
    End If
    If Len(fp) = 3 And ip <> "0" And Len(ip) <= 3 Then
        msg = "«" & t & "»: неоднозначно — это " & ip & "," & StripZeros(fp) & " или " & ip & fp & "? Для целого введите " & ip & fp _
            & ", для дробного — " & ip & "," & StripZeros(fp) & IIf(Right(fp, 1) <> "0", " (или " & ip & "," & fp & "0)", "")
        ParseQtyText = 3
        Exit Function
    End If
    fp = StripZeros(fp)
    If Len(fp) > QTY_MAX_DECIMALS Then
        msg = "«" & t & "»: не более " & QTY_MAX_DECIMALS & " знаков после запятой"
        ParseQtyText = 4
        Exit Function
    End If
    If fp = "" Then value = Val(ip) Else value = Val(ip & "." & fp)
    If value <= 0 Then
        msg = "«" & t & "»: количество должно быть больше 0"
        ParseQtyText = 5
        Exit Function
    End If
    msg = ""
    ParseQtyText = 0
End Function

Private Function StripZeros(ByVal fp As String) As String
    Do While Len(fp) > 0 And Right(fp, 1) = "0"
        fp = Left(fp, Len(fp) - 1)
    Loop
    StripZeros = fp
End Function

' quantity from a cell: text → strict parser; a number is accepted only in a plain number format (a date, time,
' percent, fraction or scientific format means Calc reinterpreted the input); formulas are refused
Function ParseQtyCell(oCell As Object, ByRef value As Double, ByRef msg As String) As Integer
    Dim ft As Long, x As Double
    value = 0
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        msg = "не указано количество"
        ParseQtyCell = 1
    Case com.sun.star.table.CellContentType.TEXT
        ParseQtyCell = ParseQtyText(oCell.getString(), value, msg)
    Case com.sun.star.table.CellContentType.VALUE
        ft = gDoc.getNumberFormats().getByKey(oCell.NumberFormat).Type
        If (ft And (2 Or 4 Or 32 Or 64 Or 128 Or 1024)) <> 0 Then
            msg = "ячейка количества распознана Calc как дата/время/процент/дробь — введите число, например 1,5"
            ParseQtyCell = 6
            Exit Function
        End If
        x = oCell.getValue()
        If x <= 0 Then
            msg = "количество должно быть больше 0"
            ParseQtyCell = 5
            Exit Function
        End If
        If Abs(x * 1000 - Int(x * 1000 + 0.5)) > 0.000001 Then
            msg = "не более " & QTY_MAX_DECIMALS & " знаков после запятой"
            ParseQtyCell = 4
            Exit Function
        End If
        value = Int(x * 1000 + 0.5) / 1000
        msg = ""
        ParseQtyCell = 0
    Case Else
        msg = "формула в ячейке количества не допускается"
        ParseQtyCell = 2
    End Select
End Function
