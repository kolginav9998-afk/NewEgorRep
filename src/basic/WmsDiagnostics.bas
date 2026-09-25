' WmsDiagnostics — minimal self-check (a subset of «WMS-Доктор», spec §49) and the full key check that runs after the
' book was saved without WMS (spec §15, §24, §30; D-032, D-034): only the needed columns, in chunks of CHECK_CHUNK_ROWS.
Option Explicit

Private mKeyProblems As Long

Private Function Ln(sStatus As String, sItem As String, sDetail As String) As String
    Ln = sStatus & " " & sItem & IIf(sDetail <> "", ": " & sDetail, "") & Chr(10)
End Function

Function SelfCheck() As String
    Dim out As String, p As String, t0 As Long, lastSeq As Long, st As String, s As String, i As Long, bi As String
    Dim ec As Double, nFail As Long, nWarn As Long, lines As Variant
    WmsInit()
    t0 = GetSystemTicks()
    If IsOdsDocument() Then
        out = out & Ln("OK", "формат", "ODS")
    Else
        out = out & Ln("FAIL", "формат", "не ODS («" & DocFilterName() & "»)")
    End If
    p = SysLayoutProblem()
    If p <> "" Then
        out = out & Ln("FAIL", "_SYS", p)
        GoTo SUMMARY
    End If
    out = out & Ln("OK", "_SYS", "схема " & SysStr(SK_SCHEMA) & ", экземпляр " & SysStr(SK_INSTANCE) & ", режим " & SysStr(SK_MODE))
    lastSeq = CLng(SysNum(SK_LAST_SEQ))
    out = out & Ln("OK", "счётчики", "LAST_SEQ " & lastSeq & ", NEXT_EI " & SysStr(SK_NEXT_EI) & ", NEXT_NO " & SysStr(SK_NEXT_NO) & ", NEXT_RET " & SysStr(SK_NEXT_RET))
    st = SysStr(SK_TX_STATE)
    bi = SysStr(SK_TX_BI)
    If st = TX_STARTED Then
        out = out & Ln("FAIL", "маркер операции", "STARTED вне операции (seq " & SysStr(SK_TX_SEQ) & ") — перезапустите WMS, операция будет отменена по снимку")
    ElseIf st <> TX_COMMITTED And st <> TX_NONE Then
        out = out & Ln("FAIL", "маркер операции", "неизвестное состояние «" & st & "»")
    ElseIf bi <> "" Then
        out = out & Ln("WARN", "маркер операции", st & ", но снимок до изменения не очищен")
    Else
        out = out & Ln("OK", "маркер операции", st & " seq " & SysStr(SK_TX_SEQ))
    End If
    For i = 1 To Len(bi)
        If Asc(Mid(bi, i, 1)) < 32 Then
            out = out & Ln("FAIL", "снимок до изменения", "содержит непечатаемые символы")
            Exit For
        End If
    Next i
    If SysStr(SK_REG_URL) = gDoc.getURL() Then
        out = out & Ln("OK", "рабочий путь", ConvertFromURL(gDoc.getURL()))
    Else
        out = out & Ln("FAIL", "рабочий путь", "книга открыта не по рабочему пути (" & SysStr(SK_REG_URL) & ")")
    End If
    s = WmsLock.VerifyOwner()
    If s = "" Then
        out = out & Ln("OK", "блокировка WMS", "принадлежит этому сеансу " & gSession)
    Else
        out = out & Ln("FAIL", "блокировка WMS", s)
    End If
    ' journal: full read, without moving the ownership baseline (a foreign write must stay detectable)
    WmsJournal.ReadJournal(lastSeq, True, False)
    If gJR_Error <> "" Then
        out = out & Ln("FAIL", "журнал", gJR_Error)
    Else
        s = "файлов " & gJR_Files & ", записей " & gJR_Lines & ", seq " & gJR_FirstSeq & ".." & gJR_MaxSeq
        If gJR_Foreign > 0 Then
            out = out & Ln("FAIL", "журнал", s & "; строк другой WMS: " & gJR_Foreign)
        ElseIf gJR_Order <> "" Then
            out = out & Ln("FAIL", "журнал", s & "; " & gJR_Order)
        ElseIf WmsJournal.StartProblem(lastSeq) <> "" Then
            out = out & Ln("FAIL", "журнал", s & "; " & WmsJournal.StartProblem(lastSeq))
        ElseIf lastSeq > 0 And Not gJR_HasLast Then
            out = out & Ln("FAIL", "журнал", s & "; нет записи последней операции книги seq " & lastSeq)
        ElseIf gJR_N > 0 Then
            out = out & Ln("FAIL", "журнал", s & "; операций вне книги: " & gJR_N)
        ElseIf gJR_Damaged > 0 Or gJR_Torn Then
            out = out & Ln("WARN", "журнал", s & "; повреждённых строк " & gJR_Damaged & IIf(gJR_Torn, ", последняя строка оборвана", ""))
        Else
            out = out & Ln("OK", "журнал", s & ", непрерывен")
        End If
    End If
    If SysStr(SK_JPOS) = "" Then
        out = out & Ln(IIf(lastSeq = 0, "OK", "WARN"), "позиция журнала", "не задана")
    ElseIf WmsJournal.PosSeq() = lastSeq And WmsJournal.PosValid(gJDir & WmsJournal.PosFile(), WmsJournal.PosOffset(), lastSeq) Then
        out = out & Ln("OK", "позиция журнала", SysStr(SK_JPOS) & " подтверждена содержимым")
    Else
        out = out & Ln("WARN", "позиция журнала", SysStr(SK_JPOS) & " не подтверждается — при запуске журнал будет перечитан (файл позиции и последующие или целиком)")
    End If
    ec = gDoc.getDocumentProperties().EditingCycles
    If gState = "CLEAN" Then
        out = out & Ln("OK", "сеанс WMS", gSession & ", запуск завершён, проведение разрешено")
    Else
        out = out & Ln("FAIL", "сеанс WMS", StateLine())
    End If
    If SysNum(SK_SAVE_STAMP) = ec Or SysNum(SK_SAVE_STAMP) = ec + 1 Then
        out = out & Ln("OK", "метка сохранения", "книга сохранялась средствами WMS")
    Else
        out = out & Ln("WARN", "метка сохранения", "книга сохранялась без WMS")
    End If
    If WmsConfig.AutoInputGet() Then
        out = out & Ln("WARN", "автоввод Calc", "включён (WMS выключает его при запуске)")
    Else
        out = out & Ln("OK", "автоввод Calc", "выключен")
    End If
    s = FullKeyCheck(False)
    If Left(s, 6) = "ОШИБКА" Or InStr(s, "; ОШИБКА") > 0 Then
        out = out & Ln("FAIL", "ключи", s)
    ElseIf mKeyProblems > 0 Then
        out = out & Ln("WARN", "ключи", s)
    Else
        out = out & Ln("OK", "ключи", s)
    End If
SUMMARY:
    out = out & Ln("INFO", "время проверки", (GetSystemTicks() - t0) & " мс")
    lines = Split(out, Chr(10))
    For i = 0 To UBound(lines)
        If Left(lines(i), 4) = "FAIL" Then nFail = nFail + 1
        If Left(lines(i), 4) = "WARN" Then nWarn = nWarn + 1
    Next i
    SelfCheck = "САМОПРОВЕРКА WMS: ошибок " & nFail & ", предупреждений " & nWarn & Chr(10) & out
End Function

' ---------------------------------------------------------------- full key check
' _SYS KEY_SHEETS = "sheet|keyCol|ctlCol|counterRow|inputCol;..." (0-based columns; counterRow = _SYS row of NEXT_*).
' A key the WMS never issued (≥ NEXT_*), an invalid key or a repeated key marks the row «КОПИЯ» (never posted);
' of repeated keys the first row in sheet order is kept as the original.

Function FullKeyCheck(bMark As Boolean) As String
    Dim specs As Variant, i As Long, res As String
    WmsInit()
    mKeyProblems = 0
    If SysStr(SK_KEY_SHEETS) = "" Then
        FullKeyCheck = "проверка ключей: учётных листов с ключами пока нет"
        Exit Function
    End If
    specs = Split(SysStr(SK_KEY_SHEETS), ";")
    For i = 0 To UBound(specs)
        If specs(i) <> "" Then res = res & IIf(res <> "", "; ", "") & CheckKeySheet(CStr(specs(i)), bMark)
    Next i
    FullKeyCheck = res
End Function

Private Function LastUsedRow(sh As Object) As Long
    Dim cur As Object
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    LastUsedRow = cur.getRangeAddress().EndRow
End Function

Private Sub MarkCopy(sh As Object, ctlCol As Integer, r As Long, txt As String)
    sh.getCellByPosition(ctlCol, r).setString("КОПИЯ: " & txt & " — не проводится, удалите командой «Очистить копию»")
End Sub

Private Function CheckKeySheet(spec As String, bMark As Boolean) As String
    Dim a As Variant, sName As String, keyCol As Integer, ctlCol As Integer, cntRow As Integer, inCol As Integer
    Dim sh As Object, last As Long, r0 As Long, r1 As Long, i As Long, j As Long, nextKey As Double, k As Variant
    Dim dKey As Variant, dCtl As Variant, dIn As Variant, buf() As Variant, nb As Long, nK As Long
    Dim nNever As Long, nBad As Long, nUnposted As Long, nDup As Long, t0 As Long, scr As Object, wasProt As Boolean
    Dim prevKey As Double, prevRow As Long, d As Variant, desc As Variant, fld(1) As New com.sun.star.table.TableSortField
    On Error GoTo EH
    t0 = GetSystemTicks()
    a = Split(spec, "|")
    sName = Unesc(CStr(a(0)))
    keyCol = CInt(a(1))
    ctlCol = CInt(a(2))
    cntRow = CInt(a(3))
    inCol = CInt(a(4))
    sh = SheetByName(sName)
    last = LastUsedRow(sh)
    nextKey = SysNum(cntRow)
    wasProt = gSysSh.isProtected()
    If wasProt Then gSysSh.unprotect(PROTECT_PWD)
    gSysSh.getCellRangeByPosition(3, 0, 4, 1048575).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.STRING)
    ' pass 1: three columns per chunk; keys that may be genuine go to a scratch area for the duplicate check
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        dKey = sh.getCellRangeByPosition(keyCol, r0, keyCol, r1).getDataArray()
        dCtl = sh.getCellRangeByPosition(ctlCol, r0, ctlCol, r1).getDataArray()
        dIn = sh.getCellRangeByPosition(inCol, r0, inCol, r1).getDataArray()
        ReDim buf(r1 - r0)
        nb = 0
        For i = 0 To r1 - r0
            k = dKey(i)(0)
            If Left(CStr(dCtl(i)(0)), 5) = "КОПИЯ" Then
                ' already marked
            ElseIf VarType(k) = 5 Then
                If k < 1 Or k <> Int(k) Then
                    nBad = nBad + 1
                    If bMark Then MarkCopy(sh, ctlCol, r0 + i, "неверный ключ " & k)
                ElseIf k >= nextKey Then
                    nNever = nNever + 1
                    If bMark Then MarkCopy(sh, ctlCol, r0 + i, "ключ " & k & " не выдавался WMS")
                Else
                    buf(nb) = Array(CDbl(k), CDbl(r0 + i))
                    nb = nb + 1
                End If
            ElseIf CStr(k) <> "" Then
                nBad = nBad + 1
                If bMark Then MarkCopy(sh, ctlCol, r0 + i, "неверный ключ «" & k & "»")
            ElseIf CStr(dIn(i)(0)) <> "" Then
                nUnposted = nUnposted + 1
            End If
        Next i
        If nb > 0 Then
            ReDim Preserve buf(nb - 1)
            gSysSh.getCellRangeByPosition(3, nK, 4, nK + nb - 1).setDataArray(buf)
            nK = nK + nb
        End If
    Next r0
    ' pass 2: repeated keys — sorted by the Calc engine (key, then row), then compared pairwise in chunks
    If nK > 1 Then
        scr = gSysSh.getCellRangeByPosition(3, 0, 4, nK - 1)
        fld(0).Field = 0
        fld(0).IsAscending = True
        fld(1).Field = 1
        fld(1).IsAscending = True
        desc = scr.createSortDescriptor()
        For i = 0 To UBound(desc)
            If desc(i).Name = "SortFields" Then desc(i).Value = fld()
            If desc(i).Name = "ContainsHeader" Then desc(i).Value = False
        Next i
        scr.sort(desc)
        prevKey = -1
        For r0 = 0 To nK - 1 Step CHECK_CHUNK_ROWS
            r1 = r0 + CHECK_CHUNK_ROWS - 1
            If r1 > nK - 1 Then r1 = nK - 1
            d = gSysSh.getCellRangeByPosition(3, r0, 4, r1).getDataArray()
            For j = 0 To r1 - r0
                If d(j)(0) = prevKey Then
                    nDup = nDup + 1
                    If bMark Then MarkCopy(sh, ctlCol, CLng(d(j)(1)), "ключ " & d(j)(0) & " повторяет строку " & (prevRow + 1))
                Else
                    prevKey = d(j)(0)
                    prevRow = CLng(d(j)(1))
                End If
            Next j
        Next r0
    End If
    gSysSh.getCellRangeByPosition(3, 0, 4, nK + 1).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.STRING)
    If wasProt Then gSysSh.protect(PROTECT_PWD)
    mKeyProblems = mKeyProblems + nNever + nDup + nBad
    CheckKeySheet = sName & ": строк " & last & ", ключей " & nK & ", не выдавались " & nNever & ", повторов " & nDup & ", неверных " & nBad _
        & ", непроведённых " & nUnposted & IIf(bMark And nNever + nDup + nBad > 0, " (помечены КОПИЯ)", "") & ", " & (GetSystemTicks() - t0) & " мс"
    Exit Function
EH:
    CheckKeySheet = "ОШИБКА проверки ключей «" & sName & "»: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    gSysSh.getCellRangeByPosition(3, 0, 4, 1048575).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.STRING)
    If wasProt Then gSysSh.protect(PROTECT_PWD)
End Function
