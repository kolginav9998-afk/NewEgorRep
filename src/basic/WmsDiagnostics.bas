' WmsDiagnostics — minimal self-check (a subset of «WMS-Доктор», spec §49) and the full key check that runs after the
' book was saved without WMS (spec §15, §24, §30; D-032, D-034): only the needed columns, in chunks of CHECK_CHUNK_ROWS.
Option Explicit

Private mKeyProblems As Long
' _RCV rows read in chunks while the sorted EI keys are walked (each chunk at most once): EI, OLID, row hint
Private mRcvChunk As Long
Private mRcvData As Variant
' the same for _RET while the sorted return № keys are walked: №, issue, EI, quantity, state, row hint (only its used rows);
' Phase 5: the same for _SPR (the lines of «Иной приход») — the registry of the key type of the sheet being checked
Private mRetChunk As Long
Private mRetData As Variant
Private mRetLast As Long
Private mRegSheet As String
Private mRegLast As Integer
Private mRegRow As Integer

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
    out = out & Ln("OK", "счётчики", "LAST_SEQ " & lastSeq & ", NEXT_EI " & SysStr(SK_NEXT_EI) & ", NEXT_NO " & SysStr(SK_NEXT_NO) & ", NEXT_RET " & SysStr(SK_NEXT_RET) _
        & ", NEXT_SPL " & SysStr(SK_NEXT_SPL) & ", NEXT_OFF " & SysStr(SK_NEXT_OFF) & ", NEXT_PROD " & SysStr(SK_NEXT_PROD) & ", NEXT_DET " & SysStr(SK_NEXT_DET) _
        & ", NEXT_OLD " & SysStr(SK_NEXT_OLD) & ", NEXT_OTH " & SysStr(SK_NEXT_OTH) & ", NEXT_ADJ " & SysStr(SK_NEXT_ADJ) & ", NEXT_CAR " & SysStr(SK_NEXT_CAR))
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
    s = RegistryCheck()
    If Left(s, 6) = "ОШИБКА" Then
        out = out & Ln("FAIL", "реестр «" & SH_STOCK & "»", Mid(s, 9))
    Else
        out = out & Ln("OK", "реестр «" & SH_STOCK & "»", s)
    End If
    s = ReceiptsCheck()
    If Left(s, 6) = "ОШИБКА" Then
        out = out & Ln("FAIL", "приходы и заказы", Mid(s, 9))
    Else
        out = out & Ln("OK", "приходы и заказы", s)
    End If
    s = WmsReturn.ReturnsCheck()
    If Left(s, 6) = "ОШИБКА" Then
        out = out & Ln("FAIL", "возвраты", Mid(s, 9))
    Else
        out = out & Ln("OK", "возвраты", s)
    End If
    s = WmsSpecial.SpecialCheck()
    If Left(s, 6) = "ОШИБКА" Then
        out = out & Ln("FAIL", "специальные приходы", Mid(s, 9))
    Else
        out = out & Ln("OK", "специальные приходы", s)
    End If
    s = WmsAdjust.AdjustCheck()
    If Left(s, 6) = "ОШИБКА" Then
        out = out & Ln("FAIL", "корректировки", Mid(s, 9))
    Else
        out = out & Ln("OK", "корректировки", s)
    End If
    s = WmsCar.CarCheck()
    If Left(s, 6) = "ОШИБКА" Then
        out = out & Ln("FAIL", "приход авто", Mid(s, 9))
    Else
        out = out & Ln("OK", "приход авто", s)
    End If
    s = WmsIssue.RecipientsDuplicates()
    If s <> "" Then
        out = out & Ln("WARN", "справочник «" & SH_RCPT & "»", "повторяются сокращения " & s & " — такие сокращения не подставляются")
    Else
        out = out & Ln("OK", "справочник «" & SH_RCPT & "»", "записей " & WmsIssue.RecipientsCount() & ", сокращения уникальны")
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

' ---------------------------------------------------------------- EI registry «Наличие» (spec §5, §49)
' Row n must hold ЕИ n (dense addressing) and a balance ≥ 0. Two columns, in chunks of CHECK_CHUNK_ROWS.
Function RegistryCheck() As String
    Dim sh As Object, cur As Object, last As Long, r0 As Long, r1 As Long, i As Long, dEI As Variant, dQ As Variant
    Dim nEI As Long, nGap As Long, nBad As Long, nNeg As Long, first As String, t0 As Long, canon As String
    On Error GoTo EH
    t0 = GetSystemTicks()
    sh = gDoc.Sheets.getByName(SH_STOCK)
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        dEI = sh.getCellRangeByPosition(SC_EI, r0, SC_EI, r1).getDataArray()
        dQ = sh.getCellRangeByPosition(SC_QTY, r0, SC_QTY, r1).getDataArray()
        For i = 0 To r1 - r0
            canon = WmsIssue.EiCanon(r0 + i)
            If CStr(dEI(i)(0)) = canon Then
                nEI = nEI + 1
                If VarType(dQ(i)(0)) <> 5 Then
                    If CStr(dQ(i)(0)) <> "" Then
                        nBad = nBad + 1
                        If first = "" Then first = "строка " & (r0 + i + 1) & ": остаток не число"
                    End If
                ElseIf dQ(i)(0) < 0 Then
                    nNeg = nNeg + 1
                    If first = "" Then first = "строка " & (r0 + i + 1) & ": отрицательный остаток " & dQ(i)(0)
                End If
            ElseIf CStr(dEI(i)(0)) = "" Then
                nGap = nGap + 1
            Else
                nBad = nBad + 1
                If first = "" Then first = "строка " & (r0 + i + 1) & ": «" & dEI(i)(0) & "» вместо " & canon
            End If
        Next i
    Next r0
    If nBad + nNeg > 0 Then
        RegistryCheck = "ОШИБКА: ЕИ " & nEI & ", нарушений адресации " & nBad & ", отрицательных остатков " & nNeg & " (" & first & ")"
    Else
        RegistryCheck = "ЕИ " & nEI & IIf(nGap > 0, ", пустых строк " & nGap, "") & ", адресация и остатки в порядке, " & (GetSystemTicks() - t0) & " мс"
    End If
    Exit Function
EH:
    RegistryCheck = "ОШИБКА: проверка реестра не выполнена: " & Error$
End Function

' ---------------------------------------------------------------- receipts and order positions (Phase 3)
' Every receipt of _RCV against «Наличие» (a live receipt has its registry row; a storno has balance 0 and the storno
' state), every position of _ORD against the sum of its live receipts, NEXT_EI above every EI of the book, the row hints.
' Chunks of CHECK_CHUNK_ROWS rows, only the needed columns (spec §24).
Function ReceiptsCheck() As String
    Dim rcv As Object, stk As Object, ords As Object, iss As Object, os As Object, nOl As Long, nextEI As Long, last As Long
    Dim r0 As Long, r1 As Long, i As Long, d As Variant, st As Variant, sumQ() As Double, cnt() As Long, nod() As Long, hint() As Long
    Dim ol As Long, nRcv As Long, nStorno As Long, nBad As Long, first As String, nPos As Long, nCancel As Long, nRest As Long
    Dim nStale As Long, maxEI As Long, n As Long, t0 As Long, canon As String, q As Variant, dv As Variant, where As String, v As Variant, x As Double
    Dim plus As Variant, px As Double
    On Error GoTo EH
    t0 = GetSystemTicks()
    rcv = gDoc.Sheets.getByName(SH_RCV)
    stk = gDoc.Sheets.getByName(SH_STOCK)
    ords = gDoc.Sheets.getByName(SH_ORD)
    os = gDoc.Sheets.getByName(SH_ORDERS)
    iss = gDoc.Sheets.getByName(SH_ISSUES)
    nOl = WmsOrders.NextOl()
    nextEI = CLng(SysNum(SK_NEXT_EI))
    ' a balance may exceed its receipt by the live positive inventory corrections of the EI (Final Core)
    plus = WmsAdjust.PlusByEI(nextEI)
    ReDim sumQ(nOl)
    ReDim cnt(nOl)
    ReDim nod(nOl)
    last = LastUsedRow(rcv)
    If last >= nextEI Then Bad(nBad, first, SH_RCV & ": приход с номером ЕИ ≥ NEXT_EI " & nextEI)
    ReDim hint(IIf(last > 0, last, 0))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = rcv.getCellRangeByPosition(0, r0, RV_LAST, r1).getDataArray()
        st = stk.getCellRangeByPosition(0, r0, SC_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            n = r0 + i
            hint(n) = -1
            If CStr(d(i)(RV_EI)) <> "" Then
                canon = EI_PREFIX & Right("0000000" & n, 8)
                If CStr(d(i)(RV_EI)) <> canon Or VarType(d(i)(RV_OL)) <> 5 Or VarType(d(i)(RV_QTY)) <> 5 Or CStr(st(i)(SC_EI)) <> canon _
                    Or CStr(d(i)(RV_STATE)) <> RV_LIVE Then where = SH_RCV & " строка " & (n + 1) & " (" & canon & ")"
                If CStr(d(i)(RV_EI)) <> canon Then
                    Bad(nBad, first, where & ": записан «" & d(i)(RV_EI) & "»")
                ElseIf VarType(d(i)(RV_OL)) <> 5 Or VarType(d(i)(RV_QTY)) <> 5 Or VarType(d(i)(RV_ROW)) <> 5 Then
                    Bad(nBad, first, where & ": повреждены OLID, строка или количество")
                ElseIf d(i)(RV_OL) < 1 Or d(i)(RV_OL) >= nOl Or d(i)(RV_QTY) <= 0 Then
                    Bad(nBad, first, where & ": OLID вне диапазона или количество ≤ 0")
                ElseIf CStr(st(i)(SC_EI)) <> canon Then
                    Bad(nBad, first, where & ": нет строки в «" & SH_STOCK & "»")
                Else
                    nRcv = nRcv + 1
                    hint(n) = CLng(d(i)(RV_ROW))
                    ol = CLng(d(i)(RV_OL))
                    q = st(i)(SC_QTY)
                    If CStr(d(i)(RV_STATE)) = RV_LIVE Then
                        sumQ(ol) = sumQ(ol) + d(i)(RV_QTY)
                        cnt(ol) = cnt(ol) + 1
                        If VarType(d(i)(RV_NODOC)) = 5 Then nod(ol) = nod(ol) + d(i)(RV_NODOC)
                        If VarType(q) <> 5 Then
                            Bad(nBad, first, SH_RCV & " строка " & (n + 1) & " (" & canon & "): остаток в реестре не число")
                        Else
                            px = 0
                            If n <= UBound(plus) Then px = plus(n)
                            If q > d(i)(RV_QTY) + px + 0.0000001 Then
                                Bad(nBad, first, SH_RCV & " строка " & (n + 1) & " (" & canon & "): остаток " & q & " больше прихода " & d(i)(RV_QTY) _
                                    & IIf(px > 0, " и излишков инвентаризации " & px, ""))
                            End If
                        End If
                    ElseIf CStr(d(i)(RV_STATE)) = RV_STORNO Then
                        nStorno = nStorno + 1
                        If VarType(q) <> 5 Then
                            Bad(nBad, first, where & ": остаток сторнированного прихода не число")
                        ElseIf q <> 0 Or CStr(st(i)(SC_STATE)) <> EI_ST_STORNO Then
                            Bad(nBad, first, where & ": приход удалён (сторно), а в реестре остаток " & q & ", состояние «" & st(i)(SC_STATE) & "»")
                        End If
                    Else
                        Bad(nBad, first, where & ": неизвестное состояние «" & d(i)(RV_STATE) & "»")
                    End If
                End If
            End If
        Next i
    Next r0
    ' positions: dense rows 1..NEXT_OL-1, totals equal to the sum of their live receipts
    last = LastUsedRow(ords)
    If last >= nOl Then Bad(nBad, first, SH_ORD & ": позиция с OLID ≥ NEXT_OL " & nOl)
    For r0 = 1 To nOl - 1 Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > nOl - 1 Then r1 = nOl - 1
        d = ords.getCellRangeByPosition(0, r0, OD_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            ol = r0 + i
            where = SH_ORD & " OLID "
            If VarType(d(i)(OD_ID)) <> 5 Then
                Bad(nBad, first, where & ol & ": строка пуста или повреждена")
            ElseIf d(i)(OD_ID) <> ol Or VarType(d(i)(OD_RCV)) <> 5 Or VarType(d(i)(OD_CNT)) <> 5 Or VarType(d(i)(OD_ORD)) <> 5 Then
                Bad(nBad, first, where & ol & ": повреждена")
            Else
                nPos = nPos + 1
                If Abs(d(i)(OD_RCV) - sumQ(ol)) > 0.0000001 Or d(i)(OD_CNT) <> cnt(ol) Or d(i)(OD_NODOC) <> nod(ol) Then
                    Bad(nBad, first, where & ol & ": получено " & d(i)(OD_RCV) & " / приходов " & d(i)(OD_CNT) & ", а по приходам " & sumQ(ol) & " / " & cnt(ol))
                End If
                If CStr(d(i)(OD_CANCEL)) = OD_CANCEL_ORDER Then
                    nCancel = nCancel + 1
                    If cnt(ol) > 0 Then Bad(nBad, first, where & ol & ": отменённая позиция имеет действующие приходы")
                ElseIf CStr(d(i)(OD_CANCEL)) = OD_CANCEL_REST Then
                    nRest = nRest + 1
                ElseIf CStr(d(i)(OD_CANCEL)) <> "" Then
                    Bad(nBad, first, where & ol & ": неизвестная отметка отмены «" & d(i)(OD_CANCEL) & "»")
                End If
            End If
        Next i
    Next r0
    ' «Заказы».V: every EI below NEXT_EI; outdated row hints (rows inserted above) are healed by the next operation
    last = LastUsedRow(os)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        dv = os.getCellRangeByPosition(OC_EI, r0, OC_EI, r1).getDataArray()
        For i = 0 To r1 - r0
            v = dv(i)(0)
            If Len(v) = 11 Then
                x = Val(Mid(v, 4))
                If x >= 1 And x <= UBound(hint) Then
                    If hint(x) >= 0 And hint(x) <> r0 + i Then nStale = nStale + 1
                End If
                If x > maxEI Then
                    If WmsOrders.StrictEI(CStr(v), n) Then maxEI = n
                End If
            End If
        Next i
    Next r0
    ' posted issues («Выдачи».L)
    last = LastUsedRow(iss)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        dv = iss.getCellRangeByPosition(IC_EI, r0, IC_EI, r1).getDataArray()
        d = iss.getCellRangeByPosition(IC_NO, r0, IC_NO, r1).getDataArray()
        For i = 0 To r1 - r0
            If VarType(d(i)(0)) = 5 Then
                v = dv(i)(0)
                If Len(v) = 11 Then
                    If Val(Mid(v, 4)) > maxEI Then
                        If WmsOrders.StrictEI(CStr(v), n) Then maxEI = n
                    End If
                End If
            End If
        Next i
    Next r0
    If LastUsedRow(stk) > maxEI Then maxEI = LastUsedRow(stk)
    If maxEI >= nextEI Then Bad(nBad, first, "NEXT_EI " & nextEI & " не больше максимального ЕИ книги " & maxEI)
    If nBad > 0 Then
        ReceiptsCheck = "ОШИБКА: расхождений " & nBad & " (" & first & ")"
    Else
        ReceiptsCheck = "приходов " & nRcv & " (из них сторно " & nStorno & "), позиций заказов " & nPos & " (отменено " & nCancel _
            & ", с отменённым остатком " & nRest & "), расхождений нет, NEXT_EI " & nextEI & " выше всех ЕИ книги" _
            & IIf(nStale > 0, ", устаревших подсказок строк " & nStale & " (исправятся следующей операцией)", "") & ", " & (GetSystemTicks() - t0) & " мс"
    End If
    Exit Function
EH:
    ReceiptsCheck = "ОШИБКА: проверка приходов не выполнена: " & Error$ & " (строка " & Erl & ")"
End Function

Private Sub Bad(ByRef nBad As Long, ByRef first As String, s As String)
    nBad = nBad + 1
    If first = "" Then first = s
End Sub

' ---------------------------------------------------------------- full key check
' _SYS KEY_SHEETS = "sheet|keyCol|ctlCol|counterRow|inputCol[|EI|RET];..." (0-based columns; counterRow = _SYS row of
' NEXT_*). A key the WMS never issued (≥ NEXT_*), an invalid key or a repeated key marks the row «КОПИЯ» (never posted);
' of repeated keys the first row in sheet order is kept as the original. Type EI («Заказы».V): the key is the canonical
' EI text; it must also be a receipt of this WMS (_RCV), and of repeated keys the row its receipt is registered at wins.
' Type RET («Возврат».A): the key is the return №; it must be registered in _RET, of repeated keys the row _RET names wins.
' Type SPR («Иной приход».A): the line №, registered in _SPR the same way.

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

' True when EI n is a receipt of this WMS; hintRow = the row its receipt is registered at
Private Function RcvKeyInfo(n As Long, ByRef hintRow As Long) As Boolean
    Dim c0 As Long, c1 As Long, row As Variant
    hintRow = -1
    If n < 1 Or n > MAX_SHEET_ROW Then Exit Function
    c0 = ((n - 1) \ CHECK_CHUNK_ROWS) * CHECK_CHUNK_ROWS + 1
    If c0 <> mRcvChunk Then
        c1 = c0 + CHECK_CHUNK_ROWS - 1
        If c1 > MAX_SHEET_ROW Then c1 = MAX_SHEET_ROW
        mRcvData = gDoc.Sheets.getByName(SH_RCV).getCellRangeByPosition(RV_EI, c0, RV_ROW, c1).getDataArray()
        mRcvChunk = c0
    End If
    row = mRcvData(n - c0)
    If CStr(row(RV_EI)) <> EI_PREFIX & Right("0000000" & n, 8) Then Exit Function
    If VarType(row(RV_ROW)) = 5 Then hintRow = CLng(row(RV_ROW))
    RcvKeyInfo = True
End Function

' True when № n is registered in the dense registry of the key type (_RET for returns, _SPR for the lines of «Иной
' приход»: № in column A); hintRow = the row of the user sheet it is registered at
Private Function RetKeyInfo(n As Long, ByRef hintRow As Long) As Boolean
    Dim c0 As Long, c1 As Long, row As Variant
    hintRow = -1
    If mRetLast < 0 Then mRetLast = LastUsedRow(gDoc.Sheets.getByName(mRegSheet))
    If n < 1 Or n > mRetLast Then Exit Function
    c0 = ((n - 1) \ CHECK_CHUNK_ROWS) * CHECK_CHUNK_ROWS + 1
    If c0 <> mRetChunk Then
        c1 = c0 + CHECK_CHUNK_ROWS - 1
        If c1 > mRetLast Then c1 = mRetLast
        mRetData = gDoc.Sheets.getByName(mRegSheet).getCellRangeByPosition(0, c0, mRegLast, c1).getDataArray()
        mRetChunk = c0
    End If
    row = mRetData(n - c0)
    If VarType(row(0)) <> 5 Then Exit Function
    If row(0) <> n Then Exit Function
    If VarType(row(mRegRow)) = 5 Then hintRow = CLng(row(mRegRow))
    RetKeyInfo = True
End Function

' a key of a registered type: registered by WMS (receipt of the EI / return №) and the row it is registered at
Private Function KeyInfo(isEI As Boolean, n As Long, ByRef hintRow As Long) As Boolean
    If isEI Then KeyInfo = RcvKeyInfo(n, hintRow) Else KeyInfo = RetKeyInfo(n, hintRow)
End Function

' «ЕИ-00000012 не создан приходом WMS» / «возврат № 5 не зарегистрирован WMS»
Private Function NotRegText(isEI As Boolean, n As Long) As String
    If isEI Then
        NotRegText = WmsIssue.EiCanon(n) & " не создан приходом WMS"
    ElseIf mRegSheet = SH_SPR Then
        NotRegText = "строка № " & n & " не зарегистрирована WMS"
    ElseIf mRegSheet = SH_ADJ Then
        NotRegText = "корректировка № " & n & " не зарегистрирована WMS"
    Else
        NotRegText = "возврат № " & n & " не зарегистрирован WMS"
    End If
End Function

Private Function LastUsedRow(sh As Object) As Long
    Dim cur As Object
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    LastUsedRow = cur.getRangeAddress().EndRow
End Function

Private Sub MarkCopy(sh As Object, ctlCol As Integer, r As Long, txt As String)
    sh.getCellByPosition(ctlCol, r).setString("КОПИЯ: " & txt & " — не проводится, очистите строку кнопкой «Очистить»")
End Sub

Private Function CheckKeySheet(spec As String, bMark As Boolean) As String
    Dim a As Variant, sName As String, keyCol As Integer, ctlCol As Integer, cntRow As Integer, inCol As Integer
    Dim sh As Object, last As Long, r0 As Long, r1 As Long, i As Long, j As Long, nextKey As Double, k As Variant
    Dim dKey As Variant, dCtl As Variant, dIn As Variant, buf() As Variant, nb As Long, nK As Long
    Dim nNever As Long, nBad As Long, nUnposted As Long, nDup As Long, t0 As Long, scr As Object, wasProt As Boolean
    Dim prevKey As Double, prevRow As Long, d As Variant, desc As Variant, fld(1) As New com.sun.star.table.TableSortField
    Dim isEI As Boolean, n As Long, reg As Boolean, hintRow As Long, origIsHint As Boolean, x As Long, kShow As String, isReg As Boolean
    On Error GoTo EH
    t0 = GetSystemTicks()
    a = Split(spec, "|")
    sName = Unesc(CStr(a(0)))
    keyCol = CInt(a(1))
    ctlCol = CInt(a(2))
    cntRow = CInt(a(3))
    inCol = CInt(a(4))
    If UBound(a) >= 5 Then
        isEI = (a(5) = "EI")
        isReg = (a(5) = "EI" Or a(5) = "RET" Or a(5) = "SPR" Or a(5) = "ADJ")
        ' the dense registry of a № key: _RET (returns), _SPR (the lines of «Иной приход») or _ADJ (the corrections), № in A,
        ' the row hint in RT_ROW / SR_ROW / AJ_ROW
        If a(5) = "SPR" Then
            mRegSheet = SH_SPR
            mRegLast = SR_LAST
            mRegRow = SR_ROW
        ElseIf a(5) = "ADJ" Then
            mRegSheet = SH_ADJ
            mRegLast = AJ_LAST
            mRegRow = AJ_ROW
        Else
            mRegSheet = SH_RET
            mRegLast = RT_LAST
            mRegRow = RT_ROW
        End If
    End If
    mRcvChunk = -1
    mRetChunk = -1
    mRetLast = -1
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
            If isEI And CStr(k) <> "" Then
                ' the EI key is text: its number is the key (a text that is not a canonical EI is invalid)
                If WmsOrders.StrictEI(CStr(k), n) Then k = CDbl(n) Else k = -1
            End If
            If Left(CStr(dCtl(i)(0)), 5) = "КОПИЯ" Then
                ' already marked
            ElseIf VarType(k) = 5 Then
                If k < 1 Or k <> Int(k) Then
                    nBad = nBad + 1
                    If bMark Then MarkCopy(sh, ctlCol, r0 + i, "неверный ключ «" & dKey(i)(0) & "»")
                ElseIf k >= nextKey Then
                    nNever = nNever + 1
                    If bMark Then MarkCopy(sh, ctlCol, r0 + i, "ключ " & dKey(i)(0) & " не выдавался WMS")
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
                    If isEI Then kShow = WmsIssue.EiCanon(CLng(d(j)(0))) Else kShow = CStr(d(j)(0))
                    x = CLng(d(j)(1))
                    If isReg And Not reg Then
                        nNever = nNever + 1
                        If bMark Then MarkCopy(sh, ctlCol, x, NotRegText(isEI, CLng(d(j)(0))))
                    ElseIf isReg And Not origIsHint And x = hintRow Then
                        ' the row the receipt / return is registered at is the original, the row seen first is the copy
                        nDup = nDup + 1
                        If bMark Then MarkCopy(sh, ctlCol, prevRow, kShow & " повторяет строку " & (x + 1))
                        prevRow = x
                        origIsHint = True
                    Else
                        nDup = nDup + 1
                        If bMark Then MarkCopy(sh, ctlCol, x, "ключ " & kShow & " повторяет строку " & (prevRow + 1))
                    End If
                Else
                    prevKey = d(j)(0)
                    prevRow = CLng(d(j)(1))
                    If isReg Then
                        reg = KeyInfo(isEI, CLng(prevKey), hintRow)
                        origIsHint = (prevRow = hintRow)
                        If Not reg Then
                            nNever = nNever + 1
                            If bMark Then MarkCopy(sh, ctlCol, prevRow, NotRegText(isEI, CLng(prevKey)))
                        End If
                    End If
                End If
            Next j
        Next r0
    ElseIf nK = 1 And isReg Then
        d = gSysSh.getCellRangeByPosition(3, 0, 4, 0).getDataArray()
        If Not KeyInfo(isEI, CLng(d(0)(0)), hintRow) Then
            nNever = nNever + 1
            If bMark Then MarkCopy(sh, ctlCol, CLng(d(0)(1)), NotRegText(isEI, CLng(d(0)(0))))
        End If
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
