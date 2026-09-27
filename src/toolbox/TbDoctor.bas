' TbDoctor — WMS_DOCTOR: аудит снимка WMS и его журнала (задание «FINAL WMS MARATHON», §6; M6 PRIME §8). Только чтение:
' ничего не исправляет, выдаёт понятный отчёт и план исправления. Проверки: целостность снимка, версии схемы и контракта
' снимка, журнал (формат, непрерывность, экземпляр, хвост), остатки (пересчёт по движениям), отрицательные остатки, ЕИ без
' карточки, связи возвратов с выдачами, счётчики, повторы номеров и копии, конфликт артикулов деталей, возможные дубли,
' непроведённые строки, защиты листов и структуры книги (по manifest снимка), транспортные визиты «Приход авто» (номера,
' незакрытые и повреждённые визиты, выезд раньше приезда, пустые обязательные поля, служебные записи, машины с прошлых
' дней), незавершённая загрузка пакетов, комплект инструментов выпуска, лишние файлы (незавершённый снимок, остаток
' обновления), свежесть резервной копии. DoctorCollect — те же проверки для «Контроля дня» WMS_MANAGER.
Option Explicit

Private Const SH_REP = "Отчёт"
Private mOut() As Variant
Private mN As Long

Private Sub AddRow(st As String, what As String, detail As String, sFix As String)
    If mN > UBound(mOut) Then Exit Sub
    mOut(mN) = Array(st, what, detail, sFix)
    mN = mN + 1
End Sub

Private Function TblRows(sName As String, nCols As Integer) As Variant
    Dim sh As Object, last As Long
    sh = ThisComponent.Sheets.getByName(sName)
    last = TbLastRow(sh)
    If last < 1 Then
        TblRows = Array()
    Else
        TblRows = sh.getCellRangeByPosition(0, 1, nCols - 1, last).getDataArray()
    End If
End Function

Private Function EINum(v As Variant) As Long
    Dim s As String
    s = CStr(v)
    If Left(s, 3) = "ЕИ-" And Len(s) = 11 And IsNumeric(Mid(s, 4)) Then EINum = CLng(Mid(s, 4))
End Function

Private Function Posted(st As Variant) As Boolean
    Posted = (Left(CStr(st), 9) = "Проведено")
End Function

' «Проверить»: the latest snapshot → the report
Sub BtnDoctor(Optional oEvent As Variant)
    TbMsg(DoctorRun())
End Sub

Function DoctorRun() As String
    Dim snap As String, sh As Object, last As Long, nFail As Long, nWarn As Long, j As Long
    On Error GoTo EH
    snap = TbLatestSnapshot()
    If snap = "" Then
        DoctorRun = "ERR:в папке " & ConvertFromURL(TbExportDir()) & " нет снимков — в WMS нажмите «Экспорт для инструментов»"
        Exit Function
    End If
    DoctorCollect(snap)
    sh = ThisComponent.Sheets.getByName(SH_REP)
    last = TbLastRow(sh)
    If last >= 3 Then sh.getCellRangeByPosition(0, 3, 3, last).clearContents(1023)
    If mN > 0 Then
        Dim r() As Variant
        ReDim r(mN - 1)
        For j = 0 To mN - 1
            r(j) = mOut(j)
            If mOut(j)(0) = "FAIL" Then nFail = nFail + 1
            If mOut(j)(0) = "WARN" Then nWarn = nWarn + 1
        Next j
        sh.getCellRangeByPosition(0, 3, 3, 3 + mN - 1).setDataArray(r)
    End If
    sh.getCellByPosition(0, 1).setString("Проверено " & Format(Now(), "DD.MM.YYYY HH:MM") & ": ошибок " & nFail & ", предупреждений " & nWarn _
        & " — снимок " & ConvertFromURL(snap))
    DoctorRun = "OK:ошибок " & nFail & ", предупреждений " & nWarn
    Exit Function
EH:
    DoctorRun = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' every check of snapshot snap (the tables are loaded into the hidden sheets of this book): the rows (status OK / WARN /
' FAIL, check, details, what to do). A snapshot that is not intact — only the row «снимок» FAIL
Function DoctorCollect(snap As String) As Variant
    Dim why As String, t As Variant, i As Long, r() As Variant
    ReDim mOut(80)
    mN = 0
    why = TbVerifySnapshot(snap)
    If why <> "" Then
        AddRow("FAIL", "снимок", why, "Сделайте новый снимок в WMS («Главная» → «Экспорт для инструментов»); этот не использовать")
    Else
        AddRow("OK", "снимок", ConvertFromURL(snap) & ": файлы и SHA-256 совпадают с manifest.csv", "")
        t = Array("stock", "orders", "special", "issues", "returns", "adjustments")
        For i = 0 To UBound(t)
            TbLoadTable(snap, t(i) & ".csv", "_" & t(i), True)
        Next i
        TbLoadTable(snap, "service/_ART.csv", "_art", True)
        CheckVersions(snap)
        CheckProtections(snap)
        CheckJournal(snap)
        CheckStock(snap)
        CheckCars(snap)
        CheckBatches(snap)
        CheckBackups(snap)
        CheckRelease()
        CheckOrphans(snap)
    End If
    If mN = 0 Then
        DoctorCollect = Array()
        Exit Function
    End If
    ReDim r(mN - 1)
    For i = 0 To mN - 1
        r(i) = mOut(i)
    Next i
    DoctorCollect = r
End Function

' the schema of _SYS (WMS-SYS-4 — WMS 0.7; WMS-SYS-3 — WMS 0.6, read without the vehicles) and the contract of the
' snapshot (1.0 — no cars.csv; 1.1 — with «Приход авто»; a newer one — a newer WMS than this toolbox)
Private Sub CheckVersions(snap As String)
    Dim sc As String, ct As String, info As String
    sc = TbManifestValue(snap, "schema")
    info = sc & ", ядро " & TbManifestValue(snap, "core_version") & ", продукт " & TbManifestValue(snap, "product_version")
    If sc = "WMS-SYS-4" Then
        AddRow("OK", "версия схемы", info, "")
    ElseIf sc = "WMS-SYS-3" Then
        AddRow("OK", "версия схемы", info & " (WMS 0.6: «Приход авто» не ведётся)", "")
    Else
        AddRow("WARN", "версия схемы", "схема «" & sc & "» — инструмент рассчитан на WMS-SYS-3 и WMS-SYS-4", "Обновите WMS_TOOLBOX до версии этой WMS")
    End If
    ct = TbManifestValue(snap, "contract")
    If ct = "" Then ct = "1.0"
    If ct = "1.0" Or ct = "1.1" Then
        AddRow("OK", "контракт снимка", "версия " & ct & IIf(ct = "1.0", " (без транспорта — снимок WMS 0.6)", " (с «Приход авто»)") & "; инструмент читает 1.0–" & TB_CONTRACT, "")
    Else
        AddRow("WARN", "контракт снимка", "версия " & ct & " новее инструмента (" & TB_CONTRACT & "): новые данные не проверяются", _
            "Обновите WMS_TOOLBOX из выпуска той же версии, что и WMS")
    End If
End Sub

' the protections of the book as WMS wrote them into the manifest («sheet;лист;защищён;виден», «structure;1»): every sheet
' protected but «Получатели» (a list kept by hand), the service sheets hidden, the structure protected — the rule of
' «Проверка перед работой» of WMS (WmsStatus.Protections)
Private Sub CheckProtections(snap As String)
    Dim lines As Variant, i As Long, f As Variant, p As String, n As Long, svc As String
    svc = "|_ORD|_RCV|_SPR|_ART|_RET|_ISS|_ADJ|_CAR|_IDX|_SYS|"
    lines = Split(Replace(TbReadText(snap & "manifest.csv"), Chr(13), ""), Chr(10))
    For i = 0 To UBound(lines)
        f = Split(lines(i), ";")
        If UBound(f) >= 3 Then
            If f(0) = "sheet" Then
                n = n + 1
                If f(2) <> "1" And f(1) <> "Получатели" Then p = p & IIf(p <> "", "; ", "") & "лист «" & f(1) & "» не защищён"
                If InStr(svc, "|" & f(1) & "|") > 0 And f(3) = "1" Then p = p & IIf(p <> "", "; ", "") & "служебный лист «" & f(1) & "» не скрыт"
            End If
        End If
    Next i
    If n = 0 Then
        AddRow("WARN", "защиты", "в снимке нет сведений о защитах листов (снимок старой версии WMS)", "Сделайте новый снимок в WMS")
        Exit Sub
    End If
    If TbManifestValue(snap, "structure") <> "1" Then p = p & IIf(p <> "", "; ", "") & "структура книги не защищена"
    If p = "" Then
        AddRow("OK", "защиты", "листов " & n & ": защищены (кроме «Получатели»), служебные скрыты, структура книги защищена", "")
    Else
        AddRow("WARN", "защиты", p, "Ответственный за WMS восстанавливает защиту: Сервис → Защитить лист / Защитить структуру книги " _
            & "(пароль WMS). Защита не даёт изменить учёт мимо WMS; пока её нет, проверяйте изменения «Самопроверкой»")
    End If
End Sub

' the journal copy of the snapshot: format of every line, continuous seq, one instance, the operations of the book
Private Sub CheckJournal(snap As String)
    Dim sfa As Object, files As Variant, i As Long, j As Long, txt As String, lines As Variant, f As Variant, seq As Long, prev As Long
    Dim nBad As Long, nLines As Long, firstBad As String, inst As String, lastSeq As Long, maxSeq As Long, nForeign As Long, p As Long, body As String
    Dim ln As String, gap As String, tmp As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(snap & "journal/") Then
        AddRow("FAIL", "журнал", "в снимке нет копии журнала", "Сделайте новый снимок")
        Exit Sub
    End If
    files = sfa.getFolderContents(snap & "journal/", False)
    inst = TbManifestValue(snap, "instance_id")
    lastSeq = Val(TbManifestValue(snap, "last_seq"))
    ' the file names sort by month
    For i = 1 To UBound(files)
        tmp = files(i)
        j = i - 1
        Do While j >= 0
            If files(j) <= tmp Then Exit Do
            files(j + 1) = files(j)
            j = j - 1
        Loop
        files(j + 1) = tmp
    Next i
    prev = -1
    For i = 0 To UBound(files)
        txt = TbReadText(files(i))
        lines = Split(txt, Chr(10))
        For j = 0 To UBound(lines)
            ln = lines(j)
            If ln <> "" Then
                nLines = nLines + 1
                p = InStr(ln, ";END;")
                f = Split(ln, ";", 6)
                If Left(ln, 3) <> "J1;" Or p = 0 Or UBound(f) < 5 Then
                    nBad = nBad + 1
                    If firstBad = "" Then firstBad = Mid(files(i), TbInStrRev(files(i), "/") + 1) & ", строка " & (j + 1)
                ElseIf Val(Mid(ln, p + 5)) <> p - 1 Then
                    nBad = nBad + 1
                    If firstBad = "" Then firstBad = Mid(files(i), TbInStrRev(files(i), "/") + 1) & ", строка " & (j + 1) & " (длина)"
                Else
                    If f(3) <> inst Then nForeign = nForeign + 1
                    seq = Val(f(1))
                    If prev >= 0 And seq <> prev + 1 And gap = "" Then gap = "после seq " & prev & " идёт " & seq
                    prev = seq
                    If seq > maxSeq Then maxSeq = seq
                End If
            End If
        Next j
    Next i
    If nBad > 0 Then
        AddRow("FAIL", "журнал: формат", "повреждённых строк " & nBad & " (первая — " & firstBad & ")", _
            "Не правьте журнал вручную. Откройте WMS: запуск покажет состояние журнала; при блокировке следуйте указаниям «Главной». Сохраните копию папки WMS_Journal для разбора")
    Else
        AddRow("OK", "журнал: формат", "строк " & nLines & ", у каждой признак конца и длина", "")
    End If
    If gap <> "" Then AddRow("FAIL", "журнал: непрерывность", gap, "Разбор ответственным: пропуск номера операции — признак потери строк журнала")
    If nForeign > 0 Then AddRow("FAIL", "журнал: экземпляр", "строк другой WMS: " & nForeign, "В папке журнала оказались файлы другой книги — уберите их и перезапустите WMS")
    If maxSeq < lastSeq Then
        AddRow("FAIL", "снимок против журнала", "в книге операций до seq " & lastSeq & ", в журнале последняя " & maxSeq, _
            "Журнал неполон: не работайте в WMS до разбора; восстановите папку WMS_Journal из резервной копии")
    ElseIf maxSeq > lastSeq Then
        AddRow("WARN", "снимок против журнала", "в журнале есть операции после состояния книги (seq " & (lastSeq + 1) & "…" & maxSeq & ")", _
            "Откройте WMS: при запуске она предложит «Восстановить» — эти операции применятся к книге")
    Else
        AddRow("OK", "снимок против журнала", "книга и журнал согласованы (seq " & lastSeq & ")", "")
    End If
End Sub

' balances against the movements, negative balances, EIs without a card, returns and issues, counters, numbers, parts
Private Sub CheckStock(snap As String)
    Dim st As Variant, o As Variant, sp As Variant, iss As Variant, rt As Variant, ad As Variant, art As Variant, i As Long, n As Long, top As Long
    Dim sumQ() As Double, hasMove() As Boolean, nEI As Long, nBadBal As Long, firstBad As String, nNeg As Long, nOrphan As Long, firstOrphan As String
    Dim issEI() As String, issSt() As Integer, maxIss As Long, nRetBad As Long, firstRet As String, nCopy As Long, nDup As Long, firstDup As String
    Dim seen As String, key As String, k As Long, nUnposted As Long, nPossDup As Long, nNoHist As Long, e As Long, q As Double
    Dim hasRcpt() As Boolean, nNoRcpt As Long, firstNoRcpt As String, bTest As Boolean, nMig As Long
    st = TblRows("_stock", 10)
    o = TblRows("_orders", 28)
    sp = TblRows("_special", 19)
    iss = TblRows("_issues", 18)
    rt = TblRows("_returns", 15)
    ad = TblRows("_adjustments", 19)
    art = TblRows("_art", 3)
    top = UBound(st) + 2
    ReDim sumQ(top)
    ReDim hasMove(top)
    ReDim hasRcpt(top)
    bTest = (TbManifestValue(snap, "mode") = "TEST")
    ' receipts of «Заказы»: live unless their status says the receipt was cancelled
    For i = 0 To UBound(o)
        If Left(CStr(o(i)(24)), 5) = "КОПИЯ" Then
            nCopy = nCopy + 1
        Else
            e = EINum(o(i)(21))
            If e > 0 And e <= top Then
                hasMove(e) = True
                hasRcpt(e) = True
                If InStr(CStr(o(i)(22)), "сторно") = 0 And VarType(o(i)(5)) = 5 Then sumQ(e) = sumQ(e) + o(i)(5)
            ElseIf CStr(o(i)(21)) <> "" Then
                nOrphan = nOrphan + 1
                If firstOrphan = "" Then firstOrphan = "«Заказы» строка " & (i + 2) & ": " & o(i)(21)
            End If
            If CStr(o(i)(27)) <> "" Then nPossDup = nPossDup + 1
            If CStr(o(i)(21)) = "" And VarType(o(i)(5)) = 5 Then nUnposted = nUnposted + 1
        End If
    Next i
    For i = 0 To UBound(sp)
        If Left(CStr(sp(i)(17)), 5) = "КОПИЯ" Then
            nCopy = nCopy + 1
        ElseIf VarType(sp(i)(0)) = 5 Then
            e = EINum(sp(i)(13))
            If e > 0 And e <= top Then
                hasMove(e) = True
                hasRcpt(e) = True
                If CStr(sp(i)(16)) <> "Удалено (сторно)" And VarType(sp(i)(5)) = 5 Then sumQ(e) = sumQ(e) + sp(i)(5)
            Else
                nOrphan = nOrphan + 1
                If firstOrphan = "" Then firstOrphan = "«Иной приход» строка " & (i + 2) & ": " & sp(i)(13)
            End If
        ElseIf CStr(sp(i)(1)) <> "" Then
            nUnposted = nUnposted + 1
        End If
    Next i
    maxIss = UBound(iss) + 2
    ReDim issEI(maxIss)
    ReDim issSt(maxIss)
    For i = 0 To UBound(iss)
        If Left(CStr(iss(i)(17)), 5) = "КОПИЯ" Then
            nCopy = nCopy + 1
        ElseIf VarType(iss(i)(0)) = 5 Then
            e = EINum(iss(i)(11))
            n = CLng(iss(i)(0))
            If n <= maxIss Then
                If issSt(n) <> 0 Then
                    nDup = nDup + 1
                    If firstDup = "" Then firstDup = "выдача № " & n
                End If
                issEI(n) = CStr(iss(i)(11))
                issSt(n) = IIf(Posted(iss(i)(17)), 1, 2)
            End If
            If e > 0 And e <= top Then
                hasMove(e) = True
                If Posted(iss(i)(17)) And VarType(iss(i)(4)) = 5 Then sumQ(e) = sumQ(e) - iss(i)(4)
            Else
                nOrphan = nOrphan + 1
                If firstOrphan = "" Then firstOrphan = "«Выдачи» строка " & (i + 2) & ": " & iss(i)(11)
            End If
        ElseIf CStr(iss(i)(11)) <> "" Then
            nUnposted = nUnposted + 1
        End If
    Next i
    For i = 0 To UBound(rt)
        If Left(CStr(rt(i)(13)), 5) = "КОПИЯ" Then
            nCopy = nCopy + 1
        ElseIf VarType(rt(i)(0)) = 5 Then
            e = EINum(rt(i)(2))
            n = IIf(VarType(rt(i)(1)) = 5, CLng(rt(i)(1)), 0)
            If CStr(rt(i)(13)) <> "Удалено (сторно)" Then
                If n < 1 Or n > maxIss Then
                    nRetBad = nRetBad + 1
                    If firstRet = "" Then firstRet = "возврат № " & rt(i)(0) & ": выдачи № " & rt(i)(1) & " нет"
                ElseIf issSt(n) <> 1 Or issEI(n) <> CStr(rt(i)(2)) Then
                    nRetBad = nRetBad + 1
                    If firstRet = "" Then firstRet = "возврат № " & rt(i)(0) & ": выдача № " & n & IIf(issSt(n) = 2, " удалена (сторно)", " другого ЕИ")
                End If
            End If
            If e > 0 And e <= top Then
                hasMove(e) = True
                If CStr(rt(i)(13)) <> "Удалено (сторно)" And VarType(rt(i)(5)) = 5 Then sumQ(e) = sumQ(e) + rt(i)(5)
            End If
        End If
    Next i
    For i = 0 To UBound(ad)
        If Left(CStr(ad(i)(16)), 5) = "КОПИЯ" Then
            nCopy = nCopy + 1
        ElseIf VarType(ad(i)(0)) = 5 Then
            e = EINum(ad(i)(2))
            If e > 0 And e <= top Then
                hasMove(e) = True
                If CStr(ad(i)(16)) <> "Удалено (сторно)" And VarType(ad(i)(15)) = 5 Then sumQ(e) = sumQ(e) + ad(i)(15)
            Else
                nOrphan = nOrphan + 1
                If firstOrphan = "" Then firstOrphan = "«Корректировки» строка " & (i + 2) & ": " & ad(i)(2)
            End If
        ElseIf CStr(ad(i)(1)) <> "" Then
            nUnposted = nUnposted + 1
        End If
    Next i
    ' the balance of every EI: received − issued + returned + corrections. A migrated EI (source «Перенос…») starts with
    ' its transferred quantity (in the journal, not in the tables); an EI of the synthetic registry of a TEST book (source
    ' «тест») with its initial balance; any other EI without a receipt is a card without an origin
    For i = 0 To UBound(st)
        If CStr(st(i)(0)) <> "" Then
            nEI = nEI + 1
            e = i + 1
            q = 0
            If VarType(st(i)(4)) = 5 Then q = st(i)(4)
            If q < -0.0000001 Then nNeg = nNeg + 1
            If Left(CStr(st(i)(8)), 7) = "Перенос" Then
                nMig = nMig + 1
            ElseIf Not hasRcpt(e) Then
                If bTest And CStr(st(i)(8)) = "тест" Then
                    nNoHist = nNoHist + 1
                Else
                    nNoRcpt = nNoRcpt + 1
                    If firstNoRcpt = "" Then firstNoRcpt = CStr(st(i)(0))
                End If
            ElseIf Abs(sumQ(e) - q) > 0.0005 Then
                nBadBal = nBadBal + 1
                If firstBad = "" Then firstBad = st(i)(0) & ": остаток " & q & ", по движениям " & Round3(sumQ(e))
            End If
        End If
    Next i
    If nBadBal > 0 Then
        AddRow("FAIL", "остатки", "не сходятся с движениями у " & nBadBal & " ЕИ (первый — " & firstBad & ")", _
            "В WMS нажмите «Самопроверка»; расхождение без ошибки самопроверки — разбор ответственным по журналу; исправление — только операциями (инвентаризация)")
    Else
        AddRow("OK", "остатки", "ЕИ " & nEI & ": остаток = сумме движений" & IIf(nMig > 0, "; перенесённых (начало — в журнале) " & nMig, "") _
            & IIf(nNoHist > 0, "; начального реестра тестовой книги " & nNoHist, ""), "")
    End If
    If nNoRcpt > 0 Then AddRow("FAIL", "ЕИ без прихода", "карточек без прихода и без переноса: " & nNoRcpt & " (первая — " & firstNoRcpt & ")", _
        "ЕИ появляется только приходом или переносом: разбор ответственным по журналу")
    If nNeg > 0 Then AddRow("FAIL", "отрицательные остатки", "ЕИ: " & nNeg, "Ошибка целостности: не работайте с этими ЕИ до разбора; WMS такое не проводит")
    If nOrphan > 0 Then
        AddRow("FAIL", "ЕИ без карточки", "строк: " & nOrphan & " (первая — " & firstOrphan & ")", "Строка ссылается на ЕИ, которого нет в «Наличие»: разбор ответственным")
    Else
        AddRow("OK", "ЕИ без карточки", "нет", "")
    End If
    If nRetBad > 0 Then
        AddRow("FAIL", "связи возвратов", "нарушений " & nRetBad & " (" & firstRet & ")", "Возврат должен ссылаться на действующую выдачу того же ЕИ: разбор ответственным")
    Else
        AddRow("OK", "связи возвратов", "каждый действующий возврат — к действующей выдаче того же ЕИ", "")
    End If
    If nDup > 0 Then AddRow("FAIL", "повторы номеров", "повторов: " & nDup & " (" & firstDup & ")", "Строка с чужим № без пометки КОПИЯ: откройте WMS — полная проверка ключей пометит копии")
    If nCopy > 0 Then AddRow("WARN", "строки-копии", "строк с пометкой КОПИЯ: " & nCopy, "Очистите их кнопкой «Очистить» на своём листе")
    CheckCounters(snap, st, iss, rt, sp, ad)
    CheckParts(st, art)
    If nPossDup > 0 Then AddRow("WARN", "возможные дубли", "строк «Заказов» с пометкой «Возможный дубль»: " & nPossDup, "Проверьте документы этих поступлений")
    If nUnposted > 0 Then AddRow("WARN", "непроведённые строки", "строк с вводом без номера: " & nUnposted, "Проведите или очистите их в WMS («Главная» показывает, где)")
End Sub

Private Function Round3(x As Double) As Double
    Round3 = Int(x * 1000 + 0.5) / 1000
End Function

Private Sub CheckCounters(snap As String, st As Variant, iss As Variant, rt As Variant, sp As Variant, ad As Variant)
    Dim i As Long, mx(4) As Long, nx(4) As Long, names As Variant, bad As String, k As Integer
    For i = 0 To UBound(st)
        If CStr(st(i)(0)) <> "" Then mx(0) = i + 1
    Next i
    For i = 0 To UBound(iss)
        If VarType(iss(i)(0)) = 5 Then
            If iss(i)(0) > mx(1) Then mx(1) = iss(i)(0)
        End If
    Next i
    For i = 0 To UBound(rt)
        If VarType(rt(i)(0)) = 5 Then
            If rt(i)(0) > mx(2) Then mx(2) = rt(i)(0)
        End If
    Next i
    For i = 0 To UBound(sp)
        If VarType(sp(i)(0)) = 5 Then
            If sp(i)(0) > mx(3) Then mx(3) = sp(i)(0)
        End If
    Next i
    For i = 0 To UBound(ad)
        If VarType(ad(i)(0)) = 5 Then
            If ad(i)(0) > mx(4) Then mx(4) = ad(i)(0)
        End If
    Next i
    names = Array("next_ei", "next_issue", "next_return", "next_line", "next_adj")
    For k = 0 To 4
        nx(k) = Val(TbManifestValue(snap, names(k)))
        If nx(k) <= mx(k) Then bad = bad & IIf(bad <> "", "; ", "") & names(k) & " " & nx(k) & " не выше " & mx(k)
    Next k
    If bad <> "" Then
        AddRow("FAIL", "счётчики", bad, "Номер мог быть выдан повторно: не проводите операции до разбора; WMS при запуске проверяет счётчики")
    Else
        AddRow("OK", "счётчики", "NEXT_EI, выдачи, возвраты, строки иного прихода, корректировки — выше всех номеров", "")
    End If
End Sub

Private Function ArtKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, a As String, b As String
    s = UCase(s)
    a = "АВЕЁКМНОРСТУХ"
    b = "ABEEKMHOPCTYX"
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = " " Or ch = Chr(9) Or ch = Chr(160) Then
            ch = ""
        ElseIf ch = "‐" Or ch = "‑" Or ch = "‒" Or ch = "–" Or ch = "—" Or ch = "−" Then
            ch = "-"
        ElseIf InStr(a, ch) > 0 Then
            ch = Mid(b, InStr(a, ch), 1)
        End If
        out = out & ch
    Next i
    ArtKey = out
End Function

Private Sub CheckParts(st As Variant, art As Variant)
    Dim i As Long, j As Long, n As Long, keys() As String, eis() As String, bad As String, nBad As Long, k As String, t As String, te As String
    ReDim keys(UBound(st) + 1)
    ReDim eis(UBound(st) + 1)
    For i = 0 To UBound(st)
        If CStr(st(i)(9)) = "Детали" And CStr(st(i)(7)) <> "Приход удалён (сторно)" Then
            keys(n) = ArtKey(CStr(st(i)(2)))
            eis(n) = CStr(st(i)(0))
            n = n + 1
        End If
    Next i
    ' sort by the key (a shell sort), then equal neighbours are conflicts
    Dim gap As Long
    gap = n \ 2
    Do While gap > 0
        For i = gap To n - 1
            k = keys(i)
            te = eis(i)
            j = i
            Do While j >= gap
                If keys(j - gap) <= k Then Exit Do
                keys(j) = keys(j - gap)
                eis(j) = eis(j - gap)
                j = j - gap
            Loop
            keys(j) = k
            eis(j) = te
        Next i
        gap = gap \ 2
    Loop
    For i = 1 To n - 1
        If keys(i) = keys(i - 1) And keys(i) <> "" Then
            nBad = nBad + 1
            If bad = "" Then bad = "артикул «" & keys(i) & "»: " & eis(i - 1) & " и " & eis(i)
        End If
    Next i
    If nBad > 0 Then
        AddRow("FAIL", "артикулы деталей", "один артикул у нескольких ЕИ: " & nBad & " (" & bad & ")", _
            "Такой артикул WMS не принимает («найдено несколько ЕИ»): решите, какой ЕИ основной; остаток другого — инвентаризацией и перемещением")
    Else
        AddRow("OK", "артикулы деталей", "деталей " & n & ", у каждого артикула один ЕИ", "")
    End If
End Sub

' the backups of the WMS folder (the parent of the folder of the snapshots: …/WMS_Backups): one made today
Private Sub CheckBackups(snap As String)
    Dim sfa As Object, d As String, lst As Variant, i As Long, best As Date, dt As Date, n As Long
    Dim dtm As New com.sun.star.util.DateTime
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbWmsDir() & "WMS_Backups/"
    If Not sfa.exists(d) Then
        AddRow("WARN", "резервные копии", "папки WMS_Backups рядом с книгой не видно (" & ConvertFromURL(d) & ")", "Проверьте диск рабочей WMS: копии создаются при запуске и кнопкой «Резервная копия»")
        Exit Sub
    End If
    lst = sfa.getFolderContents(d, False)
    For i = 0 To UBound(lst)
        dtm = sfa.getDateTimeModified(lst(i))
        dt = DateSerial(dtm.Year, dtm.Month, dtm.Day) + TimeSerial(dtm.Hours, dtm.Minutes, dtm.Seconds)
        If dt > best Then best = dt
        n = n + 1
    Next i
    If n = 0 Then
        AddRow("FAIL", "резервные копии", "в WMS_Backups нет копий", "В WMS нажмите «Резервная копия»")
    ElseIf Int(CDbl(best)) < Int(CDbl(Now())) Then
        AddRow("WARN", "резервные копии", "последняя копия " & Format(best, "DD.MM.YYYY HH:MM") & " — сегодня копии не было" _
            & IIf(Now() - best > 2, " (старше двух дней)", ""), "В WMS нажмите «Резервная копия»; проверьте, что WMS открывают каждый рабочий день")
    Else
        AddRow("OK", "резервные копии", "копий " & n & ", последняя " & Format(best, "DD.MM.YYYY HH:MM"), "")
    End If
End Sub

' ================================================================ M6: vehicles, batches, the release, stray files

Private Function DtText(v As Variant) As String
    If VarType(v) = 5 Or VarType(v) = 7 Then DtText = Format(CDate(v), "DD.MM.YYYY HH:MM") Else DtText = CStr(v)
End Function

' «Приход авто» (cars.csv, the row k of the table — visit k; empty rows — the numbers of abandoned operations) and its
' service table _CAR: numbers, statuses, the times, the required fields, the counter, the service records; the vehicles
' left «На территории» since an earlier day
Private Sub CheckCars(snap As String)
    Dim n As Long, d As Variant, i As Long, nV As Long, nOpen As Long, nCl As Long, nCanc As Long, bad As Long, first As String, st As String
    Dim p As String, old As String, nOld As Long, created As Double, cs As String, sv As Variant, nSvc As Long, maxNo As Long, nxt As Long, want As String
    If Not TbHasFile(snap, "cars.csv") Then
        AddRow("OK", "транспорт", "в снимке нет данных «Приход авто» (снимок WMS до версии 0.7)", "")
        Exit Sub
    End If
    n = TbLoadTable(snap, "cars.csv", "_cars", True)
    If TbHasFile(snap, "service/_CAR.csv") Then nSvc = TbLoadTable(snap, "service/_CAR.csv", "_car", True)
    cs = TbManifestValue(snap, "created")
    If Len(cs) >= 10 Then created = CDbl(DateSerial(Val(Left(cs, 4)), Val(Mid(cs, 6, 2)), Val(Mid(cs, 9, 2)))) Else created = Int(CDbl(Now()))
    If n > 0 Then d = TblRows("_cars", 14) Else d = Array()
    If nSvc > 0 Then sv = TblRows("_car", 3) Else sv = Array()
    For i = 0 To UBound(d)
        p = ""
        If CStr(d(i)(0)) = "" And CStr(d(i)(2)) = "" And CStr(d(i)(7)) = "" Then GoTo NEXT_CAR
        If VarType(d(i)(0)) <> 5 Then
            p = "строка " & (i + 2) & ": нет номера визита"
        ElseIf CLng(d(i)(0)) <> i + 1 Then
            p = "строка " & (i + 2) & ": визит № " & d(i)(0) & " не на своём месте (повтор или сдвиг номера)"
        Else
            If CLng(d(i)(0)) > maxNo Then maxNo = CLng(d(i)(0))
            st = CStr(d(i)(7))
            Select Case st
            Case "Отменён"
                nCanc = nCanc + 1
            Case "На территории", "Уехал"
                nV = nV + 1
                If Trim(CStr(d(i)(2))) = "" Then
                    p = "визит № " & d(i)(0) & ": пустая «Марка / госномер»"
                ElseIf Trim(CStr(d(i)(3))) = "" Then
                    p = "визит № " & d(i)(0) & ": пустой «Поставщик»"
                ElseIf VarType(d(i)(4)) <> 5 Or VarType(d(i)(1)) <> 5 Then
                    p = "визит № " & d(i)(0) & ": нет даты или времени приезда"
                ElseIf st = "Уехал" Then
                    nCl = nCl + 1
                    If VarType(d(i)(5)) <> 5 Then
                        p = "визит № " & d(i)(0) & ": «Уехал» без времени выезда"
                    ElseIf d(i)(5) < d(i)(4) - 1 / 86400 Then
                        p = "визит № " & d(i)(0) & ": выезд " & DtText(d(i)(5)) & " раньше приезда " & DtText(d(i)(4))
                    ElseIf VarType(d(i)(6)) <> 5 Then
                        p = "визит № " & d(i)(0) & ": нет длительности"
                    ElseIf Abs(d(i)(6) - (d(i)(5) - d(i)(4)) * 1440) > 1.01 Then
                        p = "визит № " & d(i)(0) & ": длительность " & d(i)(6) & " мин не равна выезд − приезд"
                    End If
                Else
                    nOpen = nOpen + 1
                    If VarType(d(i)(5)) = 5 Then
                        p = "визит № " & d(i)(0) & ": «На территории», но записан выезд"
                    ElseIf Int(d(i)(4)) < created Then
                        nOld = nOld + 1
                        If old = "" Then old = "визит № " & d(i)(0) & " «" & d(i)(2) & "» с " & DtText(d(i)(4))
                    End If
                End If
            Case Else
                p = "визит № " & d(i)(0) & ": неизвестный статус «" & st & "»"
            End Select
            ' the service record of the visit: its state is the status of the sheet
            If p = "" And nSvc > 0 Then
                want = IIf(st = "Уехал", "CLOSED", IIf(st = "Отменён", "CANCELLED", "OPEN"))
                If i > UBound(sv) Then
                    p = "визит № " & d(i)(0) & ": нет служебной записи _CAR"
                ElseIf CStr(sv(i)(0)) <> CStr(d(i)(0)) Or CStr(sv(i)(2)) <> want Then
                    p = "визит № " & d(i)(0) & ": служебная запись _CAR (" & sv(i)(2) & ") не совпадает с листом (" & st & ")"
                End If
            End If
        End If
        If p <> "" Then
            bad = bad + 1
            If first = "" Then first = p
        End If
NEXT_CAR:
    Next i
    nxt = Val(TbManifestValue(snap, "next_car"))
    If nxt <= maxNo Then
        bad = bad + 1
        If first = "" Then first = "счётчик NEXT_CAR " & nxt & " не выше последнего визита " & maxNo
    End If
    If bad > 0 Then
        AddRow("FAIL", "транспорт", "нарушений " & bad & " (первое — " & first & ")", "Визиты меняются только кнопками листа «Приход авто» (УЕХАЛ, «Исправить», " _
            & "«Отменить визит»); в WMS нажмите «Самопроверка»; повреждение без ошибки самопроверки — разбор ответственным по журналу")
    Else
        AddRow("OK", "транспорт", "визитов " & (nV + nCanc) & " (на территории " & nOpen & ", уехало " & nCl & ", отменено " & nCanc _
            & "): номера, статусы, время, обязательные поля, служебные записи, счётчик — без ошибок", "")
    End If
    If nOld > 0 Then AddRow("WARN", "транспорт: машины с прошлых дней", "на территории с прошлых дней: " & nOld & " (" & old & ")", _
        "Если машина уехала — выделите её строку на листе «Приход авто» и нажмите УЕХАЛ, затем «Исправить» время выезда")
End Sub

' a batch loaded into WMS but not posted (rows with its mark and without a number), and the batch files WMS never loaded
Private Sub CheckBatches(snap As String)
    Dim o As Variant, sp As Variant, ad As Variant, i As Long, n As Long, ids As String, loaded As String, sfa As Object, d As String, lst As Variant
    Dim id As String, lines As Variant, j As Long, nWait As Long, waitIds As String
    o = TblRows("_orders", 28)
    sp = TblRows("_special", 19)
    ad = TblRows("_adjustments", 19)
    For i = 0 To UBound(ad)
        If CStr(ad(i)(17)) <> "" Then
            loaded = loaded & "|" & ad(i)(17) & "|"
            If VarType(ad(i)(0)) <> 5 And Left(CStr(ad(i)(16)), 5) <> "КОПИЯ" Then
                n = n + 1
                If InStr(1, ids, CStr(ad(i)(17)), 0) = 0 Then ids = ids & IIf(ids <> "", ", ", "") & ad(i)(17)
            End If
        End If
    Next i
    For i = 0 To UBound(sp)
        id = BatchMark(CStr(sp(i)(18)))
        If id <> "" Then
            loaded = loaded & "|" & id & "|"
            If VarType(sp(i)(0)) <> 5 And Left(CStr(sp(i)(17)), 5) <> "КОПИЯ" Then
                n = n + 1
                If InStr(1, ids, id, 0) = 0 Then ids = ids & IIf(ids <> "", ", ", "") & id
            End If
        End If
    Next i
    For i = 0 To UBound(o)
        id = BatchMark(CStr(o(i)(25)))
        If id <> "" Then
            loaded = loaded & "|" & id & "|"
            If CStr(o(i)(21)) = "" And Left(CStr(o(i)(24)), 5) <> "КОПИЯ" Then
                n = n + 1
                If InStr(1, ids, id, 0) = 0 Then ids = ids & IIf(ids <> "", ", ", "") & id
            End If
        End If
    Next i
    If n > 0 Then
        AddRow("WARN", "пакеты: незавершённая загрузка", "строк пакетов загружено, но не проведено: " & n & " (пакеты: " & ids & ")", _
            "В WMS проведите эти строки кнопками их листа или очистите их; повторно тот же пакет WMS не загрузит")
    Else
        AddRow("OK", "пакеты: незавершённая загрузка", "строк загруженных, но не проведённых пакетов нет", "")
    End If
    ' the batch files of the folder of the batches: a file whose #id no row of the snapshot carries waits for loading
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbBatchDir()
    If sfa.exists(d) Then
        lst = sfa.getFolderContents(d, False)
        For i = 0 To UBound(lst)
            If LCase(Right(lst(i), 4)) = ".csv" Then
                lines = Split(Replace(TbReadText(lst(i)), Chr(13), ""), Chr(10))
                id = ""
                For j = 0 To IIf(UBound(lines) > 8, 8, UBound(lines))
                    If Left(lines(j), 4) = "#id;" Then id = Mid(lines(j), 5)
                Next j
                If id <> "" And InStr(1, loaded, "|" & id & "|", 0) = 0 Then
                    nWait = nWait + 1
                    If nWait <= 5 Then waitIds = waitIds & IIf(waitIds <> "", ", ", "") & id
                End If
            End If
        Next i
    End If
    If nWait > 0 Then AddRow("WARN", "пакеты: ожидают загрузки", "файлов пакетов, которых нет в WMS по снимку: " & nWait & " (" & waitIds & ")", _
        "Если пакет нужен — в WMS «Главная» → «Загрузить пакет»; ненужный файл перенесите из папки пакетов")
End Sub

' the id of the mark «[пакет <id>]» of a comment ("" — none)
Private Function BatchMark(s As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, "[пакет ", 0)
    If p = 0 Then Exit Function
    q = InStr(p, s, "]")
    If q > p + 7 Then BatchMark = Mid(s, p + 7, q - p - 7)
End Function

' the set of the release next to this book: the tool books, the launcher, the scripts, VERSION of this very version
Private Sub CheckRelease()
    Dim sfa As Object, f As Variant, miss As String, n As Integer, v As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    For Each f In Array("WMS_TOOLBOX.ods", "WMS_INVENTORY.ods", "WMS_ANALYTICS.ods", "WMS_MANAGER.ods", "WMS_SEARCH.ods", "WMS_DOCTOR.ods", "WMS_LABELS.ods", _
        "WMS_DOCS.ods", "WMS_ARCHIVE.ods", "WMS_IMPORTER.ods", "reconcile/wms_reconcile.py", "backup/wms_backup.py", "backup/wms_healthcheck.py", _
        "labels/tspl_labels.py", "insights/wms_insights.py", "VERSION")
        n = n + 1
        If Not sfa.exists(TbFolder() & f) Then miss = miss & IIf(miss <> "", ", ", "") & f
    Next f
    If sfa.exists(TbFolder() & "VERSION") Then v = Trim(Replace(Replace(TbReadText(TbFolder() & "VERSION"), Chr(10), ""), Chr(13), ""))
    If miss <> "" Then
        AddRow("WARN", "комплект инструментов", "нет файлов: " & miss, "Распакуйте папку WMS_TOOLBOX из архива выпуска заново (данные книг — журнал работы, история актов — перенесите вручную)")
    ElseIf v <> "WMS_TOOLBOX " & TB_VERSION Then
        AddRow("WARN", "комплект инструментов", "VERSION «" & v & "», а эта книга — WMS_TOOLBOX " & TB_VERSION, "Книги из разных выпусков: распакуйте WMS_TOOLBOX одного выпуска")
    Else
        AddRow("OK", "комплект инструментов", "файлов " & n & ": книги инструментов, launcher, скрипты, VERSION — " & v, "")
    End If
End Sub

' stray files of the WMS folder: an export that did not finish (.…part, a snapshot folder without manifest.csv), the rest
' of an update (…upgrade.ods)
Private Sub CheckOrphans(snap As String)
    Dim sfa As Object, w As String, e As String, lst As Variant, i As Long, nm As String, found As String, n As Long
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    w = TbWmsDir()
    e = TbExportDir()
    If sfa.exists(e) Then
        lst = sfa.getFolderContents(e, True)
        For i = 0 To UBound(lst)
            nm = Mid(lst(i), Len(e) + 1)
            If Right(nm, 1) = "/" Then nm = Left(nm, Len(nm) - 1)
            If Right(nm, 5) = ".part" Then
                n = n + 1
                found = found & IIf(found <> "", "; ", "") & "WMS_Export/" & nm & " (незавершённый снимок)"
            ElseIf Left(nm, 13) = "WMS_SNAPSHOT_" And sfa.isFolder(lst(i)) Then
                If Not sfa.exists(lst(i) & IIf(Right(lst(i), 1) = "/", "", "/") & "manifest.csv") Then
                    n = n + 1
                    found = found & IIf(found <> "", "; ", "") & "WMS_Export/" & nm & " (снимок без manifest.csv)"
                End If
            End If
        Next i
    End If
    If sfa.exists(w) Then
        lst = sfa.getFolderContents(w, False)
        For i = 0 To UBound(lst)
            nm = Mid(lst(i), Len(w) + 1)
            If Right(nm, 12) = ".upgrade.ods" Then
                n = n + 1
                found = found & IIf(found <> "", "; ", "") & nm & " (остаток незавершённого обновления)"
            End If
        Next i
    End If
    If n > 0 Then
        AddRow("WARN", "лишние файлы", found, "Эти файлы не используются: убедитесь, что WMS закрыта и обновление не идёт, и удалите их (или перенесите для разбора)")
    Else
        AddRow("OK", "лишние файлы", "незавершённых снимков и остатков обновления нет", "")
    End If
End Sub

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function DoctorProbe() As String
    DoctorProbe = "OK"
End Function
