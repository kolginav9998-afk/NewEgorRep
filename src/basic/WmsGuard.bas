' WmsGuard — 1.1: protection of the data from what the storekeeper does with the book itself.
' «Наличие» is addressed by the row: row n holds ЕИ-n (a number that was never given is an empty row). A sort (Data → Sort, the AutoFilter of the
' sheet) carries the empty rows to the end and moves every row below the first of them up — nothing is lost from the journal, but the registry is no
' longer the registry. The guard (1) finds it at every start in one formula of the Calc engine and blocks the posting with the words that say what
' to do; (2) when such a book is being saved, keeps the file that is about to be overwritten as the copy «последняя без ошибок» (the journal is
' ahead of it, «Вернуться к копии» brings it up to date) and says so aloud; (3) after a good save keeps a copy of it, not oftener than every
' GOOD_EVERY_MIN minutes. Library WmsPlus.
Option Explicit

Public Const GOOD_EVERY_MIN = 20

' the number of the rows of «Наличие» that hold an EI other than the one their row is for (row n — ЕИ-n); 0 — the registry is in order;
' -1 — the formula could not be used (the caller takes the slower check)
Function RegistryMisplaced() As Long
    Dim cur As Object, last As Long, rng As String, s As String, i As Long, c As String
    RegistryMisplaced = -1
    On Error GoTo EH
    cur = WmsIssue.StockSheet().createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    If last < 1 Then
        RegistryMisplaced = 0
        Exit Function
    End If
    rng = WmsView.ColRng(SH_STOCK, SC_EI, 1, last)
    s = WmsView.Eng("SUMPRODUCT((" & rng & "<>"""")*(VALUE(""0""&MID(" & rng & ";" & (Len(EI_PREFIX) + 1) & ";" & EI_DIGITS & "))<>ROW(" & rng & ")-1))")
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    RegistryMisplaced = CLng(s)
    Exit Function
EH:
    RegistryMisplaced = -1
End Function

' "" — the registry is in order; otherwise what is wrong, in words (the first row that is not where it should be)
Function RegistryProblem() As String
    Dim n As Long, s As String, p As Long
    n = RegistryMisplaced()
    If n = 0 Then Exit Function
    s = WmsDiagnostics.RegistryCheck()
    If Left(s, 6) <> "ОШИБКА" Then
        If n > 0 Then RegistryProblem = "реестр «" & SH_STOCK & "»: " & n & " строк стоят не на своих местах"
        Exit Function
    End If
    p = InStr(s, "(строка")
    RegistryProblem = "реестр «" & SH_STOCK & "» нарушен: строки ЕИ стоят не на своих местах" & IIf(n > 0, " (" & n & ")", "") _
        & IIf(p > 0, " — " & Mid(s, p + 1, Len(s) - p - 1), "") & ". Так бывает, если реестр отсортировали"
End Function

' the book is being saved (OnSave: the file on the disk is still the previous save): a registry out of order is not allowed to replace it unnoticed
Sub SaveGuard()
    Dim p As String, r As String
    On Error GoTo EH
    If gState = "" Then Exit Sub
    p = RegistryProblem()
    If p = "" Then Exit Sub
    ' the first save of a book that is out of order: the file that is about to be overwritten is the last one that was good
    If WmsSafe.StGet("RED") = "" Then
        ' the file on the disk was copied right after its own (good) save: the same copy is not made twice (the few copies of this kind are kept in turn)
        If WmsSafe.StGet("GOODSIZE") <> CStr(gSfa.getSize(gDoc.getURL())) Then r = WmsBackup.BackupNow("good")
        WmsSafe.StSet("RED", Format(Now(), "YYYY-MM-DD HH:MM"))
    End If
    If gBlock = "" Then WmsCore.SetBlocked(BLK_REGISTRY, p)
    WmsUi.UiMessage("ВНИМАНИЕ: " & p & "." & Chr(10) & Chr(10) _
        & "Эта книга сейчас сохраняется, но после этого проводить в ней нельзя." & Chr(10) _
        & "Последнее хорошее состояние сохранено как копия «последняя без ошибок»: «Обслуживание» → «Вернуться к копии…» вернёт его, " _
        & "а проведённое после него WMS доведёт по журналу." & Chr(10) _
        & "Если книгу ещё не сохраняли после сортировки — закройте её БЕЗ сохранения.")
    Exit Sub
EH:
End Sub

' the book is saved (OnSaveDone): in order — a copy of it as «последняя без ошибок» (every GOOD_EVERY_MIN minutes at the most), the red mark off
Sub SaveDoneGuard()
    Dim p As String, t As String, d As Double
    On Error GoTo EH
    ' the size of the file that has a copy of its own ("" — the file saved last has none: the copy of «good» is then made by SaveGuard)
    If WmsSafe.StGet("GOODSIZE") <> "" Then WmsSafe.StSet("GOODSIZE", "")
    If gState <> "CLEAN" Then Exit Sub
    If WmsSafe.StGet("RED") <> "" Then
        p = RegistryProblem()
        If p <> "" Then Exit Sub
        WmsSafe.StSet("RED", "")
    End If
    t = WmsSafe.StGet("GOOD")
    If t <> "" Then
        d = CDbl(DateSerial(Val(Left(t, 4)), Val(Mid(t, 6, 2)), Val(Mid(t, 9, 2)))) + CDbl(TimeSerial(Val(Mid(t, 12, 2)), Val(Mid(t, 15, 2)), 0))
        If Now() - d < GOOD_EVERY_MIN / 1440 And Now() >= d Then Exit Sub
    End If
    If RegistryMisplaced() <> 0 Then Exit Sub
    t = WmsBackup.BackupNow("good")
    If Left(t, 6) <> "ОШИБКА" And InStr(t, "невозможна") = 0 Then WmsSafe.StSet("GOODSIZE", CStr(gSfa.getSize(gDoc.getURL())))
    WmsSafe.StSet("GOOD", Format(Now(), "YYYY-MM-DD HH:MM"))
    Exit Sub
EH:
End Sub

' ================================================================ «Что делать?»
' The self-check says what is wrong; this says what to do about it — line by line, in the words of the storekeeper.

' what to do about a line of the self-check (the line as «FAIL заголовок: текст»); "" — nothing to add
Private Function AdviceFor(ln As String) As String
    Dim p As Long, code As String, hint As String
    If InStr(ln, "реестр «" & SH_STOCK & "»") > 0 And Left(ln, 4) = "FAIL" Then
        AdviceFor = "Реестр «Наличие» нарушен (строки сдвинуты — обычно его отсортировали). Не проводите. Если после сортировки книгу ещё не сохраняли — закройте её БЕЗ сохранения и откройте заново. " _
            & "Если сохранили — «Обслуживание» → «Вернуться к копии…» → «последняя без ошибок»."
    ElseIf InStr(ln, "сеанс WMS") > 0 And InStr(ln, "[") > 0 And InStr(ln, "BLOCKED") > 0 Then
        p = InStr(ln, "BLOCKED [")
        If p > 0 Then code = Mid(ln, p + 9, InStr(p, ln, "]") - p - 9)
        hint = WmsUi.BlockHint(code)
        If code = "TAIL" Then hint = hint & " Посмотреть, что это за операции: «Главная» → «Операции дня» → «Хвост журнала»."
        AdviceFor = hint
    ElseIf InStr(ln, "журнал:") > 0 And InStr(ln, "вне книги") > 0 Then
        AdviceFor = "В журнале есть операции, которых нет в книге — на «Главной» нажмите «Восстановить» (что это за операции — «Операции дня» → «Хвост журнала»)."
    ElseIf Left(ln, 4) = "FAIL" And (InStr(ln, "приходы и заказы") > 0 Or InStr(ln, "специальные приходы") > 0 Or InStr(ln, "возвраты") > 0 Or InStr(ln, "корректировки") > 0) Then
        AdviceFor = "Если выше красным стоит реестр «Наличие», это его следствие: сначала исправьте реестр. Иначе не сохраняйте книгу и верните её к копии: «Обслуживание» → «Вернуться к копии…»."
    ElseIf InStr(ln, "остаток = движения") > 0 Then
        AdviceFor = "Остаток ЕИ не сходится с движениями. Откройте «Сверка переноса» → «Обновить»: там названы ЕИ. Если реестр «Наличие» красный — это следствие его сдвига."
    ElseIf InStr(ln, "позиция журнала") > 0 Then
        AdviceFor = "Ничего делать не нужно: при следующем запуске WMS прочитает журнал заново. Не уходит — сделайте «Резервную копию» и перезапустите книгу."
    ElseIf InStr(ln, "ключи") > 0 Then
        AdviceFor = "Строка листа не согласована с журналом: не очищайте и не правьте ключи вручную. На строке нажмите «Почему?»; если строка скопирована — «Очистить»."
    ElseIf InStr(ln, "справочник «") > 0 And Left(ln, 4) = "WARN" Then
        AdviceFor = "Исправьте сокращения на листе «Получатели»: повторяющиеся не подставляются."
    ElseIf Left(ln, 4) = "FAIL" Then
        AdviceFor = "Не сохраняйте книгу. Сфотографируйте это окно и покажите ответственному за WMS; «Обслуживание» → «Вернуться к копии…» вернёт хорошее состояние."
    End If
End Function

' the text of the answer: the count and, under every FAIL and WARN of the self-check, what to do
Function WhatNowText() As String
    Dim rep As String, a As Variant, i As Long, nF As Long, nW As Long, s As String, adv As String, n As Long
    rep = WmsDiagnostics.SelfCheck()
    a = Split(rep, Chr(10))
    s = IIf(gState = "CLEAN", "Работа разрешена.", "ЗАБЛОКИРОВАНО (код " & gBlock & "): " & gBlockText & Chr(10) & WmsUi.BlockHint(gBlock))
    For i = 1 To UBound(a)
        If Left(a(i), 4) = "FAIL" Or Left(a(i), 4) = "WARN" Then
            If Left(a(i), 4) = "FAIL" Then nF = nF + 1 Else nW = nW + 1
            adv = AdviceFor(CStr(a(i)))
            n = n + 1
            If n <= 8 Then s = s & Chr(10) & Chr(10) & "• " & Left(a(i), 230) & IIf(adv <> "", Chr(10) & "   → " & adv, "")
        End If
    Next i
    If nF + nW = 0 Then
        s = s & Chr(10) & Chr(10) & "Самопроверка: ошибок нет, предупреждений нет — делать ничего не нужно."
    Else
        s = s & Chr(10) & Chr(10) & "Самопроверка: ошибок " & nF & ", предупреждений " & nW & IIf(n > 8, " (показаны первые 8)", "") & "."
    End If
    WhatNowText = s
End Function

Sub BtnWhatNow(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage(WhatNowText())
End Sub
