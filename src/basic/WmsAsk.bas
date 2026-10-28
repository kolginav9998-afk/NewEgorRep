' WmsAsk — 1.0: the questions asked before a posting when a value is possible but unusual (the team found them missing):
' a date in the future or more than a year back, a recipient not in the directory «Получатели» (with the names that look
' alike). The question never changes the books: it returns the value to post or «not posted». Library WmsPlus.
Option Explicit

' RetColumnsHint was shown in this session (Global: a sheet event is a macro call of its own)
Global gAskRetHint As Boolean
' 1.0: a block of rows posted at once (WmsDaily): the answers given to a question once, «|key=answer|…», reused for the rest
Global gAskMemoOn As Boolean
Global gAskMemo As String

Sub AskMemoBegin()
    gAskMemoOn = True
    gAskMemo = "|"
End Sub

Sub AskMemoEnd()
    gAskMemoOn = False
    gAskMemo = ""
End Sub

' the answer remembered for key in the block ("" — none)
Private Function MemoGet(sKey As String) As String
    Dim p As Long, q As Long
    If Not gAskMemoOn Then Exit Function
    p = InStr(gAskMemo, "|" & sKey & "=")
    If p = 0 Then Exit Function
    p = p + Len(sKey) + 2
    q = InStr(p, gAskMemo, "|")
    MemoGet = Mid(gAskMemo, p, q - p)
End Function

Private Sub MemoPut(sKey As String, sAnswer As String)
    If gAskMemoOn Then gAskMemo = gAskMemo & sKey & "=" & Replace(sAnswer, "|", "/") & "|"
End Sub

' 1.0 (К13): the answer of a block for another module (the reason of a posting in a closed month); "" — none or no block
Function MemoValue(sKey As String) As String
    MemoValue = MemoGet(sKey)
End Function

Sub MemoSet(sKey As String, sAnswer As String)
    MemoPut(sKey, sAnswer)
End Sub

' the line under a question of a block: the answer counts for the other rows of the block
Private Function MemoNote() As String
    If gAskMemoOn Then MemoNote = Chr(10) & Chr(10) & "Ответ будет применён и к остальным строкам этой пачки."
End Function

' ================================================================ a big operation (1.1)

' An issue or a write-off of a big quantity (from the setting «Большая операция» of the sheet «Настройки», 1000 by default; 0 — never asked) is asked twice, the second time with
' the quantity written in words of the action: a quantity typed with an extra zero is the mistake that costs the most. 0 — post, 1 — not posted (msg). A block of rows: every
' big row is asked, the answer of one is not given to the others
Function BigAsk(sVerb As String, sName As String, dQty As Double, sUnit As String, ByRef msg As String) As Integer
    Dim thr As Double, t As String
    thr = WmsSettings.SettingNum(SET_BIG_QTY, 1000)
    If thr <= 0 Or dQty < thr - 0.0000001 Then Exit Function
    t = sVerb & " " & WmsIssue.QtyText(dQty) & " " & sUnit & " «" & sName & "»"
    If Not WmsUi.UiConfirm("Большая операция: " & t & "." & Chr(10) & Chr(10) & "Это не меньше " & WmsIssue.QtyText(thr) & " " & sUnit & " (порог — на листе «" & SH_SET & "»). Количество верное?") Then
        msg = "большая операция не подтверждена (" & t & ") — проверьте количество в строке"
        BigAsk = 1
        Exit Function
    End If
    If Not WmsUi.UiConfirm("Ещё раз, чтобы не ошибиться:" & Chr(10) & Chr(10) & t & "." & Chr(10) & Chr(10) & "Провести это количество?") Then
        msg = "большая операция не подтверждена (" & t & ") — проверьте количество в строке"
        BigAsk = 1
    End If
End Function

' ================================================================ the date of a movement

' d — the date of the movement (a Calc date), sWhat — what it is («выдачи», «возврата»), sVerb — the action on the button
' («Провести»). 0 — post with d, 1 — not posted (msg)
Function DateAsk(d As Double, sWhat As String, sVerb As String, ByRef msg As String) As Integer
    Dim today As Double, sText As String, k As Integer, sMemo As String
    today = CDbl(Date())
    If d > today Then
        sText = "Дата " & sWhat & " " & Format(d, "DD.MM.YYYY") & " — в будущем (сегодня " & Format(today, "DD.MM.YYYY") & ")." & Chr(10) _
            & "Так бывает, только если операция оформляется заранее. Если дата набрана с ошибкой — исправьте её в строке."
    ElseIf d < today - 366 Then
        sText = "Дата " & sWhat & " " & Format(d, "DD.MM.YYYY") & " — больше года назад (сегодня " & Format(today, "DD.MM.YYYY") & ")." & Chr(10) _
            & "Проверьте год. Если дата набрана с ошибкой — исправьте её в строке."
    Else
        Exit Function
    End If
    sMemo = MemoGet("D" & CLng(d))
    If sMemo <> "" Then
        k = CInt(sMemo)
    Else
        k = WmsPrimeUi.UiButtons("WMS — проверьте дату", sText & MemoNote(), Array(sVerb & " с датой " & Format(d, "DD.MM.YYYY"), "Не проводить — исправлю дату"), _
            Array(0, 1))
        MemoPut("D" & CLng(d), CStr(IIf(k = 0, 0, 1)))
    End If
    If k <> 0 Then
        msg = "проверьте дату " & Format(d, "DD.MM.YYYY")
        DateAsk = 1
    End If
End Function

' ================================================================ the recipient

' a name as a key: lower case, «ё» as «е», dots and repeated spaces dropped («Петров И.И.» = «петров и и»)
Function WhoKey(ByVal s As String) As String
    Dim i As Long, ch As String, r As String, sp As Boolean
    s = LCase(Trim(s))
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "ё" Then ch = "е"
        If ch = "." Or ch = "," Or ch = Chr(160) Or ch = Chr(9) Then ch = " "
        If ch = " " Then
            If Not sp And r <> "" Then r = r & " "
            sp = True
        Else
            r = r & ch
            sp = False
        End If
    Next i
    WhoKey = Trim(r)
End Function

' the edit distance of two short keys, stopped above lim (lim + 1 is returned)
Private Function AskDist(a As String, b As String, lim As Integer) As Integer
    Dim la As Integer, lb As Integer, i As Integer, j As Integer, c As Integer, prev() As Integer, cur() As Integer, best As Integer
    la = Len(a)
    lb = Len(b)
    If Abs(la - lb) > lim Then
        AskDist = lim + 1
        Exit Function
    End If
    ReDim prev(lb)
    ReDim cur(lb)
    For j = 0 To lb
        prev(j) = j
    Next j
    For i = 1 To la
        cur(0) = i
        best = i
        For j = 1 To lb
            If Mid(a, i, 1) = Mid(b, j, 1) Then c = 0 Else c = 1
            cur(j) = prev(j - 1) + c
            If prev(j) + 1 < cur(j) Then cur(j) = prev(j) + 1
            If cur(j - 1) + 1 < cur(j) Then cur(j) = cur(j - 1) + 1
            If cur(j) < best Then best = cur(j)
        Next j
        If best > lim Then
            AskDist = lim + 1
            Exit Function
        End If
        For j = 0 To lb
            prev(j) = cur(j)
        Next j
    Next i
    AskDist = prev(lb)
End Function

' up to three names of the directory that look like s: the same key, one starts with the other, or 1–2 letters apart
Function WhoLike(s As String) As Variant
    Dim k As String, i As Long, n As Integer, r(2) As String, kk As String, lim As Integer, res() As String
    k = WhoKey(s)
    If Not gRcLoaded Then WmsIssue.RecipientsCount()
    lim = IIf(Len(k) <= 5, 1, 2)
    For i = 0 To gRcN - 1
        If n > 2 Then Exit For
        kk = WhoKey(gRcName(i))
        If kk = k Or (Len(k) >= 3 And Left(kk, Len(k)) = k) Or (Len(kk) >= 3 And Left(k, Len(kk)) = kk) Then
            r(n) = gRcName(i)
            n = n + 1
        ElseIf AskDist(k, kk, lim) <= lim Then
            r(n) = gRcName(i)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        WhoLike = Array()
        Exit Function
    End If
    ReDim res(n - 1)
    For i = 0 To n - 1
        res(i) = r(i)
    Next i
    WhoLike = res
End Function

' a recipient not in the directory (ResolveRecipient 2): «Записать как есть», one of the names that look alike, «Не проводить».
' sWho — in: as typed, res: the name to post. 0 — post sWho, 1 — not posted (msg)
Function WhoAsk(ByRef sWho As String, ByRef msg As String) As Integer
    Dim aLike As Variant, labels() As String, codes() As Integer, i As Integer, n As Integer, k As Integer, sText As String, sMemo As String
    ' a block: the same name answered once
    sMemo = MemoGet("W" & WhoKey(sWho))
    If sMemo = "0" Then
        Exit Function
    ElseIf sMemo = "1" Then
        msg = "проверьте получателя «" & Trim(sWho) & "»"
        WhoAsk = 1
        Exit Function
    ElseIf Left(sMemo, 1) = "=" Then
        sWho = Mid(sMemo, 2)
        Exit Function
    End If
    aLike = WhoLike(sWho)
    n = UBound(aLike) + 1
    ReDim labels(n + 1)
    ReDim codes(n + 1)
    sText = "Получателя «" & Trim(sWho) & "» нет в справочнике «" & SH_RCPT & "»."
    If n > 0 Then
        sText = sText & Chr(10) & "Похожие: " & Join(aLike, ", ") & "." & Chr(10) & "Если это опечатка, выберите правильное имя — иначе «на руках у …» разделится на два человека."
    Else
        sText = sText & Chr(10) & "Новый получатель? Его можно добавить на лист «" & SH_RCPT & "», чтобы потом выбирать из списка."
    End If
    For i = 0 To n - 1
        labels(i) = "Выдать: " & aLike(i)
        codes(i) = 10 + i
    Next i
    labels(n) = "Записать «" & Trim(sWho) & "» как есть"
    codes(n) = 0
    labels(n + 1) = "Не проводить"
    codes(n + 1) = 1
    ' the test seam answers the first button: the name as typed is not replaced by a guess
    If gUiAuto <> 0 And Not gPrChoiceSet Then
        WhoAsk = 0
        Exit Function
    End If
    k = WmsPrimeUi.UiButtons("WMS — получатель", sText & MemoNote(), labels, codes)
    sMemo = "W" & WhoKey(sWho)
    If k >= 10 Then
        sWho = aLike(k - 10)
        MemoPut(sMemo, "=" & sWho)
    ElseIf k <> 0 Then
        msg = "проверьте получателя «" & Trim(sWho) & "»"
        WhoAsk = 1
        MemoPut(sMemo, "1")
    Else
        MemoPut(sMemo, "0")
    End If
End Function

' ================================================================ the result of «Провести» for one row

' 1.0: one row posted by its button and not posted — the reason in a window (it was only in «Контроль», often beyond the
' edge of the screen). sNot — «Выдача не проведена», «Приход не проведён»… Nothing when the storekeeper chose «Не
' проводить» in a question (a SKIP that says «не проведен…»), nothing on success.
Sub ShowRowResult(sNot As String, res As String)
    If Left(res, 4) = "ERR:" Then
        WmsUi.UiMessage(sNot & ":" & Chr(10) & Chr(10) & Mid(res, 5))
    ElseIf Left(res, 5) = "SKIP:" Then
        If InStr(res, "не проведен") = 0 And InStr(res, "не подтвержден") = 0 And InStr(res, "не исправлен") = 0 Then WmsUi.UiMessage(sNot & ":" & Chr(10) & Chr(10) & Mid(res, 6))
    End If
End Sub

' ================================================================ «Вернул» / «Вернул в %» on «Выдачи» (D-004)

' something was typed into M or N of «Выдачи»: these are the notes of the old table (D-004) — WMS neither reads nor fills
' them. Explained once per session, only when a value appeared (clearing the cells says nothing)
Sub RetColumnsHint(ra As Variant)
    Dim d As Variant, i As Long, j As Long, c0 As Integer, c1 As Integer, bAny As Boolean, sh As Object
    On Error GoTo EH
    If gAskRetHint Then Exit Sub
    WmsInit()
    sh = WmsIssue.IssuesSheet()
    If ra.EndRow - ra.StartRow > 200 Then Exit Sub
    c0 = IIf(ra.StartColumn > IC_RET, ra.StartColumn, IC_RET)
    c1 = IIf(ra.EndColumn < IC_RETPCT, ra.EndColumn, IC_RETPCT)
    If c1 < c0 Then Exit Sub
    d = sh.getCellRangeByPosition(c0, ra.StartRow, c1, ra.EndRow).getDataArray()
    For i = 0 To UBound(d)
        For j = 0 To UBound(d(i))
            If CStr(d(i)(j)) <> "" Then bAny = True
        Next j
    Next i
    If Not bAny Then Exit Sub
    gAskRetHint = True
    WmsUi.UiMessage("Колонки «Вернул» и «Вернул в %» — для заметок (так было в старой таблице): WMS их не учитывает, остаток от них не меняется." _
        & Chr(10) & Chr(10) & "Возврат оформляется на листе «" & SH_RETURNS & "»: № выдачи, сколько вернули и дата — затем «Провести». " _
        & "Сколько уже вернули по выдаче, видно в карточке ЕИ на листе «Поиск».")
    Exit Sub
EH:
End Sub
