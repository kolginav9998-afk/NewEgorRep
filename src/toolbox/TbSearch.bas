' TbSearch — WMS_SEARCH: быстрый поиск по снимку и карточка ЕИ с историей (задание «FINAL WMS MARATHON», §6).
' Поиск: ЕИ, артикул, часть наименования, место, категория, источник, получатель. Карточка ЕИ: данные «Наличие» и
' история — пришло → пополнено → перемещено → выдано → возвращено → списано → корректировано (по датам). Только чтение.
Option Explicit

Private Const SH_FIND = "Поиск"
Private Const SH_CARD = "Карточка"
Private Const RES_ROW = 3              ' 0-based: the results start at row 4
Private Const MAX_RESULTS = 500

' «Обновить из снимка»: the tables of the latest snapshot into the hidden sheets
Sub BtnSearchRefresh(Optional oEvent As Variant)
    TbMsg(SearchRefresh())
End Sub

Function SearchRefresh() As String
    Dim why As String, snap As String, t As Variant, i As Integer, s As String
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        SearchRefresh = "ERR:" & why
        Exit Function
    End If
    t = Array("stock", "orders", "special", "issues", "returns", "adjustments")
    For i = 0 To UBound(t)
        s = s & IIf(s <> "", ", ", "") & t(i) & " " & TbLoadTable(snap, t(i) & ".csv", "_" & t(i), True)
    Next i
    TbPutSetting(3, snap)
    ThisComponent.Sheets.getByName(SH_FIND).getCellByPosition(0, 1).setString("Снимок: " & ConvertFromURL(snap) & " (LAST_SEQ " _
        & TbManifestValue(snap, "last_seq") & ")")
    SearchRefresh = "OK:" & s
    Exit Function
EH:
    SearchRefresh = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function Data(sName As String, nCols As Integer) As Variant
    Dim sh As Object, last As Long
    If Not ThisComponent.Sheets.hasByName(sName) Then
        Data = Array()
        Exit Function
    End If
    sh = ThisComponent.Sheets.getByName(sName)
    last = TbLastRow(sh)
    If last < 1 Then
        Data = Array()
        Exit Function
    End If
    Data = sh.getCellRangeByPosition(0, 1, nCols - 1, last).getDataArray()
End Function

Private Function Norm(ByVal s As String) As String
    Norm = LCase(Trim(Replace(Replace(s, Chr(160), " "), "ё", "е")))
End Function

' the canonical EI of a query that looks like an EI («12», «ЕИ-12», «ЕИ-00000012»), "" otherwise
Private Function AsEI(ByVal q As String) As String
    Dim t As String
    t = UCase(Trim(q))
    If Left(t, 2) = "ЕИ" Or Left(t, 2) = "EI" Then t = Mid(t, 3)
    If Left(t, 1) = "-" Then t = Mid(t, 2)
    If t <> "" And Len(t) <= 8 And IsNumeric(t) And InStr(t, ",") = 0 And InStr(t, ".") = 0 Then AsEI = "ЕИ-" & Right("00000000" & CLng(t), 8)
End Function

' «Найти»: query in B1 of «Поиск»; results from row 4
Sub BtnSearch(Optional oEvent As Variant)
    TbMsg(SearchRun(ThisComponent.Sheets.getByName(SH_FIND).getCellByPosition(1, 0).getString()))
End Sub

Function SearchRun(ByVal q As String) As String
    Dim sh As Object, st As Variant, iss As Variant, i As Long, out() As Variant, k As Long, nq As String, e As String, hay As String, last As Long
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_FIND)
    last = TbLastRow(sh)
    If last >= RES_ROW Then sh.getCellRangeByPosition(0, RES_ROW, 9, last).clearContents(1023)
    nq = Norm(q)
    If nq = "" Then
        SearchRun = "ERR:введите, что искать (ЕИ, артикул, часть наименования, место, категория, источник или получатель)"
        Exit Function
    End If
    e = AsEI(q)
    st = Data("_stock", 10)
    ReDim out(MAX_RESULTS)
    For i = 0 To UBound(st)
        If k > MAX_RESULTS Then Exit For
        If CStr(st(i)(0)) <> "" Then
            If e <> "" Then
                If st(i)(0) = e Then
                    out(k) = Array("Наличие", st(i)(0), st(i)(1), st(i)(2), st(i)(5), st(i)(6), st(i)(4), st(i)(9), "", st(i)(7))
                    k = k + 1
                End If
            Else
                hay = Norm(st(i)(0) & "|" & st(i)(1) & "|" & st(i)(2) & "|" & st(i)(5) & "|" & st(i)(6) & "|" & st(i)(8) & "|" & st(i)(9))
                If InStr(hay, nq) > 0 Then
                    out(k) = Array("Наличие", st(i)(0), st(i)(1), st(i)(2), st(i)(5), st(i)(6), st(i)(4), st(i)(9), "", st(i)(7))
                    k = k + 1
                End If
            End If
        End If
    Next i
    ' recipients: the issues of a person whose name contains the query
    If e = "" Then
        iss = Data("_issues", 18)
        For i = 0 To UBound(iss)
            If k > MAX_RESULTS Then Exit For
            If VarType(iss(i)(0)) = 5 Then
                If InStr(Norm(CStr(iss(i)(8))), nq) > 0 Then
                    out(k) = Array("Выдача № " & iss(i)(0), iss(i)(11), iss(i)(2), iss(i)(3), iss(i)(9), iss(i)(10), -iss(i)(4), iss(i)(8), _
                        IIf(VarType(iss(i)(7)) = 5, Format(CDate(iss(i)(7)), "DD.MM.YYYY"), CStr(iss(i)(7))), iss(i)(17))
                    k = k + 1
                End If
            End If
        Next i
    End If
    If k > 0 Then
        Dim rows() As Variant, j As Long
        ReDim rows(IIf(k > MAX_RESULTS, MAX_RESULTS, k) - 1)
        For j = 0 To UBound(rows)
            rows(j) = out(j)
        Next j
        sh.getCellRangeByPosition(0, RES_ROW, 9, RES_ROW + UBound(rows)).setDataArray(rows)
    End If
    SearchRun = "OK:найдено " & k & IIf(k > MAX_RESULTS, " (показаны первые " & MAX_RESULTS & ")", "")
    Exit Function
EH:
    SearchRun = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' «Показать карточку»: the EI in B1 of «Карточка» — its data and history (sorted by date, then by the kind)
Sub BtnCard(Optional oEvent As Variant)
    TbMsg(CardShow(ThisComponent.Sheets.getByName(SH_CARD).getCellByPosition(1, 0).getString()))
End Sub

Function CardShow(ByVal q As String) As String
    Dim sh As Object, e As String, st As Variant, i As Long, n As Long, h() As Variant, k As Long, t As Variant, last As Long
    Dim o As Variant, sp As Variant, iss As Variant, rt As Variant, ad As Variant, j As Long, tmp As Variant, card As Variant
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_CARD)
    last = TbLastRow(sh)
    If last >= 2 Then sh.getCellRangeByPosition(0, 2, 7, last).clearContents(1023)
    e = AsEI(q)
    If e = "" Then
        CardShow = "ERR:«" & q & "» — не номер ЕИ"
        Exit Function
    End If
    st = Data("_stock", 10)
    n = CLng(Mid(e, 4))
    If n > UBound(st) + 1 Then
        CardShow = "ERR:" & e & " нет в снимке"
        Exit Function
    End If
    card = st(n - 1)
    If CStr(card(0)) <> e Then
        CardShow = "ERR:" & e & " нет в снимке"
        Exit Function
    End If
    sh.getCellRangeByPosition(0, 2, 1, 11).setDataArray(Array(Array("ЕИ", card(0)), Array("Наименование", card(1)), Array("Артикул", card(2)), _
        Array("Единица", card(3)), Array("Остаток", card(4)), Array("Место", card(5)), Array("Категория", card(6)), Array("Состояние", card(7)), _
        Array("Источник", card(8)), Array("Тип источника", card(9))))
    ReDim h(1000)
    o = Data("_orders", 28)
    For i = 0 To UBound(o)
        If CStr(o(i)(21)) = e Then AddH(h, k, o(i)(13), "Пришло (заказ " & o(i)(0) & ")", o(i)(5), "", o(i)(11), o(i)(22))
    Next i
    sp = Data("_special", 19)
    For i = 0 To UBound(sp)
        If CStr(sp(i)(13)) = e And VarType(sp(i)(0)) = 5 Then
            AddH(h, k, sp(i)(7), IIf(InStr(CStr(sp(i)(16)), "пополнение") > 0, "Пополнено (", "Пришло (") & sp(i)(1) & " " & sp(i)(2) & ")", sp(i)(5), _
                sp(i)(8), sp(i)(10), sp(i)(16))
        End If
    Next i
    iss = Data("_issues", 18)
    For i = 0 To UBound(iss)
        If CStr(iss(i)(11)) = e And VarType(iss(i)(0)) = 5 Then AddH(h, k, iss(i)(7), "Выдано (№ " & iss(i)(0) & ")", -iss(i)(4), "", iss(i)(8), iss(i)(17))
    Next i
    rt = Data("_returns", 15)
    For i = 0 To UBound(rt)
        If CStr(rt(i)(2)) = e And VarType(rt(i)(0)) = 5 Then AddH(h, k, rt(i)(7), "Возвращено (№ " & rt(i)(0) & ", выдача № " & rt(i)(1) & ")", rt(i)(5), _
            rt(i)(9), rt(i)(8), rt(i)(13))
    Next i
    ad = Data("_adjustments", 19)
    For i = 0 To UBound(ad)
        If CStr(ad(i)(2)) = e And VarType(ad(i)(0)) = 5 Then
            Select Case CStr(ad(i)(1))
            Case "Перемещение"
                AddH(h, k, ad(i)(10), "Перемещено (№ " & ad(i)(0) & ")", 0, ad(i)(12) & " → " & ad(i)(9), ad(i)(11), ad(i)(16))
            Case "Списание"
                AddH(h, k, ad(i)(10), "Списано (№ " & ad(i)(0) & ")", ad(i)(15), "", ad(i)(11), ad(i)(16))
            Case Else
                AddH(h, k, ad(i)(10), "Корректировано (№ " & ad(i)(0) & ")", ad(i)(15), "", ad(i)(11), ad(i)(16))
            End Select
        End If
    Next i
    ' by date (insertion sort: a history is short)
    For i = 1 To k - 1
        tmp = h(i)
        j = i - 1
        Do While j >= 0
            If CStr(h(j)(0)) <= CStr(tmp(0)) Then Exit Do
            h(j + 1) = h(j)
            j = j - 1
        Loop
        h(j + 1) = tmp
    Next i
    sh.getCellRangeByPosition(3, 2, 8, 2).setDataArray(Array(Array("Дата", "Событие", "Количество", "Место", "Кто / документ", "Статус")))
    If k > 0 Then
        Dim rows() As Variant
        ReDim rows(k - 1)
        For i = 0 To k - 1
            rows(i) = h(i)
        Next i
        sh.getCellRangeByPosition(3, 3, 8, 3 + k - 1).setDataArray(rows)
    End If
    CardShow = "OK:" & e & ": событий в истории " & k
    Exit Function
EH:
    CardShow = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' one event of the history; the date as ГГГГ-ММ-ДД (it sorts and reads the same)
Private Sub AddH(h As Variant, k As Long, dt As Variant, ev As String, q As Variant, pl As Variant, who As Variant, stt As Variant)
    Dim sd As String
    If k > UBound(h) Then Exit Sub
    If VarType(dt) = 5 Or VarType(dt) = 7 Then sd = Format(CDate(dt), "YYYY-MM-DD") Else sd = CStr(dt)
    h(k) = Array(sd, ev, q, CStr(pl), CStr(who), CStr(stt))
    k = k + 1
End Sub

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function SearchProbe() As String
    SearchProbe = "OK"
End Function
