' TbSearch — WMS_SEARCH: быстрый поиск по снимку и карточка ЕИ с историей (задание «FINAL WMS MARATHON», §6; M6 PRIME §7).
' Поиск: ЕИ, артикул, часть наименования, место, категория, источник, получатель, поставщик (его заказы и машины), госномер
' или марка машины (её визиты). Карточка ЕИ: данные «Наличие» (остаток, место, происхождение) и история — пришло →
' пополнено → перемещено → выдано → возвращено → списано → корректировано (по датам). Лист «Поставщик» — все заказы и
' транспортные визиты поставщика; лист «Машина» — вся история приездов по госномеру в любом написании. Только чтение.
Option Explicit

Private Const SH_FIND = "Поиск"
Private Const SH_CARD = "Карточка"
Private Const SH_SUP = "Поставщик"
Private Const SH_VEH = "Машина"
Private Const RES_ROW = 3              ' 0-based: the results start at row 4
Private Const MAX_RESULTS = 500

' «Обновить из снимка»: the tables of the latest snapshot into the hidden sheets
Sub BtnSearchRefresh(Optional oEvent As Variant)
    TbMsg(SearchRefresh())
End Sub

Function SearchRefresh() As String
    Dim why As String, snap As String, t As Variant, i As Integer, s As String, n As Long
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
    n = TbLoadOptional(snap, "cars.csv", "_cars", True, TB_CAR_KEYS)
    s = s & ", cars " & IIf(n < 0, "нет в снимке (WMS до 0.7)", CStr(n))
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
    ' orders: the supplier or the number of the order
    Dim o As Variant, cars As Variant, pk As String, nSup As Long, supName As String
    If e = "" Then
        o = Data("_orders", 28)
        For i = 0 To UBound(o)
            If k > MAX_RESULTS Then Exit For
            If CStr(o(i)(1)) <> "" Or CStr(o(i)(0)) <> "" Then
                If InStr(Norm(CStr(o(i)(11))), nq) > 0 Or Norm(CStr(o(i)(0))) = nq Then
                    out(k) = Array("Заказ " & o(i)(0), o(i)(21), o(i)(1), o(i)(4), o(i)(20), o(i)(18), IIf(CStr(o(i)(5)) <> "", o(i)(5), o(i)(7)), o(i)(11), _
                        DateText(IIf(CStr(o(i)(13)) <> "", o(i)(13), o(i)(16))), o(i)(22))
                    k = k + 1
                    If InStr(Norm(CStr(o(i)(11))), nq) > 0 Then
                        nSup = nSup + 1
                        supName = CStr(o(i)(11))
                    End If
                End If
            End If
        Next i
    End If
    ' vehicles: the plate in any spelling (or its digits), the make, the supplier
    cars = Data("_cars", 14)
    pk = TbPlateKey(q)
    For i = 0 To UBound(cars)
        If k > MAX_RESULTS Then Exit For
        If VarType(cars(i)(0)) = 5 Then
            If (Len(pk) >= 3 And InStr(1, CStr(cars(i)(13)), pk, 0) > 0) Or (e = "" And InStr(Norm(CStr(cars(i)(2)) & "|" & CStr(cars(i)(3))), nq) > 0) Then
                out(k) = Array("Визит № " & cars(i)(0), "", cars(i)(2), "", "", "", IIf(CStr(cars(i)(7)) = "Уехал", cars(i)(6), ""), cars(i)(3), _
                    DateTimeText(cars(i)(4)), cars(i)(7))
                k = k + 1
            End If
        End If
    Next i
    ' a supplier found: its name goes to «Поставщик» (all its orders and vehicles there)
    If nSup > 0 Then ThisComponent.Sheets.getByName(SH_SUP).getCellByPosition(1, 0).setString(supName)
    If k > 0 Then
        Dim rows() As Variant, j As Long
        ReDim rows(IIf(k > MAX_RESULTS, MAX_RESULTS, k) - 1)
        For j = 0 To UBound(rows)
            rows(j) = out(j)
        Next j
        sh.getCellRangeByPosition(0, RES_ROW, 9, RES_ROW + UBound(rows)).setDataArray(rows)
    End If
    SearchRun = "OK:найдено " & k & IIf(k > MAX_RESULTS, " (показаны первые " & MAX_RESULTS & ")", "") _
        & IIf(nSup > 0, "; поставщик «" & supName & "»: все его заказы и машины — лист «" & SH_SUP & "», кнопка «Заказы и машины поставщика»", "")
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

Private Function DateText(v As Variant) As String
    If VarType(v) = 5 Or VarType(v) = 7 Then DateText = Format(CDate(v), "DD.MM.YYYY") Else DateText = CStr(v)
End Function

Private Function DateTimeText(v As Variant) As String
    If VarType(v) = 5 Or VarType(v) = 7 Then DateTimeText = Format(CDate(v), "DD.MM.YYYY HH:MM") Else DateTimeText = CStr(v)
End Function

' ================================================================ «Поставщик» and «Машина»

Private Sub Title2(sh As Object, r As Long, s As String)
    sh.getCellByPosition(0, r).setString(s)
    sh.getCellByPosition(0, r).CharWeight = 150
End Sub

Private Sub Heads(sh As Object, r As Long, h As Variant)
    Dim rng As Object
    rng = sh.getCellRangeByPosition(0, r, UBound(h), r)
    rng.setDataArray(Array(h))
    rng.CharWeight = 150
    rng.CellBackColor = RGB(231, 230, 230)
End Sub

Private Sub ClearFrom(sh As Object, r0 As Long, nCols As Integer)
    Dim last As Long
    last = TbLastRow(sh)
    If last >= r0 Then
        sh.getCellRangeByPosition(0, r0, nCols - 1, last).clearContents(1023)
        sh.getCellRangeByPosition(0, r0, nCols - 1, last).CellBackColor = -1
    End If
End Sub

Sub BtnSupplier(Optional oEvent As Variant)
    TbMsg(SupplierShow(ThisComponent.Sheets.getByName(SH_SUP).getCellByPosition(1, 0).getString()))
End Sub

' «Заказы и машины поставщика»: every row of «Заказы» of the supplier (the name contains the query, any case) and every
' visit of its vehicles; a summary line
Function SupplierShow(ByVal q As String) As String
    Dim sh As Object, o As Variant, cars As Variant, i As Long, r As Long, nq As String, k As Long, nOver As Long, nWait As Long, nV As Long, nCl As Long
    Dim sumD As Double, names As String, row As Variant, avg As String
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_SUP)
    ClearFrom(sh, 1, 9)
    nq = Norm(q)
    If nq = "" Then
        SupplierShow = "ERR:введите поставщика (часть названия) в B1"
        Exit Function
    End If
    o = Data("_orders", 28)
    If UBound(o) < 0 And Not ThisComponent.Sheets.hasByName("_orders") Then
        SupplierShow = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    r = 3
    Title2(sh, r, "ЗАКАЗЫ ПОСТАВЩИКА")
    r = r + 1
    Heads(sh, r, Array("№ заказа", "Наименование", "Поставщик", "ЕИ", "Кол-во (факт / заказано)", "Дата заказа", "Ожидается", "Получено", "Статус"))
    r = r + 1
    For i = 0 To UBound(o)
        If CStr(o(i)(11)) <> "" And InStr(Norm(CStr(o(i)(11))), nq) > 0 Then
            row = Array(CStr(o(i)(0)), CStr(o(i)(1)), CStr(o(i)(11)), CStr(o(i)(21)), IIf(CStr(o(i)(5)) <> "", o(i)(5), o(i)(7)), DateText(o(i)(15)), _
                DateText(o(i)(16)), DateText(o(i)(13)), CStr(o(i)(22)))
            sh.getCellRangeByPosition(0, r, 8, r).setDataArray(Array(row))
            r = r + 1
            k = k + 1
            If InStr(1, "|" & names & "|", "|" & o(i)(11) & "|", 0) = 0 Then names = names & IIf(names <> "", "|", "") & o(i)(11)
            If TbDaysLate(CStr(o(i)(22)), o(i)(16), Int(CDbl(Now()))) > 0 Then nOver = nOver + 1
            If CStr(o(i)(22)) = "Ожидается" Then nWait = nWait + 1
        End If
    Next i
    If k = 0 Then
        sh.getCellByPosition(0, r).setString("нет")
        r = r + 1
    End If
    r = r + 1
    Title2(sh, r, "ТРАНСПОРТНЫЕ ВИЗИТЫ ПОСТАВЩИКА («Приход авто»)")
    r = r + 1
    Heads(sh, r, Array("№ визита", "Машина", "Поставщик", "Приезд", "Выезд", "Минут", "Статус"))
    r = r + 1
    cars = Data("_cars", 14)
    For i = 0 To UBound(cars)
        If VarType(cars(i)(0)) = 5 And InStr(Norm(CStr(cars(i)(3))), nq) > 0 Then
            sh.getCellRangeByPosition(0, r, 6, r).setDataArray(Array(Array(cars(i)(0), CStr(cars(i)(2)), CStr(cars(i)(3)), DateTimeText(cars(i)(4)), _
                DateTimeText(cars(i)(5)), IIf(CStr(cars(i)(7)) = "Уехал", cars(i)(6), ""), CStr(cars(i)(7)))))
            r = r + 1
            If CStr(cars(i)(7)) <> "Отменён" Then nV = nV + 1
            If CStr(cars(i)(7)) = "Уехал" And VarType(cars(i)(6)) = 5 Then
                nCl = nCl + 1
                sumD = sumD + cars(i)(6)
            End If
        End If
    Next i
    If nV = 0 And UBound(cars) < 0 Then sh.getCellByPosition(0, r).setString(IIf(ThisComponent.Sheets.hasByName("_cars"), "нет", "нет данных транспорта в снимке"))
    ' (IIf of Basic computes both branches: the mean only when there are closed visits)
    If nCl > 0 Then avg = ", средняя стоянка " & Format(sumD / nCl, "0") & " мин"
    sh.getCellByPosition(0, 1).setString("Заказов (строк) " & k & ": ожидается " & nWait & ", ожидаемая дата прошла " & nOver & "; визитов машин " & nV _
        & avg & IIf(names <> "", " — поставщики: " & Replace(names, "|", ", "), ""))
    SupplierShow = "OK:заказов " & k & ", визитов " & nV
    Exit Function
EH:
    SupplierShow = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnVehicle(Optional oEvent As Variant)
    TbMsg(VehicleShow(ThisComponent.Sheets.getByName(SH_VEH).getCellByPosition(1, 0).getString()))
End Sub

' «История приездов»: the visits of a vehicle by its plate in any spelling («а 123 вс», «A123BC 77»); without a plate in
' the query — the visits whose make / plate text contains it
Function VehicleShow(ByVal q As String) As String
    Dim sh As Object, cars As Variant, i As Long, r As Long, pk As String, nq As String, k As Long, nOpen As Long, first As String, lastV As String
    Dim sups As String, bExact As Boolean
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_VEH)
    ClearFrom(sh, 1, 8)
    nq = Norm(q)
    If nq = "" Then
        VehicleShow = "ERR:введите госномер (или марку) в B1"
        Exit Function
    End If
    If Not ThisComponent.Sheets.hasByName("_cars") Then
        VehicleShow = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    pk = TbPlateKey(q)
    cars = Data("_cars", 14)
    ' an exact key (the plate) first; otherwise a part of the text
    For i = 0 To UBound(cars)
        If VarType(cars(i)(0)) = 5 And CStr(cars(i)(13)) = pk Then bExact = True
    Next i
    r = 3
    Title2(sh, r, "ИСТОРИЯ ПРИЕЗДОВ" & IIf(bExact, " (госномер " & pk & ")", " (текст «" & q & "»)"))
    r = r + 1
    Heads(sh, r, Array("№ визита", "Дата", "Машина (как записано)", "Поставщик", "Приезд", "Выезд", "Минут", "Статус"))
    r = r + 1
    For i = 0 To UBound(cars)
        If VarType(cars(i)(0)) = 5 Then
            If (bExact And CStr(cars(i)(13)) = pk) Or (Not bExact And (InStr(Norm(CStr(cars(i)(2))), nq) > 0 Or (Len(pk) >= 3 And InStr(1, CStr(cars(i)(13)), pk, 0) > 0))) Then
                sh.getCellRangeByPosition(0, r, 7, r).setDataArray(Array(Array(cars(i)(0), DateText(cars(i)(1)), CStr(cars(i)(2)), CStr(cars(i)(3)), _
                    DateTimeText(cars(i)(4)), DateTimeText(cars(i)(5)), IIf(CStr(cars(i)(7)) = "Уехал", cars(i)(6), ""), CStr(cars(i)(7)))))
                r = r + 1
                If CStr(cars(i)(7)) <> "Отменён" Then
                    k = k + 1
                    If first = "" Then first = DateTimeText(cars(i)(4))
                    lastV = DateTimeText(cars(i)(4))
                    If CStr(cars(i)(7)) = "На территории" Then nOpen = nOpen + 1
                    If InStr(1, "|" & sups & "|", "|" & cars(i)(3) & "|", 0) = 0 Then sups = sups & IIf(sups <> "", "|", "") & cars(i)(3)
                End If
            End If
        End If
    Next i
    If k = 0 Then
        sh.getCellByPosition(0, r).setString("визитов нет")
        sh.getCellByPosition(0, 1).setString("Визитов нет")
    Else
        sh.getCellByPosition(0, 1).setString("Визитов " & k & ": первый " & first & ", последний " & lastV & "; поставщики: " & Replace(sups, "|", ", ") _
            & IIf(nOpen > 0, "; СЕЙЧАС НА ТЕРРИТОРИИ", ""))
    End If
    VehicleShow = "OK:визитов " & k & IIf(nOpen > 0, ", на территории", "")
    Exit Function
EH:
    VehicleShow = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
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
