' WmsNav — 1.4.0: simple ways through the book. «Я хочу…» — the daily tasks in the words of the storekeeper (a short list; each line leads to the window or the sheet that does it)
' and «Все листы…» — every sheet of the book by its purpose, with a line of what it is for (first the purpose, then the sheet).
' Nothing is written here that the book does not write by its own buttons: the lines lead to the receipt window, the issue window, «Закрыть день» and the navigation of a sheet.
' Library WmsPlus.
Option Explicit

' ================================================================ «Я хочу…»

' the tasks, in the order of the day (the title of the window is «я хочу…», so each line goes on from it)
Function WantList() As Variant
    WantList = Array("принять поставку по заказу", "принять товар без заказа", "выдать со склада", "принять возврат от человека", "найти товар, ЕИ или человека", _
        "узнать, где лежит и куда положить", "исправить ошибку", "закрыть день", "посмотреть отчёты", "открыть другой лист…")
End Function

Sub BtnWant(Optional oEvent As Variant)
    Dim k As Integer
    WmsInit()
    k = WmsPrimeUi.UiChoose("WMS — я хочу…", "Выберите, что нужно сделать — WMS откроет нужное окно или лист:", WantList())
    If k >= 0 Then WantRun(k)
End Sub

' the task k of the list (0-based): what was done — OK:… / SKIP:… / ERR:… (a test seam as well)
Function WantRun(k As Integer) As String
    Dim s As String, j As Integer
    WmsInit()
    Select Case k
    Case 0
        s = NvReceive()
    Case 1
        WmsUi.NavSpecial()
        s = "OK:" & SH_SPECIAL
    Case 2
        s = WmsDaily.IssueWindow()
        If Left(s, 3) = "OK:" Then WmsUi.NavIssues()
    Case 3
        j = WmsPrimeUi.UiButtons("Принять возврат", "Как найти выдачу, по которой вернули?" & Chr(10) _
            & "«Поиск»: впишите имя человека, выделите строки выдачи и нажмите «Вернуть выделенное»." & Chr(10) _
            & "«Возврат»: впишите № выдачи и сколько вернули, затем «Провести».", Array("Открыть «Поиск» (по человеку)", "Открыть «Возврат» (по № выдачи)", "Отмена"), Array(1, 2, 0))
        If j = 1 Then
            WmsUi.NavFind()
            s = "OK:" & SH_FIND
        ElseIf j = 2 Then
            WmsUi.NavReturns()
            s = "OK:" & SH_RETURNS
        Else
            s = "SKIP:отмена"
        End If
    Case 4
        WmsUi.NavFind()
        s = "OK:" & SH_FIND
    Case 5
        WmsUi.NavMap()
        s = "OK:" & SH_MAP
    Case 6
        WmsHelp.BtnWhatIf()
        s = "OK:что делать, если…"
    Case 7
        WmsDocs.BtnCloseDay()
        s = "OK:закрыть день"
    Case 8
        WmsLedger.NavReports()
        s = "OK:Отчёты"
    Case 9
        s = SheetsMenu()
    Case Else
        s = "ERR:нет такого пункта"
    End Select
    gUiLastMsg = s
    WantRun = s
End Function

' «принять поставку по заказу»: the positions still expected (the oldest first) in a list; the chosen one is put under the cursor of «Заказы» and its receipt window opens
Private Function NvReceive() As String
    Dim res As String, i As Long, n As Long, sh As Object, d As Variant, opts() As String, rws() As Long, k As Integer, s As String
    res = WmsSupply.OpenPositions("")
    If Left(res, 4) = "ERR:" Then
        WmsUi.NavOrders()
        WmsUi.UiMessage("Список ожидаемых позиций не построен: " & Mid(res, 5) & Chr(10) & "На листе «" & SH_ORDERS & "» встаньте на строку заказа и нажмите «Принять…».")
        NvReceive = "ERR:" & Mid(res, 5)
        Exit Function
    End If
    If gSpN = 0 Then
        WmsUi.NavOrders()
        WmsUi.UiMessage("Ожидаемых позиций в заказах нет: всё заказанное уже принято или заказ ещё не записан на лист «" & SH_ORDERS & "»." & Chr(10) _
            & "Товар без заказа принимают через «" & SH_SPECIAL & "».")
        NvReceive = "SKIP:ожидаемых позиций нет"
        Exit Function
    End If
    sh = WmsOrders.OrdersSheet()
    n = gSpN
    If n > 15 Then n = 15
    ReDim opts(n)
    ReDim rws(n)
    For i = 0 To n - 1
        rws(i) = gSpRow(i)
        d = sh.getCellRangeByPosition(0, rws(i), OC_EI, rws(i)).getDataArray()(0)
        s = "№ " & Trim(CStr(d(OC_ORDER))) & " — " & Trim(CStr(d(OC_NAME)))
        If gSpRest(i) > 0 Then s = s & " — ждём " & WmsIssue.QtyText(gSpRest(i)) & " " & Trim(CStr(d(OC_UNIT)))
        If Trim(CStr(d(OC_SUPPLIER))) <> "" Then s = s & " — " & Trim(CStr(d(OC_SUPPLIER)))
        opts(i) = Left(s, 120)
    Next i
    opts(n) = "другая позиция — выберу сам на листе «" & SH_ORDERS & "»"
    k = WmsPrimeUi.UiChoose("WMS — принять поставку", "Какая позиция приехала? Ожидаемые — с самых давних" & IIf(gSpN > n, " (показаны первые " & n & " из " & gSpN & ")", "") & ":", opts)
    If k < 0 Then
        NvReceive = "SKIP:окно закрыто"
        Exit Function
    End If
    WmsUi.NavOrders()
    If k >= n Then
        NvReceive = "OK:" & SH_ORDERS
        Exit Function
    End If
    gDoc.getCurrentController().select(sh.getCellByPosition(OC_NAME, rws(k)))
    NvReceive = WmsDaily.ReceiveWindow()
End Function

' ================================================================ «Все листы…»

Sub BtnSheets(Optional oEvent As Variant)
    WmsInit()
    SheetsMenu()
End Sub

' the purposes of the sheets (the first number of a line of NvTable)
Function SheetGroups() As Variant
    SheetGroups = Array("Каждый день — приход, выдача, возврат, исправления", "Склад — что есть и где лежит", "Найти и посмотреть", "Отчёты и документы", _
        "Закупки и поставщики", "Справочники и порядок в названиях", "Служебное — настройки, обслуживание, справка", "Прочие листы")
End Function

' purpose | sheet | what it is for
Private Function NvTable() As Variant
    NvTable = Array( _
        "0|Заказы|заказы поставщикам и приёмка по ним", "0|Иной приход|товар без заказа, по сертификату, от человека", "0|Приход авто|приехала машина: номер, поставщик", _
        "0|Выдачи|выдать со склада", "0|Возврат|принять возврат от человека", "0|Корректировки|переместить, списать, пересчитать остаток", _
        "1|Наличие|весь склад: что и сколько, где лежит", "1|Наличие по названию|остаток по названию, все партии вместе", "1|Остаток деталей|детали по артикулу", _
        "1|Серийные детали|детали с серийным номером и цветом", "1|Металл|прокат: марка, размер, длина, вес", "1|Метизы|болты, гайки, шайбы, винты", _
        "1|Инструмент ЧПУ|фрезы, свёрла, пластины, цанги, метчики", "1|Карта склада|где лежит и куда положить", "1|Места|места вне стеллажей: пол, двор, офис", "1|Топология|схема склада", _
        "2|Поиск|найти по номеру, названию, артикулу или человеку", "2|Дашборд|картина дня на одном экране", "2|Операции дня|что делали за день", "2|Календарь|операции по дням месяца", _
        "3|Отчёты|отчёты за период, по остаткам, руководителю", "3|Отчёт дня|документ за день", "3|ABC-анализ|что выдаётся чаще всего", "3|Без движения|залежавшиеся товары", _
        "3|Акты|реестр актов", "3|Печать|этикетки и накладные", _
        "4|Снабжение|что заказать, что ждём, сроки поставщиков", "4|Поставщики|справочник поставщиков", "4|Написания поставщиков|привести названия к одному написанию", "4|Резервы|что отложено под заказ", _
        "5|Получатели|кому выдаём", "5|Номенклатура|справочник товаров", "5|Комплекты|наборы для выдачи", "5|Сборка деталей|из чего собирается деталь", "5|Категории|категории товаров", _
        "5|Без категории|товары без категории", "5|Проверка названий|названия, которые стоит проверить", "5|Стандарты|ГОСТ, DIN, ISO", "5|Синонимы|как ещё называют товар", _
        "6|Главная|состояние и кнопки", "6|Настройки|что включено и что выключено", "6|Обслуживание|копии, проверки, скорость", "6|Справка|как работать, что нового, горячие клавиши", "6|Сверка переноса|итоги переноса старой книги")
End Function

' the sheet is there and is shown
Private Function NvShown(sNm As String) As Boolean
    NvShown = False
    If Not gDoc.Sheets.hasByName(sNm) Then Exit Function
    NvShown = gDoc.Sheets.getByName(sNm).IsVisible
End Function

' is the sheet in the table
Private Function NvKnown(sNm As String) As Boolean
    Dim t As Variant, i As Integer
    t = NvTable()
    NvKnown = False
    For i = 0 To UBound(t)
        If Split(t(i), "|")(1) = sNm Then
            NvKnown = True
            Exit Function
        End If
    Next i
End Function

' the shown sheets of a purpose g (0…6; 7 — those that are not in the table), one per line: «sheet|what it is for»
Function SheetsOf(g As Integer) As String
    Dim t As Variant, i As Integer, a As Variant, s As String, j As Long, sh As Object
    WmsInit()
    t = NvTable()
    If g < 7 Then
        For i = 0 To UBound(t)
            a = Split(t(i), "|")
            If CInt(a(0)) = g Then
                If NvShown(CStr(a(1))) Then s = s & IIf(s <> "", Chr(10), "") & a(1) & "|" & a(2)
            End If
        Next i
    Else
        For j = 0 To gDoc.Sheets.getCount() - 1
            sh = gDoc.Sheets.getByIndex(j)
            If sh.IsVisible Then
                If Left(sh.Name, 1) <> "_" Then
                    If Not NvKnown(sh.Name) Then s = s & IIf(s <> "", Chr(10), "") & sh.Name & "|"
                End If
            End If
        Next j
    End If
    SheetsOf = s
End Function

' go to the sheet the way its own button does (the lists are built, the cursor stands where the work starts)
Function SheetsGo(sNm As String) As String
    WmsInit()
    If Not gDoc.Sheets.hasByName(sNm) Then
        SheetsGo = "ERR:нет листа «" & sNm & "»"
        gUiLastMsg = SheetsGo
        Exit Function
    End If
    Select Case sNm
    Case SH_ORDERS
        WmsUi.NavOrders()
    Case SH_SPECIAL
        WmsUi.NavSpecial()
    Case SH_CARS
        WmsUi.NavCars()
    Case SH_ISSUES
        WmsUi.NavIssues()
    Case SH_RETURNS
        WmsUi.NavReturns()
    Case SH_ADJUST
        WmsUi.NavAdjust()
    Case SH_STOCK
        WmsUi.NavStock()
    Case SH_FIND
        WmsUi.NavFind()
    Case SH_MAP
        WmsUi.NavMap()
    Case SH_DASH
        WmsUi.NavDash()
    Case SH_MAIN
        WmsUi.NavMain()
    Case "Серийные детали"
        WmsSerial.NavSerials()
    Case "Метизы"
        WmsKeys.KeyFasteners()
    Case "Инструмент ЧПУ"
        WmsTool.NavTool()
    Case "Металл"
        WmsMetal.NavMetal()
    Case "Операции дня"
        WmsOps.BtnOpsOpen()
    Case "Отчёты"
        WmsLedger.NavReports()
    Case "Снабжение"
        WmsSupply.NavSupply()
    Case "Поставщики"
        WmsSupply.NavSuppliers()
    Case "Номенклатура"
        WmsSupply.NavGoods()
    Case "Написания поставщиков"
        WmsSup.NavSpel()
    Case "Комплекты"
        WmsKit.NavKits()
    Case "Сборка деталей"
        WmsAsm.NavAsm()
    Case "Резервы"
        WmsItem.NavReserves()
    Case "Обслуживание"
        WmsSafe.NavService()
    Case "Справка"
        WmsDaily.NavHelp()
    Case Else
        WmsView.GoCell(sNm, 0, 0)
    End Select
    SheetsGo = "OK:" & sNm
    gUiLastMsg = SheetsGo
End Function

' the window: first the purpose, then the sheet
Function SheetsMenu() As String
    Dim gl As Variant, opts() As String, idx() As Integer, i As Integer, k As Integer, n As Integer, s As String, lines As Variant, o2() As String, j As Integer, a As Variant
    WmsInit()
    gl = SheetGroups()
    ReDim opts(UBound(gl))
    ReDim idx(UBound(gl))
    For i = 0 To UBound(gl)
        s = SheetsOf(i)
        If s <> "" Then
            opts(n) = gl(i) & "  (" & (UBound(Split(s, Chr(10))) + 1) & ")"
            idx(n) = i
            n = n + 1
        End If
    Next i
    If n = 0 Then
        SheetsMenu = "SKIP:листов нет"
        Exit Function
    End If
    ReDim Preserve opts(n - 1)
    k = WmsPrimeUi.UiChoose("WMS — все листы", "Для чего нужен лист? В скобках — сколько листов:", opts)
    If k < 0 Then
        SheetsMenu = "SKIP:окно закрыто"
        Exit Function
    End If
    lines = Split(SheetsOf(idx(k)), Chr(10))
    ReDim o2(UBound(lines))
    For j = 0 To UBound(lines)
        a = Split(lines(j), "|")
        o2(j) = a(0)
        If UBound(a) >= 1 Then
            If a(1) <> "" Then o2(j) = a(0) & " — " & a(1)
        End If
    Next j
    j = WmsPrimeUi.UiChoose("WMS — " & gl(idx(k)), "Куда перейти?", o2)
    If j < 0 Then
        SheetsMenu = "SKIP:окно закрыто"
        Exit Function
    End If
    SheetsMenu = SheetsGo(CStr(Split(lines(j), "|")(0)))
End Function
