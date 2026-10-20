' WmsHelp — 1.0 (этап L): помощь и обучение. «Проверить этот ПК» за секунды (П1); «Что нового в 1.0» (П2); тур «Первый
' день» с подсветкой кнопок (П3); учебная WMS — отдельная копия со своим журналом (П4); «?» на каждом листе — свой раздел
' «Справки» (П6); шпаргалка A4 (П7); «Что делать, если…» (П8); ярлык «WMS склад» на рабочем столе (П10); обезличенная копия
' книги для разработчика (П12). Library WmsPlus.
Option Explicit

' П4, П12: the role of a copy made by WMS (a user property of the document): its first start registers it with its own journal
Public Const ROLE_PROP = "WMS_ROLE"
Public Const ROLE_TRAINING = "TRAINING"
Public Const ROLE_ANON = "ANON"
Global gHpDesktopTest As String

' ================================================================ П1 «Проверить этот ПК»

Sub BtnPcCheck(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage(PcCheck())
End Sub

' seconds, no internet: LibreOffice, the macros and the trusted folder, the folders written, the free space, Python for the
' tools — each line «✓/✗ что — как есть — что сделать»
Function PcCheck() As String
    Dim s As String, why As String, lvl As Integer, trusted As Boolean, p As String, free As Double, py As String, nBad As Integer, sec As Integer
    WmsInit()
    On Error GoTo EH
    WmsView.ProgressStart("Проверка этого ПК", 5)
    lvl = WmsStatus.LoCompat(why)
    s = s & PcLine(lvl = 0, "LibreOffice " & WmsStatus.LoVersion(), IIf(lvl = 0, "проверена с WMS", why), _
        IIf(lvl = 2, "поставьте LibreOffice 24.2 или новее", "работать можно; полная проверка — пакет проверки ПК"), nBad)
    WmsView.ProgressSet(1)
    sec = WmsStatus.MacroSecurity(trusted)
    s = s & PcLine(sec >= 1 And trusted, "макросы", "уровень безопасности " & sec & IIf(trusted, ", папка WMS — доверенная", ", папка WMS не в доверенных"), _
        "Сервис → Параметры → LibreOffice → Безопасность → Безопасность макросов: «Высокий», «Надёжные источники» — добавить папку WMS", nBad)
    WmsView.ProgressSet(2)
    p = WmsStatus.Writable(gJDir)
    If p = "" Then p = WmsStatus.Writable(gBDir)
    s = s & PcLine(p = "", "запись в папку WMS", IIf(p = "", "журнал и копии пишутся", p), "дайте этой учётной записи право записи в папку " _
        & ConvertFromURL(gBaseDir), nBad)
    WmsView.ProgressSet(3)
    free = WmsSafe.DiskFreeMb(gBaseDir)
    s = s & PcLine(free < 0 Or free >= WmsSafe.DISK_WARN_MB, "место на диске", IIf(free < 0, "не определено", Format(free / 1024, "0.0") & " ГБ свободно"), _
        "освободите хотя бы 1 ГБ: копии и журнал растут каждый день", nBad)
    WmsView.ProgressSet(4)
    py = PythonVersion()
    s = s & PcLine(py <> "", "Python (для инструментов WMS_TOOLBOX)", IIf(py <> "", py, "не найден"), _
        "нужен только для набора инструментов (этикетки на термопринтер, сверка, закономерности); WMS работает и без него", nBad)
    WmsView.ProgressEnd()
    PcCheck = "Проверка этого ПК: " & IIf(nBad = 0, "всё в порядке", "замечаний " & nBad) & Chr(10) & s
    gUiLastMsg = PcCheck
    Exit Function
EH:
    WmsView.ProgressEnd()
    PcCheck = "ERR:проверка не выполнена: " & Error$
    gUiLastMsg = PcCheck
End Function

Private Function PcLine(ok As Boolean, sWhat As String, sNow As String, sTodo As String, ByRef nBad As Integer) As String
    If Not ok Then nBad = nBad + 1
    PcLine = IIf(ok, "✓ ", "✗ ") & sWhat & ": " & sNow & IIf(ok, "", Chr(10) & "    что сделать: " & sTodo) & Chr(10)
End Function

' the version of Python of this computer ("" — none): «python3 --version» (Windows: «py»/«python») into a file
Function PythonVersion() As String
    Dim f As String, s As String, cmds As Variant, i As Integer
    On Error GoTo EH
    If WmsSafe.SafeDir() = "" Then Exit Function
    If Not gSfa.exists(WmsSafe.SafeDir()) Then gSfa.createFolder(WmsSafe.SafeDir())
    f = WmsSafe.SafeDir() & "py_" & NewId(4) & ".txt"
    If GetGuiType() = 1 Then
        cmds = Array("py", "python")
        For i = 0 To 1
            Shell("cmd.exe", 0, "/c " & cmds(i) & " --version > """ & ConvertFromURL(f) & """ 2>&1", True)
            If gSfa.exists(f) Then
                s = Trim(Replace(Replace(WmsAudit.ReadText(f), Chr(13), ""), Chr(10), ""))
                gSfa.kill(f)
                If Left(s, 7) = "Python " Then Exit For
                s = ""
            End If
        Next i
    Else
        Shell("/bin/sh", 0, "-c ""python3 --version > '" & Replace(ConvertFromURL(f), "'", "'\''") & "' 2>&1""", True)
        If gSfa.exists(f) Then
            s = Trim(Replace(WmsAudit.ReadText(f), Chr(10), ""))
            gSfa.kill(f)
        End If
        If Left(s, 7) <> "Python " Then s = ""
    End If
    PythonVersion = s
    Exit Function
EH:
    PythonVersion = ""
End Function

' ================================================================ П2 «Что нового в 1.0»

' the cards: «заголовок|что это|лист, куда «Показать» ведёт»
Function WhatsNewCards() As Variant
    WhatsNewCards = Array( _
        "Выдать и принять быстрее|Окна «Выдать…» и «Принять…», выдача пачкой одним документом, «Повторить», «Подобрать ЕИ», комплекты.|" & SH_ISSUES, _
        "Фасовка, этикетки, свойства|«1 кан», «25%», этикетка ЕИ со штрихкодом, серийный номер и «годен до», резерв, фото.|" & SH_STOCK, _
        "Поиск и карточка ЕИ|Найти по номеру, имени, месту, серийному номеру; вся история ЕИ, печать карточки.|" & SH_FIND, _
        "Карта склада|Где что лежит, «Куда положить?», правка карты прямо на карте.|" & SH_MAP, _
        "Отчёты и документы|Остаток на дату, оборотная ведомость, выгрузка для 1С, акты, накладные, закрытие месяца.|Отчёты", _
        "Сохранность данных|Копия на флешку, зеркало журнала, «Вернуться к копии», проверка дня.|" & WmsSafe.SH_SERVICE)
End Function

Sub BtnWhatsNew(Optional oEvent As Variant)
    WmsInit()
    WhatsNew()
End Sub

' the window of the cards; «Показать» opens the sheet of the card. The chosen card (0-based) or -1
Function WhatsNew() As Integer
    Dim cards As Variant, opts() As String, i As Integer, k As Integer, a As Variant
    cards = WhatsNewCards()
    ReDim opts(UBound(cards))
    For i = 0 To UBound(cards)
        a = Split(cards(i), "|")
        opts(i) = (i + 1) & ". " & a(0) & " — " & a(1)
    Next i
    k = WmsPrimeUi.UiChoose("Что нового в WMS 1.0", "Выберите и нажмите OK — WMS покажет, где это. Это окно можно открыть снова: «Справка» → «Что нового».", opts)
    WhatsNew = k
    If k >= 0 Then
        a = Split(cards(k), "|")
        WmsView.GoCell(CStr(a(2)), 0, 0)
    End If
    WmsSafe.StSet("WHATSNEW", WMS_PRODUCT_VERSION)
End Function

' at the start: the cards once a version, to a person (never in a test book), after the start finished
Sub WhatsNewAtStart()
    On Error Resume Next
    If SysStr(SK_MODE) = "TEST" Or gUiAuto <> 0 Then Exit Sub
    If WmsSafe.StGet("WHATSNEW") = WMS_PRODUCT_VERSION Then Exit Sub
    HpLater()
End Sub

Private Sub HpLater()
    gHpCb = CreateUnoListener("HpCb_", "com.sun.star.awt.XCallback")
    CreateUnoService("com.sun.star.awt.AsyncCallback").addCallback(gHpCb, Array())
End Sub

Global gHpCb As Object

Sub HpCb_notify(aData As Variant)
    WhatsNew()
End Sub

' ================================================================ П3 the tour «Первый день»

' the steps: «лист|кнопка (имя) or ""|текст»
Function TourSteps() As Variant
    TourSteps = Array( _
        SH_MAIN & "||«Главная»: вверху — состояние WMS. «Работа разрешена» — можно работать. Если «ЗАБЛОКИРОВАНО» — ниже написано, что сделать.", _
        "Дашборд||«Дашборд» открывается первым: что ожидается, что пришло и выдано сегодня, «Внимание» — что требует действия (двойной щелчок — список).", _
        SH_ORDERS & "|btnReceiveWin|Приход по заказу: встаньте на строку заказа и нажмите «Принять…» — окно спросит факт, документ и место.", _
        SH_ISSUES & "|btnIssueWin|Выдача: «Выдать…» — окно; можно и вписать строку (ЕИ, количество, кому) и нажать «Провести». «1 кан», «25%» — тоже можно.", _
        SH_RETURNS & "||Возврат: впишите № выдачи и сколько вернули, нажмите «Провести».", _
        SH_FIND & "|btnFind|Поиск: впишите номер ЕИ, имя, место или получателя и нажмите Enter. Двойной щелчок по строке — карточка ЕИ.", _
        WmsSafe.SH_SERVICE & "|btnFlashCopy|Конец дня: сохраните книгу (Ctrl+S). Раз в неделю — «Копия на флешку…» здесь, на «Обслуживании».")
End Function

Sub BtnTour(Optional oEvent As Variant)
    WmsInit()
    Tour()
End Sub

' the steps one after another: the sheet opened, its button lit (yellow) while the step is read; «Закончить» stops. The steps shown
Function Tour() As Integer
    Dim steps As Variant, i As Integer, a As Variant, k As Integer, shp As Object, oldColor As Long, lit As Boolean, n As Integer
    steps = TourSteps()
    n = UBound(steps) + 1
    For i = 0 To n - 1
        a = Split(steps(i), "|")
        If gDoc.Sheets.hasByName(CStr(a(0))) Then WmsView.GoCell(CStr(a(0)), 0, 0)
        lit = False
        If a(1) <> "" Then
            shp = ButtonModel(CStr(a(0)), CStr(a(1)))
            If Not IsNull(shp) Then
                oldColor = shp.BackgroundColor
                shp.BackgroundColor = RGB(255, 230, 0)
                lit = True
            End If
        End If
        If i < n - 1 Then
            k = WmsPrimeUi.UiButtons("Первый день — шаг " & (i + 1) & " из " & n, CStr(a(2)), Array("Дальше", "Закончить"), Array(1, 0))
        Else
            k = WmsPrimeUi.UiButtons("Первый день — шаг " & (i + 1) & " из " & n, CStr(a(2)), Array("Готово", "Закончить"), Array(1, 0))
        End If
        If lit Then shp.BackgroundColor = oldColor
        Tour = i + 1
        If k <> 1 Then Exit Function
    Next i
End Function

' the model of a button of a sheet by its name (Nothing — none)
Function ButtonModel(sSheet As String, sName As String) As Object
    Dim f As Object, i As Integer
    On Error GoTo EH
    If gDoc.Sheets.getByName(sSheet).getDrawPage().getForms().getCount() = 0 Then Exit Function
    f = gDoc.Sheets.getByName(sSheet).getDrawPage().getForms().getByIndex(0)
    For i = 0 To f.getCount() - 1
        If f.getByIndex(i).Name = sName Then
            ButtonModel = f.getByIndex(i)
            Exit Function
        End If
    Next i
EH:
End Function

' ================================================================ П6 «?» — the section of «Справка» of a sheet

' the beginning of the heading of «Справка» of each sheet
Function HelpTopic(sSheet As String) As String
    Select Case sSheet
    Case SH_MAIN, "Дашборд": HelpTopic = "1. "
    Case SH_ORDERS: HelpTopic = "2. "
    Case SH_SPECIAL: HelpTopic = "3. "
    Case SH_CARS: HelpTopic = "3а. "
    Case SH_ISSUES: HelpTopic = "4. "
    Case SH_RETURNS: HelpTopic = "5. "
    Case SH_ADJUST: HelpTopic = "6. "
    Case SH_FIND: HelpTopic = "6а. "
    Case SH_MAP, "Топология", "Места": HelpTopic = "6б. "
    Case SH_STOCK, "Резервы": HelpTopic = "6в. "
    Case "Снабжение", "Поставщики", "Номенклатура": HelpTopic = "6г. "
    Case "Комплекты": HelpTopic = "6д. "
    Case "Отчёты", SH_PRINT: HelpTopic = "6е. "
    Case WmsSafe.SH_SERVICE: HelpTopic = "8а. "
    Case Else: HelpTopic = ""
    End Select
End Function

Sub BtnHelpHere(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = HelpGo(gDoc.getCurrentController().getActiveSheet().getName())
End Sub

' «Справка» opened at the section of sheet sSheet (its heading at the top). OK:<row> / ERR
Function HelpGo(sSheet As String) As String
    Dim t As String, sh As Object, d As Variant, i As Long, last As Long
    On Error GoTo EH
    t = HelpTopic(sSheet)
    sh = gDoc.Sheets.getByName("Справка")
    last = WmsOrders.LastRow(sh)
    If t <> "" Then
        d = sh.getCellRangeByPosition(0, 0, 0, last).getDataArray()
        For i = 0 To UBound(d)
            If Left(CStr(d(i)(0)), Len(t)) = t Then
                WmsView.GoCell("Справка", i, 0)
                HelpGo = "OK:" & i
                gUiLastMsg = HelpGo
                Exit Function
            End If
        Next i
    End If
    WmsView.GoCell("Справка", 0, 0)
    HelpGo = "OK:0"
    gUiLastMsg = HelpGo
    Exit Function
EH:
    HelpGo = "ERR:" & Error$
End Function

' ================================================================ П7 the cheat sheet A4

Sub BtnCheatSheet(Optional oEvent As Variant)
    WmsInit()
    CheatSheet()
End Sub

' one or two A4 pages on «Печать»: the start of the day, a receipt, an issue in parts, a return, a correction, «если красное»
Function CheatSheet() As String
    Dim sh As Object, a As Variant, i As Integer
    sh = WmsPrint.PrintBegin("WMS — шпаргалка кладовщика", "Шпаргалка")
    a = Array("#Начало дня", "Откройте книгу WMS. На «Главной» — «Работа разрешена». «Дашборд»: что ждём, что сегодня, «Внимание».", _
        "Если вверху «ЗАБЛОКИРОВАНО» — ниже написано, что сделать (чаще всего «Восстановить»).", _
        "#Приход по заказу («Заказы»)", "Встаньте на строку заказа → «Принять…» → факт, документ, место → «Принять». Всё по заказу — «Принять всё по заказу».", _
        "Новый ЕИ получает номер сам; этикетка — «Этикетка».", _
        "#Выдача («Выдачи»)", "«Выдать…» → ЕИ (427 или ЕИ-427), количество, кому → «Провести». Несколько строк — выделить и «Провести» одним документом.", _
        "Частями: «1 кан», «1,5 м», «25%». Не помните ЕИ — «Подобрать ЕИ». Комплект — «Выдать комплект…».", _
        "#Возврат («Возврат»)", "№ выдачи, сколько вернули, дата → «Провести». Не знаете № — «Найти выдачу».", _
        "#Перемещение, списание, инвентаризация («Корректировки»)", "Вид, ЕИ, количество или факт, место → «Провести». Пересчитать место — «Пересчитать место…».", _
        "#Ошиблись", "Проведённую строку не правят руками: «Исправить» (с причиной) или «Удалить» (сторно, с причиной). История — «Поиск» → «История изменений».", _
        "#Если красное", "Строка не провелась — колонка «Контроль» и кнопка «Почему?». «1 кан» не принимается — задайте «Фасовку ЕИ».", _
        "Не знаете, что делать — «Справка» → «Что делать, если…».", _
        "#Конец дня", "Сохраните книгу (Ctrl+S). Раз в неделю — «Обслуживание» → «Копия на флешку…».")
    For i = 0 To UBound(a)
        If Left(a(i), 1) = "#" Then
            WmsPrint.PrintLine(sh, "")
            WmsPrint.PrintLine(sh, Mid(a(i), 2), True)
        Else
            WmsPrint.PrintLine(sh, CStr(a(i)))
        End If
    Next i
    WmsPrint.PrintDone(sh)
    CheatSheet = "OK:шпаргалка на «Печати»"
    gUiLastMsg = CheatSheet
End Function

' ================================================================ П8 «Что делать, если…»

' the situations: «ситуация|ответ|лист «Перейти»»
Function IfCases() As Variant
    IfCases = Array( _
        "Вверху «ЗАБЛОКИРОВАНО»|Ниже на «Главной» написано, почему и что сделать. Чаще всего — «Восстановить» (доведёт книгу по журналу). Ничего не удаляйте " _
            & "и не копируйте поверх.|" & SH_MAIN, _
        "Строка не проводится|Посмотрите колонку «Контроль» строки и нажмите «Почему?» — там причина и что исправить. Не тот ЕИ, не хватает остатка, " _
            & "дата в закрытом месяце — исправьте ввод и снова «Провести».|" & SH_ISSUES, _
        "«1 кан» не принимается|У ЕИ не задана фасовка: «Наличие» → строка ЕИ → «Фасовка ЕИ» (сколько кан в 1 шт). Потом «1 кан» выдаётся.|" & SH_STOCK, _
        "Выдал не то или не тому|Найдите строку выдачи → «Исправить» (верное количество или получатель, с причиной) или «Удалить» (сторно). " _
            & "Руками проведённую строку не правьте.|" & SH_ISSUES, _
        "Остаток не сходится с полкой|«Корректировки» → «Пересчитать место…» или вид «Инвентаризация»: впишите факт — WMS проведёт разницу. " _
            & "Сначала посмотрите историю ЕИ на «Поиске».|" & SH_ADJUST, _
        "Книга испорчена, нужна старая копия|«Обслуживание» → «Вернуться к копии…»: выберите копию — WMS поставит её на место и доведёт по журналу.|" _
            & WmsSafe.SH_SERVICE, _
        "Журнал испорчен|«Обслуживание» → «Журнал из зеркала» (если зеркало включено). Иначе сообщите ответственному за WMS, файлы журнала не правьте.|" _
            & WmsSafe.SH_SERVICE, _
        "Всё стало медленно|«Обслуживание»: «Советы LibreOffice», «Лишние ячейки», «Замер скорости ПК». Книга больше 200 000 движений — нужно " _
            & "«Закрытие года».|" & WmsSafe.SH_SERVICE, _
        "Забыл сохранить, был сбой|Откройте книгу: WMS увидит операции журнала, которых нет в книге, — нажмите «Восстановить». Ничего не пропадёт.|" & SH_MAIN, _
        "Дата не принимается|Набирайте дату полностью: 20.10.2026. Короткое «20.10» Calc понимает как число 20,1 (19.01.1900), и WMS его не примет.|" _
            & SH_ISSUES, _
        "«Здоровье WMS» жёлтое или красное|В строке написано, что сделать и какой кнопкой: красное — сейчас (проверка дня, диск, журнал), жёлтое — " _
            & "в ближайшие дни (копия на флешку, учебное восстановление раз в месяц).|" & WmsSafe.SH_SERVICE)
End Function

Sub BtnWhatIf(Optional oEvent As Variant)
    Dim cases As Variant, opts() As String, i As Integer, k As Integer, a As Variant, j As Integer
    WmsInit()
    cases = IfCases()
    ReDim opts(UBound(cases))
    For i = 0 To UBound(cases)
        opts(i) = Split(cases(i), "|")(0)
    Next i
    k = WmsPrimeUi.UiChoose("Что делать, если…", "Выберите, что случилось:", opts)
    If k < 0 Then Exit Sub
    a = Split(cases(k), "|")
    j = WmsPrimeUi.UiButtons("Что делать: " & a(0), a(1), Array("Перейти", "Закрыть"), Array(1, 0))
    If j = 1 Then WmsView.GoCell(CStr(a(2)), 0, 0)
End Sub

' ================================================================ П10 the shortcut «WMS склад» on the desktop

Sub BtnShortcut(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage(Shortcut())
End Sub

' the shortcut to this book on the desktop of this computer (Windows: «WMS склад.lnk»; Linux: «WMS склад.desktop»): the book is
' always opened where it is, never an old copy. OK:… / ERR:…
Function Shortcut() As String
    Dim desk As String, book As String, sh As Object, lnk As Object, f As String, s As String
    WmsInit()
    On Error GoTo EH
    book = ConvertFromURL(gDoc.getURL())
    If GetGuiType() = 1 And gHpDesktopTest = "" Then
        sh = CreateObject("WScript.Shell")
        desk = sh.SpecialFolders("Desktop")
        lnk = sh.CreateShortcut(desk & "\WMS склад.lnk")
        lnk.TargetPath = book
        lnk.WorkingDirectory = ConvertFromURL(Left(gBaseDir, Len(gBaseDir) - 1))
        lnk.Description = "Рабочая книга WMS"
        lnk.Save()
        f = desk & "\WMS склад.lnk"
    Else
        desk = gHpDesktopTest
        If desk = "" Then desk = Environ("HOME") & "/Desktop"
        If Not gSfa.exists(ConvertToURL(desk)) Then gSfa.createFolder(ConvertToURL(desk))
        f = desk & "/WMS склад.desktop"
        s = "[Desktop Entry]" & Chr(10) & "Type=Application" & Chr(10) & "Name=WMS склад" & Chr(10) & "Comment=Рабочая книга WMS" & Chr(10) _
            & "Exec=soffice """ & book & """" & Chr(10) & "Icon=libreoffice-calc" & Chr(10) & "Terminal=false" & Chr(10)
        WmsSafe.WriteText(ConvertToURL(f), s)
    End If
    Shortcut = "OK:Ярлык «WMS склад» на рабочем столе: " & f & Chr(10) & "Он открывает эту книгу: " & book
    gUiLastMsg = Shortcut
    Exit Function
EH:
    Shortcut = "ERR:ярлык не создан: " & Error$
    gUiLastMsg = Shortcut
End Function

' ================================================================ П4, П12 the training copy and the anonymous copy

Sub BtnTraining(Optional oEvent As Variant)
    Dim s As String, k As Integer
    WmsInit()
    k = WmsPrimeUi.UiButtons("Учебная WMS", "WMS сделает отдельную учебную книгу — копию этой (на момент сохранения) в папке рядом, со своим журналом. " _
        & "В ней можно пробовать всё: «1 кан», «Исправить», «Удалить». Рабочая книга и её журнал не меняются." & Chr(10) _
        & "Учебная книга подписана «УЧЕБНАЯ» в заголовке окна и на «Дашборде».", Array("Сделать учебную WMS", "Отмена"), Array(1, 0))
    If k <> 1 Then Exit Sub
    s = MakeCopy(ROLE_TRAINING, True)
    WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
End Sub

Sub BtnAnonCopy(Optional oEvent As Variant)
    Dim s As String, k As Integer
    WmsInit()
    k = WmsPrimeUi.UiButtons("Обезличенная копия", "WMS сделает копию книги для разработчика: наименования, артикулы, люди, поставщики, документы и " _
        & "примечания заменены на «Товар 1», «Получатель 1»… Количества, места, даты и структура — как есть. Журнал не копируется. " _
        & "Рабочая книга не меняется.", Array("Сделать копию", "Отмена"), Array(1, 0))
    If k <> 1 Then Exit Sub
    s = MakeCopy(ROLE_ANON, False)
    WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
End Sub

' the copy of the book (as last saved) as a book of its own: TRAINING — «<папка WMS>_УЧЕБНАЯ/<книга>_УЧЕБНАЯ.ods», opened when bOpen;
' ANON — «WMS_Для_разработчика/<дата>/<книга>_ОБЕЗЛИЧЕННАЯ.ods», the texts replaced. Its first start registers it with a new journal
' (ROLE_PROP). OK:<path> / ERR:…
Function MakeCopy(sRole As String, bOpen As Boolean) As String
    Dim dirUrl As String, url As String, sBase As String, hd As Object, args(1) As New com.sun.star.beans.PropertyValue, up As Object, n As Long
    Dim parent As String, stamp As String
    WmsInit()
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        MakeCopy = "ERR:книга не сохранена в файл"
        Exit Function
    End If
    sBase = WmsBackup.BookBaseName()
    stamp = Format(Now(), "YYYY-MM-DD") & "_" & Format(Now(), "HHMMSS")
    If sRole = ROLE_TRAINING Then
        parent = Left(gBaseDir, Len(gBaseDir) - 1)
        dirUrl = parent & "_УЧЕБНАЯ/"
        If gSfa.exists(dirUrl) Then
            ' the earlier training copy put aside, never mixed with the new one
            gSfa.move(dirUrl, parent & "_УЧЕБНАЯ_" & stamp & "/")
        End If
        url = dirUrl & sBase & "_УЧЕБНАЯ.ods"
    Else
        dirUrl = gBaseDir & "WMS_Для_разработчика/" & stamp & "/"
        url = dirUrl & sBase & "_ОБЕЗЛИЧЕННАЯ.ods"
    End If
    If WmsSafe.DiskRefuse(gBaseDir, 2 * WmsSafe.BookMb()) <> "" Then
        MakeCopy = "ERR:" & WmsSafe.DiskRefuse(gBaseDir, 2 * WmsSafe.BookMb())
        Exit Function
    End If
    gSfa.createFolder(dirUrl)
    gSfa.copy(gDoc.getURL(), url)
    ' the copy changed without its macros: the role, the title of the window, a new registration and instance
    args(0).Name = "Hidden"
    args(0).Value = True
    args(1).Name = "MacroExecutionMode"
    args(1).Value = 0
    hd = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    up = hd.getDocumentProperties().getUserDefinedProperties()
    If up.getPropertySetInfo().hasPropertyByName(ROLE_PROP) Then up.removeProperty(ROLE_PROP)
    up.addProperty(ROLE_PROP, 128, sRole)
    hd.getDocumentProperties().Title = IIf(sRole = ROLE_TRAINING, "УЧЕБНАЯ WMS — проводки не настоящие", "ОБЕЗЛИЧЕННАЯ копия WMS (для разработчика)")
    CopySys(hd, SK_REG_URL, "")
    CopySys(hd, SK_INSTANCE, NewId(16))
    If sRole = ROLE_ANON Then n = Anonymize(hd)
    hd.store()
    hd.close(True)
    MakeCopy = "OK:" & IIf(sRole = ROLE_TRAINING, "Учебная WMS готова: ", "Обезличенная копия готова (заменено значений: " & n & "): ") & ConvertFromURL(url) _
        & IIf(sRole = ROLE_ANON, Chr(10) & "Отправьте разработчику папку " & ConvertFromURL(dirUrl), "")
    gUiLastMsg = MakeCopy
    If bOpen Then
        args(0).Name = "MacroExecutionMode"
        args(0).Value = 4
        args(1).Name = "Hidden"
        args(1).Value = False
        StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    End If
    Exit Function
EH:
    MakeCopy = "ERR:копия не сделана: " & Error$
    gUiLastMsg = MakeCopy
    On Error Resume Next
    If Not IsNull(hd) And Not IsEmpty(hd) Then hd.close(True)
End Function

Private Sub CopySys(hd As Object, k As Integer, v As String)
    Dim sh As Object
    sh = hd.Sheets.getByName(SYS_SHEET)
    If sh.isProtected() Then sh.unprotect(PROTECT_PWD)
    sh.getCellByPosition(1, k).setString(v)
    sh.protect(PROTECT_PWD)
End Sub

' the role of this book (TRAINING / ANON) or ""
Function BookRole() As String
    Dim up As Object
    On Error GoTo EH
    up = gDoc.getDocumentProperties().getUserDefinedProperties()
    If up.getPropertySetInfo().hasPropertyByName(ROLE_PROP) Then BookRole = up.getPropertyValue(ROLE_PROP)
EH:
End Function

' WmsRecovery.Analyze: a copy made by MakeCopy, not registered yet and with no journal of its own — registered with a new one
' (the start runs again inside). True when so
Function CopyRoleStart() As Boolean
    Dim r As String
    On Error GoTo EH
    r = BookRole()
    If r <> ROLE_TRAINING And r <> ROLE_ANON Then Exit Function
    If SysStr(SK_REG_URL) <> "" Then Exit Function
    If UBound(WmsJournal.ListFiles()) >= 0 Then Exit Function
    WmsRecovery.RegisterHere(True)
    CopyRoleStart = True
EH:
End Function

' П12: the texts that name goods, people, suppliers and documents replaced in every table (one value — one substitute in all of
' them); the pages that show them (search, print, dashboard) cleared. The number of the values replaced
Private Function Anonymize(hd As Object) As Long
    Dim spec As Variant, i As Integer, j As Integer, a As Variant, sh As Object, last As Long, d As Variant, r As Long, v As String, k As String
    Dim maps(6) As Object, cnt(6) As Long, kind As Integer, n As Long, nm As Variant, c As Integer, x As Long, r0 As Long, fp As Variant, q As Variant
    nm = Array("Товар", "", "Получатель", "Поставщик", "Документ", "Примечание", "Машина")
    For i = 0 To 6
        maps(i) = New Collection
    Next i
    ' sheet | column:kind, … (kind: 0 goods, 2 person, 3 supplier, 4 document, 5 note, 6 vehicle); the articles stay (the index _ART)
    spec = Array(Array(SH_STOCK, "1:0", "8:4"), _
        Array(SH_ORDERS, "1:0", "2:4", "3:4", "11:3", "12:3", "17:2", "19:2", "25:5"), _
        Array(SH_ISSUES, "1:4", "2:0", "8:2", "14:5"), _
        Array(SH_RETURNS, "3:0", "8:2", "14:5"), _
        Array(SH_SPECIAL, "3:0", "10:2", "11:4", "18:5"), _
        Array(SH_ADJUST, "3:0", "11:5", "18:5"), _
        Array(SH_CARS, CC_PLATE & ":6", CC_SUP & ":3", CC_NOTE & ":5"), _
        Array(SH_RCPT, "0:2", "1:2"), Array("Поставщики", "0:3", "1:5", "2:2", "4:5"), _
        Array("Номенклатура", "1:0", "9:3", "10:5", "11:5"), Array("Комплекты", "0:0", "1:0", "5:5"), _
        Array("Резервы", "2:2", "4:5"), Array("_PROP", "3:5"))
    For i = 0 To UBound(spec)
        If hd.Sheets.hasByName(spec(i)(0)) Then
            sh = hd.Sheets.getByName(spec(i)(0))
            If sh.isProtected() Then sh.unprotect(PROTECT_PWD)
            last = WmsOrders.LastRow(sh)
            r0 = IIf(spec(i)(0) = SH_CARS, CAR_FIRST, 1)
            If last >= r0 Then
                For j = 1 To UBound(spec(i))
                    a = Split(spec(i)(j), ":")
                    c = CInt(a(0))
                    kind = CInt(a(1))
                    d = sh.getCellRangeByPosition(c, r0, c, last).getDataArray()
                    For r = 0 To UBound(d)
                        If VarType(d(r)(0)) = 8 Then
                            If Trim(CStr(d(r)(0))) <> "" Then
                                d(r)(0) = AnonValue(maps, cnt, nm, kind, CStr(d(r)(0)), n)
                            End If
                        End If
                    Next r
                    sh.getCellRangeByPosition(c, r0, c, last).setDataArray(d)
                Next j
            End If
        End If
    Next i
    ' _ORD: the fingerprint of a position names its document, goods and supplier (A|D|B|E|H|I|L|P): the same substitutes there
    If hd.Sheets.hasByName(SH_ORD) Then
        sh = hd.Sheets.getByName(SH_ORD)
        If sh.isProtected() Then sh.unprotect(PROTECT_PWD)
        last = WmsOrders.LastRow(sh)
        If last >= 1 Then
            d = sh.getCellRangeByPosition(3, 1, 3, last).getDataArray()
            For r = 0 To UBound(d)
                fp = Split(CStr(d(r)(0)), "|")
                If UBound(fp) >= 6 Then
                    For Each q In Array(1, 2, 6)
                        If Left(fp(q), 1) = "S" And Len(fp(q)) > 1 Then fp(q) = "S" & AnonValue(maps, cnt, nm, IIf(q = 1, 4, IIf(q = 2, 0, 3)), Mid(fp(q), 2), n)
                    Next q
                    d(r)(0) = Join(fp, "|")
                End If
            Next r
            sh.getCellRangeByPosition(3, 1, 3, last).setDataArray(d)
        End If
    End If
    ' the pages that show names: emptied (they are made again by their buttons)
    For Each a In Array(SH_FIND, SH_PRINT, "Дашборд", "Сверка переноса")
        If hd.Sheets.hasByName(a) Then
            sh = hd.Sheets.getByName(a)
            If sh.isProtected() Then sh.unprotect(PROTECT_PWD)
            sh.getCellRangeByPosition(0, IIf(a = "Дашборд", 5, 6), 15, 3000).clearContents(1 + 2 + 4)
        End If
    Next a
    Anonymize = n
End Function

' the substitute of value v of kind (one value — one substitute in the whole book; n — the substitutes made)
Private Function AnonValue(maps As Variant, cnt As Variant, nm As Variant, kind As Integer, v As String, ByRef n As Long) As String
    Dim k As String, x As Long
    k = LCase(Trim(v))
    x = CollGet(maps(kind), k)
    If x < 0 Then
        cnt(kind) = cnt(kind) + 1
        x = cnt(kind)
        CollPut(maps(kind), x, k)
        n = n + 1
    End If
    AnonValue = nm(kind) & " " & x
End Function

' test seam: the folder of the desktop of «Ярлык на рабочий стол» ("" — the real one)
Function TestHelpDesktop(sPath As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then Exit Function
    gHpDesktopTest = sPath
    TestHelpDesktop = "OK"
End Function
