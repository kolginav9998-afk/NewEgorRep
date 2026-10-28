' WmsKeys — 1.1: the keyboard shortcuts of the book (the table is HOTKEYS_10 / HOTKEYS_11 of tools/build_ods.py; they are stored in the document and act in this book only): Ctrl+Shift + a letter.
' Here: the shortcuts that act by the sheet in view («Проверить», «Исправить»), the list of all shortcuts (Hotkeys, the window of «Справка», the cheat sheet) and the jump to «Метизы».
' Library WmsPlus.
Option Explicit

' Ctrl+Shift+K — «Проверить» the row under the cursor of the sheet in view
Sub KeyCheck(Optional oEvent As Variant)
    Dim sName As String
    WmsInit()
    sName = gDoc.getCurrentController().getActiveSheet().Name
    Select Case sName
    Case SH_ORDERS
        WmsOrdersUi.BtnRcvCheck()
    Case SH_RETURNS
        WmsReturnUi.BtnRetCheck()
    Case SH_SPECIAL
        WmsSpecialUi.BtnSpcCheck()
    Case SH_ADJUST
        WmsAdjustUi.BtnAdjCheck()
    Case SH_ISSUES
        WmsUi.UiMessage("На «" & SH_ISSUES & "» строка проверяется при проведении: Ctrl+Shift+P. Колонка «Контроль» говорит, что не так, кнопка «Почему?» — как исправить.")
    Case Else
        WmsUi.UiMessage("Ctrl+Shift+K — «Проверить» строку на листах «" & SH_ORDERS & "», «" & SH_RETURNS & "», «" & SH_SPECIAL & "», «" & SH_ADJUST & "».")
    End Select
End Sub

' Ctrl+Shift+U — «Исправить» the posted row under the cursor of the sheet in view (WMS asks the reason)
Sub KeyFix(Optional oEvent As Variant)
    Dim sName As String
    WmsInit()
    sName = gDoc.getCurrentController().getActiveSheet().Name
    Select Case sName
    Case SH_ISSUES
        WmsUi.BtnFix()
    Case SH_ORDERS
        WmsOrdersUi.BtnRcvFix()
    Case SH_RETURNS
        WmsReturnUi.BtnRetFix()
    Case SH_SPECIAL
        WmsSpecialUi.BtnSpcFix()
    Case SH_ADJUST
        WmsAdjustUi.BtnAdjFix()
    Case Else
        WmsUi.UiMessage("Ctrl+Shift+U — «Исправить» проведённую строку на листах «" & SH_ISSUES & "», «" & SH_ORDERS & "», «" & SH_RETURNS & "», «" & SH_SPECIAL & "», «" & SH_ADJUST & "».")
    End Select
End Sub

' Ctrl+Shift+G — «Метизы», the cursor in «Найти»
Sub KeyFasteners(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(WmsCat.SH_FAST, WmsCat.FS_FIND_ROW, 1)
End Sub

' the shortcuts, one per line: «Ctrl+Shift+X — what it does»
Function Hotkeys() As String
    Dim a As Variant, i As Long, s As String
    a = Array("#Перейти на лист", "Ctrl+Shift+H — «Главная»", "Ctrl+Shift+D — «Дашборд»", "Ctrl+Shift+F — «Поиск» (курсор в «Что найти»)", "Ctrl+Shift+M — «Карта склада»", _
        "Ctrl+Shift+O — «Заказы» (приход по заказу)", "Ctrl+Shift+N — «Иной приход»", "Ctrl+Shift+I — «Выдачи»", "Ctrl+Shift+R — «Возврат»", "Ctrl+Shift+A — «Корректировки»", _
        "Ctrl+Shift+L — «Наличие»", "Ctrl+Shift+G — «Метизы» (курсор в «Найти»)", "Ctrl+Shift+J — «Операции дня»", "Ctrl+Shift+X — «Обслуживание»", "Ctrl+Shift+B — «Справка»", _
        "#Сделать на листе", "Ctrl+Shift+P — «Провести» строку (лист в работе)", "Ctrl+Shift+K — «Проверить» строку", "Ctrl+Shift+U — «Исправить» проведённую строку", _
        "Ctrl+Shift+W — «Выдать…» (окно)", "Ctrl+Shift+T — «Принять…» (окно, на «Заказах»)", "Ctrl+Shift+E — «Подобрать ЕИ»", "Ctrl+Shift+Q — «Куда положить?»", _
        "#Если что-то не так", "Ctrl+Shift+Y — «Что делать?» (совет под красными пунктами самопроверки)")
    For i = 0 To UBound(a)
        If Left(a(i), 1) = "#" Then
            s = s & IIf(s <> "", Chr(10), "") & Mid(a(i), 2) & ":" & Chr(10)
        Else
            s = s & CStr(a(i)) & Chr(10)
        End If
    Next i
    Hotkeys = s
End Function

Sub BtnHotkeys(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage("Горячие клавиши WMS (действуют в этой книге):" & Chr(10) & Chr(10) & Hotkeys() & Chr(10) & "Привязаны к клавише, а не к букве, поэтому должны работать и при русской раскладке; если какая-то не сработает — переключитесь на английскую. Список — в «Справке» → «Шпаргалка A4».")
End Sub
