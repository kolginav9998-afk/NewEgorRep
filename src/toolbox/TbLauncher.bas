' TbLauncher — WMS_TOOLBOX.ods: одна точка входа в инструменты (задание «FINAL WMS MARATHON», §7). Инструменты
' открываются по относительным путям (книги лежат рядом с этой книгой); для работы WMS_PROD.ods книга не нужна. Лист
' «Инструменты»: что есть в папке, последний снимок WMS (имя, время, возраст), пакеты для WMS; скрипты запускаются
' отсюда же, если на ПК есть python3: «Проверка перед сменой»; «Сверка названий» (WMS_RECONCILE: файл старых названий →
' отчёт с кандидатами-ЕИ открывается для отметок «да») и «Соответствие по сверке» (отмеченный отчёт → соответствие
' старых названий и ЕИ).
Option Explicit

Private Const SH_LIST = "Инструменты"
Private Const FIRST_ROW = 3            ' 0-based: the tools from row 4 (row 3 — the headers)

Global gLnFile As String              ' test seam: the file «Сверка названий» / «Соответствие» take without the file dialog

' the tool books of WMS_TOOLBOX: file | title | what for
Function ToolList() As Variant
    ToolList = Array( _
        "WMS_INVENTORY|Инвентаризация|снимок → пересчёт → разница → пакет корректировок для WMS", _
        "WMS_ANALYTICS|Аналитика|остатки, приходы, выдачи, возвраты, просрочка, поставщики, категории, места, получатели", _
        "WMS_MANAGER|Отчёт руководителю|показатели за день / неделю / месяц и журнал работы кладовщика", _
        "WMS_SEARCH|Поиск и история|ЕИ, артикул, наименование, место, получатель; история ЕИ", _
        "WMS_DOCTOR|Диагностика|аудит журнала, остатков, ЕИ, связей, дублей; только отчёт и план", _
        "WMS_LABELS|Этикетки|ЕИ крупно, наименование, артикул, место, штрихкод; печать через системный принтер", _
        "WMS_DOCS|Акты|акты передачи (ODT, PDF) из выдач WMS", _
        "WMS_ARCHIVE|Архив|месяцы истории, архивные пакеты, контроль размера рабочей книги", _
        "WMS_IMPORTER|Массовый приход|таблица поставщика → проверка → пакет «Иной приход» / «Заказы» для WMS")
End Function

' the scripts (python3, next to the books): folder/file | title | the arguments shown to the user
Function ScriptList() As Variant
    ScriptList = Array( _
        "reconcile/wms_reconcile.py|Сверка старых названий|НАЗВАНИЯ.csv --snapshot <снимок> --out <отчёт.csv>", _
        "backup/wms_healthcheck.py|Проверка перед сменой|--wms <папка WMS>", _
        "backup/wms_backup.py|Резервная копия на носитель|--wms <папка WMS> --to <папка носителя>", _
        "labels/tspl_labels.py|Этикетки на термопринтер (TSPL)|--snapshot <снимок> --ei 1,2,3 --out labels.prn")
End Function

Sub BtnLnRefresh(Optional oEvent As Variant)
    TbMsg(LnRefresh())
End Sub

' the list: every tool book and script present or missing, the latest snapshot, the batches
Function LnRefresh() As String
    Dim sh As Object, sfa As Object, t As Variant, s As Variant, i As Integer, f As Variant, r As Long, nMiss As Integer, snap As String
    Dim created As String, why As String, nb As Long, lst As Variant, bd As String
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    sh = ThisComponent.Sheets.getByName(SH_LIST)
    sh.getCellRangeByPosition(0, FIRST_ROW, 3, FIRST_ROW + 40).clearContents(1023)
    t = ToolList()
    r = FIRST_ROW
    For i = 0 To UBound(t)
        f = Split(t(i), "|")
        sh.getCellRangeByPosition(0, r, 3, r).setDataArray(Array(Array(f(0) & ".ods", f(1), f(2), IIf(sfa.exists(TbFolder() & f(0) & ".ods"), "есть", "НЕТ ФАЙЛА"))))
        If Not sfa.exists(TbFolder() & f(0) & ".ods") Then nMiss = nMiss + 1
        r = r + 1
    Next i
    s = ScriptList()
    For i = 0 To UBound(s)
        f = Split(s(i), "|")
        sh.getCellRangeByPosition(0, r, 3, r).setDataArray(Array(Array(f(0), f(1), "python3 " & f(0) & " " & f(2), IIf(sfa.exists(TbFolder() & f(0)), "есть", "НЕТ ФАЙЛА"))))
        If Not sfa.exists(TbFolder() & f(0)) Then nMiss = nMiss + 1
        r = r + 1
    Next i
    snap = TbLatestSnapshot()
    If snap = "" Then
        sh.getCellByPosition(1, 0).setString("снимков нет в " & ConvertFromURL(TbExportDir()) & " — в WMS: «Главная» → «Экспорт для инструментов»")
    Else
        why = TbVerifySnapshot(snap)
        created = TbManifestValue(snap, "created")
        sh.getCellByPosition(1, 0).setString(Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1, 80) & " (" & Replace(created, "T", " ") & ", LAST_SEQ " _
            & TbManifestValue(snap, "last_seq") & ")" & IIf(why <> "", " — НЕ ЦЕЛ: " & why, " — проверен (SHA-256)"))
    End If
    bd = TbBatchDir()
    If sfa.exists(bd) Then
        lst = sfa.getFolderContents(bd, False)
        nb = UBound(lst) + 1
    End If
    sh.getCellByPosition(1, 1).setString("пакетов для WMS: " & nb & " в " & ConvertFromURL(bd))
    LnRefresh = IIf(nMiss = 0, "OK:", "ERR:") & "инструментов " & (UBound(t) + 1) & ", скриптов " & (UBound(s) + 1) & IIf(nMiss > 0, "; нет файлов: " & nMiss, "") _
        & "; " & IIf(snap = "", "снимков нет", "последний снимок " & Replace(created, "T", " "))
    Exit Function
EH:
    LnRefresh = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' opens a tool book of this folder by its name (WMS_INVENTORY …)
Function LnOpen(sTool As String) As String
    Dim url As String, doc As Object, args(0) As New com.sun.star.beans.PropertyValue
    url = TbFolder() & sTool & ".ods"
    If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(url) Then
        LnOpen = "ERR:нет файла " & ConvertFromURL(url)
        Exit Function
    End If
    args(0).Name = "MacroExecutionMode"
    args(0).Value = 4
    doc = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    If IsNull(doc) Then
        LnOpen = "ERR:не открылся " & ConvertFromURL(url)
    Else
        LnOpen = "OK:" & sTool
    End If
End Function

Private Sub OpenTool(k As Integer)
    Dim res As String
    res = LnOpen(Split(ToolList()(k), "|")(0))
    If Left(res, 3) <> "OK:" Then TbMsg(res)
End Sub

Sub BtnOpen0(Optional oEvent As Variant)
    OpenTool(0)
End Sub

Sub BtnOpen1(Optional oEvent As Variant)
    OpenTool(1)
End Sub

Sub BtnOpen2(Optional oEvent As Variant)
    OpenTool(2)
End Sub

Sub BtnOpen3(Optional oEvent As Variant)
    OpenTool(3)
End Sub

Sub BtnOpen4(Optional oEvent As Variant)
    OpenTool(4)
End Sub

Sub BtnOpen5(Optional oEvent As Variant)
    OpenTool(5)
End Sub

Sub BtnOpen6(Optional oEvent As Variant)
    OpenTool(6)
End Sub

Sub BtnOpen7(Optional oEvent As Variant)
    OpenTool(7)
End Sub

Sub BtnOpen8(Optional oEvent As Variant)
    OpenTool(8)
End Sub

' the report of wms_healthcheck.py for a message: the summary, the header, the lines that are not OK
Private Function HealthSummary(res As String) As String
    Dim ln As Variant, i As Long, s As String, head As String
    ln = Split(Replace(res, Chr(13), ""), Chr(10))
    For i = 0 To UBound(ln)
        If Left(ln(i), 4) = "ИТОГ" Then
            head = ln(i)
        ElseIf Left(ln(i), 2) <> "OK" And Trim(ln(i)) <> "" And Left(ln(i), 12) <> "ПРОВЕРКА WMS" Then
            s = s & Chr(10) & ln(i)
        End If
    Next i
    HealthSummary = head & Chr(10) & ln(0) & s
End Function

' «Проверка перед сменой»: backup/wms_healthcheck.py for the WMS folder of «Настройки» (B6); the report opens as text
Sub BtnLnHealth(Optional oEvent As Variant)
    TbMsg(LnHealth())
End Sub

Function LnHealth() As String
    Dim wms As String, sOut As String, sfa As Object, cmd As String, res As String
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    wms = TbResolve(IIf(TbSetting(5) <> "", TbSetting(5), ".."))
    sOut = TbFolder() & "healthcheck_last.txt"
    If sfa.exists(sOut) Then sfa.kill(sOut)
    cmd = "python3 '" & ConvertFromURL(TbFolder() & "backup/wms_healthcheck.py") & "' --wms '" & ConvertFromURL(wms) & "' --out '" & ConvertFromURL(sOut) & "'"
    Shell("/bin/sh", 0, "-c """ & cmd & """", True)
    If Not sfa.exists(sOut) Then
        LnHealth = "ERR:проверка не выполнилась (нужен python3): " & cmd
        Exit Function
    End If
    res = TbReadText(sOut)
    ' the summary line «ИТОГ: ошибок N, предупреждений M» (InStr of Basic ignores the case: «ОШИБКА» would match «ошибок»)
    ' first, then only the lines that are not OK; the whole report stays in the file
    LnHealth = IIf(InStr(1, res, "ИТОГ: ошибок 0,", 0) > 0, "OK:", "ERR:") & HealthSummary(res) & Chr(10) & "Отчёт целиком: " & ConvertFromURL(sOut)
    Exit Function
EH:
    LnHealth = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ WMS_RECONCILE (reconcile/wms_reconcile.py)

Private Function PickFile(sTitle As String) As String
    Dim fp As Object
    If gLnFile <> "" Then
        PickFile = gLnFile
        gLnFile = ""
        Exit Function
    End If
    fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    fp.setTitle(sTitle)
    fp.appendFilter("Таблица CSV / текст", "*.csv;*.txt")
    If fp.execute() = 1 Then PickFile = fp.getFiles()(0)
End Function

' python3 reconcile/wms_reconcile.py with the arguments; the output into a log file; "" or what failed
Private Function RunReconcile(sArgs As String, sLog As String) As String
    Dim sfa As Object, cmd As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(sLog) Then sfa.kill(sLog)
    cmd = "python3 '" & ConvertFromURL(TbFolder() & "reconcile/wms_reconcile.py") & "' " & sArgs & " > '" & ConvertFromURL(sLog) & "' 2>&1"
    Shell("/bin/sh", 0, "-c """ & cmd & """", True)
    If Not sfa.exists(sLog) Then RunReconcile = "сверка не выполнилась (нужен python3): " & cmd
End Function

' a CSV of the reconcile opened in Calc for the person («;», UTF-8, every column as text)
Private Sub OpenCsv(url As String)
    Dim args(1) As New com.sun.star.beans.PropertyValue
    If gTbAuto Then Exit Sub
    args(0).Name = "FilterName"
    args(0).Value = "Text - txt - csv (StarCalc)"
    args(1).Name = "FilterOptions"
    args(1).Value = "59,34,76,1,1/2/2/2/3/2/4/2/5/2/6/2/7/2/8/2/9/2/10/2/11/2/12/2"
    StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
End Sub

' a system path with the suffix before «.csv» (the other extension is dropped): /x/names.txt → /x/names<suffix>.csv
Private Function WithSuffix(sPath As String, sSuffix As String) As String
    Dim p As Long
    p = TbInStrRev(sPath, ".")
    If p > TbInStrRev(sPath, "/") Then WithSuffix = Left(sPath, p - 1) & sSuffix & ".csv" Else WithSuffix = sPath & sSuffix & ".csv"
End Function

Sub BtnLnReconcile(Optional oEvent As Variant)
    Dim f As String
    f = PickFile("Файл старых названий (CSV / TXT)")
    If f <> "" Then TbMsg(LnReconcile(f))
End Sub

' «Сверка названий»: the file of old names → <файл>_сверка.csv (up to 5 candidate EIs each) → opened for the marks «да»
Function LnReconcile(sNames As String) As String
    Dim sOut As String, sLog As String, why As String, res As String
    On Error GoTo EH
    ' the files by their system paths (a URL may carry the Cyrillic letters encoded)
    sOut = WithSuffix(ConvertFromURL(sNames), "_сверка")
    sLog = TbFolder() & "reconcile_last.txt"
    why = RunReconcile("'" & ConvertFromURL(sNames) & "' --export '" & ConvertFromURL(TbExportDir()) & "' --out '" & sOut & "'", sLog)
    If why <> "" Then
        LnReconcile = "ERR:" & why
        Exit Function
    End If
    res = Trim(Replace(TbReadText(sLog), Chr(10), " "))
    If Left(res, 6) = "ОШИБКА" Or Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(ConvertToURL(sOut)) Then
        LnReconcile = "ERR:" & res
        Exit Function
    End If
    OpenCsv(ConvertToURL(sOut))
    LnReconcile = "OK:" & res & ". Отметьте «да» в колонке «подтвердить» у верного ЕИ каждого названия, сохраните файл (CSV) и нажмите " _
        & "«Соответствие по сверке»"
    Exit Function
EH:
    LnReconcile = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnLnConfirm(Optional oEvent As Variant)
    Dim f As String
    f = PickFile("Отчёт сверки с отметками «да»")
    If f <> "" Then TbMsg(LnConfirm(f))
End Sub

' «Соответствие по сверке»: the report with the marks → <отчёт>_соответствие.csv (old name → confirmed EI) → opened
Function LnConfirm(sReport As String) As String
    Dim sOut As String, sLog As String, why As String, res As String, bad As Boolean, src As String
    On Error GoTo EH
    src = ConvertFromURL(sReport)
    ' names_сверка.csv → names_соответствие.csv; any other name → <name>_соответствие.csv
    If Right(src, Len("_сверка.csv")) = "_сверка.csv" Then
        sOut = Left(src, Len(src) - Len("_сверка.csv")) & "_соответствие.csv"
    Else
        sOut = WithSuffix(src, "_соответствие")
    End If
    sLog = TbFolder() & "reconcile_last.txt"
    why = RunReconcile("--confirm '" & src & "' --export '" & ConvertFromURL(TbExportDir()) & "' --out '" & sOut & "'", sLog)
    If why <> "" Then
        LnConfirm = "ERR:" & why
        Exit Function
    End If
    res = Trim(Replace(TbReadText(sLog), Chr(10), " "))
    If Left(res, 6) = "ОШИБКА" Or Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(ConvertToURL(sOut)) Then
        LnConfirm = "ERR:" & res
        Exit Function
    End If
    OpenCsv(ConvertToURL(sOut))
    bad = InStr(1, res, "ошибка:", 0) > 0
    LnConfirm = IIf(bad, "ERR:", "OK:") & res & IIf(bad, " — строки «ошибка» в файле соответствия: исправьте отметки в отчёте сверки и повторите", "")
    Exit Function
EH:
    LnConfirm = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Function TestLnFile(url As String) As String
    gLnFile = url
    TestLnFile = "OK"
End Function
