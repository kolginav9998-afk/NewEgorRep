' UpMain — книга WMS_UPGRADE.ods: мастер «Обновление из текущей WMS» (PRIME 0.8, 1.0, 1.0.1; D-091, D-115). Пользователь выбирает свою
' рабочую книгу WMS (0.7.x или 1.0) и новую папку и нажимает кнопки шагов; вся работа — в программе tools/upgrade.py папки выпуска
' (запускается отсюда, её ход показывается на странице). Рабочая книга только читается: новая книга 1.0 собирается в
' новой папке из побайтной копии, сверяется ячейка в ячейку и открывается WMS 1.0; рабочей она становится только по
' кнопке «4. СДЕЛАТЬ РАБОЧЕЙ». Откат — старая книга, она не меняется никогда. Эта книга не часть WMS: её данные — пути и
' результаты шагов (файлы UPGRADE_* новой папки).
Option Explicit

Public Const UP_VERSION = "1.0.2"
Public Const UP_MAIN = "ОБНОВЛЕНИЕ"
Public Const UP_REPORT = "Отчёт"
' «ОБНОВЛЕНИЕ»: rows (0-based) of the paths (column C), the steps (status in column C), the state line
Public Const UM_BOOK = 4
Public Const UM_TO = 5
Public Const UM_STEP1 = 8
Public Const UM_STATE = 13
' the longest a run of the program may take before the book stops waiting (seconds)
Public Const UP_TIMEOUT = 14400

Global gUpBusy As Boolean
Global gUpAuto As Boolean
Global gUpLastMsg As String
Global gUpLog As String
Global gUpConfirm As Integer
Global gUpLastAsk As String

' ================================================================ folders and files

Function UpInStrRev(ByVal s As String, ByVal sFind As String) As Long
    Dim i As Long
    For i = Len(s) - Len(sFind) + 1 To 1 Step -1
        If Mid(s, i, Len(sFind)) = sFind Then
            UpInStrRev = i
            Exit Function
        End If
    Next i
End Function

' the folder of this book with a trailing "/"
Function UpFolder() As String
    Dim u As String
    u = ThisComponent.getURL()
    UpFolder = Left(u, UpInStrRev(u, "/"))
End Function

' a path typed by the user (/home/…, file:///…, ~/…) as a URL; "" when empty
Function UpUrl(ByVal s As String, bDir As Boolean) As String
    s = Trim(s)
    If s = "" Then Exit Function
    If Left(s, 2) = "~/" Then s = Environ("HOME") & Mid(s, 2)
    If Left(s, 8) = "file:///" Then UpUrl = s Else UpUrl = ConvertToURL(s)
    If bDir And Right(UpUrl, 1) <> "/" Then UpUrl = UpUrl & "/"
    If Not bDir And Right(UpUrl, 1) = "/" Then UpUrl = Left(UpUrl, Len(UpUrl) - 1)
End Function

Function UpCell(r As Integer) As String
    UpCell = Trim(ThisComponent.Sheets.getByName(UP_MAIN).getCellByPosition(2, r).getString())
End Function

Sub UpSetCell(r As Integer, s As String)
    ThisComponent.Sheets.getByName(UP_MAIN).getCellByPosition(2, r).setString(s)
End Sub

' the work folder of the wizard (logs of the checks): next to this book
Function UpWork() As String
    Dim sfa As Object
    UpWork = UpFolder() & "UPGRADE_WORK/"
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.isFolder(UpWork) Then sfa.createFolder(UpWork)
End Function

Function UpReadText(url As String) As String
    Dim tis As Object, sfa As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(url) Then Exit Function
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(sfa.openFileRead(url))
    tis.setEncoding("UTF-8")
    UpReadText = tis.readString(Array(), False)
    tis.closeInput()
End Function

' a path for the shell: in single quotes (a path with a quote is refused before)
Function UpQ(url As String) As String
    UpQ = "'" & ConvertFromURL(url) & "'"
End Function

Function UpPathProblem(url As String) As String
    If InStr(ConvertFromURL(url), "'") > 0 Or InStr(ConvertFromURL(url), """") > 0 Then
        UpPathProblem = "в пути " & ConvertFromURL(url) & " есть кавычка — переименуйте папку"
    End If
End Function

' the parent folder of a folder URL (with a trailing "/")
Function UpParent(url As String) As String
    Dim u As String
    u = url
    If Right(u, 1) = "/" Then u = Left(u, Len(u) - 1)
    UpParent = Left(u, UpInStrRev(u, "/"))
End Function

Function UpBaseName(url As String) As String
    Dim u As String
    u = url
    If Right(u, 1) = "/" Then u = Left(u, Len(u) - 1)
    UpBaseName = ConvertFromURL(Mid(u, UpInStrRev(u, "/") + 1))
    If UpBaseName = "" Then UpBaseName = Mid(u, UpInStrRev(u, "/") + 1)
End Function

' a folder that does not exist yet or is empty
Function UpFolderFree(url As String) As Boolean
    Dim sfa As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(url) Then
        UpFolderFree = True
    ElseIf sfa.isFolder(url) Then
        UpFolderFree = (UBound(sfa.getFolderContents(url, True)) < 0)
    End If
End Function

' a free folder for the new WMS: «WMS_WORK» next to this book (the folder of the release, as the transfer of 0.7 made it),
' or «<папка книги>_WMS_<версия>» next to the folder of the book when that is the same folder; with _2, _3 … when taken
Function UpProposeFolder(bookUrl As String) As String
    Dim sBase As String, cand As String, k As Integer
    If UpUrl(ConvertFromURL(UpFolder()) & "WMS_WORK", True) <> UpParent(bookUrl) Then
        sBase = ConvertFromURL(UpFolder()) & "WMS_WORK"
    Else
        sBase = ConvertFromURL(UpParent(UpParent(bookUrl))) & UpBaseName(UpParent(bookUrl)) & "_WMS_" & UP_VERSION
    End If
    cand = UpUrl(sBase, True)
    k = 1
    Do While Not UpFolderFree(cand) Or cand = UpParent(bookUrl)
        k = k + 1
        cand = UpUrl(sBase & "_" & k, True)
        If k > 99 Then Exit Do
    Loop
    UpProposeFolder = cand
End Function

' the working books of WMS (0.7.x, 1.0) near this book: ./WMS_WORK, ../WMS_WORK and <every folder next to this one>/WMS_WORK
' (the release folders keep their working WMS there); the most recently changed one, "" when none
Function UpFindBook(ByRef nFound As Integer) As String
    Dim sfa As Object, c As Variant, i As Long, u As String, best As String, bestT As Double
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    c = Array(UpFolder() & "WMS_WORK/WMS_PROD.ods", UpParent(UpFolder()) & "WMS_WORK/WMS_PROD.ods")
    For i = 0 To UBound(c)
        UpConsider(sfa, c(i), best, bestT, nFound)
    Next i
    On Error Resume Next
    c = sfa.getFolderContents(UpParent(UpFolder()), True)
    For i = 0 To UBound(c)
        If sfa.isFolder(c(i)) Then
            u = c(i)
            If Right(u, 1) <> "/" Then u = u & "/"
            If u <> UpFolder() Then UpConsider(sfa, u & "WMS_WORK/WMS_PROD.ods", best, bestT, nFound)
        End If
    Next i
    UpFindBook = best
End Function

Sub UpConsider(sfa As Object, url As String, ByRef best As String, ByRef bestT As Double, ByRef nFound As Integer)
    Dim d As Object, tt As Double
    If Not sfa.exists(url) Then Exit Sub
    If url = best Then Exit Sub
    d = sfa.getDateTimeModified(url)
    tt = DateSerial(d.Year, d.Month, d.Day) + TimeSerial(d.Hours, d.Minutes, d.Seconds)
    nFound = nFound + 1
    If best = "" Or tt > bestT Then
        best = url
        bestT = tt
    End If
End Sub

' ================================================================ messages (no window in the test mode)

Sub UpMsg(s As String)
    gUpLastMsg = s
    If gUpAuto Then Exit Sub
    If Left(s, 4) = "ERR:" Then
        MsgBox Mid(s, 5), 48, "Обновление WMS — не выполнено"
    ElseIf Left(s, 3) = "OK:" Then
        MsgBox Mid(s, 4), 64, "Обновление WMS"
    Else
        MsgBox s, 64, "Обновление WMS"
    End If
End Sub

Function UpAsk(s As String) As Boolean
    gUpLastAsk = s
    If gUpAuto Then
        UpAsk = (gUpConfirm = 1)
        Exit Function
    End If
    UpAsk = (MsgBox(s, 4 + 32 + 256, "Обновление WMS") = 6)
End Function

Sub UpSetState(s As String)
    UpSetCell(UM_STATE, s)
End Sub

' ================================================================ running the program

' «python3 tools/upgrade.py <args>» next to this book; waits for its end, showing the step from progressUrl. Exit code
' (0 done, 1 a stop by the rules, 2 an error, -1 not started); gUpLog — its output.
Function UpRunEngine(sName As String, sArgs As String, progressUrl As String) As Integer
    Dim work As String, sfa As Object, logU As String, rcU As String, cmd As String, t0 As Double, s As String, eng As String
    Dim p As Variant, waited As Long
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    work = UpWork()
    eng = UpFolder() & "tools/upgrade.py"
    gUpLog = ""
    If Not sfa.exists(eng) Then
        gUpLog = "нет программы обновления " & ConvertFromURL(eng) & " — книга должна лежать в папке выпуска WMS 1.0 (рядом с папкой tools)"
        UpRunEngine = -1
        Exit Function
    End If
    s = UpPathProblem(work) & UpPathProblem(eng)
    If s <> "" Then
        gUpLog = s
        UpRunEngine = -1
        Exit Function
    End If
    logU = work & "engine_" & sName & ".log"
    rcU = work & "engine_" & sName & ".rc"
    If sfa.exists(rcU) Then sfa.kill(rcU)
    cmd = "python3 " & UpQ(eng) & " " & sArgs & " > " & UpQ(logU) & " 2>&1; echo $? > " & UpQ(rcU & ".tmp") & "; mv " & UpQ(rcU & ".tmp") _
        & " " & UpQ(rcU)
    UpSetState("Выполняется …")
    Shell("/bin/sh", 0, "-c """ & cmd & """", False)
    t0 = Now()
    Do
        Wait 300
        waited = waited + 1
        If sfa.exists(rcU) Then Exit Do
        If waited Mod 2 = 0 And progressUrl <> "" Then
            p = Trim(Replace(UpReadText(progressUrl), Chr(10), ""))
            If Len(p) > 20 Then UpSetState("Выполняется: " & Mid(p, 21) & " …  (" & Format((Now() - t0) * 86400, "0") & " с)")
        End If
        If (Now() - t0) * 86400 > UP_TIMEOUT Then
            gUpLog = "программа обновления не закончила работу за " & (UP_TIMEOUT \ 60) & " минут — см. " & ConvertFromURL(logU)
            UpRunEngine = 2
            Exit Function
        End If
    Loop
    s = Trim(Replace(UpReadText(rcU), Chr(10), ""))
    gUpLog = UpReadText(logU)
    If s = "" Then UpRunEngine = 2 Else UpRunEngine = CInt(Val(s))
End Function

' the output of the program without the empty lines (for a message)
Function UpLogText(nMax As Integer) As String
    Dim a As Variant, i As Long, s As String, k As Integer
    a = Split(Replace(gUpLog, Chr(13), ""), Chr(10))
    For i = 0 To UBound(a)
        If Trim(a(i)) <> "" Then
            s = s & IIf(s <> "", Chr(10), "") & Trim(a(i))
            k = k + 1
            If k >= nMax Then Exit For
        End If
    Next i
    UpLogText = s
End Function

' ================================================================ the state of the steps

Function UpStateValue(txt As String, key As String) As String
    Dim a As Variant, i As Long
    a = Split(Replace(txt, Chr(13), ""), Chr(10))
    For i = 0 To UBound(a)
        If Left(a(i), Len(key) + 1) = key & "=" Then
            UpStateValue = Mid(a(i), Len(key) + 2)
            Exit Function
        End If
    Next i
End Function

Function UpWhen(s As String) As String
    ' 2026-09-27T19:17:31 → 27.09.2026 19:17
    If Len(s) >= 16 Then UpWhen = Mid(s, 9, 2) & "." & Mid(s, 6, 2) & "." & Left(s, 4) & " " & Mid(s, 12, 5) Else UpWhen = s
End Function

Function UpStateOf(toUrl As String) As String
    If toUrl <> "" Then UpStateOf = UpReadText(toUrl & "UPGRADE_STATE.txt")
End Function

' the status texts of the steps and the button «СДЕЛАТЬ РАБОЧЕЙ» (only after a successful build in the chosen folder)
Sub UpRefreshState()
    Dim t As String, sh As Object, s As String, ready As Boolean, toU As String, btn As Object, chk As String
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(UP_MAIN)
    toU = UpUrl(UpCell(UM_TO), True)
    t = UpStateOf(toU)
    chk = UpReadText(UpWork() & "last_check.txt")
    If UpStateValue(chk, "book") <> UpCell(UM_BOOK) Or UpStateValue(chk, "time") = "" Then
        s = "— ещё не выполнялась"
    ElseIf UpStateValue(chk, "ok") = "1" Then
        s = "✓ " & UpStateValue(chk, "time") & " — книгу можно обновить"
    Else
        s = "✗ " & UpStateValue(chk, "time") & " — обновлять пока нельзя (причина — в строке «Состояние»)"
    End If
    sh.getCellByPosition(2, UM_STEP1).setString(s)
    s = UpReadText(UpWork() & "last_rehearsal.txt")
    If UpStateValue(s, "book") <> UpCell(UM_BOOK) Or UpStateValue(s, "time") = "" Then
        s = "— по желанию: всё как при обновлении, но во временной папке и без перехода"
    ElseIf UpStateValue(s, "ok") = "1" Then
        s = "✓ " & UpStateValue(s, "time") & " — репетиция прошла: " & UpStateValue(s, "folder")
    Else
        s = "✗ " & UpStateValue(s, "time") & " — репетиция не прошла: отчёт в " & UpStateValue(s, "folder")
    End If
    sh.getCellByPosition(2, UM_STEP1 + 1).setString(s)
    If t = "" Or UpStateValue(t, "kind") = "rehearsal" Then
        s = "— ещё не собиралась в этой папке"
    ElseIf UpStateValue(t, "build_ok") = "1" Then
        s = "✓ " & UpWhen(UpStateValue(t, "build_time")) & " — новая книга собрана и сверена: " & UpStateValue(t, "build_result")
    Else
        s = "✗ " & UpWhen(UpStateValue(t, "build_time")) & " — не собрана (отчёт — кнопка «Открыть отчёт»); выберите новую папку и повторите"
    End If
    sh.getCellByPosition(2, UM_STEP1 + 2).setString(s)
    ready = (t <> "" And UpStateValue(t, "kind") <> "rehearsal" And UpStateValue(t, "build_ok") = "1" And UpStateValue(t, "promote_ok") <> "1")
    If UpStateValue(t, "promote_ok") = "1" Then
        s = "✓ " & UpWhen(UpStateValue(t, "promote_time")) & " — рабочая WMS теперь: " & UpStateValue(t, "promote_book")
    ElseIf ready Then
        s = "… доступно: в конце рабочего дня закройте старую WMS и нажмите"
    Else
        s = "— недоступно: сначала «3. СОБРАТЬ И СВЕРИТЬ» без ошибок"
    End If
    sh.getCellByPosition(2, UM_STEP1 + 3).setString(s)
    btn = sh.getDrawPage().getForms().getByIndex(0).getByName("btnPromote")
    btn.Enabled = ready
    Exit Sub
EH:
End Sub

Sub UpWriteMark(sFile As String, ok As Boolean, folder As String)
    Dim tos As Object, sfa As Object, url As String
    url = UpWork() & sFile
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(url) Then sfa.kill(url)
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(sfa.openFileWrite(url))
    tos.setEncoding("UTF-8")
    tos.writeString("book=" & UpCell(UM_BOOK) & Chr(10) & "time=" & Format(Now(), "DD.MM.YYYY HH:MM") & Chr(10) & "ok=" & IIf(ok, "1", "0") _
        & Chr(10) & "folder=" & folder & Chr(10))
    tos.closeOutput()
End Sub

' ================================================================ the steps

Function UpBegin() As Boolean
    If gUpBusy Then
        UpMsg("ERR:программа обновления уже выполняется — дождитесь окончания")
        Exit Function
    End If
    gUpBusy = True
    UpBegin = True
End Function

' the chosen book: "" and the reason when it is not usable
Function UpBookUrl(ByRef why As String) As String
    Dim s As String, u As String, sfa As Object
    s = UpCell(UM_BOOK)
    If s = "" Then
        why = "не выбрана рабочая книга — нажмите «Выбрать…» рядом с полем «Рабочая книга WMS» (обычно WMS_WORK/WMS_PROD.ods)"
        Exit Function
    End If
    u = UpUrl(s, False)
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(u) Or sfa.isFolder(u) Then
        why = "нет файла " & s & " — проверьте путь (кнопка «Выбрать…»)"
        Exit Function
    End If
    If LCase(Right(u, 4)) <> ".ods" Then
        why = "файл " & s & " — не книга .ods"
        Exit Function
    End If
    If u = ThisComponent.getURL() Then
        why = "выбрана сама книга обновления — выберите рабочую книгу WMS"
        Exit Function
    End If
    why = UpPathProblem(u)
    If why = "" Then UpBookUrl = u
End Function

' the book is open in this LibreOffice — it must be closed first (the program reads a closed book)
Function UpBookOpenHere(u As String) As Boolean
    Dim en As Object, c As Object
    en = StarDesktop.getComponents().createEnumeration()
    Do While en.hasMoreElements()
        c = en.nextElement()
        If HasUnoInterfaces(c, "com.sun.star.frame.XModel") Then
            If c.getURL() = u Then
                UpBookOpenHere = True
                Exit Function
            End If
        End If
    Loop
End Function

' «Выбрать…» рабочую книгу
Sub BtnUpPickBook(Optional oEvent As Variant)
    Dim fp As Object, u As String
    On Error GoTo EH
    fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    fp.initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE))
    fp.setTitle("Рабочая книга WMS (обычно WMS_PROD.ods в папке WMS_WORK)")
    fp.appendFilter("Книга WMS (*.ods)", "*.ods")
    If UpCell(UM_BOOK) <> "" Then fp.setDisplayDirectory(UpParent(UpUrl(UpCell(UM_BOOK), False)))
    If fp.execute() <> 1 Then Exit Sub
    u = fp.getFiles()(0)
    UpChooseBook(ConvertFromURL(u))
    Exit Sub
EH:
    UpMsg("ERR:окно выбора файла не открылось: " & Error$ & " — впишите путь к книге в поле вручную")
End Sub

' the book is chosen (the button or a test): the new folder proposed next to it
Function UpChooseBook(sPath As String) As String
    Dim u As String
    UpSetCell(UM_BOOK, sPath)
    u = UpUrl(sPath, False)
    UpSetCell(UM_TO, ConvertFromURL(UpProposeFolder(u)))
    UpRefreshState()
    UpSetState("Книга выбрана. Новая папка предложена рядом с её папкой — её можно изменить. Дальше — «1. ПРОВЕРИТЬ».")
    UpChooseBook = "OK"
End Function

' «Выбрать…» новую папку
Sub BtnUpPickFolder(Optional oEvent As Variant)
    Dim fp As Object, u As String
    On Error GoTo EH
    fp = CreateUnoService("com.sun.star.ui.dialogs.FolderPicker")
    fp.setTitle("Новая пустая папка для WMS 1.0")
    If UpCell(UM_TO) <> "" Then fp.setDisplayDirectory(UpParent(UpUrl(UpCell(UM_TO), True)))
    If fp.execute() <> 1 Then Exit Sub
    u = fp.getDirectory()
    If Not UpFolderFree(UpUrl(u, True)) Then
        UpMsg("ERR:папка " & ConvertFromURL(u) & " не пуста — новая WMS собирается только в новой (пустой) папке. Выберите другую или " _
            & "создайте пустую в окне выбора.")
        Exit Sub
    End If
    UpSetCell(UM_TO, ConvertFromURL(u))
    UpRefreshState()
    Exit Sub
EH:
    UpMsg("ERR:окно выбора папки не открылось: " & Error$ & " — впишите путь в поле вручную")
End Sub

' 1. ПРОВЕРИТЬ — nothing is written anywhere
Sub BtnUpCheck(Optional oEvent As Variant)
    Dim u As String, why As String, rc As Integer
    If Not UpBegin() Then Exit Sub
    On Error GoTo EH
    u = UpBookUrl(why)
    If u = "" Then
        gUpBusy = False
        UpMsg("ERR:" & why)
        Exit Sub
    End If
    If UpBookOpenHere(u) Then
        gUpBusy = False
        UpMsg("ERR:рабочая книга сейчас открыта — закройте её (Файл → Закрыть) и нажмите «1. ПРОВЕРИТЬ» снова")
        Exit Sub
    End If
    UpSetState("Выполняется: проверка книги (она только читается) …")
    rc = UpRunEngine("check", "check " & UpQ(u), "")
    UpWriteMark("last_check.txt", rc = 0, "")
    UpRefreshState()
    gUpBusy = False
    UpSetState(IIf(rc = 0, "✓ ", "✗ ") & Replace(UpLogText(12), Chr(10), "  •  "))
    Select Case rc
    Case 0
        UpMsg("OK:Проверка пройдена — книгу можно обновить. Книга не изменена." & Chr(10) & Chr(10) & UpLogText(8) & Chr(10) & Chr(10) _
            & "Дальше — «2. РЕПЕТИЦИЯ» (по желанию) или сразу «3. СОБРАТЬ И СВЕРИТЬ».")
    Case 1
        UpMsg("ERR:Обновлять пока нельзя — книга не изменена. Что мешает:" & Chr(10) & Chr(10) & UpLogText(14))
    Case Else
        UpMsg("ERR:Проверка не выполнена:" & Chr(10) & UpLogText(8))
    End Select
    Exit Sub
EH:
    gUpBusy = False
    UpMsg("ERR:внутренняя ошибка книги обновления: " & Error$ & " (строка " & Erl & ")")
End Sub

' 2. РЕПЕТИЦИЯ — the whole build in a folder of its own, never made working
Sub BtnUpRehearsal(Optional oEvent As Variant)
    Dim u As String, why As String, rc As Integer, toU As String
    If Not UpBegin() Then Exit Sub
    On Error GoTo EH
    u = UpBookUrl(why)
    If u = "" Then
        gUpBusy = False
        UpMsg("ERR:" & why)
        Exit Sub
    End If
    If UpBookOpenHere(u) Then
        gUpBusy = False
        UpMsg("ERR:рабочая книга сейчас открыта — закройте её и повторите")
        Exit Sub
    End If
    toU = UpWork() & "репетиция_" & Format(Now(), "YYYYMMDD-HHMMSS") & "/"
    why = UpPathProblem(toU)
    If why <> "" Then
        gUpBusy = False
        UpMsg("ERR:" & why)
        Exit Sub
    End If
    rc = UpRunEngine("rehearsal", "build " & UpQ(u) & " --to " & UpQ(toU) & " --rehearsal", toU & "progress.txt")
    UpWriteMark("last_rehearsal.txt", rc = 0, ConvertFromURL(toU))
    UpLoadReport(toU)
    UpRefreshState()
    gUpBusy = False
    UpSetState(IIf(rc = 0, "✓ Репетиция прошла: ", "✗ Репетиция не прошла: ") & Replace(UpLogText(3), Chr(10), "  •  "))
    If rc = 0 Then
        UpMsg("OK:Репетиция прошла: новая книга собрана, сверена ячейка в ячейку и открыта WMS 1.0 без ошибок." & Chr(10) & Chr(10) _
            & "Папка репетиции: " & ConvertFromURL(toU) & Chr(10) & "Рабочей она не становится; её можно удалить." & Chr(10) _
            & "Ваша книга не изменена. Отчёт — на листе «" & UP_REPORT & "».")
    Else
        UpMsg("ERR:Репетиция не прошла — ваша книга не изменена." & Chr(10) & Chr(10) & UpLogText(6) & Chr(10) & Chr(10) _
            & "Отчёт — на листе «" & UP_REPORT & "».")
    End If
    Exit Sub
EH:
    gUpBusy = False
    UpMsg("ERR:внутренняя ошибка книги обновления: " & Error$ & " (строка " & Erl & ")")
End Sub

' 3. СОБРАТЬ И СВЕРИТЬ — the new book in the new folder
Sub BtnUpBuild(Optional oEvent As Variant)
    Dim u As String, why As String, rc As Integer, toU As String, t As String, nu As String
    If Not UpBegin() Then Exit Sub
    On Error GoTo EH
    u = UpBookUrl(why)
    If u = "" Then
        gUpBusy = False
        UpMsg("ERR:" & why)
        Exit Sub
    End If
    If UpBookOpenHere(u) Then
        gUpBusy = False
        UpMsg("ERR:рабочая книга сейчас открыта — закройте её и повторите")
        Exit Sub
    End If
    toU = UpUrl(UpCell(UM_TO), True)
    If toU = "" Then
        toU = UpProposeFolder(u)
        UpSetCell(UM_TO, ConvertFromURL(toU))
    End If
    If Not UpFolderFree(toU) Then
        t = UpStateOf(toU)
        If UpStateValue(t, "promote_ok") = "1" Then
            gUpBusy = False
            UpMsg("ERR:в папке " & ConvertFromURL(toU) & " уже рабочая WMS 1.0 (переход " & UpWhen(UpStateValue(t, "promote_time")) _
                & "). Повторно собирать не нужно — работайте в ней.")
            Exit Sub
        End If
        nu = UpProposeFolder(u)
        If Not UpAsk("Папка " & ConvertFromURL(toU) & " не пуста" & IIf(t <> "", " (в ней прошлая сборка)", "") & "." & Chr(10) & Chr(10) _
            & "Новая книга собирается только в пустую папку; в этой ничего не удаляется. Собрать в новой папке" & Chr(10) _
            & ConvertFromURL(nu) & " ?") Then
            gUpBusy = False
            UpMsg("Ничего не изменено.")
            Exit Sub
        End If
        toU = nu
        UpSetCell(UM_TO, ConvertFromURL(toU))
    End If
    why = UpPathProblem(toU)
    If why <> "" Then
        gUpBusy = False
        UpMsg("ERR:" & why)
        Exit Sub
    End If
    rc = UpRunEngine("build", "build " & UpQ(u) & " --to " & UpQ(toU), toU & "progress.txt")
    UpLoadReport(toU)
    UpRefreshState()
    gUpBusy = False
    UpSetState(IIf(rc = 0, "✓ ", "✗ ") & Replace(UpLogText(3), Chr(10), "  •  "))
    If rc = 0 Then
        UpMsg("OK:Новая книга WMS 1.0 собрана и сверена:" & Chr(10) & "• каждая ячейка каждого листа совпадает со старой книгой;" & Chr(10) _
            & "• журнал скопирован побайтно;" & Chr(10) & "• WMS 1.0 открыла её: «работа разрешена», самопроверка без ошибок, остатки = движения;" _
            & Chr(10) & "• ваша книга не изменена." & Chr(10) & Chr(10) & "До перехода работайте в старой книге как обычно. В конце дня закройте её " _
            & "и нажмите «4. СДЕЛАТЬ РАБОЧЕЙ». Если до этого проведёте операции в старой книге — мастер попросит собрать заново (это быстро и " _
            & "безопасно).")
    ElseIf rc = 1 Then
        UpMsg("ERR:Новая книга не собрана — ваша книга не изменена. Что мешает:" & Chr(10) & Chr(10) & UpLogText(10))
    Else
        UpMsg("ERR:Новая книга не собрана — ваша книга не изменена, незаконченная копия удалена:" & Chr(10) & Chr(10) & UpLogText(6) _
            & Chr(10) & Chr(10) & "Отчёт — на листе «" & UP_REPORT & "».")
    End If
    Exit Sub
EH:
    gUpBusy = False
    UpMsg("ERR:внутренняя ошибка книги обновления: " & Error$ & " (строка " & Erl & ")")
End Sub

' 4. СДЕЛАТЬ РАБОЧЕЙ
Sub BtnUpPromote(Optional oEvent As Variant)
    Dim toU As String, rc As Integer, t As String, nb As String, u As String, why As String
    If Not UpBegin() Then Exit Sub
    On Error GoTo EH
    toU = UpUrl(UpCell(UM_TO), True)
    t = UpStateOf(toU)
    If t = "" Or UpStateValue(t, "build_ok") <> "1" Or UpStateValue(t, "kind") = "rehearsal" Then
        gUpBusy = False
        UpMsg("ERR:в выбранной папке нет собранной новой книги — сначала «3. СОБРАТЬ И СВЕРИТЬ»")
        Exit Sub
    End If
    u = UpBookUrl(why)
    If u <> "" Then
        If UpBookOpenHere(u) Then
            gUpBusy = False
            UpMsg("ERR:старая WMS сейчас открыта — закройте её (Файл → Закрыть) и нажмите «4. СДЕЛАТЬ РАБОЧЕЙ» снова")
            Exit Sub
        End If
    End If
    nb = UpStateValue(t, "build_result")
    If Not UpAsk("Сделать новую книгу рабочей?" & Chr(10) & Chr(10) & "Новая рабочая WMS: " & nb & Chr(10) & Chr(10) _
        & "• После этого работайте только в ней (её можно открыть сразу)." & Chr(10) _
        & "• Старая книга остаётся как есть — это откат; в ней больше не работайте: операции в старой книге в новую не попадут " _
        & "(новая WMS предупредит при запуске, если такое случится).") Then
        gUpBusy = False
        UpMsg("Ничего не изменено.")
        Exit Sub
    End If
    rc = UpRunEngine("promote", "promote --to " & UpQ(toU), "")
    UpLoadReport(toU)
    UpRefreshState()
    gUpBusy = False
    UpSetState(IIf(rc = 0, "✓ ", "✗ ") & Replace(UpLogText(3), Chr(10), "  •  "))
    If rc = 0 Then
        If UpAsk("Готово: рабочая WMS теперь" & Chr(10) & nb & Chr(10) & Chr(10) & "Старая книга не изменена. Рядом с новой книгой — файл " _
            & "«ПЕРЕХОД_ВЫПОЛНЕН.txt»." & Chr(10) & Chr(10) & "Открыть новую WMS сейчас?") Then
            UpOpenBook(ConvertToURL(nb))
        End If
        gUpLastMsg = "OK:переход выполнен: " & nb
    Else
        UpMsg("ERR:Переход не выполнен — рабочей остаётся старая книга, ничего не потеряно:" & Chr(10) & Chr(10) & UpLogText(6))
    End If
    Exit Sub
EH:
    gUpBusy = False
    UpMsg("ERR:внутренняя ошибка книги обновления: " & Error$ & " (строка " & Erl & ")")
End Sub

Sub UpOpenBook(url As String)
    Dim args(0) As New com.sun.star.beans.PropertyValue
    args(0).Name = "MacroExecutionMode"
    args(0).Value = 4
    StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
End Sub

' «Открыть новую WMS» — after the switch only (before it the old book is the working one)
Sub BtnUpOpenNew(Optional oEvent As Variant)
    Dim t As String
    t = UpStateOf(UpUrl(UpCell(UM_TO), True))
    If UpStateValue(t, "promote_ok") <> "1" Then
        UpMsg("ERR:новая WMS ещё не рабочая — до «4. СДЕЛАТЬ РАБОЧЕЙ» работайте в старой книге")
        Exit Sub
    End If
    UpOpenBook(ConvertToURL(UpStateValue(t, "promote_book")))
End Sub

' ================================================================ the report (UPGRADE_REPORT.md of the folder) on «Отчёт»

Sub UpLoadReport(toU As String)
    Dim txt As String, a As Variant, i As Long, sh As Object, rows() As Variant, n As Long, s As String, c As Object
    sh = ThisComponent.Sheets.getByName(UP_REPORT)
    sh.getCellRangeByPosition(0, 2, 0, 3000).clearContents(com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.VALUE)
    sh.getCellRangeByPosition(0, 2, 0, 3000).CharWeight = 100
    sh.getCellRangeByPosition(0, 2, 0, 3000).CharHeight = 10
    txt = UpReadText(toU & "UPGRADE_REPORT.md")
    If txt = "" Then
        sh.getCellByPosition(0, 2).setString("Отчёта ещё нет: он появляется после «2. РЕПЕТИЦИЯ» или «3. СОБРАТЬ И СВЕРИТЬ».")
        Exit Sub
    End If
    sh.getCellByPosition(0, 1).setString("Папка: " & ConvertFromURL(toU) & "   (полный текст — файл UPGRADE_REPORT.md в ней)")
    a = Split(Replace(txt, Chr(13), ""), Chr(10))
    For i = 0 To UBound(a)
        If n > 2900 Then Exit For
        s = Replace(Replace(a(i), "**", ""), "`", "")
        c = sh.getCellByPosition(0, 2 + n)
        If Left(s, 2) = "# " Then
            c.setString(Mid(s, 3))
            c.CharWeight = 150
            c.CharHeight = 14
        ElseIf Left(s, 3) = "## " Then
            c.setString(Mid(s, 4))
            c.CharWeight = 150
            c.CharHeight = 11
        ElseIf Left(s, 2) = "- " Then
            c.setString("•  " & Mid(s, 3))
        ElseIf Left(s, 4) = "  - " Then
            c.setString("        –  " & Mid(s, 5))
        Else
            c.setString(s)
        End If
        n = n + 1
    Next i
End Sub

Sub BtnUpReport(Optional oEvent As Variant)
    Dim toU As String, sh As Object, r As String
    toU = UpUrl(UpCell(UM_TO), True)
    If UpReadText(toU & "UPGRADE_REPORT.md") = "" Then
        r = UpStateValue(UpReadText(UpWork() & "last_rehearsal.txt"), "folder")
        If r <> "" Then toU = UpUrl(r, True)
    End If
    UpLoadReport(toU)
    sh = ThisComponent.Sheets.getByName(UP_REPORT)
    ThisComponent.getCurrentController().setActiveSheet(sh)
    ThisComponent.getCurrentController().select(sh.getCellRangeByName("A1"))
End Sub

Sub BtnUpBack(Optional oEvent As Variant)
    ThisComponent.getCurrentController().setActiveSheet(ThisComponent.Sheets.getByName(UP_MAIN))
End Sub

' the document was opened: the working book of 0.7 proposed when one lies near (the folders of the releases 0.7.x)
Sub UpOnLoad(Optional oEvent As Variant)
    Dim cand As String, n As Integer
    If Not ThisComponent.hasLocation() Then Exit Sub
    On Error GoTo EH
    If UpCell(UM_BOOK) = "" Then
        cand = UpFindBook(n)
        If cand <> "" Then
            UpChooseBook(ConvertFromURL(cand))
            If n > 1 Then UpSetState("Найдено рабочих книг WMS рядом: " & n & ". Предложена самая свежая — проверьте путь (или «Выбрать…»). " _
                & "Дальше — «1. ПРОВЕРИТЬ».")
        End If
    End If
    UpRefreshState()
    Exit Sub
EH:
End Sub

' ================================================================ test seams (the GUI and the tests drive the book without windows)

Function UpTestAuto(b As Boolean) As String
    gUpAuto = b
    UpTestAuto = "OK"
End Function

Function UpTestConfirm(n As Integer) As String
    gUpConfirm = n
    UpTestConfirm = "OK"
End Function

Function UpTestLastMessage() As String
    UpTestLastMessage = gUpLastMsg
End Function

Function UpTestLastAsk() As String
    UpTestLastAsk = gUpLastAsk
End Function

Function UpTestSetFolder(sPath As String) As String
    UpSetCell(UM_TO, sPath)
    UpRefreshState()
    UpTestSetFolder = "OK"
End Function
