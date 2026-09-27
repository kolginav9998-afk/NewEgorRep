' LtMain — книга WMS_LEGACY_TRANSFER.ods: перенос старой таблицы «Заказы» в новую WMS 0.7 (задание M7 PRIME).
' Пользователь вставляет старый лист «Заказы» (A:AB) одним Ctrl+V и нажимает кнопки шагов; вся работа — в программе
' tools/legacy_transfer.py папки выпуска (запускается отсюда, её ход показывается на главной странице). Книга не часть
' боевой WMS: она не открывает ни старую таблицу, ни рабочую WMS; её данные — копия вставленного, её результат — новая
' рабочая WMS в отдельной папке. Перед каждым шагом формулы вставленных листов заменяются их значениями (ссылки на старую
' книгу не переносятся), копия книги сохраняется в рабочую папку переноса — программа читает только её.
Option Explicit

Public Const LT_VERSION = "1.0.0"
Public Const LT_MAIN = "LEGACY TRANSFER"
Public Const LT_ORDERS = "1_Вставить_Заказы"
Public Const LT_STOCK = "1B_Вставить_Наличие"
Public Const LT_ISSUES = "1C_Вставить_Выдачи"
Public Const LT_RETURNS = "1D_Вставить_Возвраты"
Public Const LT_CHECK = "2_Проверка"
Public Const LT_ERRORS = "3_Ошибки"
Public Const LT_VERIFY = "4_Сверка"
Public Const LT_SETTINGS = "Настройки"
' «Настройки»: rows (0-based) of column B
Public Const LS_CANDIDATE = 1
Public Const LS_WORK = 2
Public Const LS_PROD = 3
Public Const LS_SOURCE = 4
Public Const LS_UNIT = 5
Public Const LS_PLACE = 6
Public Const LS_BALANCE = 7
Public Const LS_NOQTY = 8
Public Const LS_NODATE = 9
Public Const LS_STOCKONLY = 10
Public Const LS_PRICE = 11
Public Const LS_EMPTY = 12
Public Const LS_UNITS = 13
Public Const LS_PLACES = 14
' «LEGACY TRANSFER»: the status cells (column C) of the steps, the message line
Public Const LM_STEP1 = 3
Public Const LM_STATE = 9
Public Const LT_RESULT_COL = 29
' the longest a run of the program may take before the book stops waiting (seconds)
Public Const LT_TIMEOUT = 14400

' a Global keeps its value between macro calls (a module-level Private does not)
Global gLtBusy As Boolean
Global gLtAuto As Boolean
Global gLtLastMsg As String
Global gLtLog As String
Global gLtConfirm As Integer
Global gLtLastAsk As String

' ================================================================ folders and files

Function LtInStrRev(ByVal s As String, ByVal sFind As String) As Long
    Dim i As Long
    For i = Len(s) - Len(sFind) + 1 To 1 Step -1
        If Mid(s, i, Len(sFind)) = sFind Then
            LtInStrRev = i
            Exit Function
        End If
    Next i
End Function

' the folder of this book with a trailing "/"
Function LtFolder() As String
    Dim u As String
    u = ThisComponent.getURL()
    LtFolder = Left(u, LtInStrRev(u, "/"))
End Function

' a path of a setting: absolute (file:///… or /…) or relative to this book; bDir — a folder (with a trailing "/")
Function LtResolve(ByVal s As String, bDir As Boolean) As String
    Dim sDir As String, parts As Variant, i As Integer
    s = Trim(s)
    If Left(s, 8) = "file:///" Then
        LtResolve = s
    ElseIf Left(s, 1) = "/" Then
        LtResolve = ConvertToURL(s)
    Else
        sDir = LtFolder()
        parts = Split(Replace(s, "\", "/"), "/")
        For i = 0 To UBound(parts)
            If parts(i) = ".." Then
                sDir = Left(sDir, LtInStrRev(Left(sDir, Len(sDir) - 1), "/"))
            ElseIf parts(i) <> "." And parts(i) <> "" Then
                sDir = sDir & parts(i) & "/"
            End If
        Next i
        LtResolve = sDir
        If Not bDir And Right(LtResolve, 1) = "/" Then LtResolve = Left(LtResolve, Len(LtResolve) - 1)
    End If
    If bDir And Right(LtResolve, 1) <> "/" Then LtResolve = LtResolve & "/"
End Function

Function LtSetting(r As Integer) As String
    LtSetting = Trim(ThisComponent.Sheets.getByName(LT_SETTINGS).getCellByPosition(1, r).getString())
End Function

' the work folder of the transfer (LtWork creates it when missing)
Function LtWorkPath() As String
    Dim s As String
    s = LtSetting(LS_WORK)
    If s = "" Then s = "LEGACY_WORK"
    LtWorkPath = LtResolve(s, True)
End Function

Function LtWork() As String
    Dim sfa As Object
    LtWork = LtWorkPath()
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.isFolder(LtWork) Then sfa.createFolder(LtWork)
End Function

Function LtReadText(url As String) As String
    Dim tis As Object, sfa As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(url) Then Exit Function
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(sfa.openFileRead(url))
    tis.setEncoding("UTF-8")
    LtReadText = tis.readString(Array(), False)
    tis.closeInput()
End Function

Sub LtWriteText(url As String, s As String)
    Dim tos As Object, sfa As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(url) Then sfa.kill(url)
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(sfa.openFileWrite(url))
    tos.setEncoding("UTF-8")
    tos.writeString(s)
    tos.closeOutput()
End Sub

' a path for the shell: in single quotes (a path with a quote is refused before)
Function LtQ(url As String) As String
    LtQ = "'" & ConvertFromURL(url) & "'"
End Function

' ================================================================ messages (no window in the test mode)

Sub LtMsg(s As String)
    gLtLastMsg = s
    If gLtAuto Then Exit Sub
    If Left(s, 4) = "ERR:" Then
        MsgBox Mid(s, 5), 48, "WMS LEGACY TRANSFER — не выполнено"
    ElseIf Left(s, 3) = "OK:" Then
        MsgBox Mid(s, 4), 64, "WMS LEGACY TRANSFER"
    Else
        MsgBox s, 64, "WMS LEGACY TRANSFER"
    End If
End Sub

' Yes (True) / No; in the test mode the answer given by LtTestConfirm
Function LtAsk(s As String) As Boolean
    gLtLastAsk = s
    If gLtAuto Then
        LtAsk = (gLtConfirm = 1)
        Exit Function
    End If
    LtAsk = (MsgBox(s, 4 + 32 + 256, "WMS LEGACY TRANSFER") = 6)
End Function

Sub LtSetState(s As String)
    ThisComponent.Sheets.getByName(LT_MAIN).getCellByPosition(2, LM_STATE).setString(s)
End Sub

' ================================================================ the staging: formulas → values, the copy for the program

' Every formula of the four paste sheets becomes its value (the value Calc shows now): a formula that refers to the old
' book or to another sheet is not carried anywhere. An error result stays visible as the text «#ОШИБКА …» — the check
' reports it. Returns "формул N (со ссылками на другие книги K), ошибок M".
Function LtFreeze() As String
    Dim names As Variant, i As Integer, sh As Object, rr As Object, addrs As Variant, k As Long, r As Long, c As Long
    Dim nF As Long, nE As Long, nX As Long, rng As Object, cell As Object, fa As Variant, j As Long, q As Long
    names = Array(LT_ORDERS, LT_STOCK, LT_ISSUES, LT_RETURNS)
    For i = 0 To UBound(names)
        sh = ThisComponent.Sheets.getByName(names(i))
        rr = sh.queryFormulaCells(com.sun.star.sheet.FormulaResult.ERROR)
        addrs = rr.getRangeAddresses()
        For k = 0 To UBound(addrs)
            For r = addrs(k).StartRow To addrs(k).EndRow
                For c = addrs(k).StartColumn To addrs(k).EndColumn
                    cell = sh.getCellByPosition(c, r)
                    cell.setString("#ОШИБКА " & cell.getString())
                    nE = nE + 1
                Next c
            Next r
        Next k
        rr = sh.queryContentCells(com.sun.star.sheet.CellFlags.FORMULA)
        addrs = rr.getRangeAddresses()
        For k = 0 To UBound(addrs)
            rng = sh.getCellRangeByPosition(addrs(k).StartColumn, addrs(k).StartRow, addrs(k).EndColumn, addrs(k).EndRow)
            fa = rng.getFormulaArray()
            For j = 0 To UBound(fa)
                For q = 0 To UBound(fa(j))
                    If InStr(fa(j)(q), "file:") > 0 Then nX = nX + 1
                Next q
            Next j
            nF = nF + (addrs(k).EndRow - addrs(k).StartRow + 1) * (addrs(k).EndColumn - addrs(k).StartColumn + 1)
            rng.setDataArray(rng.getDataArray())
        Next k
    Next i
    LtFreeze = "формул заменено значениями: " & nF & IIf(nX > 0, " (со ссылками на другие книги: " & nX & ")", "") _
        & IIf(nE > 0, ", ячеек с ошибкой формулы: " & nE, "")
End Function

' a copy of this book for the program (the book itself stays where it is, unsaved changes included)
Function LtStoreCopy(url As String) As String
    Dim args(0) As New com.sun.star.beans.PropertyValue
    On Error GoTo EH
    args(0).Name = "FilterName"
    args(0).Value = "calc8"
    ThisComponent.storeToURL(url, args())
    LtStoreCopy = ""
    Exit Function
EH:
    LtStoreCopy = "копия данных не сохранена в " & ConvertFromURL(url) & ": " & Error$
End Function

' a choice of «Настройки» → the word of the program
Function LtRule(r As Integer) As String
    Dim v As String
    v = UCase(LtSetting(r))
    Select Case v
    Case "0"
        LtRule = "zero"
    Case "F", "ПОЗИЦИЯ H=F", "ПОЗИЦИЯ Н=F"
        LtRule = "fact"
    Case "ДАТА ДОКУМЕНТА"
        LtRule = "doc"
    Case "НЕ ПЕРЕНОСИТЬ"
        LtRule = "skip"
    Case "СТАРЫЙ СКЛАД"
        LtRule = "old"
    Case "ОКРУГЛИТЬ"
        LtRule = "round"
    Case "ПРОПУСТИТЬ"
        LtRule = "skip"
    Case Else
        LtRule = "stop"
    End Select
End Function

Function LtWriteSettings(work As String) As String
    Dim s As String, src As String
    src = LtSetting(LS_SOURCE)
    If src <> "" Then src = ConvertFromURL(LtResolve(src, False))
    s = "default_unit=" & LtSetting(LS_UNIT) & Chr(10) & "default_place=" & LtSetting(LS_PLACE) & Chr(10) _
        & "unknown_balance=" & LtRule(LS_BALANCE) & Chr(10) & "no_order_qty=" & LtRule(LS_NOQTY) & Chr(10) _
        & "no_receipt_date=" & LtRule(LS_NODATE) & Chr(10) & "stock_only=" & LtRule(LS_STOCKONLY) & Chr(10) _
        & "price_round=" & LtRule(LS_PRICE) & Chr(10) & "empty_rows=" & LtRule(LS_EMPTY) & Chr(10) _
        & "units=" & LtSetting(LS_UNITS) & Chr(10) & "places=" & LtSetting(LS_PLACES) & Chr(10) & "source_file=" & src & Chr(10)
    LtWriteText(work & "settings.txt", s)
    LtWriteSettings = work & "settings.txt"
End Function

' the paths of the arguments must not contain a quote (the shell command is built from them)
Function LtPathProblem(url As String) As String
    If InStr(ConvertFromURL(url), "'") > 0 Or InStr(ConvertFromURL(url), """") > 0 Then
        LtPathProblem = "в пути " & ConvertFromURL(url) & " есть кавычка — переименуйте папку"
    End If
End Function

' ================================================================ running the program

' runs «python3 tools/legacy_transfer.py <mode> --work … <args>» next to this book and waits for its end, showing its
' progress on the main page (the window stays alive: Wait lets LibreOffice work). The exit code (0 done, 1 a stop by the
' rules, 2 an error of the program, -1 not started); gLtLog — its output.
Function LtRunEngine(sMode As String, sArgs As String) As Integer
    Dim work As String, sfa As Object, logU As String, rcU As String, cmd As String, t0 As Double, s As String, eng As String
    Dim progU As String, p As Variant, waited As Long
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    work = LtWork()
    eng = LtFolder() & "tools/legacy_transfer.py"
    gLtLog = ""
    If Not sfa.exists(eng) Then
        gLtLog = "нет программы переноса " & ConvertFromURL(eng) & " — книга должна лежать в папке выпуска WMS (рядом с папкой tools)"
        LtRunEngine = -1
        Exit Function
    End If
    s = LtPathProblem(work) & LtPathProblem(eng)
    If s <> "" Then
        gLtLog = s
        LtRunEngine = -1
        Exit Function
    End If
    logU = work & "engine_" & sMode & ".log"
    rcU = work & "engine_" & sMode & ".rc"
    progU = work & "progress.txt"
    If sfa.exists(rcU) Then sfa.kill(rcU)
    If sfa.exists(progU) Then sfa.kill(progU)
    cmd = "python3 " & LtQ(eng) & " " & sMode & " --work " & LtQ(work) & " " & sArgs & " > " & LtQ(logU) & " 2>&1; echo $? > " _
        & LtQ(rcU & ".tmp") & "; mv " & LtQ(rcU & ".tmp") & " " & LtQ(rcU)
    LtSetState("Выполняется: запуск программы переноса …")
    Shell("/bin/sh", 0, "-c """ & cmd & """", False)
    t0 = Now()
    Do
        Wait 300
        waited = waited + 1
        If sfa.exists(rcU) Then Exit Do
        If waited Mod 2 = 0 Then
            p = Split(LtReadText(progU) & "|||", "|")
            If Trim(p(3)) <> "" Then LtSetState("Выполняется: " & Trim(Replace(p(3), Chr(10), "")) & " …")
        End If
        If (Now() - t0) * 86400 > LT_TIMEOUT Then
            gLtLog = "программа переноса не закончила работу за " & (LT_TIMEOUT \ 60) & " минут — см. " & ConvertFromURL(logU)
            LtRunEngine = 2
            Exit Function
        End If
    Loop
    s = Trim(Replace(LtReadText(rcU), Chr(10), ""))
    gLtLog = LtReadText(logU)
    If s = "" Then LtRunEngine = 2 Else LtRunEngine = CInt(Val(s))
End Function

' the last lines of the output of the program (for a message)
Function LtLogTail(n As Integer) As String
    Dim a As Variant, i As Long, s As String, k As Integer
    a = Split(Replace(gLtLog, Chr(13), ""), Chr(10))
    For i = UBound(a) To 0 Step -1
        If Trim(a(i)) <> "" Then
            s = Trim(a(i)) & IIf(s <> "", Chr(10) & s, "")
            k = k + 1
            If k >= n Then Exit For
        End If
    Next i
    LtLogTail = s
End Function

' ================================================================ results of the program → the sheets of this book

' a CSV of the program (UTF-8, «;», every column as text) as an array of rows; empty when the file is missing
Function LtCsvRows(url As String, nCols As Integer) As Variant
    Dim args(2) As New com.sun.star.beans.PropertyValue, csv As Object, sh As Object, cur As Object, last As Long, i As Integer, fmt As String
    If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(url) Then
        LtCsvRows = Array()
        Exit Function
    End If
    For i = 1 To nCols
        fmt = fmt & IIf(fmt <> "", "/", "") & i & "/2"
    Next i
    args(0).Name = "FilterName"
    args(0).Value = "Text - txt - csv (StarCalc)"
    args(1).Name = "FilterOptions"
    args(1).Value = "59,34,76,1," & fmt & ",1033,false,true,false,false,false"
    args(2).Name = "Hidden"
    args(2).Value = True
    csv = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    sh = csv.Sheets.getByIndex(0)
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    If last < 1 Then
        LtCsvRows = Array()
    Else
        LtCsvRows = sh.getCellRangeByPosition(0, 1, nCols - 1, last).getDataArray()
    End If
    csv.close(True)
End Function

Sub LtPutRows(sh As Object, c0 As Integer, r0 As Long, rows As Variant, nCols As Integer, clearTo As Long)
    sh.getCellRangeByPosition(c0, r0, c0 + nCols - 1, clearTo).clearContents(com.sun.star.sheet.CellFlags.VALUE _
        + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.DATETIME)
    If UBound(rows) >= 0 Then sh.getCellRangeByPosition(c0, r0, c0 + nCols - 1, r0 + UBound(rows)).setDataArray(rows)
End Sub

' «2_Проверка», «3_Ошибки» and the result columns AD:AE of «1_Вставить_Заказы» from WORK/check/
Sub LtLoadCheck(work As String)
    Dim rows As Variant, sh As Object, i As Long, last As Long, res() As Variant, r As Long, mx As Long
    rows = LtCsvRows(work & "check/summary.csv", 4)
    LtPutRows(ThisComponent.Sheets.getByName(LT_CHECK), 0, 2, rows, 4, 400)
    rows = LtCsvRows(work & "check/errors.csv", 6)
    LtPutRows(ThisComponent.Sheets.getByName(LT_ERRORS), 0, 1, rows, 6, 1048575)
    rows = LtCsvRows(work & "check/rows.csv", 3)
    sh = ThisComponent.Sheets.getByName(LT_ORDERS)
    sh.getCellRangeByPosition(LT_RESULT_COL, 1, LT_RESULT_COL + 1, 1048575).clearContents(com.sun.star.sheet.CellFlags.VALUE _
        + com.sun.star.sheet.CellFlags.STRING)
    sh.getCellByPosition(LT_RESULT_COL, 0).setString("Результат проверки")
    sh.getCellByPosition(LT_RESULT_COL + 1, 0).setString("Что не так / что сделано")
    mx = 0
    For i = 0 To UBound(rows)
        r = CLng(Val(rows(i)(0)))
        If r > mx Then mx = r
    Next i
    If mx < 2 Then Exit Sub
    ReDim res(mx - 2)
    For i = 0 To mx - 2
        res(i) = Array("", "")
    Next i
    For i = 0 To UBound(rows)
        r = CLng(Val(rows(i)(0)))
        If r >= 2 Then res(r - 2) = Array(rows(i)(1), rows(i)(2))
    Next i
    sh.getCellRangeByPosition(LT_RESULT_COL, 1, LT_RESULT_COL + 1, mx - 1).setDataArray(res)
End Sub

Sub LtLoadVerify(work As String)
    LtPutRows(ThisComponent.Sheets.getByName(LT_VERIFY), 0, 2, LtCsvRows(work & "verify_summary.csv", 2), 2, 200)
End Sub

' ================================================================ the state of the steps (WORK/state.txt of the program)

Function LtStateValue(txt As String, key As String) As String
    Dim a As Variant, i As Long
    a = Split(Replace(txt, Chr(13), ""), Chr(10))
    For i = 0 To UBound(a)
        If Left(a(i), Len(key) + 1) = key & "=" Then
            LtStateValue = Mid(a(i), Len(key) + 2)
            Exit Function
        End If
    Next i
End Function

Function LtWhen(s As String) As String
    ' 2026-09-27T19:17:31 → 27.09.2026 19:17
    If Len(s) >= 16 Then LtWhen = Mid(s, 9, 2) & "." & Mid(s, 6, 2) & "." & Left(s, 4) & " " & Mid(s, 12, 5) Else LtWhen = s
End Function

' the status texts of the steps on the main page and the button «СДЕЛАТЬ РАБОЧЕЙ» (available only after a clean check,
' a created test WMS and a reconciliation without critical differences)
Sub LtRefreshState()
    Dim t As String, sh As Object, s As String, ready As Boolean, work As String, btn As Object
    On Error GoTo EH
    work = LtWorkPath()
    t = LtReadText(work & "state.txt")
    sh = ThisComponent.Sheets.getByName(LT_MAIN)
    If LtStateValue(t, "check_time") = "" Then
        s = "— ещё не выполнялась"
    ElseIf LtStateValue(t, "check_ok") = "1" Then
        s = "✓ " & LtWhen(LtStateValue(t, "check_time")) & " — блокирующих замечаний нет, требуют проверки: " & LtStateValue(t, "check_warnings")
    Else
        s = "✗ " & LtWhen(LtStateValue(t, "check_time")) & " — блокирующих замечаний: " & LtStateValue(t, "check_blockers") & " (лист «3_Ошибки»)"
    End If
    sh.getCellByPosition(2, LM_STEP1 + 1).setString(s)
    If LtStateValue(t, "build_time") = "" Then
        s = "— ещё не создавалась"
    ElseIf LtStateValue(t, "build_ok") = "1" Then
        s = "✓ " & LtWhen(LtStateValue(t, "build_time")) & " — " & LtStateValue(t, "build_book")
    Else
        s = "✗ " & LtWhen(LtStateValue(t, "build_time")) & " — не создана (см. состояние ниже)"
    End If
    sh.getCellByPosition(2, LM_STEP1 + 2).setString(s)
    If LtStateValue(t, "verify_time") = "" Then
        s = "— ещё не выполнялась"
    ElseIf LtStateValue(t, "verify_ok") = "1" Then
        s = "✓ " & LtWhen(LtStateValue(t, "verify_time")) & " — критических расхождений 0; отчёт: " & LtStateValue(t, "verify_report")
    Else
        s = "✗ " & LtWhen(LtStateValue(t, "verify_time")) & " — критических расхождений: " & LtStateValue(t, "verify_critical")
    End If
    sh.getCellByPosition(2, LM_STEP1 + 3).setString(s)
    ready = (LtStateValue(t, "promote_ready") = "1")
    If LtStateValue(t, "promote_done") = "1" Then
        s = "✓ " & LtWhen(LtStateValue(t, "promote_time")) & " — рабочая WMS: " & LtStateValue(t, "promote_book")
    ElseIf ready Then
        s = "… доступно: проверка, тестовая WMS и сверка без ошибок"
    Else
        s = "— недоступно: сначала шаги 2, 3 и 4 без ошибок"
    End If
    sh.getCellByPosition(2, LM_STEP1 + 4).setString(s)
    If LtStateValue(t, "message") <> "" Then LtSetState(LtStateValue(t, "message"))
    btn = sh.getDrawPage().getForms().getByIndex(0).getByName("btnPromote")
    btn.Enabled = ready
    Exit Sub
EH:
End Sub

' ================================================================ the steps

' the arguments common to the steps: the staging copy and the rules
Function LtCommonArgs(work As String, bStore As Boolean, ByRef why As String) As String
    Dim fr As String
    If bStore Then
        fr = LtFreeze()
        why = LtStoreCopy(work & "staging.ods")
        If why <> "" Then Exit Function
    End If
    LtCommonArgs = "--staging " & LtQ(work & "staging.ods") & " --settings " & LtQ(LtWriteSettings(work))
    gLtLastMsg = fr
End Function

Function LtCandidate() As String
    Dim s As String
    s = LtSetting(LS_CANDIDATE)
    If s = "" Then s = "WMS_PROD_CANDIDATE.ods"
    LtCandidate = LtResolve(s, False)
End Function

Function LtBegin() As Boolean
    If gLtBusy Then
        LtMsg("ERR:программа переноса уже выполняется — дождитесь окончания")
        Exit Function
    End If
    gLtBusy = True
    LtBegin = True
End Function

' 1. ВСТАВИТЬ ДАННЫЕ: the paste sheet with the cursor in A1 and what to do
Sub BtnLtPaste(Optional oEvent As Variant)
    Dim sh As Object
    sh = ThisComponent.Sheets.getByName(LT_ORDERS)
    ThisComponent.getCurrentController().setActiveSheet(sh)
    ThisComponent.getCurrentController().select(sh.getCellRangeByName("A1"))
    LtMsg("1. Откройте старую таблицу, на листе «Заказы» выделите колонки A:AB целиком (щелчок по заголовку A, Shift + щелчок по AB) " _
        & "и нажмите Ctrl+C." & Chr(10) & "2. Вернитесь сюда: ячейка A1 листа «" & LT_ORDERS & "» уже выбрана — нажмите Ctrl+V." & Chr(10) _
        & "Необязательно: старые листы «Наличие», «Выдачи», «Возвраты» — так же на листы 1B, 1C, 1D." & Chr(10) _
        & "Старая таблица не изменяется. Затем — «2. ПРОВЕРИТЬ».")
End Sub

' 2. ПРОВЕРИТЬ: the dry run; nothing but the work folder changes
Sub BtnLtCheck(Optional oEvent As Variant)
    Dim work As String, args As String, why As String, rc As Integer, fr As String
    If Not LtBegin() Then Exit Sub
    On Error GoTo EH
    work = LtWork()
    LtSetState("Выполняется: подготовка данных …")
    args = LtCommonArgs(work, True, why)
    fr = gLtLastMsg
    If why <> "" Then
        gLtBusy = False
        LtMsg("ERR:" & why)
        Exit Sub
    End If
    rc = LtRunEngine("check", args & " --candidate " & LtQ(LtCandidate()))
    LtLoadCheck(work)
    LtRefreshState()
    ThisComponent.getCurrentController().setActiveSheet(ThisComponent.Sheets.getByName(LT_CHECK))
    gLtBusy = False
    Select Case rc
    Case 0
        LtMsg("OK:Проверка закончена: блокирующих замечаний нет (" & fr & ")." & Chr(10) & LtLogTail(1) & Chr(10) _
            & "Следующий шаг — «3. СОЗДАТЬ ТЕСТОВУЮ WMS».")
    Case 1
        LtMsg("ERR:Есть блокирующие замечания — перенос невозможен, пока они не исправлены." & Chr(10) & LtLogTail(1) & Chr(10) _
            & "Красные строки — на листе «" & LT_ORDERS & "» (колонки AD, AE) и в «" & LT_ERRORS & "». Исправьте строку прямо здесь " _
            & "(или задайте правило в «Настройках») и нажмите «ПРОВЕРИТЬ» снова.")
    Case Else
        LtMsg("ERR:Проверка не выполнена:" & Chr(10) & LtLogTail(6))
    End Select
    Exit Sub
EH:
    gLtBusy = False
    LtMsg("ERR:внутренняя ошибка книги переноса: " & Error$ & " (строка " & Erl & ")")
End Sub

' 3. СОЗДАТЬ ТЕСТОВУЮ WMS: a copy of the clean candidate, the whole transfer in it, the checks
Sub BtnLtBuild(Optional oEvent As Variant)
    Dim work As String, args As String, why As String, rc As Integer
    If Not LtBegin() Then Exit Sub
    On Error GoTo EH
    work = LtWork()
    LtSetState("Выполняется: подготовка данных …")
    args = LtCommonArgs(work, True, why)
    If why <> "" Then
        gLtBusy = False
        LtMsg("ERR:" & why)
        Exit Sub
    End If
    rc = LtRunEngine("build", args & " --candidate " & LtQ(LtCandidate()))
    LtLoadCheck(work)
    LtRefreshState()
    gLtBusy = False
    If rc = 0 Then
        LtMsg("OK:Тестовая WMS создана и проверена (самопроверка, независимая проверка журнала, сохранение и повторное открытие)." _
            & Chr(10) & LtLogTail(1) & Chr(10) & "Следующий шаг — «4. СВЕРИТЬ».")
    Else
        LtMsg("ERR:Тестовая WMS не создана — ничего не изменено:" & Chr(10) & LtLogTail(4))
    End If
    Exit Sub
EH:
    gLtBusy = False
    LtMsg("ERR:внутренняя ошибка книги переноса: " & Error$ & " (строка " & Erl & ")")
End Sub

' 4. СВЕРИТЬ: the test WMS against the pasted data; LEGACY_TRANSFER_REPORT.html / .md
Sub BtnLtVerify(Optional oEvent As Variant)
    Dim work As String, args As String, why As String, rc As Integer
    If Not LtBegin() Then Exit Sub
    On Error GoTo EH
    work = LtWork()
    LtSetState("Выполняется: подготовка данных …")
    args = LtCommonArgs(work, True, why)
    If why <> "" Then
        gLtBusy = False
        LtMsg("ERR:" & why)
        Exit Sub
    End If
    rc = LtRunEngine("verify", args)
    LtLoadVerify(work)
    LtRefreshState()
    ThisComponent.getCurrentController().setActiveSheet(ThisComponent.Sheets.getByName(LT_VERIFY))
    gLtBusy = False
    If rc = 0 Then
        LtMsg("OK:Сверка закончена: критических расхождений нет." & Chr(10) & "Отчёт: " & ConvertFromURL(work) & "LEGACY_TRANSFER_REPORT.html" _
            & Chr(10) & "Теперь доступна кнопка «5. СДЕЛАТЬ РАБОЧЕЙ».")
    Else
        LtMsg("ERR:Сверка не пройдена — «СДЕЛАТЬ РАБОЧЕЙ» недоступно:" & Chr(10) & LtLogTail(4))
    End If
    Exit Sub
EH:
    gLtBusy = False
    LtMsg("ERR:внутренняя ошибка книги переноса: " & Error$ & " (строка " & Erl & ")")
End Sub

' 5. СДЕЛАТЬ РАБОЧЕЙ: the checked test WMS becomes WMS_WORK/WMS_PROD.ods (a new folder; an existing PROD is never touched)
Sub BtnLtPromote(Optional oEvent As Variant)
    Dim work As String, args As String, why As String, rc As Integer, dst As String, s As String
    If Not LtBegin() Then Exit Sub
    On Error GoTo EH
    work = LtWork()
    s = LtSetting(LS_PROD)
    If s = "" Then s = "WMS_WORK"
    dst = LtResolve(s, True)
    If Not LtAsk("Сделать проверенную тестовую WMS рабочей?" & Chr(10) & Chr(10) & "Будет создана папка " & ConvertFromURL(dst) _
        & " с книгой WMS_PROD.ods." & Chr(10) & "С этого момента все новые движения записываются только в новую WMS, старая таблица " _
        & "остаётся архивом только для чтения.") Then
        gLtBusy = False
        LtMsg("Ничего не изменено.")
        Exit Sub
    End If
    args = LtCommonArgs(work, True, why)
    If why <> "" Then
        gLtBusy = False
        LtMsg("ERR:" & why)
        Exit Sub
    End If
    rc = LtRunEngine("promote", args & " --to " & LtQ(dst))
    LtRefreshState()
    gLtBusy = False
    If rc = 0 Then
        LtMsg("OK:Рабочая WMS создана: " & ConvertFromURL(dst) & "WMS_PROD.ods" & Chr(10) & Chr(10) _
            & "С этого момента все новые движения записываются только в новую WMS. Старая таблица остаётся read-only архивом." & Chr(10) _
            & "Рядом: CUTOVER_README.txt, отчёт переноса (LEGACY_TRANSFER), резервная копия выпуска (backup), WMS_TOOLBOX.")
    Else
        LtMsg("ERR:Рабочая WMS не создана:" & Chr(10) & LtLogTail(4))
    End If
    Exit Sub
EH:
    gLtBusy = False
    LtMsg("ERR:внутренняя ошибка книги переноса: " & Error$ & " (строка " & Erl & ")")
End Sub

' the header rows of the paste sheets (as the builder writes them)
Function LtHeaders(sName As String) As Variant
    Select Case sName
    Case LT_ORDERS
        LtHeaders = Array("№ позиции", "Полное наименование товара", "Номер документа", "Номер счёта", "Артикул", "Фактическое количество", _
            "Количество по документу", "Заказанное количество", "Единица измерения", "Цена", "Сумма", "Площадка / Поставщик", "Продавец", _
            "Дата поступления", "Дата документа", "Дата заказа", "Ожидаемая дата поступления", "Покупатель", "Категория", "Кому назначено", _
            "Место хранения", "Внутренний код", "Статус", "Наличие", "Контроль", "Комментарий", "Срок поставки, дней", "Возможный дубль")
    Case LT_STOCK
        LtHeaders = Array("ЕИ", "Наименование", "Артикул", "Единица", "Остаток", "Место", "Категория")
    Case LT_ISSUES
        LtHeaders = Array("ЕИ", "Количество", "Дата", "Кому")
    Case Else
        LtHeaders = Array("ЕИ", "Количество", "Дата", "От кого")
    End Select
End Function

' «Очистить staging»: the pasted data and the results of the check (the work folder with the test WMS stays)
Sub BtnLtClear(Optional oEvent As Variant)
    Dim names As Variant, i As Integer, sh As Object, h As Variant, flags As Long, sfa As Object
    If gLtBusy Then
        LtMsg("ERR:программа переноса выполняется — дождитесь окончания")
        Exit Sub
    End If
    If Not LtAsk("Очистить вставленные данные (листы 1, 1B, 1C, 1D) и результаты проверки?" & Chr(10) _
        & "Старая таблица и созданные WMS не затрагиваются.") Then
        LtMsg("Ничего не изменено.")
        Exit Sub
    End If
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.DATETIME _
        + com.sun.star.sheet.CellFlags.FORMULA + com.sun.star.sheet.CellFlags.ANNOTATION
    names = Array(LT_ORDERS, LT_STOCK, LT_ISSUES, LT_RETURNS)
    For i = 0 To UBound(names)
        sh = ThisComponent.Sheets.getByName(names(i))
        sh.getCellRangeByPosition(0, 0, 60, 1048575).clearContents(flags)
        h = LtHeaders(names(i))
        sh.getCellRangeByPosition(0, 0, UBound(h), 0).setDataArray(Array(h))
    Next i
    ThisComponent.Sheets.getByName(LT_ORDERS).getCellByPosition(LT_RESULT_COL, 0).setString("Результат проверки")
    ThisComponent.Sheets.getByName(LT_ORDERS).getCellByPosition(LT_RESULT_COL + 1, 0).setString("Что не так / что сделано")
    ThisComponent.Sheets.getByName(LT_CHECK).getCellRangeByPosition(0, 2, 3, 400).clearContents(flags)
    ThisComponent.Sheets.getByName(LT_ERRORS).getCellRangeByPosition(0, 1, 5, 200000).clearContents(flags)
    ThisComponent.Sheets.getByName(LT_VERIFY).getCellRangeByPosition(0, 2, 1, 200).clearContents(flags)
    ' the steps start over: until the next check the page shows no result of the old data (the program keeps its own
    ' history in state.json; a transfer already made working stays shown)
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If LtStateValue(LtReadText(LtWorkPath() & "state.txt"), "promote_done") <> "1" And sfa.exists(LtWorkPath() & "state.txt") Then
        sfa.kill(LtWorkPath() & "state.txt")
    End If
    LtRefreshState()
    ThisComponent.Sheets.getByName(LT_MAIN).getDrawPage().getForms().getByIndex(0).getByName("btnPromote").Enabled = False
    LtSetState("Данные очищены — вставьте старый лист «Заказы» (шаг 1)")
    LtMsg("OK:Вставленные данные очищены.")
End Sub

Sub BtnLtErrors(Optional oEvent As Variant)
    Dim sh As Object
    sh = ThisComponent.Sheets.getByName(LT_ERRORS)
    ThisComponent.getCurrentController().setActiveSheet(sh)
    ThisComponent.getCurrentController().select(sh.getCellRangeByName("A2"))
End Sub

' «Экспорт отчёта»: the reports of the transfer copied into a new folder of the work folder
Sub BtnLtExport(Optional oEvent As Variant)
    Dim work As String, sfa As Object, dst As String, files As Variant, i As Integer, n As Integer
    work = LtWork()
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    dst = work & "export_" & Format(Now(), "YYYYMMDD-HHMMSS") & "/"
    files = Array("LEGACY_TRANSFER_REPORT.html", "LEGACY_TRANSFER_REPORT.md", "LEGACY_ROWS.csv", "BUILD_REPORT.md", "check/CHECK_REPORT.md", _
        "check/errors.csv", "check/summary.csv", "check/status_pairs.csv")
    For i = 0 To UBound(files)
        If sfa.exists(work & files(i)) Then
            If n = 0 Then sfa.createFolder(dst)
            sfa.copy(work & files(i), dst & Mid(files(i), LtInStrRev(files(i), "/") + 1))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        LtMsg("ERR:отчётов ещё нет — сначала «2. ПРОВЕРИТЬ»")
    Else
        LtMsg("OK:Отчёты сохранены (" & n & " файлов) в папку" & Chr(10) & ConvertFromURL(dst) & Chr(10) _
            & "Главный отчёт — LEGACY_TRANSFER_REPORT.html (открывается в браузере), если сверка уже была.")
    End If
End Sub

' «Создать резервную копию»: this book (the pasted data and the settings) into WORK/backups/
Sub BtnLtBackup(Optional oEvent As Variant)
    Dim work As String, sfa As Object, url As String, why As String
    work = LtWork()
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.isFolder(work & "backups/") Then sfa.createFolder(work & "backups/")
    url = work & "backups/WMS_LEGACY_TRANSFER_" & Format(Now(), "YYYYMMDD-HHMMSS") & ".ods"
    why = LtStoreCopy(url)
    If why <> "" Then
        LtMsg("ERR:" & why)
    Else
        LtMsg("OK:Резервная копия книги переноса: " & ConvertFromURL(url))
    End If
End Sub

' the document was opened: the states of the steps
Sub LtOnLoad(Optional oEvent As Variant)
    If ThisComponent.hasLocation() Then LtRefreshState()
End Sub

' ================================================================ test seams (the GUI and the tests drive the book without windows)

Function LtTestAuto(b As Boolean) As String
    gLtAuto = b
    LtTestAuto = "OK"
End Function

Function LtTestConfirm(n As Integer) As String
    gLtConfirm = n
    LtTestConfirm = "OK"
End Function

Function LtTestLastMessage() As String
    LtTestLastMessage = gLtLastMsg
End Function

Function LtTestLastAsk() As String
    LtTestLastAsk = gLtLastAsk
End Function
