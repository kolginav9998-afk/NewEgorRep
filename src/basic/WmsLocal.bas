' WmsLocal — 1.1: «Проверить локальность». The book looks at itself: its own macros (the text of every module of the libraries Standard and WmsPlus),
' the links to other files and sheets, the hyperlinks, the formulas that read the internet — and says in words whether anything could leave the computer.
' Read only: nothing is changed. The words that a network call is made of are written in pieces (so that this module, scanned by tools/offline_audit.py,
' is not taken for what it looks for). The same kind of check from outside the book: tools/offline_audit.py. Library WmsPlus.
Option Explicit

' how many places of each kind WMS has, by design (a place that is not in this list is shown as new): the version of Python and the free space (a command of the
' system, the answer goes to a file in the folder of WMS), the shortcut on the desktop, the opening of a photo or of the folder of the documents by the program of the computer
Private Const EXPECT_SHELL = 3
Private Const EXPECT_EXEC = 2
Private Const EXPECT_WSH = 1
Private Const EXPECT_FSO = 1

' the words of a network call (piece|what it would be)
Private Function NetWords() As Variant
    NetWords = Array( _
        "Simple" & "SystemMail|отправка почты", "Simple" & "CommandMail|отправка почты", "Simple" & "MailClient|отправка почты", "Mail" & "Merge|рассылка почты", _
        "com.sun.star." & "mail|почтовая служба LibreOffice", "com.sun.star." & "connection|сетевое соединение UNO", "com.sun.star." & "bridge|мост UNO", _
        "XML" & "Http|HTTP-запрос", "Win" & "Http|HTTP-запрос", "MSXML" & "2|HTTP-запрос", "Internet" & "Explorer|браузер", "Open" & "Hyperlink|открытие ссылки", _
        "Open" & "Url|открытие ссылки", "WEB" & "SERVICE(|функция Calc, читающая интернет", "FILTER" & "XML(|функция Calc, читающая интернет", _
        "http" & "://|адрес в интернете", "https" & "://|адрес в интернете", "ftp" & "://|адрес в интернете", "smb" & "://|сетевая папка")
End Function

' the places of the system (the opposite of the net: a program of this computer is started or asked)
Private Function SysWords() As Variant
    SysWords = Array("Shell" & "(", "System" & "ShellExecute", "WScript" & ".Shell", "Scripting" & ".FileSystemObject")
End Function

Sub BtnLocality(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage(LocalityReport())
End Sub

' one line of the report
Private Function LcLine(ok As Boolean, sWhat As String, sNow As String, ByRef nBad As Integer) As String
    If Not ok Then nBad = nBad + 1
    LcLine = IIf(ok, "✓ ", "✗ ") & sWhat & ": " & sNow & Chr(10)
End Function

' the line number (1…) of the position p of the text t
Private Function LineNo(t As String, p As Long) As Long
    Dim q As Long
    q = InStr(t, Chr(10))
    LineNo = 1
    Do While q > 0 And q < p
        LineNo = LineNo + 1
        q = InStr(q + 1, t, Chr(10))
    Loop
End Function

' true when the position p of the text t (lower case) stands in a line that is only a comment
Private Function InComment(t As String, p As Long) As Boolean
    Dim q As Long, e As Long, ln As String
    ' the line break before p (LibreOffice Basic has no search from the end): back from p, a character at a time
    q = p
    Do While q > 0
        If Mid(t, q, 1) = Chr(10) Then Exit Do
        q = q - 1
    Loop
    e = InStr(p, t, Chr(10))
    If e = 0 Then e = Len(t) + 1
    ln = LTrim(Mid(t, q + 1, e - q - 1))
    InComment = (Left(ln, 1) = "'" Or Left(ln, 4) = "rem ")
End Function

' scans the libraries of the book, a module at a time (the whole text at once: the words are looked for by the string search, not line by line):
' nMod modules, nLines lines, the net findings (module: what) in finds, the places of the system by kind in nSys()
Private Sub ScanLibraries(ByRef nMod As Long, ByRef nLines As Long, ByRef finds As String, nSys() As Long)
    Dim libs As Variant, oLib As Object, i As Long, j As Long, k As Long, src As String, low As String, p As Long
    Dim nw As Variant, sw As Variant, m As Long, parts As Variant, names As Variant, bc As Object, w As String
    nw = NetWords()
    sw = SysWords()
    bc = gDoc.BasicLibraries
    libs = Array("Standard", "WmsPlus")
    For i = 0 To UBound(libs)
        If bc.hasByName(libs(i)) Then
            If Not bc.isLibraryLoaded(libs(i)) Then bc.loadLibrary(libs(i))
            oLib = bc.getByName(libs(i))
            names = oLib.getElementNames()
            For j = 0 To UBound(names)
                src = oLib.getByName(names(j))
                low = LCase(src)
                nMod = nMod + 1
                nLines = nLines + Len(src) - Len(Replace(src, Chr(10), "")) + 1
                For m = 0 To UBound(nw)
                    parts = Split(nw(m), "|")
                    p = InStr(low, LCase(parts(0)))
                    If p > 0 Then finds = finds & "    " & names(j) & ": " & parts(1) & " (строка " & LineNo(low, p) & ")" & Chr(10)
                Next m
                For m = 0 To 3
                    w = LCase(sw(m))
                    p = InStr(low, w)
                    Do While p > 0
                        If Not InComment(low, p) Then nSys(m) = nSys(m) + 1
                        p = InStr(p + Len(w), low, w)
                    Loop
                Next m
            Next j
        End If
    Next i
End Sub

' what Calc keeps in the book: links to other files and areas; formulas that read the internet or open a link (one search of the formulas on every sheet).
' The hyperlinks that stand in the text of a cell are not looked for here (Calc gives no way to ask for all of them but cell by cell): tools/offline_audit.py reads them from the file
Private Sub ScanBook(ByRef nLink As Long, ByRef nForm As Long, ByRef what As String, ByRef failed As String)
    Dim sd As Object, i As Long, sh As Object, r As Variant, sheets As Object, hit As String
    nLink = LinkCount("SheetLinks", failed) + LinkCount("AreaLinks", failed) + LinkCount("ExternalDocLinks", failed)
    If nLink > 0 Then what = what & "связи с другими файлами: " & nLink & "; "
    sheets = gDoc.Sheets
    For i = 0 To sheets.getCount() - 1
        sh = sheets.getByIndex(i)
        hit = FormulaHit(sh, failed)
        If hit <> "" Then
            nForm = nForm + 1
            what = what & "лист «" & sh.Name & "»: " & hit & "; "
        End If
    Next i
End Sub

' the number of the links of a kind (0 and a note in `failed` when Calc cannot be asked)
Private Function LinkCount(sKind As String, ByRef failed As String) As Long
    On Error GoTo EH
    Select Case sKind
    Case "SheetLinks": LinkCount = gDoc.SheetLinks.getCount()
    Case "AreaLinks": LinkCount = gDoc.AreaLinks.getCount()
    Case Else: LinkCount = gDoc.ExternalDocLinks.getCount()
    End Select
    Exit Function
EH:
    failed = failed & sKind & " "
End Function

' the first formula of the sheet that reads the internet / opens a link: its address and the name of the function; "" — none
Private Function FormulaHit(sh As Object, ByRef failed As String) As String
    Dim sd As Object, r As Variant
    On Error GoTo EH
    sd = sh.createSearchDescriptor()
    sd.SearchRegularExpression = True
    sd.SearchCaseSensitive = False
    sd.SearchType = 0
    sd.SearchString = "(WEB" & "SERVICE|FILTER" & "XML|HYPER" & "LINK|D" & "DE)\("
    r = sh.findFirst(sd)
    If Not IsNull(r) And Not IsEmpty(r) Then FormulaHit = r.AbsoluteName
    Exit Function
EH:
    failed = failed & sh.Name & " "
End Function

' what LibreOffice itself is set to do, in words ("" — not known on this computer); information, not a finding: it is not WMS
Private Function LoOwn() As String
    Dim v As Variant, s As String
    v = WmsStatus.ConfigGet("/org.openoffice.Office.Update/Update", "Enabled")
    If Not IsEmpty(v) Then s = s & "проверка обновлений LibreOffice: " & IIf(v, "включена (отключить: Сервис → Параметры → LibreOffice → Онлайн-обновление)", "выключена") & Chr(10)
    v = WmsStatus.ConfigGet("/org.openoffice.Office.Common/Misc", "CrashReport")
    If Not IsEmpty(v) Then s = s & "отчёт об аварии LibreOffice: " & IIf(v, "после аварийного закрытия LibreOffice сама предложит отправить отчёт — это её окно, не WMS; не отправляйте", "выключен") & Chr(10)
    LoOwn = s
End Function

Function LocalityReport() As String
    Dim nMod As Long, nLines As Long, finds As String, nSys(3) As Long, nLink As Long, nForm As Long, what As String, nBad As Integer, failed As String
    Dim s As String, own As String, extra As String, i As Long, nm As Variant, t0 As Long, t1 As Long, t2 As Long
    WmsInit()
    On Error GoTo EH
    WmsView.ProgressStart("Проверка локальности", 3)
    t0 = GetSystemTicks()
    ScanLibraries(nMod, nLines, finds, nSys())
    t1 = GetSystemTicks()
    WmsView.ProgressSet(1)
    ScanBook(nLink, nForm, what, failed)
    t2 = GetSystemTicks()
    WmsView.ProgressSet(2)
    s = LcLine(nMod > 0 And finds = "", "макросы WMS (" & nMod & " модулей, " & nLines & " строк)", _
        IIf(nMod = 0, "тексты макросов не прочитаны", IIf(finds = "", "сетевых вызовов, почты и адресов нет", "найдено:" & Chr(10) & finds)), nBad)
    s = s & LcLine(nLink = 0 And nForm = 0 And failed = "", "книга", IIf(failed <> "", "не удалось проверить: " & failed, IIf(what = "", "связей с другими файлами и формул, читающих интернет, нет", what)), nBad)
    ' the places of the system: the known ones are named, a new one is shown
    nm = Array(EXPECT_SHELL, EXPECT_EXEC, EXPECT_WSH, EXPECT_FSO)
    For i = 0 To 3
        If nSys(i) <> nm(i) Then extra = extra & "    " & SysName(i) & ": " & nSys(i) & " (по замыслу " & nm(i) & ")" & Chr(10)
    Next i
    s = s & LcLine(extra = "", "запуск программ этого компьютера", IIf(extra = "", "только по замыслу: версия Python и свободное место (ответ — в файл в папке WMS), ярлык на рабочий стол, открытие фото позиции и папки документов вашей программой", _
        "число мест изменилось:" & Chr(10) & extra & "    покажите разработчику"), nBad)
    own = LoOwn()
    WmsView.ProgressEnd()
    LocalityReport = "Проверка локальности: " & IIf(nBad = 0, "сетевых возможностей в WMS не найдено", "замечаний " & nBad) & Chr(10) & s _
        & "Данные пишутся только в папку WMS (книга, журнал, копии, состояние) и на флешку, которую выбираете вы." & Chr(10) _
        & IIf(own <> "", Chr(10) & "Сама LibreOffice (не WMS):" & Chr(10) & own, "") _
        & Chr(10) & "Книга только прочитана, ничего не менялось (макросы " & (t1 - t0) & " мс, книга " & (t2 - t1) & " мс). Проверка со стороны: tools/offline_audit.py из папки выпуска."
    gUiLastMsg = LocalityReport
    Exit Function
EH:
    LocalityReport = "ERR:проверка не выполнена: " & Error$ & " (строка " & Erl & ", ошибка " & Err & ")"
    gUiLastMsg = LocalityReport
    WmsView.ProgressEnd()
End Function

Private Function SysName(i As Long) As String
    Select Case i
    Case 0: SysName = "команда системы"
    Case 1: SysName = "открытие файла программой компьютера"
    Case 2: SysName = "ярлык рабочего стола (Windows)"
    Case Else: SysName = "свободное место на диске (Windows)"
    End Select
End Function
