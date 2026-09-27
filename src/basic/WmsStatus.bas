' WmsStatus — «Состояние системы» и «Проверка перед работой» (задание «FINAL WMS MARATHON», §5: WMS_PROD_CANDIDATE).
' Только чтение: ни одна функция модуля не меняет учёт и не пишет в журнал; результат выводится в раздел «Главной»
' (строки ниже справки) и в окно. Проверка совместимости LibreOffice: версии, на которых WMS проверена, — без замечаний;
' другая версия — предупреждение (работа не блокируется: учёт защищён журналом и проверками ядра), слишком старая —
' ошибка проверки перед работой.
Option Explicit

' LibreOffice versions (major.minor) WMS was verified on: the development environment and the cross-version round trip
Public Const TESTED_LO = "24.2|26.2"
' the oldest major version WMS is expected to work on (the functions and the sheet protection options it uses)
Private Const MIN_LO_MAJOR = 7
' «Главная»: the section of this module (0-based row of its title; column A — check, B — result)
Public Const STATUS_ROW = 24
Private Const STATUS_ROWS = 30
' the book size WMS is planned for (spec §27): about 250k movements
Private Const PLANNED_MOVEMENTS = 250000

' test seam (TestSetLoVersion): a Global keeps its value between macro calls, a module-level Private is reset
Global gFakeLo As String

' ================================================================ LibreOffice and its settings

Private Function Config(sPath As String, sName As String) As Variant
    Dim cp As Object, ca As Object, args(0) As New com.sun.star.beans.PropertyValue
    On Error GoTo EH
    cp = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
    args(0).Name = "nodepath"
    args(0).Value = sPath
    ca = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", args())
    Config = ca.getByName(sName)
    Exit Function
EH:
    Config = Empty
End Function

Function LoVersion() As String
    If gFakeLo <> "" Then
        LoVersion = gFakeLo
    Else
        LoVersion = CStr(Config("/org.openoffice.Setup/Product", "ooSetupVersionAboutBox"))
    End If
End Function

' 0 — a tested version, 1 — not tested (a warning), 2 — too old or unknown (an error of the check before work)
Function LoCompat(ByRef why As String) As Integer
    Dim v As String, p As Variant
    v = LoVersion()
    p = Split(v, ".")
    why = ""
    If UBound(p) < 1 Then
        why = "версия LibreOffice не определена («" & v & "»)"
        LoCompat = 2
    ElseIf InStr("|" & TESTED_LO & "|", "|" & p(0) & "." & p(1) & "|") > 0 Then
        LoCompat = 0
    ElseIf Val(p(0)) < MIN_LO_MAJOR Then
        why = "LibreOffice " & v & " слишком старая для WMS (нужна " & MIN_LO_MAJOR & ".x или новее; проверены " & Replace(TESTED_LO, "|", ", ") & ")"
        LoCompat = 2
    Else
        why = "LibreOffice " & v & " с WMS не проверялась (проверены " & Replace(TESTED_LO, "|", ", ") & ") — до работы с настоящими данными " _
            & "выполните проверку WMS на этом ПК (пакет проверки ПК)"
        LoCompat = 1
    End If
End Function

' macro security: 0 low … 3 very high; the folder of the book among the trusted locations
Private Function MacroSecurity(ByRef trusted As Boolean) As Integer
    Dim lvl As Variant, urls As Variant, i As Integer, d As String
    lvl = Config("/org.openoffice.Office.Common/Security/Scripting", "MacroSecurityLevel")
    MacroSecurity = IIf(IsEmpty(lvl), -1, lvl)
    urls = Config("/org.openoffice.Office.Common/Security/Scripting", "SecureURL")
    trusted = False
    d = LCase(gBaseDir)
    If IsArray(urls) Then
        For i = 0 To UBound(urls)
            If urls(i) <> "" Then
                If Left(d, Len(urls(i))) = LCase(urls(i)) Or Left(d, Len(urls(i)) + 1) = LCase(urls(i)) & "/" Then trusted = True
            End If
        Next i
    End If
End Function

' ================================================================ files

' "" when a probe file can be written and removed in the folder
Private Function Writable(sDir As String) As String
    Dim f As String, n As Integer
    On Error GoTo EH
    If Not gSfa.exists(sDir) Then gSfa.createFolder(sDir)
    f = sDir & ".wms_probe_" & NewId(6)
    n = FreeFile()
    Open f For Output As #n
    Print #n, "probe"
    Close #n
    gSfa.kill(f)
    Writable = ""
    Exit Function
EH:
    Writable = Error$
End Function

' the newest backup of this book: its name and age in days (-1 — none)
Private Function LastBackup(ByRef nm As String, ByRef n As Long) As Double
    Dim files As Variant, i As Long, f As String, sPrefix As String, t As Double, tBest As Double
    LastBackup = -1
    nm = ""
    n = 0
    If Not gSfa.exists(gBDir) Then Exit Function
    files = gSfa.getFolderContents(gBDir, False)
    sPrefix = WmsBackup.BookBaseName() & "_"
    For i = 0 To UBound(files)
        f = FileNameOf(files(i))
        If Left(f, Len(sPrefix)) = sPrefix And LCase(Right(f, 4)) = ".ods" Then
            n = n + 1
            t = DtToDouble(gSfa.getDateTimeModified(files(i)))
            If nm = "" Or t > tBest Then
                nm = f
                tBest = t
            End If
        End If
    Next i
    If nm <> "" Then LastBackup = CDbl(Now()) - tBest
End Function

Private Function DtToDouble(t As Variant) As Double
    DtToDouble = CDbl(DateSerial(t.Year, t.Month, t.Day)) + CDbl(TimeSerial(t.Hours, t.Minutes, t.Seconds))
End Function

Private Function SizeMb(url As String) As Double
    On Error GoTo EH
    SizeMb = Int(gSfa.getSize(url) / 10485.76) / 100
    Exit Function
EH:
    SizeMb = -1
End Function

' the rows of the movement sheets (the size of the book, spec §27)
Private Function Movements() As Long
    Dim t As Variant, i As Integer, n As Long
    t = Array(SH_ORDERS, SH_SPECIAL, SH_ISSUES, SH_RETURNS, SH_ADJUST)
    For i = 0 To UBound(t)
        If gDoc.Sheets.hasByName(t(i)) Then n = n + WmsOrders.LastRow(gDoc.Sheets.getByName(t(i)))
    Next i
    Movements = n
End Function

' ================================================================ the two reports

Private Function L(sLevel As String, sWhat As String, sText As String) As String
    L = sLevel & Chr(9) & sWhat & Chr(9) & sText & Chr(10)
End Function

' «Состояние системы»: versions, the working file, the journal, the backups, the size — lines «уровень TAB что TAB значение»
Function SystemStatus() As String
    Dim s As String, why As String, lvl As Integer, trusted As Boolean, nm As String, n As Long, age As Double, jf As String
    WmsInit()
    If gState = "" Then WmsStartup()
    s = L("INFO", "WMS", "продукт " & WMS_PRODUCT_VERSION & ", ядро " & WMS_CORE_VERSION & ", схема " & SysStr(SK_SCHEMA) & " (ядро ожидает " & WMS_SYS_SCHEMA & ")")
    s = s & L(IIf(gState = "CLEAN", "OK", "FAIL"), "состояние", IIf(gState = "CLEAN", "работа разрешена", gState & " [" & gBlock & "] " & gBlockText))
    s = s & L("INFO", "экземпляр", SysStr(SK_INSTANCE) & ", режим " & SysStr(SK_MODE))
    s = s & L("INFO", "рабочий файл", ConvertFromURL(gDoc.getURL()) & IIf(SysStr(SK_REG_URL) = gDoc.getURL(), " (зарегистрирован)", " (НЕ рабочий путь)"))
    lvl = LoCompat(why)
    s = s & L(Choose(lvl + 1, "OK", "WARN", "FAIL"), "LibreOffice", LoVersion() & IIf(why <> "", " — " & why, " — проверенная версия"))
    jf = Mid(SysStr(SK_JPOS), 1, InStr(SysStr(SK_JPOS) & "|", "|") - 1)
    s = s & L("INFO", "журнал", ConvertFromURL(gJDir) & "; последняя операция " & SysStr(SK_LAST_SEQ) & IIf(jf <> "", "; файл " & jf, ""))
    age = LastBackup(nm, n)
    s = s & L(IIf(age < 0, "WARN", "INFO"), "резервные копии", IIf(age < 0, "нет ни одной — «Резервная копия» на «Главной»", n & " шт.; последняя " & nm _
        & " (" & Format(age, "0.0") & " дн. назад)") & "; папка " & ConvertFromURL(gBDir))
    s = s & L("INFO", "размер книги", SizeMb(gDoc.getURL()) & " МБ; строк движений " & Movements() & " (ориентир — до " & PLANNED_MOVEMENTS & ")")
    s = s & L("INFO", "проверено", Format(Now(), "DD.MM.YYYY HH:MM:SS"))
    SystemStatus = s
End Function

' «Проверка перед работой»: the quick checks a storekeeper runs before a working day (a few seconds, read-only)
Function PreWorkCheck() As String
    Dim s As String, why As String, lvl As Integer, trusted As Boolean, nm As String, n As Long, age As Double, p As String, sec As Integer
    Dim auto As Variant, autoMin As Variant, userAuto As Variant, mv As Long, nFail As Integer, nWarn As Integer, k As Long, ln As Variant
    WmsInit()
    If gState = "" Then WmsStartup()
    s = L(IIf(gState = "CLEAN", "OK", "FAIL"), "WMS", IIf(gState = "CLEAN", "работа разрешена", "ЗАБЛОКИРОВАНО: " & gBlockText & " — см. «Что сделать»"))
    lvl = LoCompat(why)
    s = s & L(Choose(lvl + 1, "OK", "WARN", "FAIL"), "LibreOffice", LoVersion() & IIf(why <> "", " — " & why, ""))
    If Not IsOdsDocument() Then
        s = s & L("FAIL", "формат", "книга не в формате ODS — сохраняйте только в ODS")
    ElseIf gDoc.isReadonly() Then
        s = s & L("FAIL", "файл", "книга открыта только для чтения")
    Else
        s = s & L("OK", "файл", "ODS, запись разрешена")
    End If
    sec = MacroSecurity(trusted)
    If sec = 0 Then
        s = s & L("WARN", "макросы", "уровень безопасности «низкий» — поднимите до «высокого» и добавьте папку WMS в доверенные расположения")
    Else
        s = s & L("OK", "макросы", "уровень безопасности " & sec & IIf(trusted, "; папка WMS — доверенное расположение", "; папка WMS не в списке доверенных (макросы разрешены иначе)"))
    End If
    p = Writable(gJDir)
    s = s & L(IIf(p = "", "OK", "FAIL"), "журнал", IIf(p = "", "папка журнала доступна для записи", "запись в " & ConvertFromURL(gJDir) & " невозможна: " & p))
    p = Writable(gBDir)
    age = LastBackup(nm, n)
    If p <> "" Then
        s = s & L("FAIL", "резервные копии", "запись в " & ConvertFromURL(gBDir) & " невозможна: " & p)
    ElseIf age < 0 Then
        s = s & L("WARN", "резервные копии", "копий нет — нажмите «Резервная копия»")
    ElseIf age > 7 Then
        s = s & L("WARN", "резервные копии", "последняя копия " & Format(age, "0") & " дн. назад — нажмите «Резервная копия» и перенесите копии на другой носитель")
    Else
        s = s & L("OK", "резервные копии", n & " шт.; последняя " & Format(age, "0.0") & " дн. назад")
    End If
    auto = Config("/org.openoffice.Office.Recovery/AutoSave", "Enabled")
    autoMin = Config("/org.openoffice.Office.Recovery/AutoSave", "TimeIntervall")
    userAuto = Config("/org.openoffice.Office.Recovery/AutoSave", "UserAutoSave")
    If Not IsEmpty(userAuto) And userAuto = True Then
        s = s & L("WARN", "автосохранение", "включено автосохранение самого документа — WMS рекомендует сохранять кнопкой (данные защищены журналом)")
    Else
        s = s & L("OK", "автосохранение", IIf(Not IsEmpty(auto) And auto = True, "данные восстановления каждые " & autoMin & " мин.", "выключено") _
            & "; документ сохраняется только пользователем")
    End If
    mv = Movements()
    s = s & L(IIf(mv > PLANNED_MOVEMENTS * 0.8, "WARN", "OK"), "размер", mv & " строк движений" & IIf(mv > PLANNED_MOVEMENTS * 0.8, _
        " — близко к ориентиру " & PLANNED_MOVEMENTS & ": пора архивировать историю (WMS_ARCHIVE)", ""))
    p = Protections()
    s = s & L(IIf(p = "", "OK", "WARN"), "защиты", IIf(p = "", "листы защищены, служебные листы скрыты, структура книги защищена", _
        p & " — сообщите ответственному за WMS: защита не даёт изменить учёт мимо WMS"))
    k = WmsIssue.UnpostedCount()
    If k > 0 Then s = s & L("INFO", "непроведённые", "выдачи: " & k & " строк(и) без № — проверьте перед началом работы")
    If gDoc.isModified() Then s = s & L("INFO", "сохранение", "есть несохранённые изменения (проведённые операции уже в журнале) — сохраните книгу")
    ln = Split(s, Chr(10))
    For k = 0 To UBound(ln)
        If Left(ln(k), 4) = "FAIL" Then nFail = nFail + 1
        If Left(ln(k), 4) = "WARN" Then nWarn = nWarn + 1
    Next k
    PreWorkCheck = "ПРОВЕРКА ПЕРЕД РАБОТОЙ: ошибок " & nFail & ", предупреждений " & nWarn & Chr(10) & s
End Function

' "" when the protections of the book are in place: every sheet protected but «Получатели» (a list kept by hand), the
' service sheets of WMS hidden, the structure of the book protected; otherwise what is not (the same rule as WMS_DOCTOR)
Function Protections() As String
    Dim i As Integer, sh As Object, p As String, svc As String
    svc = "|" & SH_ORD & "|" & SH_RCV & "|" & SH_SPR & "|" & SH_ART & "|" & SH_RET & "|" & SH_ISS & "|" & SH_ADJ & "|" & SH_IDX & "|" & SYS_SHEET & "|"
    For i = 0 To gDoc.Sheets.getCount() - 1
        sh = gDoc.Sheets.getByIndex(i)
        If Not sh.isProtected() And sh.Name <> SH_RCPT Then p = p & IIf(p <> "", "; ", "") & "лист «" & sh.Name & "» не защищён"
        If InStr(svc, "|" & sh.Name & "|") > 0 And sh.IsVisible Then p = p & IIf(p <> "", "; ", "") & "служебный лист «" & sh.Name & "» не скрыт"
    Next i
    If Not gDoc.isProtected() Then p = p & IIf(p <> "", "; ", "") & "структура книги не защищена"
    Protections = p
End Function

' ================================================================ «Главная»

' the lines «уровень TAB что TAB значение» into the section of «Главная» (A — what, B — level and value)
Sub ShowSection(sTitle As String, sText As String)
    Dim sh As Object, ln As Variant, i As Integer, f As Variant, r As Long, wasModified As Boolean, c As Object
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_MAIN) Then Exit Sub
    wasModified = gDoc.isModified()
    sh = gDoc.Sheets.getByName(SH_MAIN)
    For r = STATUS_ROW To STATUS_ROW + STATUS_ROWS
        For i = 0 To 1
            c = sh.getCellByPosition(i, r)
            If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(c)
        Next i
    Next r
    sh.getCellByPosition(0, STATUS_ROW).setString(sTitle)
    ln = Split(sText, Chr(10))
    r = STATUS_ROW + 1
    For i = 0 To UBound(ln)
        f = Split(ln(i), Chr(9))
        If UBound(f) >= 2 And r <= STATUS_ROW + STATUS_ROWS Then
            sh.getCellByPosition(0, r).setString(f(1))
            c = sh.getCellByPosition(1, r)
            c.setString(Choose(InStr("OK  INFOWARNFAIL", Left(f(0) & "    ", 4)) \ 4 + 1, "✓ ", "· ", "! ", "✗ ") & f(2))
            r = r + 1
        ElseIf i = 0 And UBound(f) < 2 And ln(i) <> "" Then
            sh.getCellByPosition(1, STATUS_ROW).setString(ln(i))
        End If
    Next i
    If Not wasModified Then gDoc.setModified(False)
    Exit Sub
EH:
    ' the section is informational: a failure to show it never stops WMS
End Sub

Sub BtnSystemStatus(Optional oEvent As Variant)
    WmsInit()
    ShowSection("СОСТОЯНИЕ СИСТЕМЫ", SystemStatus())
End Sub

Sub BtnPreWorkCheck(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PreWorkCheck()
    ShowSection("ПРОВЕРКА ПЕРЕД РАБОТОЙ", s)
    WmsUi.UiMessage(Left(s, InStr(s & Chr(10), Chr(10)) - 1) & " — подробности на «Главной», раздел «ПРОВЕРКА ПЕРЕД РАБОТОЙ»")
End Sub

' ================================================================ test seam (inert unless _SYS MODE = TEST)

Function TestSetLoVersion(v As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSetLoVersion = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gFakeLo = v
    TestSetLoVersion = "OK"
End Function
