' WmsCare — 1.0 (этап K): скорость и обслуживание книги. Время открытия, сохранения и карты и сторож роста (С4); рост книги
' по месяцам и прогноз (С3); «Лишние ячейки», «Удалить старые снимки» (С3); советы по настройкам LibreOffice (С6); замер
' скорости этого ПК — «держит до ~N движений» (С2); сообщение при сохранении (С1). Всё хранится в WMS_State/wms_state.txt
' (WmsSafe.StGet / StSet) — книгу не меняет. Library WmsPlus.
Option Explicit

' С4: the thresholds of the watchdog of the growth (ms)
Private Const CARE_MAP_SLOW = 5000
Private Const CARE_SAVE_SLOW = 30000
Private Const CARE_OPEN_SLOW = 20000
' С3: the size the book is planned for and the one at which «Закрытие года» is due
Private Const CARE_PLANNED = 250000
Private Const CARE_YEAR_CLOSE = 200000
' С2: the reference of the speed test — the computer WMS 1.0 was measured on with 250 000 movements (ms of each part)
Private Const CARE_REF_BASIC = 260
Private Const CARE_REF_CALC = 90
Private Const CARE_REF_DISK = 60
Global gCrSaveT0 As Long

' ================================================================ С4 the times

' a time (ms) of kind sKind (OPEN, SAVE, MAP) remembered: T_<kind>=ms|date
Sub TimingSet(sKind As String, ms As Long)
    On Error Resume Next
    WmsSafe.StSet("T_" & sKind, ms & "|" & Format(Now(), "YYYY-MM-DD HH:MM"))
End Sub

' the last time (ms) of kind sKind; -1 — not known
Function TimingGet(sKind As String) As Long
    Dim s As String
    TimingGet = -1
    On Error GoTo EH
    s = WmsSafe.StGet("T_" & sKind)
    If s <> "" Then TimingGet = CLng(Val(Split(s, "|")(0)))
EH:
End Function

' С4: the map takes more than CARE_MAP_SLOW — it is built again only by «Обновить»
Function MapSlow() As Boolean
    MapSlow = (TimingGet("MAP") > CARE_MAP_SLOW)
End Function

' С1, С4: «Сохранение» begins — the estimate in the status bar, the time measured
Sub SaveStart()
    Dim t As Long
    On Error Resume Next
    gCrSaveT0 = GetSystemTicks()
    t = TimingGet("SAVE")
    If t >= 3000 Then WmsView.ProgressStart("Сохранение книги ~" & Format(t / 1000, "0") & " с — не закрывайте окно", 1)
End Sub

Sub SaveDone()
    On Error Resume Next
    WmsView.ProgressEnd()
    If gCrSaveT0 > 0 Then TimingSet("SAVE", GetSystemTicks() - gCrSaveT0)
    gCrSaveT0 = 0
End Sub

' С4: the advice when the book grew ("" — nothing to say)
Function GrowthNote() As String
    Dim s As String, t As Long
    t = TimingGet("SAVE")
    If t > CARE_SAVE_SLOW Then s = "сохранение книги идёт " & Format(t / 1000, "0") & " с"
    t = TimingGet("OPEN")
    If t > CARE_OPEN_SLOW Then s = s & IIf(s <> "", ", ", "") & "запуск WMS — " & Format(t / 1000, "0") & " с"
    t = TimingGet("MAP")
    If t > CARE_MAP_SLOW Then s = s & IIf(s <> "", ", ", "") & "карта строится " & Format(t / 1000, "0") & " с (теперь обновляется только кнопкой «Обновить»)"
    If s <> "" Then GrowthNote = "Книга выросла: " & s & ". Советы — «" & WmsSafe.SH_SERVICE & "» → «Советы LibreOffice»; к " _
        & WmsView.NumText(CARE_YEAR_CLOSE) & " движений — «Закрытие года»"
End Function

' ================================================================ С3 the growth of the book

' the rows of the sheets of the movements (the size of the book)
Function Movements() As Long
    Dim t As Variant, i As Integer, n As Long
    t = Array(SH_ORDERS, SH_SPECIAL, SH_ISSUES, SH_RETURNS, SH_ADJUST)
    For i = 0 To UBound(t)
        If gDoc.Sheets.hasByName(t(i)) Then n = n + WmsOrders.LastRow(gDoc.Sheets.getByName(t(i)))
    Next i
    Movements = n
End Function

' the first start of a month: the size of the book and its movements remembered (GROW_<ГГГГ-ММ>=МБ|движений)
Sub GrowthMark()
    Dim k As String
    On Error Resume Next
    k = "GROW_" & Format(Now(), "YYYY-MM")
    If WmsSafe.StGet(k) = "" Then WmsSafe.StSet(k, Format(WmsSafe.BookMb(), "0.00") & "|" & Movements())
End Sub

' «+3,2 МБ и +4 100 движений в месяц; до 200 000 движений — ~14 мес.» ("" — less than a month of history)
Function GrowthText() As String
    Dim a As Variant, i As Long, k As String, first As String, months As Long, y As Integer, m As Integer, v As Variant, mb0 As Double, mv0 As Double
    Dim mb As Double, mv As Double, dMb As Double, dMv As Double, s As String
    On Error GoTo EH
    ' the oldest month remembered within the last 24
    For i = 24 To 1 Step -1
        y = Year(Now())
        m = Month(Now()) - i
        Do While m < 1
            m = m + 12
            y = y - 1
        Loop
        k = "GROW_" & y & "-" & Right("0" & m, 2)
        If WmsSafe.StGet(k) <> "" Then
            first = WmsSafe.StGet(k)
            months = i
            Exit For
        End If
    Next i
    mb = WmsSafe.BookMb()
    mv = Movements()
    If first = "" Then
        GrowthText = Format(mb, "0.0") & " МБ, движений " & WmsView.NumText(mv) & " (рост посчитается через месяц)"
        Exit Function
    End If
    v = Split(first, "|")
    mb0 = Val(v(0))
    mv0 = Val(v(1))
    dMb = (mb - mb0) / months
    dMv = (mv - mv0) / months
    s = Format(mb, "0.0") & " МБ, движений " & WmsView.NumText(mv) & "; в месяц +" & Format(dMb, "0.0") & " МБ и +" & WmsView.NumText(CLng(dMv)) & " движений"
    If dMv > 0 And mv < CARE_YEAR_CLOSE Then s = s & "; до " & WmsView.NumText(CARE_YEAR_CLOSE) & " движений («Закрытие года») — ~" & CLng((CARE_YEAR_CLOSE - mv) / dMv) & " мес."
    If mv >= CARE_YEAR_CLOSE Then s = "✗ " & s & "; пора «Закрытие года»"
    GrowthText = s
    Exit Function
EH:
    GrowthText = ""
End Function

' «запуск 1,2 с · сохранение 6 с · карта 2,1 с»
Function TimingText() As String
    Dim s As String, k As Variant, nm As Variant, i As Integer, t As Long
    k = Array("OPEN", "SAVE", "MAP")
    nm = Array("запуск WMS", "сохранение", "карта")
    For i = 0 To 2
        t = TimingGet(CStr(k(i)))
        If t >= 0 Then s = s & IIf(s <> "", " · ", "") & nm(i) & " " & Format(t / 1000, "0.0") & " с"
    Next i
    If s = "" Then s = "ещё не замерено"
    If GrowthNote() <> "" Then s = "✗ " & s Else s = "✓ " & s
    TimingText = s
End Function

' ================================================================ С3 the extra cells, the old snapshots

Sub BtnExtraCells(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage(ExtraCells())
End Sub

' the sheets whose used area goes far below their data (empty texts, formulas of nothing, spaces left there make the file
' bigger and slower): «Лист: занято до строки N, данные до M». Nothing is changed
Function ExtraCells() As String
    Dim i As Integer, sh As Object, cur As Object, endRow As Long, endCol As Long, dataRow As Long, s As String, n As Integer, sErrTxt As String
    On Error GoTo EH
    WmsView.ProgressStart("Лишние ячейки: просмотр листов", gDoc.Sheets.getCount())
    For i = 0 To gDoc.Sheets.getCount() - 1
        WmsView.ProgressSet(i)
        sh = gDoc.Sheets.getByIndex(i)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        endRow = cur.getRangeAddress().EndRow
        endCol = cur.getRangeAddress().EndColumn
        If endRow > 2000 Then
            dataRow = LastRealRow(sh, endRow, endCol)
            If endRow - dataRow > 2000 Then
                s = s & Chr(10) & "• «" & sh.Name & "»: занято до строки " & WmsView.NumText(endRow + 1) & ", данные — до " & WmsView.NumText(dataRow + 1)
                n = n + 1
            End If
        End If
    Next i
    WmsView.ProgressEnd()
    If n = 0 Then
        ExtraCells = "Лишних ячеек нет: у каждого листа занятая область кончается там, где кончаются данные."
    Else
        ExtraCells = "Листов с лишними ячейками внизу: " & n & s & Chr(10) & Chr(10) & "Ниже данных остались пустой текст, пробелы или формулы без " _
            & "значения — файл от них больше и медленнее. Удалите эти строки (выделить от строки после данных до конца, «Удалить строки») — на копии " _
            & "книги, потом «Самопроверка»; или сообщите ответственному за WMS."
    End If
    gUiLastMsg = ExtraCells
    Exit Function
EH:
    sErrTxt = Error$
    WmsView.ProgressEnd()
    ExtraCells = "ERR:листы не просмотрены: " & sErrTxt
    gUiLastMsg = ExtraCells
End Function

' the last row with a value or a visible text, looked for from the bottom up in portions
Private Function LastRealRow(sh As Object, endRow As Long, endCol As Long) As Long
    Dim r1 As Long, r0 As Long, d As Variant, i As Long, j As Long, v As Variant
    r1 = endRow
    Do While r1 >= 0
        r0 = r1 - 1999
        If r0 < 0 Then r0 = 0
        d = sh.getCellRangeByPosition(0, r0, endCol, r1).getDataArray()
        For i = UBound(d) To 0 Step -1
            For j = 0 To UBound(d(i))
                v = d(i)(j)
                If VarType(v) = 5 Then
                    LastRealRow = r0 + i
                    Exit Function
                ElseIf Trim(CStr(v)) <> "" Then
                    LastRealRow = r0 + i
                    Exit Function
                End If
            Next j
        Next i
        r1 = r0 - 1
    Loop
End Function

Sub BtnPruneSnapshots(Optional oEvent As Variant)
    Dim k As Integer
    WmsInit()
    k = WmsPrimeUi.UiButtons("Старые снимки", "Удалить снимки «Экспорт для инструментов» кроме двух последних? Снимок делается заново из книги в любой момент; " _
        & "данные книги и журнал не затрагиваются.", Array("Удалить старые", "Отмена"), Array(1, 0))
    If k <> 1 Then Exit Sub
    WmsUi.UiMessage(CarePrune(2))
End Sub

' the snapshots of WMS_Export but the nKeep newest removed: «удалено N, освобождено X МБ»
Function CarePrune(nKeep As Integer) As String
    Dim d As String, f As Variant, names() As String, n As Long, i As Long, j As Long, t As String, nm As String, freed As Double, nDel As Long
    On Error GoTo EH
    d = gBaseDir & EXPORT_DIR & "/"
    If Not gSfa.exists(d) Then
        CarePrune = "Снимков нет."
        Exit Function
    End If
    f = gSfa.getFolderContents(d, True)
    n = -1
    For i = 0 To UBound(f)
        nm = FileNameOf(f(i))
        If Right(nm, 1) = "/" Then nm = Left(nm, Len(nm) - 1)
        If gSfa.isFolder(f(i)) And (Left(nm, 13) = "WMS_SNAPSHOT_" Or Left(nm, 10) = "WMS_STOCK_") Then
            n = n + 1
            ReDim Preserve names(n)
            names(n) = Mid(nm, InStr(nm, "_2") + 1) & "|" & nm
        End If
    Next i
    For i = 1 To n
        t = names(i)
        j = i - 1
        Do While j >= 0
            If names(j) <= t Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = t
    Next i
    For i = 0 To n - nKeep
        nm = Split(names(i), "|")(1)
        freed = freed + FolderBytes(d & nm & "/")
        gSfa.kill(d & nm & "/")
        nDel = nDel + 1
    Next i
    CarePrune = "Старых снимков удалено: " & nDel & ", освобождено " & Format(freed / 1048576, "0.0") & " МБ (оставлены последние " & nKeep & ")."
    gUiLastMsg = CarePrune
    Exit Function
EH:
    CarePrune = "ERR:снимки не удалены: " & Error$
    gUiLastMsg = CarePrune
End Function

Private Function FolderBytes(d As String) As Double
    Dim f As Variant, i As Long, n As Double
    On Error Resume Next
    f = gSfa.getFolderContents(d, True)
    For i = 0 To UBound(f)
        If gSfa.isFolder(f(i)) Then n = n + FolderBytes(f(i) & IIf(Right(f(i), 1) = "/", "", "/")) Else n = n + gSfa.getSize(f(i))
    Next i
    FolderBytes = n
End Function

' ================================================================ С6 the settings of LibreOffice

Private Function Cfg(sPath As String, sName As String) As Variant
    Dim cp As Object, ca As Object, args(0) As New com.sun.star.beans.PropertyValue
    On Error GoTo EH
    cp = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
    args(0).Name = "nodepath"
    args(0).Value = sPath
    ca = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", args())
    Cfg = ca.getByName(sName)
    Exit Function
EH:
    Cfg = Empty
End Function

Private Function CfgSet(sPath As String, sName As String, v As Variant) As Boolean
    Dim cp As Object, ca As Object, args(0) As New com.sun.star.beans.PropertyValue
    On Error GoTo EH
    cp = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
    args(0).Name = "nodepath"
    args(0).Value = sPath
    ca = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", args())
    ca.replaceByName(sName, v)
    ca.commitChanges()
    CfgSet = True
EH:
End Function

' the settings that matter for a big book: «текущее → советуем»; nBad — how many differ
Function LoTips(ByRef nBad As Integer) As String
    Dim s As String, v As Variant, w As Variant
    nBad = 0
    v = Cfg("/org.openoffice.Office.Common/Save/Document", "CreateBackup")
    s = s & Tip(Not IsEmpty(v) And v = True, "«Всегда создавать резервную копию» (Сервис → Параметры → Загрузка/сохранение)", IIf(v = True, "включено", "выключено"), _
        "включить: при сбое сохранения останется предыдущая версия файла (.bak)", nBad)
    v = Cfg("/org.openoffice.Office.Recovery/AutoSave", "Enabled")
    w = Cfg("/org.openoffice.Office.Recovery/AutoSave", "TimeIntervall")
    s = s & Tip(Not IsEmpty(v) And v = True And Not IsEmpty(w) And w >= 5 And w <= 15, "данные для восстановления", IIf(v = True, "каждые " & w & " мин.", "выключено"), _
        "каждые 10 мин.: после сбоя LibreOffice предложит восстановить окно (данные WMS и так в журнале)", nBad)
    v = Cfg("/org.openoffice.Office.Recovery/AutoSave", "UserAutoSave")
    s = s & Tip(IsEmpty(v) Or v = False, "автосохранение самого документа", IIf(v = True, "включено", "выключено"), _
        "выключить: большую книгу сохраняйте кнопкой, когда удобно (проведённое — в журнале)", nBad)
    v = Cfg("/org.openoffice.Office.Calc/Formula/Load", "ODFRecalcMode")
    s = s & Tip(Not IsEmpty(v) And v = 1, "пересчёт формул при загрузке ODF", IIf(IsEmpty(v), "?", Choose(v + 1, "всегда", "никогда", "спрашивать")), _
        "«никогда» — большая книга открывается быстрее (результаты формул хранятся в файле)", nBad)
    v = Cfg("/org.openoffice.Office.Common/Undo", "Steps")
    s = s & Tip(Not IsEmpty(v) And v <= 100, "шагов отмены", IIf(IsEmpty(v), "?", CStr(v)), "не больше 100 — меньше памяти на большой книге", nBad)
    v = Cfg("/org.openoffice.Office.Calc/Layout/Line", "GridOnColoredCells")
    s = s & Tip(Not IsEmpty(v) And v = True, "линии сетки на цветных ячейках (Сервис → Параметры → LibreOffice Calc → Вид)", _
        IIf(v = True, "показываются", "не показываются"), "показывать: серые колонки WMS не сливаются в полосу", nBad)
    LoTips = s
End Function

Private Function Tip(ok As Boolean, sWhat As String, sNow As String, sAdvice As String, ByRef nBad As Integer) As String
    If Not ok Then nBad = nBad + 1
    Tip = IIf(ok, "✓ ", "✗ ") & sWhat & ": " & sNow & IIf(ok, "", " — советуем " & sAdvice) & Chr(10)
End Function

Sub BtnLoTips(Optional oEvent As Variant)
    Dim s As String, n As Integer, k As Integer
    WmsInit()
    s = LoTips(n)
    If n = 0 Then
        WmsUi.UiMessage("Настройки LibreOffice для большой книги — в порядке:" & Chr(10) & s)
        Exit Sub
    End If
    k = WmsPrimeUi.UiButtons("Советы по настройкам LibreOffice", "Настройки LibreOffice для большой книги:" & Chr(10) & s & Chr(10) _
        & "Применить советы? Они меняют только настройки LibreOffice этого компьютера, не книгу.", Array("Применить советы", "Не менять"), Array(1, 0))
    If k = 1 Then WmsUi.UiMessage(LoApply())
End Sub

' the advice applied to the settings of LibreOffice of this computer
Function LoApply() As String
    Dim n As Integer, s As String
    CfgSet("/org.openoffice.Office.Common/Save/Document", "CreateBackup", True)
    CfgSet("/org.openoffice.Office.Recovery/AutoSave", "Enabled", True)
    CfgSet("/org.openoffice.Office.Recovery/AutoSave", "TimeIntervall", 10)
    CfgSet("/org.openoffice.Office.Recovery/AutoSave", "UserAutoSave", False)
    CfgSet("/org.openoffice.Office.Calc/Formula/Load", "ODFRecalcMode", 1)
    If Cfg("/org.openoffice.Office.Common/Undo", "Steps") > 100 Then CfgSet("/org.openoffice.Office.Common/Undo", "Steps", 100)
    CfgSet("/org.openoffice.Office.Calc/Layout/Line", "GridOnColoredCells", True)
    s = LoTips(n)
    LoApply = IIf(n = 0, "Советы применены:", "Применено не всё (" & n & "):") & Chr(10) & s
    gUiLastMsg = LoApply
End Function

' 1.0.1: the grid of the cells is drawn over the coloured ones too (the grey columns WMS writes, the colours of the statuses) —
' at every start, once: without it the cells of those columns run together. A setting of LibreOffice of this computer, not of
' the book
Sub GridOnColored()
    On Error Resume Next
    If Cfg("/org.openoffice.Office.Calc/Layout/Line", "GridOnColoredCells") <> True Then
        CfgSet("/org.openoffice.Office.Calc/Layout/Line", "GridOnColoredCells", True)
    End If
End Sub

' ================================================================ С2 the speed of this computer

Sub BtnSpeedTest(Optional oEvent As Variant)
    WmsInit()
    WmsUi.UiMessage(SpeedTest())
End Sub

' three short measures (the Basic, the formula engine of Calc, the disk) against the computer WMS was measured on with 250 000
' movements: «этот ПК комфортно держит до ~N движений». A few seconds; nothing in the book changes
Function SpeedTest() As String
    Dim t0 As Long, tB As Long, tC As Long, tD As Long, i As Long, x As Double, s As String, f As String, xs As Object, tos As Object, r As Double
    Dim n As Long, chunk As String, mv As Long, sErrTxt As String
    On Error GoTo EH
    WmsView.ProgressStart("Замер скорости ПК", 3)
    t0 = GetSystemTicks()
    For i = 1 To 300000
        x = x + (i Mod 7) * 1.5
        If i Mod 1000 = 0 Then s = Left(s & "ab", 20)
    Next i
    tB = GetSystemTicks() - t0
    WmsView.ProgressSet(1)
    t0 = GetSystemTicks()
    s = WmsView.Eng("SUMPRODUCT((MOD(ROW($A$1:$A$400000);7)=3)*1)")
    tC = GetSystemTicks() - t0
    WmsView.ProgressSet(2)
    f = WmsSafe.SafeDir() & "speed_" & NewId(4) & ".tmp"
    If Not gSfa.exists(WmsSafe.SafeDir()) Then gSfa.createFolder(WmsSafe.SafeDir())
    chunk = String(1024, "Ж")
    t0 = GetSystemTicks()
    xs = gSfa.openFileWrite(f)
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(xs.getOutputStream())
    tos.setEncoding("UTF-8")
    For i = 1 To 2048
        tos.writeString(chunk)
    Next i
    tos.closeOutput()
    s = WmsAudit.ReadText(f)
    gSfa.kill(f)
    tD = GetSystemTicks() - t0
    WmsView.ProgressEnd()
    ' the speed relative to the reference: the geometric mean of the three ratios (> 1 — faster)
    r = Exp((Log(CARE_REF_BASIC / IIf(tB > 0, tB, 1)) + Log(CARE_REF_CALC / IIf(tC > 0, tC, 1)) + Log(CARE_REF_DISK / IIf(tD > 0, tD, 1))) / 3)
    n = CLng(CARE_PLANNED * r / 10000) * 10000
    If n < 30000 Then n = 30000
    If n > 600000 Then n = 600000
    mv = Movements()
    SpeedTest = "Замер скорости этого ПК: вычисления " & tB & " мс, формулы Calc " & tC & " мс, диск " & tD & " мс (эталон " & CARE_REF_BASIC & " / " _
        & CARE_REF_CALC & " / " & CARE_REF_DISK & " мс)." & Chr(10) & "Этот ПК " & IIf(r >= 1, "быстрее эталона в " & Format(r, "0.0"), "медленнее эталона в " _
        & Format(1 / r, "0.0")) & " раза: комфортно держит книгу до ~" & WmsView.NumText(n) & " движений. Сейчас в книге " & WmsView.NumText(mv) & "." _
        & IIf(mv > n * 0.8, Chr(10) & "Книга близка к пределу этого ПК — запланируйте «Закрытие года» или более быстрый компьютер.", "")
    WmsSafe.StSet("SPEED", Format(Now(), "YYYY-MM-DD") & "|" & n & "|" & tB & "|" & tC & "|" & tD)
    gUiLastMsg = SpeedTest
    Exit Function
EH:
    sErrTxt = Error$
    WmsView.ProgressEnd()
    SpeedTest = "ERR:замер не выполнен: " & sErrTxt
    gUiLastMsg = SpeedTest
End Function

' ================================================================ С8 the light snapshot for the labels

Sub BtnLiteSnapshot(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = WmsExport.ExportSnapshot(True)
    If Left(s, 3) = "OK:" Then
        WmsOrdersUi.UiInfo("Лёгкий снимок для этикеток готов: " & Mid(s, 4))
    Else
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    End If
End Sub
