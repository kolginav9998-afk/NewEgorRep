' TbCommon — общий модуль инструментов WMS_TOOLBOX (задание «FINAL WMS MARATHON», §6; контракт docs/EXPORT_CONTRACT.md).
' Инструменты никогда не открывают WMS_PROD: они читают снимок «Экспорт для инструментов» (WMS-SNAPSHOT-1) и пишут
' пакеты (WMS-BATCH-1), которые WMS проверяет и проводит сама. Здесь: папка снимков, выбор и проверка снимка (SHA-256),
' загрузка таблицы снимка в лист инструмента (фильтр CSV LibreOffice + importSheet, без цикла Basic по строкам), пакет,
' запись отчёта.
Option Explicit

Public Const TB_VERSION = "1.1.1"
' the contract versions of the snapshot this toolbox reads: 1.0 (WMS 0.6, no vehicles) and 1.1 (WMS 0.7, «Приход авто»)
Public Const TB_CONTRACT = "1.1"
Public Const TB_SNAPSHOT_FORMAT = "WMS-SNAPSHOT-1"
Public Const TB_BATCH_FORMAT = "WMS-BATCH-1"
' the sheet «Настройки» of every tool: B2 the folder of the snapshots, B3 the folder of the batches, B4 the snapshot in use
Public Const TB_SETTINGS = "Настройки"

Global gTbSnap As String
Global gTbAuto As Boolean
Global gTbLastMsg As String

' ================================================================ folders

' the position of the last sFind in s (0 — none): InStrRev of LibreOffice Basic exists only in the VBA mode
Function TbInStrRev(ByVal s As String, ByVal sFind As String) As Long
    Dim i As Long
    For i = Len(s) - Len(sFind) + 1 To 1 Step -1
        If Mid(s, i, Len(sFind)) = sFind Then
            TbInStrRev = i
            Exit Function
        End If
    Next i
End Function

' the folder of this tool book, with a trailing "/"
Function TbFolder() As String
    Dim u As String, i As Long
    u = ThisComponent.getURL()
    For i = Len(u) To 1 Step -1
        If Mid(u, i, 1) = "/" Then
            TbFolder = Left(u, i)
            Exit Function
        End If
    Next i
End Function

' a folder of a setting: absolute (file:///… or /…) or relative to this book («../WMS_Export»); with a trailing "/"
Function TbResolve(ByVal s As String) As String
    Dim sDir As String, parts As Variant, i As Integer
    s = Trim(s)
    If s = "" Then Exit Function
    If Left(s, 8) = "file:///" Then
        TbResolve = s
    ElseIf Left(s, 1) = "/" Then
        TbResolve = ConvertToURL(s)
    Else
        sDir = TbFolder()
        parts = Split(Replace(s, "\", "/"), "/")
        For i = 0 To UBound(parts)
            If parts(i) = ".." Then
                sDir = Left(sDir, TbInStrRev(Left(sDir, Len(sDir) - 1), "/"))
            ElseIf parts(i) <> "." And parts(i) <> "" Then
                sDir = sDir & parts(i) & "/"
            End If
        Next i
        TbResolve = sDir
    End If
    If Right(TbResolve, 1) <> "/" Then TbResolve = TbResolve & "/"
End Function

Function TbSetting(r As Integer) As String
    If ThisComponent.Sheets.hasByName(TB_SETTINGS) Then TbSetting = ThisComponent.Sheets.getByName(TB_SETTINGS).getCellByPosition(1, r).getString()
End Function

Sub TbPutSetting(r As Integer, s As String)
    Dim sh As Object
    sh = ThisComponent.Sheets.getByName(TB_SETTINGS)
    sh.getCellByPosition(1, r).setString(s)
End Sub

' the folder of the snapshots (setting B2, default ../WMS_Export next to the toolbox folder)
Function TbExportDir() As String
    TbExportDir = TbResolve(IIf(TbSetting(1) <> "", TbSetting(1), "../WMS_Export"))
End Function

Function TbBatchDir() As String
    TbBatchDir = TbResolve(IIf(TbSetting(2) <> "", TbSetting(2), "../WMS_Batches"))
End Function

' ================================================================ the snapshot

' the latest complete snapshot (the greatest name WMS_SNAPSHOT_… with a manifest.csv), "" when none
Function TbLatestSnapshot() As String
    Dim sfa As Object, d As String, lst As Variant, i As Long, best As String, nm As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbExportDir()
    If Not sfa.exists(d) Then Exit Function
    lst = sfa.getFolderContents(d, True)
    For i = 0 To UBound(lst)
        nm = Mid(lst(i), Len(d) + 1)
        If Right(nm, 1) = "/" Then nm = Left(nm, Len(nm) - 1)
        If Left(nm, 13) = "WMS_SNAPSHOT_" And sfa.isFolder(lst(i)) Then
            If sfa.exists(d & nm & "/manifest.csv") And nm > best Then best = nm
        End If
    Next i
    If best <> "" Then TbLatestSnapshot = d & best & "/"
End Function

Function TbReadText(url As String) As String
    Dim tis As Object
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(url))
    tis.setEncoding("UTF-8")
    TbReadText = tis.readString(Array(), False)
    tis.closeInput()
End Function

Sub TbWriteText(url As String, s As String)
    Dim tos As Object, sfa As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(url) Then sfa.kill(url)
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(sfa.openFileWrite(url))
    tos.setEncoding("UTF-8")
    tos.writeString(s)
    tos.closeOutput()
End Sub

' the value of a key of manifest.csv ("" when absent)
Function TbManifestValue(snap As String, key As String) As String
    Dim lines As Variant, i As Long, f As Variant
    lines = Split(Replace(TbReadText(snap & "manifest.csv"), Chr(13), ""), Chr(10))
    For i = 0 To UBound(lines)
        f = Split(lines(i), ";")
        If UBound(f) >= 1 Then
            If f(0) = key Then
                TbManifestValue = f(1)
                Exit Function
            End If
        End If
    Next i
End Function

Function TbSha256File(url As String, ByRef nBytes As Double) As String
    Dim nss As Object, dc As Object, inp As Object, buf() As Byte, n As Long, h As Variant, i As Long, s As String
    On Error GoTo EH
    nBytes = 0
    nss = CreateUnoService("com.sun.star.xml.crypto.NSSInitializer")
    dc = nss.getDigestContext(com.sun.star.xml.crypto.DigestID.SHA256, Array())
    inp = CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(url)
    Do
        n = inp.readBytes(buf, 1048576)
        If n <= 0 Then Exit Do
        nBytes = nBytes + n
        ' readBytes returns a sequence of exactly the bytes read
        dc.updateDigest(buf)
    Loop While n = 1048576
    inp.closeInput()
    h = dc.finalizeDigestAndDispose()
    For i = 0 To UBound(h)
        s = s & Right("0" & LCase(Hex(h(i) And 255)), 2)
    Next i
    TbSha256File = s
    Exit Function
EH:
    TbSha256File = ""
End Function

' "" when the snapshot is complete and unchanged: format, every file of the manifest present with its size and SHA-256
Function TbVerifySnapshot(snap As String) As String
    Dim lines As Variant, i As Long, f As Variant, nb As Double, h As String, sfa As Object, nFiles As Long
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(snap & "manifest.csv") Then
        TbVerifySnapshot = "нет manifest.csv"
        Exit Function
    End If
    If TbManifestValue(snap, "format") <> TB_SNAPSHOT_FORMAT Then
        TbVerifySnapshot = "формат снимка «" & TbManifestValue(snap, "format") & "», инструмент ожидает " & TB_SNAPSHOT_FORMAT
        Exit Function
    End If
    lines = Split(Replace(TbReadText(snap & "manifest.csv"), Chr(13), ""), Chr(10))
    For i = 0 To UBound(lines)
        f = Split(lines(i), ";")
        If UBound(f) >= 4 Then
            If f(0) = "file" Then
                nFiles = nFiles + 1
                If Not sfa.exists(snap & f(1)) Then
                    TbVerifySnapshot = "нет файла " & f(1)
                    Exit Function
                End If
                h = TbSha256File(snap & f(1), nb)
                If h <> f(4) Or nb <> Val(f(3)) Then
                    TbVerifySnapshot = "файл " & f(1) & " изменён после снимка (SHA-256 или размер не совпадают)"
                    Exit Function
                End If
            End If
        End If
    Next i
    If nFiles = 0 Then TbVerifySnapshot = "в manifest.csv нет файлов"
    Exit Function
EH:
    TbVerifySnapshot = "снимок не прочитан: " & Error$
End Function

' ================================================================ tables of the snapshot

' column formats of the CSV filter for a table of the contract (1 standard, 2 text, 5 date Y-M-D): numbers and dates are
' read as values, identifiers and names as text (a leading zero of an article is never lost)
Function TbColumnFormats(sFile As String) As String
    Dim t As String, i As Integer, s As String, a As Variant
    Select Case sFile
    Case "stock.csv"
        t = "2222122222"
    Case "orders.csv"
        t = "2222211121122555522222212212"
    Case "special.csv"
        t = "1222212522222211222"
    Case "issues.csv"
        t = "122211252222222112"
    Case "returns.csv"
        t = "112221252221122"
    Case "adjustments.csv"
        t = "1222221112522111222"
    Case "recipients.csv"
        t = "22"
    Case "cars.csv"
        t = "15221112221222"
    Case Else
        t = ""
    End Select
    For i = 1 To Len(t)
        s = s & IIf(s <> "", "/", "") & i & "/" & Mid(t, i, 1)
    Next i
    TbColumnFormats = s
End Function

' loads table sFile of snapshot snap into sheet sDest of this book (the sheet is replaced; hidden when bHide); the number of
' data rows (the first row of the sheet — the column keys of the contract)
Function TbLoadTable(snap As String, sFile As String, sDest As String, bHide As Boolean) As Long
    Dim args(2) As New com.sun.star.beans.PropertyValue, csv As Object, doc As Object, pos As Integer, sh As Object, cur As Object
    doc = ThisComponent
    args(0).Name = "FilterName"
    args(0).Value = "Text - txt - csv (StarCalc)"
    args(1).Name = "FilterOptions"
    args(1).Value = "59,34,76,1," & TbColumnFormats(sFile) & ",1033,false,true,false,false,false"
    args(2).Name = "Hidden"
    args(2).Value = True
    csv = StarDesktop.loadComponentFromURL(snap & sFile, "_blank", 0, args())
    pos = doc.Sheets.getCount()
    If doc.Sheets.hasByName(sDest) Then
        pos = TbSheetIndex(sDest)
        doc.Sheets.removeByName(sDest)
    End If
    doc.Sheets.importSheet(csv, csv.Sheets.getByIndex(0).Name, pos)
    csv.close(True)
    sh = doc.Sheets.getByIndex(pos)
    sh.Name = sDest
    If bHide Then sh.IsVisible = False
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    TbLoadTable = cur.getRangeAddress().EndRow
End Function

' is file sFile listed in the manifest of the snapshot (cars.csv — only from contract 1.1, WMS 0.7)
Function TbHasFile(snap As String, sFile As String) As Boolean
    Dim lines As Variant, i As Long
    lines = Split(Replace(TbReadText(snap & "manifest.csv"), Chr(13), ""), Chr(10))
    For i = 0 To UBound(lines)
        If Left(lines(i), Len(sFile) + 6) = "file;" & sFile & ";" Then
            TbHasFile = CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(snap & sFile)
            Exit Function
        End If
    Next i
End Function

' TbLoadTable for a table a snapshot may lack (cars.csv of a WMS 0.6 snapshot): -1 and an empty sheet sDest (hidden when
' bHide) with the column keys sKeys ("a|b|…") when the file is not in the snapshot
Function TbLoadOptional(snap As String, sFile As String, sDest As String, bHide As Boolean, sKeys As String) As Long
    Dim sh As Object, k As Variant, i As Integer
    If TbHasFile(snap, sFile) Then
        TbLoadOptional = TbLoadTable(snap, sFile, sDest, bHide)
        Exit Function
    End If
    If ThisComponent.Sheets.hasByName(sDest) Then ThisComponent.Sheets.removeByName(sDest)
    ThisComponent.Sheets.insertNewByName(sDest, ThisComponent.Sheets.getCount())
    sh = ThisComponent.Sheets.getByName(sDest)
    k = Split(sKeys, "|")
    For i = 0 To UBound(k)
        sh.getCellByPosition(i, 0).setString(k(i))
    Next i
    If bHide Then sh.IsVisible = False
    TbLoadOptional = -1
End Function

Public Const TB_CAR_KEYS = "visit_no|date|vehicle|supplier|arrived|departed|duration_min|status|weekday|month|day_no|control|note|vehicle_key"

' a CSV file (UTF-8, «;») into sheet sDest of this book (replaced; hidden when bHide) with the column types sTypes of the
' CSV filter («1» standard, «2» text, «5» date Y-M-D per column; "" — every column standard); the number of data rows
Function TbLoadCsv(url As String, sDest As String, bHide As Boolean, sTypes As String) As Long
    Dim args(2) As New com.sun.star.beans.PropertyValue, csv As Object, doc As Object, pos As Integer, sh As Object, i As Integer, s As String
    For i = 1 To Len(sTypes)
        s = s & IIf(s <> "", "/", "") & i & "/" & Mid(sTypes, i, 1)
    Next i
    doc = ThisComponent
    args(0).Name = "FilterName"
    args(0).Value = "Text - txt - csv (StarCalc)"
    args(1).Name = "FilterOptions"
    args(1).Value = "59,34,76,1," & s & ",1033,false,true,false,false,false"
    args(2).Name = "Hidden"
    args(2).Value = True
    csv = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    pos = doc.Sheets.getCount()
    If doc.Sheets.hasByName(sDest) Then
        pos = TbSheetIndex(sDest)
        doc.Sheets.removeByName(sDest)
    End If
    doc.Sheets.importSheet(csv, csv.Sheets.getByIndex(0).Name, pos)
    csv.close(True)
    sh = doc.Sheets.getByIndex(pos)
    sh.Name = sDest
    If bHide Then sh.IsVisible = False
    TbLoadCsv = TbLastRow(sh)
End Function

' the folder of the working WMS: the parent of the folder of the snapshots (…/WMS_Export → …/), with a trailing "/"
Function TbWmsDir() As String
    Dim d As String
    d = TbExportDir()
    TbWmsDir = Left(d, TbInStrRev(Left(d, Len(d) - 1), "/"))
End Function

' python3 <this folder>/sScript with the arguments (already quoted), the output and the errors into sLog (a URL);
' "" — ran (the log exists), otherwise why not
Function TbRunPython(sScript As String, sArgs As String, sLog As String) As String
    Dim sfa As Object, cmd As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(TbFolder() & sScript) Then
        TbRunPython = "нет скрипта " & ConvertFromURL(TbFolder() & sScript) & " — соберите WMS_TOOLBOX заново из выпуска"
        Exit Function
    End If
    If sfa.exists(sLog) Then sfa.kill(sLog)
    cmd = "python3 '" & ConvertFromURL(TbFolder() & sScript) & "' " & sArgs & " > '" & ConvertFromURL(sLog) & "' 2>&1"
    Shell("/bin/sh", 0, "-c """ & cmd & """", True)
    If Not sfa.exists(sLog) Then TbRunPython = "скрипт не выполнился (нужен python3): " & cmd
End Function

' the key of a vehicle — the rule of WMS (WmsCar.PlateKey): upper case, the Latin look-alikes of the letters of a plate as
' Cyrillic, Ё → Е, only letters and digits; a Russian plate in the text (letter, three digits, two letters; the last one)
' without the region, otherwise the whole text: «Газель А123ВС 77», «а 123 вс» and «A123BC» are one vehicle
Function TbPlateKey(ByVal s As String) As String
    Dim i As Long, ch As String, p As Integer, t As String, a As Long
    s = UCase(s)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        p = InStr(1, "ABEKMHOPCTYX", ch, 0)
        If p > 0 Then ch = Mid("АВЕКМНОРСТУХ", p, 1)
        If ch = "Ё" Then ch = "Е"
        a = Asc(ch)
        If (a >= 48 And a <= 57) Or (a >= 65 And a <= 90) Or (a >= 1040 And a <= 1071) Then t = t & ch
    Next i
    For i = Len(t) - 5 To 1 Step -1
        If PlateLetter(Mid(t, i, 1)) And PlateDigit(Mid(t, i + 1, 1)) And PlateDigit(Mid(t, i + 2, 1)) And PlateDigit(Mid(t, i + 3, 1)) _
            And PlateLetter(Mid(t, i + 4, 1)) And PlateLetter(Mid(t, i + 5, 1)) Then
            TbPlateKey = Mid(t, i, 6)
            Exit Function
        End If
    Next i
    TbPlateKey = t
End Function

Private Function PlateLetter(ch As String) As Boolean
    PlateLetter = (InStr(1, "АВЕКМНОРСТУХ", ch, 0) > 0)
End Function

Private Function PlateDigit(ch As String) As Boolean
    PlateDigit = (ch >= "0" And ch <= "9" And Len(ch) = 1)
End Function

Function TbSheetIndex(sName As String) As Integer
    Dim i As Integer
    For i = 0 To ThisComponent.Sheets.getCount() - 1
        If ThisComponent.Sheets.getByIndex(i).Name = sName Then
            TbSheetIndex = i
            Exit Function
        End If
    Next i
    TbSheetIndex = -1
End Function

' the snapshot to use: the one of setting B4 when it still exists, otherwise the latest; verified. "" and why when none
Function TbPickSnapshot(ByRef why As String) As String
    Dim snap As String
    snap = TbLatestSnapshot()
    If snap = "" Then
        why = "в папке " & ConvertFromURL(TbExportDir()) & " нет снимков — в WMS нажмите «Экспорт для инструментов» на листе «Главная»"
        Exit Function
    End If
    why = TbVerifySnapshot(snap)
    If why <> "" Then
        why = "снимок " & ConvertFromURL(snap) & ": " & why
        Exit Function
    End If
    TbPickSnapshot = snap
End Function

' the first day of month m of year y as a date number; m may be outside 1..12 (DateSerial of LibreOffice Basic refuses that)
Function TbMonthStart(ByVal y As Integer, ByVal m As Integer) As Double
    Do While m < 1
        m = m + 12
        y = y - 1
    Loop
    Do While m > 12
        m = m - 12
        y = y + 1
    Loop
    TbMonthStart = CDbl(DateSerial(y, m, 1))
End Function

' the last used row of a sheet (0-based; 0 — only the header or nothing)
Function TbLastRow(sh As Object) As Long
    Dim cur As Object
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    TbLastRow = cur.getRangeAddress().EndRow
End Function

' ================================================================ CSV and batches

Function TbCsv(v As Variant) As String
    Dim s As String
    Select Case VarType(v)
    Case 0, 1
        TbCsv = ""
    Case 2, 3, 4, 5, 6
        TbCsv = Trim(Str(v))
    Case Else
        s = CStr(v)
        If InStr(s, ";") > 0 Or InStr(s, Chr(34)) > 0 Or InStr(s, Chr(10)) > 0 Or InStr(s, Chr(13)) > 0 Then
            TbCsv = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
        Else
            TbCsv = s
        End If
    End Select
End Function

' writes a batch file (WMS-BATCH-1) into the folder of the batches: the path of the file
Function TbWriteBatch(sTarget As String, sId As String, sSource As String, sSnapshot As String, keys As Variant, rows As Variant, _
    sExtra As String) As String
    Dim s As String, i As Long, j As Integer, a() As String, sfa As Object, d As String
    s = "#" & TB_BATCH_FORMAT & Chr(10) & "#id;" & sId & Chr(10) & "#target;" & sTarget & Chr(10) & "#source;" & sSource & Chr(10) _
        & "#snapshot;" & sSnapshot & Chr(10) & sExtra & Join(keys, ";") & Chr(10)
    For i = 0 To UBound(rows)
        ReDim a(UBound(rows(i)))
        For j = 0 To UBound(rows(i))
            a(j) = TbCsv(rows(i)(j))
        Next j
        s = s & Join(a, ";") & Chr(10)
    Next i
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbBatchDir()
    If Not sfa.exists(d) Then sfa.createFolder(d)
    TbWriteText(d & sId & ".csv", s)
    TbWriteBatch = d & sId & ".csv"
End Function

' «Все инструменты» (every tool book): the launcher WMS_TOOLBOX.ods of this folder; an open one comes to the front
Function TbOpenLauncher() As String
    Dim url As String, doc As Object, args(0) As New com.sun.star.beans.PropertyValue
    url = TbFolder() & "WMS_TOOLBOX.ods"
    If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(url) Then
        TbOpenLauncher = "ERR:нет файла " & ConvertFromURL(url) & " — книги инструментов лежат в одной папке WMS_TOOLBOX"
        Exit Function
    End If
    args(0).Name = "MacroExecutionMode"
    args(0).Value = 4
    doc = StarDesktop.loadComponentFromURL(url, "_default", 0, args())
    If IsNull(doc) Then TbOpenLauncher = "ERR:не открылся " & ConvertFromURL(url) Else TbOpenLauncher = "OK:WMS_TOOLBOX"
End Function

Sub BtnTbLauncher(Optional oEvent As Variant)
    Dim res As String
    res = TbOpenLauncher()
    If Left(res, 3) <> "OK:" Then TbMsg(res)
End Sub

' a message for the user (recorded for the tests, no window in the test mode): «OK:…» — information, «ERR:…» — a refusal
' with the warning icon; the prefixes themselves are not shown
Sub TbMsg(s As String)
    gTbLastMsg = s
    If gTbAuto Then Exit Sub
    If Left(s, 4) = "ERR:" Then
        MsgBox Mid(s, 5), 48, "WMS_TOOLBOX — не выполнено"
    ElseIf Left(s, 3) = "OK:" Then
        MsgBox Mid(s, 4), 64, "WMS_TOOLBOX"
    Else
        MsgBox s, 64, "WMS_TOOLBOX"
    End If
End Sub

Function TbTestAuto(b As Boolean) As String
    gTbAuto = b
    TbTestAuto = "OK"
End Function

Function TbTestLastMessage() As String
    TbTestLastMessage = gTbLastMsg
End Function
