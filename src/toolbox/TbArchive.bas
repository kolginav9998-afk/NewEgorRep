' TbArchive — WMS_ARCHIVE: безопасная работа с историей (задание «FINAL WMS MARATHON», §6). Сначала только анализ и
' пакеты — история из рабочей WMS никогда не удаляется автоматически. «Анализ» — размер книги и журнала, строк по
' месяцам на каждом листе движений, какие месяцы закрыты (старше порога «Настройки», без непроведённых строк) и могут
' войти в архив. «Создать архивный пакет» — папка ../WMS_Archive/ARCHIVE_<с>_<по>/: строки движений этих месяцев из
' снимка, файлы журнала этих месяцев и manifest.csv с SHA-256 каждого файла; пакет проверяется сразу после записи.
' «Открыть пакет» — таблицы пакета только для просмотра. Физическое сокращение рабочей книги — только отдельной
' подтверждаемой процедурой release/upgrade с тестом восстановления (в инструмент не входит).
Option Explicit

Private Const SH_AN = "Анализ"

Sub BtnArcAnalyze(Optional oEvent As Variant)
    TbMsg(ArcAnalyze(0))
End Sub

Private Function Months(t As String, dateCol As Integer, n As Long, ByRef mon() As String, ByRef cnt() As Long, ByRef nm As Integer) As Long
    Dim d As Variant, i As Long, m As String, j As Integer, found As Boolean
    If n < 1 Then Exit Function
    d = ThisComponent.Sheets.getByName(t).getCellRangeByPosition(dateCol, 1, dateCol, n).getDataArray()
    For i = 0 To UBound(d)
        If VarType(d(i)(0)) = 5 Then
            m = Format(CDate(d(i)(0)), "YYYY-MM")
            found = False
            For j = 0 To nm - 1
                If mon(j) = m Then
                    cnt(j) = cnt(j) + 1
                    found = True
                    Exit For
                End If
            Next j
            If Not found And nm <= UBound(mon) Then
                mon(nm) = m
                cnt(nm) = 1
                nm = nm + 1
            End If
            Months = Months + 1
        End If
    Next i
End Function

' dToday: the date the age of a month is counted from (0 — today; the tests fix it)
Function ArcAnalyze(dToday As Double) As String
    Dim why As String, snap As String, t As Variant, dc As Variant, i As Integer, n As Long, sh As Object, r As Long, mon(240) As String, cnt(240) As Long
    Dim nm As Integer, j As Integer, sfa As Object, book As String, keep As Long, a As String, closedTo As String, last As Long, tmp As String, tc As Long
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        ArcAnalyze = "ERR:" & why
        Exit Function
    End If
    If dToday = 0 Then dToday = Int(CDbl(Now()))
    keep = Val(TbSetting(5))
    If keep < 1 Then keep = 12
    t = Array("orders", "special", "issues", "returns", "adjustments")
    dc = Array(13, 7, 7, 7, 10)
    For i = 0 To 4
        n = TbLoadTable(snap, t(i) & ".csv", "_" & t(i), True)
        Months("_" & t(i), dc(i), n, mon, cnt, nm)
    Next i
    ' months in order
    For i = 1 To nm - 1
        tmp = mon(i)
        tc = cnt(i)
        j = i - 1
        Do While j >= 0
            If mon(j) <= tmp Then Exit Do
            mon(j + 1) = mon(j)
            cnt(j + 1) = cnt(j)
            j = j - 1
        Loop
        mon(j + 1) = tmp
        cnt(j + 1) = tc
    Next i
    sh = ThisComponent.Sheets.getByName(SH_AN)
    last = TbLastRow(sh)
    If last >= 0 Then sh.getCellRangeByPosition(0, 0, 3, last + 1).clearContents(1023)
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    book = ConvertToURL(TbManifestValue(snap, "book"))
    sh.getCellByPosition(0, 0).setString("Книга WMS")
    sh.getCellByPosition(1, 0).setString(ConvertFromURL(book))
    sh.getCellByPosition(0, 1).setString("Размер книги, МБ")
    If sfa.exists(book) Then sh.getCellByPosition(1, 1).setValue(Int(sfa.getSize(book) / 10485.76) / 100)
    sh.getCellByPosition(0, 2).setString("Хранить в рабочей книге, месяцев")
    sh.getCellByPosition(1, 2).setValue(keep)
    ' keep N months: the current month and the N − 1 before it stay in the working book; older closed months may be archived
    a = Format(CDate(TbMonthStart(Year(CDate(dToday)), Month(CDate(dToday)) - keep + 1)), "YYYY-MM")
    r = 4
    sh.getCellRangeByPosition(0, r, 2, r).setDataArray(Array(Array("Месяц", "Строк движений", "Можно в архив")))
    r = r + 1
    For i = 0 To nm - 1
        sh.getCellRangeByPosition(0, r, 2, r).setDataArray(Array(Array(mon(i), cnt(i), IIf(mon(i) < a, "да", "нет"))))
        If mon(i) < a Then closedTo = mon(i)
        r = r + 1
    Next i
    ArcAnalyze = "OK:месяцев с движениями " & nm & IIf(closedTo <> "", ", можно в архив — до " & closedTo & " включительно", ", архивировать пока нечего")
    Exit Function
EH:
    ArcAnalyze = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnArcPackage(Optional oEvent As Variant)
    TbMsg(ArcPackage(Trim(TbSetting(6)), Trim(TbSetting(7))))
End Sub

' the package of the months sFrom..sTo (ГГГГ-ММ): the rows of the movement tables dated in these months, the journal files
' of these months, manifest.csv with the SHA-256 of every file; verified after writing
Function ArcPackage(sFrom As String, sTo As String) As String
    Dim why As String, snap As String, sfa As Object, sArc As String, dst As String, t As Variant, dc As Variant, i As Integer, j As Long, txt As String
    Dim lines As Variant, keep() As String, k As Long, m As String, man As String, h As String, nb As Double, files As Variant
    Dim ok As String, nRows As Long, jn As String, part() As String, q As Long
    On Error GoTo EH
    If Len(sFrom) <> 7 Or Len(sTo) <> 7 Or sFrom > sTo Then
        ArcPackage = "ERR:укажите месяцы «с» и «по» в формате ГГГГ-ММ (лист «Настройки», B7 и B8)"
        Exit Function
    End If
    snap = TbPickSnapshot(why)
    If snap = "" Then
        ArcPackage = "ERR:" & why
        Exit Function
    End If
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    sArc = TbResolve(IIf(TbSetting(8) <> "", TbSetting(8), "../WMS_Archive"))
    If Not sfa.exists(sArc) Then sfa.createFolder(sArc)
    dst = sArc & "ARCHIVE_" & sFrom & "_" & sTo & "/"
    If sfa.exists(dst) Then
        ArcPackage = "ERR:пакет " & ConvertFromURL(dst) & " уже есть — пакеты не перезаписываются"
        Exit Function
    End If
    sfa.createFolder(dst)
    sfa.createFolder(dst & "journal/")
    man = "key;value" & Chr(10) & "format;WMS-ARCHIVE-1" & Chr(10) & "from;" & sFrom & Chr(10) & "to;" & sTo & Chr(10) & "snapshot;" _
        & Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1, 60) & Chr(10) & "instance_id;" & TbManifestValue(snap, "instance_id") & Chr(10) _
        & "created;" & Format(Now(), "YYYY-MM-DD") & "T" & Format(Now(), "HH:MM:SS") & Chr(10)
    t = Array("orders", "special", "issues", "returns", "adjustments")
    dc = Array(13, 7, 7, 7, 10)
    For i = 0 To 4
        ' the date column of the contract is ГГГГ-ММ-ДД: its first 7 characters are the month (the field of a CSV line)
        txt = TbReadText(snap & t(i) & ".csv")
        lines = Split(txt, Chr(10))
        ReDim keep(UBound(lines))
        keep(0) = lines(0)
        k = 1
        For j = 1 To UBound(lines)
            If lines(j) <> "" Then
                m = Left(FieldOf(lines(j), dc(i)), 7)
                If m >= sFrom And m <= sTo And Len(m) = 7 Then
                    keep(k) = lines(j)
                    k = k + 1
                End If
            End If
        Next j
        ReDim part(k - 1)
        For q = 0 To k - 1
            part(q) = keep(q)
        Next q
        TbWriteText(dst & t(i) & ".csv", Join(part, Chr(10)) & Chr(10))
        nRows = nRows + k - 1
        h = TbSha256File(dst & t(i) & ".csv", nb)
        man = man & "file;" & t(i) & ".csv;" & (k - 1) & ";" & Trim(Str(nb)) & ";" & h & Chr(10)
    Next i
    files = sfa.getFolderContents(snap & "journal/", False)
    For i = 0 To UBound(files)
        jn = Mid(files(i), TbInStrRev(files(i), "/") + 1)
        m = Mid(jn, Len("WMS_journal_") + 1, 7)
        If m >= sFrom And m <= sTo Then
            sfa.copy(files(i), dst & "journal/" & jn)
            h = TbSha256File(dst & "journal/" & jn, nb)
            man = man & "file;journal/" & jn & ";-1;" & Trim(Str(nb)) & ";" & h & Chr(10)
        End If
    Next i
    TbWriteText(dst & "manifest.csv", man)
    ok = ArcVerify(dst)
    If ok <> "" Then
        ArcPackage = "ERR:пакет записан, но не прошёл проверку: " & ok
        Exit Function
    End If
    ArcPackage = "OK:архивный пакет " & ConvertFromURL(dst) & ": строк движений " & nRows & "; проверен (SHA-256). Рабочая книга не изменялась"
    Exit Function
EH:
    ArcPackage = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' the value of field k (0-based) of a CSV line of the contract (quotes respected)
Function FieldOf(ln As String, k As Integer) As String
    Dim i As Long, n As Integer, inQ As Boolean, ch As String, s As String, q As String
    q = Chr(34)
    For i = 1 To Len(ln)
        ch = Mid(ln, i, 1)
        If ch = q Then
            inQ = Not inQ
        ElseIf ch = ";" And Not inQ Then
            If n = k Then
                FieldOf = s
                Exit Function
            End If
            n = n + 1
            s = ""
        Else
            s = s & ch
        End If
    Next i
    If n = k Then FieldOf = s
End Function

' "" when every file of the package matches its manifest
Function ArcVerify(dst As String) As String
    Dim lines As Variant, i As Long, f As Variant, h As String, nb As Double
    lines = Split(TbReadText(dst & "manifest.csv"), Chr(10))
    For i = 0 To UBound(lines)
        f = Split(lines(i), ";")
        If UBound(f) >= 4 Then
            If f(0) = "file" Then
                h = TbSha256File(dst & f(1), nb)
                If h <> f(4) Then
                    ArcVerify = "файл " & f(1) & ": SHA-256 не совпадает"
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Sub BtnArcOpen(Optional oEvent As Variant)
    Dim fp As Object, url As String, n As Long
    fp = CreateUnoService("com.sun.star.ui.dialogs.FolderPicker")
    fp.setDisplayDirectory(TbResolve(IIf(TbSetting(8) <> "", TbSetting(8), "../WMS_Archive")))
    If fp.execute() <> 1 Then Exit Sub
    url = fp.getDirectory()
    If Right(url, 1) <> "/" Then url = url & "/"
    TbMsg(ArcOpen(url))
End Sub

' a package for reading: verified, its tables into the sheets «Архив: …»
Function ArcOpen(url As String) As String
    Dim why As String, t As Variant, i As Integer, s As String
    On Error GoTo EH
    why = ArcVerify(url)
    If why <> "" Then
        ArcOpen = "ERR:" & why
        Exit Function
    End If
    t = Array("orders", "special", "issues", "returns", "adjustments")
    For i = 0 To UBound(t)
        s = s & IIf(s <> "", ", ", "") & t(i) & " " & TbLoadTable(url, t(i) & ".csv", "Архив " & t(i), False)
    Next i
    ArcOpen = "OK:" & s
    Exit Function
EH:
    ArcOpen = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function
