' TbDocs — WMS_DOCS: акты ODT из движений WMS (задание «FINAL WMS MARATHON», §6). Лист «Акт»: вид акта (акт передачи,
' передача деталей — у каждой позиции артикул, передача сторонней организации — организация обязательна), дата, передал,
' принял, организация и позиции (ЕИ, наименование, артикул, единица, количество, примечание) — вручную или «Подобрать по
' выдачам» (получатель и период из снимка). «Создать документ» — документ Writer (ODT, и PDF по желанию) с номером,
' датой, сторонами, таблицей позиций и подписями в папке ../WMS_Docs; лист «История» — созданные документы; номер и
' история сохраняются вместе с книгой инструмента, существующий акт с тем же номером не перезаписывается. Книгу WMS
' инструмент не открывает.
Option Explicit

Private Const SH_ACT = "Акт"
Private Const SH_HIST = "История"
Private Const POS_ROW = 8            ' 0-based: the positions start at row 9 (row 8 — their headers)
Public Const KIND_PARTS = "Передача деталей"
Public Const KIND_THIRD = "Передача сторонней организации"

Sub BtnDocsRefresh(Optional oEvent As Variant)
    Dim why As String, snap As String, n As Long
    snap = TbPickSnapshot(why)
    If snap = "" Then
        TbMsg("ERR:" & why)
        Exit Sub
    End If
    n = TbLoadTable(snap, "issues.csv", "_issues", True)
    TbLoadTable(snap, "stock.csv", "_stock", True)
    TbPutSetting(3, snap)
    TbMsg("OK:выдач в снимке: " & n)
End Sub

Sub BtnDocsPick(Optional oEvent As Variant)
    TbMsg(DocsPick())
End Sub

' «Подобрать по выдачам»: the posted issues of the recipient (B4 «Принял») in the period B5..B6 (dates) into the positions
Function DocsPick() As String
    Dim act As Object, iss As Variant, who As String, a As Double, b As Double, i As Long, k As Long, out() As Variant, last As Long, n As Long
    On Error GoTo EH
    If Not ThisComponent.Sheets.hasByName("_issues") Then
        DocsPick = "ERR:сначала «Обновить из снимка»"
        Exit Function
    End If
    act = ThisComponent.Sheets.getByName(SH_ACT)
    who = LCase(Trim(act.getCellByPosition(1, 3).getString()))
    a = act.getCellByPosition(1, 4).getValue()
    b = act.getCellByPosition(1, 5).getValue()
    If who = "" Or a = 0 Or b = 0 Then
        DocsPick = "ERR:укажите «Принял» (B4) и период (B5, B6)"
        Exit Function
    End If
    n = TbLastRow(ThisComponent.Sheets.getByName("_issues"))
    If n < 1 Then
        DocsPick = "ERR:в снимке нет выдач"
        Exit Function
    End If
    iss = ThisComponent.Sheets.getByName("_issues").getCellRangeByPosition(0, 1, 17, n).getDataArray()
    ReDim out(n)
    For i = 0 To UBound(iss)
        If Left(CStr(iss(i)(17)), 9) = "Проведено" And LCase(Trim(CStr(iss(i)(8)))) = who And VarType(iss(i)(7)) = 5 Then
            If iss(i)(7) >= a And iss(i)(7) <= b Then
                out(k) = Array(iss(i)(11), iss(i)(2), iss(i)(3), iss(i)(6), iss(i)(4), "выдача № " & iss(i)(0) & " от " & Format(CDate(iss(i)(7)), "DD.MM.YYYY"))
                k = k + 1
            End If
        End If
    Next i
    last = TbLastRow(act)
    If last > POS_ROW Then act.getCellRangeByPosition(0, POS_ROW + 1, 5, last).clearContents(1023)
    If k > 0 Then
        Dim rows() As Variant, j As Long
        ReDim rows(k - 1)
        For j = 0 To k - 1
            rows(j) = out(j)
        Next j
        act.getCellRangeByPosition(0, POS_ROW + 1, 5, POS_ROW + k).setDataArray(rows)
    End If
    DocsPick = "OK:позиций " & k
    Exit Function
EH:
    DocsPick = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnDocsCreate(Optional oEvent As Variant)
    TbMsg(DocsCreate(False))
End Sub

' «Создать документ»: the ODT (and PDF when bPdf or setting B6 = «да») of the act; the number from setting B7
Function DocsCreate(bPdf As Boolean) As String
    Dim act As Object, kind As String, dt As Double, who1 As String, who2 As String, org As String, last As Long, d As Variant, i As Long, n As Long
    Dim num As Long, doc As Object, txt As Object, oCur As Object, tbl As Object, sfa As Object, sDir As String, url As String, args(0) As New com.sun.star.beans.PropertyValue
    Dim hist As Object, hr As Long, k As Long, title As String, fl As Variant
    On Error GoTo EH
    act = ThisComponent.Sheets.getByName(SH_ACT)
    kind = Trim(act.getCellByPosition(1, 0).getString())
    If kind = "" Then kind = "Акт передачи"
    dt = act.getCellByPosition(1, 1).getValue()
    If dt = 0 Then dt = Int(CDbl(Now()))
    who1 = Trim(act.getCellByPosition(1, 2).getString())
    who2 = Trim(act.getCellByPosition(1, 3).getString())
    org = Trim(act.getCellByPosition(1, 6).getString())
    last = TbLastRow(act)
    If last <= POS_ROW Then
        DocsCreate = "ERR:нет позиций (строки с " & (POS_ROW + 2) & ")"
        Exit Function
    End If
    d = act.getCellRangeByPosition(0, POS_ROW + 1, 5, last).getDataArray()
    For i = 0 To UBound(d)
        If CStr(d(i)(0)) <> "" Or CStr(d(i)(1)) <> "" Then n = n + 1
    Next i
    If n = 0 Or who1 = "" Or who2 = "" Then
        DocsCreate = "ERR:укажите «Передал» (B3), «Принял» (B4) и хотя бы одну позицию"
        Exit Function
    End If
    If kind = KIND_THIRD And org = "" Then
        DocsCreate = "ERR:«" & KIND_THIRD & "»: укажите «Организация» (B7)"
        Exit Function
    End If
    If kind = KIND_PARTS Then
        k = 0
        For i = 0 To UBound(d)
            If CStr(d(i)(0)) <> "" Or CStr(d(i)(1)) <> "" Then
                k = k + 1
                If Trim(CStr(d(i)(2))) = "" Then
                    DocsCreate = "ERR:«" & KIND_PARTS & "»: у позиции " & k & " (строка " & (POS_ROW + 2 + i) & ") нет артикула"
                    Exit Function
                End If
            End If
        Next i
        k = 0
    End If
    num = Val(TbSetting(6))
    If num < 1 Then num = 1
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    sDir = TbResolve(IIf(TbSetting(7) <> "", TbSetting(7), "../WMS_Docs"))
    ' a Latin file name: a URL with Cyrillic letters is not accepted everywhere
    url = sDir & "ACT-" & Right("00000" & num, 5) & "_" & Format(CDate(dt), "YYYY-MM-DD") & ".odt"
    ' (Filter() of Basic exists only in the VBA mode: a loop)
    If sfa.exists(sDir) Then
        fl = sfa.getFolderContents(sDir, False)
        For i = 0 To UBound(fl)
            If InStr(fl(i), "/ACT-" & Right("00000" & num, 5) & "_") > 0 Then
                DocsCreate = "ERR:акт № " & num & " уже есть в " & ConvertFromURL(sDir) & " — проверьте «Номер следующего акта» («Настройки», B7); " _
                    & "созданный акт не перезаписывается"
                Exit Function
            End If
        Next i
    End If
    title = UCase(Left(kind, 1)) & Mid(kind, 2) & " № " & num & " от " & Format(CDate(dt), "DD.MM.YYYY")
    args(0).Name = "Hidden"
    args(0).Value = True
    doc = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, args())
    txt = doc.getText()
    oCur = txt.createTextCursor()
    oCur.CharHeight = 14
    oCur.CharWeight = 150
    oCur.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
    txt.insertString(oCur, title, False)
    txt.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    oCur.CharHeight = 11
    oCur.CharWeight = 100
    oCur.ParaAdjust = com.sun.star.style.ParagraphAdjust.LEFT
    txt.insertString(oCur, "Передал: " & who1, False)
    txt.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    txt.insertString(oCur, "Принял: " & who2 & IIf(org <> "", " (" & org & ")", ""), False)
    txt.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    tbl = doc.createInstance("com.sun.star.text.TextTable")
    tbl.initialize(n + 1, 7)
    txt.insertTextContent(oCur, tbl, False)
    tbl.getCellRangeByName("A1:G1").setDataArray(Array(Array("№", "ЕИ", "Наименование", "Артикул", "Ед.", "Количество", "Примечание")))
    For i = 0 To UBound(d)
        If CStr(d(i)(0)) <> "" Or CStr(d(i)(1)) <> "" Then
            k = k + 1
            tbl.getCellByName("A" & (k + 1)).setString(CStr(k))
            tbl.getCellByName("B" & (k + 1)).setString(CStr(d(i)(0)))
            tbl.getCellByName("C" & (k + 1)).setString(CStr(d(i)(1)))
            tbl.getCellByName("D" & (k + 1)).setString(CStr(d(i)(2)))
            tbl.getCellByName("E" & (k + 1)).setString(CStr(d(i)(3)))
            tbl.getCellByName("F" & (k + 1)).setString(Replace(CStr(d(i)(4)), ".", ","))
            tbl.getCellByName("G" & (k + 1)).setString(CStr(d(i)(5)))
        End If
    Next i
    txt.insertControlCharacter(txt.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    txt.insertString(txt.getEnd(), "Всего позиций: " & n, False)
    txt.insertControlCharacter(txt.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    txt.insertControlCharacter(txt.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    txt.insertString(txt.getEnd(), "Передал: ____________________ / " & who1 & " /", False)
    txt.insertControlCharacter(txt.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    txt.insertControlCharacter(txt.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    txt.insertString(txt.getEnd(), "Принял: ____________________ / " & who2 & " /", False)
    If Not sfa.exists(sDir) Then sfa.createFolder(sDir)
    args(0).Name = "FilterName"
    args(0).Value = "writer8"
    doc.storeToURL(url, args())
    If bPdf Or LCase(Left(Trim(TbSetting(5)), 1)) = "д" Then
        args(0).Value = "writer_pdf_Export"
        doc.storeToURL(Left(url, Len(url) - 4) & ".pdf", args())
    End If
    doc.close(True)
    hist = ThisComponent.Sheets.getByName(SH_HIST)
    hr = TbLastRow(hist) + 1
    hist.getCellRangeByPosition(0, hr, 5, hr).setDataArray(Array(Array(num, Format(CDate(dt), "DD.MM.YYYY"), kind, who2, n, ConvertFromURL(url))))
    TbPutSetting(6, CStr(num + 1))
    DocsCreate = "OK:" & title & " — " & ConvertFromURL(url) & SaveSelf()
    Exit Function
EH:
    DocsCreate = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' the number and the history live in this book: it is saved at once ("" — saved; otherwise what to do)
Private Function SaveSelf() As String
    On Error GoTo EH
    ThisComponent.store()
    Exit Function
EH:
    SaveSelf = "; ВНИМАНИЕ: книга WMS_DOCS не сохранена (" & Error$ & ") — сохраните её (Ctrl+S), иначе номер и история акта пропадут"
End Function

' the compile check of the toolbox calls one function of every module (tools/tb_compile_check.py)
Function DocsProbe() As String
    DocsProbe = "OK"
End Function
