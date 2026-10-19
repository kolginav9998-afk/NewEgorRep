' WmsPrint — 1.0: документы на бумаге. Одна страница «Печать», её заполняет документ — акт расхождения с поставщиком (З12),
' акт списания (К3), опись и сличительная ведомость (К4), накладная выдачи с подписями (К11, К12), карточка ЕИ (К14), отчёты;
' «Печать…» — на принтер, «Сохранить PDF» — файл в папке «WMS_Документы» рядом с книгой, без интернета. Страница — вид:
' ничего не проводит, в журнал не пишет. Library WmsPlus.
Option Explicit

Public Const SH_PRINT = "Печать"
Public Const PR_COLS = 7
Public Const PR_MAX = 2000
Public Const DOCS_DIR = "WMS_Документы"

' the document on the page: the next free row, its name (the file of the PDF), the sheet it was made from («Назад»)
Global gPrRow As Long
Global gPrTitle As String
Global gPrFrom As String
Global gPrLastPdf As String

' ================================================================ the page

' the page emptied for a new document: its title (bold, large) and the name of the PDF file
Function PrintBegin(sTitle As String, sFile As String) As Object
    Dim sh As Object, rng As Object
    gPrFrom = gDoc.getCurrentController().getActiveSheet().getName()
    If gPrFrom = SH_PRINT Then gPrFrom = ""
    sh = WmsView.ViewOpen(SH_PRINT)
    rng = sh.getCellRangeByPosition(0, 0, PR_COLS - 1, PR_MAX)
    rng.clearContents(1 + 2 + 4 + 16 + 32)
    rng.CharHeight = 10
    rng.CharColor = RGB(38, 50, 56)
    rng.IsTextWrapped = False
    sh.getCellByPosition(0, 0).setString(sTitle)
    sh.getCellByPosition(0, 0).CharHeight = 14
    sh.getCellByPosition(0, 0).CharWeight = 150
    gPrRow = 2
    gPrTitle = sFile
    PrintBegin = sh
End Function

' one line of the head of the document («Поставщик: …», «Основание: …»)
Sub PrintLine(sh As Object, s As String, Optional bBold As Variant)
    sh.getCellByPosition(0, gPrRow).setString(s)
    If Not IsMissing(bBold) Then
        If bBold Then sh.getCellByPosition(0, gPrRow).CharWeight = 150
    End If
    gPrRow = gPrRow + 1
End Sub

' a table: the headers (up to 7, columns A..G), n rows; numeric columns (numbers of «nums», «3,5») right-aligned
Sub PrintTable(sh As Object, heads As Variant, rows As Variant, n As Long, nums As String)
    Dim nc As Integer, r0 As Long, i As Long, hd As Object, body As Object, c As Integer, a() As Variant, row() As Variant, j As Integer
    nc = UBound(heads) + 1
    If nc > PR_COLS Then nc = PR_COLS
    gPrRow = gPrRow + 1
    r0 = gPrRow
    hd = sh.getCellRangeByPosition(0, r0, nc - 1, r0)
    ReDim row(nc - 1)
    For j = 0 To nc - 1
        row(j) = heads(j)
    Next j
    hd.setDataArray(Array(row))
    hd.CharWeight = 150
    hd.CharHeight = 9
    hd.IsTextWrapped = True
    hd.VertJustify = 2
    hd.CellBackColor = RGB(231, 236, 242)
    If n > 0 Then
        ReDim a(n - 1)
        For i = 0 To n - 1
            ReDim row(nc - 1)
            For j = 0 To nc - 1
                If j <= UBound(rows(i)) Then row(j) = rows(i)(j) Else row(j) = ""
            Next j
            a(i) = row
        Next i
        body = sh.getCellRangeByPosition(0, r0 + 1, nc - 1, r0 + n)
        body.setDataArray(a)
        body.IsTextWrapped = True
        body.VertJustify = 1
        For j = 0 To nc - 1
            If InStr("," & nums & ",", "," & j & ",") > 0 Then sh.getCellRangeByPosition(j, r0 + 1, j, r0 + n).HoriJustify = 3
        Next j
    End If
    Grid(sh.getCellRangeByPosition(0, r0, nc - 1, r0 + IIf(n > 0, n, 0)))
    gPrRow = r0 + n + 2
End Sub

' thin lines around and between the cells of a table
Private Sub Grid(rng As Object)
    Dim b As New com.sun.star.table.TableBorder2, ln As New com.sun.star.table.BorderLine2
    ln.OuterLineWidth = 18
    ln.LineWidth = 18
    ln.Color = RGB(120, 130, 140)
    b.TopLine = ln
    b.BottomLine = ln
    b.LeftLine = ln
    b.RightLine = ln
    b.HorizontalLine = ln
    b.VerticalLine = ln
    b.IsTopLineValid = True
    b.IsBottomLineValid = True
    b.IsLeftLineValid = True
    b.IsRightLineValid = True
    b.IsHorizontalLineValid = True
    b.IsVerticalLineValid = True
    rng.TableBorder2 = b
End Sub

' the lines of the signatures: «Сдал ____________ / ____________ /»
Sub PrintSign(sh As Object, labels As Variant)
    Dim i As Integer
    gPrRow = gPrRow + 1
    For i = 0 To UBound(labels)
        sh.getCellByPosition(0, gPrRow).setString(labels(i) & ":   ______________________   /   ______________________   /")
        sh.getCellByPosition(0, gPrRow + 1).setString("                    подпись                                   фамилия, инициалы")
        sh.getCellByPosition(0, gPrRow + 1).CharHeight = 7
        sh.getCellByPosition(0, gPrRow + 1).CharColor = RGB(120, 130, 140)
        gPrRow = gPrRow + 3
    Next i
End Sub

' the document finished: the print area (A..G of its rows), the page shown. OK:<name of the document>
Function PrintDone(sh As Object) As String
    Dim pr(0) As New com.sun.star.table.CellRangeAddress
    On Error Resume Next
    pr(0).Sheet = sh.getRangeAddress().Sheet
    pr(0).StartColumn = 0
    pr(0).StartRow = 0
    pr(0).EndColumn = PR_COLS - 1
    pr(0).EndRow = gPrRow
    sh.setPrintAreas(pr())
    sh.getCellByPosition(0, gPrRow + 1).setString("WMS · " & Format(Now(), "DD.MM.YYYY HH:MM"))
    sh.getCellByPosition(0, gPrRow + 1).CharHeight = 7
    sh.getCellByPosition(0, gPrRow + 1).CharColor = RGB(150, 150, 150)
    On Error GoTo 0
    WmsView.ViewDone(sh)
    WmsView.GoCell(SH_PRINT, 0, 0)
    PrintDone = "OK:" & gPrTitle
    gUiLastMsg = PrintDone
End Function

' ================================================================ the buttons of «Печать»

Sub BtnPrint(Optional oEvent As Variant)
    Dim d As Object
    WmsInit()
    WmsView.GoCell(SH_PRINT, 0, 0)
    d = CreateUnoService("com.sun.star.frame.DispatchHelper")
    d.executeDispatch(gDoc.getCurrentController().getFrame(), ".uno:Print", "", 0, Array())
End Sub

Sub BtnPdf(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PdfSave()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsOrdersUi.UiInfo("Сохранено: " & Mid(s, 4))
End Sub

Sub BtnBack(Optional oEvent As Variant)
    WmsInit()
    If gPrFrom <> "" Then WmsView.GoCell(gPrFrom, 1, 1) Else WmsUi.NavMain()
End Sub

' the document of the page as a PDF in «WMS_Документы» next to the book (a name of the document and the date; «(2)» … when the
' file exists). OK:<path> / ERR:…
Function PdfSave() As String
    Dim sh As Object, sFolder As String, sBase As String, url As String, k As Integer, fd(0) As New com.sun.star.beans.PropertyValue
    Dim pv(1) As New com.sun.star.beans.PropertyValue, ra As Variant, sel As Object
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        PdfSave = "ERR:книга не сохранена в файл — PDF некуда положить"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_PRINT)
    ra = sh.getPrintAreas()
    If UBound(ra) < 0 Then
        PdfSave = "ERR:на странице «" & SH_PRINT & "» ещё нет документа — сделайте его кнопкой на рабочем листе"
        Exit Function
    End If
    sFolder = gBaseDir & DOCS_DIR & "/"
    If Not gSfa.exists(sFolder) Then gSfa.createFolder(sFolder)
    sBase = FileSafe(IIf(gPrTitle <> "", gPrTitle, "документ")) & " " & Format(Now(), "YYYY-MM-DD")
    url = sFolder & sBase & ".pdf"
    k = 1
    Do While gSfa.exists(url)
        k = k + 1
        url = sFolder & sBase & " (" & k & ").pdf"
    Loop
    sel = sh.getCellRangeByPosition(ra(0).StartColumn, ra(0).StartRow, ra(0).EndColumn, ra(0).EndRow)
    fd(0).Name = "Selection"
    fd(0).Value = sel
    pv(0).Name = "FilterName"
    pv(0).Value = "calc_pdf_Export"
    pv(1).Name = "FilterData"
    pv(1).Value = fd()
    gDoc.storeToURL(url, pv())
    gPrLastPdf = ConvertFromURL(url)
    PdfSave = "OK:" & gPrLastPdf
    gUiLastMsg = PdfSave
    Exit Function
EH:
    PdfSave = "ERR:PDF не сохранён: " & Error$
    gUiLastMsg = PdfSave
End Function

' a file name of the text: the characters a file name cannot hold become «_»
Private Function FileSafe(s As String) As String
    Dim i As Integer, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If InStr("\/:*?""<>|«»№", ch) > 0 Then
            r = r & "_"
        Else
            r = r & ch
        End If
    Next i
    FileSafe = Trim(Left(r, 80))
End Function

' a quantity of EI n with its parts when it has a packing («0,25 шт (1 кан)»)
Function PrQty(n As Long, q As Double, sUnit As String) As String
    If n > 0 Then PrQty = WmsPack.QtyWithParts(n, q, sUnit) Else PrQty = WmsIssue.QtyText(q) & IIf(sUnit <> "", " " & sUnit, "")
End Function

' ================================================================ З12 the act of discrepancy with the supplier

Sub BtnDiscrepancyAct(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = DiscrepancyAct()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Акт расхождения»: the delivery of the row under the cursor — every received row of «Заказы» with the same supplier (L) and
' document (C): by the document (G) and in fact (F), the shortage or the surplus; the returns to the supplier of those EIs
' (brak, wrong goods). A document for the claim, on «Печать». OK:… / ERR:…
Function DiscrepancyAct() As String
    Dim osh As Object, r As Long, sup As String, docNo As String, last As Long, d As Variant, i As Long, rows() As Variant, n As Long, k As Long
    Dim f As Double, g As Double, hasG As Boolean, diff As Double, nDiff As Long, sh As Object, eis As String, rr() As Variant, nr As Long
    Dim ddate As String, ordNo As String, rdate As String, ash As Object, lastA As Long, a As Variant, j As Long, msg As String, en As Long
    osh = WmsOrders.OrdersSheet()
    r = WmsView.ViewRow(SH_ORDERS)
    If r < 1 Then
        DiscrepancyAct = "ERR:встаньте на строку прихода на листе «" & SH_ORDERS & "» и нажмите «Акт расхождения»"
        gUiLastMsg = DiscrepancyAct
        Exit Function
    End If
    sup = Trim(osh.getCellByPosition(OC_SUPPLIER, r).getString())
    docNo = Trim(osh.getCellByPosition(OC_DOC, r).getString())
    If osh.getCellByPosition(OC_EI, r).getString() = "" Then
        DiscrepancyAct = "ERR:строка ещё не принята (нет ЕИ) — акт составляется по проведённому приходу"
        gUiLastMsg = DiscrepancyAct
        Exit Function
    End If
    last = WmsOrders.LastRow(osh)
    d = osh.getCellRangeByPosition(0, 1, OC_EI, last).getDataArray()
    ReDim rows(last)
    For i = 0 To last - 1
        If CStr(d(i)(OC_EI)) <> "" And SameText(CStr(d(i)(OC_SUPPLIER)), sup) And SameText(CStr(d(i)(OC_DOC)), docNo) Then
            If (docNo <> "" Or i + 1 = r) Then
                f = 0
                g = 0
                hasG = False
                If Not WmsOrders.ParseQtyValueAs(d(i)(OC_FACT), "F", f, msg) Then f = 0
                If CStr(d(i)(OC_DOCQTY)) <> "" Then hasG = WmsOrders.ParseQtyValueAs(d(i)(OC_DOCQTY), "G", g, msg)
                diff = 0
                If hasG Then diff = WmsIssue.Round3(f - g)
                If Abs(diff) > 0.0000005 Then nDiff = nDiff + 1
                n = n + 1
                rows(n - 1) = Array(n, CStr(d(i)(OC_NAME)) & IIf(CStr(d(i)(OC_ART)) <> "", " (" & d(i)(OC_ART) & ")", ""), CStr(d(i)(OC_EI)), _
                    IIf(hasG, WmsIssue.QtyText(g), "—"), WmsIssue.QtyText(f), IIf(Not hasG, "нет документа", IIf(diff < 0, "недостача " & WmsIssue.QtyText(-diff), _
                    IIf(diff > 0, "излишек " & WmsIssue.QtyText(diff), "совпадает"))), CStr(d(i)(OC_UNIT)))
                eis = eis & "|" & CStr(d(i)(OC_EI)) & "|"
                If ddate = "" Then ddate = DayText(d(i)(OC_DDATE))
                If rdate = "" Then rdate = DayText(d(i)(OC_RDATE))
                If ordNo = "" Then ordNo = CStr(d(i)(OC_ORDER))
            End If
        End If
    Next i
    ' the returns to the supplier of those EIs (live rows «Возврат поставщику»)
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    lastA = WmsOrders.LastRow(ash)
    ReDim rr(IIf(lastA > 0, lastA, 0))
    If lastA >= 1 Then
        a = ash.getCellRangeByPosition(0, 1, AC_CTL, lastA).getDataArray()
        For j = 0 To lastA - 1
            If CStr(a(j)(AC_KIND)) = "Возврат поставщику" And Left(CStr(a(j)(AC_CTL)), Len(ST_POSTED)) = ST_POSTED And InStr(eis, "|" & CStr(a(j)(AC_EI)) & "|") > 0 Then
                en = WmsView.EiNo(CStr(a(j)(AC_EI)), CLng(SysNum(SK_NEXT_EI)))
                nr = nr + 1
                rr(nr - 1) = Array(nr, CStr(a(j)(AC_NAME)), CStr(a(j)(AC_EI)), PrQty(en, CDbl(a(j)(AC_QTY)), CStr(a(j)(AC_UNIT))), DayText(a(j)(AC_DATE)), _
                    CStr(a(j)(AC_REASON)), IIf(CStr(a(j)(AC_PLACE)) <> "", "док. " & a(j)(AC_PLACE), ""))
            End If
        Next j
    End If
    sh = PrintBegin("АКТ О РАСХОЖДЕНИИ ПРИ ПРИЁМКЕ ТОВАРА", "Акт расхождения " & IIf(sup <> "", sup & " ", "") & IIf(docNo <> "", docNo, ""))
    PrintLine(sh, "Дата составления: " & Format(Date(), "DD.MM.YYYY"))
    PrintLine(sh, "Поставщик: " & IIf(sup <> "", sup, "—"))
    PrintLine(sh, "Документ поставщика: " & IIf(docNo <> "", "№ " & docNo, "—") & IIf(ddate <> "", " от " & ddate, "") & IIf(ordNo <> "", " · заказ " & ordNo, "") _
        & IIf(rdate <> "", " · принято " & rdate, ""))
    PrintLine(sh, "Позиций в поставке: " & n & IIf(nDiff > 0, " · с расхождением по количеству: " & nDiff, " · по количеству расхождений нет") _
        & IIf(nr > 0, " · возвращено как брак / пересорт: " & nr, ""), True)
    PrintTable(sh, Array("№", "Наименование (артикул)", "ЕИ", "По документу", "Фактически", "Расхождение", "Ед."), rows, n, "0,3,4")
    If nr > 0 Then
        PrintLine(sh, "Брак, пересорт — возвращено поставщику:", True)
        PrintTable(sh, Array("№", "Наименование", "ЕИ", "Количество", "Дата", "Причина", "Документ"), rr, nr, "0,3")
    End If
    PrintLine(sh, "Претензия: прошу допоставить недостающий товар, заменить бракованный или уменьшить сумму к оплате.")
    PrintSign(sh, Array("Представитель склада", "Представитель поставщика (водитель)"))
    DiscrepancyAct = PrintDone(sh)
End Function

Private Function SameText(a As String, b As String) As Boolean
    SameText = (UCase(Trim(a)) = UCase(Trim(b)))
End Function

' a date of a cell as «19.10.2026» ("" when it is not a date)
Function DayText(v As Variant) As String
    If VarType(v) = 5 Or VarType(v) = 7 Then
        If v > 1 Then DayText = Format(CDate(Int(v)), "DD.MM.YYYY")
    ElseIf VarType(v) = 8 Then
        DayText = Trim(CStr(v))
    End If
End Function
