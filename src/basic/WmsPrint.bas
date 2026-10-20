' WmsPrint — 1.0: документы на бумаге. Одна страница «Печать», её заполняет документ — акт расхождения с поставщиком (З12),
' акт списания (К3), опись и сличительная ведомость (К4), накладная выдачи с подписями (К11, К12), карточка ЕИ (К14), отчёты;
' «Печать…» — на принтер, «Сохранить PDF» — файл в папке «WMS_Документы» рядом с книгой, без интернета. Страница — вид:
' ничего не проводит, в журнал не пишет. Library WmsPlus.
Option Explicit

Public Const SH_PRINT = "Печать"
Public Const PR_COLS = 7
Public Const PR_MAX = 2000
Public Const DOCS_DIR = "WMS_Документы"
Public Const RC_MAX = 500

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
Function FileSafe(s As String) As String
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

' ================================================================ helpers of the documents

' the rows (0-based, header excluded) selected on sheet sName — every range of the selection; Array() when another sheet
Function SelRows(sName As String) As Variant
    Dim sel As Object, addrs As Variant, i As Long, r As Long, out() As Long, n As Long
    SelRows = Array()
    On Error GoTo EH
    If gDoc.getCurrentController().getActiveSheet().getName() <> sName Then Exit Function
    sel = gDoc.getCurrentSelection()
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then addrs = sel.getRangeAddresses() Else addrs = Array(sel.getRangeAddress())
    ReDim out(100)
    For i = 0 To UBound(addrs)
        For r = IIf(addrs(i).StartRow < 1, 1, addrs(i).StartRow) To addrs(i).EndRow
            If r > WmsOrders.LastRow(gDoc.Sheets.getByName(sName)) Then Exit For
            If n > UBound(out) Then
                ReDim Preserve out(2 * n)
            End If
            out(n) = r
            n = n + 1
        Next r
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve out(n - 1)
    SelRows = out
EH:
End Function

' the price of EI n (the price J of its receipt; a split EI — that of its source); -1 when there is none
Function EiPrice(ByVal n As Long) As Double
    Dim rv As Variant, r As Long, canon As String, k As Integer, m As Long, c As Object
    EiPrice = -1
    On Error GoTo EH
    For k = 0 To 5
        canon = WmsIssue.EiCanon(n)
        If WmsOrders.RcvRegistered(n, rv) Then
            r = WmsOrders.ReceiptRowOf(canon, rv(RV_ROW))
            If r >= 1 Then
                c = WmsOrders.OrdersSheet().getCellByPosition(OC_PRICE, r)
                If c.getType() = com.sun.star.table.CellContentType.VALUE Then EiPrice = c.getValue()
            End If
            Exit Function
        End If
        m = WmsSupply.SplitSourceOf(canon)
        If m < 1 Then
            ' 1.0 (Ф8): a kit — the money of its components ÷ the number of kits
            EiPrice = WmsKit.KitPrice(canon)
            Exit Function
        End If
        n = m
    Next k
EH:
End Function

' the money of q at price p as text («125,50»); "" without a price
Function MoneyText(q As Double, p As Double) As String
    If p < 0 Then Exit Function
    MoneyText = Money(q * p)
End Function

' money for people, whatever the language of the PC: «1 250,50», «−3,00»
Function Money(x As Double) As String
    Dim m As Double, ip As Double, fr As Long, s As String, g As String
    m = Int(Abs(x) * 100 + 0.5)
    ip = Int(m / 100)
    fr = CLng(m - ip * 100)
    s = Format(ip, "0")
    Do While Len(s) > 3
        g = " " & Right(s, 3) & g
        s = Left(s, Len(s) - 3)
    Loop
    Money = IIf(x < 0 And m > 0, "−", "") & s & g & "," & Right("0" & fr, 2)
End Function

' money for a file (a CSV of the accounting): «1250,50» — no spaces, a comma
Function MoneyCsv(x As Double) As String
    Dim m As Double, ip As Double, fr As Long
    m = Int(Abs(x) * 100 + 0.5)
    ip = Int(m / 100)
    fr = CLng(m - ip * 100)
    MoneyCsv = IIf(x < 0 And m > 0, "-", "") & Format(ip, "0") & "," & Right("0" & fr, 2)
End Function

' a quantity in the parts first when the EI has a packing («1 кан (0,25 шт)»), otherwise «2 шт» (К12)
Function PartsFirst(n As Long, q As Double, sUnit As String) As String
    Dim pu As String, pn As Double
    If n > 0 Then
        If WmsPack.PackOf(n, pu, pn) Then
            PartsFirst = WmsIssue.QtyText(WmsPack.PkRound(q * pn, 3)) & " " & pu & " (" & WmsIssue.QtyText(q) & IIf(sUnit <> "", " " & sUnit, "") & ")"
            Exit Function
        End If
    End If
    PartsFirst = WmsIssue.QtyText(q) & IIf(sUnit <> "", " " & sUnit, "")
End Function

' ================================================================ К3 the act of a write-off

Sub BtnWriteOffAct(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = WriteOffAct()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Акт списания»: the posted rows «Списание» of the selection on «Корректировки» (one row — that one): the number, the date, the
' reasons, the positions in their parts, the price and the sum, the signatures. OK:… / ERR:…
Function WriteOffAct() As String
    Dim ash As Object, sel As Variant, i As Long, r As Long, d As Variant, rows() As Variant, n As Long, en As Long, p As Double, sum As Double
    Dim no0 As String, no1 As String, dt As String, reasons As String, sh As Object, q As Double, noPrice As Long
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    sel = SelRows(SH_ADJUST)
    If UBound(sel) < 0 Then
        WriteOffAct = "ERR:выделите проведённые строки «Списание» на листе «" & SH_ADJUST & "» и нажмите «Акт списания»"
        gUiLastMsg = WriteOffAct
        Exit Function
    End If
    ReDim rows(UBound(sel))
    For i = 0 To UBound(sel)
        r = sel(i)
        d = ash.getCellRangeByPosition(0, r, AC_LAST, r).getDataArray()(0)
        If CStr(d(AC_KIND)) = "Списание" And Left(CStr(d(AC_CTL)), Len(ST_POSTED)) = ST_POSTED And VarType(d(AC_QTY)) = 5 Then
            en = WmsView.EiNo(CStr(d(AC_EI)), CLng(SysNum(SK_NEXT_EI)))
            q = d(AC_QTY)
            p = EiPrice(en)
            If p < 0 Then noPrice = noPrice + 1 Else sum = sum + q * p
            n = n + 1
            rows(n - 1) = Array(n, CStr(d(AC_EI)), CStr(d(AC_NAME)) & IIf(CStr(d(AC_ART)) <> "", " (" & d(AC_ART) & ")", ""), PartsFirst(en, q, CStr(d(AC_UNIT))), _
                CStr(d(AC_UNIT)), IIf(p >= 0, Money(p), "—"), IIf(p >= 0, MoneyText(q, p), "—"))
            If no0 = "" Then no0 = CStr(d(AC_NO))
            no1 = CStr(d(AC_NO))
            If dt = "" Then dt = DayText(d(AC_DATE))
            If InStr("|" & reasons & "|", "|" & CStr(d(AC_REASON)) & "|") = 0 Then reasons = reasons & IIf(reasons <> "", "|", "") & CStr(d(AC_REASON))
        End If
    Next i
    If n = 0 Then
        WriteOffAct = "ERR:среди выделенных строк нет проведённых списаний — акт составляется по проведённым строкам вида «Списание»"
        gUiLastMsg = WriteOffAct
        Exit Function
    End If
    sh = PrintBegin("АКТ СПИСАНИЯ № " & no0 & IIf(no1 <> no0, "–" & no1, "") & " от " & dt, "Акт списания " & no0)
    PrintLine(sh, "Основание (причина): " & Replace(reasons, "|", "; "))
    PrintLine(sh, "Списано позиций: " & n & " · на сумму " & Money(sum) & IIf(noPrice > 0, " (без цены: " & noPrice & ")", ""), True)
    PrintTable(sh, Array("№", "ЕИ", "Наименование (артикул)", "Количество", "Ед.", "Цена", "Сумма"), rows, n, "0,5,6")
    PrintLine(sh, "Итого на сумму: " & Money(sum), True)
    PrintSign(sh, Array("Списал (кладовщик)", "Проверил", "Утвердил"))
    WriteOffAct = PrintDone(sh)
End Function

' ================================================================ К11, К12 the delivery note with the signatures

Sub BtnIssueNote(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = IssueNote()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Накладная»: the posted issues of the selection on «Выдачи»; one row — that row or «всё, что выдано <кому> <дата>». The
' quantities in the parts («1 кан (0,25 шт)»); the number of the note goes into B «№ документа выдачи» of the issues that had
' none (a field of the storekeeper). OK:… / ERR:…
Function IssueNote() As String
    Dim ish As Object, sel As Variant, rs() As Long, nR As Long, i As Long, r As Long, d As Variant, who As String, dt As Double, k As Integer
    Dim rows() As Variant, n As Long, en As Long, docNo As String, sh As Object, last As Long, a As Variant, j As Long, whoKey As String, res As String
    ish = WmsIssue.IssuesSheet()
    sel = SelRows(SH_ISSUES)
    If UBound(sel) < 0 Then
        IssueNote = "ERR:выделите проведённые выдачи на листе «" & SH_ISSUES & "» (или встаньте на одну) и нажмите «Накладная»"
        gUiLastMsg = IssueNote
        Exit Function
    End If
    ReDim rs(UBound(sel))
    For i = 0 To UBound(sel)
        rs(i) = sel(i)
    Next i
    nR = UBound(sel) + 1
    If nR = 1 Then
        d = ish.getCellRangeByPosition(0, rs(0), IC_LAST, rs(0)).getDataArray()(0)
        who = Trim(CStr(d(IC_WHO)))
        If VarType(d(IC_DATE)) = 5 Then dt = Int(d(IC_DATE))
        If who <> "" And dt > 0 Then
            k = WmsPrimeUi.UiButtons("Накладная", "Накладная на выдачу № " & d(IC_NO) & " — " & who & ", " & Format(dt, "DD.MM.YYYY") & ".", _
                Array("Только эта выдача", "Всё, что выдано «" & who & "» " & Format(dt, "DD.MM.YYYY"), "Отмена"), Array(1, 2, 0))
            If k = 0 Or k = -1 Then
                IssueNote = "SKIP:окно закрыто"
                gUiLastMsg = IssueNote
                Exit Function
            End If
            If k = 2 Then
                ' every posted issue of that recipient and that day
                whoKey = WmsAsk.WhoKey(who)
                last = WmsOrders.LastRow(ish)
                a = ish.getCellRangeByPosition(0, 1, IC_LAST, last).getDataArray()
                ReDim rs(last)
                nR = 0
                For j = 0 To last - 1
                    If VarType(a(j)(IC_DATE)) = 5 Then
                        If Int(a(j)(IC_DATE)) = dt And WmsAsk.WhoKey(CStr(a(j)(IC_WHO))) = whoKey And Left(CStr(a(j)(IC_CTL)), Len(ST_POSTED)) = ST_POSTED Then
                            rs(nR) = j + 1
                            nR = nR + 1
                        End If
                    End If
                Next j
            End If
        End If
    End If
    ReDim rows(nR)
    For i = 0 To nR - 1
        r = rs(i)
        d = ish.getCellRangeByPosition(0, r, IC_LAST, r).getDataArray()(0)
        If Left(CStr(d(IC_CTL)), Len(ST_POSTED)) = ST_POSTED And VarType(d(IC_QTY)) = 5 Then
            en = WmsView.EiNo(CStr(d(IC_EI)), CLng(SysNum(SK_NEXT_EI)))
            n = n + 1
            rows(n - 1) = Array(n, CStr(d(IC_EI)), CStr(d(IC_NAME)) & IIf(CStr(d(IC_ART)) <> "", " (" & d(IC_ART) & ")", ""), _
                PartsFirst(en, CDbl(d(IC_QTY)), CStr(d(IC_UNIT))), CStr(d(IC_WHO)), DayText(d(IC_DATE)), "№ " & d(IC_NO))
            If docNo = "" And Trim(CStr(d(IC_DOC))) <> "" Then docNo = Trim(CStr(d(IC_DOC)))
            If who = "" Then who = Trim(CStr(d(IC_WHO)))
            If dt = 0 And VarType(d(IC_DATE)) = 5 Then dt = Int(d(IC_DATE))
            rs(n - 1) = r
        End If
    Next i
    If n = 0 Then
        IssueNote = "ERR:среди выделенных строк нет проведённых выдач"
        gUiLastMsg = IssueNote
        Exit Function
    End If
    ' the number of the note: the document of the issues, or «Н-<№ первой выдачи>» written into B where it was empty (К12)
    If docNo = "" Then docNo = "Н-" & ish.getCellByPosition(IC_NO, rs(0)).getString()
    res = NoteNumber(ish, rs, n, docNo)
    If Left(res, 4) = "ERR:" Or Left(res, 4) = "ERR-" Then
        IssueNote = "ERR:номер накладной не записан в выдачи — " & Mid(res, InStr(res, ":") + 1)
        gUiLastMsg = IssueNote
        Exit Function
    End If
    sh = PrintBegin("НАКЛАДНАЯ НА ВЫДАЧУ № " & docNo & " от " & Format(IIf(dt > 0, dt, WmsOrders.Today()), "DD.MM.YYYY"), "Накладная " & docNo)
    PrintLine(sh, "Получатель: " & IIf(who <> "", who, "—"))
    PrintLine(sh, "Позиций: " & n & ". Количество — как выдано: в частях фасовки, в скобках — в единицах учёта.", True)
    PrintTable(sh, Array("№", "ЕИ", "Наименование (артикул)", "Количество", "Кому", "Дата", "Выдача"), rows, n, "0")
    PrintSign(sh, Array("Выдал (кладовщик)", "Получил"))
    IssueNote = PrintDone(sh)
End Function

' К12: the number of the note into B of the issues that had none — one operation ISSUE_DOC (journal, recovery); "" when nothing
' to write, otherwise the result of ApplyOperation
Private Function NoteNumber(ish As Object, rs() As Long, n As Long, docNo As String) As String
    Dim i As Long, k As Long, nos As String
    For i = 0 To n - 1
        If Trim(ish.getCellByPosition(IC_DOC, rs(i)).getString()) = "" Then
            If k = 0 Then PlanBegin("ISSUE_DOC", docNo)
            nos = nos & IIf(k > 0, ",", "") & ish.getCellByPosition(IC_NO, rs(i)).getString()
            PlanSetValue(SH_ISSUES, rs(i), IC_DOC, docNo, True)
            k = k + 1
        End If
    Next i
    If k = 0 Then Exit Function
    PlanField("DOC", docNo)
    PlanField("NOS", nos)
    NoteNumber = ApplyOperation(0, 0)
End Function

' ================================================================ К14 the card of an EI on paper

Sub BtnCardPrint(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = CardPrint(WmsFind.FindCardEI())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Печать карточки»: the card of EI canon («Поиск» builds it) — its movements with the running balance — and every correction and
' storno of the journal with «было → стало», the time of the posting and the reason. OK:… / ERR:…
Function CardPrint(canon As String) As String
    Dim n As Long, msg As String, fd As Object, d As Variant, i As Long, rows() As Variant, nr As Long, sh As Object, head As String, info As String
    Dim files As Variant, f As Integer, txt As String, a As Variant, j As Long, seq As Long, ts As String, inst As String, sType As String, fields As Variant
    Dim fx() As Variant, nf As Long, res As String, sIn As String
    sIn = canon
    If Not WmsIssue.NormalizeEI(sIn, n, canon, msg) Then
        CardPrint = "ERR:откройте карточку ЕИ на странице «Поиск» (впишите номер ЕИ) и нажмите «Печать карточки»"
        gUiLastMsg = CardPrint
        Exit Function
    End If
    res = WmsFind.CardShow(n)
    If Left(res, 4) = "ERR:" Then
        CardPrint = res
        Exit Function
    End If
    fd = gDoc.Sheets.getByName(SH_FIND)
    head = fd.getCellByPosition(1, FD_HEAD).getString()
    info = fd.getCellByPosition(1, FD_INFO).getString()
    d = fd.getCellRangeByPosition(1, FD_TABLE, 8, FD_TABLE + 400).getDataArray()
    ReDim rows(400)
    For i = 0 To UBound(d)
        If CStr(d(i)(0)) = "" And CStr(d(i)(1)) = "" Then Exit For
        rows(nr) = Array(d(i)(0), d(i)(1), d(i)(2), d(i)(4), d(i)(5), d(i)(6), d(i)(7))
        nr = nr + 1
    Next i
    ' the corrections and the storno of the EI in the journal
    ReDim fx(200)
    files = WmsJournal.ListFiles()
    gAuHeadSheets = "|"
    gAuHead = New Collection
    gAuHeadN = 0
    ReDim gAuHeadTxt(300)
    For f = 0 To UBound(files)
        txt = WmsAudit.ReadText(gJDir & files(f))
        a = Split(Replace(txt, Chr(13), ""), Chr(10))
        For j = 0 To UBound(a)
            If InStr(a(j), canon) > 0 Then
                If WmsJournal.ParseLine(a(j), seq, ts, inst, sType, fields) Then
                    If (Right(sType, 4) = "_FIX" Or Right(sType, 4) = "_DEL") And WmsView.JField(fields, "EI") = canon And nf <= 200 Then
                        fx(nf) = Array(nf + 1, Mid(ts, 9, 2) & "." & Mid(ts, 6, 2) & "." & Left(ts, 4) & " " & Mid(ts, 12, 5), WmsView.OpName(sType), _
                            WmsAudit.Changes(fields) & IIf(WmsView.JField(fields, "WHY") <> "", " · причина: " & WmsView.JField(fields, "WHY"), ""))
                        nf = nf + 1
                    End If
                End If
            End If
        Next j
    Next f
    sh = PrintBegin("КАРТОЧКА " & canon, "Карточка " & canon)
    PrintLine(sh, head, True)
    PrintLine(sh, info)
    PrintTable(sh, Array("Дата", "Операция", "Документ / кому", "Приход", "Расход", "Остаток", "Подробно"), rows, nr, "3,4,5")
    If nf > 0 Then
        PrintLine(sh, "Исправления и сторно (из журнала WMS): " & nf, True)
        PrintTable(sh, Array("№", "Проведено", "Операция", "Было → стало, причина"), fx, nf, "0")
    Else
        PrintLine(sh, "Исправлений и сторно не было.")
    End If
    PrintSign(sh, Array("Кладовщик"))
    CardPrint = PrintDone(sh)
End Function

' ================================================================ К4 the inventory list and the comparison sheet

Sub BtnInventoryList(Optional oEvent As Variant)
    Dim s As String, v(0) As String
    WmsInit()
    If Not WmsPrimeUi.UiForm("Опись для пересчёта", "Все ЕИ с остатком по порядку мест, колонка «Факт» пустая — для записи от руки. Ряд «5», секция «5-12», " _
        & "место «5-12-3» или пусто — весь склад.", Array("Место или зона"), v, Array(""), "Сделать опись") Then Exit Sub
    s = InventoryList(Trim(v(0)))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the EIs with a balance on the places beginning with sZone ("" — all) by the order of the places: the book balance in the parts,
' an empty column «Факт». OK:… / ERR:…
Function InventoryList(sZone As String) As String
    Dim st As Object, last As Long, r0 As Long, r1 As Long, d As Variant, i As Long, rows() As Variant, key() As Double, ord() As Long, n As Long
    Dim pls As Variant, zp As Variant, sh As Object, tmp() As Variant, k As Long, pl As String, en As Long
    zp = Array()
    If sZone <> "" Then zp = ZoneParts(sZone)
    If sZone <> "" And UBound(zp) < 0 Then
        InventoryList = "ERR:«" & sZone & "» — не ряд, секция или место (например 5, 5-12, 5-12-3)"
        gUiLastMsg = InventoryList
        Exit Function
    End If
    st = WmsIssue.StockSheet()
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    ReDim rows(IIf(last > 0, last, 0))
    ReDim key(IIf(last > 0, last, 0))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = st.getCellRangeByPosition(0, r0, SC_STATE, r1).getDataArray()
        For i = 0 To r1 - r0
            If VarType(d(i)(SC_QTY)) = 5 Then
                If d(i)(SC_QTY) > 0.0000005 Then
                    pl = Trim(CStr(d(i)(SC_PLACE)))
                    pls = WmsView.PlaceList(pl)
                    If InZone(pls, zp) Then
                        en = r0 + i
                        rows(n) = Array(0, IIf(pl <> "", pl, "без места"), CStr(d(i)(SC_EI)), CStr(d(i)(SC_NAME)), PartsFirst(en, CDbl(d(i)(SC_QTY)), CStr(d(i)(SC_UNIT))), "", "")
                        key(n) = PlaceOrder(pls) * 1000000 + en
                        n = n + 1
                    End If
                End If
            End If
        Next i
    Next r0
    ReDim ord(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        ord(i) = i
    Next i
    WmsView.SortIdx(key, ord, n, False)
    ReDim tmp(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        tmp(i) = rows(ord(i))
        tmp(i) = Array(i + 1, tmp(i)(1), tmp(i)(2), tmp(i)(3), tmp(i)(4), "", "")
    Next i
    If n = 0 Then
        InventoryList = "ERR:в «" & IIf(sZone <> "", sZone, "складе") & "» нет ЕИ с остатком — проверьте написание: ряд «5», секция «5-12», место «5-12-3»"
        gUiLastMsg = InventoryList
        Exit Function
    End If
    sh = PrintBegin("ИНВЕНТАРИЗАЦИОННАЯ ОПИСЬ от " & Format(WmsOrders.Today(), "DD.MM.YYYY"), "Опись " & IIf(sZone <> "", sZone, "склад"))
    PrintLine(sh, "Зона: " & IIf(sZone <> "", "«" & sZone & "»", "весь склад") & " · ЕИ с остатком: " & n & ". Впишите в «Факт» пересчитанное количество " _
        & "(в тех же частях); найденное без этикетки — в конце листа.", True)
    PrintTable(sh, Array("№", "Место", "ЕИ", "Наименование", "Учёт", "Факт", "Примечание"), tmp, n, "0")
    PrintSign(sh, Array("Пересчитал", "Проверил"))
    InventoryList = PrintDone(sh)
End Function

' ================================================================ К5 recount a place or a zone

Sub BtnRecountZone(Optional oEvent As Variant)
    Dim s As String, v(0) As String
    WmsInit()
    If Not WmsPrimeUi.UiForm("Пересчитать место или зону", "WMS добавит в конец листа строки «Инвентаризация» для всех ЕИ с остатком в этом месте: " _
        & "учётный остаток, дата — сегодня, основание. Останется вписать фактический остаток (H) и провести; совпавшее запишется с разницей 0 — " _
        & "«пересчитано». Ряд «5», секция «5-12» или место «5-12-3».", Array("Место или зона"), v, Array(""), "Заполнить строки") Then Exit Sub
    s = RecountZone(Trim(v(0)))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsOrdersUi.UiInfo(Mid(s, 4))
End Sub

' the rows «Инвентаризация» of every EI with a balance in zone sZone at the end of «Корректировки» (B C I K L R: the kind, the EI, the
' book balance, today, the reason, the batch «Пересчёт <зона> <дата>»), the place order; an EI that already has an unposted row
' «Инвентаризация» is not added twice. The cells are the storekeeper's inputs (nothing posted); one Undo removes them. OK:… / ERR:…
Function RecountZone(sZone As String) As String
    Dim zp As Variant, st As Object, ash As Object, last As Long, r0 As Long, r1 As Long, d As Variant, i As Long, n As Long, pls As Variant, pl As String
    Dim eis() As Long, key() As Double, ord() As Long, bk() As Double, sWait As String, a As Variant, j As Long, r As Long, um As Object, inCtx As Boolean
    Dim today As Double, sBatch As String, nSkip As Long, lastA As Long, x As Long
    If sZone = "" Then
        RecountZone = "ERR:укажите место или зону — ряд «5», секцию «5-12» или место «5-12-3»"
        gUiLastMsg = RecountZone
        Exit Function
    End If
    zp = ZoneParts(sZone)
    If UBound(zp) < 0 Then
        RecountZone = "ERR:«" & sZone & "» — не ряд, секция или место (например 5, 5-12, 5-12-3)"
        gUiLastMsg = RecountZone
        Exit Function
    End If
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    lastA = WmsOrders.LastRow(ash)
    ' the EIs that already wait for their count (an unposted row «Инвентаризация»)
    sWait = "|"
    If lastA >= 1 Then
        a = ash.getCellRangeByPosition(0, 1, AC_CTL, lastA).getDataArray()
        For j = 0 To lastA - 1
            If CStr(a(j)(AC_NO)) = "" And CStr(a(j)(AC_KIND)) = "Инвентаризация" And CStr(a(j)(AC_EI)) <> "" Then sWait = sWait & UCase(Trim(CStr(a(j)(AC_EI)))) & "|"
        Next j
    End If
    st = WmsIssue.StockSheet()
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    ReDim eis(IIf(last > 0, last, 0))
    ReDim key(IIf(last > 0, last, 0))
    ReDim bk(IIf(last > 0, last, 0))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = st.getCellRangeByPosition(0, r0, SC_STATE, r1).getDataArray()
        For i = 0 To r1 - r0
            If VarType(d(i)(SC_QTY)) = 5 Then
                If d(i)(SC_QTY) > 0.0000005 Then
                    pl = Trim(CStr(d(i)(SC_PLACE)))
                    pls = WmsView.PlaceList(pl)
                    If UBound(pls) >= 0 Then
                        If InZone(pls, zp) Then
                            If InStr(sWait, "|" & UCase(CStr(d(i)(SC_EI))) & "|") > 0 Then
                                nSkip = nSkip + 1
                            ElseIf n < RC_MAX Then
                                eis(n) = r0 + i
                                bk(n) = d(i)(SC_QTY)
                                key(n) = PlaceOrder(pls) * 1000000 + r0 + i
                                n = n + 1
                            Else
                                x = x + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    Next r0
    If n = 0 Then
        RecountZone = "ERR:в «" & sZone & "» нет ЕИ с остатком" & IIf(nSkip > 0, ", кроме уже ждущих пересчёта (" & nSkip & ")", "")
        gUiLastMsg = RecountZone
        Exit Function
    End If
    ReDim ord(n - 1)
    For i = 0 To n - 1
        ord(i) = i
    Next i
    WmsView.SortIdx(key, ord, n, False)
    today = WmsOrders.Today()
    sBatch = "Пересчёт " & sZone & " " & Format(today, "DD.MM.YYYY")
    r = lastA + 1
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: пересчитать " & sZone)
    inCtx = True
    WmsAdjust.HandlerOff(True)
    For i = 0 To n - 1
        ash.getCellRangeByPosition(AC_KIND, r + i, AC_EI, r + i).setDataArray(Array(Array("Инвентаризация", st.getCellByPosition(SC_EI, eis(ord(i))).getString())))
        ash.getCellByPosition(AC_BOOK, r + i).setValue(bk(ord(i)))
        ash.getCellByPosition(AC_DATE, r + i).setValue(today)
        ash.getCellByPosition(AC_REASON, r + i).setString("пересчёт места " & sZone)
        ash.getCellByPosition(AC_BATCH, r + i).setString(sBatch)
    Next i
    For i = 0 To n - 1
        WmsAdjust.PreviewAdjustRow(r + i)
    Next i
    WmsAdjust.HandlerOff(False)
    um.leaveUndoContext()
    inCtx = False
    WmsView.GoCell(SH_ADJUST, r, AC_FACT)
    RecountZone = "OK:добавлено строк пересчёта: " & n & " (строки " & (r + 1) & "–" & (r + n) & ", пакет «" & sBatch & "»)" _
        & IIf(nSkip > 0, "; уже ждут пересчёта: " & nSkip, "") & IIf(x > 0, "; ещё " & x & " ЕИ не добавлены — не больше " & RC_MAX & " за раз", "") _
        & ". Впишите фактический остаток в H и проведите выделенные строки."
    gUiLastMsg = RecountZone
    Exit Function
EH:
    RecountZone = "ERR:строки пересчёта не добавлены полностью: " & Error$ & " — отмените (Ctrl+Z) и повторите"
    gUiLastMsg = RecountZone
    On Error Resume Next
    WmsAdjust.HandlerOff(False)
    If inCtx Then um.leaveUndoContext()
End Function

' the parts of a zone («5», «5-12», «5-12-3») as the keys have them; Array() when none
Function ZoneParts(ByVal s As String) As Variant
    Dim k As String, a As Variant
    ZoneParts = Array()
    s = Trim(s)
    If s = "" Then Exit Function
    k = WmsView.PlaceKey(s)
    If k = "" Then
        k = WmsView.PlaceKey(s & "-0")
        If k = "" Then Exit Function
        a = Split(k, "-")
        ZoneParts = Array(a(0))
        Exit Function
    End If
    ZoneParts = Split(k, "-")
End Function

' one of the places pls lies in the zone zp (no zone — every place, «без места» too)
Function InZone(pls As Variant, zp As Variant) As Boolean
    Dim j As Integer
    If UBound(zp) < 0 Then
        InZone = True
        Exit Function
    End If
    For j = 0 To UBound(pls)
        If WmsMap.KeyRenamed(CStr(pls(j)), zp, zp) <> "" Then
            InZone = True
            Exit Function
        End If
    Next j
End Function

' the order of the first place of a list: ряд, секция, ярус (letters by their code; no place — last)
Private Function PlaceOrder(pls As Variant) As Double
    Dim a As Variant, i As Integer, x As Double, w As Variant, v As Double
    PlaceOrder = 99999999
    If UBound(pls) < 0 Then Exit Function
    a = Split(pls(0), "-")
    w = Array(100000, 100, 1)
    For i = 0 To UBound(a)
        If i > 2 Then Exit For
        If IsNumeric(a(i)) Then v = Val(a(i)) Else v = 900 + (AscW(Left(a(i), 1)) Mod 99)
        If v > IIf(i = 2, 99, 999) Then v = IIf(i = 2, 99, 999)
        x = x + v * w(i)
    Next i
    PlaceOrder = x
End Function

Sub BtnInventoryCompare(Optional oEvent As Variant)
    Dim s As String, v(0) As String, d As Double, msg As String
    WmsInit()
    v(0) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    If Not WmsPrimeUi.UiForm("Сличительная ведомость", "Проведённые строки «Инвентаризация» за день: излишки и недостачи с суммами.", Array("Дата пересчёта"), v, _
        Array(""), "Сделать ведомость") Then Exit Sub
    If Not WmsOrders.ParseDateTextAs(Trim(v(0)), "дата", d, msg) Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = InventoryCompare(d)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the posted counts («Инвентаризация») of day dt: book, actual, the surplus or the shortage, the price and the sum. OK:… / ERR:…
Function InventoryCompare(dt As Double) As String
    Dim ash As Object, last As Long, d As Variant, i As Long, rows() As Variant, n As Long, en As Long, p As Double, diff As Double
    Dim plus As Double, minus As Double, nEq As Long, sh As Object
    ash = gDoc.Sheets.getByName(SH_ADJUST)
    last = WmsOrders.LastRow(ash)
    If last < 1 Then
        InventoryCompare = "ERR:на листе «" & SH_ADJUST & "» нет пересчётов"
        gUiLastMsg = InventoryCompare
        Exit Function
    End If
    d = ash.getCellRangeByPosition(0, 1, AC_LAST, last).getDataArray()
    ReDim rows(last)
    For i = 0 To last - 1
        If CStr(d(i)(AC_KIND)) = "Инвентаризация" And Left(CStr(d(i)(AC_CTL)), Len(ST_POSTED)) = ST_POSTED And VarType(d(i)(AC_DATE)) = 5 Then
            If Int(d(i)(AC_DATE)) = dt Then
                en = WmsView.EiNo(CStr(d(i)(AC_EI)), CLng(SysNum(SK_NEXT_EI)))
                diff = 0
                If VarType(d(i)(AC_DIFF)) = 5 Then diff = d(i)(AC_DIFF)
                p = EiPrice(en)
                If Abs(diff) < 0.0000005 Then
                    nEq = nEq + 1
                ElseIf diff > 0 Then
                    If p >= 0 Then plus = plus + diff * p
                Else
                    If p >= 0 Then minus = minus - diff * p
                End If
                n = n + 1
                rows(n - 1) = Array(n, CStr(d(i)(AC_EI)), CStr(d(i)(AC_NAME)), WmsIssue.QtyText(CDbl(d(i)(AC_BOOK))) & " / " & WmsIssue.QtyText(CDbl(d(i)(AC_FACT))), _
                    IIf(Abs(diff) < 0.0000005, "совпало", IIf(diff > 0, "излишек ", "недостача ") & WmsIssue.QtyText(Abs(diff)) & " " & d(i)(AC_UNIT)), _
                    IIf(p >= 0, Money(p), "—"), IIf(p >= 0 And Abs(diff) > 0.0000005, MoneyText(Abs(diff), p), ""))
            End If
        End If
    Next i
    If n = 0 Then
        InventoryCompare = "ERR:за " & Format(dt, "DD.MM.YYYY") & " проведённых пересчётов нет"
        gUiLastMsg = InventoryCompare
        Exit Function
    End If
    sh = PrintBegin("СЛИЧИТЕЛЬНАЯ ВЕДОМОСТЬ от " & Format(dt, "DD.MM.YYYY"), "Сличительная ведомость " & Format(dt, "YYYY-MM-DD"))
    PrintLine(sh, "Пересчитано ЕИ: " & n & " · совпало " & nEq & " · излишки на сумму " & Money(plus) & " · недостачи на сумму " & Money(minus), True)
    PrintTable(sh, Array("№", "ЕИ", "Наименование", "Учёт / факт", "Результат", "Цена", "Сумма"), rows, n, "0,5,6")
    PrintSign(sh, Array("Председатель комиссии", "Члены комиссии", "Материально ответственное лицо"))
    InventoryCompare = PrintDone(sh)
End Function
