' TbLabels — WMS_LABELS: этикетки ЕИ по снимку (задание «FINAL WMS MARATHON», §6). Лист «Этикетки»: список ЕИ и число
' копий; «Построить этикетки» — лист «Макет»: одна этикетка — одна страница размера этикетки (Настройки: ширина и
' высота, мм); на этикетке ЕИ крупно, наименование, артикул, место и штрихкод Code 128 (номер ЕИ, 8 цифр) — рисуется
' прямоугольниками, внешние шрифты не нужны. «Печать» — системный принтер LibreOffice / CUPS (обычное окно печати);
' «В PDF» — файл для проверки. Драйвер конкретного принтера ядро WMS не касается; адаптер прямого протокола (TSPL) —
' отдельный скрипт WMS_TOOLBOX/labels/tspl_labels.py.
Option Explicit

Private Const SH_LIST = "Этикетки"
Private Const SH_LAYOUT = "Макет"
Private Const ROWS_PER_LABEL = 5

' the Code 128 patterns (bar, space, bar, space, bar, space widths in modules); 105 Start C, 106 Stop (7 elements)
Private Function Patterns() As Variant
    Patterns = Split("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 113222 123122 " _
        & "123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 131123 131321 " _
        & "112313 132113 132311 211313 231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 " _
        & "311321 331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 " _
        & "142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 412121 111143 " _
        & "111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 211232 2331112", " ")
End Function

' the bar widths (modules) of Code 128 C for an even string of digits: start C, the pairs, the checksum, stop
Function Code128C(digits As String) As String
    Dim p As Variant, s As String, i As Integer, v As Integer, sum As Long, w As Integer
    p = Patterns()
    s = p(105)
    sum = 105
    w = 1
    For i = 1 To Len(digits) Step 2
        v = CInt(Mid(digits, i, 2))
        s = s & p(v)
        sum = sum + w * v
        w = w + 1
    Next i
    s = s & p(sum Mod 103) & p(106)
    Code128C = s
End Function

' «Загрузить снимок»: the registry of the latest snapshot (the EIs to print are listed on «Этикетки»: A ЕИ, B копий)
Sub BtnLblRefresh(Optional oEvent As Variant)
    Dim why As String, snap As String, n As Long
    snap = TbPickSnapshot(why)
    If snap = "" Then
        TbMsg("ERR:" & why)
        Exit Sub
    End If
    n = TbLoadTable(snap, "stock.csv", "_stock", True)
    TbPutSetting(3, snap)
    TbMsg("OK:ЕИ в снимке: " & n)
End Sub

Private Function Setting(r As Integer, dflt As Double) As Double
    Dim s As String
    s = Trim(Replace(TbSetting(r), ",", "."))
    If s = "" Then Setting = dflt Else Setting = Val(s)
End Function

Private Function Flag(r As Integer) As Boolean
    Flag = (LCase(Left(Trim(TbSetting(r)), 1)) <> "н")
End Function

Sub BtnLblBuild(Optional oEvent As Variant)
    TbMsg(LblBuild())
End Sub

' «Построить этикетки»: the layout sheet from the list; «ERR:» when an EI is not in the snapshot
Function LblBuild() As String
    Dim lst As Object, st As Object, lay As Object, last As Long, d As Variant, i As Long, k As Long, c As Long, n As Long, e As String, row As Variant
    Dim w As Double, h As Double, r As Long, bad As String, nBad As Long, dp As Object, ps As Object, stName As String, x As Long
    On Error GoTo EH
    If Not ThisComponent.Sheets.hasByName("_stock") Then
        LblBuild = "ERR:сначала «Загрузить снимок»"
        Exit Function
    End If
    lst = ThisComponent.Sheets.getByName(SH_LIST)
    st = ThisComponent.Sheets.getByName("_stock")
    last = TbLastRow(lst)
    If last < 1 Then
        LblBuild = "ERR:на листе «" & SH_LIST & "» нет ЕИ (колонка A, начиная со строки 2)"
        Exit Function
    End If
    d = lst.getCellRangeByPosition(0, 1, 1, last).getDataArray()
    w = Setting(5, 58)
    h = Setting(6, 40)
    If ThisComponent.Sheets.hasByName(SH_LAYOUT) Then ThisComponent.Sheets.removeByName(SH_LAYOUT)
    ThisComponent.Sheets.insertNewByName(SH_LAYOUT, ThisComponent.Sheets.getCount())
    lay = ThisComponent.Sheets.getByName(SH_LAYOUT)
    ' a page style of the label size: one label — one page
    stName = "WMS_Этикетка"
    If Not ThisComponent.StyleFamilies.getByName("PageStyles").hasByName(stName) Then
        ps = ThisComponent.createInstance("com.sun.star.style.PageStyle")
        ThisComponent.StyleFamilies.getByName("PageStyles").insertByName(stName, ps)
    End If
    ps = ThisComponent.StyleFamilies.getByName("PageStyles").getByName(stName)
    ps.Width = w * 100
    ps.Height = h * 100
    ps.LeftMargin = 200
    ps.RightMargin = 200
    ps.TopMargin = 200
    ps.BottomMargin = 200
    ps.HeaderIsOn = False
    ps.FooterIsOn = False
    ps.PrintGrid = False
    ps.CenterHorizontally = True
    lay.PageStyle = stName
    lay.getColumns().getByIndex(0).Width = (w - 4) * 100
    dp = lay.getDrawPage()
    For i = 0 To UBound(d)
        e = NormEI(CStr(d(i)(0)))
        If e <> "" Then
            n = CLng(Mid(e, 4))
            row = st.getCellRangeByPosition(0, n, 9, n).getDataArray()(0)
            If CStr(row(0)) <> e Then
                nBad = nBad + 1
                If bad = "" Then bad = e & " нет в снимке"
            Else
                c = 1
                If VarType(d(i)(1)) = 5 Then c = CLng(d(i)(1))
                If c < 1 Then c = 1
                For x = 1 To c
                    PutLabel(lay, dp, r, row, w, h)
                    r = r + ROWS_PER_LABEL
                    k = k + 1
                Next x
            End If
        ElseIf CStr(d(i)(0)) <> "" Then
            nBad = nBad + 1
            If bad = "" Then bad = "«" & d(i)(0) & "» — не номер ЕИ"
        End If
    Next i
    If nBad > 0 Then
        LblBuild = "ERR:ошибок в списке: " & nBad & " (" & bad & "); построено этикеток " & k
    Else
        LblBuild = "OK:этикеток " & k & " (" & w & "×" & h & " мм) на листе «" & SH_LAYOUT & "»"
    End If
    Exit Function
EH:
    LblBuild = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function NormEI(ByVal s As String) As String
    Dim t As String
    t = UCase(Trim(s))
    If Left(t, 2) = "ЕИ" Or Left(t, 2) = "EI" Then t = Mid(t, 3)
    If Left(t, 1) = "-" Then t = Mid(t, 2)
    If t <> "" And Len(t) <= 8 And IsNumeric(t) And InStr(t, ",") = 0 And InStr(t, ".") = 0 Then NormEI = "ЕИ-" & Right("00000000" & CLng(t), 8)
End Function

' one label from row r: EI large, name, article, place, the barcode below; the first row starts a new page
Private Sub PutLabel(lay As Object, dp As Object, r As Long, row As Variant, w As Double, h As Double)
    Dim hrow As Double, rows As Object, i As Integer, c As Object, bars As String, x As Long, j As Integer, bw As Long, y As Long
    Dim shp As Object, pt As New com.sun.star.awt.Point, sz As New com.sun.star.awt.Size, modW As Long, bcH As Long
    hrow = (h - 4) * 100 / ROWS_PER_LABEL
    rows = lay.getRows()
    For i = 0 To ROWS_PER_LABEL - 1
        rows.getByIndex(r + i).Height = hrow
    Next i
    rows.getByIndex(r).IsStartOfNewPage = True
    c = lay.getCellByPosition(0, r)
    c.setString(row(0))
    c.CharHeight = Setting(11, 20)
    c.CharWeight = 150
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    If Flag(7) Then SetLine(lay, r + 1, Left(CStr(row(1)), 60), 9)
    If Flag(8) Then SetLine(lay, r + 2, IIf(CStr(row(2)) <> "", "Арт. " & row(2), ""), 9)
    If Flag(9) Then SetLine(lay, r + 3, IIf(CStr(row(5)) <> "", "Место: " & row(5), ""), 9)
    If Flag(10) Then
        bars = Code128C(Mid(CStr(row(0)), 4))
        modW = 30
        bcH = hrow * 0.8
        x = lay.getCellByPosition(0, r + 4).Position.X + ((w - 4) * 100 - modW * 101) / 2
        y = lay.getCellByPosition(0, r + 4).Position.Y + hrow * 0.1
        For j = 1 To Len(bars)
            bw = CLng(Mid(bars, j, 1)) * modW
            If j Mod 2 = 1 Then
                shp = ThisComponent.createInstance("com.sun.star.drawing.RectangleShape")
                pt.X = x
                pt.Y = y
                sz.Width = bw
                sz.Height = bcH
                shp.setPosition(pt)
                shp.setSize(sz)
                dp.add(shp)
                shp.FillColor = RGB(0, 0, 0)
                shp.LineStyle = com.sun.star.drawing.LineStyle.NONE
                shp.Anchor = lay.getCellByPosition(0, r + 4)
            End If
            x = x + bw
        Next j
    End If
End Sub

Private Sub SetLine(lay As Object, r As Long, s As String, sz As Double)
    Dim c As Object
    c = lay.getCellByPosition(0, r)
    c.setString(s)
    c.CharHeight = sz
    c.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    c.IsTextWrapped = True
End Sub

' «Печать»: the usual print window of LibreOffice for the layout (the system printer / CUPS)
Sub BtnLblPrint(Optional oEvent As Variant)
    Dim d As Object, args(0) As New com.sun.star.beans.PropertyValue
    If Not ThisComponent.Sheets.hasByName(SH_LAYOUT) Then
        TbMsg("ERR:сначала «Построить этикетки»")
        Exit Sub
    End If
    ThisComponent.getCurrentController().setActiveSheet(ThisComponent.Sheets.getByName(SH_LAYOUT))
    d = CreateUnoService("com.sun.star.frame.DispatchHelper")
    d.executeDispatch(ThisComponent.getCurrentController().getFrame(), ".uno:Print", "", 0, args())
End Sub

' «В PDF»: the layout into ../WMS_Labels/labels_<time>.pdf (to check it without a printer)
Function LblPdf() As String
    Dim d As String, sfa As Object, url As String, args(1) As New com.sun.star.beans.PropertyValue, fd(0) As New com.sun.star.beans.PropertyValue
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    d = TbResolve("../WMS_Labels")
    If Not sfa.exists(d) Then sfa.createFolder(d)
    url = d & "labels_" & Format(Now(), "YYYYMMDD-HHMMSS") & ".pdf"
    fd(0).Name = "Selection"
    fd(0).Value = ThisComponent.Sheets.getByName(SH_LAYOUT)
    args(0).Name = "FilterName"
    args(0).Value = "calc_pdf_Export"
    args(1).Name = "FilterData"
    args(1).Value = fd()
    ThisComponent.storeToURL(url, args())
    LblPdf = "OK:" & ConvertFromURL(url)
    Exit Function
EH:
    LblPdf = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Sub BtnLblPdf(Optional oEvent As Variant)
    TbMsg(LblPdf())
End Sub
