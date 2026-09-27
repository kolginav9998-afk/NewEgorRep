' TbImporter — WMS_IMPORTER: безопасная подготовка массовых приходов из CSV / ODS / XLSX (задание «FINAL WMS MARATHON», §6).
' «Загрузить таблицу» — файл поставщика (накладная, список, выгрузка) на лист «Таблица» без ручного копирования строк;
' лист «Сопоставление» — какая колонка файла идёт в какую колонку WMS (угадывается по заголовкам, можно поправить); цель —
' «Иной приход» или «Заказы» (лист «Настройки»). «Проверить» — каждая строка по правилам WMS (обязательные поля, количество
' по правилам ядра, даты) и повторы: строка, совпадающая с другой строкой файла, и поступление, которое уже есть в WMS
' (по последнему снимку: «Иной приход» — тип, наименование, артикул, количество, дата; «Заказы» — № заказа и
' наименование); итог — в колонке «Проверка», а лист «Предпросмотр» показывает ровно то, что уйдёт в пакет; «Сформировать
' пакет» — только когда ошибок нет: файл WMS-BATCH-1 для WMS («Главная» → «Загрузить пакет»), где каждая строка ещё раз
' проверяется и проводится обычной операцией. Инструмент книгу WMS не открывает и ничего в учёт не пишет.
Option Explicit

Private Const SH_DATA = "Таблица"
Private Const SH_MAP = "Сопоставление"
Private Const SH_PREV = "Предпросмотр"
Private Const MAP_ROW = 1              ' 0-based: the keys from row 2 of «Сопоставление» (A key, B column of WMS, C column of the file, D required)
Private Const TARGET_SPECIAL = "Иной приход"
Private Const TARGET_ORDERS = "Заказы"

Global gImpFile As String             ' test seam: the file «Загрузить таблицу» takes without the file dialog

' ================================================================ the keys of the batch (docs/EXPORT_CONTRACT.md §3)

' key | name on the sheet of WMS | required (1) | kind (T text, Q quantity, D date, N number) | header synonyms (lower case, «/»)
Private Function Keys(sTarget As String) As Variant
    If sTarget = TARGET_ORDERS Then
        Keys = Array( _
            "order_no|№ заказа|0|T|№ заказа/номер заказа/заказ", _
            "name|Полное наименование|1|T|наименование/полное наименование/название/товар/номенклатура", _
            "doc_no|Номер документа|0|T|номер документа/документ/упд/накладная/№ документа", _
            "invoice_no|Номер счёта|0|T|номер счёта/номер счета/счёт/счет", _
            "article|Артикул|0|T|артикул/код товара/код/sku", _
            "qty_fact|Фактическое количество|0|Q|фактическое количество/факт/получено/принято", _
            "qty_doc|Количество по документу|0|Q|количество по документу/по документу", _
            "qty_ordered|Заказанное количество|1|Q|заказанное количество/заказано/количество/кол-во/кол.", _
            "unit|Единица измерения|1|T|единица измерения/единица/ед./ед. изм./ед.изм./ед", _
            "price|Цена|0|N|цена/цена за единицу", _
            "amount|Сумма|0|N|сумма/стоимость", _
            "supplier|Площадка / Поставщик|1|T|поставщик/площадка/площадка / поставщик/контрагент", _
            "seller|Продавец|0|T|продавец", _
            "date_received|Дата поступления|0|D|дата поступления/дата прихода/дата приёмки", _
            "date_doc|Дата документа|0|D|дата документа/дата упд/дата накладной", _
            "date_order|Дата заказа|0|D|дата заказа", _
            "date_expected|Ожидаемая дата|0|D|ожидаемая дата/срок поставки/дата поставки", _
            "buyer|Покупатель|0|T|покупатель/заказчик", _
            "category|Категория|0|T|категория/группа", _
            "assignee|Кому назначено|0|T|кому назначено/для кого", _
            "place|Место|0|T|место/место хранения/ячейка", _
            "note|Комментарий|0|T|комментарий/примечание", _
            "lead_days|Срок поставки, дней|0|N|срок поставки, дней/дней")
    Else
        Keys = Array( _
            "type|Тип прихода|1|T|тип прихода/тип/источник/тип источника", _
            "event|№ поступления|0|T|№ поступления/номер поступления", _
            "name|Наименование|1|T|наименование/название/товар/номенклатура", _
            "article|Артикул|0|T|артикул/код товара/код/sku", _
            "qty|Количество|1|Q|количество/кол-во/кол./количество (остаток)", _
            "unit|Единица измерения|1|T|единица измерения/единица/ед./ед. изм./ед.изм./ед", _
            "date|Дата поступления|1|D|дата поступления/дата прихода/дата", _
            "place|Место хранения|1|T|место хранения/место/ячейка", _
            "category|Категория|0|T|категория/группа", _
            "from_who|Кто передал|0|T|кто передал/от кого/передал", _
            "doc|Документ|0|T|документ/накладная/номер документа", _
            "old_mark|Старая маркировка|0|T|старая маркировка/маркировка/инв. номер/инвентарный номер", _
            "note|Комментарий|0|T|комментарий/примечание")
    End If
End Function

Private Function Target() As String
    Target = IIf(Trim(TbSetting(5)) = TARGET_ORDERS, TARGET_ORDERS, TARGET_SPECIAL)
End Function

Private Function NormH(ByVal s As String) As String
    s = LCase(Trim(Replace(Replace(s, Chr(160), " "), "ё", "е")))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If Right(s, 1) = ":" Then s = Left(s, Len(s) - 1)
    NormH = s
End Function

' ================================================================ «Загрузить таблицу»

Sub BtnImpLoad(Optional oEvent As Variant)
    Dim fp As Object, url As String
    If gImpFile <> "" Then
        url = gImpFile
    Else
        fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
        fp.appendFilter("Таблицы (CSV, ODS, XLSX)", "*.csv;*.ods;*.xlsx;*.xls")
        If fp.execute() <> 1 Then Exit Sub
        url = fp.getFiles()(0)
    End If
    TbMsg(ImpLoad(url))
End Sub

' the first sheet of the file into «Таблица» (values as the file holds them); the mapping guessed by the headers
Function ImpLoad(url As String) As String
    Dim args(2) As New com.sun.star.beans.PropertyValue, src As Object, d As Variant, sh As Object, cur As Object, last As Long, lastC As Long
    Dim enc As String, sfa As Object, head As String, sep As String
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(url) Then
        ImpLoad = "ERR:файл " & ConvertFromURL(url) & " не найден"
        Exit Function
    End If
    args(0).Name = "Hidden"
    args(0).Value = True
    args(1).Name = "ReadOnly"
    args(1).Value = True
    If LCase(Right(url, 4)) = ".csv" Then
        ' separators «;», «,», tab; every column as text (a leading zero of a code is kept; WMS parses quantities itself);
        ' the encoding of «Настройки» (UTF-8 by default, Windows-1251 for old programs)
        enc = IIf(InStr(LCase(TbSetting(8)), "1251") > 0, "33", "76")
        ' one separator — the one of the header line (a comma inside «1,5» must not split a row): «;», else tab, else «,»
        head = TbReadText(url)
        If InStr(head, Chr(10)) > 0 Then head = Left(head, InStr(head, Chr(10)) - 1)
        If InStr(head, ";") > 0 Then
            sep = "59"
        ElseIf InStr(head, Chr(9)) > 0 Then
            sep = "9"
        Else
            sep = "44"
        End If
        args(2).Name = "FilterOptions"
        args(2).Value = sep & ",34," & enc & ",1,1/2/2/2/3/2/4/2/5/2/6/2/7/2/8/2/9/2/10/2/11/2/12/2/13/2/14/2/15/2/16/2/17/2/18/2/19/2/20/2/21/2/22/2/23/2/24/2/25/2/26/2/27/2/28/2/29/2/30/2"
    Else
        args(2).Name = "MacroExecutionMode"
        args(2).Value = 0
    End If
    src = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    If IsNull(src) Then
        ImpLoad = "ERR:файл не открылся"
        Exit Function
    End If
    cur = src.Sheets.getByIndex(0).createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    lastC = cur.getRangeAddress().EndColumn
    If last < 1 Then
        src.close(True)
        ImpLoad = "ERR:в файле нет строк данных (первая строка — заголовки колонок)"
        Exit Function
    End If
    d = src.Sheets.getByIndex(0).getCellRangeByPosition(0, 0, lastC, last).getDataArray()
    src.close(True)
    sh = ThisComponent.Sheets.getByName(SH_DATA)
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    sh.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn + 1, cur.getRangeAddress().EndRow).clearContents(1023)
    sh.getCellRangeByPosition(0, 0, lastC, last).setDataArray(d)
    sh.getCellByPosition(lastC + 1, 0).setString("Проверка")
    sh.getCellRangeByPosition(0, 0, lastC + 1, 0).CharWeight = 150
    TbPutSetting(3, ConvertFromURL(url))
    ImpLoad = "OK:строк " & last & ", колонок " & (lastC + 1) & "; " & ImpAutoMap()
    Exit Function
EH:
    ImpLoad = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' «Сопоставление»: every key of the target → the column of the file with a matching header ("" — none)
Function ImpAutoMap() As String
    Dim ks As Variant, i As Integer, f As Variant, syn As Variant, j As Integer, c As Integer, heads As Variant, sh As Object, mp As Object
    Dim nCols As Integer, found As String, nMapped As Integer, missing As String, rows() As Variant
    sh = ThisComponent.Sheets.getByName(SH_DATA)
    nCols = DataCols(sh)
    If nCols = 0 Then
        ImpAutoMap = "таблица не загружена"
        Exit Function
    End If
    heads = sh.getCellRangeByPosition(0, 0, nCols - 1, 0).getDataArray()(0)
    ks = Keys(Target())
    ReDim rows(UBound(ks))
    For i = 0 To UBound(ks)
        f = Split(ks(i), "|")
        syn = Split(f(4), "/")
        found = ""
        For c = 0 To nCols - 1
            For j = 0 To UBound(syn)
                If NormH(CStr(heads(c))) = syn(j) Then
                    found = CStr(heads(c))
                    Exit For
                End If
            Next j
            If found <> "" Then Exit For
        Next c
        rows(i) = Array(f(0), f(1), found, IIf(f(2) = "1", "обязательно", ""))
        If found <> "" Then
            nMapped = nMapped + 1
        ElseIf f(2) = "1" Then
            missing = missing & IIf(missing <> "", ", ", "") & f(1)
        End If
    Next i
    mp = ThisComponent.Sheets.getByName(SH_MAP)
    mp.getCellRangeByPosition(0, MAP_ROW, 3, MAP_ROW + 40).clearContents(1023)
    mp.getCellRangeByPosition(0, MAP_ROW, 3, MAP_ROW + UBound(rows)).setDataArray(rows)
    ImpAutoMap = "цель «" & Target() & "»: сопоставлено колонок " & nMapped & IIf(missing <> "", "; нет обязательных: " & missing & _
        " (впишите заголовок колонки файла в «" & SH_MAP & "», колонка C" & IIf(Target() = TARGET_SPECIAL, ", или тип прихода по умолчанию в «Настройки»", "") & ")", "")
End Function

' the number of columns of the file on «Таблица» (without «Проверка»)
Private Function DataCols(sh As Object) As Integer
    Dim cur As Object, heads As Variant, c As Integer, n As Integer
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    If cur.getRangeAddress().EndRow < 0 Then Exit Function
    heads = sh.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, 0).getDataArray()(0)
    For c = 0 To UBound(heads)
        If CStr(heads(c)) = "Проверка" Then Exit For
        n = c + 1
    Next c
    DataCols = n
End Function

' ================================================================ «Проверить»

Sub BtnImpCheck(Optional oEvent As Variant)
    TbMsg(ImpCheck())
End Sub

' the rules of WMS for a quantity (docs D-030): "" when valid, else why; value in q
Function QtyProblem(v As Variant, ByRef q As Double) As String
    Dim t As String, i As Integer, ch As String, ip As String, fp As String, nSep As Integer
    q = 0
    If VarType(v) = 5 Then
        If v <= 0 Then
            QtyProblem = "количество должно быть больше 0"
        ElseIf Abs(v * 1000 - Int(v * 1000 + 0.5)) > 0.0001 Then
            QtyProblem = "больше 3 знаков после запятой"
        Else
            q = v
        End If
        Exit Function
    End If
    t = Trim(Replace(CStr(v), Chr(160), ""))
    If t = "" Then
        QtyProblem = "нет количества"
        Exit Function
    End If
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            If nSep = 0 Then ip = ip & ch Else fp = fp & ch
        ElseIf ch = "," Or ch = "." Then
            nSep = nSep + 1
        ElseIf ch = "-" Then
            QtyProblem = "«" & t & "» — отрицательное количество"
            Exit Function
        Else
            QtyProblem = "«" & t & "» — не число"
            Exit Function
        End If
    Next i
    If nSep > 1 Or ip = "" Then
        QtyProblem = "«" & t & "» — не число"
    ElseIf Len(fp) = 3 And Len(ip) <= 3 Then
        QtyProblem = "«" & t & "» — неоднозначно (тысячи или дробь?): запишите «" & ip & fp & "» или «" & ip & "," & fp & "»"
    ElseIf Len(fp) > 3 Then
        QtyProblem = "«" & t & "» — больше 3 знаков после запятой"
    Else
        q = Val(ip & IIf(fp <> "", "." & fp, ""))
        If q <= 0 Then QtyProblem = "количество должно быть больше 0"
    End If
End Function

' a date as dd.mm.yyyy ("" and why when not a date): a date cell, dd.mm.yyyy or ГГГГ-ММ-ДД
Function DateText(v As Variant, ByRef why As String) As String
    Dim t As String, p As Variant, d As Double
    why = ""
    If VarType(v) = 5 Or VarType(v) = 7 Then
        d = CDbl(v)
    Else
        t = Trim(CStr(v))
        If t = "" Then Exit Function
        If Len(t) = 10 And Mid(t, 5, 1) = "-" And Mid(t, 8, 1) = "-" Then
            p = Split(t, "-")
            If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then d = CDbl(DateSerial(Val(p(0)), Val(p(1)), Val(p(2))))
        Else
            p = Split(Replace(t, "/", "."), ".")
            If UBound(p) = 2 Then
                If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) And Len(p(2)) = 4 Then
                    If Val(p(1)) >= 1 And Val(p(1)) <= 12 And Val(p(0)) >= 1 And Val(p(0)) <= 31 Then d = CDbl(DateSerial(Val(p(2)), Val(p(1)), Val(p(0))))
                End If
            End If
        End If
        If d = 0 Then
            why = "«" & t & "» — не дата (нужно дд.мм.гггг)"
            Exit Function
        End If
    End If
    If d > Int(CDbl(Now())) + 0.5 Then
        why = Format(CDate(d), "DD.MM.YYYY") & " — дата в будущем"
        Exit Function
    End If
    DateText = Format(CDate(d), "DD.MM.YYYY")
End Function

' the key of a possible repeat: a special receipt — type, name, article, quantity, date; an order — № заказа and name
' ("" — an order without a number is not compared)
Private Function DupKey(sType As String, sName As String, sArt As String, sQty As String, sDate As String, sOrder As String) As String
    If Target() = TARGET_ORDERS Then
        If Trim(sOrder) <> "" Then DupKey = LCase(Trim(sOrder)) & Chr(1) & LCase(Trim(sName))
    Else
        DupKey = LCase(Trim(sType)) & Chr(1) & LCase(Trim(sName)) & Chr(1) & UCase(Replace(Replace(Trim(sArt), " ", ""), "-", "")) _
            & Chr(1) & sQty & Chr(1) & sDate
    End If
End Function

' the item of a collection by its key, Empty when there is none
Private Function Found(c As Collection, k As String) As Variant
    On Error GoTo NONE
    Found = c.Item(k)
    Exit Function
NONE:
    Found = Empty
End Function

' the posted receipts of the latest snapshot by DupKey → where they are in WMS; "" or why they were not read
Private Function WmsReceipts(c As Collection) As String
    Dim why As String, snap As String, sh As Object, d As Variant, i As Long, k As String, last As Long, q As String
    snap = TbPickSnapshot(why)
    If snap = "" Then
        WmsReceipts = why
        Exit Function
    End If
    If Target() = TARGET_ORDERS Then
        last = TbLoadTable(snap, "orders.csv", "_orders", True)
        If last < 1 Then Exit Function
        d = ThisComponent.Sheets.getByName("_orders").getCellRangeByPosition(0, 1, 1, last).getDataArray()
        For i = 0 To UBound(d)
            k = DupKey("", CStr(d(i)(1)), "", "", "", CStr(d(i)(0)))
            If k <> "" And IsEmpty(Found(c, k)) Then c.Add("«" & TARGET_ORDERS & "», строка " & (i + 2) & " (заказ " & d(i)(0) & ")", k)
        Next i
    Else
        last = TbLoadTable(snap, "special.csv", "_special", True)
        If last < 1 Then Exit Function
        d = ThisComponent.Sheets.getByName("_special").getCellRangeByPosition(0, 1, 16, last).getDataArray()
        For i = 0 To UBound(d)
            If Left(CStr(d(i)(16)), 9) = "Проведено" And VarType(d(i)(5)) = 5 And VarType(d(i)(7)) = 5 Then
                q = Replace(Trim(Str(d(i)(5))), ".", ",")
                k = DupKey(CStr(d(i)(1)), CStr(d(i)(3)), CStr(d(i)(4)), q, Format(CDate(d(i)(7)), "DD.MM.YYYY"), "")
                If IsEmpty(Found(c, k)) Then c.Add("«" & TARGET_SPECIAL & "», строка " & (i + 2) & " (" & d(i)(13) & ")", k)
            End If
        Next i
    End If
End Function

' checks every row; writes «OK» or the problems into «Проверка»; fills the prepared rows for the batch (keys, values)
Function ImpPrepare(ByRef outKeys As Variant, ByRef outRows As Variant, ByRef nErr As Long) As String
    Dim sh As Object, mp As Object, nCols As Integer, heads As Variant, ks As Variant, i As Integer, f As Variant, col() As Integer, c As Integer
    Dim cur As Object, last As Long, d As Variant, r As Long, msg As String, v As Variant, q As Double, why As String, vals() As Variant, used() As Integer
    Dim nUsed As Integer, k As Integer, rows() As Variant, nRows As Long, status() As Variant, sKind As String, defKind As String, m As Variant
    Dim seen As New Collection, wms As New Collection, noWms As String, pos(5) As Integer, kn As Variant, allKey As String, dk As String, w As Variant
    Dim nDup As Long, pv As Object, heads2() As Variant
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_DATA)
    nCols = DataCols(sh)
    If nCols = 0 Then
        ImpPrepare = "ERR:сначала «Загрузить таблицу»"
        Exit Function
    End If
    heads = sh.getCellRangeByPosition(0, 0, nCols - 1, 0).getDataArray()(0)
    ks = Keys(Target())
    mp = ThisComponent.Sheets.getByName(SH_MAP).getCellRangeByPosition(0, MAP_ROW, 2, MAP_ROW + UBound(ks)).getDataArray()
    ReDim col(UBound(ks))
    ReDim used(UBound(ks))
    For i = 0 To UBound(ks)
        f = Split(ks(i), "|")
        col(i) = -1
        If CStr(mp(i)(0)) <> f(0) Then
            ImpPrepare = "ERR:лист «" & SH_MAP & "» не соответствует цели «" & Target() & "» — нажмите «Загрузить таблицу» ещё раз"
            Exit Function
        End If
        If Trim(CStr(mp(i)(2))) <> "" Then
            For c = 0 To nCols - 1
                If NormH(CStr(heads(c))) = NormH(CStr(mp(i)(2))) Then col(i) = c
            Next c
            If col(i) < 0 Then
                ImpPrepare = "ERR:в таблице нет колонки «" & mp(i)(2) & "» (" & SH_MAP & ", " & f(1) & ")"
                Exit Function
            End If
        End If
    Next i
    defKind = Trim(TbSetting(6))
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    d = sh.getCellRangeByPosition(0, 1, nCols - 1, last).getDataArray()
    ' the keys of the batch: mapped ones, and «type» when it comes from the default of «Настройки»
    For i = 0 To UBound(ks)
        f = Split(ks(i), "|")
        If col(i) >= 0 Or (f(0) = "type" And defKind <> "") Then
            used(nUsed) = i
            nUsed = nUsed + 1
        End If
    Next i
    ' the positions of the keys of a repeat among the values of a row (-1 — not in the batch)
    kn = Array("type", "name", "article", "qty", "date", "order_no")
    For i = 0 To 5
        pos(i) = -1
        For k = 0 To nUsed - 1
            If Split(ks(used(k)), "|")(0) = kn(i) Then pos(i) = k
        Next k
    Next i
    noWms = WmsReceipts(wms)
    ReDim rows(UBound(d))
    ReDim status(UBound(d))
    For r = 0 To UBound(d)
        msg = ""
        ReDim vals(nUsed - 1)
        m = ""
        For c = 0 To nCols - 1
            m = m & CStr(d(r)(c))
        Next c
        If Trim(m) = "" Then
            status(r) = Array("")
        Else
            For k = 0 To nUsed - 1
                i = used(k)
                f = Split(ks(i), "|")
                If col(i) >= 0 Then v = d(r)(col(i)) Else v = ""
                If f(0) = "type" And Trim(CStr(v)) = "" Then v = defKind
                Select Case f(3)
                Case "Q"
                    If Trim(CStr(v)) = "" Then
                        vals(k) = ""
                    Else
                        why = QtyProblem(v, q)
                        If why <> "" Then
                            msg = msg & IIf(msg <> "", "; ", "") & f(1) & ": " & why
                        Else
                            vals(k) = Replace(Trim(Str(q)), ".", ",")
                        End If
                    End If
                Case "D"
                    vals(k) = DateText(v, why)
                    If why <> "" Then msg = msg & IIf(msg <> "", "; ", "") & f(1) & ": " & why
                Case "N"
                    ' IIf of Basic evaluates both branches: the number and the text are handled apart
                    If VarType(v) = 5 Then
                        vals(k) = Replace(Trim(Str(v)), ".", ",")
                    ElseIf Trim(CStr(v)) = "" Or IsNumeric(Replace(CStr(v), ",", ".")) Then
                        vals(k) = Trim(CStr(v))
                    Else
                        msg = msg & IIf(msg <> "", "; ", "") & f(1) & ": «" & v & "» не число"
                    End If
                Case Else
                    If VarType(v) = 5 Then
                        If Int(v) = v Then vals(k) = Trim(Str(v)) Else vals(k) = Replace(Trim(Str(v)), ".", ",")
                    Else
                        vals(k) = Trim(CStr(v))
                    End If
                End Select
                If f(2) = "1" And Trim(CStr(vals(k))) = "" And InStr(msg, f(1)) = 0 Then msg = msg & IIf(msg <> "", "; ", "") & "нет «" & f(1) & "»"
            Next k
            ' the kind of a special receipt; a part needs its article
            If Target() = TARGET_SPECIAL Then
                sKind = ""
                For k = 0 To nUsed - 1
                    If Split(ks(used(k)), "|")(0) = "type" Then sKind = CStr(vals(k))
                Next k
                If sKind <> "" And InStr("|Офис|Производство|Детали|Старый склад|Иной|", "|" & sKind & "|") = 0 Then
                    msg = msg & IIf(msg <> "", "; ", "") & "тип прихода «" & sKind & "» (Офис, Производство, Детали, Старый склад, Иной)"
                End If
            End If
            For i = 0 To UBound(ks)
                If col(i) < 0 And Split(ks(i), "|")(2) = "1" And Not (Split(ks(i), "|")(0) = "type" And defKind <> "") Then
                    If InStr(msg, Split(ks(i), "|")(1)) = 0 Then msg = msg & IIf(msg <> "", "; ", "") & "нет «" & Split(ks(i), "|")(1) & "»"
                End If
            Next i
            ' repeats: the same row earlier in the file; a receipt that WMS already has
            If msg = "" Then
                allKey = LCase(Join(vals, Chr(1)))
                w = Found(seen, allKey)
                If Not IsEmpty(w) Then
                    msg = "повтор строки " & w & " файла (все значения совпадают) — удалите повтор или укажите различие, например в «Комментарий»"
                Else
                    seen.Add(r + 2, allKey)
                    dk = DupKey(PosVal(vals, pos(0)), PosVal(vals, pos(1)), PosVal(vals, pos(2)), PosVal(vals, pos(3)), PosVal(vals, pos(4)), PosVal(vals, pos(5)))
                    If dk <> "" Then
                        w = Found(wms, dk)
                        If Not IsEmpty(w) Then msg = "уже есть в WMS: " & w & IIf(Target() = TARGET_ORDERS, " — новую поставку по заказу вводят в WMS " _
                            & "(«Ещё поступление»)", " — если это действительно новое поступление, проведите его в WMS вручную")
                    End If
                End If
                If msg <> "" Then nDup = nDup + 1
            End If
            If msg <> "" Then
                nErr = nErr + 1
                status(r) = Array(msg)
            Else
                status(r) = Array("OK")
                rows(nRows) = vals
                nRows = nRows + 1
            End If
        End If
    Next r
    sh.getCellRangeByPosition(nCols, 1, nCols, last).setDataArray(status)
    ' «Предпросмотр»: exactly the rows and the columns of the batch
    pv = ThisComponent.Sheets.getByName(SH_PREV)
    pv.clearContents(1023)
    ReDim heads2(nUsed - 1)
    For k = 0 To nUsed - 1
        heads2(k) = Split(ks(used(k)), "|")(1)
    Next k
    pv.getCellRangeByPosition(0, 0, nUsed - 1, 0).setDataArray(Array(heads2))
    pv.getCellRangeByPosition(0, 0, nUsed - 1, 0).CharWeight = 150
    For r = 0 To nRows - 1
        pv.getCellRangeByPosition(0, r + 1, nUsed - 1, r + 1).setDataArray(Array(rows(r)))
    Next r
    Dim kk() As String
    ReDim kk(nUsed - 1)
    For k = 0 To nUsed - 1
        kk(k) = Split(ks(used(k)), "|")(0)
    Next k
    outKeys = kk
    If nRows > 0 Then
        Dim rr() As Variant
        ReDim rr(nRows - 1)
        For r = 0 To nRows - 1
            rr(r) = rows(r)
        Next r
        outRows = rr
    Else
        outRows = Array()
    End If
    ImpPrepare = "OK:строк без ошибок " & nRows & ", с ошибками " & nErr & IIf(nDup > 0, " (из них повторов " & nDup & ")", "") _
        & IIf(noWms <> "", "; повторы с WMS не проверены: " & noWms, "; повторы с WMS проверены по последнему снимку")
    Exit Function
EH:
    ImpPrepare = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function PosVal(vals As Variant, k As Integer) As String
    If k >= 0 Then PosVal = CStr(vals(k))
End Function

Function ImpCheck() As String
    Dim ks As Variant, rs As Variant, nErr As Long
    ImpCheck = ImpPrepare(ks, rs, nErr)
    If Left(ImpCheck, 3) = "OK:" And nErr > 0 Then ImpCheck = "ERR:" & Mid(ImpCheck, 4) & " — исправьте строки (колонка «Проверка») и проверьте снова"
End Function

' ================================================================ «Сформировать пакет»

Sub BtnImpBatch(Optional oEvent As Variant)
    TbMsg(ImpBatch())
End Sub

Function ImpBatch() As String
    Dim ks As Variant, rs As Variant, nErr As Long, res As String, id As String, path As String, extra As String
    On Error GoTo EH
    res = ImpPrepare(ks, rs, nErr)
    If Left(res, 3) <> "OK:" Then
        ImpBatch = res
        Exit Function
    End If
    If nErr > 0 Then
        ImpBatch = "ERR:в таблице строк с ошибками " & nErr & " — пакет не создан (он создаётся только целиком)"
        Exit Function
    End If
    If UBound(rs) < 0 Then
        ImpBatch = "ERR:нет строк для пакета"
        Exit Function
    End If
    If UBound(rs) + 1 > 5000 Then
        ImpBatch = "ERR:строк " & (UBound(rs) + 1) & " — WMS принимает пакет до 5000 строк; разделите файл"
        Exit Function
    End If
    id = IIf(Target() = TARGET_ORDERS, "ORD-", "SPR-") & Format(Now(), "YYYYMMDD-HHMMSS")
    If Target() = TARGET_SPECIAL And LCase(Left(Trim(TbSetting(7)), 1)) = "д" Then extra = "#shared;yes" & Chr(10)
    path = TbWriteBatch(Target(), id, "WMS_IMPORTER", Mid(TbSetting(3), TbInStrRev(TbSetting(3), "/") + 1), ks, rs, extra)
    ImpBatch = "OK:пакет " & id & ": строк " & (UBound(rs) + 1) & " для листа «" & Target() & "» — " & ConvertFromURL(path) _
        & ". В WMS: «Главная» → «Загрузить пакет»"
    Exit Function
EH:
    ImpBatch = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ test seams

Function TestImpFile(url As String) As String
    gImpFile = url
    TestImpFile = "OK"
End Function
