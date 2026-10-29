' WmsLedger — 1.0 (этап G «Контроль»): учёт для бухгалтерии и ревизора, страница «Отчёты». Остаток на любую прошедшую дату
' (К7), оборотная ведомость за период в единицах и рублях (К8), выгрузка движений с документами, ценами и суммами в файл (К9),
' сверка остатков с файлом из 1С (К10), закрытие месяца (К13): итоги месяца — в файлах и в _CLOSE (операция MONTH_CLOSE),
' операция с датой в закрытом месяце проводится только после предупреждения и с причиной (BACKDATED, BACKWHY в журнале).
' Отчёты только читают книгу; файлы — в «WMS_Документы» рядом с книгой, без интернета. Library WmsPlus.
Option Explicit

Private Const SH_REPORTS = "Отчёты"
Public Const SH_CLOSE = "_CLOSE"
' _CLOSE: one row per closing of a month (written only by MONTH_CLOSE and MONTH_OPEN)
Private Const CL_MONTH = 0
Private Const CL_TO = 1
Private Const CL_TS = 2
Private Const CL_STATE = 3
Private Const CL_QTY = 4
Private Const CL_SUM = 5
Private Const CL_FILE = 6
Private Const CL_WHY = 7
Private Const CL_LAST = 7
Private Const CL_MAX = 600
Private Const CL_CLOSED = "CLOSED"
Private Const CL_OPENED = "OPENED"
' the scratch columns of «Печать» where Calc sorts for WMS (outside the print area A:G, cleared after each sort)
Private Const LD_SORT_COL = 29
' the kinds of a movement
Private Const LK_RCV = 1
Private Const LK_SPEC = 2
Private Const LK_ISS = 3
Private Const LK_RET = 4
Private Const LK_WOFF = 5
Private Const LK_INV = 6
Private Const LK_SUP = 7
Private Const LK_SPLIT = 8
Private Const LK_LEG = 9
' 1.0 (Ф8): a kit build — the components out (−), the kit in (+)
Private Const LK_KIT = 10

' the movements read by LedgerRead: EI number, date, kind, quantity (+ in, − out), operation, document, counterparty
Global gLdN As Long
Global gLdEi() As Long
Global gLdDate() As Double
Global gLdKind() As Integer
Global gLdQty() As Double
Global gLdOp() As String
Global gLdDoc() As String
Global gLdWho() As String
' the registry by EI number 1 … gLdMax: balance of «Наличие», price (−1 none), card
Global gLdMax As Long
Global gLdNow() As Double
Global gLdPrice() As Double
Global gLdCanon() As String
Global gLdName() As String
Global gLdArt() As String
Global gLdUnit() As String
Global gLdPlace() As String
' test seam: the file of «Сверка с 1С» instead of the file picker
Global gLdFileTest As String
' А2: the money of the stock, valid for LAST_SEQ gLdMoneySeq
Global gLdMoneyOk As Boolean
Global gLdMoneySeq As Long
Global gLdMoney As Double
Global gLdMoneyN As Long

' ================================================================ the movements of the book

' every movement of the book read from its sheets once: the posted receipts of «Заказы» (by _RCV), «Иной приход», «Выдачи»,
' «Возврат», «Корректировки» and the new EIs of the splits (_ADJ); the registry «Наличие» and the price of every EI (J of its
' receipt; a split EI — that of its source). A storno row is no movement: its effect was taken back
Sub LedgerRead()
    Dim st As Object, r0 As Long, r1 As Long, d As Variant, i As Long, n As Long
    gLdMax = CLng(SysNum(SK_NEXT_EI)) - 1
    If gLdMax < 0 Then gLdMax = 0
    ReDim gLdNow(gLdMax)
    ReDim gLdPrice(gLdMax)
    ReDim gLdCanon(gLdMax)
    ReDim gLdName(gLdMax)
    ReDim gLdArt(gLdMax)
    ReDim gLdUnit(gLdMax)
    ReDim gLdPlace(gLdMax)
    gLdN = 0
    ReDim gLdEi(1023)
    ReDim gLdDate(1023)
    ReDim gLdKind(1023)
    ReDim gLdQty(1023)
    ReDim gLdOp(1023)
    ReDim gLdDoc(1023)
    ReDim gLdWho(1023)
    For n = 0 To gLdMax
        gLdPrice(n) = -1
    Next n
    st = WmsIssue.StockSheet()
    For r0 = 1 To gLdMax Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > gLdMax Then r1 = gLdMax
        d = st.getCellRangeByPosition(0, r0, SC_PLACE, r1).getDataArray()
        For i = 0 To r1 - r0
            n = r0 + i
            gLdCanon(n) = CStr(d(i)(SC_EI))
            gLdName(n) = CStr(d(i)(SC_NAME))
            gLdArt(n) = CStr(d(i)(SC_ART))
            gLdUnit(n) = CStr(d(i)(SC_UNIT))
            gLdPlace(n) = CStr(d(i)(SC_PLACE))
            If VarType(d(i)(SC_QTY)) = 5 Then gLdNow(n) = d(i)(SC_QTY)
        Next i
    Next r0
    LdReceipts()
    LdSheet(SH_SPECIAL)
    LdSheet(SH_ISSUES)
    LdSheet(SH_RETURNS)
    LdSheet(SH_ADJUST)
    LdSplits()
End Sub

Private Sub LdAdd(n As Long, dt As Double, kind As Integer, q As Double, sOp As String, sDoc As String, sWho As String)
    If n < 1 Or n > gLdMax Then Exit Sub
    If gLdN > UBound(gLdEi) Then
        ReDim Preserve gLdEi(2 * gLdN + 1)
        ReDim Preserve gLdDate(2 * gLdN + 1)
        ReDim Preserve gLdKind(2 * gLdN + 1)
        ReDim Preserve gLdQty(2 * gLdN + 1)
        ReDim Preserve gLdOp(2 * gLdN + 1)
        ReDim Preserve gLdDoc(2 * gLdN + 1)
        ReDim Preserve gLdWho(2 * gLdN + 1)
    End If
    gLdEi(gLdN) = n
    gLdDate(gLdN) = dt
    gLdKind(gLdN) = kind
    gLdQty(gLdN) = q
    gLdOp(gLdN) = sOp
    gLdDoc(gLdN) = sDoc
    gLdWho(gLdN) = sWho
    gLdN = gLdN + 1
End Sub

' a quantity of a cell: a number, or the text WMS wrote; 0 otherwise
Private Function LdQ(v As Variant) As Double
    Dim q As Double, msg As String
    If VarType(v) = 5 Then
        LdQ = v
    ElseIf CStr(v) <> "" Then
        If WmsOrders.ParseQtyValueAs(v, "F", q, msg) Then LdQ = q
    End If
End Function

Private Function LdPosted(v As Variant) As Boolean
    LdPosted = (Left(CStr(v), Len(ST_POSTED)) = ST_POSTED)
End Function

' the receipts: the row of «Заказы» that _RCV names for each EI (a live one is a movement; its price is the price of the EI);
' a receipt of the transfer from the old table also brings what was used before the transfer
Private Sub LdReceipts()
    Dim rcv As Object, osh As Object, lgs As Object, r0 As Long, r1 As Long, a As Variant, i As Long, n As Long, last As Long
    Dim rvRow() As Long, rvLive() As Boolean, rvAdd() As Boolean, lgKind() As String, lgPre() As Double, dt As Double, sDoc As String
    ReDim rvRow(gLdMax)
    ReDim rvLive(gLdMax)
    ReDim rvAdd(gLdMax)
    ReDim lgKind(gLdMax)
    ReDim lgPre(gLdMax)
    rcv = WmsOrders.RcvSheet()
    For r0 = 1 To gLdMax Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > gLdMax Then r1 = gLdMax
        a = rcv.getCellRangeByPosition(0, r0, RV_KIND, r1).getDataArray()
        For i = 0 To r1 - r0
            rvRow(r0 + i) = -1
            If CStr(a(i)(RV_EI)) <> "" And VarType(a(i)(RV_ROW)) = 5 Then
                rvRow(r0 + i) = CLng(a(i)(RV_ROW))
                rvLive(r0 + i) = (CStr(a(i)(RV_STATE)) = RV_LIVE)
                rvAdd(r0 + i) = (CStr(a(i)(RV_KIND)) = RV_ADD)
            End If
        Next i
    Next r0
    If gDoc.Sheets.hasByName(SH_LEG) Then
        lgs = gDoc.Sheets.getByName(SH_LEG)
        For r0 = 1 To gLdMax Step CHECK_CHUNK_ROWS
            r1 = r0 + CHECK_CHUNK_ROWS - 1
            If r1 > gLdMax Then r1 = gLdMax
            a = lgs.getCellRangeByPosition(0, r0, LG_PRE, r1).getDataArray()
            For i = 0 To r1 - r0
                lgKind(r0 + i) = CStr(a(i)(LG_KIND))
                If VarType(a(i)(LG_PRE)) = 5 Then lgPre(r0 + i) = a(i)(LG_PRE)
            Next i
        Next r0
    End If
    osh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(osh)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        a = osh.getCellRangeByPosition(0, r0, OC_EI, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(a(i)(OC_EI)) <> "" Then
                n = WmsView.EiNo(CStr(a(i)(OC_EI)), gLdMax + 1)
                If n >= 1 Then
                    If rvRow(n) = r0 + i Then
                        If VarType(a(i)(OC_PRICE)) = 5 Then gLdPrice(n) = a(i)(OC_PRICE)
                        If rvLive(n) Then
                            dt = WmsView.DateOf(a(i)(OC_RDATE))
                            sDoc = CStr(a(i)(OC_DOC)) & IIf(CStr(a(i)(OC_INVOICE)) <> "", " · счёт " & a(i)(OC_INVOICE), "")
                            LdAdd(n, dt, LK_RCV, LdQ(a(i)(OC_FACT)), IIf(rvAdd(n), "Ещё поступление", "Приход по заказу") & " № " & a(i)(OC_ORDER), _
                                sDoc, CStr(a(i)(OC_SUPPLIER)))
                            If (lgKind(n) = LG_RECEIPT Or lgKind(n) = LG_ADD) And lgPre(n) > 0.0000005 Then
                                LdAdd(n, dt, LK_LEG, -lgPre(n), "Израсходовано до переноса", "старая таблица", "")
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    Next r0
End Sub

' the posted rows of a sheet of movements
Private Sub LdSheet(sName As String)
    Dim sh As Object, last As Long, r0 As Long, r1 As Long, a As Variant, i As Long, n As Long, cLast As Integer, cEI As Integer
    Dim kind As String, ch As Double, k As Integer, sDoc As String
    If Not gDoc.Sheets.hasByName(sName) Then Exit Sub
    sh = gDoc.Sheets.getByName(sName)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    Select Case sName
    Case SH_SPECIAL
        cLast = XC_LAST
    Case SH_ISSUES
        cLast = IC_LAST
    Case SH_RETURNS
        cLast = RC_LAST
    Case Else
        cLast = AC_LAST
    End Select
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        a = sh.getCellRangeByPosition(0, r0, cLast, r1).getDataArray()
        For i = 0 To r1 - r0
            Select Case sName
            Case SH_SPECIAL
                If LdPosted(a(i)(XC_STATUS)) And CStr(a(i)(XC_EI)) <> "" Then
                    n = WmsView.EiNo(CStr(a(i)(XC_EI)), gLdMax + 1)
                    LdAdd(n, WmsView.DateOf(a(i)(XC_DATE)), LK_SPEC, LdQ(a(i)(XC_QTY)), CStr(a(i)(XC_TYPE)) & IIf(CStr(a(i)(XC_EVENT)) <> "", " № " & a(i)(XC_EVENT), ""), _
                        CStr(a(i)(XC_DOC)), CStr(a(i)(XC_WHO)))
                End If
            Case SH_ISSUES
                If LdPosted(a(i)(IC_CTL)) And VarType(a(i)(IC_NO)) = 5 Then
                    n = WmsView.EiNo(CStr(a(i)(IC_EI)), gLdMax + 1)
                    LdAdd(n, WmsView.DateOf(a(i)(IC_DATE)), LK_ISS, -LdQ(a(i)(IC_QTY)), "Выдача № " & a(i)(IC_NO), CStr(a(i)(IC_DOC)), CStr(a(i)(IC_WHO)))
                End If
            Case SH_RETURNS
                If LdPosted(a(i)(RC_CTL)) And VarType(a(i)(RC_NO)) = 5 Then
                    n = WmsView.EiNo(CStr(a(i)(RC_EI)), gLdMax + 1)
                    LdAdd(n, WmsView.DateOf(a(i)(RC_DATE)), LK_RET, LdQ(a(i)(RC_QTY)), "Возврат № " & a(i)(RC_NO), "по выдаче № " & a(i)(RC_ISSUE), CStr(a(i)(RC_WHO)))
                End If
            Case Else
                If LdPosted(a(i)(AC_CTL)) And VarType(a(i)(AC_NO)) = 5 And VarType(a(i)(AC_BEFORE)) = 5 And VarType(a(i)(AC_AFTER)) = 5 Then
                    kind = CStr(a(i)(AC_KIND))
                    ch = WmsIssue.Round3(a(i)(AC_AFTER) - a(i)(AC_BEFORE))
                    k = 0
                    If kind = "Списание" Then
                        k = LK_WOFF
                    ElseIf kind = "Инвентаризация" Then
                        k = LK_INV
                    ElseIf kind = "Разделение" Then
                        k = LK_SPLIT
                    ElseIf kind = "Возврат поставщику" Then
                        k = LK_SUP
                    ElseIf kind = "Сборка" Then
                        k = LK_KIT
                    End If
                    If k > 0 And Abs(ch) > 0.0000005 Then
                        n = WmsView.EiNo(CStr(a(i)(AC_EI)), gLdMax + 1)
                        sDoc = CStr(a(i)(AC_BATCH))
                        If k = LK_SUP And CStr(a(i)(AC_PLACE)) <> "" Then sDoc = CStr(a(i)(AC_PLACE))
                        LdAdd(n, WmsView.DateOf(a(i)(AC_DATE)), k, ch, kind & " № " & a(i)(AC_NO), sDoc, CStr(a(i)(AC_REASON)))
                    End If
                End If
            End Select
        Next i
    Next r0
End Sub

' the new EI of a live split starts with what was split off (its price — that of the source)
Private Sub LdSplits()
    Dim sh As Object, last As Long, a As Variant, i As Long, m As Long, src As Long, kitSum() As Double, kitBad() As Boolean, e As Long
    If Not gDoc.Sheets.hasByName(SH_ADJ) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_ADJ)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    a = sh.getCellRangeByPosition(0, 1, AJ_LAST, last).getDataArray()
    ReDim kitSum(gLdMax)
    ReDim kitBad(gLdMax)
    ' 1.0 (Ф8): the price of a kit — the money of its components ÷ the kits (the rows of a build come before its kit row, a
    ' kit built of kits finds their prices already known)
    For i = 0 To last - 1
        If CStr(a(i)(AJ_KIND)) = "KIT" And CStr(a(i)(AJ_STATE)) = RV_LIVE And VarType(a(i)(AJ_QTY)) = 5 And VarType(a(i)(AJ_FACT)) <> 5 Then
            m = WmsView.EiNo(CStr(a(i)(AJ_NEW)), gLdMax + 1)
            e = WmsView.EiNo(CStr(a(i)(AJ_EI)), gLdMax + 1)
            If m >= 1 And e >= 1 Then
                If e = m Then
                    If Not kitBad(m) And a(i)(AJ_QTY) > 0 Then gLdPrice(m) = kitSum(m) / a(i)(AJ_QTY)
                ElseIf gLdPrice(e) < 0 Then
                    kitBad(m) = True
                Else
                    kitSum(m) = kitSum(m) - a(i)(AJ_QTY) * gLdPrice(e)
                End If
            End If
        End If
    Next i
    For i = 0 To last - 1
        If CStr(a(i)(AJ_KIND)) = "SPLIT" And CStr(a(i)(AJ_STATE)) = RV_LIVE And CStr(a(i)(AJ_NEW)) <> "" And VarType(a(i)(AJ_QTY)) = 5 Then
            m = WmsView.EiNo(CStr(a(i)(AJ_NEW)), gLdMax + 1)
            src = WmsView.EiNo(CStr(a(i)(AJ_EI)), gLdMax + 1)
            If m >= 1 And src >= 1 Then
                If gLdPrice(m) < 0 Then gLdPrice(m) = gLdPrice(src)
            End If
            LdAdd(m, WmsView.DateOf(a(i)(AJ_DATE)), LK_SPLIT, -a(i)(AJ_QTY), "Отделён от " & a(i)(AJ_EI), "корректировка № " & a(i)(AJ_NO), "")
        End If
    Next i
End Sub

' ================================================================ helpers

' the order of n keys (numbers or texts; ascending, texts without regard to case) sorted by Calc itself — fast for 100 000:
' written to a scratch area of «Печать» (outside the print area), sorted, read back, cleared. out(i) — the number of the i-th key
Function SortByCalc(keys As Variant, n As Long) As Variant
    Dim sh As Object, a() As Variant, i As Long, rng As Object, out() As Long, b As Variant, wasProt As Boolean, wasMod As Boolean, um As Object
    Dim fld(0) As New com.sun.star.table.TableSortField, sd(1) As New com.sun.star.beans.PropertyValue
    If n <= 0 Then
        SortByCalc = Array()
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_PRINT)
    wasMod = gDoc.isModified()
    um = gDoc.getUndoManager()
    um.lock()
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    ReDim a(n - 1)
    For i = 0 To n - 1
        a(i) = Array(keys(i), i)
    Next i
    rng = sh.getCellRangeByPosition(LD_SORT_COL, 0, LD_SORT_COL + 1, n - 1)
    rng.setDataArray(a)
    fld(0).Field = 0
    fld(0).IsAscending = True
    fld(0).IsCaseSensitive = False
    sd(0).Name = "SortFields"
    sd(0).Value = fld()
    sd(1).Name = "ContainsHeader"
    sd(1).Value = False
    rng.sort(sd())
    b = sh.getCellRangeByPosition(LD_SORT_COL + 1, 0, LD_SORT_COL + 1, n - 1).getDataArray()
    ReDim out(n - 1)
    For i = 0 To n - 1
        out(i) = CLng(b(i)(0))
    Next i
    rng.clearContents(1 + 2 + 4 + 16)
    If wasProt Then sh.protect(PROTECT_PWD)
    um.unlock()
    If Not wasMod Then gDoc.setModified(False)
    SortByCalc = out
End Function

' the goods of an EI as a key: the name, the article and the unit (case and outer spaces do not matter)
Private Function LdGoodsKey(n As Long) As String
    LdGoodsKey = LCase(Trim(gLdName(n))) & " ¦ " & LCase(Trim(gLdArt(n))) & " ¦ " & LCase(Trim(gLdUnit(n)))
End Function

' a name to compare with another program: lower case, «ё» as «е», no quotes, single spaces
Function NameKey(ByVal s As String) As String
    s = LCase(Trim(s))
    s = Replace(s, "ё", "е")
    s = Replace(s, Chr(160), " ")
    s = Replace(s, Chr(9), " ")
    s = Replace(s, """", "")
    s = Replace(s, "«", "")
    s = Replace(s, "»", "")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NameKey = Trim(s)
End Function

' «Наименование (артикул)» of EI n
Private Function LdGoodsText(n As Long) As String
    LdGoodsText = gLdName(n) & IIf(Trim(gLdArt(n)) <> "", " (" & Trim(gLdArt(n)) & ")", "")
End Function

' a cell of a CSV file (the separator «;»): quoted when it holds «;», a quote or a line break
Private Function LdCsv(ByVal s As String) As String
    If InStr(s, ";") > 0 Or InStr(s, """") > 0 Or InStr(s, Chr(10)) > 0 Or InStr(s, Chr(13)) > 0 Then
        LdCsv = """" & Replace(s, """", """""") & """"
    Else
        LdCsv = s
    End If
End Function

' lines of a CSV file into «WMS_Документы» (UTF-8 with the mark Excel needs; «;»; a comma in the numbers). OK:<path> / ERR:…
Function CsvSave(sName As String, lines() As String, n As Long) As String
    Dim sFolder As String, url As String, k As Integer, a() As String, i As Long, sBase As String
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        CsvSave = "ERR:книга не сохранена в файл — файл некуда положить"
        Exit Function
    End If
    sFolder = gBaseDir & DOCS_DIR & "/"
    If Not gSfa.exists(sFolder) Then gSfa.createFolder(sFolder)
    sBase = WmsPrint.FileSafe(sName)
    url = sFolder & sBase & ".csv"
    k = 1
    Do While gSfa.exists(url)
        k = k + 1
        url = sFolder & sBase & " (" & k & ").csv"
    Loop
    ReDim a(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        a(i) = lines(i)
    Next i
    WriteTextFile(url, Chr(65279) & Join(a, Chr(13) & Chr(10)) & Chr(13) & Chr(10))
    CsvSave = "OK:" & ConvertFromURL(url)
    Exit Function
EH:
    CsvSave = "ERR:файл не записан: " & Error$
End Function

' «ДД.ММ.ГГГГ» of a form → a date; 0 when it is not one (msg)
Private Function LdFormDay(s As String, sWhat As String, ByRef msg As String) As Double
    Dim d As Double
    If WmsOrders.ParseDateTextAs(Trim(s), sWhat, d, msg) Then LdFormDay = Int(d)
End Function

' «09.2026», «9.2026», «сентябрь 2026» → year and month; False when not a month
Private Function LdParseMonth(ByVal s As String, ByRef y As Integer, ByRef m As Integer) As Boolean
    Dim a As Variant, names As Variant, i As Integer
    s = LCase(Trim(Replace(Replace(s, "/", "."), "-", ".")))
    names = Array("январ", "феврал", "март", "апрел", "ма", "июн", "июл", "август", "сентябр", "октябр", "ноябр", "декабр")
    If InStr(s, ".") > 0 Then
        a = Split(s, ".")
        If UBound(a) <> 1 Then Exit Function
        If Not IsNumeric(a(0)) Or Not IsNumeric(a(1)) Then Exit Function
        m = CInt(a(0))
        y = CInt(a(1))
    Else
        a = Split(s, " ")
        If UBound(a) < 1 Then Exit Function
        If Not IsNumeric(a(UBound(a))) Then Exit Function
        y = CInt(a(UBound(a)))
        For i = 11 To 0 Step -1
            If Left(a(0), Len(names(i))) = names(i) Then
                m = i + 1
                Exit For
            End If
        Next i
    End If
    If y < 100 Then y = y + 2000
    LdParseMonth = (m >= 1 And m <= 12 And y >= 2000 And y <= 2100)
End Function

Private Function LdMonthWord(m As Integer) As String
    LdMonthWord = Array("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")(m - 1)
End Function

' ================================================================ К7 the balance on a date

Sub BtnStockOnDate(Optional oEvent As Variant)
    Dim s As String, v(1) As String, d As Double, msg As String
    WmsInit()
    v(0) = Format(WmsOrders.Today() - 1, "DD.MM.YYYY")
    v(1) = "на страницу «Печать»"
    If Not WmsPrimeUi.UiForm("Остаток на дату", "Остаток всех ЕИ на конец выбранного дня: сегодняшний остаток минус всё, что было позже. " _
        & "Цена — из заказа; где цены нет, так и написано.", Array("Дата", "Куда"), v, Array("", "на страницу «Печать»;в файл CSV (Excel, 1С)"), "Показать") Then Exit Sub
    d = LdFormDay(v(0), "дата", msg)
    If d = 0 Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = StockOnDate(d, Left(v(1), 7) = "в файл ")
    LdReport(s)
End Sub

' the result of a report for people: the page is shown already; a file — its path
Private Sub LdReport(s As String)
    If Left(s, 4) = "ERR:" Then
        WmsUi.UiMessage(Mid(s, 5))
    ElseIf Left(s, 5) = "FILE:" Then
        WmsOrdersUi.UiInfo("Файл сохранён:" & Chr(10) & Mid(s, 6))
    End If
End Sub

' the balance of every EI at the end of day dt = the balance now minus every movement after dt. On «Печать» (by EI, the price
' and the sum, the total) or into a CSV file. OK:<document> / FILE:<path> / ERR:…
Function StockOnDate(dt As Double, bFile As Boolean) As String
    Dim aft() As Double, k As Long, n As Long, b As Double, rows() As Variant, nr As Long, sum As Double, nNoPrice As Long, nNeg As Long, p As Double
    Dim lines() As String, sh As Object, shown As Long, t0 As Long
    On Error GoTo EH
    t0 = GetSystemTicks()
    LedgerRead()
    ReDim aft(gLdMax)
    For k = 0 To gLdN - 1
        If gLdDate(k) > dt Then aft(gLdEi(k)) = aft(gLdEi(k)) + gLdQty(k)
    Next k
    ReDim rows(IIf(gLdMax > 0, gLdMax - 1, 0))
    ReDim lines(gLdMax + 1)
    lines(0) = "Остаток на " & Format(dt, "DD.MM.YYYY")
    lines(1) = "ЕИ;Наименование;Артикул;Ед.;Количество;Цена;Сумма;Место (сейчас)"
    For n = 1 To gLdMax
        b = WmsIssue.Round3(gLdNow(n) - aft(n))
        If b < -0.0000005 Then nNeg = nNeg + 1
        If b > 0.0000005 Then
            p = gLdPrice(n)
            If p >= 0 Then sum = sum + b * p Else nNoPrice = nNoPrice + 1
            rows(nr) = Array(nr + 1, gLdCanon(n), LdGoodsText(n), WmsPrint.PartsFirst(n, b, gLdUnit(n)), gLdPlace(n), IIf(p >= 0, WmsPrint.Money(p), "—"), _
                IIf(p >= 0, WmsPrint.Money(b * p), "—"))
            lines(nr + 2) = LdCsv(gLdCanon(n)) & ";" & LdCsv(gLdName(n)) & ";" & LdCsv(gLdArt(n)) & ";" & LdCsv(gLdUnit(n)) & ";" & WmsIssue.QtyText(b) & ";" _
                & IIf(p >= 0, WmsPrint.MoneyCsv(p), "") & ";" & IIf(p >= 0, WmsPrint.MoneyCsv(b * p), "") & ";" & LdCsv(gLdPlace(n))
            nr = nr + 1
        End If
    Next n
    If bFile Then
        StockOnDate = CsvSave("Остаток на " & Format(dt, "YYYY-MM-DD"), lines, nr + 2)
        If Left(StockOnDate, 3) = "OK:" Then StockOnDate = "FILE:" & Mid(StockOnDate, 4)
        gUiLastMsg = StockOnDate
        Exit Function
    End If
    shown = nr
    If shown > PR_MAX - 80 Then shown = PR_MAX - 80
    sh = WmsPrint.PrintBegin("ОСТАТОК НА " & Format(dt, "DD.MM.YYYY"), "Остаток на " & Format(dt, "YYYY-MM-DD"))
    WmsPrint.PrintLine(sh, "ЕИ с остатком: " & nr & " · на сумму " & WmsPrint.Money(sum) & IIf(nNoPrice > 0, " (без цены: " & nNoPrice & " ЕИ)", "") & ". Место — текущее.", True)
    If nNeg > 0 Then WmsPrint.PrintLine(sh, "Внимание: у " & nNeg & " ЕИ остаток на эту дату получился меньше нуля — движения не сходятся с «" & SH_STOCK _
        & "»; откройте их карточки («Поиск») или «Самопроверку».")
    WmsPrint.PrintTable(sh, Array("№", "ЕИ", "Наименование (артикул)", "Количество", "Место", "Цена", "Сумма"), rows, shown, "0,5,6")
    If shown < nr Then WmsPrint.PrintLine(sh, "Показаны первые " & shown & " ЕИ из " & nr & " — полный список: «Остаток на дату» → «в файл CSV».", True)
    WmsPrint.PrintLine(sh, "Итого на сумму: " & WmsPrint.Money(sum), True)
    WmsPrint.PrintSign(sh, Array("Кладовщик", "Проверил"))
    StockOnDate = WmsPrint.PrintDone(sh)
    Exit Function
EH:
    StockOnDate = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = StockOnDate
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' ================================================================ К8 the turnover sheet

Sub BtnTurnover(Optional oEvent As Variant)
    Dim s As String, v(2) As String, d0 As Double, d1 As Double, msg As String, t As Double
    WmsInit()
    t = WmsOrders.Today()
    v(0) = Format(DateSerial(Year(t), Month(t), 1), "DD.MM.YYYY")
    v(1) = Format(t, "DD.MM.YYYY")
    v(2) = "на страницу «Печать»"
    If Not WmsPrimeUi.UiForm("Оборотная ведомость", "За период по товарам: остаток на начало, приход, выдано, возвращено, списано, остаток на конец — " _
        & "в единицах и в рублях (по цене из заказа).", Array("С", "По", "Куда"), v, Array("", "", "на страницу «Печать»;в файл CSV (Excel, 1С)"), "Показать") Then Exit Sub
    d0 = LdFormDay(v(0), "дата начала", msg)
    If d0 > 0 Then d1 = LdFormDay(v(1), "дата конца", msg)
    If d0 = 0 Or d1 = 0 Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = Turnover(d0, d1, Left(v(2), 7) = "в файл ")
    LdReport(s)
End Sub

' the turnover of [d0, d1] by goods (name, article, unit): the start, the receipts (with the surplus of a count and the splits),
' the issues, the returns, the write-offs (with the returns to the supplier and the shortage of a count), the end; the money by
' the price of every EI. Start + receipts − issues + returns − write-offs = end, by construction. OK:… / FILE:… / ERR:…
Function Turnover(d0 As Double, d1 As Double, bFile As Boolean) As String
    Dim res As Variant
    On Error GoTo EH
    If d1 < d0 Then
        Turnover = "ERR:дата конца раньше даты начала"
        gUiLastMsg = Turnover
        Exit Function
    End If
    res = LdTurnoverData(d0, d1)
    Turnover = LdTurnoverOut(res, d0, d1, bFile, "Оборотная ведомость " & Format(d0, "YYYY-MM-DD") & " — " & Format(d1, "YYYY-MM-DD"))
    Exit Function
EH:
    Turnover = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = Turnover
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' the rows of the turnover: Array(n, keys of goods (text), q(6 per goods), money(6 per goods), totals q, totals money, EIs per goods)
Private Function LdTurnoverData(d0 As Double, d1 As Double) As Variant
    Dim a0() As Double, a1() As Double, bk() As Double, k As Long, n As Long, j As Integer, e As Long, q As Double, keys() As String, idx() As Long
    Dim ne As Long, ord As Variant, g As Long, i As Long, prevKey As String, gq() As Double, gm() As Double, gN() As Long, gFirst() As Long, tq(5) As Double
    Dim tm(5) As Double, start As Double, fin As Double, p As Double, has As Boolean, nNoPrice As Long
    LedgerRead()
    ReDim a0(gLdMax)
    ReDim a1(gLdMax)
    ReDim bk(gLdMax, 3)
    For k = 0 To gLdN - 1
        e = gLdEi(k)
        q = gLdQty(k)
        If gLdDate(k) >= d0 Then a0(e) = a0(e) + q
        If gLdDate(k) > d1 Then a1(e) = a1(e) + q
        If gLdDate(k) >= d0 And gLdDate(k) <= d1 Then
            Select Case gLdKind(k)
            Case LK_RCV, LK_SPEC, LK_SPLIT
                bk(e, 0) = bk(e, 0) + q
            Case LK_ISS, LK_LEG
                bk(e, 1) = bk(e, 1) - q
            Case LK_RET
                bk(e, 2) = bk(e, 2) + q
            Case LK_WOFF, LK_SUP
                bk(e, 3) = bk(e, 3) - q
            Case LK_INV, LK_KIT
                If q > 0 Then bk(e, 0) = bk(e, 0) + q Else bk(e, 3) = bk(e, 3) - q
            End Select
        End If
    Next k
    ' the EIs with anything in the period
    ReDim keys(IIf(gLdMax > 0, gLdMax - 1, 0))
    ReDim idx(IIf(gLdMax > 0, gLdMax - 1, 0))
    For n = 1 To gLdMax
        start = WmsIssue.Round3(gLdNow(n) - a0(n))
        fin = WmsIssue.Round3(gLdNow(n) - a1(n))
        has = (Abs(start) > 0.0000005 Or Abs(fin) > 0.0000005)
        For j = 0 To 3
            If Abs(bk(n, j)) > 0.0000005 Then has = True
        Next j
        If has Then
            keys(ne) = LdGoodsKey(n)
            idx(ne) = n
            ne = ne + 1
        End If
    Next n
    ord = SortByCalc(keys, ne)
    ReDim gq(IIf(ne > 0, ne - 1, 0), 5)
    ReDim gm(IIf(ne > 0, ne - 1, 0), 5)
    ReDim gN(IIf(ne > 0, ne - 1, 0))
    ReDim gFirst(IIf(ne > 0, ne - 1, 0))
    g = -1
    prevKey = Chr(1)
    For i = 0 To ne - 1
        n = idx(ord(i))
        If LCase(keys(ord(i))) <> prevKey Then
            g = g + 1
            prevKey = LCase(keys(ord(i)))
            gFirst(g) = n
        End If
        gN(g) = gN(g) + 1
        start = WmsIssue.Round3(gLdNow(n) - a0(n))
        fin = WmsIssue.Round3(gLdNow(n) - a1(n))
        p = gLdPrice(n)
        If p < 0 Then nNoPrice = nNoPrice + 1
        For j = 0 To 5
            Select Case j
            Case 0
                q = start
            Case 5
                q = fin
            Case Else
                q = bk(n, j - 1)
            End Select
            gq(g, j) = gq(g, j) + q
            tq(j) = tq(j) + q
            If p >= 0 Then
                gm(g, j) = gm(g, j) + q * p
                tm(j) = tm(j) + q * p
            End If
        Next j
    Next i
    LdTurnoverData = Array(g + 1, gFirst, gq, gm, gN, tq, tm, nNoPrice, ne)
End Function

' the turnover on «Печать» (quantities by goods, the money in a table of totals) or into a CSV (both, by goods)
Private Function LdTurnoverOut(res As Variant, d0 As Double, d1 As Double, bFile As Boolean, sName As String) As String
    Dim ng As Long, gFirst As Variant, gq As Variant, gm As Variant, gN As Variant, tq As Variant, tm As Variant, nNoPrice As Long, rows() As Variant
    Dim g As Long, n As Long, lines() As String, j As Integer, s As String, sh As Object, shown As Long, heads As Variant
    ng = res(0)
    gFirst = res(1)
    gq = res(2)
    gm = res(3)
    gN = res(4)
    tq = res(5)
    tm = res(6)
    nNoPrice = res(7)
    heads = Array("Начало", "Приход", "Выдано", "Возвращено", "Списано", "Конец")
    If bFile Then
        ReDim lines(ng + 2)
        lines(0) = "Оборотная ведомость " & Format(d0, "DD.MM.YYYY") & " — " & Format(d1, "DD.MM.YYYY")
        lines(1) = "Товар;Артикул;Ед.;ЕИ;Начало;Приход;Выдано;Возвращено;Списано;Конец;Начало, руб;Приход, руб;Выдано, руб;Возвращено, руб;Списано, руб;Конец, руб"
        For g = 0 To ng - 1
            n = gFirst(g)
            s = LdCsv(gLdName(n)) & ";" & LdCsv(gLdArt(n)) & ";" & LdCsv(gLdUnit(n)) & ";" & gN(g)
            For j = 0 To 5
                s = s & ";" & WmsIssue.QtyText(gq(g, j))
            Next j
            For j = 0 To 5
                s = s & ";" & WmsPrint.MoneyCsv(gm(g, j))
            Next j
            lines(g + 2) = s
        Next g
        s = "Итого;;;"
        For j = 0 To 5
            s = s & ";"
        Next j
        For j = 0 To 5
            s = s & ";" & WmsPrint.MoneyCsv(tm(j))
        Next j
        lines(ng + 2) = s
        LdTurnoverOut = CsvSave(sName, lines, ng + 3)
        If Left(LdTurnoverOut, 3) = "OK:" Then LdTurnoverOut = "FILE:" & Mid(LdTurnoverOut, 4)
        gUiLastMsg = LdTurnoverOut
        Exit Function
    End If
    ReDim rows(IIf(ng > 0, ng - 1, 0))
    For g = 0 To ng - 1
        n = gFirst(g)
        rows(g) = Array(LdGoodsText(n) & IIf(gLdUnit(n) <> "", ", " & gLdUnit(n), ""), LdQ0(gq(g, 0)), LdQ0(gq(g, 1)), LdQ0(gq(g, 2)), LdQ0(gq(g, 3)), LdQ0(gq(g, 4)), _
            LdQ0(gq(g, 5)))
    Next g
    shown = ng
    If shown > PR_MAX - 100 Then shown = PR_MAX - 100
    sh = WmsPrint.PrintBegin("ОБОРОТНАЯ ВЕДОМОСТЬ за " & Format(d0, "DD.MM.YYYY") & " — " & Format(d1, "DD.MM.YYYY"), sName)
    WmsPrint.PrintLine(sh, "Товаров: " & ng & " · ЕИ: " & res(8) & IIf(nNoPrice > 0, " · без цены: " & nNoPrice & " ЕИ (в рублях не учтены)", "") _
        & ". Начало + приход − выдано + возвращено − списано = конец.", True)
    WmsPrint.PrintLine(sh, "Приход — с излишками пересчёта и собранными комплектами; списано — с возвратами поставщику, недостачами пересчёта и " _
        & "позициями, ушедшими в комплекты.")
    WmsPrint.PrintTable(sh, Array("Товар, ед.", heads(0), heads(1), heads(2), heads(3), heads(4), heads(5)), rows, shown, "1,2,3,4,5,6")
    If shown < ng Then WmsPrint.PrintLine(sh, "Показаны первые " & shown & " товаров из " & ng & " — полностью: «Оборотная ведомость» → «в файл CSV».", True)
    WmsPrint.PrintLine(sh, "В рублях:", True)
    WmsPrint.PrintTable(sh, Array("", heads(0), heads(1), heads(2), heads(3), heads(4), heads(5)), Array(Array("Итого, руб.", WmsPrint.Money(tm(0)), _
        WmsPrint.Money(tm(1)), WmsPrint.Money(tm(2)), WmsPrint.Money(tm(3)), WmsPrint.Money(tm(4)), WmsPrint.Money(tm(5)))), 1, "1,2,3,4,5,6")
    WmsPrint.PrintSign(sh, Array("Кладовщик", "Бухгалтер"))
    LdTurnoverOut = WmsPrint.PrintDone(sh)
End Function

' a quantity of a table: «» for zero
Private Function LdQ0(q As Double) As String
    If Abs(q) > 0.0000005 Then LdQ0 = WmsIssue.QtyText(q)
End Function

' ================================================================ К9 the movements for the accounting

Sub BtnExport(Optional oEvent As Variant)
    Dim s As String, v(1) As String, d0 As Double, d1 As Double, msg As String, t As Double
    WmsInit()
    t = WmsOrders.Today()
    v(0) = Format(DateSerial(Year(t), Month(t), 1), "DD.MM.YYYY")
    v(1) = Format(t, "DD.MM.YYYY")
    If Not WmsPrimeUi.UiForm("Выгрузка для бухгалтерии", "Файл CSV (открывается в Excel и загружается в 1С): каждое движение за период — дата, " _
        & "операция, документ, ЕИ и товар, приход или расход, цена, сумма, контрагент.", Array("С", "По"), v, Array("", ""), "Сохранить файл") Then Exit Sub
    d0 = LdFormDay(v(0), "дата начала", msg)
    If d0 > 0 Then d1 = LdFormDay(v(1), "дата конца", msg)
    If d0 = 0 Or d1 = 0 Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = ExportMoves(d0, d1)
    LdReport(s)
End Sub

' every movement of [d0, d1] by date into a CSV of «WMS_Документы». FILE:<path> (and the count) / ERR:…
Function ExportMoves(d0 As Double, d1 As Double) As String
    Dim k As Long, keys() As Double, idx() As Long, m As Long, ord As Variant, i As Long, lines() As String, n As Long, q As Double, p As Double
    Dim sIn As Double, sOut As Double, res As String
    On Error GoTo EH
    If d1 < d0 Then
        ExportMoves = "ERR:дата конца раньше даты начала"
        gUiLastMsg = ExportMoves
        Exit Function
    End If
    LedgerRead()
    ReDim keys(IIf(gLdN > 0, gLdN - 1, 0))
    ReDim idx(IIf(gLdN > 0, gLdN - 1, 0))
    For k = 0 To gLdN - 1
        If gLdDate(k) >= d0 And gLdDate(k) <= d1 Then
            keys(m) = gLdDate(k) * 10000000 + k
            idx(m) = k
            m = m + 1
        End If
    Next k
    ord = SortByCalc(keys, m)
    ReDim lines(m + 2)
    lines(0) = "Движения склада " & Format(d0, "DD.MM.YYYY") & " — " & Format(d1, "DD.MM.YYYY")
    lines(1) = "Дата;Операция;Документ;ЕИ;Наименование;Артикул;Ед.;Приход;Расход;Цена;Сумма;Контрагент / получатель / причина"
    For i = 0 To m - 1
        k = idx(ord(i))
        n = gLdEi(k)
        q = gLdQty(k)
        p = gLdPrice(n)
        lines(i + 2) = Format(gLdDate(k), "DD.MM.YYYY") & ";" & LdCsv(gLdOp(k)) & ";" & LdCsv(gLdDoc(k)) & ";" & LdCsv(gLdCanon(n)) & ";" & LdCsv(gLdName(n)) & ";" _
            & LdCsv(gLdArt(n)) & ";" & LdCsv(gLdUnit(n)) & ";" & IIf(q > 0, WmsIssue.QtyText(q), "") & ";" & IIf(q < 0, WmsIssue.QtyText(-q), "") & ";" _
            & IIf(p >= 0, WmsPrint.MoneyCsv(p), "") & ";" & IIf(p >= 0, WmsPrint.MoneyCsv(Abs(q) * p), "") & ";" & LdCsv(gLdWho(k))
        If p >= 0 Then
            If q > 0 Then sIn = sIn + q * p Else sOut = sOut - q * p
        End If
    Next i
    lines(m + 2) = "Итого;;;;;;;;;;приход " & WmsPrint.MoneyCsv(sIn) & " / расход " & WmsPrint.MoneyCsv(sOut) & ";"
    res = CsvSave("Движения " & Format(d0, "YYYY-MM-DD") & " — " & Format(d1, "YYYY-MM-DD"), lines, m + 3)
    If Left(res, 3) = "OK:" Then res = "FILE:" & Mid(res, 4) & Chr(10) & "Движений: " & m & " · приход на " & WmsPrint.Money(sIn) & " · расход на " & WmsPrint.Money(sOut)
    ExportMoves = res
    gUiLastMsg = res
    Exit Function
EH:
    ExportMoves = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = ExportMoves
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' ================================================================ К10 the comparison with 1С

Sub BtnCompare1C(Optional oEvent As Variant)
    Dim fp As Object, s As String, k As Integer, url As String
    WmsInit()
    On Error GoTo EH
    If gLdFileTest <> "" Then
        url = gLdFileTest
    Else
        fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
        fp.initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE))
        fp.setTitle("Остатки из 1С: файл CSV, XLSX, XLS или ODS (колонки «Наименование» и «Остаток» или «Количество»)")
        fp.appendFilter("Таблицы и CSV (*.csv; *.txt; *.xlsx; *.xls; *.ods)", "*.csv;*.txt;*.xlsx;*.xls;*.ods")
        If fp.execute() <> 1 Then Exit Sub
        url = fp.getFiles()(0)
    End If
    k = WmsPrimeUi.UiButtons("Сверка с 1С", "Как сопоставлять товары 1С и WMS?", Array("По наименованию", "По артикулу", "Отмена"), Array(1, 2, 0))
    If k <= 0 Then Exit Sub
    s = Compare1C(url, k = 2)
    LdReport(s)
    Exit Sub
EH:
    WmsUi.UiMessage("Сверка не выполнена: " & Error$)
End Sub

' the balances of a 1С file against the balances of WMS by goods (the name, or the article): only in 1С, only in WMS, the
' quantity differs. Nothing is posted. OK:… / ERR:…
Function Compare1C(url As String, bByArt As Boolean) As String
    Dim grid As Variant, nr As Long, msg As String, h As Long, cName As Integer, cArt As Integer, cQty As Integer, cUnit As Integer, i As Long
    Dim keys() As String, side() As Integer, qty() As Double, nm() As String, un() As String, m As Long, n As Long, ord As Variant, kk As String
    Dim rows() As Variant, nd As Long, nEq As Long, n1 As Long, nW As Long, nDiff As Long, j As Long, q1 As Double, qW As Double, s1 As String
    Dim sW As String, u1 As String, uW As String, nm1 As String, nmW As String, sh As Object, cnt As Long, x As Double, shown As Long, nIn1 As Long
    On Error GoTo EH
    If Not LdRead1C(url, grid, nr, msg) Then
        Compare1C = "ERR:" & msg
        gUiLastMsg = Compare1C
        Exit Function
    End If
    If Not LdFind1CColumns(grid, nr, h, cName, cArt, cQty, cUnit) Then
        Compare1C = "ERR:в файле не найдены колонки «Наименование» (или «Номенклатура», «Товар») и «Остаток» (или «Количество») в первых 30 строках"
        gUiLastMsg = Compare1C
        Exit Function
    End If
    If bByArt And cArt < 0 Then
        Compare1C = "ERR:в файле нет колонки «Артикул» — сравните по наименованию"
        gUiLastMsg = Compare1C
        Exit Function
    End If
    LedgerRead()
    ReDim keys(nr + gLdMax)
    ReDim side(nr + gLdMax)
    ReDim qty(nr + gLdMax)
    ReDim nm(nr + gLdMax)
    ReDim un(nr + gLdMax)
    For i = h + 1 To nr - 1
        nm1 = Trim(LdCellText(grid(i), cName))
        If nm1 <> "" And LCase(Left(nm1, 5)) <> "итого" And LCase(Left(nm1, 4)) <> "всего" Then
            If LdNum1C(LdCellOf(grid(i), cQty), x) Then
                kk = IIf(bByArt, NameKey(LdCellText(grid(i), cArt)), NameKey(nm1))
                If kk <> "" Then
                    keys(m) = kk
                    side(m) = 1
                    qty(m) = x
                    nm(m) = nm1 & IIf(cArt >= 0, IIf(Trim(LdCellText(grid(i), cArt)) <> "", " (" & Trim(LdCellText(grid(i), cArt)) & ")", ""), "")
                    un(m) = IIf(cUnit >= 0, Trim(LdCellText(grid(i), cUnit)), "")
                    m = m + 1
                    nIn1 = nIn1 + 1
                End If
            End If
        End If
    Next i
    If nIn1 = 0 Then
        Compare1C = "ERR:в файле нет строк с наименованием и числом в колонке остатка (под строкой заголовков " & (h + 1) & ")"
        gUiLastMsg = Compare1C
        Exit Function
    End If
    For n = 1 To gLdMax
        If gLdNow(n) > 0.0000005 Then
            kk = IIf(bByArt, NameKey(gLdArt(n)), NameKey(gLdName(n)))
            If kk <> "" Then
                keys(m) = kk
                side(m) = 2
                qty(m) = gLdNow(n)
                nm(m) = LdGoodsText(n)
                un(m) = gLdUnit(n)
                m = m + 1
            End If
        End If
    Next n
    ord = SortByCalc(keys, m)
    ReDim rows(m)
    i = 0
    Do While i < m
        kk = LCase(keys(ord(i)))
        q1 = 0
        qW = 0
        s1 = ""
        sW = ""
        u1 = ""
        uW = ""
        cnt = 0
        j = i
        Do While j < m
            If LCase(keys(ord(j))) <> kk Then Exit Do
            If side(ord(j)) = 1 Then
                q1 = q1 + qty(ord(j))
                If s1 = "" Then s1 = nm(ord(j))
                If u1 = "" Then u1 = un(ord(j))
            Else
                qW = qW + qty(ord(j))
                cnt = cnt + 1
                If sW = "" Then sW = nm(ord(j))
                If uW = "" Then uW = un(ord(j))
            End If
            j = j + 1
        Loop
        q1 = WmsIssue.Round3(q1)
        qW = WmsIssue.Round3(qW)
        If s1 <> "" And sW = "" Then
            n1 = n1 + 1
            rows(nd) = Array(nd + 1, s1, u1, WmsIssue.QtyText(q1), "—", "", "нет в WMS")
            nd = nd + 1
        ElseIf s1 = "" And sW <> "" Then
            nW = nW + 1
            rows(nd) = Array(nd + 1, sW, uW, "—", WmsIssue.QtyText(qW), "", "нет в 1С (ЕИ: " & cnt & ")")
            nd = nd + 1
        ElseIf Abs(q1 - qW) > 0.0000005 Then
            nDiff = nDiff + 1
            rows(nd) = Array(nd + 1, sW, uW, WmsIssue.QtyText(q1), WmsIssue.QtyText(qW), IIf(qW > q1, "+", "") & WmsIssue.QtyText(WmsIssue.Round3(qW - q1)), _
                "ЕИ: " & cnt & IIf(u1 <> "" And LCase(Replace(u1, ".", "")) <> LCase(Replace(uW, ".", "")), " · ед. в 1С: " & u1, ""))
            nd = nd + 1
        Else
            nEq = nEq + 1
        End If
        i = j
    Loop
    shown = nd
    If shown > PR_MAX - 80 Then shown = PR_MAX - 80
    sh = WmsPrint.PrintBegin("СВЕРКА ОСТАТКОВ С 1С от " & Format(WmsOrders.Today(), "DD.MM.YYYY"), "Сверка с 1С " & Format(WmsOrders.Today(), "YYYY-MM-DD"))
    WmsPrint.PrintLine(sh, "Файл 1С: " & ConvertFromURL(url))
    WmsPrint.PrintLine(sh, "Сопоставлено " & IIf(bByArt, "по артикулу", "по наименованию") & ". Строк в файле: " & nIn1 & ". Остатки WMS — сейчас («" & SH_STOCK & "»).")
    WmsPrint.PrintLine(sh, "Совпало: " & nEq & " · количество расходится: " & nDiff & " · есть только в 1С: " & n1 & " · есть только в WMS: " & nW _
        & ". Ничего не проведено — расхождения разберите пересчётом («" & SH_ADJUST & "» → «Пересчитать место…»).", True)
    If nd > 0 Then
        WmsPrint.PrintTable(sh, Array("№", "Товар", "Ед.", "В 1С", "В WMS", "Разница", "Примечание"), rows, shown, "0,3,4,5")
        If shown < nd Then WmsPrint.PrintLine(sh, "Показаны первые " & shown & " расхождений из " & nd & ".", True)
    Else
        WmsPrint.PrintLine(sh, "Расхождений нет.", True)
    End If
    WmsPrint.PrintSign(sh, Array("Кладовщик", "Бухгалтер"))
    Compare1C = WmsPrint.PrintDone(sh)
    Exit Function
EH:
    Compare1C = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = Compare1C
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' the cells of a file as rows of texts and numbers: a CSV / TXT (UTF-8 or Windows-1251, «;», tab or «,»), or the first sheet of
' a table file (opened hidden, read-only, without its macros)
Private Function LdRead1C(url As String, ByRef grid As Variant, ByRef nr As Long, ByRef msg As String) As Boolean
    Dim ext As String, txt As String, a As Variant, i As Long, sep As String, src As Object, sh As Object, cur As Object, d As Variant, rows() As Variant
    Dim args(2) As New com.sun.star.beans.PropertyValue, c1 As Long, c2 As Long, c3 As Long
    On Error GoTo EH
    If Not gSfa.exists(url) Then
        msg = "файл не найден: " & ConvertFromURL(url)
        Exit Function
    End If
    ext = LCase(Mid(url, LdLastPos(url, ".") + 1))
    If ext = "csv" Or ext = "txt" Then
        txt = LdReadAs(url, "UTF-8")
        If InStr(txt, Chr(65533)) > 0 Then txt = LdReadAs(url, "windows-1251")
        If Left(txt, 1) = Chr(65279) Then txt = Mid(txt, 2)
        a = Split(Replace(txt, Chr(13), ""), Chr(10))
        For i = 0 To UBound(a)
            If i > 30 Then Exit For
            c1 = c1 + UBound(Split(a(i), ";"))
            c2 = c2 + UBound(Split(a(i), Chr(9)))
            c3 = c3 + UBound(Split(a(i), ","))
        Next i
        sep = ";"
        If c2 > c1 And c2 >= c3 Then sep = Chr(9)
        If c3 > c1 And c3 > c2 Then sep = ","
        ReDim rows(UBound(a))
        For i = 0 To UBound(a)
            rows(i) = LdSplitCsv(a(i), sep)
        Next i
        grid = rows
        nr = UBound(a) + 1
        LdRead1C = True
        Exit Function
    End If
    args(0).Name = "Hidden"
    args(0).Value = True
    args(1).Name = "ReadOnly"
    args(1).Value = True
    args(2).Name = "MacroExecutionMode"
    args(2).Value = 0
    src = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    If IsNull(src) Then
        msg = "файл не открылся: " & ConvertFromURL(url)
        Exit Function
    End If
    sh = src.Sheets.getByIndex(0)
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    d = sh.getCellRangeByPosition(0, 0, cur.getRangeAddress().EndColumn, cur.getRangeAddress().EndRow).getDataArray()
    src.close(True)
    grid = d
    nr = UBound(d) + 1
    LdRead1C = True
    Exit Function
EH:
    msg = "файл не прочитан: " & Error$
    On Error Resume Next
    If Not IsNull(src) And Not IsEmpty(src) Then src.close(True)
End Function

' the position of the last ch in s (0 — none)
Private Function LdLastPos(s As String, ch As String) As Long
    Dim p As Long
    p = InStr(s, ch)
    Do While p > 0
        LdLastPos = p
        p = InStr(p + 1, s, ch)
    Loop
End Function

Private Function LdReadAs(url As String, sEnc As String) As String
    Dim tis As Object
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(gSfa.openFileRead(url))
    tis.setEncoding(sEnc)
    LdReadAs = tis.readString(Array(), False)
    tis.closeInput()
End Function

' a line of a CSV into its cells (quotes «"…"» kept together, «""» — one quote)
Private Function LdSplitCsv(s As String, sep As String) As Variant
    Dim out() As String, n As Integer, i As Long, ch As String, cur As String, q As Boolean
    ReDim out(63)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If q Then
            If ch = """" Then
                If Mid(s, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    q = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            q = True
        ElseIf ch = sep Then
            If n <= 63 Then out(n) = cur
            n = n + 1
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    If n <= 63 Then out(n) = cur
    If n > 63 Then n = 63
    ReDim Preserve out(n)
    LdSplitCsv = out
End Function

Private Function LdCellOf(row As Variant, c As Integer) As Variant
    If c < 0 Or c > UBound(row) Then
        LdCellOf = ""
    Else
        LdCellOf = row(c)
    End If
End Function

Private Function LdCellText(row As Variant, c As Integer) As String
    LdCellText = CStr(LdCellOf(row, c))
End Function

' a number of 1С: a number cell or «1 250,000» (spaces, a comma); False when it is not a number
Private Function LdNum1C(v As Variant, ByRef x As Double) As Boolean
    Dim s As String
    If VarType(v) = 5 Then
        x = v
        LdNum1C = True
        Exit Function
    End If
    s = Replace(Replace(Replace(Trim(CStr(v)), " ", ""), Chr(160), ""), Chr(8239), "")
    s = Replace(s, ",", ".")
    If s = "" Then Exit Function
    If Not IsNumeric(Replace(s, ".", "0")) Then Exit Function
    If Len(s) - Len(Replace(s, ".", "")) > 1 Then Exit Function
    x = Val(s)
    LdNum1C = True
End Function

' the row of the headers (the first of 30 with a name and a quantity column) and the columns; -1 when absent
Private Function LdFind1CColumns(grid As Variant, nr As Long, ByRef h As Long, ByRef cName As Integer, ByRef cArt As Integer, ByRef cQty As Integer, _
    ByRef cUnit As Integer) As Boolean
    Dim i As Long, c As Integer, t As String, row As Variant, qBest As Integer, qRank As Integer, rk As Integer
    For i = 0 To nr - 1
        If i > 30 Then Exit For
        row = grid(i)
        cName = -1
        cArt = -1
        cQty = -1
        cUnit = -1
        qRank = 0
        For c = 0 To UBound(row)
            t = LCase(Trim(CStr(row(c))))
            If t <> "" Then
                If cName < 0 And (InStr(t, "номенклатур") > 0 Or InStr(t, "наименован") > 0 Or t = "товар" Or Left(t, 6) = "товар " Or t = "товары") Then
                    cName = c
                ElseIf cArt < 0 And InStr(t, "артикул") > 0 Then
                    cArt = c
                ElseIf cUnit < 0 And (Left(t, 2) = "ед" Or InStr(t, "единиц") > 0) Then
                    cUnit = c
                Else
                    rk = 0
                    If InStr(t, "конечн") > 0 And InStr(t, "остат") > 0 Then
                        rk = 4
                    ElseIf InStr(t, "остат") > 0 And InStr(t, "начал") = 0 Then
                        rk = 3
                    ElseIf InStr(t, "количеств") > 0 Or InStr(t, "кол-во") > 0 Or t = "кол." Then
                        rk = 2
                    End If
                    If rk > qRank Then
                        qRank = rk
                        cQty = c
                    End If
                End If
            End If
        Next c
        If cName >= 0 And cQty >= 0 Then
            h = i
            LdFind1CColumns = True
            Exit Function
        End If
    Next i
End Function

' ================================================================ К13 the closing of a month

' the last day of the closed period (the latest closing still in force); 0 — no month is closed
Function ClosedTo() As Double
    Dim sh As Object, d As Variant, i As Long, x As Double
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_CLOSE) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_CLOSE)
    ' every posting asks: two years of closings are read at once, the rest only when they exist
    d = sh.getCellRangeByPosition(0, 1, CL_STATE, 24).getDataArray()
    If CStr(d(23)(CL_MONTH)) <> "" Then d = sh.getCellRangeByPosition(0, 1, CL_STATE, CL_MAX).getDataArray()
    For i = 0 To UBound(d)
        If CStr(d(i)(CL_MONTH)) = "" Then Exit For
        If CStr(d(i)(CL_STATE)) = CL_CLOSED And VarType(d(i)(CL_TO)) = 5 Then
            If d(i)(CL_TO) > x Then x = d(i)(CL_TO)
        End If
    Next i
    ClosedTo = x
EH:
End Function

' ApplyOperation (a live operation): a date of the plan (DATE, OLD_DATE) inside the closed period — the warning and the reason
' (once per block of rows), the fields BACKDATED and BACKWHY added to the journal line. "" — post; SKIP:… — not posted
Function ClosedAsk() As String
    Dim ct As Double, i As Long, f As String, d As Double, dMin As Double, sWhy As String, sText As String
    Select Case gPlanType
    Case "MONTH_CLOSE", "MONTH_OPEN", "MOVE", "MOVE_FIX", "MOVE_DEL", "EI_PACK", "ISSUE_DOC", "START", "ABANDON"
        Exit Function
    End Select
    ct = ClosedTo()
    If ct <= 0 Then Exit Function
    For i = 0 To gPlanFieldN - 1
        f = gPlanFields(i)
        d = 0
        If Left(f, 5) = "DATE=" Then
            d = LdIsoDay(Mid(f, 6))
        ElseIf Left(f, 9) = "OLD_DATE=" Then
            d = LdIsoDay(Mid(f, 10))
        End If
        If d > 0 And d <= ct Then
            If dMin = 0 Or d < dMin Then dMin = d
        End If
    Next i
    If dMin = 0 Then Exit Function
    sWhy = WmsAsk.MemoValue("CLOSED")
    If sWhy = "" Then
        sText = "Период по " & Format(ct, "DD.MM.YYYY") & " закрыт: остатки и оборотная ведомость на конец месяца зафиксированы." & Chr(10) _
            & "Операция «" & WmsView.OpName(gPlanType) & "» с датой " & Format(dMin, "DD.MM.YYYY") & " изменит закрытый месяц." & Chr(10) & Chr(10) _
            & "Если это действительно нужно — напишите причину: она останется в журнале. Если дата набрана с ошибкой — закройте окно и исправьте её."
        If Not WmsPrimeUi.ReasonAsk("Операция задним числом", sText, "Провести задним числом", sWhy) Then
            ClosedAsk = "SKIP:не проведено — дата " & Format(dMin, "DD.MM.YYYY") & " в закрытом периоде (по " & Format(ct, "DD.MM.YYYY") & ")"
            gUiLastMsg = ClosedAsk
            Exit Function
        End If
        sWhy = Trim(sWhy)
        If sWhy = "" Then
            ClosedAsk = "SKIP:не проведено — для операции задним числом нужна причина"
            gUiLastMsg = ClosedAsk
            Exit Function
        End If
        WmsAsk.MemoSet("CLOSED", sWhy)
    End If
    PlanField("BACKDATED", Format(ct, "YYYY-MM-DD"))
    PlanField("BACKWHY", sWhy)
End Function

' «2026-09-30» → a date; 0 when not one
Private Function LdIsoDay(s As String) As Double
    On Error GoTo EH
    If Len(s) >= 10 And Mid(s, 5, 1) = "-" Then LdIsoDay = CDbl(DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 6, 2)), CInt(Mid(s, 9, 2))))
EH:
End Function

Sub BtnCloseMonth(Optional oEvent As Variant)
    Dim s As String, v(0) As String, y As Integer, m As Integer, t As Double
    WmsInit()
    t = DateSerial(Year(WmsOrders.Today()), Month(WmsOrders.Today()), 1) - 1
    v(0) = Format(t, "MM.YYYY")
    If Not WmsPrimeUi.UiForm("Закрыть месяц", "Остатки и оборотная ведомость на конец месяца сохраняются в файлы и фиксируются в журнале. " _
        & "Потом операция с датой в закрытом месяце проводится только после предупреждения и с причиной.", Array("Месяц (ММ.ГГГГ)"), v, Array(""), _
        "Закрыть месяц") Then Exit Sub
    If Not LdParseMonth(v(0), y, m) Then
        WmsUi.UiMessage("«" & v(0) & "» — не месяц: напишите, например, «09.2026».")
        Exit Sub
    End If
    s = CloseMonth(y, m)
    If Left(s, 3) = "OK:" Then WmsOrdersUi.UiInfo(Mid(s, 4)) Else LdReport(s)
End Sub

' the closing of month m of year y: the turnover of the month and the balances at its end into two files («Закрытие …»), the
' turnover on «Печать», then the operation MONTH_CLOSE (the row of _CLOSE, the journal). OK:… / ERR:… / SKIP:…
Function CloseMonth(y As Integer, m As Integer) As String
    Dim d0 As Double, d1 As Double, ct As Double, why As String, res As Variant, f1 As String, f2 As String, k As Long, sh As Object, d As Variant
    Dim sMonth As String, r As String, qEnd As Double, sumEnd As Double
    On Error GoTo EH
    d0 = CDbl(DateSerial(y, m, 1))
    d1 = CDbl(DateSerial(y, m + 1, 1)) - 1
    sMonth = Format(y, "0000") & "-" & Format(m, "00")
    If d1 >= WmsOrders.Today() Then
        CloseMonth = "ERR:" & LdMonthWord(m) & " " & y & " ещё не закончился — закрыть можно с " & Format(d1 + 1, "DD.MM.YYYY")
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    ct = ClosedTo()
    If d1 <= ct Then
        CloseMonth = "ERR:" & LdMonthWord(m) & " " & y & " уже закрыт (закрыто по " & Format(ct, "DD.MM.YYYY") & ")"
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        CloseMonth = "ERR:" & Mid(why, InStr(why, ":") + 1)
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    If Not WmsUi.UiConfirm("Закрыть " & LdMonthWord(m) & " " & y & "?" & Chr(10) & Chr(10) & "Остатки на " & Format(d1, "DD.MM.YYYY") _
        & " и оборотная ведомость месяца будут сохранены в папку «" & DOCS_DIR & "» и зафиксированы в журнале. Операции с датой по " _
        & Format(d1, "DD.MM.YYYY") & " потом проводятся только с причиной.") Then
        CloseMonth = "SKIP:не закрыт"
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    res = LdTurnoverData(d0, d1)
    qEnd = res(5)(5)
    sumEnd = res(6)(5)
    f1 = LdTurnoverOut(res, d0, d1, True, "Закрытие " & sMonth & " — оборотная ведомость")
    If Left(f1, 5) <> "FILE:" Then
        CloseMonth = "ERR:месяц не закрыт — " & Mid(f1, 5)
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    f2 = StockOnDate(d1, True)
    If Left(f2, 5) <> "FILE:" Then
        CloseMonth = "ERR:месяц не закрыт — " & Mid(f2, 5)
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_CLOSE)
    d = sh.getCellRangeByPosition(0, 1, 0, CL_MAX).getDataArray()
    For k = 0 To UBound(d)
        If CStr(d(k)(0)) = "" Then Exit For
    Next k
    If k > UBound(d) Then
        CloseMonth = "ERR:таблица закрытий " & SH_CLOSE & " заполнена"
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    k = k + 1
    PlanBegin("MONTH_CLOSE", sMonth)
    PlanField("MONTH", sMonth)
    PlanField("TO", Format(d1, "YYYY-MM-DD"))
    PlanField("QTY_END", WmsIssue.Round3(qEnd))
    PlanField("SUM_END", Int(sumEnd * 100 + 0.5) / 100)
    PlanField("FILE", LdFileName(Mid(f1, 6)))
    PlanSetValue(SH_CLOSE, k, CL_MONTH, sMonth, False)
    PlanSetValue(SH_CLOSE, k, CL_TO, d1, False)
    PlanSetValue(SH_CLOSE, k, CL_TS, Format(Now(), "DD.MM.YYYY HH:MM"), False)
    PlanSetValue(SH_CLOSE, k, CL_STATE, CL_CLOSED, False)
    PlanSetValue(SH_CLOSE, k, CL_QTY, WmsIssue.Round3(qEnd), False)
    PlanSetValue(SH_CLOSE, k, CL_SUM, Int(sumEnd * 100 + 0.5) / 100, False)
    PlanSetValue(SH_CLOSE, k, CL_FILE, LdFileName(Mid(f1, 6)) & " · " & LdFileName(Mid(f2, 6)), False)
    r = ApplyOperation(0, 0)
    If Left(r, 3) <> "OK:" Then
        CloseMonth = "ERR:месяц не закрыт — " & Mid(r, InStr(r, ":") + 1)
        gUiLastMsg = CloseMonth
        Exit Function
    End If
    LdTurnoverOut(res, d0, d1, False, "Закрытие " & sMonth & " — оборотная ведомость")
    CloseMonth = "OK:" & LdMonthWord(m) & " " & y & " закрыт (по " & Format(d1, "DD.MM.YYYY") & "). Остаток на конец на сумму " & WmsPrint.Money(sumEnd) _
        & ". Файлы:" & Chr(10) & Mid(f1, 6) & Chr(10) & Mid(f2, 6)
    gUiLastMsg = CloseMonth
    Exit Function
EH:
    CloseMonth = "ERR:" & "не выполнено: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = CloseMonth
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

Private Function LdFileName(sPath As String) As String
    LdFileName = Mid(sPath, LdLastPos(Replace(sPath, "\", "/"), "/") + 1)
End Function

Sub BtnOpenMonth(Optional oEvent As Variant)
    Dim s As String, sWhy As String
    WmsInit()
    If ClosedTo() <= 0 Then
        WmsUi.UiMessage("Закрытых месяцев нет.")
        Exit Sub
    End If
    If Not WmsPrimeUi.ReasonAsk("Открыть месяц", "Снять закрытие последнего закрытого месяца (по " & Format(ClosedTo(), "DD.MM.YYYY") & "). Операции с датой в нём " _
        & "снова проводятся без предупреждения. Файлы закрытия остаются в «" & DOCS_DIR & "».", "Открыть месяц", sWhy) Then Exit Sub
    s = OpenMonth(sWhy)
    If Left(s, 3) = "OK:" Then WmsOrdersUi.UiInfo(Mid(s, 4)) Else LdReport(s)
End Sub

' the latest closing in force taken back (MONTH_OPEN, the reason is required). OK:… / ERR:…
Function OpenMonth(ByVal sWhy As String) As String
    Dim sh As Object, d As Variant, i As Long, k As Long, x As Double, r As String
    On Error GoTo EH
    sWhy = Trim(sWhy)
    If sWhy = "" Then
        OpenMonth = "ERR:нужна причина — почему месяц открывается снова"
        gUiLastMsg = OpenMonth
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_CLOSE)
    d = sh.getCellRangeByPosition(0, 1, CL_STATE, CL_MAX).getDataArray()
    k = -1
    For i = 0 To UBound(d)
        If CStr(d(i)(CL_MONTH)) = "" Then Exit For
        If CStr(d(i)(CL_STATE)) = CL_CLOSED And VarType(d(i)(CL_TO)) = 5 Then
            If d(i)(CL_TO) > x Then
                x = d(i)(CL_TO)
                k = i + 1
            End If
        End If
    Next i
    If k < 0 Then
        OpenMonth = "ERR:закрытых месяцев нет"
        gUiLastMsg = OpenMonth
        Exit Function
    End If
    PlanBegin("MONTH_OPEN", CStr(d(k - 1)(CL_MONTH)))
    PlanField("MONTH", CStr(d(k - 1)(CL_MONTH)))
    PlanField("WHY", sWhy)
    PlanSetValue(SH_CLOSE, k, CL_STATE, CL_OPENED, False)
    PlanSetValue(SH_CLOSE, k, CL_WHY, sWhy, False)
    r = ApplyOperation(0, 0)
    If Left(r, 3) <> "OK:" Then
        OpenMonth = "ERR:месяц не открыт — " & Mid(r, InStr(r, ":") + 1)
    Else
        OpenMonth = "OK:месяц " & d(k - 1)(CL_MONTH) & " открыт" & IIf(ClosedTo() > 0, "; закрыто теперь по " & Format(ClosedTo(), "DD.MM.YYYY"), "; закрытых месяцев нет")
    End If
    gUiLastMsg = OpenMonth
    Exit Function
EH:
    OpenMonth = "ERR:" & "не выполнено: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = OpenMonth
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' ================================================================ А2 the money on the stock

' the money of the stock by the price of the order, cached until the next operation (LAST_SEQ): the receipts of «Заказы»
' (J × X — the balance their EI shows there; a copy does not count) by the engine of Calc, the EIs split off by the price of
' their source. Array(sum, EIs with a balance and a price)
Function StockMoney() As Variant
    Dim seq As Long, osh As Object, last As Long, cond As String, a As Variant, sum As Double, nP As Long, ash As Object, items As Variant
    Dim nItems As Long, i As Long, d As Variant, m As Long, q As Double, p As Double, st As Object, J As String, X As String, nextEI As Long
    seq = CLng(SysNum(SK_LAST_SEQ))
    If gLdMoneyOk And gLdMoneySeq = seq Then
        StockMoney = Array(gLdMoney, gLdMoneyN)
        Exit Function
    End If
    osh = WmsOrders.OrdersSheet()
    last = WmsOrders.LastRow(osh)
    If last >= 1 Then
        J = WmsView.ColRng(SH_ORDERS, OC_PRICE, 1, last)
        X = WmsView.ColRng(SH_ORDERS, OC_STOCK, 1, last)
        cond = "(LEFT(" & WmsView.ColRng(SH_ORDERS, OC_CTL, 1, last) & ";5)<>""КОПИЯ"")*ISNUMBER(" & J & ")*ISNUMBER(" & X & ")*(" _
            & WmsView.ColRng(SH_ORDERS, OC_EI, 1, last) & "<>"""")"
        a = Split(WmsView.Eng("SUMPRODUCT(IF(" & cond & ";" & J & "*" & X & ";0))&""|""&SUMPRODUCT(" & cond & "*(" & X & ">0.0000005))"), "|")
        If UBound(a) >= 1 Then
            sum = Val(Replace(a(0), ",", "."))
            nP = Val(a(1))
        End If
    End If
    If gDoc.Sheets.hasByName(SH_ADJ) Then
        ash = gDoc.Sheets.getByName(SH_ADJ)
        last = WmsOrders.LastRow(ash)
        If last >= 1 Then
            ' the split EIs and (1.0, Ф8) the kits — their own build row (AJ_EI = AJ_NEW)
            items = WmsView.EngList("TEXTJOIN("";"";1;IF(((" & WmsView.ColRng(SH_ADJ, AJ_KIND, 1, last) & "=""SPLIT"")+(" & WmsView.ColRng(SH_ADJ, AJ_KIND, 1, last) _
                & "=""KIT"")*(" & WmsView.ColRng(SH_ADJ, AJ_EI, 1, last) & "=" & WmsView.ColRng(SH_ADJ, AJ_NEW, 1, last) & "))*(" & WmsView.ColRng(SH_ADJ, AJ_STATE, 1, last) _
                & "=""" & RV_LIVE & """);ROW(" & WmsView.ColRng(SH_ADJ, AJ_KIND, 1, last) & ")-1;""""))", nItems)
            st = WmsIssue.StockSheet()
            nextEI = CLng(SysNum(SK_NEXT_EI))
            For i = 0 To nItems - 1
                d = ash.getCellRangeByPosition(0, CLng(items(i)), AJ_LAST, CLng(items(i))).getDataArray()(0)
                m = WmsView.EiNo(CStr(d(AJ_NEW)), nextEI)
                If m >= 1 Then
                    q = 0
                    If st.getCellByPosition(SC_QTY, m).getType() = com.sun.star.table.CellContentType.VALUE Then q = st.getCellByPosition(SC_QTY, m).getValue()
                    If q > 0.0000005 Then
                        p = WmsPrint.EiPrice(m)
                        If p >= 0 Then
                            sum = sum + q * p
                            nP = nP + 1
                        End If
                    End If
                End If
            Next i
        End If
    End If
    gLdMoney = sum
    gLdMoneyN = nP
    gLdMoneySeq = seq
    gLdMoneyOk = True
    StockMoney = Array(sum, nP)
End Function

' ================================================================ А1 the report of the manager

Sub BtnManagerReport(Optional oEvent As Variant)
    Dim k As Integer, d0 As Double, d1 As Double, t As Double, v(1) As String, msg As String, s As String
    WmsInit()
    t = WmsOrders.Today()
    k = WmsPrimeUi.UiButtons("Отчёт руководителю", "За какой срок?", Array("Сегодня", "Неделя", "Месяц", "Период…", "Отмена"), Array(1, 2, 3, 4, 0))
    Select Case k
    Case 1
        d0 = t
        d1 = t
    Case 2
        d0 = t - 6
        d1 = t
    Case 3
        d0 = CDbl(DateSerial(Year(t), Month(t), 1))
        d1 = t
    Case 4
        v(0) = Format(CDbl(DateSerial(Year(t), Month(t), 1)), "DD.MM.YYYY")
        v(1) = Format(t, "DD.MM.YYYY")
        If Not WmsPrimeUi.UiForm("Отчёт руководителю", "Период отчёта.", Array("С", "По"), v, Array("", ""), "Показать") Then Exit Sub
        d0 = LdFormDay(v(0), "дата начала", msg)
        If d0 > 0 Then d1 = LdFormDay(v(1), "дата конца", msg)
        If d0 = 0 Or d1 = 0 Then
            WmsUi.UiMessage(msg)
            Exit Sub
        End If
    Case Else
        Exit Sub
    End Select
    s = ManagerReport(d0, d1)
    LdReport(s)
End Sub

' «Отчёт руководителю» for [d0, d1] on «Печать»: the operations and their money, the stock at the end, the recipients, the
' losses by reason, the suppliers, the vehicles. OK:… / ERR:…
Function ManagerReport(d0 As Double, d1 As Double) As String
    Dim k As Long, n As Long, q As Double, p As Double, cnt(10) As Long, mon(10) As Double, noP(10) As Long, sh As Object, rows() As Variant, i As Long
    Dim aft() As Double, b As Double, nStock As Long, sumStock As Double, names As Variant, kk As Integer, g As Variant, ng As Long, csh As Object
    Dim last As Long, cd As Variant, nCar As Long, nGone As Long, durSum As Double, nCancel As Long, plus As Double, minus As Double, s As String
    On Error GoTo EH
    If d1 < d0 Then
        ManagerReport = "ERR:дата конца раньше даты начала"
        gUiLastMsg = ManagerReport
        Exit Function
    End If
    LedgerRead()
    ReDim aft(gLdMax)
    For k = 0 To gLdN - 1
        n = gLdEi(k)
        If gLdDate(k) > d1 Then aft(n) = aft(n) + gLdQty(k)
        If gLdDate(k) >= d0 And gLdDate(k) <= d1 Then
            kk = gLdKind(k)
            q = gLdQty(k)
            p = gLdPrice(n)
            ' a kit build counts its kits (the rows of the components are inside them)
            If kk <> LK_KIT Or q > 0 Then
                cnt(kk) = cnt(kk) + 1
                If p >= 0 Then mon(kk) = mon(kk) + Abs(q) * p Else noP(kk) = noP(kk) + 1
            End If
            If kk = LK_INV And p >= 0 Then
                If q > 0 Then plus = plus + q * p Else minus = minus - q * p
            End If
        End If
    Next k
    For n = 1 To gLdMax
        b = WmsIssue.Round3(gLdNow(n) - aft(n))
        If b > 0.0000005 Then
            nStock = nStock + 1
            If gLdPrice(n) >= 0 Then sumStock = sumStock + b * gLdPrice(n)
        End If
    Next n
    names = Array("", "Приход по заказам", "Иной приход", "Выдачи", "Возвраты на склад", "Списания", "Пересчёты", "Возвраты поставщику", "Разделения ЕИ", _
        "Израсходовано до переноса", "Собрано комплектов (ЕИ)")
    ReDim rows(10)
    i = 0
    For kk = 1 To 10
        If cnt(kk) > 0 And kk <> LK_SPLIT Then
            rows(i) = Array(names(kk), cnt(kk), IIf(kk = LK_INV, "излишки " & WmsPrint.Money(plus) & " · недостачи " & WmsPrint.Money(minus), WmsPrint.Money(mon(kk))), _
                IIf(noP(kk) > 0, "без цены: " & noP(kk), ""))
            i = i + 1
        End If
    Next kk
    sh = WmsPrint.PrintBegin("ОТЧЁТ РУКОВОДИТЕЛЮ за " & Format(d0, "DD.MM.YYYY") & IIf(d1 <> d0, " — " & Format(d1, "DD.MM.YYYY"), ""), _
        "Отчёт руководителю " & Format(d0, "YYYY-MM-DD") & IIf(d1 <> d0, " — " & Format(d1, "YYYY-MM-DD"), ""))
    WmsPrint.PrintLine(sh, "На складе на " & Format(d1, "DD.MM.YYYY") & ": ЕИ с остатком " & nStock & " на сумму " & WmsPrint.Money(sumStock) _
        & " (по цене из заказа; ЕИ без цены в сумму не входят).", True)
    WmsPrint.PrintLine(sh, "Операции за период:", True)
    If i > 0 Then
        WmsPrint.PrintTable(sh, Array("Операция", "Строк", "На сумму", "Примечание"), rows, i, "1,2")
    Else
        WmsPrint.PrintLine(sh, "операций нет")
    End If
    ' the recipients and the losses (the first ten), the suppliers
    g = LdGroupBy(d0, d1, "WHO", ng)
    WmsPrint.PrintLine(sh, "Получатели — больше всего выдано (на сумму):", True)
    LdPrintGroups(sh, g, ng, 10, Array("№", "Получатель", "Выдач", "Выдано на сумму", "Возвращено на сумму", "На руках на сумму", "Без цены"))
    g = LdGroupBy(d0, d1, "LOSS", ng)
    WmsPrint.PrintLine(sh, "Потери — списания и недостачи пересчёта по причинам:", True)
    LdPrintGroups(sh, g, ng, 10, Array("№", "Причина", "Строк", "На сумму", "", "", "Без цены"))
    g = LdGroupBy(d0, d1, "SUP", ng)
    WmsPrint.PrintLine(sh, "Поставщики — принято за период:", True)
    LdPrintGroups(sh, g, ng, 10, Array("№", "Поставщик", "Позиций", "На сумму", "", "", "Без цены"))
    ' the vehicles of «Приход авто»
    If gDoc.Sheets.hasByName(SH_CARS) Then
        csh = gDoc.Sheets.getByName(SH_CARS)
        last = WmsOrders.LastRow(csh)
        If last >= CAR_FIRST Then
            cd = csh.getCellRangeByPosition(0, CAR_FIRST, CC_STATUS, last).getDataArray()
            For i = 0 To UBound(cd)
                If VarType(cd(i)(CC_DATE)) = 5 Or VarType(cd(i)(CC_DATE)) = 7 Then
                    If Int(cd(i)(CC_DATE)) >= d0 And Int(cd(i)(CC_DATE)) <= d1 Then
                        If CStr(cd(i)(CC_STATUS)) = CS_CANCEL Then
                            nCancel = nCancel + 1
                        Else
                            nCar = nCar + 1
                            If CStr(cd(i)(CC_STATUS)) = CS_GONE And VarType(cd(i)(CC_DUR)) = 5 Then
                                nGone = nGone + 1
                                durSum = durSum + cd(i)(CC_DUR)
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        ' (IIf computes both of its branches: the division only under an If)
        s = ""
        If nGone > 0 Then s = " · средняя стоянка " & CLng(durSum / nGone * 1440) & " мин"
        WmsPrint.PrintLine(sh, "Машины («" & SH_CARS & "»): визитов " & nCar & s & IIf(nCancel > 0, " · отменено " & nCancel, ""), True)
    End If
    ManagerReport = WmsPrint.PrintDone(sh)
    Exit Function
EH:
    ManagerReport = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = ManagerReport
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' the movements of [d0, d1] grouped: "WHO" — the recipients (issues, returns), "LOSS" — the reasons of the write-offs and the
' shortages of a count, "SUP" — the suppliers of the receipts by order. Array of Array(name, count, money1, money2, money3,
' no price), the biggest money first; ng — their number
Private Function LdGroupBy(d0 As Double, d1 As Double, sBy As String, ByRef ng As Long) As Variant
    Dim k As Long, m As Long, keys() As String, idx() As Long, ord As Variant, i As Long, g As Long, kk As String, prev As String, out() As Variant
    Dim q As Double, p As Double, n As Long, take As Boolean, nm() As String, c() As Long, m1() As Double, m2() As Double, np() As Long, okey() As Double
    Dim ord2 As Variant, res() As Variant
    ReDim keys(IIf(gLdN > 0, gLdN - 1, 0))
    ReDim idx(IIf(gLdN > 0, gLdN - 1, 0))
    For k = 0 To gLdN - 1
        If gLdDate(k) >= d0 And gLdDate(k) <= d1 Then
            take = False
            Select Case sBy
            Case "WHO"
                take = (gLdKind(k) = LK_ISS Or gLdKind(k) = LK_RET)
            Case "LOSS"
                take = (gLdKind(k) = LK_WOFF Or (gLdKind(k) = LK_INV And gLdQty(k) < 0))
            Case "SUP"
                take = (gLdKind(k) = LK_RCV)
            End Select
            If take Then
                keys(m) = LCase(Trim(IIf(Trim(gLdWho(k)) <> "", gLdWho(k), "—")))
                idx(m) = k
                m = m + 1
            End If
        End If
    Next k
    ng = 0
    If m = 0 Then
        LdGroupBy = Array()
        Exit Function
    End If
    ord = SortByCalc(keys, m)
    ReDim nm(m - 1)
    ReDim c(m - 1)
    ReDim m1(m - 1)
    ReDim m2(m - 1)
    ReDim np(m - 1)
    g = -1
    prev = Chr(1)
    For i = 0 To m - 1
        k = idx(ord(i))
        kk = LCase(keys(ord(i)))
        If kk <> prev Then
            g = g + 1
            prev = kk
            nm(g) = IIf(Trim(gLdWho(k)) <> "", Trim(gLdWho(k)), "—")
        End If
        n = gLdEi(k)
        q = gLdQty(k)
        p = gLdPrice(n)
        If gLdKind(k) <> LK_RET Then c(g) = c(g) + 1
        If p < 0 Then
            np(g) = np(g) + 1
        ElseIf gLdKind(k) = LK_RET Then
            m2(g) = m2(g) + q * p
        Else
            m1(g) = m1(g) + Abs(q) * p
        End If
    Next i
    ng = g + 1
    ReDim okey(ng - 1)
    For g = 0 To ng - 1
        okey(g) = -(m1(g) + 0.000001 * c(g))
    Next g
    ord2 = SortByCalc(okey, ng)
    ReDim res(ng - 1)
    For i = 0 To ng - 1
        g = ord2(i)
        res(i) = Array(nm(g), c(g), m1(g), m2(g), m1(g) - m2(g), np(g))
    Next i
    LdGroupBy = res
End Function

' the first nMax groups as a table (the money columns of heads that are "" are left out of the rows)
Private Sub LdPrintGroups(sh As Object, g As Variant, ng As Long, nMax As Long, heads As Variant)
    Dim rows() As Variant, i As Long, n As Long, r As Variant
    If ng = 0 Then
        WmsPrint.PrintLine(sh, "нет")
        Exit Sub
    End If
    n = IIf(ng < nMax, ng, nMax)
    ReDim rows(n - 1)
    For i = 0 To n - 1
        r = g(i)
        rows(i) = Array(i + 1, r(0), r(1), WmsPrint.Money(r(2)), IIf(heads(4) <> "", WmsPrint.Money(r(3)), ""), IIf(heads(5) <> "", WmsPrint.Money(r(4)), ""), _
            IIf(r(5) > 0, r(5), ""))
    Next i
    WmsPrint.PrintTable(sh, heads, rows, n, "0,2,3,4,5,6")
    If ng > nMax Then WmsPrint.PrintLine(sh, "… и ещё " & (ng - nMax))
End Sub

' ================================================================ А3 the consumption by the recipients and the departments

Sub BtnConsumers(Optional oEvent As Variant)
    Dim s As String, v(1) As String, d0 As Double, d1 As Double, msg As String, t As Double
    WmsInit()
    t = WmsOrders.Today()
    v(0) = Format(CDbl(DateSerial(Year(t), Month(t), 1)), "DD.MM.YYYY")
    v(1) = Format(t, "DD.MM.YYYY")
    If Not WmsPrimeUi.UiForm("Расход по получателям", "Сколько выдано, возвращено и на руках у каждого получателя и подразделения (колонка C «" _
        & SH_RCPT & "») — на какую сумму.", Array("С", "По"), v, Array("", ""), "Показать") Then Exit Sub
    d0 = LdFormDay(v(0), "дата начала", msg)
    If d0 > 0 Then d1 = LdFormDay(v(1), "дата конца", msg)
    If d0 = 0 Or d1 = 0 Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = Consumers(d0, d1)
    LdReport(s)
End Sub

' the recipients of [d0, d1] (every one) and their departments (C of «Получатели»): issues, the money issued, returned and in
' hand. OK:… / ERR:…
Function Consumers(d0 As Double, d1 As Double) As String
    Dim g As Variant, ng As Long, sh As Object, rows() As Variant, i As Long, r As Variant, shown As Long, dep As Object, a As Variant, rs As Object
    Dim last As Long, j As Long, dn() As String, dc() As Long, d1m() As Double, d2m() As Double, nd As Long, k As Long, dk As String, t1 As Double
    Dim t2 As Double, tc As Long, depName() As String
    On Error GoTo EH
    If d1 < d0 Then
        Consumers = "ERR:дата конца раньше даты начала"
        gUiLastMsg = Consumers
        Exit Function
    End If
    LedgerRead()
    g = LdGroupBy(d0, d1, "WHO", ng)
    ' the departments of the recipients
    dep = New Collection
    ReDim depName(0)
    If gDoc.Sheets.hasByName(SH_RCPT) Then
        rs = gDoc.Sheets.getByName(SH_RCPT)
        last = WmsOrders.LastRow(rs)
        If last >= 1 And rs.getCellByPosition(2, 0).getString() = "Подразделение" Then
            a = rs.getCellRangeByPosition(0, 1, 2, last).getDataArray()
            ReDim depName(UBound(a))
            For j = 0 To UBound(a)
                If Trim(CStr(a(j)(1))) <> "" And Trim(CStr(a(j)(2))) <> "" Then
                    If WmsView.CollPut(dep, j, LCase(Trim(CStr(a(j)(1))))) Then depName(j) = Trim(CStr(a(j)(2)))
                End If
            Next j
        End If
    End If
    ReDim rows(IIf(ng > 0, ng - 1, 0))
    ReDim dn(IIf(ng > 0, ng, 0))
    ReDim dc(IIf(ng > 0, ng, 0))
    ReDim d1m(IIf(ng > 0, ng, 0))
    ReDim d2m(IIf(ng > 0, ng, 0))
    For i = 0 To ng - 1
        r = g(i)
        j = WmsView.CollGet(dep, LCase(Trim(CStr(r(0)))))
        dk = ""
        If j >= 0 Then dk = depName(j)
        rows(i) = Array(i + 1, r(0) & IIf(dk <> "", " (" & dk & ")", ""), r(1), WmsPrint.Money(r(2)), WmsPrint.Money(r(3)), WmsPrint.Money(r(4)), IIf(r(5) > 0, r(5), ""))
        If dk = "" Then dk = "без подразделения"
        For k = 0 To nd - 1
            If dn(k) = dk Then Exit For
        Next k
        If k = nd Then
            dn(nd) = dk
            nd = nd + 1
        End If
        dc(k) = dc(k) + r(1)
        d1m(k) = d1m(k) + r(2)
        d2m(k) = d2m(k) + r(3)
        tc = tc + r(1)
        t1 = t1 + r(2)
        t2 = t2 + r(3)
    Next i
    shown = ng
    If shown > PR_MAX - 120 Then shown = PR_MAX - 120
    sh = WmsPrint.PrintBegin("РАСХОД ПО ПОЛУЧАТЕЛЯМ за " & Format(d0, "DD.MM.YYYY") & " — " & Format(d1, "DD.MM.YYYY"), "Расход по получателям " _
        & Format(d0, "YYYY-MM-DD") & " — " & Format(d1, "YYYY-MM-DD"))
    WmsPrint.PrintLine(sh, "Получателей: " & ng & " · выдач " & tc & " на сумму " & WmsPrint.Money(t1) & " · возвращено на " & WmsPrint.Money(t2) _
        & " · на руках на " & WmsPrint.Money(t1 - t2) & ". По цене из заказа.", True)
    If ng > 0 Then
        WmsPrint.PrintTable(sh, Array("№", "Получатель (подразделение)", "Выдач", "Выдано на сумму", "Возвращено", "На руках", "Без цены"), rows, shown, "0,2,3,4,5,6")
    Else
        WmsPrint.PrintLine(sh, "выдач и возвратов за период нет")
    End If
    If nd > 1 Or (nd = 1 And dn(0) <> "без подразделения") Then
        ReDim rows(nd - 1)
        For k = 0 To nd - 1
            rows(k) = Array(k + 1, dn(k), dc(k), WmsPrint.Money(d1m(k)), WmsPrint.Money(d2m(k)), WmsPrint.Money(d1m(k) - d2m(k)), "")
        Next k
        WmsPrint.PrintLine(sh, "По подразделениям:", True)
        WmsPrint.PrintTable(sh, Array("№", "Подразделение", "Выдач", "Выдано на сумму", "Возвращено", "На руках", ""), rows, nd, "0,2,3,4,5")
    End If
    Consumers = WmsPrint.PrintDone(sh)
    Exit Function
EH:
    Consumers = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = Consumers
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' ================================================================ А4 the dead stock

Sub BtnDeadStock(Optional oEvent As Variant)
    Dim s As String, v(1) As String
    WmsInit()
    v(0) = "90"
    v(1) = "по дням"
    If Not WmsPrimeUi.UiForm("Мёртвый груз", "ЕИ с остатком, у которых не было движения дольше порога (приход, выдача, возврат, корректировка), с суммой.", _
        Array("Порог, дней", "Сортировка"), v, Array("90;180;365", "по дням;по сумме"), "Показать") Then Exit Sub
    If Not IsNumeric(Trim(v(0))) Then
        WmsUi.UiMessage("Порог — число дней: 90, 180 или 365.")
        Exit Sub
    End If
    s = DeadStock(CLng(Trim(v(0))), Left(v(1), 8) = "по сумме")
    LdReport(s)
End Sub

' every EI with a balance that did not move for more than nDays (no movement at all in the book — the longest), the days and the
' money; sorted by the days or by the money; the total. OK:… / ERR:…
Function DeadStock(nDays As Long, bByMoney As Boolean) As String
    Dim lastMv() As Double, k As Long, n As Long, today As Double, keys() As Double, idx() As Long, m As Long, ord As Variant, rows() As Variant, i As Long
    Dim sum As Double, p As Double, b As Double, sh As Object, shown As Long, nNo As Long, dd As Variant
    On Error GoTo EH
    If nDays < 1 Then
        DeadStock = "ERR:порог — число дней больше нуля"
        gUiLastMsg = DeadStock
        Exit Function
    End If
    LedgerRead()
    today = WmsOrders.Today()
    ReDim lastMv(gLdMax)
    For k = 0 To gLdN - 1
        If gLdDate(k) > lastMv(gLdEi(k)) Then lastMv(gLdEi(k)) = gLdDate(k)
    Next k
    ReDim keys(IIf(gLdMax > 0, gLdMax - 1, 0))
    ReDim idx(IIf(gLdMax > 0, gLdMax - 1, 0))
    For n = 1 To gLdMax
        b = gLdNow(n)
        If b > 0.0000005 And (lastMv(n) = 0 Or today - lastMv(n) > nDays) Then
            p = gLdPrice(n)
            If p >= 0 Then sum = sum + b * p
            If lastMv(n) = 0 Then nNo = nNo + 1
            If bByMoney Then
                keys(m) = -(IIf(p >= 0, b * p, 0)) - 0.000000001 * n
            Else
                keys(m) = lastMv(n) + 0.000000001 * n
            End If
            idx(m) = n
            m = m + 1
        End If
    Next n
    ord = SortByCalc(keys, m)
    ReDim rows(IIf(m > 0, m - 1, 0))
    For i = 0 To m - 1
        n = idx(ord(i))
        p = gLdPrice(n)
        dd = IIf(lastMv(n) > 0, CLng(today - lastMv(n)), "нет движений")
        rows(i) = Array(i + 1, gLdCanon(n), LdGoodsText(n), WmsPrint.PartsFirst(n, gLdNow(n), gLdUnit(n)), gLdPlace(n), dd, IIf(p >= 0, WmsPrint.Money(gLdNow(n) * p), "—"))
    Next i
    shown = m
    If shown > PR_MAX - 80 Then shown = PR_MAX - 80
    sh = WmsPrint.PrintBegin("МЁРТВЫЙ ГРУЗ — без движения больше " & nDays & " дней", "Мёртвый груз " & nDays & " дней")
    WmsPrint.PrintLine(sh, "ЕИ: " & m & IIf(nNo > 0, " (из них без движений в книге: " & nNo & ")", "") & " · на сумму " & WmsPrint.Money(sum) _
        & " · сортировка " & IIf(bByMoney, "по сумме", "по дням") & ". На " & Format(today, "DD.MM.YYYY") & ".", True)
    If m > 0 Then
        WmsPrint.PrintTable(sh, Array("№", "ЕИ", "Наименование (артикул)", "Остаток", "Место", "Дней", "Сумма"), rows, shown, "0,5,6")
        If shown < m Then WmsPrint.PrintLine(sh, "Показаны первые " & shown & " из " & m & ".", True)
    Else
        WmsPrint.PrintLine(sh, "Мёртвого груза нет.", True)
    End If
    WmsPrint.PrintLine(sh, "Итого на сумму: " & WmsPrint.Money(sum), True)
    DeadStock = WmsPrint.PrintDone(sh)
    Exit Function
EH:
    DeadStock = "ERR:" & "отчёт не построен: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = DeadStock
    On Error Resume Next
    WmsView.ViewDone(gDoc.Sheets.getByName(SH_PRINT))
End Function

' ================================================================ the page «Отчёты»

Sub NavReports(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_REPORTS, 0, 0)
End Sub

Sub TestLedgerFile(s As String)
    gLdFileTest = s
End Sub
