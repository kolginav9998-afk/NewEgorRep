' WmsOrders — лист «Заказы»: служебная идентичность позиций заказа и приходов, статус позиции, поиск строк, антидубль,
' зеркало остатка ЕИ в «Наличие» (X), лёгкий обработчик листа и обновление статусов, зависящих от даты («Ожидается»,
' «Просрочено», «Частично получено / просрочено») (MASTER SPEC v0.3 §2, §4–§6, §13–§16, §25; D-006, D-007, D-033,
' D-045…D-059; задание Core Phase 3).
'
' Модель OrderLineID (OLID):
'  - позиция заказа получает постоянный OLID своей первой операцией (первый приход или «Отменить заказ»); _ORD — плотная
'    таблица, строка = OLID: подсказка строки исходной позиции на «Заказы», ключ (ЕИ исходной строки), отпечаток позиции
'    (A|D|B|E|H|I|L|P), заказано, получено (сумма действующих приходов), число приходов, из них без документов, отмена;
'  - каждый приход — строка _RCV с номером своего ЕИ (плотно, строка = номер ЕИ): OLID позиции, подсказка строки, вид
'    (исходная или дополнительная строка), количество, состояние (LIVE / STORNO), ключ антидубля;
'  - номер строки — только подсказка: перед использованием строка проверяется ключом (ЕИ в V) или отпечатком; при
'    несовпадении строка ищется одной формулой MATCH движка Calc (лист _IDX), и подсказку исправляет та же операция.
Option Explicit

' lookup freshness: every lookup writes a new mark, the echo cell proves that the formulas were recalculated
Global gIdxNonce As Double
' test seam: "today" for the overdue status (0 = the real date); honoured only in TEST books
Global gTodayOverride As Double
' result of the last status refresh (startup report, «Главная»); how its partly received rows were found (FORMULA / ROWS)
Global gRefreshNote As String
Global gRefreshPath As String
' test seam: the refresh does not use the array formula (the fallback path is tested on its own)
Global gRefreshNoFormula As Boolean
' details of the last OrderRowKind: EI of the row, OLID of its position, source row or not, row hint outdated, why FOREIGN
Global gKn As Long
Global gKcanon As String
Global gKol As Long
Global gKsrc As Boolean
Global gKmoved As Boolean
Global gKwhy As String
' SourceRowOf / ReceiptRowOf: the row hint was outdated (rows were inserted above) and must be corrected
Global gSmoved As Boolean
Global gRmoved As Boolean
' UnpostedReceipts stopped at UNPOSTED_MAX_BLOCKS: its result is a lower bound
Global gRcvUnpostedMore As Boolean

Private mInHandler As Boolean

' ================================================================ sheets

Function OrdersSheet() As Object
    OrdersSheet = gDoc.Sheets.getByName(SH_ORDERS)
End Function

Function OrdSheet() As Object
    OrdSheet = gDoc.Sheets.getByName(SH_ORD)
End Function

Function RcvSheet() As Object
    RcvSheet = gDoc.Sheets.getByName(SH_RCV)
End Function

Function IdxSheet() As Object
    IdxSheet = gDoc.Sheets.getByName(SH_IDX)
End Function

Function LastRow(sh As Object) As Long
    Dim cur As Object
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    LastRow = cur.getRangeAddress().EndRow
End Function

' "" when the sheets of the order scenario exist and have the expected layout (checked at every start, fail-closed)
Function SheetsProblem() As String
    Dim names As Variant, i As Integer, c As Object, x As Double, sh As Object
    names = Array(SH_ORDERS, SH_ORD, SH_RCV, SH_IDX)
    For i = 0 To UBound(names)
        If Not gDoc.Sheets.hasByName(names(i)) Then
            SheetsProblem = "нет листа «" & names(i) & "»"
            Exit Function
        End If
    Next i
    sh = OrdersSheet()
    If sh.getCellByPosition(OC_ORDQTY, 0).getString() <> "Заказанное количество" Or sh.getCellByPosition(OC_EI, 0).getString() <> "Внутренний код" _
        Or sh.getCellByPosition(OC_DUP, 0).getString() <> "Возможный дубль" Then
        SheetsProblem = "лист «" & SH_ORDERS & "»: заголовок не совпадает с A:AB спецификации"
        Exit Function
    End If
    If sh.getCellByPosition(OC_BLOCK, 0).getString() <> ORDER_BLOCK_HEADER Then
        SheetsProblem = "лист «" & SH_ORDERS & "»: нет служебной колонки AC «" & ORDER_BLOCK_HEADER & "» (книга старой версии — обновите её)"
        Exit Function
    End If
    sh = OrdSheet()
    c = sh.getCellByPosition(OD_NEXT_COL, 0)
    If sh.getCellByPosition(OD_ID, 0).getString() <> "OLID" Or sh.getCellByPosition(OD_NEXT_COL - 1, 0).getString() <> "NEXT_OL" _
        Or c.getType() <> com.sun.star.table.CellContentType.VALUE Then
        SheetsProblem = "служебный лист " & SH_ORD & " повреждён (заголовок или NEXT_OL)"
        Exit Function
    End If
    x = c.getValue()
    If x < 1 Or x <> Int(x) Or x > MAX_SHEET_ROW Then
        SheetsProblem = SH_ORD & ": NEXT_OL должен быть целым числом от 1, найдено «" & c.getString() & "»"
        Exit Function
    End If
    If RcvSheet().getCellByPosition(RV_EI, 0).getString() <> "ЕИ" Or RcvSheet().getCellByPosition(RV_DUPKEY, 0).getString() <> "Ключ антидубля" Then
        SheetsProblem = "служебный лист " & SH_RCV & " повреждён (заголовок)"
        Exit Function
    End If
    sh = IdxSheet()
    For i = IX_ECHO To IX_LAST
        If i <> IX_LIST And sh.getCellByPosition(1, i).getType() <> com.sun.star.table.CellContentType.FORMULA Then
            SheetsProblem = "служебный лист " & SH_IDX & " повреждён: нет формулы поиска в строке " & (i + 1)
            Exit Function
        End If
    Next i
    SheetsProblem = ""
End Function

Function NextOl() As Long
    NextOl = CLng(OrdSheet().getCellByPosition(OD_NEXT_COL, 0).getValue())
End Function

' A..I of _ORD row olid
Function OlRow(olid As Long) As Variant
    OlRow = OrdSheet().getCellRangeByPosition(0, olid, OD_LAST, olid).getDataArray()(0)
End Function

' True when olid is an issued OrderLineID whose _ORD row holds it (dense addressing is checked, never assumed)
Function OlValid(olid As Long, od As Variant) As Boolean
    If olid < 1 Or olid >= NextOl() Then Exit Function
    od = OlRow(olid)
    If VarType(od(OD_ID)) <> 5 Then Exit Function
    OlValid = (od(OD_ID) = olid)
End Function

' A..I of _RCV row n (the receipt of EI n)
Function RcvRow(n As Long) As Variant
    RcvRow = RcvSheet().getCellRangeByPosition(0, n, RV_LAST, n).getDataArray()(0)
End Function

' True when EI n was created by a receipt of this WMS (its _RCV row holds it)
Function RcvRegistered(n As Long, d As Variant) As Boolean
    If n < 1 Or n > MAX_SHEET_ROW Then Exit Function
    d = RcvRow(n)
    RcvRegistered = (CStr(d(RV_EI)) = WmsIssue.EiCanon(n))
End Function

' the canonical form only (ЕИ-00000123): V is written by WMS, anything else in V is not a WMS key. The number is read
' with Val and the text must equal the canonical form rebuilt from it (3 times faster than a character loop — this runs
' for every row in the full checks)
Function StrictEI(v As String, ByRef n As Long) As Boolean
    Dim x As Double
    n = 0
    If Len(v) <> 11 Then Exit Function
    x = Val(Mid(v, 4))
    If x < 1 Or x > 99999999 Then Exit Function
    If v <> EI_PREFIX & Right("0000000" & CLng(x), 8) Then Exit Function
    n = CLng(x)
    StrictEI = True
End Function

' ================================================================ lookups through the Calc engine (_IDX, spec §25)
' The criterion is written into _IDX!B1 (and the start offset into B2), the formula of row ixRow answers. A new mark in B3
' and its echo prove that the formulas were recalculated (automatic calculation may be switched off by the user); if not,
' the document is recalculated, and as the last resort FunctionAccess computes the same formula. The undo stack and the
' «modified» flag are not touched.

' the value of the lookup formula of row ixRow; -1 when it is #N/A (not found)
Private Function IdxLookup(ixRow As Integer, key As Variant, start As Long) As Double
    Dim sh As Object, um As Object, wasMod As Boolean, c As Object, locked As Boolean, res As Double, fresh As Boolean
    On Error GoTo EH
    sh = IdxSheet()
    um = gDoc.getUndoManager()
    wasMod = gDoc.isModified()
    um.lock()
    locked = True
    ' a mark never used before in this book (a saved mark of an earlier session must not look fresh)
    If gIdxNonce < 1 Then gIdxNonce = Int(CDbl(Now()) * 86400) * 1000 + Int(Rnd() * 1000)
    gIdxNonce = gIdxNonce + 1
    If VarType(key) = 8 Then
        sh.getCellByPosition(1, IX_KEY).setString(key)
    Else
        sh.getCellByPosition(1, IX_KEY).setValue(CDbl(key))
    End If
    sh.getCellByPosition(1, IX_START).setValue(start)
    sh.getCellByPosition(1, IX_NONCE).setValue(gIdxNonce)
    fresh = (sh.getCellByPosition(1, IX_ECHO).getValue() = gIdxNonce)
    If Not fresh Then
        gDoc.calculate()
        fresh = (sh.getCellByPosition(1, IX_ECHO).getValue() = gIdxNonce)
    End If
    If fresh Then
        c = sh.getCellByPosition(1, ixRow)
        If c.getError() <> 0 Then res = -1 Else res = c.getValue()
    Else
        res = FaLookup(ixRow, key, start)
    End If
    um.unlock()
    locked = False
    If Not wasMod Then gDoc.setModified(False)
    IdxLookup = res
    Exit Function
EH:
    ' the lookup cells are damaged: the same formula through FunctionAccess (slower, copies the range)
    On Error Resume Next
    If locked Then um.unlock()
    If Not wasMod Then gDoc.setModified(False)
    On Error GoTo 0
    IdxLookup = FaLookup(ixRow, key, start)
End Function

Private Function FaLookup(ixRow As Integer, key As Variant, start As Long) As Double
    Dim fa As Object, col As Integer, sh As Object, rng As Object, x As Variant
    fa = CreateUnoService("com.sun.star.sheet.FunctionAccess")
    Select Case ixRow
    Case IX_V_MATCH, IX_V_COUNT
        sh = OrdersSheet()
        col = OC_EI
    Case IX_A_MATCH
        sh = OrdersSheet()
        col = OC_ORDER
    Case IX_W_MATCH
        sh = OrdersSheet()
        col = OC_STATUS
    Case IX_I_MATCH
        sh = gDoc.Sheets.getByName(SH_ISSUES)
        col = IC_NO
    Case IX_RA_MATCH, IX_RA_COUNT
        sh = gDoc.Sheets.getByName(SH_RETURNS)
        col = RC_NO
    Case IX_XA_MATCH, IX_XA_COUNT
        sh = gDoc.Sheets.getByName(SH_SPECIAL)
        col = XC_NO
    Case IX_ART_MATCH, IX_ART_COUNT
        sh = gDoc.Sheets.getByName(SH_ART)
        col = AR_KEY
    Case IX_XD_MATCH, IX_XD_COUNT
        sh = gDoc.Sheets.getByName(SH_SPR)
        col = SR_DUP
    Case IX_AA_MATCH, IX_AA_COUNT
        sh = gDoc.Sheets.getByName(SH_ADJUST)
        col = AC_NO
    Case IX_CO_MATCH, IX_CO_COUNT
        sh = gDoc.Sheets.getByName(SH_CAR)
        col = CR_OKEY
    Case IX_CK_COUNT
        sh = gDoc.Sheets.getByName(SH_CAR)
        col = CR_KEY
    Case IX_CP_MATCH
        sh = gDoc.Sheets.getByName(SH_CAR)
        col = CD_PKEY
    Case IX_CS_MATCH
        sh = gDoc.Sheets.getByName(SH_CAR)
        col = CD_SKEY
    Case Else
        sh = RcvSheet()
        col = RV_DUP
    End Select
    If ixRow = IX_V_COUNT Or ixRow = IX_DUP_COUNT Or ixRow = IX_RA_COUNT Or ixRow = IX_XA_COUNT Or ixRow = IX_ART_COUNT Or ixRow = IX_XD_COUNT _
        Or ixRow = IX_AA_COUNT Or ixRow = IX_CO_COUNT Or ixRow = IX_CK_COUNT Then
        FaLookup = fa.callFunction("COUNTIF", Array(sh.getCellRangeByPosition(col, 1, col, MAX_SHEET_ROW), key))
        Exit Function
    End If
    rng = sh.getCellRangeByPosition(col, start + 1, col, MAX_SHEET_ROW)
    On Error GoTo NF
    x = fa.callFunction("MATCH", Array(key, rng, 0))
    FaLookup = CDbl(x) + start
    Exit Function
NF:
    FaLookup = -1
End Function

' 0-based «Заказы» row of the first V equal to canon, searching from data row start + 1; -1 when none
Function FindEIRow(canon As String, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_V_MATCH, canon, start)
    If x < 1 Then FindEIRow = -1 Else FindEIRow = CLng(x)
End Function

' the first row holding canon in V that is not marked КОПИЯ (the key of a receipt row found after rows moved); -1 when none
Function FindKeyRow(canon As String) As Long
    Dim r As Long, start As Long, i As Integer, sh As Object
    sh = OrdersSheet()
    FindKeyRow = -1
    start = 0
    For i = 1 To 20
        r = FindEIRow(canon, start)
        If r < 1 Then Exit Function
        If Left(sh.getCellByPosition(OC_CTL, r).getString(), 5) <> "КОПИЯ" Then
            FindKeyRow = r
            Exit Function
        End If
        start = r
    Next i
End Function

Function CountEI(canon As String) As Long
    CountEI = CLng(IdxLookup(IX_V_COUNT, canon, 0))
End Function

' Phase 4: 0-based «Выдачи» row of the first issue № k (A), from data row start + 1; -1 when none
Function FindIssueNoRow(k As Long, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_I_MATCH, CDbl(k), start)
    If x < 1 Then FindIssueNoRow = -1 Else FindIssueNoRow = CLng(x)
End Function

' Phase 4: 0-based «Возврат» row of the first return № n (A), from data row start + 1; -1 when none
Function FindReturnNoRow(n As Long, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_RA_MATCH, CDbl(n), start)
    If x < 1 Then FindReturnNoRow = -1 Else FindReturnNoRow = CLng(x)
End Function

Function CountReturnNo(n As Long) As Long
    CountReturnNo = CLng(IdxLookup(IX_RA_COUNT, CDbl(n), 0))
End Function

' Phase 5: 0-based «Иной приход» row of the first line № n (A), from data row start + 1; -1 when none
Function FindSpecialNoRow(n As Long, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_XA_MATCH, CDbl(n), start)
    If x < 1 Then FindSpecialNoRow = -1 Else FindSpecialNoRow = CLng(x)
End Function

Function CountSpecialNo(n As Long) As Long
    CountSpecialNo = CLng(IdxLookup(IX_XA_COUNT, CDbl(n), 0))
End Function

' Final Core: 0-based «Корректировки» row of the first correction № n (A), from data row start + 1; -1 when none
Function FindAdjustNoRow(n As Long, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_AA_MATCH, CDbl(n), start)
    If x < 1 Then FindAdjustNoRow = -1 Else FindAdjustNoRow = CLng(x)
End Function

Function CountAdjustNo(n As Long) As Long
    CountAdjustNo = CLng(IdxLookup(IX_AA_COUNT, CDbl(n), 0))
End Function

' M6: the visit № (= its _CAR row) of the first open visit of the vehicle key after visit start; -1 when none. A key of a
' vehicle holds only letters and digits (WmsCar.PlateKey); the pattern "?*" (any key) finds the open visits one by one
Function FindCarOpen(key As String, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_CO_MATCH, key, start)
    If x < 1 Then FindCarOpen = -1 Else FindCarOpen = CLng(x)
End Function

' M6: the number of open visits of the vehicle key ("?*": of all vehicles)
Function CountCarOpen(key As String) As Long
    CountCarOpen = CLng(IdxLookup(IX_CO_COUNT, key, 0))
End Function

' M6: the number of visits (open, closed, cancelled) of the vehicle key
Function CountCarVisits(key As String) As Long
    CountCarVisits = CLng(IdxLookup(IX_CK_COUNT, Replace(Replace(Replace(key, "~", "~~"), "*", "~*"), "?", "~?"), 0))
End Function

' M6: the entry (= _CAR row) of the vehicle key in the dictionary of vehicles; -1 when none
Function FindCarPlate(key As String) As Long
    Dim x As Double
    x = IdxLookup(IX_CP_MATCH, Replace(Replace(Replace(key, "~", "~~"), "*", "~*"), "?", "~?"), 0)
    If x < 1 Then FindCarPlate = -1 Else FindCarPlate = CLng(x)
End Function

' M6: the entry (= _CAR row) of the supplier key in the dictionary of suppliers; -1 when none
Function FindCarSupplier(key As String) As Long
    Dim x As Double
    x = IdxLookup(IX_CS_MATCH, Replace(Replace(Replace(key, "~", "~~"), "*", "~*"), "?", "~?"), 0)
    If x < 1 Then FindCarSupplier = -1 Else FindCarSupplier = CLng(x)
End Function

' Phase 5: the _ART row of the first article key equal to key (text, compared as a whole cell, wildcards taken
' literally), from data row start + 1; -1 when none. COUNTIF is not used for text keys: its criterion would read a
' leading "<", ">" or "=" of an article as an operator — a second MATCH after the first one tells a repeated key.
Function FindArtRow(key As String, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_ART_MATCH, Replace(Replace(Replace(key, "~", "~~"), "*", "~*"), "?", "~?"), start)
    If x < 1 Then FindArtRow = -1 Else FindArtRow = CLng(x)
End Function

' Phase 5: the line № (= _SPR row) of the first special receipt line with antidubl hash h after line start; -1 when none
Function FindSpecialDupLine(h As Double, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_XD_MATCH, h, start)
    If x < 1 Then FindSpecialDupLine = -1 Else FindSpecialDupLine = CLng(x)
End Function

' One array formula of the Calc engine, evaluated at once in the unlocked scratch cell _IDX!B<IX_LIST> (written, read and
' cleared; no Undo, the book does not become «modified»): its text result, or ok = False when it could not be used
' (then the caller takes a slower path). A tag proves the value was calculated for this call (automatic calculation
' may be off). Used where the answer is a list (TEXTJOIN) — rows of «Заказы» to refresh, issues of one EI.
Function ScratchEval(body As String, ByRef ok As Boolean) As String
    Dim ix As Object, c As Object, um As Object, wasMod As Boolean, locked As Boolean, tag As String, s As String, flags As Long
    On Error GoTo EH
    ok = False
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    ix = IdxSheet()
    c = ix.getCellRangeByPosition(1, IX_LIST, 1, IX_LIST)
    um = gDoc.getUndoManager()
    wasMod = gDoc.isModified()
    um.lock()
    locked = True
    If gIdxNonce < 1 Then gIdxNonce = Int(CDbl(Now()) * 86400) * 1000 + Int(Rnd() * 1000)
    gIdxNonce = gIdxNonce + 1
    tag = "T" & Format(gIdxNonce, "0") & "|"
    c.setArrayFormula("=""" & tag & """&" & body)
    s = ix.getCellByPosition(1, IX_LIST).getString()
    If Left(s, Len(tag)) <> tag Then
        gDoc.calculate()
        s = ix.getCellByPosition(1, IX_LIST).getString()
    End If
    c.clearContents(flags)
    um.unlock()
    locked = False
    If Not wasMod Then gDoc.setModified(False)
    If Left(s, Len(tag)) <> tag Then Exit Function
    ScratchEval = Mid(s, Len(tag) + 1)
    ok = True
    Exit Function
EH:
    On Error Resume Next
    If Not IsNull(c) And Not IsEmpty(c) Then c.clearContents(flags)
    If locked Then um.unlock()
    If Not wasMod Then gDoc.setModified(False)
End Function

' EI number of the first _RCV row with antidubl hash h after EI start; -1 when none
Private Function FindDupEI(h As Double, start As Long) As Long
    Dim x As Double
    x = IdxLookup(IX_DUP_MATCH, h, start)
    If x < 1 Then FindDupEI = -1 Else FindDupEI = CLng(x)
End Function

Function CountDup(h As Double) As Long
    CountDup = CLng(IdxLookup(IX_DUP_COUNT, h, 0))
End Function

' 0-based row of the first status W equal to s; -1 when none
Private Function FindWRow(s As String) As Long
    Dim x As Double
    x = IdxLookup(IX_W_MATCH, s, 0)
    If x < 1 Then FindWRow = -1 Else FindWRow = CLng(x)
End Function

' 0-based row of the first order № (A) equal to v (text as typed or a number), from data row start + 1; -1 when none
Private Function FindOrderRow(v As Variant, start As Long) As Long
    Dim x As Double, s As String
    If VarType(v) = 8 Then
        ' the lookup cells honour wildcards (document setting): * ? ~ are taken literally
        s = Replace(Replace(Replace(CStr(v), "~", "~~"), "*", "~*"), "?", "~?")
        x = IdxLookup(IX_A_MATCH, s, start)
    Else
        x = IdxLookup(IX_A_MATCH, v, start)
    End If
    If x < 1 Then FindOrderRow = -1 Else FindOrderRow = CLng(x)
End Function

' ================================================================ row model of «Заказы»

' True when the user filled any of the order columns A..U of row r
Function RowHasInput(r As Long) As Boolean
    Dim d As Variant, i As Integer
    d = OrdersSheet().getCellRangeByPosition(0, r, OC_PLACE, r).getDataArray()(0)
    For i = 0 To OC_PLACE
        If CStr(d(i)) <> "" Then
            RowHasInput = True
            Exit Function
        End If
    Next i
End Function

' Kind of «Заказы» row r: EMPTY, OPEN (an order position without a receipt), RECEIVED (the row's receipt is live), STORNO
' (the row's receipt was cancelled), CANCELLED (the position was cancelled before any receipt), COPY (marked КОПИЯ), FOREIGN
' (V holds an EI that WMS did not register at this row: a copy or a forged row). RECEIVED/STORNO fill gKn, gKcanon (EI),
' gKol (OLID), gKsrc (source row of the position), gKmoved (the receipt's row hint is outdated: rows were inserted above).
Function OrderRowKind(r As Long) As String
    Dim sh As Object, v As String, n As Long, canon As String, d As Variant, hint As Long, x As Variant
    sh = OrdersSheet()
    gKn = 0
    gKcanon = ""
    gKol = 0
    gKsrc = False
    gKmoved = False
    gKwhy = ""
    If Left(sh.getCellByPosition(OC_CTL, r).getString(), 5) = "КОПИЯ" Then
        OrderRowKind = "COPY"
        Exit Function
    End If
    v = sh.getCellByPosition(OC_EI, r).getString()
    If v = "" Then
        If sh.getCellByPosition(OC_STATUS, r).getString() = OS_CANCELLED Then
            OrderRowKind = "CANCELLED"
        ElseIf RowHasInput(r) Then
            OrderRowKind = "OPEN"
        Else
            OrderRowKind = "EMPTY"
        End If
        Exit Function
    End If
    If Not StrictEI(v, n) Then
        gKwhy = "«" & v & "» в «Внутренний код» — не ЕИ WMS"
        OrderRowKind = "FOREIGN"
        Exit Function
    End If
    canon = WmsIssue.EiCanon(n)
    If n >= SysNum(SK_NEXT_EI) Then
        gKwhy = canon & " WMS не выдавала"
        OrderRowKind = "FOREIGN"
        Exit Function
    End If
    If Not RcvRegistered(n, d) Then
        gKwhy = canon & " не создан приходом WMS"
        OrderRowKind = "FOREIGN"
        Exit Function
    End If
    x = d(RV_ROW)
    If VarType(x) = 5 Then hint = CLng(x) Else hint = -1
    If hint <> r Then
        ' the registered row is elsewhere: when it still holds the key this row is a copy, otherwise rows moved
        If hint >= 1 And hint <= MAX_SHEET_ROW Then
            If sh.getCellByPosition(OC_EI, hint).getString() = canon Then
                gKwhy = canon & " повторяет строку " & (hint + 1)
                OrderRowKind = "FOREIGN"
                Exit Function
            End If
        End If
        If CountEI(canon) > 1 Then
            ' the key in several rows and none at the registered one: the first row in sheet order that is not a marked
            ' copy is taken as the original (as the full key check does)
            If FindKeyRow(canon) <> r Then
                gKwhy = canon & " повторяется"
                OrderRowKind = "FOREIGN"
                Exit Function
            End If
        End If
        gKmoved = True
    End If
    gKn = n
    gKcanon = canon
    gKol = CLng(d(RV_OL))
    gKsrc = (CStr(d(RV_KIND)) = RV_SRC)
    If CStr(d(RV_STATE)) = RV_STORNO Then OrderRowKind = "STORNO" Else OrderRowKind = "RECEIVED"
End Function

' the part of the position fingerprint for one cell: E, N<number>, S<trimmed text>
Private Function FpCell(c As Object) As String
    Select Case c.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        FpCell = "E"
    Case com.sun.star.table.CellContentType.VALUE
        FpCell = "N" & NumStr(c.getValue())
    Case com.sun.star.table.CellContentType.TEXT
        FpCell = "S" & Esc(Trim(c.getString()))
    Case Else
        FpCell = "F" & Esc(c.getFormula())
    End Select
End Function

' fingerprint of an order position A|D|B|E|H|I|L|P. For a row being posted H and P are passed as their normalized values
' (ordQty > 0; pDate = 0 when P is empty); for a posted row (ordQty < 0) the cells are read as they are.
Function Fingerprint(r As Long, ordQty As Double, pDate As Double) As String
    Dim sh As Object, s As String
    sh = OrdersSheet()
    s = FpCell(sh.getCellByPosition(OC_ORDER, r)) & "|" & FpCell(sh.getCellByPosition(OC_INVOICE, r)) & "|" _
        & FpCell(sh.getCellByPosition(OC_NAME, r)) & "|" & FpCell(sh.getCellByPosition(OC_ART, r)) & "|"
    If ordQty >= 0 Then s = s & "N" & NumStr(ordQty) Else s = s & FpCell(sh.getCellByPosition(OC_ORDQTY, r))
    s = s & "|" & FpCell(sh.getCellByPosition(OC_UNIT, r)) & "|" & FpCell(sh.getCellByPosition(OC_SUPPLIER, r)) & "|"
    If ordQty >= 0 Then
        If pDate > 0 Then s = s & "N" & NumStr(pDate) Else s = s & "E"
    Else
        s = s & FpCell(sh.getCellByPosition(OC_ODATE, r))
    End If
    Fingerprint = s
End Function

' 0-based «Заказы» row of the source row of position olid (od = its _ORD row); -1 when it cannot be found.
' Keyed positions (a receipt was made on the source row): the hint is checked by the EI in V, else one MATCH on V.
' Positions cancelled before any receipt have no key: the hint is checked by the fingerprint, else the rows with the same
' order № (A) are examined one by one. gSmoved = True when the hint was outdated.
Function SourceRowOf(od As Variant) As Long
    Dim sh As Object, hint As Long, key As String, r As Long, fp As String, a As Variant, t As String, i As Integer, start As Long
    sh = OrdersSheet()
    gSmoved = False
    SourceRowOf = -1
    If VarType(od(OD_ROW)) = 5 Then hint = CLng(od(OD_ROW)) Else hint = -1
    key = CStr(od(OD_KEY))
    If key <> "" Then
        If hint >= 1 And hint <= MAX_SHEET_ROW Then
            If sh.getCellByPosition(OC_EI, hint).getString() = key Then
                SourceRowOf = hint
                Exit Function
            End If
        End If
        r = FindKeyRow(key)
        If r >= 1 Then
            gSmoved = True
            SourceRowOf = r
        End If
        Exit Function
    End If
    fp = CStr(od(OD_FP))
    If hint >= 1 And hint <= MAX_SHEET_ROW Then
        If Fingerprint(hint, -1, 0) = fp And sh.getCellByPosition(OC_EI, hint).getString() = "" Then
            SourceRowOf = hint
            Exit Function
        End If
    End If
    ' the order № is the first part of the fingerprint: N<number> or S<escaped text>
    t = Split(fp, "|")(0)
    If Left(t, 1) = "N" Then
        a = Val(Mid(t, 2))
    ElseIf Left(t, 1) = "S" Then
        a = Unesc(Mid(t, 2))
    Else
        Exit Function
    End If
    start = 0
    For i = 1 To 200
        r = FindOrderRow(a, start)
        If r < 1 Then Exit Function
        If Fingerprint(r, -1, 0) = fp And sh.getCellByPosition(OC_EI, r).getString() = "" Then
            gSmoved = True
            SourceRowOf = r
            Exit Function
        End If
        start = r
    Next i
End Function

' 0-based «Заказы» row of the receipt of EI n (hint = its _RCV row hint); -1 when not found; gRmoved when outdated
Function ReceiptRowOf(canon As String, hint As Variant) As Long
    Dim h As Long, r As Long
    gRmoved = False
    ReceiptRowOf = -1
    If VarType(hint) = 5 Then h = CLng(hint) Else h = -1
    If h >= 1 And h <= MAX_SHEET_ROW Then
        If OrdersSheet().getCellByPosition(OC_EI, h).getString() = canon Then
            ReceiptRowOf = h
            Exit Function
        End If
    End If
    r = FindKeyRow(canon)
    If r >= 1 Then
        gRmoved = True
        ReceiptRowOf = r
    End If
End Function

' «Заказы».X shows the current balance of the very EI of a receipt row. Every operation that changes the balance of an EI
' created by a receipt (issue, correction, storno) adds this write to its plan, correcting the row hint when rows moved.
' Nothing for an EI that no receipt created (migrated or test EIs).
Sub PlanStockMirror(n As Long, bal As Double)
    Dim d As Variant, canon As String, r As Long
    If Not RcvRegistered(n, d) Then Exit Sub
    canon = WmsIssue.EiCanon(n)
    r = ReceiptRowOf(canon, d(RV_ROW))
    If r < 1 Then Exit Sub
    If gRmoved Then PlanSetValue(SH_RCV, n, RV_ROW, r, False)
    PlanDerived(SH_ORDERS, r, OC_STOCK, bal)
End Sub

' ================================================================ status of an order position (spec §6, v0.1 §15)

Function Today() As Double
    If gTodayOverride > 0 Then
        Today = gTodayOverride
    Else
        Today = Int(CDbl(Now()))
    End If
End Function

' The business status of a position: its ordered quantity, the sum of its live receipts, how many of them have no
' documents, the cancellation, the expected date Q (0 = none). The expected date has passed (Q before today): nothing
' received — «Просрочено», partly received — «Частично получено / просрочено» (D-047); a cancelled rest and a fully
' received position are never overdue.
Function PositionStatus(ordQty As Double, rcvQty As Double, nodoc As Double, cancel As String, edate As Double) As String
    If cancel = OD_CANCEL_ORDER Then
        PositionStatus = OS_CANCELLED
    ElseIf rcvQty <= 0.0000001 Then
        If cancel = OD_CANCEL_REST Then
            PositionStatus = OS_CANCELLED
        ElseIf edate > 0 And edate < Today() Then
            PositionStatus = OS_OVERDUE
        Else
            PositionStatus = OS_WAITING
        End If
    ElseIf rcvQty >= ordQty - 0.0000001 Then
        If nodoc > 0 Then PositionStatus = OS_NODOCS Else PositionStatus = OS_RECEIVED
    ElseIf cancel = OD_CANCEL_REST Then
        PositionStatus = OS_REST_CANCELLED
    ElseIf edate > 0 And edate < Today() Then
        PositionStatus = OS_PARTIAL_OVERDUE
    Else
        PositionStatus = OS_PARTIAL
    End If
End Function

' the position summary for «Контроль» of the source row
Function PositionSummary(ordQty As Double, rcvQty As Double, cancel As String, unit As String) As String
    Dim u As String
    u = IIf(unit <> "", " " & unit, "")
    If cancel = OD_CANCEL_ORDER Then
        PositionSummary = "позиция: заказ отменён"
    ElseIf rcvQty <= 0.0000001 Then
        If cancel = OD_CANCEL_REST Then
            PositionSummary = "позиция: ничего не получено, остаток отменён"
        Else
            PositionSummary = "позиция: ожидается " & WmsIssue.QtyText(ordQty) & u
        End If
    ElseIf rcvQty >= ordQty - 0.0000001 Then
        PositionSummary = "позиция: получено " & WmsIssue.QtyText(rcvQty) & " из " & WmsIssue.QtyText(ordQty) & u _
            & IIf(rcvQty > ordQty + 0.0000001, " (больше заказа на " & WmsIssue.QtyText(WmsIssue.Round3(rcvQty - ordQty)) & ")", "")
    ElseIf cancel = OD_CANCEL_REST Then
        PositionSummary = "позиция: получено " & WmsIssue.QtyText(rcvQty) & " из " & WmsIssue.QtyText(ordQty) & u & ", остаток " _
            & WmsIssue.QtyText(WmsIssue.Round3(ordQty - rcvQty)) & " отменён"
    Else
        PositionSummary = "позиция: получено " & WmsIssue.QtyText(rcvQty) & " из " & WmsIssue.QtyText(ordQty) & u & ", осталось " _
            & WmsIssue.QtyText(WmsIssue.Round3(ordQty - rcvQty))
    End If
End Function

' quantity control of one receipt (G against F) and the documents (C, O) — spec §6 «Контроль количества», «Документы»
Function ReceiptControl(fact As Double, hasDocQty As Boolean, docQty As Double, hasDoc As Boolean, hasDdate As Boolean, unit As String) As String
    Dim s As String, u As String
    u = IIf(unit <> "", " " & unit, "")
    If hasDocQty Then
        If Abs(docQty - fact) < 0.0000001 Then
            s = "Норма"
        ElseIf docQty > fact Then
            s = "Недостача: " & WmsIssue.QtyText(WmsIssue.Round3(docQty - fact)) & u & " (по документу " & WmsIssue.QtyText(docQty) & ")"
        Else
            s = "Излишек: " & WmsIssue.QtyText(WmsIssue.Round3(fact - docQty)) & u & " (по документу " & WmsIssue.QtyText(docQty) & ")"
        End If
    End If
    If Not hasDoc And Not hasDdate Then
        s = s & IIf(s <> "", "; ", "") & "нет документа (C, O)"
    ElseIf Not hasDoc Then
        s = s & IIf(s <> "", "; ", "") & "нет номера документа (C)"
    ElseIf Not hasDdate Then
        s = s & IIf(s <> "", "; ", "") & "нет даты документа (O)"
    End If
    ReceiptControl = s
End Function

' W and Y of the source row of a position, from its _ORD row (after the operation) and its expected date
Function SourceControl(ownControl As String, ordQty As Double, rcvQty As Double, cancel As String, unit As String) As String
    SourceControl = ownControl & IIf(ownControl <> "", "; ", "") & PositionSummary(ordQty, rcvQty, cancel, unit)
End Function

' ================================================================ antidubl (spec §30 of v0.1): supplier + document + article
' or name + quantity + date; an index (hash in _RCV!D, MATCH of the Calc engine), never a scan of the sheet. A warning only.

' lower case letters and digits; every other run of characters becomes one "_"; at most maxLen characters
Private Function NormPart(ByVal s As String, maxLen As Integer) As String
    Dim i As Long, ch As String, out As String, gap As Boolean, code As Long
    s = LCase(Trim(s))
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = Asc(ch)
        ' ё = е; digits, Latin and Cyrillic letters by their code points (no collation involved)
        If code = 1105 Then
            ch = Chr(1077)
            code = 1077
        End If
        If (code >= 48 And code <= 57) Or (code >= 97 And code <= 122) Or (code >= 1072 And code <= 1103) Then
            If gap And out <> "" Then out = out & "_"
            out = out & ch
            gap = False
            If Len(out) >= maxLen Then Exit For
        Else
            gap = True
        End If
    Next i
    NormPart = out
End Function

Function DupKeyText(supplier As String, doc As String, art As String, sName As String, fact As Double, dt As Double) As String
    Dim p3 As String
    If Trim(art) <> "" Then p3 = "a" & NormPart(art, 40) Else p3 = "n" & NormPart(sName, 60)
    DupKeyText = NormPart(supplier, 40) & "-" & NormPart(doc, 40) & "-" & p3 & "-" & Replace(NumStr(fact), ".", "d") & "-" & NumStr(Int(dt))
End Function

' polynomial hash of the key text below 2^35 (every intermediate product stays an exact integer of a Double)
Function DupHash(s As String) As Double
    Dim i As Long, h As Double, p As Double, q As Double
    p = 34359738337
    h = 7
    For i = 1 To Len(s)
        h = h * 131071 + Asc(Mid(s, i, 1))
        q = Int(h / p)
        h = h - q * p
        If h < 0 Then h = h + p
        If h >= p Then h = h - p
    Next i
    DupHash = h + 1
End Function

' the first other live receipt with the same key: "ЕИ-… (строка N)", or ""
Function FindDuplicate(keyText As String, exceptN As Long) As String
    Dim h As Double, n As Long, start As Long, i As Integer, d As Variant, r As Long
    h = DupHash(keyText)
    start = 0
    For i = 1 To 50
        n = FindDupEI(h, start)
        If n < 1 Then Exit Function
        If n <> exceptN Then
            d = RcvRow(n)
            If CStr(d(RV_STATE)) = RV_LIVE And CStr(d(RV_DUPKEY)) = keyText Then
                r = ReceiptRowOf(WmsIssue.EiCanon(n), d(RV_ROW))
                FindDuplicate = WmsIssue.EiCanon(n) & IIf(r >= 1, " (строка " & (r + 1) & ")", "")
                Exit Function
            End If
        End If
        start = n
    Next i
End Function

' ================================================================ a new EI (spec §4: only after the safe maximum)

' "" when EI n (= NEXT_EI) may be created now: it is not in «Наличие», not among the receipts, not in «Заказы» (a migrated
' old code), and no EI at or above NEXT_EI exists anywhere; otherwise the reason (the receipt is refused)
Function NewEIProblem(n As Long) As String
    Dim canon As String, last As Long
    canon = WmsIssue.EiCanon(n)
    If n < 1 Or n >= MAX_SHEET_ROW Then
        NewEIProblem = "номера ЕИ исчерпаны (NEXT_EI " & n & ")"
        Exit Function
    End If
    last = LastRow(WmsIssue.StockSheet())
    If last >= n Then
        NewEIProblem = "в «" & SH_STOCK & "» уже есть строки с номером ЕИ не меньше NEXT_EI " & n & " (последняя строка реестра " & (last + 1) _
            & ") — счётчик ЕИ отстаёт от реестра, новый ЕИ не выдаётся. Нужна проверка ответственным (самопроверка)"
        Exit Function
    End If
    last = LastRow(RcvSheet())
    If last >= n Then
        NewEIProblem = "среди проведённых приходов уже есть ЕИ с номером не меньше NEXT_EI " & n & " — счётчик ЕИ отстаёт, новый ЕИ не выдаётся"
        Exit Function
    End If
    If CountEI(canon) > 0 Then
        NewEIProblem = canon & " уже записан в «" & SH_ORDERS & "» (старый код?) — новый ЕИ с этим номером не выдаётся. Нужна проверка ответственным"
        Exit Function
    End If
    NewEIProblem = ""
End Function

' ================================================================ input checks shared by the preview, «Проверить» and posting

' a date from a cell (a Calc date 2000–2099 or a text dd.mm.yyyy); False with msg (naming the field) otherwise
Function ParseDateCellAs(oCell As Object, sField As String, ByRef d As Double, ByRef msg As String) As Boolean
    Dim ft As Long
    d = 0
    msg = ""
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        msg = "не указана " & sField
    Case com.sun.star.table.CellContentType.TEXT
        ParseDateCellAs = ParseDateTextAs(oCell.getString(), sField, d, msg)
    Case com.sun.star.table.CellContentType.VALUE
        ft = gDoc.getNumberFormats().getByKey(oCell.NumberFormat).Type
        If (ft And com.sun.star.util.NumberFormat.DATE) = 0 Then
            msg = sField & " должна быть датой, например 25.09.2026"
        Else
            ParseDateCellAs = DateOk(Int(oCell.getValue()), sField, d, msg)
        End If
    Case Else
        msg = "формула в поле «" & sField & "» не допускается"
    End Select
End Function

Function ParseDateTextAs(ByVal s As String, sField As String, ByRef d As Double, ByRef msg As String) As Boolean
    Dim a As Variant, dd As Integer, mm As Integer, yy As Integer
    s = Trim(s)
    a = Split(s, ".")
    msg = sField & ": «" & s & "» не является датой (формат дд.мм.гггг)"
    If UBound(a) <> 2 Then Exit Function
    If Not IsDigits(CStr(a(0))) Or Not IsDigits(CStr(a(1))) Or Not IsDigits(CStr(a(2))) Then Exit Function
    If Len(a(0)) > 2 Or Len(a(1)) > 2 Or Len(a(2)) <> 4 Then Exit Function
    dd = CInt(a(0))
    mm = CInt(a(1))
    yy = CInt(a(2))
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function
    If yy < 2000 Or yy > 2099 Then
        msg = sField & " вне допустимого диапазона 2000–2099"
        Exit Function
    End If
    If dd > 30 And (mm = 4 Or mm = 6 Or mm = 9 Or mm = 11) Then Exit Function
    If mm = 2 Then
        If dd > 29 Then Exit Function
        If dd = 29 And Not ((yy Mod 4 = 0 And yy Mod 100 <> 0) Or yy Mod 400 = 0) Then Exit Function
    End If
    ParseDateTextAs = DateOk(CDbl(DateSerial(yy, mm, dd)), sField, d, msg)
End Function

Private Function DateOk(x As Double, sField As String, ByRef d As Double, ByRef msg As String) As Boolean
    If x < CDbl(DateSerial(2000, 1, 1)) Or x > CDbl(DateSerial(2099, 12, 31)) Then
        msg = sField & " вне допустимого диапазона 2000–2099"
        Exit Function
    End If
    d = x
    msg = ""
    DateOk = True
End Function

' a quantity from a cell through the locale-independent parser (spec §12), with the field name and MAX_QTY
Function ParseQtyAs(oCell As Object, sField As String, ByRef q As Double, ByRef msg As String) As Boolean
    Dim st As Integer
    st = ParseQtyCell(oCell, q, msg)
    If st <> 0 Then
        msg = sField & ": " & msg
        Exit Function
    End If
    If q > SysNum(SK_MAX_QTY) Then
        msg = sField & ": больше порога MAX_QTY (" & SysStr(SK_MAX_QTY) & ")"
        Exit Function
    End If
    ParseQtyAs = True
End Function

' a quantity given as text (dialog) or a number (test)
Function ParseQtyValueAs(v As Variant, sField As String, ByRef q As Double, ByRef msg As String) As Boolean
    Dim st As Integer, x As Double
    If VarType(v) = 8 Then
        st = ParseQtyText(CStr(v), q, msg)
    Else
        x = CDbl(v)
        If x <= 0 Then
            st = 5
            msg = "количество должно быть больше 0"
        ElseIf Abs(x * 1000 - Int(x * 1000 + 0.5)) > 0.000001 Then
            st = 4
            msg = "не более " & QTY_MAX_DECIMALS & " знаков после запятой"
        Else
            q = Int(x * 1000 + 0.5) / 1000
        End If
    End If
    If st <> 0 Then
        msg = sField & ": " & msg
        Exit Function
    End If
    If q > SysNum(SK_MAX_QTY) Then
        msg = sField & ": больше порога MAX_QTY (" & SysStr(SK_MAX_QTY) & ")"
        Exit Function
    End If
    ParseQtyValueAs = True
End Function

' J «Цена»: empty, or a number ≥ 0 (the same parser as quantities; 0 allowed). No accounting rule is applied.
Function ParsePriceCell(oCell As Object, ByRef x As Double, ByRef has As Boolean, ByRef msg As String) As Boolean
    Dim st As Integer, ft As Long
    x = 0
    has = False
    msg = ""
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.EMPTY
        ParsePriceCell = True
        Exit Function
    Case com.sun.star.table.CellContentType.TEXT
        ParsePriceCell = ParsePriceText(oCell.getString(), x, has, msg)
    Case com.sun.star.table.CellContentType.VALUE
        ft = gDoc.getNumberFormats().getByKey(oCell.NumberFormat).Type
        If (ft And (2 Or 4 Or 32 Or 64 Or 128 Or 1024)) <> 0 Then
            msg = "Цена: ячейка распознана Calc как дата/процент/дробь — введите число, например 125,50"
        ElseIf oCell.getValue() < 0 Then
            msg = "Цена: отрицательное значение не допускается"
        ElseIf Abs(oCell.getValue() * 1000 - Int(oCell.getValue() * 1000 + 0.5)) > 0.000001 Then
            msg = "Цена: не более " & QTY_MAX_DECIMALS & " знаков после запятой"
        Else
            x = Int(oCell.getValue() * 1000 + 0.5) / 1000
            has = True
            ParsePriceCell = True
        End If
    Case Else
        msg = "Цена: формула не допускается"
    End Select
End Function

Function ParsePriceText(s As String, ByRef x As Double, ByRef has As Boolean, ByRef msg As String) As Boolean
    Dim st As Integer
    x = 0
    has = False
    If Trim(s) = "" Then
        ParsePriceText = True
        Exit Function
    End If
    st = ParseQtyText(s, x, msg)
    If st = 5 Then
        ' zero: a price may be 0
        x = 0
        st = 0
    End If
    If st <> 0 Then
        msg = "Цена: " & msg
        Exit Function
    End If
    has = True
    msg = ""
    ParsePriceText = True
End Function

' L «Площадка / Поставщик»: "" or the name of the special receipt (not an ordinary receipt, next stage)
Function SpecialSupplier(s As String) As String
    Dim a As Variant, i As Integer, t As String
    t = LCase(Trim(s))
    If t = "" Then Exit Function
    a = Split(SPECIAL_SUPPLIERS, "|")
    For i = 0 To UBound(a)
        If t = a(i) Then
            SpecialSupplier = Trim(s)
            Exit Function
        End If
    Next i
End Function

' the first problem among the filled inputs of an open order row (preview); "" when none
Function OpenRowProblem(r As Long) As String
    Dim sh As Object, c As Object, q As Double, d As Double, msg As String, has As Boolean, x As Double, sp As String
    sh = OrdersSheet()
    sp = SpecialSupplier(sh.getCellByPosition(OC_SUPPLIER, r).getString())
    If sp <> "" Then
        OpenRowProblem = "«" & sp & "» — специальный приход: он проводится на листе «" & SH_SPECIAL & "» (тип прихода «" & sp & "»), " _
            & "обычным приходом «Заказов» не проводится"
        Exit Function
    End If
    c = sh.getCellByPosition(OC_ORDQTY, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not ParseQtyAs(c, "Заказанное количество", q, msg) Then
            OpenRowProblem = msg
            Exit Function
        End If
    End If
    c = sh.getCellByPosition(OC_FACT, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not ParseQtyAs(c, "Фактическое количество", q, msg) Then
            OpenRowProblem = msg
            Exit Function
        End If
    End If
    c = sh.getCellByPosition(OC_DOCQTY, r)
    If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        If Not ParseQtyAs(c, "Количество по документу", q, msg) Then
            OpenRowProblem = msg
            Exit Function
        End If
    End If
    If Not ParsePriceCell(sh.getCellByPosition(OC_PRICE, r), x, has, msg) Then
        OpenRowProblem = msg
        Exit Function
    End If
    If Not DateIfFilled(sh.getCellByPosition(OC_RDATE, r), "дата поступления", msg) Then
        OpenRowProblem = msg
    ElseIf Not DateIfFilled(sh.getCellByPosition(OC_DDATE, r), "дата документа", msg) Then
        OpenRowProblem = msg
    ElseIf Not DateIfFilled(sh.getCellByPosition(OC_ODATE, r), "дата заказа", msg) Then
        OpenRowProblem = msg
    ElseIf Not DateIfFilled(sh.getCellByPosition(OC_EDATE, r), "ожидаемая дата поступления", msg) Then
        OpenRowProblem = msg
    End If
End Function

Private Function DateIfFilled(c As Object, sField As String, ByRef msg As String) As Boolean
    Dim d As Double
    msg = ""
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then
        DateIfFilled = True
    Else
        DateIfFilled = ParseDateCellAs(c, sField, d, msg)
    End If
End Function

' the expected date Q of row r as a serial (0 when empty or not a valid date)
Function ExpectedDate(r As Long) As Double
    ExpectedDate = QSerial(OrdersSheet().getCellRangeByPosition(OC_EDATE, r, OC_EDATE, r).getDataArray()(0)(0))
End Function

' the expected date from a cell value (as getDataArray returns it): a date of 2000–2099 given as a number (the column is
' date-formatted) or as a text dd.mm.yyyy; 0 = no expected date. The one reading of Q for the operations, the change
' handler and «Обновить статусы», so all three agree on «просрочено».
Function QSerial(v As Variant) As Double
    Dim d As Double, msg As String
    If VarType(v) = 5 Then
        If DateOk(Int(v), "дата", d, msg) Then QSerial = d
    ElseIf VarType(v) = 8 Then
        If v <> "" Then
            If ParseDateTextAs(CStr(v), "дата", d, msg) Then QSerial = d
        End If
    End If
End Function

' the status of an open order row: «Ожидается» or «Просрочено» (Q before today); "" for a row without order data
Function OpenRowStatus(r As Long) As String
    Dim sh As Object, e As Double
    sh = OrdersSheet()
    If sh.getCellByPosition(OC_ORDER, r).getString() = "" And sh.getCellByPosition(OC_NAME, r).getString() = "" _
        And sh.getCellByPosition(OC_ORDQTY, r).getString() = "" Then Exit Function
    e = ExpectedDate(r)
    If e > 0 And e < Today() Then OpenRowStatus = OS_OVERDUE Else OpenRowStatus = OS_WAITING
End Function

' ================================================================ light change handler (spec §13, D-008)

' «Заказы» sheet event "Content changed": the status and the first input problem of open rows (W, Y). Leaves at once for
' columns that do not matter, never scans the sheet, is guarded against re-entry and ignores changes made by operations.
Sub OnOrdersChange(Optional oTarget As Variant)
    Dim addrs As Variant, i As Long
    If mInHandler Or gBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If oTarget.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        addrs = oTarget.getRangeAddresses()
    Else
        addrs = Array(oTarget.getRangeAddress())
    End If
    For i = 0 To UBound(addrs)
        HandleRange(addrs(i))
    Next i
    Exit Sub
EH:
    mInHandler = False
End Sub

Private Sub HandleRange(ra As Variant)
    Dim r As Long, r1 As Long, n As Long, um As Object, inCtx As Boolean, last As Long
    If ra.EndRow < 1 Then Exit Sub
    ' the inputs A..U and V (a key that appeared by paste) matter; W X Y AB are written by WMS, Z AA are free text
    If ra.StartColumn > OC_EI Then Exit Sub
    mInHandler = True
    WmsInit()
    r1 = ra.EndRow
    If r1 - ra.StartRow + 1 > PREVIEW_MAX_ROWS Then
        last = LastRow(OrdersSheet())
        If r1 > last Then r1 = last
    End If
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: проверка строки заказа")
    inCtx = True
    For r = IIf(ra.StartRow < 1, 1, ra.StartRow) To r1
        n = n + 1
        ' more rows at once (paste, fill): the rest keeps its status until «Обновить статусы» (spec §13)
        If n > PREVIEW_MAX_ROWS Then Exit For
        PreviewOrderRow(r, (ra.StartColumn <= OC_EI And ra.EndColumn >= OC_EI), (ra.StartColumn <= OC_EDATE And ra.EndColumn >= OC_EDATE))
    Next r
    um.leaveUndoContext()
    mInHandler = False
    Exit Sub
EH:
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    mInHandler = False
End Sub

' W and Y of an open row: the status by the expected date, the first problem of the filled inputs. Of a row with a
' receipt only W of the source row of its position is recalculated, and only when its expected date Q (open after the
' receipt) was changed: a partly received position becomes or stops being overdue at once (D-047). Cancelled rows and
' copies are never touched here; a key in V that WMS did not write marks the row КОПИЯ.
Sub PreviewOrderRow(r As Long, bKeyTouched As Boolean, bDateTouched As Boolean)
    Dim sh As Object, kind As String, p As String, nChanged As Long
    sh = OrdersSheet()
    kind = OrderRowKind(r)
    Select Case kind
    Case "FOREIGN"
        CopyCheckOrder(r)
    Case "EMPTY"
        ClearDerived(r)
    Case "OPEN"
        SetIfDiff(sh.getCellByPosition(OC_STATUS, r), OpenRowStatus(r))
        p = OpenRowProblem(r)
        If p <> "" Then p = "Ошибка: " & p
        SetIfDiff(sh.getCellByPosition(OC_CTL, r), p)
        SetIfDiff(sh.getCellByPosition(OC_STOCK, r), "")
        SetIfDiff(sh.getCellByPosition(OC_DUP, r), "")
    Case "RECEIVED", "STORNO"
        If bKeyTouched And gKmoved Then CopyCheckOrder(r)
        If bDateTouched And gKsrc Then RefreshRow(r, nChanged, Today(), p)
    End Select
End Sub

Private Sub ClearDerived(r As Long)
    Dim sh As Object, cols As Variant, i As Integer
    sh = OrdersSheet()
    cols = Array(OC_STATUS, OC_STOCK, OC_CTL, OC_DUP)
    For i = 0 To UBound(cols)
        If sh.getCellByPosition(cols(i), r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(cols(i), r))
    Next i
End Sub

Sub SetIfDiff(c As Object, v As Variant)
    Select Case VarType(v)
    Case 2, 3, 4, 5, 6
        If c.getType() = com.sun.star.table.CellContentType.VALUE Then
            If Abs(c.getValue() - CDbl(v)) < 0.0000001 Then Exit Sub
        End If
        c.setValue(CDbl(v))
    Case Else
        If CStr(v) = "" Then
            If c.getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(c)
        ElseIf c.getString() <> CStr(v) Or c.getType() <> com.sun.star.table.CellContentType.TEXT Then
            c.setString(CStr(v))
        End If
    End Select
End Sub

' a row whose V holds an EI that WMS did not register there (a pasted or forged copy) is marked КОПИЯ (spec §15)
Sub CopyCheckOrder(r As Long)
    Dim sh As Object, kind As String, why As String
    sh = OrdersSheet()
    kind = OrderRowKind(r)
    If kind = "FOREIGN" Then
        why = gKwhy
    ElseIf (kind = "RECEIVED" Or kind = "STORNO") And gKmoved Then
        ' the key is registered at another row that no longer holds it: this row is the original moved by an insertion
        Exit Sub
    Else
        Exit Sub
    End If
    sh.getCellByPosition(OC_CTL, r).setString("КОПИЯ: " & why & " — не проводится, очистите строку кнопкой «Очистить»")
End Sub

' ================================================================ «Обновить статусы» (spec: «Просрочено» depends on Q and the date)
' Only the rows whose status can change with the date are examined, found by the Calc engine without a loop over the sheet
' (for each status: one MATCH for its first row, then one query of the W cells that differ from it):
'  - W «Ожидается» and «Просрочено» — every such row (their status depends on Q only);
'  - W «Частично получено» — only the rows whose expected date Q has passed, W «Частично получено / просрочено» — only the
'    rows whose Q has not passed (edited or cleared): Q is read only where these rows meet a filled Q cell (one query of the
'    filled Q cells, one read per common block), so partly received positions cost nothing while their Q stays as it was;
'  - order rows without a status yet (queries of empty W cells and filled B cells).
' The derived status W is written directly (not an accounting change, not journaled); every operation also recalculates
' the statuses of the positions it touches.

Function RefreshStatuses() As String
    Dim sh As Object, last As Long, t0 As Long, words As Variant, w As Integer, blocks As Variant, got As Variant, qb As Variant
    Dim i As Long, j As Long, rr As Long, nChanged As Long, nOver As Long, nOpen As Long, nPartOver As Long, sOld As String
    Dim flags As Long, st As String, more As Boolean, rows() As Long, nr As Long, bs() As Long, be() As Long, nb As Long
    Dim tday As Double, found As Boolean
    WmsInit()
    t0 = GetSystemTicks()
    sh = OrdersSheet()
    last = LastRow(sh)
    If last < 1 Then
        gRefreshNote = "статусы заказов: строк заказов нет"
        RefreshStatuses = gRefreshNote
        Exit Function
    End If
    tday = Today()
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    ' 1. the candidate rows are collected first (the queries see the statuses before any change), then recalculated
    ReDim rows(255)
    words = Array(OS_WAITING, OS_OVERDUE)
    For w = 0 To UBound(words)
        nb = WordBlocks(sh, last, CStr(words(w)), bs, be)
        For i = 0 To nb - 1
            For rr = bs(i) To be(i)
                AddRow(rows, nr, rr)
            Next rr
        Next i
    Next w
    ' partly received positions: only the rows where the date condition and the shown status disagree — found by one array
    ' formula of the Calc engine; if it cannot be used, by reading the rows of these statuses where Q is filled
    gRefreshPath = "FORMULA"
    If gRefreshNoFormula Then
        found = False
    Else
        found = PartialDueRows(last, tday, rows, nr, nPartOver)
    End If
    If Not found Then
        gRefreshPath = "ROWS"
        qb = sh.getCellRangeByPosition(OC_EDATE, 1, OC_EDATE, last).queryContentCells(flags).getRangeAddresses()
        nb = WordBlocks(sh, last, OS_PARTIAL, bs, be)
        AddByDate(sh, bs, be, nb, qb, True, tday, rows, nr)
        nb = WordBlocks(sh, last, OS_PARTIAL_OVERDUE, bs, be)
        nPartOver = 0
        For i = 0 To nb - 1
            nPartOver = nPartOver + be(i) - bs(i) + 1
        Next i
        AddByDate(sh, bs, be, nb, qb, False, tday, rows, nr)
    End If
    ' order rows without a status (entered with macros off, or more rows at once than the preview handles)
    blocks = sh.getCellRangeByPosition(OC_STATUS, 1, OC_STATUS, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= REFRESH_MAX_BLOCKS Then
            more = True
            Exit For
        End If
        got = sh.getCellRangeByPosition(OC_NAME, blocks(i).StartRow, OC_NAME, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            For rr = got(j).StartRow To got(j).EndRow
                AddRow(rows, nr, rr)
            Next rr
        Next j
    Next i
    ' 2. recalculation
    For i = 0 To nr - 1
        st = RefreshRow(rows(i), nChanged, tday, sOld)
        If st = OS_OVERDUE Then nOver = nOver + 1
        If st = OS_OVERDUE Or st = OS_WAITING Then nOpen = nOpen + 1
        If st = OS_PARTIAL_OVERDUE And sOld <> OS_PARTIAL_OVERDUE Then nPartOver = nPartOver + 1
        If st <> OS_PARTIAL_OVERDUE And sOld = OS_PARTIAL_OVERDUE Then nPartOver = nPartOver - 1
    Next i
    gRefreshNote = "статусы заказов: ожидаемых позиций " & nOpen & ", из них просрочено " & nOver _
        & "; частично получено и просрочено " & nPartOver & "; обновлено статусов " & nChanged _
        & IIf(more, " (строк без статуса больше, чем проверено за раз — нажмите «Обновить статусы» ещё раз)", "") & ", " & (GetSystemTicks() - t0) & " мс"
    RefreshStatuses = gRefreshNote
End Function

' The rows of «Заказы» (1..last) whose status must be checked for the date: «Частично получено» with a passed expected date
' Q, «Частично получено / просрочено» whose Q has not passed or is not a date (a text Q is always checked: RefreshRow reads
' it). One array formula over the used rows — TEXTJOIN of the row numbers, 30 ms on 100 000 rows instead of reading every
' partly received row — written into the unlocked scratch cell _IDX!B<IX_LIST>, read and cleared at once (no Undo, the
' book does not become «modified»); a tag proves the value was calculated for this call. nPO: the rows shown «Частично
' получено / просрочено» now. False when the formula could not be used.
Private Function PartialDueRows(last As Long, tday As Double, rows() As Long, ByRef nr As Long, ByRef nPO As Long) As Boolean
    Dim f As String, w As String, q As String, s As String, a As Variant, cond As String, p As Long, k As Long, ok As Boolean
    On Error GoTo EH
    w = "$'" & SH_ORDERS & "'.$W$2:$W$" & (last + 1)
    q = "$'" & SH_ORDERS & "'.$Q$2:$Q$" & (last + 1)
    ' a valid date of 2000–2099 before today, as QSerial reads a number
    cond = "ISNUMBER(" & q & ")*(" & q & ">=" & Format(CDbl(DateSerial(2000, 1, 1)), "0") & ")*(" & q & "<=" _
        & Format(CDbl(DateSerial(2099, 12, 31)), "0") & ")*(" & q & "<" & Format(tday, "0") & ")"
    f = "COUNTIF(" & w & ";""" & OS_PARTIAL_OVERDUE & """)&""|""&TEXTJOIN("";"";1;IF((" & w & "=""" & OS_PARTIAL _
        & """)*(" & cond & "+ISTEXT(" & q & "))+(" & w & "=""" & OS_PARTIAL_OVERDUE & """)*(1-" & cond & ");ROW(" & w & ")-1;""""))"
    s = ScratchEval(f, ok)
    If Not ok Then Exit Function
    p = InStr(s, "|")
    If p < 2 Then Exit Function
    nPO = CLng(Left(s, p - 1))
    s = Mid(s, p + 1)
    If s <> "" Then
        a = Split(s, ";")
        For k = 0 To UBound(a)
            AddRow(rows, nr, CLng(a(k)))
        Next k
    End If
    PartialDueRows = True
    Exit Function
EH:
    PartialDueRows = False
End Function

' the blocks of rows 1..last whose W equals s (one MATCH for the first such row, then one query of the W cells that differ
' from it); returns the number of blocks, their first and last rows in bs and be
Private Function WordBlocks(sh As Object, last As Long, s As String, bs() As Long, be() As Long) As Long
    Dim r0 As Long, a As New com.sun.star.table.CellAddress, diff As Variant, i As Long, prev As Long, nb As Long
    r0 = FindWRow(s)
    If r0 < 1 Then Exit Function
    a.Sheet = sh.getRangeAddress().Sheet
    a.Column = OC_STATUS
    a.Row = r0
    diff = sh.getCellRangeByPosition(OC_STATUS, 1, OC_STATUS, last).queryColumnDifferences(a).getRangeAddresses()
    ReDim bs(UBound(diff) + 1)
    ReDim be(UBound(diff) + 1)
    prev = 1
    For i = 0 To UBound(diff)
        If diff(i).StartRow > prev Then
            bs(nb) = prev
            be(nb) = diff(i).StartRow - 1
            nb = nb + 1
        End If
        If diff(i).EndRow + 1 > prev Then prev = diff(i).EndRow + 1
    Next i
    If prev <= last Then
        bs(nb) = prev
        be(nb) = last
        nb = nb + 1
    End If
    WordBlocks = nb
End Function

' adds the rows of the blocks bs/be whose expected date Q has passed (bPassed) or has not passed (Not bPassed). Both block
' lists are sorted: they are merged in one pass, Q is read only where a block meets a filled Q cell (qb); a row without Q
' has no expected date, so it has not passed.
Private Sub AddByDate(sh As Object, bs() As Long, be() As Long, nb As Long, qb As Variant, bPassed As Boolean, tday As Double, _
        rows() As Long, ByRef nr As Long)
    Dim i As Long, j As Long, rr As Long, pos As Long, qs As Long, qe As Long, got As Variant, e As Double, passed As Boolean
    For i = 0 To nb - 1
        pos = bs(i)
        Do While pos <= be(i)
            ' the first block of filled Q that ends at or after pos
            Do While j <= UBound(qb)
                If qb(j).EndRow >= pos Then Exit Do
                j = j + 1
            Loop
            qs = be(i) + 1
            If j <= UBound(qb) Then
                If qb(j).StartRow <= be(i) Then
                    qs = qb(j).StartRow
                    If qs < pos Then qs = pos
                End If
            End If
            ' rows pos .. qs - 1: no expected date
            If Not bPassed Then
                For rr = pos To qs - 1
                    AddRow(rows, nr, rr)
                Next rr
            End If
            If qs > be(i) Then Exit Do
            qe = qb(j).EndRow
            If qe > be(i) Then qe = be(i)
            got = sh.getCellRangeByPosition(OC_EDATE, qs, OC_EDATE, qe).getDataArray()
            For rr = qs To qe
                e = QSerial(got(rr - qs)(0))
                passed = (e > 0 And e < tday)
                If passed = bPassed Then AddRow(rows, nr, rr)
            Next rr
            pos = qe + 1
        Loop
    Next i
End Sub

Private Sub AddRow(rows() As Long, ByRef nr As Long, r As Long)
    If nr > UBound(rows) Then
        ReDim Preserve rows(2 * nr + 1)
    End If
    rows(nr) = r
    nr = nr + 1
End Sub

' recalculates W of one row whose status depends on the date; returns the status it now shows, sOld the one it showed
' (nChanged counts writes). The row is read once; only a row with a receipt (V) needs the service data of its position.
Private Function RefreshRow(r As Long, ByRef nChanged As Long, tday As Double, ByRef sOld As String) As String
    Dim sh As Object, d As Variant, want As String, od As Variant, rv As Variant, n As Long, i As Integer, has As Boolean, e As Double
    sh = OrdersSheet()
    d = sh.getCellRangeByPosition(0, r, OC_LAST, r).getDataArray()(0)
    sOld = CStr(d(OC_STATUS))
    RefreshRow = sOld
    If Left(CStr(d(OC_CTL)), 5) = "КОПИЯ" Then Exit Function
    If CStr(d(OC_EI)) = "" Then
        If CStr(d(OC_STATUS)) = OS_CANCELLED Then Exit Function
        For i = 0 To OC_PLACE
            If CStr(d(i)) <> "" Then
                has = True
                Exit For
            End If
        Next i
        If Not has Then
            want = ""
        ElseIf CStr(d(OC_ORDER)) = "" And CStr(d(OC_NAME)) = "" And CStr(d(OC_ORDQTY)) = "" Then
            want = ""
        Else
            e = QSerial(d(OC_EDATE))
            If e > 0 And e < tday Then want = OS_OVERDUE Else want = OS_WAITING
        End If
    Else
        ' a row with a receipt: only the source row of its position shows a date-dependent status; the copy and move
        ' checks of OrderRowKind are not needed for a display value (a marked copy was skipped above)
        If Not StrictEI(CStr(d(OC_EI)), n) Then Exit Function
        If Not RcvRegistered(n, rv) Then Exit Function
        If CStr(rv(RV_KIND)) <> RV_SRC Then Exit Function
        If VarType(rv(RV_OL)) <> 5 Then Exit Function
        If Not OlValid(CLng(rv(RV_OL)), od) Then Exit Function
        want = PositionStatus(od(OD_ORD), od(OD_RCV), od(OD_NODOC), CStr(od(OD_CANCEL)), QSerial(d(OC_EDATE)))
    End If
    If CStr(d(OC_STATUS)) <> want Then
        SetIfDiff(sh.getCellByPosition(OC_STATUS, r), want)
        nChanged = nChanged + 1
    End If
    RefreshRow = want
End Function

' end of a clean start: the statuses that changed with the date are refreshed (spec: recalculated at the opening)
Sub RefreshAtStartup()
    Dim s As String
    On Error GoTo EH
    If gState <> "CLEAN" Then Exit Sub
    If Not gDoc.Sheets.hasByName(SH_ORDERS) Then Exit Sub
    s = RefreshStatuses()
    AddNote(s)
    Exit Sub
EH:
    AddNote("статусы заказов не обновлены: " & Error$)
End Sub

' ================================================================ «Главная»: receipts not posted yet (spec §23)

' rows with a fact (F) but no EI (V): the goods were counted but the receipt was not posted. Two Calc queries per block
' instead of a row loop (like WmsIssue.UnpostedCount). -1 when it fails; gRcvUnpostedMore when the count is a lower bound.

Function UnpostedReceipts() As Long
    Dim sh As Object, last As Long, blocks As Variant, got As Variant, i As Long, j As Long, n As Long, flags As Long
    On Error GoTo EH
    gRcvUnpostedMore = False
    sh = OrdersSheet()
    last = LastRow(sh)
    If last < 1 Then Exit Function
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    blocks = sh.getCellRangeByPosition(OC_EI, 1, OC_EI, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= UNPOSTED_MAX_BLOCKS Then
            gRcvUnpostedMore = True
            Exit For
        End If
        got = sh.getCellRangeByPosition(OC_FACT, blocks(i).StartRow, OC_FACT, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            n = n + got(j).EndRow - got(j).StartRow + 1
        Next j
    Next i
    UnpostedReceipts = n
    Exit Function
EH:
    UnpostedReceipts = -1
End Function

' ================================================================ «Очистить» — an open order row or a copy (no accounting change)

Function ClearOrderRow(r As Long) As String
    Dim sh As Object, kind As String, wasProt As Boolean, i As Integer, pr As Object, n As Long, d As Variant
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        ClearOrderRow = "ERR:выберите строку заказа (не заголовок)"
        Exit Function
    End If
    If gBusy Then
        ClearOrderRow = "BUSY:операция уже выполняется"
        Exit Function
    End If
    sh = OrdersSheet()
    kind = OrderRowKind(r)
    Select Case kind
    Case "RECEIVED"
        ClearOrderRow = "ERR:строка — проведённый приход " & gKcanon & " — для отмены используйте «Удалить»"
        Exit Function
    Case "STORNO"
        ClearOrderRow = "ERR:приход " & gKcanon & " удалён (сторно) — строка хранит историю и не очищается"
        Exit Function
    Case "CANCELLED"
        ClearOrderRow = "ERR:позиция отменена — строка хранит историю и не очищается"
        Exit Function
    Case "COPY"
        ' defence: never the registered row of a receipt, even when it carries a copy mark
        If StrictEI(sh.getCellByPosition(OC_EI, r).getString(), n) Then
            If RcvRegistered(n, d) Then
                If VarType(d(RV_ROW)) = 5 Then
                    If CLng(d(RV_ROW)) = r Then
                        ClearOrderRow = "ERR:строка — зарегистрированный приход " & WmsIssue.EiCanon(n) & "; копией является другая строка с этим ЕИ"
                        Exit Function
                    End If
                End If
            End If
        End If
    End Select
    mInHandler = True
    On Error GoTo EH
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    sh.getCellRangeByPosition(0, r, OC_LAST, r).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME _
        + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA)
    ' a row inserted between posted rows or a copy carries the protection of its origin: back to the input layout
    For i = 0 To OC_LAST
        pr = sh.getCellByPosition(i, r).CellProtection
        pr.IsLocked = (Mid(ORDER_LOCKS_OPEN, i + 1, 1) = "1")
        sh.getCellByPosition(i, r).CellProtection = pr
    Next i
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
    ClearOrderRow = "OK:" & IIf(kind = "COPY" Or kind = "FOREIGN", "копия очищена", "строка очищена")
    Exit Function
EH:
    ClearOrderRow = "ERR:не удалось очистить строку: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
End Function

' handler guard for the operation module (its own writes into «Заказы» outside ApplyOperation, e.g. «Проверить»)
Sub HandlerOff(b As Boolean)
    mInHandler = b
End Sub

' ================================================================ test seam (inert unless _SYS MODE = TEST)

' mode 1: «Обновить статусы» finds the partly received rows by reading them (the fallback), 0: as in production
Function TestRefreshMode(mode As Integer) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestRefreshMode = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gRefreshNoFormula = (mode = 1)
    TestRefreshMode = "OK"
End Function

Function TestRefreshPath() As String
    TestRefreshPath = gRefreshPath
End Function

Function TestSetToday(sDate As String) As String
    Dim d As Double, msg As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSetToday = "REFUSED:не тестовая книга"
        Exit Function
    End If
    If sDate = "" Then
        gTodayOverride = 0
    ElseIf ParseDateTextAs(sDate, "дата", d, msg) Then
        gTodayOverride = d
    Else
        TestSetToday = "ERR:" & msg
        Exit Function
    End If
    TestSetToday = "OK"
End Function

Function TestLookupEI(canon As String) As String
    TestLookupEI = FindEIRow(canon, 0) & "|" & CountEI(canon)
End Function
