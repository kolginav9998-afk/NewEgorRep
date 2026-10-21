' WmsLegacy — PRIME 0.8: the registry of the transfer (_LEG), «Исправить ЕИ переноса» (MIGRATE_FIX, with «Разобрать» for
' an EI «Требует разбора»), the rule «остаток = сумма движений» for every EI (BalanceCheck) and the sheet «Сверка
' переноса» (LegacyReport). The corrections of transferred receipts of «Заказы» (LEGACY_FIX) live in WmsReceipt.
' The rule: the balance of an EI equals its live receipts (order receipts, special receipts, the balance a MIGRATE brought)
' minus what the old table consumed before the transfer (_LEG «Израсходовано до переноса», for a live transferred receipt)
' minus the posted issues plus the live returns plus the live corrections (write-off, inventory).
Option Explicit

' the result of the last BalanceCheck (the report and the tests read them)
Global gLgEIs As Long
Global gLgLegacy As Long
Global gLgBad As Long
Global gLgFlags As Long

Function LegSheet() As Object
    LegSheet = gDoc.Sheets.getByName(SH_LEG)
End Function

' "" when the registry of the transfer and the sheet of its report exist with their headers (checked at every start)
Function LegSheetProblem() As String
    Dim sh As Object
    If Not gDoc.Sheets.hasByName(SH_LEG) Then
        LegSheetProblem = "нет служебного листа " & SH_LEG & " (реестр переноса)"
        Exit Function
    End If
    sh = LegSheet()
    If sh.getCellByPosition(LG_EI, 0).getString() <> "ЕИ" Or sh.getCellByPosition(LG_PRE, 0).getString() <> "Израсходовано до переноса" _
        Or sh.getCellByPosition(LG_ACT, 0).getString() <> "Последняя коррекция" Then
        LegSheetProblem = "служебный лист " & SH_LEG & ": заголовок повреждён"
        Exit Function
    End If
    If Not gDoc.Sheets.hasByName(SH_LEGREP) Then
        LegSheetProblem = "нет листа «" & SH_LEGREP & "»"
        Exit Function
    End If
    LegSheetProblem = ""
End Function

' the _LEG row of EI n (d = A..K): True when the EI came from the transfer of the old table
Function LegRow(n As Long, ByRef d As Variant) As Boolean
    If n < 1 Or n > MAX_SHEET_ROW Then Exit Function
    d = LegSheet().getCellRangeByPosition(0, n, LG_LAST, n).getDataArray()(0)
    LegRow = (CStr(d(LG_EI)) <> "")
End Function

' the writes of a new _LEG row (the operations of the transfer: LEGACY_RECEIPT, LEGACY_RECEIPT_ADD, MIGRATE); the seq is the
' one ApplyOperation gives the operation being planned
Sub PlanLegRow(n As Long, canon As String, sKind As String, qty0 As Double, bal0 As Double, sX As String, sSrc As String, sOrigin As String)
    PlanSetValue(SH_LEG, n, LG_EI, canon, False)
    PlanSetValue(SH_LEG, n, LG_KIND, sKind, False)
    PlanSetValue(SH_LEG, n, LG_QTY0, qty0, False)
    PlanSetValue(SH_LEG, n, LG_BAL0, bal0, False)
    PlanSetValue(SH_LEG, n, LG_PRE, WmsIssue.Round3(qty0 - bal0), False)
    If sX <> "" Then PlanSetValue(SH_LEG, n, LG_X, sX, False)
    If sSrc <> "" Then PlanSetValue(SH_LEG, n, LG_SRC, sSrc, False)
    If sOrigin <> "" Then PlanSetValue(SH_LEG, n, LG_ORIGIN, sOrigin, False)
    PlanSetValue(SH_LEG, n, LG_SEQ, SysNum(SK_LAST_SEQ) + 1, False)
    PlanSetValue(SH_LEG, n, LG_FIXES, 0, False)
End Sub

' the kind of source of a migrated EI as WMS writes it («Наличие».J), or "" when unknown
Private Function KnownStype(ByVal s As String) As String
    Dim names As Variant, i As Integer
    names = Split(STYPE_SUPPLIER & "|" & SRC_STYPES, "|")
    For i = 0 To UBound(names)
        If LCase(Trim(s)) = LCase(names(i)) Then
            KnownStype = names(i)
            Exit Function
        End If
    Next i
    KnownStype = ""
End Function

' the number of live special receipts (refills) of EI canon on «Иной приход» (_SPR)
Private Function LiveSpecialLines(canon As String) As Long
    Dim spr As Object, last As Long, r0 As Long, r1 As Long, i As Long, d As Variant, k As Long
    If Not gDoc.Sheets.hasByName(SH_SPR) Then Exit Function
    spr = gDoc.Sheets.getByName(SH_SPR)
    last = WmsOrders.LastRow(spr)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = spr.getCellRangeByPosition(0, r0, SR_STATE, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(d(i)(SR_EI)) = canon And CStr(d(i)(SR_STATE)) = RV_LIVE Then k = k + 1
        Next i
    Next r0
    LiveSpecialLines = k
End Function

' ================================================================ «Исправить ЕИ переноса» — MIGRATE_FIX
' An EI the transfer brought without an order receipt (MIGRATE: only its «Наличие» row) gets its card and its balance at
' the transfer corrected: name, article, unit (only while nothing moved), category, kind of source — an EI «Требует
' разбора» (kind «Иной приход») gets its real kind here, which is «Разобрать» for such EIs — and the balance at the transfer
' (the current balance moves by its change; an act is required then). A part («Детали») keeps one EI per article.
' OK:<seq>, SKIP:<why>, ERR:<why> or an ApplyOperation error.
Function MigrateFixEI(sEI As String, sName As String, sArt As String, sUnit As String, sCat As String, sType As String, sBal0 As String, _
    sAct As String, sActDate As String, sReason As String) As String
    Dim n As Long, canon As String, msg As String, why As String, lg As Variant, d As Variant, s As Double, b0 As Double, b2 As Double
    Dim s2 As Double, st As Integer, newType As String, oldType As String, chg As String, nIss As Long, nRet As Long, nAdj As Long
    Dim dAct As Double, actText As String, fixes As Double, key As String, oldKey As String, en As Long, ecanon As String, aRow As Long
    Dim oRow As Long, newState As String, a As Integer
    WmsInit()
    On Error GoTo EH
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        MigrateFixEI = "ERR:" & msg
        Exit Function
    End If
    If Not LegRow(n, lg) Then
        MigrateFixEI = "ERR:" & canon & " не перенесён из старой таблицы — его карточку исправляет приход, которым он создан"
        Exit Function
    End If
    If CStr(lg(LG_KIND)) <> LG_MIGRATE Then
        MigrateFixEI = "ERR:" & canon & " перенесён приходом по заказу — его исправляет «Исправить» на листе «" & SH_ORDERS & "» (позиция или коррекция переноса)"
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        MigrateFixEI = why
        Exit Function
    End If
    why = WmsIssue.StockRowProblem(n, canon)
    If why <> "" Then
        MigrateFixEI = "ERR:" & why
        Exit Function
    End If
    d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
    If Not WmsIssue.StockQty(n, s) Then
        MigrateFixEI = "ERR:остаток " & canon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If VarType(lg(LG_BAL0)) = 5 Then b0 = lg(LG_BAL0)
    If Trim(sName) = "" Or Trim(sUnit) = "" Then
        MigrateFixEI = "ERR:наименование и единица измерения обязательны"
        Exit Function
    End If
    newType = KnownStype(sType)
    If newType = "" Then
        MigrateFixEI = "ERR:тип источника «" & sType & "» неизвестен — выберите один из: " & Replace(STYPE_SUPPLIER & "|" & SRC_STYPES, "|", ", ")
        Exit Function
    End If
    oldType = CStr(d(SC_STYPE))
    If Trim(sBal0) = "" Then
        b2 = b0
    Else
        st = ParseQtyText(Trim(sBal0), b2, msg)
        If st = 5 And b2 = 0 And InStr(sBal0, "-") = 0 Then
            b2 = 0
        ElseIf st <> 0 Then
            MigrateFixEI = "ERR:остаток на момент переноса «" & sBal0 & "» — " & msg
            Exit Function
        End If
    End If
    If Trim(sName) <> Trim(CStr(d(SC_NAME))) Then chg = chg & "B"
    If Trim(sArt) <> Trim(CStr(d(SC_ART))) Then chg = chg & "C"
    If Trim(sUnit) <> Trim(CStr(d(SC_UNIT))) Then chg = chg & "D"
    If Trim(sCat) <> Trim(CStr(d(SC_CAT))) Then chg = chg & "G"
    If newType <> oldType Then chg = chg & "J"
    If Abs(b2 - b0) > 0.0000001 Then chg = chg & "Q"
    If chg = "" Then
        MigrateFixEI = "SKIP:ничего не изменилось"
        Exit Function
    End If
    If InStr(chg, "Q") > 0 Or Trim(sAct) <> "" Then
        If Trim(sAct) = "" Then
            MigrateFixEI = "ERR:остаток на момент переноса меняется только с основанием — укажите № акта и дату"
            Exit Function
        End If
        If Not WmsOrders.ParseDateTextAs(sActDate, "дата акта", dAct, msg) Then
            MigrateFixEI = "ERR:" & msg
            Exit Function
        End If
        actText = "акт № " & Trim(sAct) & " от " & Format(dAct, "DD.MM.YYYY") & IIf(Trim(sReason) <> "", ": " & Trim(sReason), "")
    End If
    s2 = WmsIssue.Round3(s + b2 - b0)
    If s2 < -0.0000001 Then
        MigrateFixEI = "ERR:после переноса из " & canon & " уже выдано или списано " & WmsIssue.QtyText(WmsIssue.Round3(b0 - s)) & " " & d(SC_UNIT) _
            & " — остаток на момент переноса не может быть меньше этого (сейчас остаток " & WmsIssue.QtyText(s) & ")"
        Exit Function
    End If
    If InStr(chg, "D") > 0 Then
        If Not WmsReturn.LiveDependents(canon, nIss, nRet) Or Not WmsAdjust.LiveAdjustments(canon, nAdj) Then
            MigrateFixEI = "ERR-SYS:не удалось проверить движения " & canon & " (формула движка Calc) — исправление не выполнено"
            Exit Function
        End If
        nAdj = nAdj + LiveSpecialLines(canon)
        If nIss + nRet + nAdj > 0 Then
            MigrateFixEI = "ERR:единицу измерения нельзя изменить: по " & canon & " есть действующие выдачи (" & nIss & "), возвраты (" & nRet _
                & ") или приходы и корректировки (" & nAdj & ") в «" & d(SC_UNIT) & "»"
            Exit Function
        End If
    End If
    ' a part keeps one EI per article: the index _ART follows the article; a part does not stop being a part here
    If oldType = SRC_STYPES_PART And newType <> SRC_STYPES_PART Then
        MigrateFixEI = "ERR:" & canon & " — деталь (один артикул — один ЕИ): тип «Детали» здесь не снимается"
        Exit Function
    End If
    oRow = -1
    If newType = SRC_STYPES_PART And (InStr(chg, "C") > 0 Or InStr(chg, "J") > 0) Then
        key = WmsSpecial.ArticleKey(sArt)
        If key = "" Then
            MigrateFixEI = "ERR:у детали должен быть артикул"
            Exit Function
        End If
        a = WmsSpecial.ArtLookup(key, en, ecanon, aRow, why)
        If a = 1 And en <> n Then
            MigrateFixEI = "ERR:артикул «" & sArt & "» уже у " & ecanon & " — один артикул детали — один ЕИ"
            Exit Function
        ElseIf a >= 2 Then
            MigrateFixEI = "ERR:" & why
            Exit Function
        End If
        If oldType = SRC_STYPES_PART Then
            oldKey = WmsSpecial.ArticleKey(CStr(d(SC_ART)))
            a = WmsSpecial.ArtLookup(oldKey, en, ecanon, oRow, why)
            If a <> 1 Or en <> n Then oRow = -1
        End If
        If oRow < 1 Then oRow = WmsOrders.LastRow(WmsSpecial.ArtSheet()) + 1
    End If
    If VarType(lg(LG_FIXES)) = 5 Then fixes = lg(LG_FIXES)
    newState = CStr(d(SC_STATE))
    If InStr(chg, "J") > 0 Then newState = IIf(newType = "Иной приход", EI_ST_REVIEW, EI_ST_ACTIVE)
    PlanBegin("MIGRATE_FIX", canon)
    PlanField("EI", canon)
    PlanField("CHANGED", chg)
    If InStr(chg, "B") > 0 Then
        PlanField("OLD_NAME", CStr(d(SC_NAME)))
        PlanField("NAME", Trim(sName))
    End If
    If InStr(chg, "C") > 0 Then
        PlanField("OLD_ART", CStr(d(SC_ART)))
        PlanField("ART", Trim(sArt))
    End If
    If InStr(chg, "D") > 0 Then
        PlanField("OLD_UNIT", CStr(d(SC_UNIT)))
        PlanField("UNIT", Trim(sUnit))
    End If
    If InStr(chg, "G") > 0 Then
        PlanField("OLD_CAT", CStr(d(SC_CAT)))
        PlanField("CAT", Trim(sCat))
    End If
    If InStr(chg, "J") > 0 Then
        PlanField("OLD_STYPE", oldType)
        PlanField("STYPE", newType)
    End If
    PlanField("OLD_BAL0", b0)
    PlanField("BAL0", b2)
    PlanField("BAL_BEFORE", s)
    PlanField("BAL_AFTER", s2)
    If actText <> "" Then
        PlanField("ACT", Trim(sAct))
        PlanField("ACT_DATE", Format(dAct, "YYYY-MM-DD"))
        If Trim(sReason) <> "" Then PlanField("REASON", Trim(sReason))
    End If
    If InStr(chg, "B") > 0 Then PlanSetValue(SH_STOCK, n, SC_NAME, Trim(sName), False)
    If InStr(chg, "C") > 0 Then PlanSetValue(SH_STOCK, n, SC_ART, Trim(sArt), False)
    If InStr(chg, "D") > 0 Then PlanSetValue(SH_STOCK, n, SC_UNIT, Trim(sUnit), False)
    If InStr(chg, "G") > 0 Then PlanSetValue(SH_STOCK, n, SC_CAT, Trim(sCat), False)
    If InStr(chg, "J") > 0 Then
        PlanSetValue(SH_STOCK, n, SC_STYPE, newType, False)
        If newState <> CStr(d(SC_STATE)) Then PlanSetValue(SH_STOCK, n, SC_STATE, newState, False)
    End If
    If oRow >= 1 Then
        PlanSetValue(SH_ART, oRow, AR_KEY, key, False)
        PlanSetValue(SH_ART, oRow, AR_EI, canon, False)
        PlanSetValue(SH_ART, oRow, AR_ART, Trim(sArt), False)
    End If
    If InStr(chg, "Q") > 0 Then
        PlanSetValue(SH_STOCK, n, SC_QTY, s2, False)
        PlanSetValue(SH_LEG, n, LG_QTY0, b2, False)
        PlanSetValue(SH_LEG, n, LG_BAL0, b2, False)
    End If
    PlanSetValue(SH_LEG, n, LG_FIXES, fixes + 1, False)
    If actText <> "" Then PlanSetValue(SH_LEG, n, LG_ACT, actText, False)
    MigrateFixEI = ApplyOperation(0, 0)
    Exit Function
EH:
    MigrateFixEI = "ERR:внутренняя ошибка проверки исправления ЕИ переноса: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' the card of a transferred EI without an order receipt for the window: "name|article|unit|category|kind|balance at the
' transfer|current balance|state|place", or "" when EI sEI is not such an EI
Function MigrateFixInfo(sEI As String) As String
    Dim n As Long, canon As String, msg As String, lg As Variant, d As Variant
    WmsInit()
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then Exit Function
    If Not LegRow(n, lg) Then Exit Function
    If CStr(lg(LG_KIND)) <> LG_MIGRATE Then Exit Function
    d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_LAST, n).getDataArray()(0)
    MigrateFixInfo = d(SC_NAME) & "|" & d(SC_ART) & "|" & d(SC_UNIT) & "|" & d(SC_CAT) & "|" & d(SC_STYPE) & "|" _
        & IIf(VarType(lg(LG_BAL0)) = 5, WmsIssue.QtyText(lg(LG_BAL0)), "") & "|" & IIf(VarType(d(SC_QTY)) = 5, WmsIssue.QtyText(d(SC_QTY)), "") _
        & "|" & d(SC_STATE) & "|" & d(SC_PLACE)
End Function

' ================================================================ «остаток = сумма движений» (BalanceCheck)
' Expected balances of EIs 1 .. NEXT_EI-1 from the service tables and the posted rows (one pass per table, by chunks):
' exp(n) = live order receipt − consumed before the transfer (live transferred receipt) + balance a MIGRATE brought + live
' special receipts − posted issues + live returns + live corrections. src(n): 1 order receipt, 2 transferred receipt, 3
' MIGRATE, 4 special receipt only, 0 none (an EI WMS knows no source of — test registries).
Sub ExpectedBalances(ByRef ex() As Double, ByRef mv() As Double, ByRef src() As Integer, ByRef nextEI As Long)
    Dim sh As Object, last As Long, r0 As Long, r1 As Long, i As Long, d As Variant, n As Long, st As String, q As Double
    Dim live() As Boolean, k As Long, v As String
    nextEI = CLng(SysNum(SK_NEXT_EI))
    ReDim ex(nextEI)
    ReDim mv(nextEI)
    ReDim src(nextEI)
    ReDim live(nextEI)
    ' order receipts (_RCV, row = EI)
    sh = WmsOrders.RcvSheet()
    last = WmsOrders.LastRow(sh)
    If last >= nextEI Then last = nextEI - 1
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = sh.getCellRangeByPosition(0, r0, RV_QTY, r1).getDataArray()
        For i = 0 To r1 - r0
            n = r0 + i
            If CStr(d(i)(RV_EI)) <> "" Then
                src(n) = 1
                If CStr(d(i)(RV_STATE)) = RV_LIVE And VarType(d(i)(RV_QTY)) = 5 Then
                    ex(n) = ex(n) + d(i)(RV_QTY)
                    live(n) = True
                End If
            End If
        Next i
    Next r0
    ' the transfer (_LEG, row = EI)
    sh = LegSheet()
    last = WmsOrders.LastRow(sh)
    If last >= nextEI Then last = nextEI - 1
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = sh.getCellRangeByPosition(0, r0, LG_PRE, r1).getDataArray()
        For i = 0 To r1 - r0
            n = r0 + i
            v = CStr(d(i)(LG_KIND))
            If v = LG_MIGRATE Then
                src(n) = 3
                If VarType(d(i)(LG_BAL0)) = 5 Then ex(n) = ex(n) + d(i)(LG_BAL0)
            ElseIf v = LG_RECEIPT Or v = LG_ADD Then
                src(n) = 2
                If live(n) And VarType(d(i)(LG_PRE)) = 5 Then ex(n) = ex(n) - d(i)(LG_PRE)
            End If
        Next i
    Next r0
    ' special receipts (_SPR, one row per line)
    If gDoc.Sheets.hasByName(SH_SPR) Then
        sh = gDoc.Sheets.getByName(SH_SPR)
        last = WmsOrders.LastRow(sh)
        For r0 = 1 To last Step CHECK_CHUNK_ROWS
            r1 = r0 + CHECK_CHUNK_ROWS - 1
            If r1 > last Then r1 = last
            d = sh.getCellRangeByPosition(0, r0, SR_STATE, r1).getDataArray()
            For i = 0 To r1 - r0
                n = EiNum(CStr(d(i)(SR_EI)), nextEI)
                If n > 0 Then
                    If src(n) = 0 Then src(n) = 4
                    If CStr(d(i)(SR_STATE)) = RV_LIVE And VarType(d(i)(SR_QTY)) = 5 Then
                        ex(n) = ex(n) + d(i)(SR_QTY)
                        mv(n) = mv(n) + d(i)(SR_QTY)
                    End If
                End If
            Next i
        Next r0
    End If
    ' posted issues («Выдачи»: № in A, «Контроль» starts with «Проведено»)
    sh = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(sh)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = sh.getCellRangeByPosition(0, r0, IC_CTL, r1).getDataArray()
        For i = 0 To r1 - r0
            If VarType(d(i)(IC_NO)) = 5 And Left(CStr(d(i)(IC_CTL)), Len(ST_POSTED)) = ST_POSTED Then
                n = EiNum(CStr(d(i)(IC_EI)), nextEI)
                If n > 0 And VarType(d(i)(IC_QTY)) = 5 Then
                    ex(n) = ex(n) - d(i)(IC_QTY)
                    mv(n) = mv(n) - d(i)(IC_QTY)
                End If
            End If
        Next i
    Next r0
    ' live returns (_RET)
    sh = gDoc.Sheets.getByName(SH_RET)
    last = WmsOrders.LastRow(sh)
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = sh.getCellRangeByPosition(0, r0, RT_STATE, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(d(i)(RT_STATE)) = RV_LIVE And VarType(d(i)(RT_QTY)) = 5 Then
                n = EiNum(CStr(d(i)(RT_EI)), nextEI)
                If n > 0 Then
                    ex(n) = ex(n) + d(i)(RT_QTY)
                    mv(n) = mv(n) + d(i)(RT_QTY)
                End If
            End If
        Next i
    Next r0
    ' live corrections (_ADJ: the change of the balance; a move changes nothing)
    If gDoc.Sheets.hasByName(SH_ADJ) Then
        sh = gDoc.Sheets.getByName(SH_ADJ)
        last = WmsOrders.LastRow(sh)
        For r0 = 1 To last Step CHECK_CHUNK_ROWS
            r1 = r0 + CHECK_CHUNK_ROWS - 1
            If r1 > last Then r1 = last
            d = sh.getCellRangeByPosition(0, r0, AJ_NEW, r1).getDataArray()
            For i = 0 To r1 - r0
                If CStr(d(i)(AJ_STATE)) = RV_LIVE And VarType(d(i)(AJ_QTY)) = 5 Then
                    n = EiNum(CStr(d(i)(AJ_EI)), nextEI)
                    If n > 0 Then
                        ex(n) = ex(n) + d(i)(AJ_QTY)
                        mv(n) = mv(n) + d(i)(AJ_QTY)
                    End If
                End If
                ' 1.0 (Ф8): a kit EI has its source in its own build row (its AJ_EI is its AJ_NEW)
                If CStr(d(i)(AJ_KIND)) = "KIT" And CStr(d(i)(AJ_EI)) = CStr(d(i)(AJ_NEW)) Then
                    n = EiNum(CStr(d(i)(AJ_NEW)), nextEI)
                    If n > 0 Then
                        If src(n) = 0 Then src(n) = 5
                    End If
                End If
                ' 1.0 (М1): the new EI of a split starts with what was taken from the source
                If CStr(d(i)(AJ_KIND)) = "SPLIT" Then
                    n = EiNum(CStr(d(i)(AJ_NEW)), nextEI)
                    If n > 0 Then
                        If src(n) = 0 Then src(n) = 5
                        If CStr(d(i)(AJ_STATE)) = RV_LIVE And VarType(d(i)(AJ_QTY)) = 5 Then
                            ex(n) = ex(n) - d(i)(AJ_QTY)
                            mv(n) = mv(n) - d(i)(AJ_QTY)
                        End If
                    End If
                End If
            Next i
        Next r0
    End If
End Sub

' the number of EI text s («ЕИ-00000017»), or 0 when it is not an EI below NEXT_EI
Function EiNum(s As String, nextEI As Long) As Long
    Dim x As Long
    If Len(s) <> 11 Or Left(s, 3) <> EI_PREFIX Then Exit Function
    If Not IsDigits(Mid(s, 4)) Then Exit Function
    x = CLng(Mid(s, 4))
    If x >= 1 And x < nextEI Then EiNum = x
End Function

' The rule «остаток = сумма движений» for every EI that has a source. "" (and the counters gLg*) when every balance agrees,
' otherwise «ЕИ: ожидается X, в «Наличие» Y» of the first disagreements. Read-only.
Function BalanceCheck() As String
    Dim ex() As Double, mv() As Double, src() As Integer, nextEI As Long, st As Variant, sh As Object, last As Long, r0 As Long, r1 As Long
    Dim n As Long, q As Variant, s As String, i As Long
    WmsInit()
    ExpectedBalances(ex, mv, src, nextEI)
    gLgEIs = 0
    gLgLegacy = 0
    gLgBad = 0
    sh = WmsIssue.StockSheet()
    last = nextEI - 1
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        st = sh.getCellRangeByPosition(0, r0, SC_QTY, r1).getDataArray()
        For i = 0 To r1 - r0
            n = r0 + i
            If src(n) > 0 And CStr(st(i)(SC_EI)) <> "" Then
                gLgEIs = gLgEIs + 1
                If src(n) = 2 Or src(n) = 3 Then gLgLegacy = gLgLegacy + 1
                q = st(i)(SC_QTY)
                If VarType(q) <> 5 Then
                    gLgBad = gLgBad + 1
                    If gLgBad <= 5 Then s = s & IIf(s <> "", "; ", "") & st(i)(SC_EI) & ": остаток не число"
                ElseIf Abs(q - ex(n)) > 0.0000005 Then
                    gLgBad = gLgBad + 1
                    If gLgBad <= 5 Then s = s & IIf(s <> "", "; ", "") & st(i)(SC_EI) & ": по движениям " & WmsIssue.QtyText(WmsIssue.Round3(ex(n))) _
                        & ", в «" & SH_STOCK & "» " & WmsIssue.QtyText(q)
                End If
            End If
        Next i
    Next r0
    BalanceCheck = s
End Function

' 1.0: the rule «остаток = сумма движений» for the snapshot of the tools (WmsExport.ExportSnapshot): manifest lines
' balance_eis, balance_legacy (transferred EIs — their start is in the journal and _LEG, which the tools do not read),
' balance_bad and balance_first (the first disagreement, without «;»)
Function BalanceManifest() As String
    Dim s As String, f As String
    On Error GoTo EH
    s = BalanceCheck()
    f = s
    If InStr(f, ";") > 0 Then f = Left(f, InStr(f, ";") - 1)
    f = Replace(Replace(f, Chr(10), " "), ";", ",")
    BalanceManifest = "balance_eis;" & gLgEIs & Chr(10) & "balance_legacy;" & gLgLegacy & Chr(10) & "balance_bad;" & gLgBad & Chr(10) _
        & "balance_first;" & f & Chr(10)
    Exit Function
EH:
    BalanceManifest = ""
End Function

' 1.0: the lines of «Самопроверка» (WmsDiagnostics.SelfCheck) for the rule «остаток = сумма движений» of every EI and for
' the packing table _PACK — «WARN …» / «OK …», one per line. A disagreement is a warning: it is named, the details are on
' «Сверка переноса»; nothing is changed
Function SelfCheckLines() As String
    Dim s As String, r As String, t0 As Long
    On Error GoTo EH
    t0 = GetSystemTicks()
    s = BalanceCheck()
    If s <> "" Then
        r = "WARN остаток = движения: не сходится у ЕИ: " & gLgBad & " (" & s & IIf(gLgBad > 5, "; …", "") & ") — подробности на листе «" & SH_LEGREP _
            & "», кнопка «Обновить»" & Chr(10)
    Else
        r = "OK остаток = движения: сходится у всех ЕИ (" & gLgEIs & "), " & (GetSystemTicks() - t0) & " мс" & Chr(10)
    End If
    s = WmsPack.PackProblems()
    If s <> "" Then
        r = r & "WARN фасовка: " & s & Chr(10)
    Else
        r = r & "OK фасовка: строки " & SH_PACK & " — на своих ЕИ" & Chr(10)
    End If
    SelfCheckLines = r
    Exit Function
EH:
    SelfCheckLines = "WARN остаток = движения: проверка не выполнена: " & Error$ & Chr(10)
End Function

' ================================================================ the sheet «Сверка переноса»
' One row per transferred EI (and per other EI whose balance disagrees with its movements): the numbers of the transfer,
' where the balance came from, what moved after it, the expected and the actual balance, and what to check. Rows that need
' a look come first. The sheet is read-only and rebuilt by its button; nothing of the accounting is changed.

Private Function SrcText(s As String) As String
    Select Case s
    Case "X"
        SrcText = "колонка X «Наличие»"
    Case "1B"
        SrcText = "лист «Наличие» (1B)"
    Case "1C/1D"
        SrcText = "по выдачам и возвратам (1C/1D)"
    Case "RULE0"
        SrcText = "по правилу: 0"
    Case "RULEF"
        SrcText = "по правилу: всё полученное"
    Case "MIGRATE"
        SrcText = "перенос ЕИ без заказа"
    Case Else
        SrcText = s
    End Select
End Function

' rebuilds «Сверка переноса»; returns the summary («OK:…» / «ERR:…»)
Function LegacyReport() As String
    Dim ex() As Double, mv() As Double, src() As Integer, nextEI As Long, sh As Object, st As Variant, lg As Variant, rows() As Variant, k As Long
    Dim n As Long, last As Long, r0 As Long, r1 As Long, i As Long, q As Variant, diff As Double, note As String, prio As Integer
    Dim out() As Variant, nOut As Long, p As Integer, j As Long, rep As Object, wasProt As Boolean, t0 As Long, x As Variant, b0 As Variant
    Dim nBad As Long, nFlag As Long, nLeg As Long, row As Variant, kind As String, moved As Double, wasModified As Boolean
    WmsInit()
    If gState = "" Then WmsStartup()
    On Error GoTo EH
    t0 = GetSystemTicks()
    ExpectedBalances(ex, mv, src, nextEI)
    sh = WmsIssue.StockSheet()
    last = nextEI - 1
    ReDim rows(IIf(last > 0, last, 0))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        st = sh.getCellRangeByPosition(0, r0, SC_LAST, r1).getDataArray()
        lg = LegSheet().getCellRangeByPosition(0, r0, LG_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            n = r0 + i
            If CStr(st(i)(SC_EI)) = "" Then GoTo NextEI
            q = st(i)(SC_QTY)
            kind = CStr(lg(i)(LG_KIND))
            If src(n) = 0 And kind = "" Then GoTo NextEI
            diff = 0
            If VarType(q) = 5 Then diff = WmsIssue.Round3(q - ex(n))
            If kind = "" And Abs(diff) < 0.0000005 Then GoTo NextEI
            note = ""
            prio = 3
            If VarType(q) <> 5 Then
                note = "остаток в «" & SH_STOCK & "» не число"
                prio = 1
            ElseIf Abs(diff) > 0.0000005 Then
                note = "остаток не равен сумме движений (разница " & WmsIssue.QtyText(diff) & ") — пересчитайте ЕИ на месте"
                prio = 1
            End If
            b0 = lg(i)(LG_BAL0)
            x = lg(i)(LG_X)
            If kind <> "" And VarType(lg(i)(LG_FIXES)) = 5 And CStr(lg(i)(LG_ACT)) <> "" Then
                ' corrected by an act: the numbers of the transfer were checked by a person
                If lg(i)(LG_FIXES) > 0 Then
                    note = note & IIf(note <> "", "; ", "") & "исправлено: " & lg(i)(LG_ACT)
                    nLeg = nLeg + 1
                    GoTo RowDone
                End If
            End If
            If kind <> "" Then
                If CStr(lg(i)(LG_SRC)) = "RULE0" Or CStr(lg(i)(LG_SRC)) = "RULEF" Then
                    note = note & IIf(note <> "", "; ", "") & "остаток при переносе взят " & SrcText(CStr(lg(i)(LG_SRC))) & " — проверьте по полке"
                    If prio > 2 Then prio = 2
                End If
                If CStr(x) <> "" And VarType(b0) = 5 Then
                    If IsNumeric(Replace(CStr(x), ",", ".")) Then
                        If Abs(Val(Replace(CStr(x), ",", ".")) - b0) > 0.0000005 Then
                            note = note & IIf(note <> "", "; ", "") & "в старой таблице X = «" & x & "», перенесено " & WmsIssue.QtyText(b0)
                            If prio > 2 Then prio = 2
                        End If
                    ElseIf CStr(lg(i)(LG_SRC)) <> "X" Then
                        note = note & IIf(note <> "", "; ", "") & "в старой таблице X = «" & x & "» (не число)"
                        If prio > 2 Then prio = 2
                    End If
                End If
                nLeg = nLeg + 1
            End If
RowDone:
            If prio = 1 Then nBad = nBad + 1
            If prio = 2 Then nFlag = nFlag + 1
            moved = WmsIssue.Round3(mv(n))
            row = Array(prio, CStr(st(i)(SC_EI)), CStr(st(i)(SC_NAME)), CStr(st(i)(SC_UNIT)), _
                IIf(kind = LG_MIGRATE, "перенос без заказа", IIf(kind <> "", "перенос: приход по заказу", "приход WMS")), _
                IIf(kind = LG_MIGRATE, "", lg(i)(LG_QTY0)), IIf(kind <> "", b0, ""), IIf(kind <> "" And kind <> LG_MIGRATE, lg(i)(LG_PRE), ""), _
                CStr(x), SrcText(CStr(lg(i)(LG_SRC))), IIf(kind <> "", moved, ""), WmsIssue.Round3(ex(n)), q, _
                IIf(VarType(q) = 5, diff, ""), lg(i)(LG_FIXES), CStr(lg(i)(LG_ACT)), CStr(st(i)(SC_PLACE)), IIf(note = "", "в порядке", note))
            rows(k) = row
            k = k + 1
NextEI:
        Next i
    Next r0
    ' rows that need a look first: priority 1, then 2, then the rest, each in the order of the EIs
    nOut = k
    ReDim out(IIf(nOut > 0, nOut - 1, 0))
    j = 0
    For p = 1 To 3
        For i = 0 To k - 1
            If rows(i)(0) = p Then
                row = rows(i)
                out(j) = Array(row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), _
                    row(14), row(15), row(16), row(17))
                j = j + 1
            End If
        Next i
    Next p
    rep = gDoc.Sheets.getByName(SH_LEGREP)
    wasModified = gDoc.isModified()
    wasProt = rep.isProtected()
    If wasProt Then rep.unprotect(PROTECT_PWD)
    rep.getCellRangeByPosition(0, LEGREP_HEAD + 1, LEGREP_COLS - 1, MAX_SHEET_ROW).clearContents(1023)
    rep.getCellByPosition(0, 1).setString("Обновлено " & Format(Now(), "DD.MM.YYYY HH:MM") & ": перенесённых ЕИ " & nLeg & "; остаток не равен " _
        & "сумме движений — " & nBad & "; проверить по полке — " & nFlag & IIf(nBad + nFlag = 0, " — всё сходится", "") _
        & ". Исправление: «Заказы» → «Исправить» → «Коррекция переноса»; ЕИ без заказа — «" & SH_STOCK & "» → «Исправить ЕИ переноса»; " _
        & "пересчёт по полке — «Корректировки» → «Инвентаризация»")
    If nOut > 0 Then rep.getCellRangeByPosition(0, LEGREP_HEAD + 1, LEGREP_COLS - 1, LEGREP_HEAD + nOut).setDataArray(out)
    If wasProt Then rep.protect(PROTECT_PWD)
    If Not wasModified Then gDoc.setModified(False)
    gLgBad = nBad
    gLgFlags = nFlag
    gLgLegacy = nLeg
    LegacyReport = "OK:перенесённых ЕИ " & nLeg & ", расхождений с движениями " & nBad & ", проверить по полке " & nFlag & ", строк " & nOut _
        & ", " & (GetSystemTicks() - t0) & " мс"
    Exit Function
EH:
    On Error Resume Next
    If wasProt Then rep.protect(PROTECT_PWD)
    LegacyReport = "ERR:сверка не построена: " & Error$ & " (строка " & Erl & ")"
End Function
