' WmsSpecial — специальные приходы: Офис, Производство, Детали, Старый склад, Иной (задание Core Phase 5 PRIME; MASTER SPEC
' v0.3 §4, §5, §7, §8, §14–§18, §25; v0.1 §10–§12, §23; D-059, D-069).
'
' Один общий механизм поступления для всех источников, не пять систем. Общий уровень: номер строки и номер события
' (OFF/PROD/DET/OLD/OTH), новый ЕИ или пополнение существующего ЕИ, строка «Наличие» (остаток, место, единица, категория,
' состояние, источник и тип источника), служебная запись строки _SPR, индекс артикулов деталей _ART, защита строки,
' журнал и восстановление (всё — одной операцией WmsCore.ApplyOperation), исправление, сторно с зависимостями выдач и
' возвратов, антидубль, статусы. Источник задаёт только правила (SourceRule): обязательные поля, новый ЕИ или ЕИ по
' артикулу (детали), состояние нового ЕИ («Требует разбора» у иного прихода).
'
' Служебные данные (лист «Иной приход» — только пользовательская таблица A:S):
'  - _SPR — плотно, строка = № строки: событие, тип, ЕИ, режим (NEW — строка создала ЕИ, ADD — пополнение ЕИ детали),
'    количество, состояние LIVE / STORNO, подсказка строки листа, дата, антидубль, артикул детали, «разобрано как»;
'  - _ART — индекс «артикул детали → её ЕИ» (строка на артикул; ключ — нормализованный артикул); поиск — формулой MATCH
'    движка Calc, найденная строка проверяется по реестру «Наличие» (проверяемый индекс, без прохода по листам);
'  - счётчики в _SYS (схема WMS-SYS-2): NEXT_SPL (№ строки), NEXT_OFF, NEXT_PROD, NEXT_DET, NEXT_OLD, NEXT_OTH.
' Номер строки — только подсказка: перед использованием строка проверяется ключом (№ в A), при несовпадении ищется одной
' формулой MATCH (_IDX), и подсказку исправляет та же операция.
Option Explicit

' details of the last SpecialRowKind: line № of the row, the row hint was outdated (rows inserted above), why FOREIGN
Global gXn As Long
Global gXmoved As Boolean
Global gXwhy As String
' «Провести» of a block as one acceptance: the event № each source got in the block so far (reset by the first block)
Global gXev(4) As String

Private mInHandler As Boolean

' the row being checked (CheckRow)
Private mType As Integer
Private mCode As String
Private mQty As Double
Private mHasQty As Boolean
Private mDate As Double
Private mHasDate As Boolean
Private mName As String
Private mNameIn As String
Private mArt As String
Private mArtKey As String
Private mUnit As String
Private mPlace As String
Private mCat As String
Private mWho As String
Private mDoc As String
Private mMark As String
Private mEvent As String
Private mMode As String
Private mEi As Long
Private mCanon As String
Private mSdata As Variant
Private mBal As Double
Private mPlaceFrom As String
Private mPlaceMove As Boolean
Private mStateBack As Boolean
Private mWarn As String
Private mMissing As String
Private mSysProblem As Boolean

' ================================================================ sheets

Function SpecialSheet() As Object
    SpecialSheet = gDoc.Sheets.getByName(SH_SPECIAL)
End Function

Function SprSheet() As Object
    SprSheet = gDoc.Sheets.getByName(SH_SPR)
End Function

Function ArtSheet() As Object
    ArtSheet = gDoc.Sheets.getByName(SH_ART)
End Function

' "" when the sheets of the special receipts exist and have the expected layout (checked at every start, fail-closed)
Function SpecialSheetsProblem() As String
    Dim names As Variant, i As Integer, sh As Object
    names = Array(SH_SPECIAL, SH_SPR, SH_ART)
    For i = 0 To UBound(names)
        If Not gDoc.Sheets.hasByName(names(i)) Then
            SpecialSheetsProblem = "нет листа «" & names(i) & "»"
            Exit Function
        End If
    Next i
    sh = SpecialSheet()
    If sh.getCellByPosition(XC_NO, 0).getString() <> "№ строки" Or sh.getCellByPosition(XC_TYPE, 0).getString() <> "Тип прихода" _
        Or sh.getCellByPosition(XC_EI, 0).getString() <> "Внутренний код" Or sh.getCellByPosition(XC_CTL, 0).getString() <> "Контроль" _
        Or sh.getCellByPosition(XC_NOTE, 0).getString() <> "Комментарий" Then
        SpecialSheetsProblem = "лист «" & SH_SPECIAL & "»: заголовок не совпадает с A:S"
        Exit Function
    End If
    If SprSheet().getCellByPosition(SR_NO, 0).getString() <> "№ строки" Or SprSheet().getCellByPosition(SR_FIXES, 0).getString() <> "Исправлений" Then
        SpecialSheetsProblem = "служебный лист " & SH_SPR & " повреждён (заголовок)"
        Exit Function
    End If
    If ArtSheet().getCellByPosition(AR_KEY, 0).getString() <> "Артикул (норм.)" Or ArtSheet().getCellByPosition(AR_EI, 0).getString() <> "ЕИ" Then
        SpecialSheetsProblem = "служебный лист " & SH_ART & " повреждён (заголовок)"
        Exit Function
    End If
    If WmsIssue.StockSheet().getCellByPosition(SC_STYPE, 0).getString() <> "Тип источника" Then
        SpecialSheetsProblem = "лист «" & SH_STOCK & "»: нет колонки «Тип источника» (J)"
        Exit Function
    End If
    SpecialSheetsProblem = ""
End Function

' ================================================================ sources (the only thing that differs between them)

Function SrcCode(i As Integer) As String
    SrcCode = Split(SRC_CODES, "|")(i)
End Function

Function SrcName(i As Integer) As String
    SrcName = Split(SRC_NAMES, "|")(i)
End Function

' «Тип источника» of the EI in «Наличие»
Function SrcStype(i As Integer) As String
    SrcStype = Split(SRC_STYPES, "|")(i)
End Function

Function SrcIndexOfCode(code As String) As Integer
    Dim a As Variant, i As Integer
    SrcIndexOfCode = -1
    a = Split(SRC_CODES, "|")
    For i = 0 To UBound(a)
        If a(i) = code Then
            SrcIndexOfCode = i
            Exit Function
        End If
    Next i
End Function

' B «Тип прихода» as typed (any case, extra spaces) → the source; -1 when unknown
Function SrcOfText(ByVal s As String) As Integer
    s = LCase(Trim(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Select Case s
    Case "офис"
        SrcOfText = 0
    Case "производство"
        SrcOfText = 1
    Case "детали", "деталь"
        SrcOfText = 2
    Case "старый склад"
        SrcOfText = 3
    Case "иной", "иной приход"
        SrcOfText = 4
    Case Else
        SrcOfText = -1
    End Select
End Function

' the rules of a source: the parts need their article and find their EI by it; the others always create a new EI.
' Every source needs the name (for a known part it comes from its EI), the quantity, the unit, the date and the place.
Private Function ByArticle(i As Integer) As Boolean
    ByArticle = (i = 2)
End Function

' the state of a new EI: an unidentified receipt is stock at once, but «Требует разбора»
Private Function NewEIState(i As Integer) As String
    If i = 4 Then NewEIState = EI_ST_REVIEW Else NewEIState = EI_ST_ACTIVE
End Function

' the counter of the event numbers of source i in _SYS
Private Function EventCounterKey(i As Integer) As Integer
    EventCounterKey = SK_NEXT_OFF + i
End Function

' ================================================================ event IDs (OFF-00000001 …): one acceptance, several lines

Function EventId(i As Integer, num As Long) As String
    EventId = SrcCode(i) & "-" & Right(String(EVENT_DIGITS, "0") & num, EVENT_DIGITS)
End Function

' "OFF-00000012", "off-12", "ОFF 12" (a Cyrillic lookalike letter) → source and number; False when not an event ID
Function ParseEvent(ByVal s As String, ByRef i As Integer, ByRef num As Long) As Boolean
    Dim t As String, p As Long, pre As String, dg As String
    i = -1
    num = 0
    t = LatinLook(UCase(Trim(s)))
    For p = 1 To Len(t)
        If IsDigits(Mid(t, p, 1)) Then Exit For
    Next p
    If p > Len(t) Or p < 2 Then Exit Function
    pre = Trim(Left(t, p - 1))
    If Right(pre, 1) = "-" Then pre = Trim(Left(pre, Len(pre) - 1))
    dg = Mid(t, p)
    If Not IsDigits(dg) Or Len(dg) > EVENT_DIGITS Then Exit Function
    i = SrcIndexOfCode(pre)
    If i < 0 Then Exit Function
    num = CLng(dg)
    If num < 1 Then Exit Function
    ParseEvent = True
End Function

' ================================================================ the article of a part (spec §8: one article — one EI)

' Cyrillic letters that look like Latin ones (upper case) become Latin: А В Е Ё К М Н О Р С Т У Х
Private Function LatinLook(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case Asc(ch)
        Case 1040
            ch = "A"
        Case 1042
            ch = "B"
        Case 1045, 1025
            ch = "E"
        Case 1050
            ch = "K"
        Case 1052
            ch = "M"
        Case 1053
            ch = "H"
        Case 1054
            ch = "O"
        Case 1056
            ch = "P"
        Case 1057
            ch = "C"
        Case 1058
            ch = "T"
        Case 1059
            ch = "Y"
        Case 1061
            ch = "X"
        End Select
        out = out & ch
    Next i
    LatinLook = out
End Function

' The key of an article: upper case, lookalike Cyrillic letters as Latin, every kind of dash as "-", no spaces (a space,
' a no-break space, a tab), no soft hyphen. "abc-100", " ABC‑100 ", "АВС-100" (Cyrillic) → "ABC-100". Everything else
' stays as typed: "ABC100" and "ABC-100" are different articles.
Function ArticleKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, code As Long
    s = UCase(s)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = Asc(ch)
        Select Case code
        Case 9, 10, 13, 32, 160, 8194, 8195, 8200, 8201, 8239, 173
            ch = ""
        Case 8208 To 8213, 8722
            ch = "-"
        End Select
        out = out & ch
    Next i
    ArticleKey = LatinLook(out)
End Function

' The EI of the article key: 0 not found, 1 found (n, canon, aRow — the _ART row), 2 several EIs share the key (why:
' fail-closed, the article needs a review), 3 the index does not agree with «Наличие» (why: a problem of WMS). Two MATCH
' formulas of the Calc engine (the first entry and whether there is a second one); the entry is checked against the
' registry row of its EI: a part, with this very article.
Function ArtLookup(key As String, ByRef n As Long, ByRef canon As String, ByRef aRow As Long, ByRef why As String) As Integer
    Dim r1 As Long, r2 As Long, d As Variant, sd As Variant, p As String
    n = 0
    canon = ""
    aRow = -1
    why = ""
    r1 = WmsOrders.FindArtRow(key, 0)
    If r1 < 1 Then
        ArtLookup = 0
        Exit Function
    End If
    d = ArtSheet().getCellRangeByPosition(0, r1, AR_LAST, r1).getDataArray()(0)
    r2 = WmsOrders.FindArtRow(key, r1)
    If r2 >= 1 Then
        why = "Для артикула «" & key & "» найдено несколько ЕИ (" & d(AR_EI) & ", " & ArtSheet().getCellByPosition(AR_EI, r2).getString() _
            & "). Требуется разбор"
        ArtLookup = 2
        Exit Function
    End If
    If CStr(d(AR_KEY)) <> key Then
        why = "индекс артикулов " & SH_ART & ": строка " & (r1 + 1) & " «" & d(AR_KEY) & "» не совпадает с ключом «" & key & "» — нужна самопроверка"
        ArtLookup = 3
        Exit Function
    End If
    If Not WmsOrders.StrictEI(CStr(d(AR_EI)), n) Then
        why = "индекс артикулов " & SH_ART & " повреждён: в строке " & (r1 + 1) & " ЕИ «" & d(AR_EI) & "» — нужна самопроверка"
        ArtLookup = 3
        Exit Function
    End If
    canon = CStr(d(AR_EI))
    p = WmsIssue.StockRowProblem(n, canon)
    If p <> "" Then
        why = "индекс артикулов указывает на " & canon & ": " & p & " — нужна самопроверка"
        ArtLookup = 3
        Exit Function
    End If
    sd = WmsIssue.StockData(n)
    If CStr(sd(SC_STYPE)) <> SrcStype(2) Or ArticleKey(CStr(sd(SC_ART))) <> key Then
        why = "индекс артикулов не согласован с «" & SH_STOCK & "»: у " & canon & " тип «" & sd(SC_STYPE) & "», артикул «" & sd(SC_ART) _
            & "», а в индексе «" & key & "» — нужна самопроверка"
        ArtLookup = 3
        Exit Function
    End If
    aRow = r1
    ArtLookup = 1
End Function

' ================================================================ service rows (dense addressing, the key is checked)

' A..N of _SPR row n (line № n)
Function SprRow(n As Long) As Variant
    SprRow = SprSheet().getCellRangeByPosition(0, n, SR_LAST, n).getDataArray()(0)
End Function

' True when line № n is registered in _SPR (its row holds n)
Function SprRegistered(n As Long, d As Variant) As Boolean
    If n < 1 Or n > MAX_SHEET_ROW Then Exit Function
    d = SprRow(n)
    If VarType(d(SR_NO)) <> 5 Then Exit Function
    SprRegistered = (d(SR_NO) = n)
End Function

' the live lines of EI canon (the lines of _SPR that still count): one formula of the Calc engine over the used rows of
' _SPR (COUNTIFS; no loop). False when the formula could not be used.
Function LiveLinesOfEI(canon As String, ByRef nLive As Long) As Boolean
    Dim lr As Long, f As String, s As String, ok As Boolean
    nLive = 0
    lr = WmsOrders.LastRow(SprSheet())
    If lr < 1 Then
        LiveLinesOfEI = True
        Exit Function
    End If
    f = "COUNTIFS($'" & SH_SPR & "'.$D$2:$D$" & (lr + 1) & ";""" & Replace(canon, """", """""") & """;$'" & SH_SPR & "'.$G$2:$G$" & (lr + 1) _
        & ";""" & RV_LIVE & """)"
    s = WmsOrders.ScratchEval(f, ok)
    If Not ok Or Not IsDigits(s) Then Exit Function
    nLive = CLng(s)
    LiveLinesOfEI = True
End Function

' ================================================================ row model of «Иной приход»

' Kind of row r: "EMPTY", "OPEN" (not posted), "POSTED", "DELETED", "COPY" (marked), "FOREIGN" (a № that WMS did not
' register at this row: never issued, not a line of this WMS, or a copy of a registered row). gXn, gXmoved, gXwhy.
Function SpecialRowKind(r As Long) As String
    Dim sh As Object, d As Variant, x As Double, n As Long, rd As Variant, hint As Long, c As Object, i As Integer
    sh = SpecialSheet()
    gXn = 0
    gXmoved = False
    gXwhy = ""
    d = sh.getCellRangeByPosition(0, r, XC_LAST, r).getDataArray()(0)
    If Left(CStr(d(XC_CTL)), 5) = "КОПИЯ" Then
        SpecialRowKind = "COPY"
        Exit Function
    End If
    If CStr(d(XC_NO)) = "" Then
        SpecialRowKind = "EMPTY"
        For i = XC_TYPE To XC_MARK
            If CStr(d(i)) <> "" Then
                SpecialRowKind = "OPEN"
                Exit For
            End If
        Next i
        Exit Function
    End If
    If VarType(d(XC_NO)) <> 5 Then
        gXwhy = "№ «" & d(XC_NO) & "» не выдавался WMS"
        SpecialRowKind = "FOREIGN"
        Exit Function
    End If
    x = d(XC_NO)
    If x < 1 Or x <> Int(x) Or x > MAX_SHEET_ROW Then
        gXwhy = "неверный № «" & d(XC_NO) & "»"
        SpecialRowKind = "FOREIGN"
        Exit Function
    End If
    n = CLng(x)
    If n >= SysNum(SK_NEXT_SPL) Then
        gXwhy = "№ строки " & n & " WMS не выдавала"
        SpecialRowKind = "FOREIGN"
        Exit Function
    End If
    If Not SprRegistered(n, rd) Then
        gXwhy = "строка № " & n & " не зарегистрирована WMS"
        SpecialRowKind = "FOREIGN"
        Exit Function
    End If
    If VarType(rd(SR_ROW)) = 5 Then hint = CLng(rd(SR_ROW)) Else hint = -1
    If hint <> r Then
        ' the registered row is elsewhere: when it still holds the key this row is a copy, otherwise rows moved
        If hint >= 1 And hint <= MAX_SHEET_ROW Then
            c = sh.getCellByPosition(XC_NO, hint)
            If c.getType() = com.sun.star.table.CellContentType.VALUE Then
                If c.getValue() = n Then
                    gXwhy = "строка № " & n & " повторяет строку " & (hint + 1)
                    SpecialRowKind = "FOREIGN"
                    Exit Function
                End If
            End If
        End If
        If WmsOrders.CountSpecialNo(n) > 1 Then
            If FirstSpecialRow(n) <> r Then
                gXwhy = "строка № " & n & " повторяется"
                SpecialRowKind = "FOREIGN"
                Exit Function
            End If
        End If
        gXmoved = True
    End If
    gXn = n
    If CStr(rd(SR_STATE)) = RV_STORNO Then SpecialRowKind = "DELETED" Else SpecialRowKind = "POSTED"
End Function

' the first row holding line № n in A that is not marked КОПИЯ; -1 when none
Function FirstSpecialRow(n As Long) As Long
    Dim r As Long, start As Long, i As Integer, sh As Object
    sh = SpecialSheet()
    FirstSpecialRow = -1
    For i = 1 To 20
        r = WmsOrders.FindSpecialNoRow(n, start)
        If r < 1 Then Exit Function
        If Left(sh.getCellByPosition(XC_CTL, r).getString(), 5) <> "КОПИЯ" Then
            FirstSpecialRow = r
            Exit Function
        End If
        start = r
    Next i
End Function

' 0-based «Иной приход» row of line № n: the hint of _SPR checked by the key, otherwise found by MATCH; -1 when none
Function SpecialRowOf(n As Long, hint As Variant) As Long
    Dim sh As Object, c As Object, h As Long
    sh = SpecialSheet()
    If VarType(hint) = 5 Then h = CLng(hint) Else h = -1
    If h >= 1 And h <= MAX_SHEET_ROW Then
        c = sh.getCellByPosition(XC_NO, h)
        If c.getType() = com.sun.star.table.CellContentType.VALUE Then
            If c.getValue() = n Then
                SpecialRowOf = h
                Exit Function
            End If
        End If
    End If
    SpecialRowOf = FirstSpecialRow(n)
End Function

' a № in A that WMS did not register at this row marks the row КОПИЯ (spec §15)
Sub CopyCheckSpecial(r As Long)
    If SpecialRowKind(r) <> "FOREIGN" Then Exit Sub
    SpecialSheet().getCellByPosition(XC_CTL, r).setString("КОПИЯ: " & gXwhy & " — не проводится, очистите строку кнопкой «Очистить»")
End Sub

' writes the text of «Контроль» (R) of an unposted row
Private Sub SetCtl(r As Long, s As String)
    WmsOrders.SetIfDiff(SpecialSheet().getCellByPosition(XC_CTL, r), s)
End Sub

' ================================================================ checks of a row before posting (shared with the preview)

' lower case letters and digits, every other run of characters one "_" (names and documents in the antidubl key)
Private Function NormText(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, gap As Boolean, code As Long
    s = LCase(Trim(s))
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = Asc(ch)
        If code = 1105 Then
            ch = Chr(1077)
            code = 1077
        End If
        If (code >= 48 And code <= 57) Or (code >= 97 And code <= 122) Or (code >= 1072 And code <= 1103) Then
            If gap And out <> "" Then out = out & "_"
            out = out & ch
            gap = False
        Else
            gap = True
        End If
    Next i
    NormText = out
End Function

Private Sub AddWarn(s As String)
    If mWarn = "" Then mWarn = s Else mWarn = mWarn & "; " & s
End Sub

Private Sub AddMissing(s As String)
    If mMissing = "" Then mMissing = s Else mMissing = mMissing & ", " & s
End Sub

' Validates row r: "" or the reason; fills the m* values. bNeedAll: posting needs every required field (the preview
' checks only what is filled and collects what is missing in mMissing). sJoin: the event of the block the row is posted
' with ("" — the row's own C or a new event). mSysProblem: the reason is a problem of WMS (the index of the parts does
' not agree with the registry), not of the row.
Private Function CheckRow(r As Long, bNeedAll As Boolean, sJoin As String) As String
    Dim sh As Object, d As Variant, t As String, ti As Integer, num As Long, c As Object, msg As String, st As Integer, why As String
    Dim aRow As Long, cu As String, cn As String, cp As String, cc As String
    sh = SpecialSheet()
    mSysProblem = False
    mWarn = ""
    mMissing = ""
    mMode = SR_NEW
    mEi = 0
    mCanon = ""
    mBal = 0
    mPlaceMove = False
    mStateBack = False
    mArtKey = ""
    mHasQty = False
    mHasDate = False
    d = sh.getCellRangeByPosition(0, r, XC_LAST, r).getDataArray()(0)
    ' B the source
    t = Trim(CStr(d(XC_TYPE)))
    If t = "" Then
        CheckRow = "не указан тип прихода (B): Офис, Производство, Детали, Старый склад или Иной"
        Exit Function
    End If
    mType = SrcOfText(t)
    If mType < 0 Then
        CheckRow = "тип прихода «" & t & "» неизвестен — выберите Офис, Производство, Детали, Старый склад или Иной"
        Exit Function
    End If
    mCode = SrcCode(mType)
    ' C the event: an issued event of the same source joins this line to it; empty — the event of the block or a new one
    t = Trim(CStr(d(XC_EVENT)))
    mEvent = ""
    If t <> "" Then
        If Not ParseEvent(t, ti, num) Then
            CheckRow = "№ поступления (C) «" & t & "» — не номер поступления (пример: " & EventId(mType, 1) & "); оставьте C пустым — WMS выдаст новый №"
            Exit Function
        End If
        If ti <> mType Then
            CheckRow = "№ поступления (C) " & EventId(ti, num) & " относится к типу «" & SrcName(ti) & "», а в строке тип «" & SrcName(mType) & "»"
            Exit Function
        End If
        If num >= SysNum(EventCounterKey(ti)) Then
            CheckRow = "поступление " & EventId(ti, num) & " ещё не проводилось — оставьте № поступления (C) пустым, WMS выдаст новый №"
            Exit Function
        End If
        mEvent = EventId(ti, num)
    ElseIf sJoin <> "" Then
        mEvent = sJoin
    End If
    ' F the quantity, H the date
    c = sh.getCellByPosition(XC_QTY, r)
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then
        If bNeedAll Then
            CheckRow = "не указано количество (F)"
            Exit Function
        End If
        AddMissing("количество (F)")
    Else
        If Not WmsOrders.ParseQtyAs(c, "Количество (F)", mQty, msg) Then
            CheckRow = msg
            Exit Function
        End If
        mHasQty = True
    End If
    c = sh.getCellByPosition(XC_DATE, r)
    If c.getType() = com.sun.star.table.CellContentType.EMPTY Then
        If bNeedAll Then
            CheckRow = "не указана дата поступления (H)"
            Exit Function
        End If
        AddMissing("дату поступления (H)")
    Else
        If Not WmsOrders.ParseDateCellAs(c, "дата поступления (H)", mDate, msg) Then
            CheckRow = msg
            Exit Function
        End If
        mHasDate = True
    End If
    mNameIn = Trim(CStr(d(XC_NAME)))
    mArt = Trim(CStr(d(XC_ART)))
    mUnit = Trim(CStr(d(XC_UNIT)))
    mPlace = Trim(CStr(d(XC_PLACE)))
    mCat = Trim(CStr(d(XC_CAT)))
    mWho = Trim(CStr(d(XC_WHO)))
    mDoc = Trim(CStr(d(XC_DOC)))
    mMark = Trim(CStr(d(XC_MARK)))
    If ByArticle(mType) Then
        ' parts: the article is the key — the same article adds to the same EI (spec §8)
        If mArt = "" Then
            CheckRow = IIf(bNeedAll, "для типа «Детали» артикул (E) обязателен — по нему WMS находит ЕИ детали", _
                "укажите артикул (E) — для детали он обязателен: по нему WMS находит её ЕИ")
            Exit Function
        End If
        mArtKey = ArticleKey(mArt)
        If mArtKey = "" Then
            CheckRow = "артикул (E) «" & mArt & "» пуст"
            Exit Function
        End If
        st = ArtLookup(mArtKey, mEi, mCanon, aRow, why)
        If st = 2 Then
            CheckRow = why
            Exit Function
        End If
        If st = 3 Then
            mSysProblem = True
            CheckRow = why
            Exit Function
        End If
        If st = 1 Then
            mMode = SR_ADD
            mSdata = WmsIssue.StockData(mEi)
            If Not WmsIssue.StockQty(mEi, mBal) Then
                mSysProblem = True
                CheckRow = "остаток " & mCanon & " в «" & SH_STOCK & "» не число — реестр повреждён"
                Exit Function
            End If
            cu = Trim(CStr(mSdata(SC_UNIT)))
            If mUnit <> "" And LCase(mUnit) <> LCase(cu) Then
                CheckRow = "Для " & mCanon & " используется единица «" & cu & "». Пересчёт единиц не настроен"
                Exit Function
            End If
            mUnit = cu
            cn = CStr(mSdata(SC_NAME))
            mName = cn
            If mNameIn <> "" And NormText(mNameIn) <> NormText(cn) Then
                AddWarn("Артикул уже существует: " & mCanon & ", название «" & cn & "». Будет использовано существующее наименование")
            End If
            cp = CStr(mSdata(SC_PLACE))
            mPlaceFrom = cp
            If mPlace = "" Then
                mPlace = cp
            ElseIf mPlace <> cp Then
                mPlaceMove = True
                AddWarn("место " & mCanon & " изменится: «" & cp & "» → «" & mPlace & "»")
            End If
            cc = CStr(mSdata(SC_CAT))
            If mCat <> "" And LCase(mCat) <> LCase(cc) Then
                AddWarn("категория карточки " & mCanon & " «" & cc & "» сохраняется (в строке «" & mCat & "»)")
            End If
            mCat = cc
            mStateBack = (CStr(mSdata(SC_STATE)) = EI_ST_STORNO)
            CheckRow = ""
            Exit Function
        End If
    End If
    ' a new EI: the name, the unit and the place are needed
    mName = mNameIn
    If mName = "" Then
        If bNeedAll Then
            CheckRow = IIf(mType = 4, "не указано наименование или временное описание (D)", "не указано наименование (D)")
            Exit Function
        End If
        AddMissing("наименование (D)")
    End If
    If mUnit = "" Then
        If bNeedAll Then
            CheckRow = "не указана единица измерения (G)"
            Exit Function
        End If
        AddMissing("единицу (G)")
    End If
    If mPlace = "" Then
        If bNeedAll Then
            CheckRow = "не указано место хранения (I)"
            Exit Function
        End If
        AddMissing("место хранения (I)")
    End If
    CheckRow = ""
End Function

' what the row will do (after CheckRow): «новый ЕИ (Офис)», «пополнение ЕИ-… «…»: остаток 100 → 150 шт»
Private Function Describe() As String
    Dim s As String
    If mMode = SR_ADD Then
        s = "пополнение " & mCanon & " «" & mName & "»: остаток " & WmsIssue.QtyText(mBal)
        If mHasQty Then s = s & " → " & WmsIssue.QtyText(WmsIssue.Round3(mBal + mQty))
        s = s & " " & mUnit
        If mStateBack Then s = s & " (приходы этого ЕИ были удалены — он снова станет активным)"
    ElseIf mType = 4 Then
        s = "новый ЕИ, требует разбора"
    Else
        s = "новый ЕИ (" & SrcName(mType) & ")"
    End If
    If mEvent <> "" Then s = s & ", поступление " & mEvent
    Describe = s
End Function

' Describe() with a capital letter (the text of «Контроль»)
Private Function DescribeCap() As String
    Dim s As String
    s = Describe()
    DescribeCap = UCase(Left(s, 1)) & Mid(s, 2)
End Function

' "" when line № n (= NEXT_SPL) may be handed out now
Private Function NewSplProblem(n As Long) As String
    If n < 1 Or n >= MAX_SHEET_ROW Then
        NewSplProblem = "номера строк иного прихода исчерпаны (NEXT_SPL " & n & ")"
    ElseIf WmsOrders.LastRow(SprSheet()) >= n Then
        NewSplProblem = "служебная таблица " & SH_SPR & " не согласована со счётчиком NEXT_SPL " & n & " — нужна самопроверка"
    ElseIf WmsOrders.CountSpecialNo(n) > 0 Then
        NewSplProblem = "№ строки " & n & " уже записан на листе «" & SH_SPECIAL & "» — новый номер не выдаётся. Нужна проверка ответственным"
    Else
        NewSplProblem = ""
    End If
End Function

' the antidubl key of a line: source + article (a part) or name + quantity + date + document
Private Function DupKeyOf(code As String, artKey As String, sName As String, q As Double, dt As Double, sDoc As String) As String
    Dim p As String
    If artKey <> "" Then p = "a" & artKey Else p = "n" & NormText(sName)
    DupKeyOf = LCase(code) & "|" & p & "|" & Replace(NumStr(q), ".", "d") & "|" & NumStr(Int(dt)) & "|" & NormText(sDoc)
End Function

' the first other live line with the same key: "строка № 7 (PROD-00000003)", or ""
Function FindSpecialDuplicate(keyText As String, exceptN As Long) As String
    Dim h As Double, n As Long, start As Long, i As Integer, d As Variant
    h = WmsOrders.DupHash(keyText)
    start = 0
    For i = 1 To 50
        n = WmsOrders.FindSpecialDupLine(h, start)
        If n < 1 Then Exit Function
        If n <> exceptN Then
            d = SprRow(n)
            If CStr(d(SR_STATE)) = RV_LIVE And CStr(d(SR_DUPKEY)) = keyText Then
                FindSpecialDuplicate = "строка № " & n & " (" & d(SR_EVENT) & ", " & d(SR_EI) & ")"
                Exit Function
            End If
        End If
        start = n
    Next i
End Function

' ================================================================ plan helpers

' a new «Наличие» row of the common level: EI, name, article, unit, balance, place, category, state, source, kind of source
Sub PlanNewStockRow(n As Long, canon As String, qty As Double, sName As String, sArt As String, sUnit As String, sPlace As String, _
    sCat As String, sState As String, sSrc As String, sType As String)
    PlanSetValue(SH_STOCK, n, SC_EI, canon, False)
    PlanSetValue(SH_STOCK, n, SC_NAME, sName, False)
    If sArt <> "" Then PlanSetValue(SH_STOCK, n, SC_ART, sArt, False)
    PlanSetValue(SH_STOCK, n, SC_UNIT, sUnit, False)
    PlanSetValue(SH_STOCK, n, SC_QTY, qty, False)
    PlanSetValue(SH_STOCK, n, SC_PLACE, sPlace, False)
    If sCat <> "" Then PlanSetValue(SH_STOCK, n, SC_CAT, sCat, False)
    PlanSetValue(SH_STOCK, n, SC_STATE, sState, False)
    PlanSetValue(SH_STOCK, n, SC_SRC, sSrc, False)
    PlanSetValue(SH_STOCK, n, SC_STYPE, sType, False)
End Sub

Private Sub PlanSprRow(n As Long, evId As String, canon As String, q As Double, r As Long, dupKey As String)
    PlanSetValue(SH_SPR, n, SR_NO, n, False)
    PlanSetValue(SH_SPR, n, SR_EVENT, evId, False)
    PlanSetValue(SH_SPR, n, SR_TYPE, mCode, False)
    PlanSetValue(SH_SPR, n, SR_EI, canon, False)
    PlanSetValue(SH_SPR, n, SR_MODE, mMode, False)
    PlanSetValue(SH_SPR, n, SR_QTY, q, False)
    PlanSetValue(SH_SPR, n, SR_STATE, RV_LIVE, False)
    PlanSetValue(SH_SPR, n, SR_ROW, r, False)
    PlanSetValue(SH_SPR, n, SR_DATE, mDate, False)
    PlanSetValue(SH_SPR, n, SR_DUP, WmsOrders.DupHash(dupKey), False)
    PlanSetValue(SH_SPR, n, SR_DUPKEY, dupKey, False)
    If mArtKey <> "" Then PlanSetValue(SH_SPR, n, SR_ARTKEY, mArtKey, False)
End Sub

' the values of a posted row: A №, B C D E F G H I J (the checked inputs), K L M kept as typed, N EI, O P balances, Q R
Private Sub PlanRowValues(r As Long, n As Long, evId As String, canon As String, bHasBefore As Boolean, before As Double, after As Double, _
    status As String, ctl As String)
    Dim sh As Object, c As Integer
    sh = SpecialSheet()
    PlanSetValue(SH_SPECIAL, r, XC_NO, n, False)
    PlanSetValue(SH_SPECIAL, r, XC_TYPE, SrcName(mType), True)
    PlanSetValue(SH_SPECIAL, r, XC_EVENT, evId, True)
    PlanSetValue(SH_SPECIAL, r, XC_NAME, mName, True)
    If mArt <> "" Then PlanSetValue(SH_SPECIAL, r, XC_ART, mArt, True)
    PlanSetValue(SH_SPECIAL, r, XC_QTY, mQty, True)
    PlanSetValue(SH_SPECIAL, r, XC_UNIT, mUnit, True)
    PlanSetValue(SH_SPECIAL, r, XC_DATE, mDate, True)
    PlanSetValue(SH_SPECIAL, r, XC_PLACE, mPlace, True)
    If mCat <> "" Then
        PlanSetValue(SH_SPECIAL, r, XC_CAT, mCat, True)
    ElseIf sh.getCellByPosition(XC_CAT, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        PlanSetValue(SH_SPECIAL, r, XC_CAT, "", True)
    End If
    For c = XC_WHO To XC_MARK
        If sh.getCellByPosition(c, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then PlanInput(SH_SPECIAL, r, c)
    Next c
    PlanSetValue(SH_SPECIAL, r, XC_EI, canon, False)
    If bHasBefore Then
        PlanSetValue(SH_SPECIAL, r, XC_BEFORE, before, False)
    ElseIf sh.getCellByPosition(XC_BEFORE, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        PlanSetValue(SH_SPECIAL, r, XC_BEFORE, "", False)
    End If
    PlanSetValue(SH_SPECIAL, r, XC_AFTER, after, False)
    PlanSetValue(SH_SPECIAL, r, XC_STATUS, status, False)
    PlanDerived(SH_SPECIAL, r, XC_CTL, ctl)
End Sub

' Q «Статус» of a line of kind k (XK_*), after a correction or not
Function LineStatus(k As String, bFixed As Boolean) As String
    If bFixed Then LineStatus = XS_FIXED & k Else LineStatus = XS_POSTED & k
End Function

' the kind of a registered line (XK_*): new EI, a part added to, not identified yet, identified
Private Function LineKind(sd As Variant) As String
    If CStr(sd(SR_MODE)) = SR_ADD Then
        LineKind = XK_ADD
    ElseIf CStr(sd(SR_TYPE)) = "OTH" Then
        If CStr(sd(SR_IDENT)) <> "" Then LineKind = XK_IDENT Else LineKind = XK_REVIEW
    Else
        LineKind = XK_NEW
    End If
End Function

' ================================================================ «Провести» — SP_RECEIPT (new EI) / SP_REFILL (a part: its EI)

' The place a posting of row r would move an existing EI to: "ЕИ-…: «A-1» → «B-2»", or "" (no move, or the row does not
' pass its checks). The button asks for the move before posting (spec: a place is never changed silently).
Function SpecialPlaceMove(r As Long) As String
    WmsInit()
    On Error GoTo EH
    If SpecialRowKind(r) <> "OPEN" Then Exit Function
    mInHandler = True
    If CheckRow(r, True, "") = "" Then
        If mMode = SR_ADD And mPlaceMove Then SpecialPlaceMove = mCanon & ": «" & mPlaceFrom & "» → «" & mPlace & "»"
    End If
    mInHandler = False
    Exit Function
EH:
    mInHandler = False
End Function

' Posts row r (0-based) of «Иной приход». sJoin: the event of the block ("" — C of the row or a new event); bPlaceOk: the
' user confirmed that a part's EI moves to the place of the row. OK:<seq>, SKIP:<why>, ERR:<why> (a problem of the row,
' the reason in R), ERR-SYS:<why> (a problem of WMS itself: a counter or the index of the parts is not safe, an internal
' error), BLOCKED:<why> or an ApplyOperation error (ERR-RB, ERR-CRITICAL).
Function SpecialPostRow(r As Long, sJoin As String, bPlaceOk As Boolean) As String
    Dim why As String, kind As String, nL As Long, n As Long, canon As String, res As String, evNum As Long, evId As String
    Dim newEvent As Boolean, dupKey As String, dup As String, bal2 As Double, status As String, ctl As String, aRow As Long
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        SpecialPostRow = "ERR:выберите строку иного прихода (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = SpecialRowKind(r)
    Select Case kind
    Case "POSTED"
        SpecialPostRow = "SKIP:уже проведено (строка № " & gXn & ")"
        Exit Function
    Case "DELETED"
        SpecialPostRow = "SKIP:строка № " & gXn & " удалена (сторно)"
        Exit Function
    Case "COPY"
        SpecialPostRow = "SKIP:строка помечена как КОПИЯ — очистите её кнопкой «Очистить»"
        Exit Function
    Case "FOREIGN"
        mInHandler = True
        CopyCheckSpecial(r)
        mInHandler = False
        SpecialPostRow = "SKIP:в строке № строки, который WMS здесь не проводила — строка не проводится"
        Exit Function
    Case "EMPTY"
        SpecialPostRow = "ERR:строка пуста"
        Exit Function
    End Select
    why = PostingBlockReason()
    If why <> "" Then
        If Left(why, 8) = "BLOCKED:" Then SetCtl(r, "Не проведено: WMS заблокирована — " & Mid(why, 9) & ". См. лист «" & SH_MAIN & "»")
        SpecialPostRow = why
        Exit Function
    End If
    why = CheckRow(r, True, sJoin)
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        SpecialPostRow = IIf(mSysProblem, "ERR-SYS:", "ERR:") & why
        Exit Function
    End If
    If mMode = SR_ADD And mPlaceMove And Not bPlaceOk Then
        why = "место " & mCanon & " «" & mPlaceFrom & "» отличается от указанного «" & mPlace & "» — перемещение ЕИ нужно подтвердить: " _
            & "проведите строку отдельно (WMS спросит) или очистите место (I)"
        SetCtl(r, "Не проведено: " & why)
        SpecialPostRow = "ERR:" & why
        Exit Function
    End If
    ' the numbers must be safe: a failure here is a problem of WMS, not of this row (D-059)
    nL = CLng(SysNum(SK_NEXT_SPL))
    why = NewSplProblem(nL)
    If why = "" And mEvent = "" Then
        newEvent = True
        evNum = CLng(SysNum(EventCounterKey(mType)))
        If evNum < 1 Or evNum > 99999999 Then why = "номера поступлений «" & SrcName(mType) & "» исчерпаны (" & evNum & ")"
        evId = EventId(mType, evNum)
    Else
        evId = mEvent
    End If
    If why = "" And mMode = SR_NEW Then
        n = CLng(SysNum(SK_NEXT_EI))
        why = WmsOrders.NewEIProblem(n)
        canon = WmsIssue.EiCanon(n)
        If why = "" And ByArticle(mType) Then aRow = WmsOrders.LastRow(ArtSheet()) + 1
    Else
        n = mEi
        canon = mCanon
    End If
    If why <> "" Then
        SetCtl(r, "Не проведено: " & why)
        SpecialPostRow = "ERR-SYS:" & why
        Exit Function
    End If
    dupKey = DupKeyOf(mCode, mArtKey, mName, mQty, mDate, mDoc)
    dup = FindSpecialDuplicate(dupKey, 0)
    If mMode = SR_ADD Then
        bal2 = WmsIssue.Round3(mBal + mQty)
        status = LineStatus(XK_ADD, False)
        ctl = "Пополнение " & canon & ": остаток " & WmsIssue.QtyText(mBal) & " → " & WmsIssue.QtyText(bal2) & " " & mUnit & " · " & evId
    Else
        bal2 = mQty
        status = LineStatus(IIf(mType = 4, XK_REVIEW, XK_NEW), False)
        ctl = "Новый ЕИ " & canon & " · " & evId & IIf(mType = 4, " · требует разбора: позднее — «Разобрать»", "")
    End If
    If mWarn <> "" Then ctl = ctl & " · " & mWarn
    If dup <> "" Then ctl = ctl & " · Возможный дубль: " & dup
    PlanBegin(IIf(mMode = SR_ADD, "SP_REFILL", "SP_RECEIPT"), CStr(nL))
    PlanField("SPL", nL)
    PlanField("EVENT", evId)
    PlanField("EVENT_NEW", IIf(newEvent, 1, 0))
    PlanField("SRC", mCode)
    PlanField("EI", canon)
    PlanField("MODE", mMode)
    PlanField("QTY", mQty)
    PlanField("UNIT", mUnit)
    PlanField("NAME", mName)
    If mNameIn <> "" And mNameIn <> mName Then PlanField("NAME_INPUT", mNameIn)
    PlanField("ART", mArt)
    If mArtKey <> "" Then PlanField("ART_KEY", mArtKey)
    PlanField("DATE", Format(mDate, "YYYY-MM-DD"))
    PlanField("PLACE", mPlace)
    If mPlaceMove Then PlanField("PLACE_BEFORE", mPlaceFrom)
    PlanField("CAT", mCat)
    PlanField("WHO", mWho)
    PlanField("DOC", mDoc)
    PlanField("MARK", mMark)
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", IIf(mMode = SR_ADD, mBal, 0))
    PlanField("BAL_AFTER", bal2)
    If dup <> "" Then PlanField("DUP", dup)
    ' one operation, the writes grouped by sheet: _SYS (event №, line №, new EI) → «Наличие» → _SPR → _ART → «Иной приход»
    If newEvent Then PlanSetValue(SYS_SHEET, EventCounterKey(mType), 1, evNum + 1, False)
    PlanSetValue(SYS_SHEET, SK_NEXT_SPL, 1, nL + 1, False)
    If mMode = SR_NEW Then
        PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, n + 1, False)
        PlanNewStockRow(n, canon, mQty, mName, mArt, mUnit, mPlace, mCat, NewEIState(mType), SrcStype(mType) & " " & evId, SrcStype(mType))
    Else
        PlanSetValue(SH_STOCK, n, SC_QTY, bal2, False)
        If mPlaceMove Then PlanSetValue(SH_STOCK, n, SC_PLACE, mPlace, False)
        If mStateBack Then PlanSetValue(SH_STOCK, n, SC_STATE, EI_ST_ACTIVE, False)
    End If
    PlanSprRow(nL, evId, canon, mQty, r, dupKey)
    If mMode = SR_NEW And ByArticle(mType) Then
        PlanSetValue(SH_ART, aRow, AR_KEY, mArtKey, False)
        PlanSetValue(SH_ART, aRow, AR_EI, canon, False)
        PlanSetValue(SH_ART, aRow, AR_ART, mArt, False)
    End If
    PlanRowValues(r, nL, evId, canon, mMode = SR_ADD, mBal, bal2, status, ctl)
    PlanLockBits(SH_SPECIAL, r, 0, XC_LAST, SPECIAL_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(n, bal2)
    res = ApplyOperation(0, 0)
    If Left(res, 6) = "ERR-RB" Then SetCtl(r, "Не проведено: " & Mid(res, 8))
    SpecialPostRow = res
    Exit Function
EH:
    ' only the checks can get here (ApplyOperation handles its own errors): nothing was written
    SpecialPostRow = "ERR-SYS:внутренняя ошибка проверки строки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    On Error Resume Next
    mInHandler = False
    SetCtl(r, "Не проведено: " & Mid(SpecialPostRow, 9))
End Function

' «Провести» for a block of rows (the rule D-059): every visible row with a quantity F is posted by its own operation;
' a row without F is skipped, so is a row already posted, cancelled or a copy; hidden rows (a filter) are not touched.
' A problem of one row does not stop the others; a system error of WMS stops the whole block. bShared: the lines of one
' source in the block get one event № (one acceptance with several lines: the first line takes a new №, the next ones
' join it); a row with its own № in C keeps it. bContinue: a further block of the same selection (the events of the
' previous blocks stay). Result as WmsReceipt.ReceiptPostRange:
' "OK=…;ERR=…;SKIP=…;STOP=<1-based row, 0 = none>", first rejection, system error.
Function SpecialPostRange(r0 As Long, r1 As Long, bShared As Boolean, bContinue As Boolean) As String
    Dim r As Long, res As String, nOk As Long, nErr As Long, nSkip As Long, sh As Object, firstErr As String, sysErr As String
    Dim stopRow As Long, ti As Integer, sJoin As String, hadC As Boolean, i As Integer
    On Error GoTo EH
    If Not bContinue Then
        For i = 0 To 4
            gXev(i) = ""
        Next i
    End If
    sh = SpecialSheet()
    For r = r0 To r1
        If sh.getRows().getByIndex(r).IsVisible Then
            If sh.getCellByPosition(XC_QTY, r).getType() = com.sun.star.table.CellContentType.EMPTY Then
                nSkip = nSkip + 1
            ElseIf SpecialRowKind(r) = "OPEN" Then
                ti = SrcOfText(sh.getCellByPosition(XC_TYPE, r).getString())
                hadC = (Trim(sh.getCellByPosition(XC_EVENT, r).getString()) <> "")
                sJoin = ""
                If bShared And ti >= 0 And Not hadC Then sJoin = gXev(ti)
                res = SpecialPostRow(r, sJoin, False)
                If WmsReceipt.IsSystemResult(res) Then
                    If Left(res, 3) = "OK:" Then nOk = nOk + 1
                    sysErr = "строка " & (r + 1) & ": " & res
                    stopRow = r + 1
                    Exit For
                ElseIf Left(res, 3) = "OK:" Then
                    nOk = nOk + 1
                    If bShared And ti >= 0 And Not hadC And gXev(ti) = "" Then gXev(ti) = sh.getCellByPosition(XC_EVENT, r).getString()
                ElseIf Left(res, 5) = "SKIP:" Then
                    nSkip = nSkip + 1
                Else
                    nErr = nErr + 1
                    If firstErr = "" Then firstErr = "строка " & (r + 1) & ": " & Mid(res, InStr(res, ":") + 1)
                End If
            Else
                nSkip = nSkip + 1
            End If
        End If
    Next r
DONE:
    SpecialPostRange = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";STOP=" & stopRow & Chr(10) & firstErr & Chr(10) & sysErr
    Exit Function
EH:
    sysErr = "строка " & (r + 1) & ": ERR-SYS:внутренняя ошибка проведения блока: " & Error$ & " (код " & Err & ")"
    stopRow = r + 1
    Resume DONE
End Function

' the number of visible rows of r0..r1 that «Провести» would post, and how many sources they have (for the question
' «одно поступление?»)
Function PostableRows(r0 As Long, r1 As Long, ByRef nSources As Integer) As Long
    Dim r As Long, sh As Object, n As Long, seen(4) As Boolean, ti As Integer, i As Integer
    sh = SpecialSheet()
    nSources = 0
    For r = r0 To r1
        If sh.getRows().getByIndex(r).IsVisible Then
            If sh.getCellByPosition(XC_QTY, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
                If SpecialRowKind(r) = "OPEN" Then
                    n = n + 1
                    ti = SrcOfText(sh.getCellByPosition(XC_TYPE, r).getString())
                    If ti >= 0 Then seen(ti) = True
                End If
            End If
        End If
    Next r
    For i = 0 To 4
        If seen(i) Then nSources = nSources + 1
    Next i
    PostableRows = n
End Function

' ================================================================ a posted line: its registry data

' the posted line of row r (after SpecialRowKind = POSTED / DELETED): _SPR data, its EI; "" or the reason
Private Function PostedLine(r As Long, ByRef nL As Long, ByRef sd As Variant, ByRef n As Long, ByRef canon As String) As String
    nL = gXn
    If Not SprRegistered(nL, sd) Then
        PostedLine = "строка № " & nL & " не найдена в " & SH_SPR
        Exit Function
    End If
    canon = CStr(sd(SR_EI))
    If Not WmsOrders.StrictEI(canon, n) Then
        PostedLine = "в " & SH_SPR & " у строки № " & nL & " неверный ЕИ «" & canon & "»"
        Exit Function
    End If
    PostedLine = WmsIssue.StockRowProblem(n, canon)
End Function

' "" when a row of this kind is a posted line that an action may change
Private Function PostedRowProblem(kind As String, sAction As String) As String
    Select Case kind
    Case "POSTED"
        PostedRowProblem = ""
    Case "DELETED"
        PostedRowProblem = "строка № " & gXn & " уже удалена (сторно)"
    Case "COPY", "FOREIGN"
        PostedRowProblem = "строка — КОПИЯ, её можно только очистить кнопкой «Очистить»"
    Case "OPEN"
        PostedRowProblem = "строка не проведена — «" & sAction & "» нужно только для проведённых строк"
    Case Else
        PostedRowProblem = "строка пуста"
    End Select
End Function

' ================================================================ «Исправить» — SP_FIX (one composite operation)
' Quantity, date, place, document, «кто передал», old marking — any line. The line that created its EI (NEW) also: name,
' article (a part: only to an article no other EI has), category, unit (only while nothing was issued out of the EI
' and no other line added to it) — they are the card of the EI. The source, the EI and the event never change (another
' identity of the goods: storno and a new receipt). The balance of the EI must not become negative.

Function SpecialFixRow(r As Long, vQty As Variant, vDate As Variant, vPlace As Variant, vCat As Variant, vDoc As Variant, vWho As Variant, _
    vMark As Variant, vName As Variant, vArt As Variant, vUnit As Variant) As String
    Dim why As String, kind As String, nL As Long, sd As Variant, n As Long, canon As String, sh As Object, q1 As Double, q2 As Double
    Dim d1 As Double, d2 As Double, msg As String, p1 As String, p2 As String, c1 As String, c2 As String, doc1 As String, doc2 As String
    Dim who1 As String, who2 As String, mk1 As String, mk2 As String, n1 As String, n2 As String, a1 As String, a2 As String, u1 As String
    Dim u2 As String, card As Variant, bal As Double, bal2 As Double, isNew As Boolean, isPart As Boolean, key1 As String, key2 As String
    Dim st As Integer, aRow As Long, en As Long, ecanon As String, nLive As Long, nIss As Long, nRet As Long, res As String, placeNow As String
    Dim dupKey As String, dup As String, fixes As Long, ctl As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        SpecialFixRow = "ERR:выберите строку иного прихода (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = SpecialSheet()
    kind = SpecialRowKind(r)
    why = PostedRowProblem(kind, "Исправить")
    If why <> "" Then
        SpecialFixRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        SpecialFixRow = why
        Exit Function
    End If
    why = PostedLine(r, nL, sd, n, canon)
    If why <> "" Then
        SpecialFixRow = "ERR-SYS:" & why
        Exit Function
    End If
    mType = SrcIndexOfCode(CStr(sd(SR_TYPE)))
    mCode = CStr(sd(SR_TYPE))
    isNew = (CStr(sd(SR_MODE)) = SR_NEW)
    isPart = (ByArticle(mType) Or CStr(sd(SR_IDENT)) = SrcStype(2))
    card = WmsIssue.StockData(n)
    If Not WmsIssue.StockQty(n, bal) Then
        SpecialFixRow = "ERR-SYS:остаток " & canon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    q1 = sd(SR_QTY)
    d1 = sh.getCellByPosition(XC_DATE, r).getValue()
    p1 = sh.getCellByPosition(XC_PLACE, r).getString()
    c1 = sh.getCellByPosition(XC_CAT, r).getString()
    doc1 = sh.getCellByPosition(XC_DOC, r).getString()
    who1 = sh.getCellByPosition(XC_WHO, r).getString()
    mk1 = sh.getCellByPosition(XC_MARK, r).getString()
    n1 = sh.getCellByPosition(XC_NAME, r).getString()
    a1 = sh.getCellByPosition(XC_ART, r).getString()
    u1 = sh.getCellByPosition(XC_UNIT, r).getString()
    If Not WmsOrders.ParseQtyValueAs(vQty, "количество", q2, msg) Then
        SpecialFixRow = "ERR:" & msg
        Exit Function
    End If
    If VarType(vDate) = 8 Then
        If Not WmsOrders.ParseDateTextAs(CStr(vDate), "дата поступления", d2, msg) Then
            SpecialFixRow = "ERR:" & msg
            Exit Function
        End If
    Else
        d2 = Int(CDbl(vDate))
    End If
    p2 = Trim(CStr(vPlace))
    If p2 = "" Then p2 = p1
    c2 = Trim(CStr(vCat))
    doc2 = Trim(CStr(vDoc))
    who2 = Trim(CStr(vWho))
    mk2 = Trim(CStr(vMark))
    n2 = Trim(CStr(vName))
    a2 = Trim(CStr(vArt))
    u2 = Trim(CStr(vUnit))
    If Not isNew Then
        ' an addition to a part: the name, the article, the unit and the category are those of its EI (its card)
        If n2 <> n1 Or ArticleKey(a2) <> ArticleKey(a1) Or LCase(u2) <> LCase(u1) Or LCase(c2) <> LCase(c1) Then
            SpecialFixRow = "ERR:у строки пополнения наименование, артикул, единица и категория — от карточки " & canon _
                & "; их исправляют в строке, создавшей ЕИ, или сторно и новым приходом"
            Exit Function
        End If
    Else
        If n2 = "" Then
            SpecialFixRow = "ERR:наименование не может быть пустым"
            Exit Function
        End If
        If u2 = "" Then
            SpecialFixRow = "ERR:единица измерения не может быть пустой"
            Exit Function
        End If
    End If
    ' a part: the article is its key — another article only when no other EI has it
    key1 = ArticleKey(a1)
    key2 = ArticleKey(a2)
    If isNew And isPart And key2 <> key1 Then
        If key2 = "" Then
            SpecialFixRow = "ERR:артикул детали не может быть пустым"
            Exit Function
        End If
        st = ArtLookup(key2, en, ecanon, aRow, why)
        If st = 1 Then
            SpecialFixRow = "ERR:артикул «" & a2 & "» уже принадлежит " & ecanon & " — ЕИ детали не меняется исправлением; нужен разбор (сторно и приход в " & ecanon & ")"
            Exit Function
        ElseIf st = 2 Then
            SpecialFixRow = "ERR:" & why
            Exit Function
        ElseIf st = 3 Then
            SpecialFixRow = "ERR-SYS:" & why
            Exit Function
        End If
        st = ArtLookup(key1, en, ecanon, aRow, why)
        If st <> 1 Or en <> n Then
            SpecialFixRow = "ERR-SYS:" & IIf(why <> "", why, "в индексе артикулов нет записи «" & key1 & "» → " & canon & " — нужна самопроверка")
            Exit Function
        End If
    End If
    ' the unit: only while the EI has no history in the old one
    If isNew And LCase(u2) <> LCase(u1) Then
        If Not LiveLinesOfEI(canon, nLive) Or Not WmsReturn.LiveDependents(canon, nIss, nRet) Then
            SpecialFixRow = "ERR-SYS:не удалось проверить движения " & canon & " (формула движка Calc) — исправление не выполнено"
            Exit Function
        End If
        If nLive > 1 Or nIss > 0 Or nRet > 0 Then
            SpecialFixRow = "ERR:единицу " & canon & " нельзя изменить: по нему есть " & IIf(nLive > 1, "другие приходы", "действующие выдачи/возвраты") _
                & " в единице «" & u1 & "». Пересчёт единиц не настроен"
            Exit Function
        End If
    End If
    bal2 = WmsIssue.Round3(bal - q1 + q2)
    If bal2 < -0.0000001 Then
        SpecialFixRow = "ERR:остаток " & canon & " — " & WmsIssue.QtyText(bal) & ": из прихода уже выдано " & WmsIssue.QtyText(WmsIssue.Round3(q1 - bal)) _
            & ", уменьшение до " & WmsIssue.QtyText(q2) & " сделало бы остаток отрицательным; сначала удалите (сторно) выдачи"
        Exit Function
    End If
    If Abs(q2 - q1) < 0.0000001 And d2 = Int(d1) And p2 = p1 And c2 = c1 And doc2 = doc1 And who2 = who1 And mk2 = mk1 And n2 = n1 And a2 = a1 _
        And u2 = u1 Then
        SpecialFixRow = "SKIP:ничего не изменилось"
        Exit Function
    End If
    placeNow = CStr(card(SC_PLACE))
    mArtKey = IIf(isPart, key2, "")
    dupKey = DupKeyOf(mCode, mArtKey, IIf(isNew, n2, n1), q2, d2, doc2)
    dup = FindSpecialDuplicate(dupKey, nL)
    If VarType(sd(SR_FIXES)) = 5 Then fixes = CLng(sd(SR_FIXES))
    ctl = "Исправлено: " & canon & ", остаток " & WmsIssue.QtyText(bal) & " → " & WmsIssue.QtyText(bal2) & " " & u2 & " · " & sd(SR_EVENT)
    If p2 <> p1 And p2 <> placeNow Then ctl = ctl & " · место ЕИ «" & placeNow & "» → «" & p2 & "»"
    If dup <> "" Then ctl = ctl & " · Возможный дубль: " & dup
    PlanBegin("SP_FIX", CStr(nL))
    PlanField("SPL", nL)
    PlanField("EVENT", sd(SR_EVENT))
    PlanField("SRC", mCode)
    PlanField("EI", canon)
    PlanField("MODE", sd(SR_MODE))
    PlanField("OLD_QTY", q1)
    PlanField("QTY", q2)
    PlanField("OLD_DATE", Format(d1, "YYYY-MM-DD"))
    PlanField("DATE", Format(d2, "YYYY-MM-DD"))
    PlanField("OLD_PLACE", p1)
    PlanField("PLACE", p2)
    If n2 <> n1 Then PlanField("NAME", n2)
    If a2 <> a1 Then
        PlanField("OLD_ART", a1)
        PlanField("ART", a2)
        If isPart Then PlanField("ART_KEY", key2)
    End If
    If u2 <> u1 Then PlanField("UNIT", u2)
    If c2 <> c1 Then PlanField("CAT", c2)
    PlanField("DOC", doc2)
    PlanField("WHO", who2)
    PlanField("MARK", mk2)
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", bal)
    PlanField("BAL_AFTER", bal2)
    If dup <> "" Then PlanField("DUP", dup)
    ' «Наличие» → _SPR → _ART → «Иной приход»
    If Abs(bal2 - bal) > 0.0000001 Then PlanSetValue(SH_STOCK, n, SC_QTY, bal2, False)
    If p2 <> placeNow And p2 <> p1 Then PlanSetValue(SH_STOCK, n, SC_PLACE, p2, False)
    If isNew Then
        If n2 <> CStr(card(SC_NAME)) Then PlanSetValue(SH_STOCK, n, SC_NAME, n2, False)
        If a2 <> CStr(card(SC_ART)) Then PlanSetValue(SH_STOCK, n, SC_ART, a2, False)
        If u2 <> CStr(card(SC_UNIT)) Then PlanSetValue(SH_STOCK, n, SC_UNIT, u2, False)
        If c2 <> CStr(card(SC_CAT)) Then PlanSetValue(SH_STOCK, n, SC_CAT, c2, False)
    End If
    PlanSetValue(SH_SPR, nL, SR_QTY, q2, False)
    If CLng(sd(SR_ROW)) <> r Then PlanSetValue(SH_SPR, nL, SR_ROW, r, False)
    PlanSetValue(SH_SPR, nL, SR_DATE, d2, False)
    PlanSetValue(SH_SPR, nL, SR_DUP, WmsOrders.DupHash(dupKey), False)
    PlanSetValue(SH_SPR, nL, SR_DUPKEY, dupKey, False)
    If isPart And key2 <> key1 Then PlanSetValue(SH_SPR, nL, SR_ARTKEY, key2, False)
    PlanSetValue(SH_SPR, nL, SR_FIXES, fixes + 1, False)
    If isNew And isPart And key2 <> key1 Then
        PlanSetValue(SH_ART, aRow, AR_KEY, key2, False)
        PlanSetValue(SH_ART, aRow, AR_ART, a2, False)
    End If
    PlanSetValue(SH_SPECIAL, r, XC_QTY, q2, False)
    PlanSetValue(SH_SPECIAL, r, XC_DATE, d2, False)
    PlanSetValue(SH_SPECIAL, r, XC_PLACE, p2, False)
    If c2 <> c1 Then PlanSetValue(SH_SPECIAL, r, XC_CAT, c2, False)
    If doc2 <> doc1 Then PlanSetValue(SH_SPECIAL, r, XC_DOC, doc2, False)
    If who2 <> who1 Then PlanSetValue(SH_SPECIAL, r, XC_WHO, who2, False)
    If mk2 <> mk1 Then PlanSetValue(SH_SPECIAL, r, XC_MARK, mk2, False)
    If n2 <> n1 Then PlanSetValue(SH_SPECIAL, r, XC_NAME, n2, False)
    If a2 <> a1 Then PlanSetValue(SH_SPECIAL, r, XC_ART, a2, False)
    If u2 <> u1 Then PlanSetValue(SH_SPECIAL, r, XC_UNIT, u2, False)
    If isNew Then
        If sh.getCellByPosition(XC_BEFORE, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then PlanSetValue(SH_SPECIAL, r, XC_BEFORE, "", False)
    Else
        PlanSetValue(SH_SPECIAL, r, XC_BEFORE, WmsIssue.Round3(bal - q1), False)
    End If
    PlanSetValue(SH_SPECIAL, r, XC_AFTER, bal2, False)
    PlanSetValue(SH_SPECIAL, r, XC_STATUS, LineStatus(LineKind(sd), True), False)
    PlanDerived(SH_SPECIAL, r, XC_CTL, ctl)
    PlanLockBits(SH_SPECIAL, r, 0, XC_LAST, SPECIAL_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(n, bal2)
    res = ApplyOperation(0, 0)
    SpecialFixRow = res
    Exit Function
EH:
    SpecialFixRow = "ERR-SYS:внутренняя ошибка проверки исправления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Удалить» — SP_DEL (storno)
' The line leaves its EI: the balance goes down by its quantity, _SPR gets STORNO, the row stays as the history with the
' status «Удалено (сторно)», its № and the event № are never reused. A line that is the last live one of its EI (every EI
' of Офис, Производство, Старый склад, Иной; the only line of a part): refused while the EI has live issues or returns
' (D-069) — first their storno — and the EI becomes «Приход удалён (сторно)». A part with other live lines: refused only
' when the balance would become negative (what was issued out of the EI stays covered by the other lines).

Function SpecialDeleteRow(r As Long) As String
    SpecialDeleteRow = SpDelete(r, False)
End Function

' every check of «Удалить» without the operation: "OK:" when the storno may be done (the button asks for the confirmation
' only then), otherwise exactly the refusal SpecialDeleteRow would give. Nothing is written.
Function SpecialDeleteCheck(r As Long) As String
    SpecialDeleteCheck = SpDelete(r, True)
End Function

Private Function SpDelete(r As Long, checkOnly As Boolean) As String
    Dim why As String, kind As String, nL As Long, sd As Variant, n As Long, canon As String, sh As Object, q As Double, bal As Double
    Dim bal2 As Double, nLive As Long, nIss As Long, nRet As Long, bLast As Boolean, unit As String, nAdj As Long
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        SpDelete = "ERR:выберите строку иного прихода (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = SpecialSheet()
    kind = SpecialRowKind(r)
    If kind = "DELETED" Then
        SpDelete = "SKIP:строка № " & gXn & " уже удалена (сторно)"
        Exit Function
    End If
    why = PostedRowProblem(kind, "Удалить")
    If why <> "" Then
        SpDelete = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        SpDelete = why
        Exit Function
    End If
    why = PostedLine(r, nL, sd, n, canon)
    If why <> "" Then
        SpDelete = "ERR-SYS:" & why
        Exit Function
    End If
    q = sd(SR_QTY)
    unit = sh.getCellByPosition(XC_UNIT, r).getString()
    If Not WmsIssue.StockQty(n, bal) Then
        SpDelete = "ERR-SYS:остаток " & canon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If Not LiveLinesOfEI(canon, nLive) Then
        SpDelete = "ERR-SYS:не удалось проверить строки прихода " & canon & " (формула движка Calc) — сторно не выполнено"
        Exit Function
    End If
    ' a migrated EI (source «Перенос…») existed before its receipt lines: cancelling the last of them keeps the EI
    bLast = (nLive <= 1) And Left(CStr(WmsIssue.StockData(n)(SC_SRC)), 7) <> "Перенос"
    If bLast Then
        ' D-069: no live chain «issue → return» may refer to a cancelled receipt
        If Not WmsReturn.LiveDependents(canon, nIss, nRet) Then
            SpDelete = "ERR-SYS:не удалось проверить связанные выдачи и возвраты " & canon & " (формула движка Calc) — сторно не выполнено"
            Exit Function
        End If
        If nIss > 0 Or nRet > 0 Then
            SpDelete = "ERR:По этому ЕИ (" & canon & ") есть связанные выдачи/возвраты: действующих выдач " & nIss & ", возвратов " & nRet _
                & IIf(bal < q - 0.0000001, " (уже выдано " & WmsIssue.QtyText(WmsIssue.Round3(q - bal)) & " " & unit & ")", "") _
                & ". Сначала выполните сторно зависимых операций: " & IIf(nRet > 0, "возвратов этого ЕИ (лист «" & SH_RETURNS & "»), затем ", "") _
                & "выдач (лист «" & SH_ISSUES & "»)"
            Exit Function
        End If
        ' D-069 for the corrections (Final Core): a live move, write-off or inventory correction refers to the receipt too
        If Not WmsAdjust.LiveAdjustments(canon, nAdj) Then
            SpDelete = "ERR-SYS:не удалось проверить корректировки " & canon & " (формула движка Calc) — сторно не выполнено"
            Exit Function
        End If
        If nAdj > 0 Then
            SpDelete = "ERR:По этому ЕИ (" & canon & ") есть действующие корректировки: " & nAdj & " (лист «" & SH_ADJUST _
                & "»: перемещения, списания, инвентаризация). Сначала выполните их сторно"
            Exit Function
        End If
    End If
    If bal < q - 0.0000001 Then
        SpDelete = "ERR:остаток " & canon & " — " & WmsIssue.QtyText(bal) & " " & unit & ", а приход строки — " & WmsIssue.QtyText(q) _
            & ": из него уже выдано " & WmsIssue.QtyText(WmsIssue.Round3(q - bal)) & ", сторно сделало бы остаток отрицательным (" _
            & WmsIssue.QtyText(WmsIssue.Round3(q - bal)) & " не покрыто); сначала удалите (сторно) выдачи"
        Exit Function
    End If
    If checkOnly Then
        SpDelete = "OK:сторно строки № " & nL & " допустимо"
        Exit Function
    End If
    bal2 = WmsIssue.Round3(bal - q)
    PlanBegin("SP_DEL", CStr(nL))
    PlanField("SPL", nL)
    PlanField("EVENT", sd(SR_EVENT))
    PlanField("SRC", sd(SR_TYPE))
    PlanField("EI", canon)
    PlanField("MODE", sd(SR_MODE))
    PlanField("QTY", q)
    PlanField("LAST", IIf(bLast, 1, 0))
    PlanField("ROW", r + 1)
    PlanField("BAL_BEFORE", bal)
    PlanField("BAL_AFTER", bal2)
    PlanSetValue(SH_STOCK, n, SC_QTY, bal2, False)
    If bLast Then PlanSetValue(SH_STOCK, n, SC_STATE, EI_ST_STORNO, False)
    PlanSetValue(SH_SPR, nL, SR_STATE, RV_STORNO, False)
    If CLng(sd(SR_ROW)) <> r Then PlanSetValue(SH_SPR, nL, SR_ROW, r, False)
    PlanSetValue(SH_SPECIAL, r, XC_STATUS, ST_DELETED, False)
    PlanDerived(SH_SPECIAL, r, XC_CTL, "Удалено (сторно): остаток " & canon & " " & WmsIssue.QtyText(bal) & " → " & WmsIssue.QtyText(bal2) & " " & unit _
        & IIf(bLast, " · ЕИ — «" & EI_ST_STORNO & "»", ""))
    PlanLockBits(SH_SPECIAL, r, 0, XC_LAST, SPECIAL_LOCKS_POSTED)
    WmsOrders.PlanStockMirror(n, bal2)
    SpDelete = ApplyOperation(0, 0)
    Exit Function
EH:
    SpDelete = "ERR-SYS:внутренняя ошибка проверки удаления: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Разобрать» — SP_IDENTIFY (the same EI, the same physical batch)
' A line of «Иной» whose EI still «Требует разбора»: the kind of source (Поставщик, Офис, Производство, Детали, Старый
' склад), the normal name, the article and the category become the card of the very same EI — no new EI. A part needs an
' article no other EI has (then the index learns it). Splitting one unidentified line into several goods is not done here.

' «Тип источника» chosen for an identified line → its text; "" when unknown
Function IdentStype(ByVal s As String) As String
    Dim i As Integer
    s = LCase(Trim(s))
    If s = LCase(STYPE_SUPPLIER) Then
        IdentStype = STYPE_SUPPLIER
        Exit Function
    End If
    i = SrcOfText(s)
    If i >= 0 And i <= 3 Then IdentStype = SrcStype(i)
End Function

Function SpecialIdentifyRow(r As Long, vType As Variant, vName As Variant, vArt As Variant, vCat As Variant) As String
    Dim why As String, kind As String, nL As Long, sd As Variant, n As Long, canon As String, sh As Object, card As Variant, stype As String
    Dim n2 As String, a2 As String, c2 As String, key As String, st As Integer, en As Long, ecanon As String, aRow As Long, src As String
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        SpecialIdentifyRow = "ERR:выберите строку иного прихода (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    sh = SpecialSheet()
    kind = SpecialRowKind(r)
    why = PostedRowProblem(kind, "Разобрать")
    If why <> "" Then
        SpecialIdentifyRow = "ERR:" & why
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        SpecialIdentifyRow = why
        Exit Function
    End If
    why = PostedLine(r, nL, sd, n, canon)
    If why <> "" Then
        SpecialIdentifyRow = "ERR-SYS:" & why
        Exit Function
    End If
    If CStr(sd(SR_TYPE)) <> "OTH" Or CStr(sd(SR_MODE)) <> SR_NEW Then
        SpecialIdentifyRow = "ERR:«Разобрать» — только для строки иного прихода (тип «Иной»), создавшей ЕИ"
        Exit Function
    End If
    card = WmsIssue.StockData(n)
    If CStr(card(SC_STATE)) <> EI_ST_REVIEW Then
        SpecialIdentifyRow = "ERR:" & canon & " уже разобран (состояние «" & card(SC_STATE) & "»)"
        Exit Function
    End If
    stype = IdentStype(CStr(vType))
    If stype = "" Then
        SpecialIdentifyRow = "ERR:укажите, что это за приход: Поставщик, Офис, Производство, Детали или Старый склад"
        Exit Function
    End If
    n2 = Trim(CStr(vName))
    a2 = Trim(CStr(vArt))
    c2 = Trim(CStr(vCat))
    If n2 = "" Then
        SpecialIdentifyRow = "ERR:укажите нормальное наименование"
        Exit Function
    End If
    If stype = SrcStype(2) Then
        key = ArticleKey(a2)
        If key = "" Then
            SpecialIdentifyRow = "ERR:для детали артикул обязателен — по нему WMS находит её ЕИ"
            Exit Function
        End If
        st = ArtLookup(key, en, ecanon, aRow, why)
        If st = 1 Then
            SpecialIdentifyRow = "ERR:артикул «" & a2 & "» уже принадлежит " & ecanon & " — это та же деталь: удалите (сторно) этот иной приход " _
                & "и проведите приход детали — он пополнит " & ecanon
            Exit Function
        ElseIf st = 2 Then
            SpecialIdentifyRow = "ERR:" & why
            Exit Function
        ElseIf st = 3 Then
            SpecialIdentifyRow = "ERR-SYS:" & why
            Exit Function
        End If
        aRow = WmsOrders.LastRow(ArtSheet()) + 1
    End If
    src = CStr(card(SC_SRC)) & " → " & stype
    PlanBegin("SP_IDENTIFY", CStr(nL))
    PlanField("SPL", nL)
    PlanField("EVENT", sd(SR_EVENT))
    PlanField("EI", canon)
    PlanField("TYPE", stype)
    PlanField("OLD_NAME", card(SC_NAME))
    PlanField("NAME", n2)
    PlanField("OLD_ART", card(SC_ART))
    PlanField("ART", a2)
    If key <> "" Then PlanField("ART_KEY", key)
    PlanField("OLD_CAT", card(SC_CAT))
    PlanField("CAT", c2)
    PlanField("ROW", r + 1)
    ' «Наличие» → _SPR → _ART → «Иной приход»
    If n2 <> CStr(card(SC_NAME)) Then PlanSetValue(SH_STOCK, n, SC_NAME, n2, False)
    If a2 <> CStr(card(SC_ART)) Then PlanSetValue(SH_STOCK, n, SC_ART, a2, False)
    If c2 <> CStr(card(SC_CAT)) Then PlanSetValue(SH_STOCK, n, SC_CAT, c2, False)
    PlanSetValue(SH_STOCK, n, SC_STATE, EI_ST_ACTIVE, False)
    PlanSetValue(SH_STOCK, n, SC_SRC, src, False)
    PlanSetValue(SH_STOCK, n, SC_STYPE, stype, False)
    PlanSetValue(SH_SPR, nL, SR_IDENT, stype, False)
    If key <> "" Then PlanSetValue(SH_SPR, nL, SR_ARTKEY, key, False)
    If CLng(sd(SR_ROW)) <> r Then PlanSetValue(SH_SPR, nL, SR_ROW, r, False)
    If key <> "" Then
        PlanSetValue(SH_ART, aRow, AR_KEY, key, False)
        PlanSetValue(SH_ART, aRow, AR_EI, canon, False)
        PlanSetValue(SH_ART, aRow, AR_ART, a2, False)
    End If
    If n2 <> sh.getCellByPosition(XC_NAME, r).getString() Then PlanSetValue(SH_SPECIAL, r, XC_NAME, n2, False)
    If a2 <> sh.getCellByPosition(XC_ART, r).getString() Then PlanSetValue(SH_SPECIAL, r, XC_ART, a2, False)
    If c2 <> sh.getCellByPosition(XC_CAT, r).getString() Then PlanSetValue(SH_SPECIAL, r, XC_CAT, c2, False)
    PlanSetValue(SH_SPECIAL, r, XC_STATUS, LineStatus(XK_IDENT, CLng(Val(CStr(sd(SR_FIXES)))) > 0), False)
    PlanDerived(SH_SPECIAL, r, XC_CTL, "Разобрано: " & stype & ", «" & n2 & "»" & IIf(a2 <> "", ", артикул " & a2, "") & " — тот же " & canon)
    PlanLockBits(SH_SPECIAL, r, 0, XC_LAST, SPECIAL_LOCKS_POSTED)
    SpecialIdentifyRow = ApplyOperation(0, 0)
    Exit Function
EH:
    SpecialIdentifyRow = "ERR-SYS:внутренняя ошибка проверки разбора: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ «Проверить»

' An open row: all the checks of «Провести», the result in R. A posted row: the line, its EI and the registry —
' OK:<text> or ERR:<text> for a message.
Function SpecialCheckRow(r As Long) As String
    Dim kind As String, why As String, nL As Long, sd As Variant, n As Long, canon As String, bal As Double, card As Variant
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        SpecialCheckRow = "ERR:выберите строку иного прихода (не заголовок)"
        Exit Function
    End If
    On Error GoTo EH
    kind = SpecialRowKind(r)
    Select Case kind
    Case "OPEN"
        mInHandler = True
        why = CheckRow(r, True, "")
        If why <> "" Then
            SetCtl(r, "Ошибка: " & why)
            SpecialCheckRow = "ERR:" & why
        Else
            SetCtl(r, "Можно провести: " & DescribeCap() & IIf(mWarn <> "", " · " & mWarn, ""))
            SpecialCheckRow = "OK:можно провести: " & Describe() & IIf(mWarn <> "", " · " & mWarn, "")
        End If
        mInHandler = False
    Case "POSTED", "DELETED"
        why = PostedLine(r, nL, sd, n, canon)
        If why <> "" Then
            SpecialCheckRow = "ERR:строка № " & nL & ": " & why
            Exit Function
        End If
        card = WmsIssue.StockData(n)
        If Not WmsIssue.StockQty(n, bal) Then bal = 0
        SpecialCheckRow = "OK:строка № " & nL & " (" & IIf(kind = "DELETED", "удалена (сторно)", "действует") & "): " & sd(SR_EVENT) & ", " _
            & IIf(CStr(sd(SR_MODE)) = SR_ADD, "пополнение ", "") & canon & " «" & card(SC_NAME) & "», " & WmsIssue.QtyText(sd(SR_QTY)) & " " _
            & card(SC_UNIT) & "; остаток ЕИ — " & WmsIssue.QtyText(bal) & ", место «" & card(SC_PLACE) & "», состояние «" & card(SC_STATE) & "»"
    Case "COPY", "FOREIGN"
        SpecialCheckRow = "ERR:строка — КОПИЯ" & IIf(gXwhy <> "", " (" & gXwhy & ")", "") & ", её можно только очистить кнопкой «Очистить»"
    Case Else
        SpecialCheckRow = "ERR:строка пуста — укажите тип прихода, наименование, количество, единицу, дату и место"
    End Select
    Exit Function
EH:
    mInHandler = False
    SpecialCheckRow = "ERR:внутренняя ошибка проверки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' ================================================================ the change handler: preview of open rows

' «Иной приход» sheet event "Content changed". Leaves at once for columns that do not matter, never scans the sheet, is
' guarded against re-entry and ignores changes made by operations.
Sub OnSpecialChange(Optional oTarget As Variant)
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
    ' the inputs A (a key that appeared by paste) and B..M matter; N..R are written by WMS, S is free text
    If ra.StartColumn > XC_MARK Then Exit Sub
    mInHandler = True
    WmsInit()
    r1 = ra.EndRow
    If r1 - ra.StartRow + 1 > PREVIEW_MAX_ROWS Then
        last = WmsOrders.LastRow(SpecialSheet())
        If r1 > last Then r1 = last
    End If
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: проверка строки иного прихода")
    inCtx = True
    For r = IIf(ra.StartRow < 1, 1, ra.StartRow) To r1
        n = n + 1
        If n > PREVIEW_MAX_ROWS Then Exit For
        PreviewSpecialRow(r)
    Next r
    um.leaveUndoContext()
    mInHandler = False
    Exit Sub
EH:
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
    mInHandler = False
End Sub

' The preview of an open row: a part with a known article shows its EI (N), the current balance (O), and — when empty —
' its name, unit, place and category (D G I J); «Контроль» R: the first problem of the filled inputs, otherwise what
' the row will do and what is still missing. A value the user typed is never replaced. Posted, cancelled and copy rows
' are never touched here; a № in A that WMS did not write marks the row КОПИЯ.
Sub PreviewSpecialRow(r As Long)
    Dim sh As Object, kind As String, why As String
    sh = SpecialSheet()
    kind = SpecialRowKind(r)
    Select Case kind
    Case "FOREIGN"
        CopyCheckSpecial(r)
        Exit Sub
    Case "EMPTY"
        ClearDerived(r, True)
        Exit Sub
    Case "OPEN"
    Case Else
        Exit Sub
    End Select
    why = CheckRow(r, False, "")
    If why <> "" Then
        ClearDerived(r, False)
        SetCtl(r, "Ошибка: " & why)
        Exit Sub
    End If
    If mMode = SR_ADD Then
        If Trim(sh.getCellByPosition(XC_NAME, r).getString()) = "" Then sh.getCellByPosition(XC_NAME, r).setString(mName)
        If Trim(sh.getCellByPosition(XC_UNIT, r).getString()) = "" Then sh.getCellByPosition(XC_UNIT, r).setString(mUnit)
        If Trim(sh.getCellByPosition(XC_PLACE, r).getString()) = "" And mPlace <> "" Then sh.getCellByPosition(XC_PLACE, r).setString(mPlace)
        If Trim(sh.getCellByPosition(XC_CAT, r).getString()) = "" And mCat <> "" Then sh.getCellByPosition(XC_CAT, r).setString(mCat)
        WmsOrders.SetIfDiff(sh.getCellByPosition(XC_EI, r), mCanon)
        WmsOrders.SetIfDiff(sh.getCellByPosition(XC_BEFORE, r), mBal)
        If mHasQty Then
            WmsOrders.SetIfDiff(sh.getCellByPosition(XC_AFTER, r), WmsIssue.Round3(mBal + mQty))
        ElseIf sh.getCellByPosition(XC_AFTER, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then
            ClearCell(sh.getCellByPosition(XC_AFTER, r))
        End If
    Else
        ClearDerived(r, False)
    End If
    SetCtl(r, IIf(mMissing = "", "Можно провести: ", "") & DescribeCap() & IIf(mMissing <> "", " — укажите " & mMissing, "") _
        & IIf(mWarn <> "", " · " & mWarn, ""))
End Sub

' the WMS-filled cells of an open row (N O P Q, and R when bAll)
Private Sub ClearDerived(r As Long, bAll As Boolean)
    Dim sh As Object, cols As Variant, i As Integer
    sh = SpecialSheet()
    cols = Array(XC_EI, XC_BEFORE, XC_AFTER, XC_STATUS)
    For i = 0 To UBound(cols)
        If sh.getCellByPosition(cols(i), r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(cols(i), r))
    Next i
    If bAll Then
        If sh.getCellByPosition(XC_CTL, r).getType() <> com.sun.star.table.CellContentType.EMPTY Then ClearCell(sh.getCellByPosition(XC_CTL, r))
    End If
End Sub

' handler guard for the UI module (its own writes into «Иной приход» outside operations)
Sub SpecialHandlerOff(b As Boolean)
    mInHandler = b
End Sub

' ================================================================ «Очистить» — an unposted row or a copy (no accounting change)

Function SpecialClearRow(r As Long) As String
    Dim sh As Object, kind As String, wasProt As Boolean, i As Integer, pr As Object, x As Double, d As Variant
    WmsInit()
    If r < 1 Or r > MAX_SHEET_ROW Then
        SpecialClearRow = "ERR:выберите строку иного прихода (не заголовок)"
        Exit Function
    End If
    If gBusy Then
        SpecialClearRow = "BUSY:операция уже выполняется"
        Exit Function
    End If
    sh = SpecialSheet()
    kind = SpecialRowKind(r)
    Select Case kind
    Case "POSTED"
        SpecialClearRow = "ERR:строка — проведённая строка № " & gXn & " — для отмены используйте «Удалить»"
        Exit Function
    Case "DELETED"
        SpecialClearRow = "ERR:строка № " & gXn & " удалена (сторно) — строка хранит историю и не очищается"
        Exit Function
    Case "COPY"
        ' defence: never the registered row of a line, even when it carries a copy mark
        If sh.getCellByPosition(XC_NO, r).getType() = com.sun.star.table.CellContentType.VALUE Then
            x = sh.getCellByPosition(XC_NO, r).getValue()
            If x >= 1 And x = Int(x) And x <= MAX_SHEET_ROW Then
                If SprRegistered(CLng(x), d) Then
                    If VarType(d(SR_ROW)) = 5 Then
                        If CLng(d(SR_ROW)) = r Then
                            SpecialClearRow = "ERR:строка — зарегистрированная строка № " & CLng(x) & "; копией является другая строка с этим №"
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End Select
    mInHandler = True
    On Error GoTo EH
    wasProt = sh.isProtected()
    If wasProt Then sh.unprotect(PROTECT_PWD)
    sh.getCellRangeByPosition(0, r, XC_LAST, r).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME _
        + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA)
    ' a row inserted between posted rows or a copy carries the protection of its origin: back to the input layout
    For i = 0 To XC_LAST
        pr = sh.getCellByPosition(i, r).CellProtection
        pr.IsLocked = (Mid(SPECIAL_LOCKS_OPEN, i + 1, 1) = "1")
        sh.getCellByPosition(i, r).CellProtection = pr
    Next i
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
    SpecialClearRow = "OK:" & IIf(kind = "COPY" Or kind = "FOREIGN", "копия очищена", "строка очищена")
    Exit Function
EH:
    SpecialClearRow = "ERR:не удалось очистить строку: " & Error$
    On Error Resume Next
    If wasProt Then sh.protect(PROTECT_PWD)
    mInHandler = False
End Function

' ================================================================ «Главная»: special receipts not posted yet (spec §23)

' rows with a quantity (F) but no № (A): two Calc queries per block instead of a row loop. -1 when it fails.
Function UnpostedSpecial() As Long
    Dim sh As Object, last As Long, blocks As Variant, got As Variant, i As Long, j As Long, n As Long, flags As Long
    On Error GoTo EH
    sh = SpecialSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    flags = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + com.sun.star.sheet.CellFlags.STRING _
        + com.sun.star.sheet.CellFlags.FORMULA
    blocks = sh.getCellRangeByPosition(XC_NO, 1, XC_NO, last).queryEmptyCells().getRangeAddresses()
    For i = 0 To UBound(blocks)
        If i >= UNPOSTED_MAX_BLOCKS Then Exit For
        got = sh.getCellRangeByPosition(XC_QTY, blocks(i).StartRow, XC_QTY, blocks(i).EndRow).queryContentCells(flags).getRangeAddresses()
        For j = 0 To UBound(got)
            n = n + got(j).EndRow - got(j).StartRow + 1
        Next j
    Next i
    UnpostedSpecial = n
    Exit Function
EH:
    UnpostedSpecial = -1
End Function

' ================================================================ self-check (spec §49, a part of «WMS-Доктор»)
' Every line of _SPR (№, event, source, EI, mode, quantity, state), every EI of the special receipts against «Наличие»
' (its kind of source, a balance not above its live lines, a cancelled EI at 0 with the storno state, exactly one line that
' created it, additions only to parts), the index of the parts (each entry — a part with this article; no key twice; every
' part in the index), the counters above every number. Chunks, only the needed columns.

Function SpecialCheck() As String
    Dim spr As Object, stk As Object, art As Object, last As Long, r0 As Long, r1 As Long, i As Long, d As Variant, t0 As Long
    Dim nextSpl As Long, nextEI As Long, maxEv(4) As Long, nextEv(4) As Long, ti As Integer, num As Long, n As Long, nBad As Long
    Dim first As String, nLines As Long, nStorno As Long, liveQ() As Double, nNew() As Integer, isPart() As Boolean, anyLive() As Boolean
    Dim stype() As String, touched() As Boolean, maxN As Long, sd As Variant, k As Long, x As Double, nEI As Long, nParts As Long
    Dim nIdx As Long, keys() As String, keyEI() As Long, j As Long, nStale As Long, lastRow As Long, colA As Variant, colA0 As Long
    Dim hasEntry() As Boolean, st As String, key As String, entryKey() As String, need As Boolean, stNames As Variant
    Dim kNo As Integer, kEv As Integer, kTy As Integer, kEi As Integer, kMo As Integer, kQt As Integer, kSt As Integer, kRo As Integer
    Dim kId As Integer, sLive As String, sStorno As String, sNew As String, sAdd As String, plus As Variant, migrated As Boolean
    On Error GoTo EH
    t0 = GetSystemTicks()
    stNames = Split(SRC_STYPES, "|")
    spr = SprSheet()
    stk = WmsIssue.StockSheet()
    art = ArtSheet()
    nextSpl = CLng(SysNum(SK_NEXT_SPL))
    nextEI = CLng(SysNum(SK_NEXT_EI))
    For i = 0 To 4
        nextEv(i) = CLng(SysNum(EventCounterKey(i)))
    Next i
    ' a balance may exceed its live receipts by the live positive inventory corrections of the EI (Final Core)
    plus = WmsAdjust.PlusByEI(nextEI)
    ReDim liveQ(nextEI)
    ReDim nNew(nextEI)
    ReDim isPart(nextEI)
    ReDim anyLive(nextEI)
    ReDim stype(nextEI)
    ReDim touched(nextEI)
    ReDim hasEntry(nextEI)
    lastRow = WmsOrders.LastRow(SpecialSheet())
    colA0 = -1
    ' the columns and values used in the loop as locals: a Public Const of another module is looked up at every use (the loop
    ' over _SPR is about a third faster)
    kNo = SR_NO
    kEv = SR_EVENT
    kTy = SR_TYPE
    kEi = SR_EI
    kMo = SR_MODE
    kQt = SR_QTY
    kSt = SR_STATE
    kRo = SR_ROW
    kId = SR_IDENT
    sLive = RV_LIVE
    sStorno = RV_STORNO
    sNew = SR_NEW
    sAdd = SR_ADD
    ' 1. _SPR: every registered line
    last = WmsOrders.LastRow(spr)
    If last >= nextSpl Then Bad(nBad, first, SH_SPR & ": строк больше, чем выдано № (NEXT_SPL " & nextSpl & ")")
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        d = spr.getCellRangeByPosition(0, r0, SR_LAST, r1).getDataArray()
        For i = 0 To r1 - r0
            k = r0 + i
            If CStr(d(i)(kNo)) <> "" Then
                If VarType(d(i)(kNo)) <> 5 Then
                    Bad(nBad, first, SH_SPR & " строка " & (k + 1) & ": № не число")
                ElseIf d(i)(kNo) <> k Then
                    Bad(nBad, first, SH_SPR & " строка " & (k + 1) & ": записан № " & d(i)(kNo) & " (плотная адресация нарушена)")
                ElseIf Not CanonEvent(CStr(d(i)(kEv)), ti, num) Or CodeIndex(CStr(d(i)(kTy))) <> ti Then
                    Bad(nBad, first, "строка № " & k & ": событие «" & d(i)(kEv) & "» не согласовано с типом «" & d(i)(kTy) & "»")
                ElseIf num >= nextEv(ti) Then
                    Bad(nBad, first, "строка № " & k & ": событие " & d(i)(kEv) & " не меньше счётчика " & EventId(ti, nextEv(ti)))
                ElseIf Not WmsOrders.StrictEI(CStr(d(i)(kEi)), n) Or n >= nextEI Then
                    Bad(nBad, first, "строка № " & k & ": ЕИ «" & d(i)(kEi) & "» неверный или не меньше NEXT_EI")
                ElseIf VarType(d(i)(kQt)) <> 5 Then
                    Bad(nBad, first, "строка № " & k & ": количество не число")
                ElseIf d(i)(kQt) <= 0 Then
                    Bad(nBad, first, "строка № " & k & ": количество не больше 0")
                Else
                    nLines = nLines + 1
                    If num > maxEv(ti) Then maxEv(ti) = num
                    touched(n) = True
                    If CStr(d(i)(kMo)) = sNew Then
                        nNew(n) = nNew(n) + 1
                        If CStr(d(i)(kId)) <> "" Then stype(n) = CStr(d(i)(kId)) Else stype(n) = stNames(ti)
                        If ti = 2 Or CStr(d(i)(kId)) = stNames(2) Then isPart(n) = True
                    ElseIf CStr(d(i)(kMo)) <> sAdd Then
                        Bad(nBad, first, "строка № " & k & ": неизвестный режим «" & d(i)(kMo) & "»")
                    ElseIf ti <> 2 Then
                        Bad(nBad, first, "строка № " & k & ": пополнение существующего ЕИ не у детали")
                    End If
                    st = CStr(d(i)(kSt))
                    If st = sLive Then
                        liveQ(n) = liveQ(n) + d(i)(kQt)
                        anyLive(n) = True
                    ElseIf st = sStorno Then
                        nStorno = nStorno + 1
                    Else
                        Bad(nBad, first, "строка № " & k & ": неизвестное состояние «" & st & "»")
                    End If
                    ' the row hint: «Иной приход».A of that row read in chunks (a stale hint is not an error: rows inserted above)
                    x = -1
                    If VarType(d(i)(kRo)) = 5 Then x = d(i)(kRo)
                    If x < 1 Or x > lastRow Then
                        nStale = nStale + 1
                    Else
                        If CLng(x) < colA0 Or CLng(x) >= colA0 + CHECK_CHUNK_ROWS Or colA0 < 0 Then
                            colA0 = CLng(x)
                            colA = SpecialSheet().getCellRangeByPosition(XC_NO, colA0, XC_NO, IIf(colA0 + CHECK_CHUNK_ROWS - 1 > lastRow, lastRow, _
                                colA0 + CHECK_CHUNK_ROWS - 1)).getDataArray()
                        End If
                        If VarType(colA(CLng(x) - colA0)(0)) <> 5 Then
                            nStale = nStale + 1
                        ElseIf colA(CLng(x) - colA0)(0) <> k Then
                            nStale = nStale + 1
                        End If
                    End If
                End If
            End If
        Next i
    Next r0
    For i = 0 To 4
        If maxEv(i) >= nextEv(i) Then Bad(nBad, first, SrcCode(i) & ": счётчик " & nextEv(i) & " не больше максимального № события " & maxEv(i))
    Next i
    ' 2. the index of the parts: every entry, no key twice (a sorted copy of the keys), the key of each EI
    ReDim entryKey(nextEI)
    last = WmsOrders.LastRow(art)
    If last >= 1 Then
        d = art.getCellRangeByPosition(0, 1, AR_LAST, last).getDataArray()
        ReDim keys(last - 1)
        ReDim keyEI(last - 1)
        For i = 0 To last - 1
            key = CStr(d(i)(AR_KEY))
            If key = "" Then
                Bad(nBad, first, SH_ART & " строка " & (i + 2) & ": пустой артикул")
            ElseIf Not WmsOrders.StrictEI(CStr(d(i)(AR_EI)), n) Or n >= nextEI Then
                Bad(nBad, first, SH_ART & " строка " & (i + 2) & ": ЕИ «" & d(i)(AR_EI) & "» неверный")
            Else
                nIdx = nIdx + 1
                keys(i) = key
                keyEI(i) = n
                If hasEntry(n) And entryKey(n) <> key Then
                    Bad(nBad, first, "индекс артикулов: у " & WmsIssue.EiCanon(n) & " две записи («" & entryKey(n) & "», «" & key & "»)")
                End If
                hasEntry(n) = True
                entryKey(n) = key
            End If
        Next i
        SortKeys(keys, keyEI)
        For i = 1 To UBound(keys)
            If keys(i) <> "" And keys(i) = keys(i - 1) Then
                Bad(nBad, first, "Для артикула «" & keys(i) & "» найдено несколько ЕИ (" & WmsIssue.EiCanon(keyEI(i - 1)) & ", " _
                    & WmsIssue.EiCanon(keyEI(i)) & "). Требуется разбор")
            End If
        Next i
    End If
    ' 3. «Наличие» in chunks (only the chunks holding an EI of the special receipts or of the index): every such EI against
    ' its lines and its index entry
    For r0 = 1 To nextEI - 1 Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > nextEI - 1 Then r1 = nextEI - 1
        need = False
        For n = r0 To r1
            If touched(n) Or hasEntry(n) Then
                need = True
                Exit For
            End If
        Next n
        If need Then
            d = stk.getCellRangeByPosition(0, r0, SC_LAST, r1).getDataArray()
            For n = r0 To r1
                If touched(n) Or hasEntry(n) Then
                    sd = d(n - r0)
                    If CStr(sd(SC_EI)) <> WmsIssue.EiCanon(n) Then
                        Bad(nBad, first, WmsIssue.EiCanon(n) & ": нет строки в «" & SH_STOCK & "»")
                    Else
                        ' a migrated part (source «Перенос…», MIGRATE) existed before its refills: no line created it, and its balance
                        ' includes the transferred quantity (in the journal) — only the kind and the number are checked here
                        migrated = (Left(CStr(sd(SC_SRC)), 7) = "Перенос")
                        If touched(n) And migrated And nNew(n) = 0 Then
                            nEI = nEI + 1
                            nParts = nParts + 1
                            If CStr(sd(SC_STYPE)) <> stNames(2) Then
                                Bad(nBad, first, WmsIssue.EiCanon(n) & ": пополнение «Иного прихода» у перенесённого ЕИ не детали («" & sd(SC_STYPE) & "»)")
                            ElseIf VarType(sd(SC_QTY)) <> 5 Then
                                Bad(nBad, first, WmsIssue.EiCanon(n) & ": остаток не число")
                            ElseIf sd(SC_QTY) < 0 Then
                                Bad(nBad, first, WmsIssue.EiCanon(n) & ": отрицательный остаток")
                            End If
                        ElseIf touched(n) Then
                            If nNew(n) <> 1 Then
                                Bad(nBad, first, WmsIssue.EiCanon(n) & ": строк иного прихода, создавших ЕИ, " & nNew(n) & " (должна быть одна)")
                            Else
                                nEI = nEI + 1
                                If isPart(n) Then nParts = nParts + 1
                                If CStr(sd(SC_STYPE)) <> stype(n) Then
                                    Bad(nBad, first, WmsIssue.EiCanon(n) & ": тип источника «" & sd(SC_STYPE) & "», по строкам «" & stype(n) & "»")
                                ElseIf VarType(sd(SC_QTY)) <> 5 Then
                                    Bad(nBad, first, WmsIssue.EiCanon(n) & ": остаток не число")
                                ElseIf anyLive(n) Then
                                    If sd(SC_QTY) > liveQ(n) + plus(n) + 0.0005 Then
                                        Bad(nBad, first, WmsIssue.EiCanon(n) & ": остаток " & sd(SC_QTY) & " больше суммы действующих приходов " _
                                            & WmsIssue.QtyText(liveQ(n)) & IIf(plus(n) > 0, " и излишков инвентаризации " & WmsIssue.QtyText(plus(n)), ""))
                                    ElseIf CStr(sd(SC_STATE)) = EI_ST_STORNO Then
                                        Bad(nBad, first, WmsIssue.EiCanon(n) & ": есть действующие приходы, а состояние «" & EI_ST_STORNO & "»")
                                    End If
                                ElseIf sd(SC_QTY) <> 0 Or CStr(sd(SC_STATE)) <> EI_ST_STORNO Then
                                    Bad(nBad, first, WmsIssue.EiCanon(n) & ": все приходы удалены (сторно), а остаток " & sd(SC_QTY) & ", состояние «" _
                                        & sd(SC_STATE) & "»")
                                End If
                            End If
                        End If
                        If hasEntry(n) Then
                            If CStr(sd(SC_STYPE)) <> stNames(2) Then
                                Bad(nBad, first, "индекс артикулов: «" & entryKey(n) & "» → " & WmsIssue.EiCanon(n) & ", но это не ЕИ детали («" _
                                    & sd(SC_STYPE) & "»)")
                            ElseIf ArticleKey(CStr(sd(SC_ART))) <> entryKey(n) Then
                                Bad(nBad, first, "индекс артикулов: «" & entryKey(n) & "» → " & WmsIssue.EiCanon(n) & ", а у ЕИ артикул «" & sd(SC_ART) & "»")
                            End If
                        End If
                    End If
                End If
            Next n
        End If
    Next r0
    For n = 1 To nextEI - 1
        If isPart(n) And Not hasEntry(n) Then Bad(nBad, first, WmsIssue.EiCanon(n) & ": деталь не записана в индексе артикулов")
    Next n
    If nBad > 0 Then
        SpecialCheck = "ОШИБКА: расхождений " & nBad & " (" & first & IIf(nBad > 3, "; …", "") & ")"
    Else
        SpecialCheck = "строк " & nLines & " (из них сторно " & nStorno & "), ЕИ специальных приходов " & nEI & " (деталей " & nParts _
            & "), записей индекса артикулов " & nIdx & ", расхождений нет, счётчики выше всех номеров" _
            & IIf(nStale > 0, ", устаревших подсказок строк " & nStale & " (исправятся при следующей операции)", "") & ", " & (GetSystemTicks() - t0) & " мс"
    End If
    Exit Function
EH:
    SpecialCheck = "ОШИБКА: проверка специальных приходов не выполнена: " & Error$ & " (строка " & Erl & ")"
End Function

' a problem found by the self-check: counted; the first three are kept for the message
' the event ID exactly as WMS writes it ("PROD-00000125"): the source and the number (the self-check: a line of _SPR holds
' only this form); False for anything else
Private Function CanonEvent(s As String, ByRef ti As Integer, ByRef num As Long) As Boolean
    Dim p As Long, dg As String
    p = InStr(s, "-")
    If p < 2 Or Len(s) - p <> EVENT_DIGITS Then Exit Function
    ti = CodeIndex(Left(s, p - 1))
    If ti < 0 Then Exit Function
    dg = Mid(s, p + 1)
    If Not IsDigits(dg) Then Exit Function
    num = CLng(dg)
    CanonEvent = (num >= 1)
End Function

' the source of a code of _SPR (OFF, PROD, DET, OLD, OTH); -1 otherwise
Private Function CodeIndex(c As String) As Integer
    Select Case c
    Case "OFF"
        CodeIndex = 0
    Case "PROD"
        CodeIndex = 1
    Case "DET"
        CodeIndex = 2
    Case "OLD"
        CodeIndex = 3
    Case "OTH"
        CodeIndex = 4
    Case Else
        CodeIndex = -1
    End Select
End Function

Private Sub Bad(ByRef nBad As Long, ByRef first As String, s As String)
    nBad = nBad + 1
    If nBad <= 3 Then
        If first = "" Then first = s Else first = first & "; " & s
    End If
End Sub

' keys sorted with their EIs (Shell sort in memory: the index is small, one row per article)
Private Sub SortKeys(keys() As String, ei() As Long)
    Dim gap As Long, i As Long, j As Long, tk As String, te As Long, n As Long
    n = UBound(keys) + 1
    gap = n \ 2
    Do While gap > 0
        For i = gap To n - 1
            tk = keys(i)
            te = ei(i)
            j = i
            Do While j >= gap
                If keys(j - gap) <= tk Then Exit Do
                keys(j) = keys(j - gap)
                ei(j) = ei(j - gap)
                j = j - gap
            Loop
            keys(j) = tk
            ei(j) = te
        Next i
        gap = gap \ 2
    Loop
End Sub

' ================================================================ the index of the parts after a migration (spec §31)
' Rebuilds _ART from «Наличие»: every EI whose kind of source is «Детали» gets one entry «its article key → EI». Old
' conflicts (one article at several EIs) are written as they are — every lookup of such an article then fails closed
' with «найдено несколько ЕИ» — and are listed in the result; nothing is merged or renumbered. Maintenance of a migrated
' book: allowed only when WMS is ready (CLEAN), without a journal record (the index is derived from «Наличие» only).
Function ArticleIndexRebuild() As String
    Dim stk As Object, art As Object, last As Long, r0 As Long, r1 As Long, i As Long, dT As Variant, dA As Variant, out() As Variant
    Dim nOut As Long, keys() As String, eis() As Long, conf As String, nConf As Long, wasProt As Boolean, why As String, t0 As Long
    Dim um As Object
    WmsInit()
    why = PostingBlockReason()
    If why <> "" Then
        ArticleIndexRebuild = "ОШИБКА: " & Mid(why, InStr(why, ":") + 1)
        Exit Function
    End If
    On Error GoTo EH
    t0 = GetSystemTicks()
    stk = WmsIssue.StockSheet()
    art = ArtSheet()
    last = WmsOrders.LastRow(stk)
    ReDim out(IIf(last > 0, last, 1))
    For r0 = 1 To last Step CHECK_CHUNK_ROWS
        r1 = r0 + CHECK_CHUNK_ROWS - 1
        If r1 > last Then r1 = last
        dT = stk.getCellRangeByPosition(SC_STYPE, r0, SC_STYPE, r1).getDataArray()
        dA = stk.getCellRangeByPosition(SC_ART, r0, SC_ART, r1).getDataArray()
        For i = 0 To r1 - r0
            If CStr(dT(i)(0)) = SrcStype(2) And ArticleKey(CStr(dA(i)(0))) <> "" Then
                out(nOut) = Array(ArticleKey(CStr(dA(i)(0))), WmsIssue.EiCanon(r0 + i), CStr(dA(i)(0)))
                nOut = nOut + 1
            End If
        Next i
    Next r0
    um = gDoc.getUndoManager()
    um.lock()
    wasProt = art.isProtected()
    If wasProt Then art.unprotect(PROTECT_PWD)
    art.getCellRangeByPosition(0, 1, AR_LAST, MAX_SHEET_ROW).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.STRING)
    If nOut > 0 Then
        ReDim Preserve out(nOut - 1)
        art.getCellRangeByPosition(0, 1, AR_LAST, nOut).setDataArray(out)
    End If
    If wasProt Then art.protect(PROTECT_PWD)
    um.unlock()
    If nOut > 1 Then
        ReDim keys(nOut - 1)
        ReDim eis(nOut - 1)
        For i = 0 To nOut - 1
            keys(i) = out(i)(0)
            eis(i) = CLng(Mid(out(i)(1), 4))
        Next i
        SortKeys(keys, eis)
        For i = 1 To nOut - 1
            If keys(i) = keys(i - 1) Then
                nConf = nConf + 1
                If Len(conf) < 400 Then conf = conf & IIf(conf <> "", "; ", "") & keys(i) & ": " & WmsIssue.EiCanon(eis(i - 1)) & ", " & WmsIssue.EiCanon(eis(i))
            End If
        Next i
    End If
    ArticleIndexRebuild = "индекс артикулов перестроен: записей " & nOut & ", конфликтов «один артикул — несколько ЕИ» " & nConf _
        & IIf(nConf > 0, " (такие артикулы не проводятся до разбора: " & conf & ")", "") & ", " & (GetSystemTicks() - t0) & " мс"
    Exit Function
EH:
    ArticleIndexRebuild = "ОШИБКА: индекс артикулов не перестроен: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    If wasProt Then art.protect(PROTECT_PWD)
    um.unlock()
End Function
