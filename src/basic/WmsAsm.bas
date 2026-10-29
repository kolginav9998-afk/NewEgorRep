' WmsAsm — 1.2.0: «Сборка деталей». A part that the storekeeper puts together from other parts of the warehouse (a headlight from its parts): the sheet «Сборка деталей» is his own
' directory — one block of rows per part: its name and permanent code (article) and, row by row, what goes into ONE piece (name, article or the number of an EI, quantity, unit). Once
' the part is described, «Собрать деталь…» is one button: the parts that go in are taken from their EIs (the oldest first, the EIs are shown before the posting), the assembled
' pieces go to the PERMANENT EI of the part — the EI of its article (the index _ART): the first build makes it, every next build and every receipt of the finished part («Иной приход»,
' type «Детали») adds to it. The first time the part can be put together without a description: the EIs selected on «Наличие» and «Собрать деталь из ЕИ…» — the description is
' written on the sheet by itself, the next time it is the button only. Serial numbers of the pieces are optional (the same as at the receipt: WmsSerial). One operation of the
' journal for the whole build (ASM_BUILD: the parts leave their EIs as corrections «Сборка», the part comes as one more correction «Сборка»; _ADJ kind KIT, the fields AJ_BOOK and
' AJ_FACT hold the balances before and after — they mark a build of a part), the storno is ASM_DEL («Удалить» on any row of the build). Library WmsPlus.
Option Explicit

Private Const AS_SHEET = "Сборка деталей"
Private Const AS_PART = 0
Private Const AS_ART = 1
Private Const AS_NAME = 2
Private Const AS_CART = 3
Private Const AS_QTY = 4
Private Const AS_UNIT = 5
Private Const AS_NOTE = 6
Private Const AS_CAT = 7
Private Const AS_PUNIT = 8
Private Const AS_SER = 9
Private Const AS_LAST = 9
Private Const AS_MAX = 200
Private Const AS_FROM_EIS_MAX = 12
Private Const AS_UNIT_DEFAULT = "шт"
Private Const AS_CAT_DEFAULT = "Детали"
Private Const AS_COMPS_TEXT = 160

' ================================================================ the directory «Сборка деталей»

' the names of the parts in their order on the sheet (a row with an empty A continues the part above)
Function AsmNames() As Variant
    Dim sh As Object, last As Long, d As Variant, i As Long, cur As String, out() As String, n As Long, j As Long, seen As Boolean
    AsmNames = Array()
    If Not gDoc.Sheets.hasByName(AS_SHEET) Then Exit Function
    sh = gDoc.Sheets.getByName(AS_SHEET)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, AS_LAST, last).getDataArray()
    ReDim out(UBound(d))
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(AS_PART))) <> "" Then cur = Trim(CStr(d(i)(AS_PART)))
        If cur <> "" And Trim(CStr(d(i)(AS_NAME))) <> "" Then
            seen = False
            For j = 0 To n - 1
                If LCase(out(j)) = LCase(cur) Then
                    seen = True
                    Exit For
                End If
            Next j
            If Not seen Then
                out(n) = cur
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve out(n - 1)
    AsmNames = out
End Function

' one attribute of the part that stands in the first row of its block (article, category, unit, serial numbers): the first filled cell counts, a different value in a later row is a mistake
Private Function AsmAttr(sCur As String, ByVal v As String, r As Long, sWhat As String, ByRef why As String) As String
    v = Trim(v)
    AsmAttr = sCur
    If v = "" Then Exit Function
    If sCur = "" Then
        AsmAttr = v
    ElseIf WmsKit.KtKey(v) <> WmsKit.KtKey(sCur) Then
        If why = "" Then why = "«" & AS_SHEET & "» строка " & r & ": " & sWhat & " «" & v & "», а выше для этой детали «" & sCur & "» — оставьте одно значение"
    End If
End Function

' the description of the part sName into the globals of WmsKit (gKt*: what goes into one piece) and sArt / sUnit / sCat / bSer (what the part is). "" or the problem: no sheet, no part,
' a quantity that is not a number, a part without article or with two, a part that goes into itself
Function AsmRead(sName As String, ByRef sArt As String, ByRef sUnit As String, ByRef sCat As String, ByRef bSer As Boolean) As String
    Dim sh As Object, last As Long, d As Variant, i As Long, cur As String, q As Double, msg As String, why As String, sSer As String, key As String, ck As String
    sArt = ""
    sUnit = ""
    sCat = ""
    sSer = ""
    bSer = False
    gKtN = 0
    ReDim gKtName(AS_MAX)
    ReDim gKtArt(AS_MAX)
    ReDim gKtUnit(AS_MAX)
    ReDim gKtQty(AS_MAX)
    ReDim gKtEi(AS_MAX)
    If Not gDoc.Sheets.hasByName(AS_SHEET) Then
        AsmRead = "листа «" & AS_SHEET & "» нет"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(AS_SHEET)
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        d = sh.getCellRangeByPosition(0, 1, AS_LAST, last).getDataArray()
        For i = 0 To UBound(d)
            If Trim(CStr(d(i)(AS_PART))) <> "" Then cur = Trim(CStr(d(i)(AS_PART)))
            If LCase(cur) = LCase(Trim(sName)) Then
                sArt = AsmAttr(sArt, CStr(d(i)(AS_ART)), i + 2, "артикул детали", why)
                sCat = AsmAttr(sCat, CStr(d(i)(AS_CAT)), i + 2, "категория", why)
                sUnit = AsmAttr(sUnit, CStr(d(i)(AS_PUNIT)), i + 2, "единица детали", why)
                sSer = AsmAttr(sSer, CStr(d(i)(AS_SER)), i + 2, "серийные номера", why)
                If why <> "" Then
                    AsmRead = why
                    Exit Function
                End If
                If Trim(CStr(d(i)(AS_NAME))) <> "" Then
                    If VarType(d(i)(AS_QTY)) = 5 Then
                        q = d(i)(AS_QTY)
                    ElseIf Not WmsOrders.ParseQtyValueAs(d(i)(AS_QTY), "E", q, msg) Then
                        AsmRead = "«" & AS_SHEET & "» строка " & (i + 2) & ": количество «" & d(i)(AS_QTY) & "» — " & msg
                        Exit Function
                    End If
                    If q <= 0 Then
                        AsmRead = "«" & AS_SHEET & "» строка " & (i + 2) & ": количество на 1 шт должно быть больше нуля"
                        Exit Function
                    End If
                    If gKtN >= AS_MAX Then
                        AsmRead = "в описании детали больше " & AS_MAX & " позиций"
                        Exit Function
                    End If
                    gKtName(gKtN) = Trim(CStr(d(i)(AS_NAME)))
                    gKtArt(gKtN) = Trim(CStr(d(i)(AS_CART)))
                    gKtUnit(gKtN) = Trim(CStr(d(i)(AS_UNIT)))
                    gKtQty(gKtN) = q
                    gKtEi(gKtN) = WmsKit.KtEiRef(gKtArt(gKtN))
                    If gKtEi(gKtN) = 0 And gKtArt(gKtN) = "" Then gKtEi(gKtN) = WmsKit.KtEiRef(gKtName(gKtN))
                    If gKtEi(gKtN) < 0 Then
                        AsmRead = "«" & AS_SHEET & "» строка " & (i + 2) & ": «" & IIf(gKtArt(gKtN) <> "", gKtArt(gKtN), gKtName(gKtN)) & "» — такого ЕИ нет"
                        Exit Function
                    End If
                    gKtN = gKtN + 1
                End If
            End If
        Next i
    End If
    If gKtN = 0 Then
        AsmRead = "детали «" & sName & "» нет на листе «" & AS_SHEET & "» (или в её описании нет ни одной позиции)"
        Exit Function
    End If
    If sArt = "" Then
        AsmRead = "у детали «" & sName & "» на листе «" & AS_SHEET & "» не указан артикул (колонка B) — это её постоянный код: по нему WMS находит ЕИ детали"
        Exit Function
    End If
    key = WmsSpecial.ArticleKey(sArt)
    If key = "" Then
        AsmRead = "артикул «" & sArt & "» детали «" & sName & "» пуст после нормализации — впишите постоянный код"
        Exit Function
    End If
    For i = 0 To gKtN - 1
        ck = ""
        If gKtEi(i) = 0 And gKtArt(i) <> "" Then ck = WmsSpecial.ArticleKey(gKtArt(i))
        If ck = key Then
            AsmRead = "деталь «" & sName & "» не может входить сама в себя: позиция «" & gKtName(i) & "» с тем же артикулом " & sArt
            Exit Function
        End If
        If gKtEi(i) = 0 And gKtArt(i) = "" And WmsKit.KtKey(gKtName(i)) = WmsKit.KtKey(sName) Then
            AsmRead = "деталь «" & sName & "» не может входить сама в себя: позиция с тем же названием"
            Exit Function
        End If
    Next i
    If sUnit = "" Then sUnit = AS_UNIT_DEFAULT
    If sCat = "" Then sCat = AS_CAT_DEFAULT
    sSer = LCase(Trim(sSer))
    bSer = (Left(sSer, 1) = "д" Or sSer = "1" Or sSer = "+")
End Function

Sub NavAsm(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(AS_SHEET, 1, 0)
End Sub

' ================================================================ how many can be built now

' how many pieces of the part sName the stock allows (the parts that go in, picked as the build would pick them); sLimit — what limits it («Болт: на 1 шт нужно 4, есть 10») or,
' when the description is wrong, the problem (the result is -1)
Function AsmCapacity(sName As String, ByRef sLimit As String) As Long
    Dim art As String, unit As String, cat As String, ser As Boolean, why As String, j As Long, tot As Double, c As Long, nMax As Long
    sLimit = ""
    why = AsmRead(sName, art, unit, cat, ser)
    If why <> "" Then
        sLimit = why
        AsmCapacity = -1
        Exit Function
    End If
    nMax = -1
    For j = 0 To gKtN - 1
        gKpN = 0
        ReDim gKpEi(31)
        ReDim gKpQty(31)
        ReDim gKpPart(31)
        tot = WmsKit.KitPick(j, 1000000000#)
        c = CLng(Int(tot / gKtQty(j) + 0.0000001))
        If nMax < 0 Or c < nMax Then
            nMax = c
            sLimit = gKtName(j) & IIf(gKtArt(j) <> "", " (" & gKtArt(j) & ")", "") & ": на 1 шт нужно " & WmsIssue.QtyText(gKtQty(j)) & IIf(gKtUnit(j) <> "", " " & gKtUnit(j), "") _
                & ", есть " & WmsIssue.QtyText(tot)
        End If
    Next j
    ' two positions may take from the same EI: the build is tried with the number found
    Do While nMax > 0
        If WmsKit.KitPickAll(nMax) = "" Then Exit Do
        nMax = nMax - 1
        sLimit = "позиции берут из одних и тех же ЕИ"
    Loop
    AsmCapacity = nMax
End Function

' ================================================================ the build (ASM_BUILD)

' the place of the permanent EI of the article (the index _ART), "" when the part has no EI yet
Function AsmPlaceOf(sArt As String) As String
    Dim key As String, n As Long, canon As String, aRow As Long, why As String
    key = WmsSpecial.ArticleKey(sArt)
    If key = "" Then Exit Function
    If WmsSpecial.ArtLookup(key, n, canon, aRow, why) = 1 Then AsmPlaceOf = CStr(WmsIssue.StockData(n)(SC_PLACE))
End Function

' nQty pieces of the part sName (described on «Сборка деталей») built on day dt. OK:<EI>|NEW or ADD|<balance before>|<after>|<text> / ERR:… / ERR-SYS:… / SKIP:… / BLOCKED:…
Function AsmBuild(sName As String, nQty As Double, sPlace As String, dt As Double, sReason As String) As String
    Dim art As String, unit As String, cat As String, ser As Boolean, why As String
    WmsInit()
    why = AsmRead(sName, art, unit, cat, ser)
    If why <> "" Then
        AsmBuild = "ERR:" & why
        gUiLastMsg = AsmBuild
        Exit Function
    End If
    AsmBuild = AsmBuildCore(sName, art, unit, cat, nQty, sPlace, dt, sReason)
End Function

' the build itself: the components are in gKt* (AsmRead or AsmFromEis), the part is described by sName / sArt / sUnit / sCat
Private Function AsmBuildCore(sName As String, sArt As String, sUnit As String, sCat As String, nQty As Double, sPlace As String, dt As Double, sReason As String) As String
    Dim why As String, sShort As String, key As String, a As Integer, en As Long, ecanon As String, aRow As Long, bAdd As Boolean, m As Long, canonK As String, sd As Variant
    Dim b0 As Double, b1 As Double, nAdj0 As Long, i As Long, n As Long, r As Long, r0 As Long, st As Object, d As Variant, ash As Object, comps As String, txt As String
    Dim sBatch As String, res As String, cb0 As Double, cb1 As Double, sNote As String, sPl As String, sCtl As String, nm As String, pu As String, pn As Double
    WmsInit()
    On Error GoTo EH
    If nQty <= 0 Or nQty <> Int(nQty) Or nQty > 100000 Then
        AsmBuildCore = "ERR:сколько собрать — целое число больше нуля"
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    If sReason = "" Then sReason = "сборка детали"
    key = WmsSpecial.ArticleKey(sArt)
    If key = "" Then
        AsmBuildCore = "ERR:у детали «" & sName & "» нет артикула — это её постоянный код"
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    ' the permanent EI of the article
    a = WmsSpecial.ArtLookup(key, en, ecanon, aRow, why)
    If a >= 2 Then
        AsmBuildCore = "ERR:" & why
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    bAdd = (a = 1)
    If bAdd Then
        sd = WmsIssue.StockData(en)
        If VarType(sd(SC_QTY)) <> 5 Then
            AsmBuildCore = "ERR-SYS:остаток " & ecanon & " в «" & SH_STOCK & "» не число — нужна самопроверка"
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
        If CStr(sd(SC_STATE)) <> EI_ST_ACTIVE And CStr(sd(SC_STATE)) <> EI_ST_STORNO Then
            AsmBuildCore = "ERR:" & ecanon & " в состоянии «" & sd(SC_STATE) & "» — собранные детали в него не добавляются"
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
        If LCase(Replace(Trim(CStr(sd(SC_UNIT))), ".", "")) <> LCase(Replace(Trim(sUnit), ".", "")) Then
            AsmBuildCore = "ERR:" & ecanon & " ведётся в «" & sd(SC_UNIT) & "», а деталь собирается в «" & sUnit & "» — пересчёта нет; поправьте единицу на листе «" & AS_SHEET & "»"
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
        m = en
        canonK = ecanon
        b0 = sd(SC_QTY)
        sPl = CStr(sd(SC_PLACE))
        nm = CStr(sd(SC_NAME))
    Else
        If Trim(sPlace) = "" Then
            AsmBuildCore = "ERR:укажите место для нового ЕИ детали «" & sName & "» — куда положить собранные детали"
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
        m = CLng(SysNum(SK_NEXT_EI))
        why = WmsOrders.NewEIProblem(m)
        If why <> "" Then
            AsmBuildCore = "ERR-SYS:" & why
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
        canonK = WmsIssue.EiCanon(m)
        sPl = Trim(sPlace)
        nm = sName
        aRow = WmsOrders.LastRow(WmsSpecial.ArtSheet()) + 1
    End If
    b1 = WmsIssue.Round3(b0 + nQty)
    sShort = WmsKit.KitPickAll(nQty)
    If sShort <> "" Then
        AsmBuildCore = "ERR:деталь «" & sName & "» ×" & WmsIssue.QtyText(nQty) & " не собрать — не хватает: " & sShort
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    For i = 0 To gKpN - 1
        If bAdd And gKpEi(i) = m Then
            AsmBuildCore = "ERR:в составе детали оказался её собственный ЕИ " & canonK & " — проверьте описание на листе «" & AS_SHEET & "»"
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
    Next i
    why = PostingBlockReason()
    If why <> "" Then
        AsmBuildCore = why
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    nAdj0 = CLng(SysNum(SK_NEXT_ADJ))
    For i = 0 To gKpN
        why = WmsAdjust.NewAdjProblem(nAdj0 + i)
        If why <> "" Then
            AsmBuildCore = "ERR-SYS:" & why
            gUiLastMsg = AsmBuildCore
            Exit Function
        End If
    Next i
    st = WmsIssue.StockSheet()
    For i = 0 To gKpN - 1
        d = st.getCellRangeByPosition(0, gKpEi(i), SC_LAST, gKpEi(i)).getDataArray()(0)
        txt = txt & Chr(10) & "  " & d(SC_EI) & " — " & d(SC_NAME) & ": " & WmsPack.QtyWithParts(gKpEi(i), gKpQty(i), CStr(d(SC_UNIT)))
    Next i
    If Not WmsUi.UiConfirm("Собрать деталь «" & nm & "» (арт. " & sArt & ") ×" & WmsIssue.QtyText(nQty) & "?" & Chr(10) & Chr(10) & "Будет снято:" & txt & Chr(10) & Chr(10) _
        & IIf(bAdd, "Соберётся в постоянный " & canonK & ": было " & WmsIssue.QtyText(b0) & ", станет " & WmsIssue.QtyText(b1) & " " & sUnit & " (место " & sPl & ")", _
        "Деталь появится на складе новым ЕИ " & canonK & " (место «" & sPl & "») — это её постоянный ЕИ: следующие сборки и приходы готовой детали идут в него")) Then
        AsmBuildCore = "SKIP:сборка не подтверждена"
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    ash = WmsAdjust.AdjustSheet()
    r0 = WmsOrders.LastRow(ash) + 1
    sBatch = "Сборка детали «" & sName & "» № " & nAdj0
    PlanBegin("ASM_BUILD", CStr(nAdj0))
    For i = 0 To gKpN - 1
        n = nAdj0 + i
        r = r0 + i
        d = st.getCellRangeByPosition(0, gKpEi(i), SC_LAST, gKpEi(i)).getDataArray()(0)
        cb0 = d(SC_QTY)
        cb1 = WmsIssue.Round3(cb0 - gKpQty(i))
        comps = comps & IIf(comps <> "", ";", "") & d(SC_EI) & ":" & NumStr(gKpQty(i)) & ":" & NumStr(cb0)
        sNote = sNote & IIf(sNote <> "", ", ", "") & d(SC_EI) & " ×" & WmsIssue.QtyText(gKpQty(i))
        PlanSetValue(SH_STOCK, gKpEi(i), SC_QTY, cb1, False)
        WmsKit.KtAdjRow(n, CStr(d(SC_EI)), -gKpQty(i), r, CStr(d(SC_PLACE)), "", dt, canonK)
        PlanSetValue(SH_ADJ, n, AJ_BOOK, cb0, False)
        PlanSetValue(SH_ADJ, n, AJ_FACT, cb1, False)
        WmsKit.KtSheetRow(r, n, CStr(d(SC_EI)), CStr(d(SC_NAME)), CStr(d(SC_ART)), CStr(d(SC_UNIT)), gKpQty(i), "", dt, sReason, CStr(d(SC_PLACE)), cb0, cb1, _
            ST_POSTED & " · в сборку " & canonK & " (" & nm & ")", sBatch)
        WmsOrders.PlanStockMirror(gKpEi(i), cb1)
    Next i
    ' the part: its permanent EI (made now or added to) and its own row «Сборка» (+N)
    n = nAdj0 + gKpN
    r = r0 + gKpN
    If bAdd Then
        PlanSetValue(SH_STOCK, m, SC_QTY, b1, False)
        If CStr(sd(SC_STATE)) <> EI_ST_ACTIVE Then PlanSetValue(SH_STOCK, m, SC_STATE, EI_ST_ACTIVE, False)
    Else
        WmsSpecial.PlanNewStockRow(m, canonK, nQty, sName, sArt, sUnit, sPl, sCat, EI_ST_ACTIVE, "Сборка: корректировки № " & nAdj0 & "–" & n, WmsSpecial.SrcStype(2))
        PlanSetValue(SH_ART, aRow, AR_KEY, key, False)
        PlanSetValue(SH_ART, aRow, AR_EI, canonK, False)
        PlanSetValue(SH_ART, aRow, AR_ART, sArt, False)
        PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, m + 1, False)
    End If
    WmsOrders.PlanStockMirror(m, b1)
    WmsKit.KtAdjRow(n, canonK, nQty, r, "", sPl, dt, canonK)
    If bAdd Then PlanSetValue(SH_ADJ, n, AJ_BOOK, b0, False)
    PlanSetValue(SH_ADJ, n, AJ_FACT, b1, False)
    sCtl = Left(sNote, AS_COMPS_TEXT)
    If Len(sNote) > AS_COMPS_TEXT Then sCtl = sCtl & "…"
    WmsKit.KtSheetRow(r, n, canonK, nm, sArt, sUnit, nQty, sPl, dt, sReason, "", b0, b1, ST_POSTED & " · деталь из " & gKpN & " поз.: " & sCtl, sBatch)
    PlanSetValue(SYS_SHEET, SK_NEXT_ADJ, 1, n + 1, False)
    PlanField("MODE", IIf(bAdd, "ADD", "NEW"))
    PlanField("NAME", nm)
    PlanField("ART", sArt)
    PlanField("ART_KEY", key)
    PlanField("UNIT", sUnit)
    PlanField("CAT", sCat)
    PlanField("N", nQty)
    PlanField("NEW_EI", canonK)
    PlanField("ADJ", nAdj0)
    PlanField("ADJ_KIT", n)
    PlanField("COMPS", comps)
    PlanField("PLACE", sPl)
    PlanField("DATE", Format(dt, "YYYY-MM-DD"))
    PlanField("REASON", sReason)
    PlanField("BATCH", sBatch)
    PlanField("BAL_BEFORE", b0)
    PlanField("BAL_AFTER", b1)
    res = ApplyOperation(0, 0)
    If Left(res, 3) <> "OK:" Then
        AsmBuildCore = res
        gUiLastMsg = AsmBuildCore
        Exit Function
    End If
    AsmBuildCore = "OK:" & canonK & "|" & IIf(bAdd, "ADD", "NEW") & "|" & NumStr(b0) & "|" & NumStr(b1) & "|собрана «" & nm & "» ×" & WmsIssue.QtyText(nQty) & " из " & gKpN & " ЕИ: " _
        & IIf(bAdd, "добавлено в постоянный " & canonK & " (было " & WmsIssue.QtyText(b0) & ", стало " & WmsIssue.QtyText(b1) & ")", "новый постоянный " & canonK & " (остаток " & WmsIssue.QtyText(b1) & ")") _
        & "; корректировки № " & nAdj0 & "–" & n & " на листе «" & SH_ADJUST & "»"
    gUiLastMsg = AsmBuildCore
    Exit Function
EH:
    AsmBuildCore = "ERR-SYS:внутренняя ошибка сборки детали: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = AsmBuildCore
End Function

' ================================================================ the first time: from the EIs selected on «Наличие»

' nQty pieces of the part sName / sArt put together from the EIs eis (numbers), q(i) of EI eis(i) for ONE piece (a text: «2», «0,5», «1 кан»): the same build as AsmBuild with this
' composition; bSave — the composition written on «Сборка деталей» (the parts by their articles: the next time it is the button only). OK:… / ERR:… / SKIP:…
Function AsmFromEis(eis As Variant, q As Variant, sName As String, sArt As String, nQty As Double, sPlace As String, dt As Double, sReason As String, bSave As Boolean) As String
    Dim i As Long, j As Long, n As Long, st As Object, d As Variant, x As Double, msg As String, res As String, names As Variant, b As Double, info As String, key As String, ck As String
    WmsInit()
    On Error GoTo EH
    If Trim(sName) = "" Then
        AsmFromEis = "ERR:дайте детали название — так она будет называться в «" & SH_STOCK & "»"
        gUiLastMsg = AsmFromEis
        Exit Function
    End If
    key = WmsSpecial.ArticleKey(sArt)
    If key = "" Then
        AsmFromEis = "ERR:впишите артикул детали — её постоянный код: по нему WMS находит ЕИ детали (и пополняет его при следующих сборках и приходах)"
        gUiLastMsg = AsmFromEis
        Exit Function
    End If
    If UBound(eis) < 0 Or UBound(eis) >= AS_FROM_EIS_MAX Or UBound(q) <> UBound(eis) Then
        AsmFromEis = "ERR:ЕИ для сборки — от 1 до " & AS_FROM_EIS_MAX
        gUiLastMsg = AsmFromEis
        Exit Function
    End If
    If bSave Then
        names = AsmNames()
        For i = 0 To UBound(names)
            If LCase(names(i)) = LCase(Trim(sName)) Then
                AsmFromEis = "ERR:деталь «" & sName & "» уже описана на листе «" & AS_SHEET & "» — соберите её кнопкой «Собрать деталь…» или ответьте «нет» на «Записать состав»"
                gUiLastMsg = AsmFromEis
                Exit Function
            End If
        Next i
    End If
    st = WmsIssue.StockSheet()
    gKtN = 0
    ReDim gKtName(AS_MAX)
    ReDim gKtArt(AS_MAX)
    ReDim gKtUnit(AS_MAX)
    ReDim gKtQty(AS_MAX)
    ReDim gKtEi(AS_MAX)
    For i = 0 To UBound(eis)
        n = eis(i)
        If n >= 1 Then d = st.getCellRangeByPosition(0, n, SC_STATE, n).getDataArray()(0)
        If n < 1 Then
            AsmFromEis = "ERR:ЕИ-" & n & " нет в «" & SH_STOCK & "»"
            gUiLastMsg = AsmFromEis
            Exit Function
        ElseIf CStr(d(SC_EI)) <> WmsIssue.EiCanon(n) Then
            AsmFromEis = "ERR:ЕИ-" & n & " нет в «" & SH_STOCK & "»"
            gUiLastMsg = AsmFromEis
            Exit Function
        End If
        For j = 0 To i - 1
            If eis(j) = n Then
                AsmFromEis = "ERR:" & d(SC_EI) & " выбран дважды"
                gUiLastMsg = AsmFromEis
                Exit Function
            End If
        Next j
        If Trim(CStr(d(SC_ART))) <> "" Then
            If WmsSpecial.ArticleKey(CStr(d(SC_ART))) = key Then
                AsmFromEis = "ERR:" & d(SC_EI) & " — это сама деталь «" & sName & "» (тот же артикул): деталь не может входить сама в себя"
                gUiLastMsg = AsmFromEis
                Exit Function
            End If
        End If
        If VarType(d(SC_QTY)) = 5 Then b = d(SC_QTY) Else b = 0
        If WmsPack.QtyEIText(CStr(q(i)), n, b, CStr(d(SC_UNIT)), "OUT", 0, x, info, msg) <> 0 Then
            AsmFromEis = "ERR:" & d(SC_EI) & ": на 1 шт «" & q(i) & "» — " & msg
            gUiLastMsg = AsmFromEis
            Exit Function
        End If
        If x <= 0 Then
            AsmFromEis = "ERR:" & d(SC_EI) & ": на 1 шт — больше нуля"
            gUiLastMsg = AsmFromEis
            Exit Function
        End If
        gKtName(i) = CStr(d(SC_NAME))
        gKtArt(i) = CStr(d(SC_EI))
        gKtUnit(i) = CStr(d(SC_UNIT))
        gKtQty(i) = x
        gKtEi(i) = n
        gKtN = gKtN + 1
    Next i
    res = AsmBuildCore(Trim(sName), Trim(sArt), AS_UNIT_DEFAULT, AS_CAT_DEFAULT, nQty, sPlace, dt, sReason)
    If Left(res, 3) = "OK:" And bSave Then res = res & AsmSaveComposition(Trim(sName), Trim(sArt), dt)
    AsmFromEis = res
    gUiLastMsg = AsmFromEis
    Exit Function
EH:
    AsmFromEis = "ERR-SYS:внутренняя ошибка сборки из ЕИ: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = AsmFromEis
End Function

' the composition in gKt* (the parts are the EIs of the build) written on «Сборка деталей»: a part with an article — by the article (the next build takes it from any of its EIs, the oldest
' first), a part without — by the name and the unit; two EIs of one article are one line. The text for the message: "; состав записан …" or why it is not
Private Function AsmSaveComposition(sName As String, sArt As String, dt As Double) As String
    Dim ks As Object, r As Long, out() As Variant, i As Long, j As Long, n As Long, st As Object, d As Variant, k As Long, ak As String, merged As Boolean
    Dim cn() As String, ca() As String, cu() As String, cq() As Double
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(AS_SHEET) Then
        AsmSaveComposition = "; состав не записан: листа «" & AS_SHEET & "» нет"
        Exit Function
    End If
    ks = gDoc.Sheets.getByName(AS_SHEET)
    st = WmsIssue.StockSheet()
    ReDim cn(gKtN)
    ReDim ca(gKtN)
    ReDim cu(gKtN)
    ReDim cq(gKtN)
    For i = 0 To gKtN - 1
        d = st.getCellRangeByPosition(0, gKtEi(i), SC_LAST, gKtEi(i)).getDataArray()(0)
        ak = ""
        If Trim(CStr(d(SC_ART))) <> "" Then ak = Trim(CStr(d(SC_ART)))
        merged = False
        If ak <> "" Then
            For j = 0 To k - 1
                If ca(j) <> "" Then
                    If WmsSpecial.ArticleKey(ca(j)) = WmsSpecial.ArticleKey(ak) And LCase(Replace(cu(j), ".", "")) = LCase(Replace(CStr(d(SC_UNIT)), ".", "")) Then
                        cq(j) = WmsIssue.Round3(cq(j) + gKtQty(i))
                        merged = True
                        Exit For
                    End If
                End If
            Next j
        End If
        If Not merged Then
            cn(k) = CStr(d(SC_NAME))
            ca(k) = ak
            cu(k) = CStr(d(SC_UNIT))
            cq(k) = gKtQty(i)
            k = k + 1
        End If
    Next i
    r = WmsOrders.LastRow(ks) + 1
    If r < 1 Then r = 1
    ReDim out(k - 1)
    For i = 0 To k - 1
        out(i) = Array(IIf(i = 0, sName, ""), IIf(i = 0, sArt, ""), cn(i), ca(i), cq(i), cu(i), IIf(i = 0, "собрана из ЕИ " & Format(dt, "DD.MM.YYYY"), ""), "", "", "")
    Next i
    ks.getCellRangeByPosition(0, r, AS_LAST, r + k - 1).setDataArray(out())
    AsmSaveComposition = "; состав записан на лист «" & AS_SHEET & "» (строки " & (r + 1) & "–" & (r + k) & ") — в следующий раз достаточно «Собрать деталь…»"
    Exit Function
EH:
    AsmSaveComposition = "; состав на лист «" & AS_SHEET & "» не записан: " & Error$
End Function

' ================================================================ the storno of a build (ASM_DEL)

' true when the record rd of _ADJ is a row of a build of a part (the records of a build of a part carry the balances before and after in AJ_BOOK / AJ_FACT)
Function AsmIsBuild(rd As Variant) As Boolean
    AsmIsBuild = (CStr(rd(AJ_KIND)) = "KIT" And VarType(rd(AJ_FACT)) = 5)
End Function

' the records (rows of _ADJ, any state) of the build of the part EI canon that holds the record n: into nums(), their number; nPart — the record of the part itself (AJ_EI = AJ_NEW), the last of a build
Private Function AsmGroup(canon As String, n As Long, nums() As Long, ByRef nPart As Long) As Long
    Dim sh As Object, last As Long, d As Variant, i As Long, rows() As Long, isPart() As Boolean, cnt As Long, k As Long, start As Long, j As Long
    nPart = 0
    ReDim nums(AS_MAX + 1)
    sh = gDoc.Sheets.getByName(SH_ADJ)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, AJ_LAST, last).getDataArray()
    ReDim rows(UBound(d))
    ReDim isPart(UBound(d))
    For i = 0 To UBound(d)
        If CStr(d(i)(AJ_KIND)) = "KIT" And CStr(d(i)(AJ_NEW)) = canon And VarType(d(i)(AJ_FACT)) = 5 Then
            rows(cnt) = i + 1
            isPart(cnt) = (CStr(d(i)(AJ_EI)) = canon)
            cnt = cnt + 1
        End If
    Next i
    start = 0
    For k = 0 To cnt - 1
        If isPart(k) Then
            If n >= rows(start) And n <= rows(k) And k - start <= AS_MAX Then
                For j = start To k
                    nums(j - start) = rows(j)
                Next j
                nPart = rows(k)
                AsmGroup = k - start + 1
                Exit Function
            End If
            start = k + 1
        End If
    Next k
End Function

' "" when the build of the record rd may be cancelled now: the part EI holds exactly what the build left in it (nothing came or went since); a build that made the EI also needs the EI untouched
Function AsmDeleteProblem(rd As Variant) As String
    Dim canon As String, m As Long, nums() As Long, nPart As Long, cnt As Long, i As Long, ad As Object, p As Variant, bk As Double, bAdd As Boolean, d As Variant
    Dim nIss As Long, nRet As Long, nAdj As Long, nLive As Long
    canon = CStr(rd(AJ_NEW))
    If Not WmsOrders.StrictEI(canon, m) Then
        AsmDeleteProblem = "ERR-SYS:у сборки № " & rd(AJ_NO) & " неверный ЕИ детали «" & canon & "»"
        Exit Function
    End If
    cnt = AsmGroup(canon, CLng(rd(AJ_NO)), nums, nPart)
    If cnt < 2 Or nPart = 0 Then
        AsmDeleteProblem = "ERR-SYS:сборка детали " & canon & " в " & SH_ADJ & " неполная (записей " & cnt & ") — нужна самопроверка"
        Exit Function
    End If
    ad = gDoc.Sheets.getByName(SH_ADJ)
    For i = 0 To cnt - 1
        d = ad.getCellRangeByPosition(0, nums(i), AJ_LAST, nums(i)).getDataArray()(0)
        If CStr(d(AJ_STATE)) <> RV_LIVE Then
            AsmDeleteProblem = "ERR-SYS:в сборке детали " & canon & " запись № " & nums(i) & " уже не действует, а сборка не отменена — нужна самопроверка"
            Exit Function
        End If
    Next i
    p = ad.getCellRangeByPosition(0, nPart, AJ_LAST, nPart).getDataArray()(0)
    bAdd = (VarType(p(AJ_BOOK)) = 5)
    If Not WmsIssue.StockQty(m, bk) Then
        AsmDeleteProblem = "ERR:остаток " & canon & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If Abs(bk - p(AJ_FACT)) > 0.0000005 Then
        AsmDeleteProblem = "ERR:у " & canon & " остаток сейчас " & WmsIssue.QtyText(bk) & ", а после этой сборки был " & WmsIssue.QtyText(p(AJ_FACT)) & " — по детали были движения после неё " _
            & "(выдача, возврат, разделение по серийным номерам, другая сборка или приход). Сначала верните их (сторно) — последние первыми"
        Exit Function
    End If
    If Not bAdd Then
        If Not WmsReturn.LiveDependents(canon, nIss, nRet) Or Not WmsAdjust.LiveAdjustments(canon, nAdj) Or Not WmsSpecial.LiveLinesOfEI(canon, nLive) Then
            AsmDeleteProblem = "ERR-SYS:не удалось проверить движения " & canon & " (формула движка Calc) — отмена сборки не выполнена"
            Exit Function
        End If
        If nIss + nRet + (nAdj - 1) + nLive > 0 Then
            AsmDeleteProblem = "ERR:у " & canon & " есть действующие движения (выдач " & nIss & ", возвратов " & nRet & ", корректировок " & (nAdj - 1) & ", приходов " & nLive & ") — сначала их сторно"
            Exit Function
        End If
    End If
    AsmDeleteProblem = ""
End Function

' the build of the record rd cancelled whole: every part back on its EI, the pieces taken from the permanent EI of the part (an EI that this build made is closed, its number and its place in
' the index of the articles stay for the next build), every row of the build «Удалено (сторно)». One operation ASM_DEL. The result of ApplyOperation
Function AsmDelete(rd As Variant) As String
    Dim canon As String, m As Long, nums() As Long, nPart As Long, cnt As Long, i As Long, ad As Object, ash As Object, d As Variant, n As Long, e As Long, bAdd As Boolean
    Dim b0 As Double, b1 As Double, comps As String, r As Long, s As String, why As String, p As Variant, bk As Double, nq As Double
    why = AsmDeleteProblem(rd)
    If why <> "" Then
        AsmDelete = why
        Exit Function
    End If
    canon = CStr(rd(AJ_NEW))
    WmsOrders.StrictEI(canon, m)
    cnt = AsmGroup(canon, CLng(rd(AJ_NO)), nums, nPart)
    ad = gDoc.Sheets.getByName(SH_ADJ)
    ash = WmsAdjust.AdjustSheet()
    p = ad.getCellRangeByPosition(0, nPart, AJ_LAST, nPart).getDataArray()(0)
    bAdd = (VarType(p(AJ_BOOK)) = 5)
    nq = p(AJ_QTY)
    WmsIssue.StockQty(m, bk)
    PlanBegin("ASM_DEL", CStr(nums(0)))
    If Trim(gOpReason) <> "" Then PlanField("WHY", Trim(gOpReason))
    For i = 0 To cnt - 1
        n = nums(i)
        d = ad.getCellRangeByPosition(0, n, AJ_LAST, n).getDataArray()(0)
        r = WmsAdjust.AdjustRowOf(n, d(AJ_ROW))
        If r < 1 Then
            PlanBegin("", "")
            AsmDelete = "ERR-SYS:строка корректировки № " & n & " (сборка " & canon & ") не найдена на листе «" & SH_ADJUST & "»"
            Exit Function
        End If
        If n <> nPart Then
            WmsOrders.StrictEI(CStr(d(AJ_EI)), e)
            b0 = WmsIssue.StockSheet().getCellByPosition(SC_QTY, e).getValue()
            b1 = WmsIssue.Round3(b0 - d(AJ_QTY))
            comps = comps & IIf(comps <> "", ";", "") & d(AJ_EI) & ":" & NumStr(-d(AJ_QTY)) & ":" & NumStr(b0)
            PlanSetValue(SH_STOCK, e, SC_QTY, b1, False)
            WmsOrders.PlanStockMirror(e, b1)
        End If
        PlanSetValue(SH_ADJ, n, AJ_STATE, RV_STORNO, False)
        If CLng(d(AJ_ROW)) <> r Then PlanSetValue(SH_ADJ, n, AJ_ROW, r, False)
        PlanSetValue(SH_ADJUST, r, AC_CTL, ST_DELETED, False)
        If Trim(gOpReason) <> "" Then
            s = Trim(ash.getCellByPosition(AC_NOTE, r).getString())
            PlanDerived(SH_ADJUST, r, AC_NOTE, s & IIf(s <> "", " · ", "") & "разобрано " & Format(WmsOrders.Today(), "DD.MM.YYYY") & " — " & Trim(gOpReason))
        End If
        PlanLockBits(SH_ADJUST, r, 0, AC_LAST, ADJUST_LOCKS_POSTED)
        If n = nPart Then PlanField("DATE", Format(d(AJ_DATE), "YYYY-MM-DD"))
    Next i
    If bAdd Then
        PlanSetValue(SH_STOCK, m, SC_QTY, WmsIssue.Round3(bk - nq), False)
        WmsOrders.PlanStockMirror(m, WmsIssue.Round3(bk - nq))
    Else
        PlanSetValue(SH_STOCK, m, SC_QTY, 0, False)
        PlanSetValue(SH_STOCK, m, SC_STATE, EI_ST_STORNO, False)
        WmsOrders.PlanStockMirror(m, 0)
    End If
    PlanField("MODE", IIf(bAdd, "ADD", "NEW"))
    PlanField("NEW_EI", canon)
    PlanField("ADJ", nums(0))
    PlanField("ADJ_KIT", nPart)
    PlanField("N", nq)
    PlanField("COMPS", comps)
    PlanField("BAL_BEFORE", bk)
    PlanField("BAL_AFTER", IIf(bAdd, WmsIssue.Round3(bk - nq), 0))
    AsmDelete = ApplyOperation(0, 0)
End Function

' ================================================================ the windows

' the part under the cursor: on «Сборка деталей» — the part of the block the cursor stands in; on «Наличие» — the part whose article is that of the EI under the cursor; "" — none
Private Function AsmUnderCursor(names As Variant) As String
    Dim sh As Object, r As Long, sn As String, d As Variant, i As Long, cur As String, last As Long, art As String, key As String, nUnit As String, nArt As String, nCat As String, nSer As Boolean
    sn = gDoc.getCurrentController().getActiveSheet().Name
    r = gDoc.getCurrentController().getSelection().getRangeAddress().StartRow
    If r < 1 Then Exit Function
    If sn = AS_SHEET Then
        sh = gDoc.Sheets.getByName(AS_SHEET)
        last = WmsOrders.LastRow(sh)
        If r > last Then Exit Function
        d = sh.getCellRangeByPosition(0, 1, 0, r).getDataArray()
        For i = 0 To UBound(d)
            If Trim(CStr(d(i)(0))) <> "" Then cur = Trim(CStr(d(i)(0)))
        Next i
        For i = 0 To UBound(names)
            If LCase(names(i)) = LCase(cur) Then AsmUnderCursor = names(i)
        Next i
    ElseIf sn = SH_STOCK Then
        art = Trim(WmsIssue.StockSheet().getCellByPosition(SC_ART, r).getString())
        If art = "" Then Exit Function
        key = WmsSpecial.ArticleKey(art)
        For i = 0 To UBound(names)
            If AsmRead(CStr(names(i)), nArt, nUnit, nCat, nSer) = "" Then
                If WmsSpecial.ArticleKey(nArt) = key Then
                    AsmUnderCursor = names(i)
                    Exit Function
                End If
            End If
        Next i
    End If
End Function

' the lines of the description of the part for a window: what goes into one piece
Private Function AsmCompText() As String
    Dim j As Long
    For j = 0 To gKtN - 1
        AsmCompText = AsmCompText & IIf(AsmCompText <> "", "; ", "") & gKtName(j) & IIf(gKtArt(j) <> "" And gKtEi(j) = 0, " (" & gKtArt(j) & ")", "") & " ×" & WmsIssue.QtyText(gKtQty(j)) _
            & IIf(gKtUnit(j) <> "", " " & gKtUnit(j), "")
    Next j
End Function

' «Собрать деталь…»: the part (the one under the cursor, otherwise chosen from the list of the sheet), how many, where (only for a new EI), the date, the reason, the serial numbers
Sub BtnAsmBuild(Optional oEvent As Variant)
    Dim names As Variant, sName As String, v(4) As String, labels As Variant, lists(4) As String, art As String, unit As String, cat As String, ser As Boolean, why As String, cap As Long
    Dim sLimit As String, info As String, nQty As Double, dt As Double, msg As String, sers As Variant, res As String, k As Long, nEI As Long
    Dim bSers As Boolean, sPlace As String, nExist As Long, canon As String, aRow As Long, key As String, bExists As Boolean
    WmsInit()
    On Error GoTo EH
    names = AsmNames()
    If UBound(names) < 0 Then
        WmsUi.UiMessage("Деталей для сборки пока нет: опишите их на листе «" & AS_SHEET & "» (в каждой строке — деталь, её артикул, что входит, сколько на 1 шт) " _
            & "или соберите первый раз из ЕИ: выделите их на «" & SH_STOCK & "» и нажмите «Собрать деталь из ЕИ…» — описание запишется само.")
        Exit Sub
    End If
    sName = AsmUnderCursor(names)
    If sName = "" Then
        v(0) = names(0)
        If Not WmsPrimeUi.UiForm("Собрать деталь", "Какую деталь собрать? Список — детали, описанные на листе «" & AS_SHEET & "».", Array("Деталь"), v, Array(Join(names, "|")), "Дальше") Then Exit Sub
        sName = Trim(v(0))
        k = -1
        For nExist = 0 To UBound(names)
            If LCase(names(nExist)) = LCase(sName) Then
                k = nExist
                sName = names(nExist)
            End If
        Next nExist
        If k < 0 Then
            WmsUi.UiMessage("Детали «" & sName & "» нет в списке листа «" & AS_SHEET & "». Выберите деталь из списка или опишите её на листе.")
            Exit Sub
        End If
    End If
    cap = AsmCapacity(sName, sLimit)
    If cap < 0 Then
        WmsUi.UiMessage("Деталь «" & sName & "»: " & sLimit & ".")
        Exit Sub
    End If
    why = AsmRead(sName, art, unit, cat, ser)
    key = WmsSpecial.ArticleKey(art)
    bExists = (WmsSpecial.ArtLookup(key, nEI, canon, aRow, why) = 1)
    info = "Деталь «" & sName & "», артикул " & art & ". На 1 шт: " & AsmCompText() & "." & Chr(10) _
        & "Сейчас можно собрать: " & cap & " шт" & IIf(sLimit <> "" And cap = 0, " (не хватает: " & sLimit & ")", IIf(sLimit <> "", " (ограничивает: " & sLimit & ")", "")) & "." & Chr(10) _
        & IIf(bExists, "Собранные детали пойдут в постоянный " & canon & " (место «" & AsmPlaceOf(art) & "»).", "У детали ещё нет ЕИ: первая сборка создаст постоянный ЕИ — укажите место.")
    labels = Array("Сколько собрать", "Место (для нового ЕИ)", "Дата", "Основание", "Серийные номера (да/нет)")
    lists(4) = "нет|да"
    v(0) = "1"
    v(1) = IIf(bExists, AsmPlaceOf(art), "")
    v(2) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    v(3) = "сборка детали"
    v(4) = IIf(ser, "да", "нет")
    If Not WmsPrimeUi.UiForm("Собрать деталь — " & sName, info, labels, v, lists, "Собрать") Then Exit Sub
    If Not IsNumeric(Replace(Trim(v(0)), ",", ".")) Then
        WmsUi.UiMessage("Сколько собрать — целое число, например 2.")
        Exit Sub
    End If
    nQty = Val(Replace(Trim(v(0)), ",", "."))
    If Not WmsOrders.ParseDateTextAs(Trim(v(2)), "дата", dt, msg) Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    sPlace = Trim(v(1))
    bSers = (LCase(Left(Trim(v(4)), 1)) = "д")
    If bSers Then
        If Not AsmAskSerials(art, nQty, bExists, nEI, sers) Then Exit Sub
    End If
    res = AsmBuild(sName, nQty, sPlace, Int(dt), Trim(v(3)))
    AsmReport(res, bSers, sers)
    Exit Sub
EH:
    WmsUi.UiMessage("Сборка детали не выполнена: " & Error$ & " (строка " & Erl & ")")
End Sub

' the window of the serial numbers of the build: the list typed (one to a line), checked before anything is posted: the count equals the number of pieces, no number twice or taken, the EI of
' the part (if it is there) has no packing and a whole balance. True and sers, or False (the window closed or the text refused — the text stays for the next try)
Private Function AsmAskSerials(sArt As String, nQty As Double, bExists As Boolean, nEI As Long, ByRef sers As Variant) As Boolean
    Dim why As String, info As String, pu As String, pn As Double, bk As Double
    If bExists Then
        If WmsPack.PackOf(nEI, pu, pn) Then
            WmsUi.UiMessage(WmsIssue.EiCanon(nEI) & " имеет фасовку («" & pu & "») — серийные номера записывают штучным деталям; снимите фасовку («Фасовка ЕИ») или соберите без номеров.")
            Exit Function
        End If
        WmsIssue.StockQty(nEI, bk)
        If Abs(bk - Int(bk + 0.5)) > 0.0000001 Then
            WmsUi.UiMessage("Остаток " & WmsIssue.EiCanon(nEI) & " — " & WmsIssue.QtyText(bk) & ": серийные номера у целых деталей. Исправьте остаток и повторите.")
            Exit Function
        End If
    End If
    info = "Деталь " & sArt & " ×" & WmsIssue.QtyText(nQty) & ". Впишите серийные номера собранных деталей, по одному на строку (можно вставить столбец из таблицы; запятая и «;» тоже разделяют). " _
        & "Номеров должно быть ровно " & WmsIssue.QtyText(nQty) & ": каждая деталь станет своим ЕИ с тем же артикулом, своим номером и этикеткой."
    Do
        If Not WmsPrimeUi.UiMemo("Серийные номера собранных деталей", info, gSzText, "Дальше") Then Exit Function
        If WmsSerial.SerialsParse(gSzText, sers, why) Then
            If UBound(sers) + 1 = nQty Then
                AsmAskSerials = True
                Exit Function
            End If
            WmsUi.UiMessage("Серийных номеров " & (UBound(sers) + 1) & ", а собирается " & WmsIssue.QtyText(nQty) & " шт — должно быть поровну.")
        Else
            WmsUi.UiMessage("Список номеров: " & why & ".")
        End If
    Loop
End Function

' the answer to the storekeeper after AsmBuild; with serial numbers (bSers) they are written to the pieces now
Private Sub AsmReport(res As String, bSers As Boolean, sers As Variant)
    Dim a As Variant, canon As String, msg As String, nEI As Long, r2 As String, eis As String, k As Integer, b As Variant
    gUiLastMsg = res
    If Left(res, 3) <> "OK:" Then
        If Left(res, 5) <> "SKIP:" Then WmsUi.UiMessage(Mid(res, InStr(res, ":") + 1))
        Exit Sub
    End If
    a = Split(Mid(res, 4), "|", 5)
    canon = CStr(a(0))
    msg = "Готово: " & CStr(a(4)) & "."
    WmsItem.StockInfoRefresh(False)
    nEI = WmsView.EiNo(canon, CLng(SysNum(SK_NEXT_EI)))
    If bSers Then
        r2 = WmsSerial.SerialApply(canon, nEI, sers)
        If Left(r2, 3) = "OK:" Then
            b = Split(Mid(r2, 4), "|", 2)
            eis = CStr(b(1))
            gSzText = ""
            msg = msg & Chr(10) & "Серийные номера записаны: деталей " & b(0) & " (" & eis & ")."
            If gSzReplaced <> "" Then msg = msg & Chr(10) & "Прежний номер " & Mid(gSzReplaced, InStr(gSzReplaced, "|") + 1) & " ЕИ " _
                & Left(gSzReplaced, InStr(gSzReplaced, "|") - 1) & " записан в его примечание («ранее № …»)."
        Else
            WmsUi.UiMessage(msg & Chr(10) & Chr(10) & "Но серийные номера записаны не все: " & Mid(r2, InStr(r2, "|") + 1) & Chr(10) & "Номера остались в окне: встаньте на " & canon _
                & " на «" & SH_STOCK & "», нажмите «Серийные…» и впишите недостающие.")
            gUiLastMsg = res & " | " & r2
            Exit Sub
        End If
    End If
    k = WmsPrimeUi.UiButtons("Сборка детали", msg & Chr(10) & Chr(10) & "Напечатать этикетки?", Array("Напечатать этикетки", "Не сейчас"), Array(1, 0))
    If k = 1 Then AsmLabels(IIf(bSers, eis, canon))
End Sub

' the labels of the EIs "ЕИ-…,ЕИ-…", one of each (the page «Печать»)
Private Sub AsmLabels(sEis As String)
    Dim a As Variant, eis() As Long, counts() As Long, i As Long, msg As String
    a = Split(sEis, ",")
    ReDim eis(UBound(a))
    ReDim counts(UBound(a))
    For i = 0 To UBound(a)
        eis(i) = WmsView.EiNo(Trim(CStr(a(i))), CLng(SysNum(SK_NEXT_EI)))
        counts(i) = 1
    Next i
    msg = WmsItem.LabelPrintN(eis, counts, WmsItem.LabelSize())
    If Left(msg, 4) = "ERR:" Then WmsUi.UiMessage(Mid(msg, 5))
End Sub

' «Что можно собрать»: every described part with the number the stock allows now (or what is missing)
Sub BtnAsmCan(Optional oEvent As Variant)
    Dim names As Variant, i As Long, s As String, cap As Long, sLimit As String, n As Long
    WmsInit()
    On Error GoTo EH
    names = AsmNames()
    If UBound(names) < 0 Then
        WmsUi.UiMessage("Деталей для сборки пока нет: опишите их на листе «" & AS_SHEET & "».")
        Exit Sub
    End If
    For i = 0 To UBound(names)
        cap = AsmCapacity(CStr(names(i)), sLimit)
        If n < 40 Then s = s & Chr(10) & names(i) & ": " & IIf(cap < 0, "описание неверно — " & sLimit, IIf(cap = 0, "собрать нельзя — не хватает: " & sLimit, "можно собрать " & cap & " шт"))
        n = n + 1
    Next i
    If n > 40 Then s = s & Chr(10) & "… и ещё " & (n - 40)
    WmsUi.UiMessage("Что можно собрать из того, что есть на складе сейчас:" & s)
    Exit Sub
EH:
    WmsUi.UiMessage("Список не построен: " & Error$ & " (строка " & Erl & ")")
End Sub

' «Собрать деталь из ЕИ…» (on «Наличие»): the EIs of the rows selected there, a window — the name and the article of the part, how many, how much of each EI goes into ONE piece, the place,
' the date, the reason, the serial numbers, whether to write the composition on «Сборка деталей» — then the build (AsmFromEis)
Sub BtnAsmFromEis(Optional oEvent As Variant)
    Dim rows As Variant, i As Long, n As Long, st As Object, d As Variant, eis() As Long, nE As Long, labels() As String, v() As String, lists() As String
    Dim j As Long, s As String, nm As String, q() As String, nQty As Double, dt As Double, msg As String, seen As Boolean, sers As Variant, bSers As Boolean, res As String
    Dim art As String, key As String, canon As String, nEI As Long, aRow As Long, bExists As Boolean, why As String
    WmsInit()
    On Error GoTo EH
    rows = WmsPrint.SelRows(SH_STOCK)
    If UBound(rows) < 0 Then
        WmsUi.UiMessage("Выделите на листе «" & SH_STOCK & "» строки ЕИ, из которых собрать деталь (несколько строк — с Ctrl), и нажмите «Собрать деталь из ЕИ…» там же.")
        If gDoc.getCurrentController().getActiveSheet().getName() <> SH_STOCK Then WmsView.GoCell(SH_STOCK, 0, 1)
        Exit Sub
    End If
    st = WmsIssue.StockSheet()
    ReDim eis(UBound(rows))
    For i = 0 To UBound(rows)
        n = rows(i)
        d = st.getCellRangeByPosition(0, n, SC_STATE, n).getDataArray()(0)
        If CStr(d(SC_EI)) = WmsIssue.EiCanon(n) Then
            seen = False
            For j = 0 To nE - 1
                If eis(j) = n Then seen = True
            Next j
            If Not seen Then
                eis(nE) = n
                nE = nE + 1
            End If
        End If
    Next i
    If nE = 0 Then
        WmsUi.UiMessage("В выделенных строках «" & SH_STOCK & "» нет ЕИ. Выделите строки нужных ЕИ и нажмите «Собрать деталь из ЕИ…» снова.")
        Exit Sub
    End If
    If nE > AS_FROM_EIS_MAX Then
        WmsUi.UiMessage("Выделено ЕИ: " & nE & ". За один раз — не больше " & AS_FROM_EIS_MAX & "; для большой детали опишите состав на листе «" & AS_SHEET & "» и нажмите «Собрать деталь…».")
        Exit Sub
    End If
    ReDim Preserve eis(nE - 1)
    ReDim labels(nE + 7)
    ReDim v(nE + 7)
    ReDim lists(nE + 7)
    labels(0) = "Название детали"
    labels(1) = "Артикул детали (постоянный код)"
    labels(2) = "Сколько собрать"
    v(2) = "1"
    For i = 0 To nE - 1
        d = st.getCellRangeByPosition(0, eis(i), SC_STATE, eis(i)).getDataArray()(0)
        nm = CStr(d(SC_NAME))
        If Len(nm) > 16 Then nm = Left(nm, 15) & "…"
        labels(3 + i) = "ЕИ-" & eis(i) & ", " & d(SC_UNIT) & ": " & nm
        v(3 + i) = "1"
        If VarType(d(SC_QTY)) = 5 Then
            If d(SC_QTY) > 0 And d(SC_QTY) < 1 Then v(3 + i) = WmsIssue.QtyText(d(SC_QTY))
        End If
    Next i
    labels(nE + 3) = "Место (для нового ЕИ)"
    labels(nE + 4) = "Дата"
    v(nE + 4) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    labels(nE + 5) = "Основание"
    v(nE + 5) = "сборка детали"
    labels(nE + 6) = "Серийные номера (да/нет)"
    v(nE + 6) = "нет"
    lists(nE + 6) = "нет|да"
    labels(nE + 7) = "Записать состав на «" & AS_SHEET & "»"
    v(nE + 7) = "да"
    lists(nE + 7) = "да|нет"
    If Not WmsPrimeUi.UiForm("Собрать деталь из ЕИ", "Для каждого ЕИ — сколько его идёт на 1 деталь (в его единице). Детали снимутся именно с этих ЕИ, собранные пойдут в постоянный ЕИ детали " _
        & "(по артикулу: нет такого ЕИ — WMS заведёт его). Состав запишется на «" & AS_SHEET & "» — в следующий раз достаточно «Собрать деталь…».", labels, v, lists, "Собрать") Then Exit Sub
    If Not IsNumeric(Replace(Trim(v(2)), ",", ".")) Then
        WmsUi.UiMessage("Сколько собрать — целое число, например 2.")
        Exit Sub
    End If
    nQty = Val(Replace(Trim(v(2)), ",", "."))
    If Not WmsOrders.ParseDateTextAs(Trim(v(nE + 4)), "дата", dt, msg) Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    ReDim q(nE - 1)
    For i = 0 To nE - 1
        q(i) = Trim(v(3 + i))
    Next i
    art = Trim(v(1))
    key = WmsSpecial.ArticleKey(art)
    bSers = (LCase(Left(Trim(v(nE + 6)), 1)) = "д")
    If bSers And key <> "" Then
        bExists = (WmsSpecial.ArtLookup(key, nEI, canon, aRow, why) = 1)
        If Not AsmAskSerials(art, nQty, bExists, nEI, sers) Then Exit Sub
    End If
    res = AsmFromEis(eis, q, Trim(v(0)), art, nQty, Trim(v(nE + 3)), Int(dt), Trim(v(nE + 5)), LCase(Left(Trim(v(nE + 7)), 1)) <> "н")
    AsmReport(res, bSers, sers)
    Exit Sub
EH:
    WmsUi.UiMessage("Сборка из ЕИ не выполнена: " & Error$ & " (строка " & Erl & ")")
End Sub

' ================================================================ test seams (TEST books only)

' "<how many can be built>|<what limits it>" of the part sName
Function TestAsmCapacity(sName As String) As String
    Dim cap As Long, s As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestAsmCapacity = "REFUSED:не тестовая книга"
        Exit Function
    End If
    cap = AsmCapacity(sName, s)
    TestAsmCapacity = cap & "|" & s
End Function

' "<article>|<unit>|<category>|<serial numbers 1/0>|<the parts that go in: name/article/EI/quantity/unit, …>" of the part sName, or "ERR:<why>"
Function TestAsmRead(sName As String) As String
    Dim art As String, unit As String, cat As String, ser As Boolean, why As String, j As Long, s As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestAsmRead = "REFUSED:не тестовая книга"
        Exit Function
    End If
    why = AsmRead(sName, art, unit, cat, ser)
    If why <> "" Then
        TestAsmRead = "ERR:" & why
        Exit Function
    End If
    For j = 0 To gKtN - 1
        s = s & IIf(j > 0, ", ", "") & gKtName(j) & "/" & gKtArt(j) & "/" & gKtEi(j) & "/" & WmsIssue.QtyText(gKtQty(j)) & "/" & gKtUnit(j)
    Next j
    TestAsmRead = art & "|" & unit & "|" & cat & "|" & IIf(ser, "1", "0") & "|" & s
End Function
