' WmsKit — 1.0 (Ф8 и сборка): комплекты. Лист «Комплекты» — свой у кладовщика: состав каждого комплекта (позиция,
' артикул, количество на 1 комплект). «Выдать комплект…» заполняет строки «Выдач» нужными ЕИ — старые ЕИ первыми — и
' выделяет их: проводит их обычное «Провести». «Собрать комплект…» — одна операция журнала KIT_BUILD: позиции снимаются
' с их ЕИ строками «Сборка» на «Корректировках», готовые комплекты становятся новым ЕИ (единица «компл»). «Удалить» любой
' строки сборки разбирает комплект целиком (KIT_DEL), пока у ЕИ комплекта не было движений. 1.0.1: позиция состава может
' быть номером ЕИ («ЕИ-427» в «Артикул» или «Наименование») — тогда она берётся только с этого ЕИ; «Собрать из ЕИ…» на
' «Наличии» собирает комплект из выделенных ЕИ (та же операция KIT_BUILD) и может записать его состав на «Комплекты».
' Library WmsPlus.
Option Explicit

Public Const SH_KITS = "Комплекты"
Private Const KT_KIT = 0
Private Const KT_NAME = 1
Private Const KT_ART = 2
Private Const KT_QTY = 3
Private Const KT_UNIT = 4
Private Const KT_NOTE = 5
Private Const KT_LAST = 5
Private Const KIT_UNIT = "компл"
Private Const KIT_KIND_NAME = "Сборка"
Private Const KIT_MAX_PARTS = 200

' the components of the kit read by KitParts: name, article, unit, quantity for one kit
Global gKtN As Long
Global gKtName() As String
Global gKtArt() As String
Global gKtUnit() As String
Global gKtQty() As Double
' 1.0.1: the EI a component is taken from (its number; 0 — any EI of its article or name), and the components set by
' «Собрать из ЕИ…» instead of the sheet (KitParts keeps them once)
Global gKtEi() As Long
Global gKtPreset As Boolean
Private Const KIT_FROM_EIS_MAX = 12
' the EIs picked by KitPick: EI number, quantity, the component it serves
Global gKpN As Long
Global gKpEi() As Long
Global gKpQty() As Double
Global gKpPart() As Long

' ================================================================ the directory «Комплекты»

' the names of the kits in their order on the sheet (a row with an empty A continues the kit above)
Function KitNames() As Variant
    Dim sh As Object, last As Long, d As Variant, i As Long, cur As String, out() As String, n As Long, j As Long, seen As Boolean
    KitNames = Array()
    If Not gDoc.Sheets.hasByName(SH_KITS) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_KITS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, KT_LAST, last).getDataArray()
    ReDim out(UBound(d))
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(KT_KIT))) <> "" Then cur = Trim(CStr(d(i)(KT_KIT)))
        If cur <> "" And Trim(CStr(d(i)(KT_NAME))) <> "" Then
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
    KitNames = out
End Function

' the components of kit sKit into gKt*; "" or the problem (an unknown kit, a quantity that is not a number)
Function KitParts(sKit As String) As String
    Dim sh As Object, last As Long, d As Variant, i As Long, cur As String, q As Double, msg As String
    If gKtPreset Then
        gKtPreset = False
        Exit Function
    End If
    gKtN = 0
    ReDim gKtName(KIT_MAX_PARTS)
    ReDim gKtArt(KIT_MAX_PARTS)
    ReDim gKtUnit(KIT_MAX_PARTS)
    ReDim gKtQty(KIT_MAX_PARTS)
    ReDim gKtEi(KIT_MAX_PARTS)
    If Not gDoc.Sheets.hasByName(SH_KITS) Then
        KitParts = "листа «" & SH_KITS & "» нет"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_KITS)
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        d = sh.getCellRangeByPosition(0, 1, KT_LAST, last).getDataArray()
        For i = 0 To UBound(d)
            If Trim(CStr(d(i)(KT_KIT))) <> "" Then cur = Trim(CStr(d(i)(KT_KIT)))
            If LCase(cur) = LCase(Trim(sKit)) And Trim(CStr(d(i)(KT_NAME))) <> "" Then
                If VarType(d(i)(KT_QTY)) = 5 Then
                    q = d(i)(KT_QTY)
                ElseIf Not WmsOrders.ParseQtyValueAs(d(i)(KT_QTY), "D", q, msg) Then
                    KitParts = "«" & SH_KITS & "» строка " & (i + 2) & ": количество «" & d(i)(KT_QTY) & "» — " & msg
                    Exit Function
                End If
                If q <= 0 Then
                    KitParts = "«" & SH_KITS & "» строка " & (i + 2) & ": количество на 1 комплект должно быть больше нуля"
                    Exit Function
                End If
                If gKtN > KIT_MAX_PARTS Then
                    KitParts = "в комплекте больше " & KIT_MAX_PARTS & " позиций"
                    Exit Function
                End If
                gKtName(gKtN) = Trim(CStr(d(i)(KT_NAME)))
                gKtArt(gKtN) = Trim(CStr(d(i)(KT_ART)))
                gKtUnit(gKtN) = Trim(CStr(d(i)(KT_UNIT)))
                gKtQty(gKtN) = q
                ' 1.0.1: a number of an EI in «Артикул» (or in «Наименование») — the component is taken from that EI only
                gKtEi(gKtN) = KtEiRef(gKtArt(gKtN))
                If gKtEi(gKtN) = 0 And gKtArt(gKtN) = "" Then gKtEi(gKtN) = KtEiRef(gKtName(gKtN))
                If gKtEi(gKtN) < 0 Then
                    KitParts = "«" & SH_KITS & "» строка " & (i + 2) & ": «" & IIf(gKtArt(gKtN) <> "", gKtArt(gKtN), gKtName(gKtN)) & "» — такого ЕИ нет"
                    Exit Function
                End If
                gKtN = gKtN + 1
            End If
        Next i
    End If
    If gKtN = 0 Then KitParts = "комплекта «" & sKit & "» нет на листе «" & SH_KITS & "» (или у него нет позиций)"
End Function

' a text as Calc compares it: lower case, single spaces
Function KtKey(ByVal s As String) As String
    s = LCase(Trim(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    KtKey = s
End Function

' 1.0.1: the EI named by s («ЕИ-427», «ЕИ-00000427», «EI 427» — always with «ЕИ» / «EI»: a bare number is an article): its
' number; 0 — s does not name an EI; -1 — it names an EI that WMS does not have
Function KtEiRef(ByVal s As String) As Long
    Dim t As String
    s = Trim(s)
    t = UCase(Left(s, 2))
    If t <> "ЕИ" And t <> "EI" Then Exit Function
    t = Trim(Mid(s, 3))
    If Left(t, 1) = "-" Then t = Trim(Mid(t, 2))
    If t = "" Or Not IsDigits(t) Then Exit Function
    KtEiRef = -1
    If Len(t) > 8 Then Exit Function
    If CLng(t) >= 1 And CLng(t) < CLng(SysNum(SK_NEXT_EI)) Then KtEiRef = CLng(t)
End Function

' the EIs for nNeed of component j, the oldest first (by the article when the component has one, otherwise by the name; the
' unit when it is given), what an earlier component already took counted out. Appended to gKp*; the quantity taken
Function KitPick(j As Long, nNeed As Double) As Double
    Dim last As Long, cond As String, key As String, items As Variant, nItems As Long, i As Long, n As Long, st As Object, q As Double, used As Double
    Dim k As Long, take As Double, rest As Double, d As Variant
    rest = WmsIssue.Round3(nNeed)
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last < 1 Then Exit Function
    If gKtEi(j) > 0 Then
        ' 1.0.1: the one EI named in the composition — what is left of it after the components above
        n = gKtEi(j)
        d = WmsIssue.StockSheet().getCellRangeByPosition(0, n, SC_STATE, n).getDataArray()(0)
        If CStr(d(SC_EI)) = WmsIssue.EiCanon(n) And VarType(d(SC_QTY)) = 5 And CStr(d(SC_STATE)) <> EI_ST_STORNO And CStr(d(SC_STATE)) <> EI_ST_REVIEW Then
            items = Array(CStr(n))
            nItems = 1
        End If
    ElseIf gKtArt(j) <> "" Then
        key = KtKey(gKtArt(j))
        cond = "(LOWER(TRIM(" & WmsView.ColRng(SH_STOCK, SC_ART, 1, last) & "))=""" & Replace(key, """", """""") & """)"
    Else
        key = KtKey(gKtName(j))
        cond = "(LOWER(TRIM(" & WmsView.ColRng(SH_STOCK, SC_NAME, 1, last) & "))=""" & Replace(key, """", """""") & """)"
    End If
    If gKtEi(j) = 0 Then
        If gKtUnit(j) <> "" Then cond = cond & "*(LOWER(SUBSTITUTE(TRIM(" & WmsView.ColRng(SH_STOCK, SC_UNIT, 1, last) & ");""."";""""))=""" _
            & Replace(Replace(KtKey(gKtUnit(j)), ".", ""), """", """""") & """)"
        cond = cond & "*ISNUMBER(" & WmsView.ColRng(SH_STOCK, SC_QTY, 1, last) & ")*(" & WmsView.ColRng(SH_STOCK, SC_QTY, 1, last) & ">0.0000005)*(" _
            & WmsView.ColRng(SH_STOCK, SC_STATE, 1, last) & "<>""" & EI_ST_STORNO & """)*(" & WmsView.ColRng(SH_STOCK, SC_STATE, 1, last) & "<>""" & EI_ST_REVIEW & """)"
    End If
    If gKtEi(j) = 0 Then items = WmsView.EngList("TEXTJOIN("";"";1;IF(" & cond & ";ROW(" & WmsView.ColRng(SH_STOCK, SC_QTY, 1, last) & ")-1;""""))", nItems)
    st = WmsIssue.StockSheet()
    For i = 0 To nItems - 1
        If rest <= 0.0000005 Then Exit For
        n = CLng(items(i))
        q = st.getCellByPosition(SC_QTY, n).getValue()
        used = 0
        For k = 0 To gKpN - 1
            If gKpEi(k) = n Then used = used + gKpQty(k)
        Next k
        q = WmsIssue.Round3(q - used)
        If q > 0.0000005 Then
            take = IIf(q < rest, q, rest)
            If gKpN > UBound(gKpEi) Then
                ReDim Preserve gKpEi(2 * gKpN + 1)
                ReDim Preserve gKpQty(2 * gKpN + 1)
                ReDim Preserve gKpPart(2 * gKpN + 1)
            End If
            gKpEi(gKpN) = n
            gKpQty(gKpN) = take
            gKpPart(gKpN) = j
            gKpN = gKpN + 1
            rest = WmsIssue.Round3(rest - take)
            KitPick = WmsIssue.Round3(KitPick + take)
        End If
    Next i
End Function

' every component of the kit for nKits kits picked (gKp*); the shortages as text («Кабель ВВГ: нужно 10 м, есть 4 м»), "" — all
Function KitPickAll(nKits As Double) As String
    Dim j As Long, need As Double, got As Double, s As String
    gKpN = 0
    ReDim gKpEi(31)
    ReDim gKpQty(31)
    ReDim gKpPart(31)
    For j = 0 To gKtN - 1
        need = WmsIssue.Round3(gKtQty(j) * nKits)
        got = KitPick(j, need)
        If got < need - 0.0000005 Then s = s & IIf(s <> "", "; ", "") & gKtName(j) & IIf(gKtArt(j) <> "", " (" & gKtArt(j) & ")", "") & ": нужно " _
            & WmsIssue.QtyText(need) & IIf(gKtUnit(j) <> "", " " & gKtUnit(j), "") & ", есть " & WmsIssue.QtyText(got)
    Next j
    KitPickAll = s
End Function

Sub NavKits(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SH_KITS, 1, 0)
End Sub

' ================================================================ Ф8 «Выдать комплект…»

Sub BtnKitIssue(Optional oEvent As Variant)
    Dim s As String, v(3) As String, names As Variant, n As Double, dt As Double, msg As String
    WmsInit()
    names = KitNames()
    If UBound(names) < 0 Then
        WmsUi.UiMessage("Комплектов нет: опишите их на листе «" & SH_KITS & "» — в каждой строке комплект, позиция, артикул, количество на 1 комплект.")
        Exit Sub
    End If
    v(0) = names(0)
    v(1) = "1"
    v(3) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    If Not WmsPrimeUi.UiForm("Выдать комплект", "WMS заполнит строки «Выдач» нужными ЕИ (старые первыми) и выделит их — проверьте и нажмите «Провести».", _
        Array("Комплект", "Сколько комплектов", "Кому", "Дата"), v, Array(Join(names, "|"), "", "", ""), "Заполнить строки") Then Exit Sub
    If Not IsNumeric(Replace(Trim(v(1)), ",", ".")) Then
        WmsUi.UiMessage("Сколько комплектов — число, например 2.")
        Exit Sub
    End If
    n = Val(Replace(Trim(v(1)), ",", "."))
    If Not WmsOrders.ParseDateTextAs(Trim(v(3)), "дата", dt, msg) Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = KitIssueFill(Trim(v(0)), n, Trim(v(2)), Int(dt))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5)) Else WmsOrdersUi.UiInfo(Mid(s, 4))
End Sub

' the rows of «Выдачи» for nKits kits sKit to sWho on day dt (B «Комплект … ×n»; nothing posted); the rows selected. A shortage
' is reported, what there is is filled. OK:… / ERR:…
Function KitIssueFill(sKit As String, nKits As Double, sWho As String, dt As Double) As String
    Dim why As String, sShort As String, sh As Object, r0 As Long, i As Long, um As Object, inCtx As Boolean, sDoc As String, sel As Object
    WmsInit()
    If nKits <= 0 Then
        KitIssueFill = "ERR:сколько комплектов — больше нуля"
        gUiLastMsg = KitIssueFill
        Exit Function
    End If
    why = KitParts(sKit)
    If why <> "" Then
        KitIssueFill = "ERR:" & why
        gUiLastMsg = KitIssueFill
        Exit Function
    End If
    sShort = KitPickAll(nKits)
    If gKpN = 0 Then
        KitIssueFill = "ERR:ни одной позиции комплекта «" & sKit & "» нет в остатке: " & sShort
        gUiLastMsg = KitIssueFill
        Exit Function
    End If
    sh = WmsIssue.IssuesSheet()
    r0 = WmsOrders.LastRow(sh) + 1
    sDoc = "Комплект «" & sKit & "»" & IIf(nKits <> 1, " ×" & WmsIssue.QtyText(nKits), "")
    On Error GoTo EH
    um = gDoc.getUndoManager()
    um.enterUndoContext("WMS: выдать комплект " & sKit)
    inCtx = True
    For i = 0 To gKpN - 1
        sh.getCellByPosition(IC_DOC, r0 + i).setString(sDoc)
        sh.getCellByPosition(IC_EI, r0 + i).setString(WmsIssue.EiCanon(gKpEi(i)))
        sh.getCellByPosition(IC_QTY, r0 + i).setString(WmsIssue.QtyText(gKpQty(i)))
        sh.getCellByPosition(IC_DATE, r0 + i).setValue(dt)
        If sWho <> "" Then sh.getCellByPosition(IC_WHO, r0 + i).setString(sWho)
        WmsIssue.PreviewRow(r0 + i, True, sWho <> "")
    Next i
    um.leaveUndoContext()
    inCtx = False
    sel = sh.getCellRangeByPosition(0, r0, IC_LAST, r0 + gKpN - 1)
    gDoc.getCurrentController().setActiveSheet(sh)
    gDoc.getCurrentController().select(sel)
    KitIssueFill = "OK:заполнено строк выдачи: " & gKpN & " (строки " & (r0 + 1) & "–" & (r0 + gKpN) & ", документ «" & sDoc & "»)" _
        & IIf(sShort <> "", Chr(10) & "Не хватает: " & sShort, "") & Chr(10) & "Строки выделены — проверьте «Контроль» и нажмите «Провести»."
    gUiLastMsg = KitIssueFill
    Exit Function
EH:
    KitIssueFill = "ERR:строки не заполнены полностью: " & Error$ & " — отмените (Ctrl+Z) и повторите"
    gUiLastMsg = KitIssueFill
    On Error Resume Next
    If inCtx Then um.leaveUndoContext()
End Function

' ================================================================ the build of a kit (KIT_BUILD)

Sub BtnKitBuild(Optional oEvent As Variant)
    Dim s As String, v(4) As String, names As Variant, n As Double, dt As Double, msg As String
    WmsInit()
    names = KitNames()
    If UBound(names) < 0 Then
        WmsUi.UiMessage("Комплектов нет: опишите их на листе «" & SH_KITS & "» — в каждой строке комплект, позиция, артикул, количество на 1 комплект.")
        Exit Sub
    End If
    v(0) = names(0)
    v(1) = "1"
    v(3) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    v(4) = "сборка комплекта"
    If Not WmsPrimeUi.UiForm("Собрать комплект", "Позиции снимутся с их ЕИ (старые первыми), готовые комплекты станут новым ЕИ с единицей «" & KIT_UNIT & "». " _
        & "Одна операция журнала; «Удалить» любой её строки на «Корректировках» разберёт комплект.", Array("Комплект", "Сколько собрать", "Место готовых", _
        "Дата", "Основание"), v, Array(Join(names, "|"), "", "", "", ""), "Собрать") Then Exit Sub
    If Not IsNumeric(Replace(Trim(v(1)), ",", ".")) Then
        WmsUi.UiMessage("Сколько собрать — число, например 2.")
        Exit Sub
    End If
    n = Val(Replace(Trim(v(1)), ",", "."))
    If Not WmsOrders.ParseDateTextAs(Trim(v(3)), "дата", dt, msg) Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    s = KitBuild(Trim(v(0)), n, Trim(v(2)), Int(dt), Trim(v(4)))
    If Left(s, 3) = "OK:" Then
        WmsOrdersUi.UiInfo(Mid(s, 4))
    ElseIf Left(s, 5) <> "SKIP:" Then
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    End If
End Sub

' nKits kits sKit built on day dt: the components taken from their EIs (rows «Сборка» on «Корректировки»), the kits a new EI
' on place sPlace. One operation KIT_BUILD. OK:<new EI …> / ERR:… / SKIP:…
Function KitBuild(sKit As String, nKits As Double, sPlace As String, dt As Double, sReason As String) As String
    Dim why As String, sShort As String, nAdj0 As Long, m As Long, canonK As String, ash As Object, r0 As Long, i As Long, n As Long, r As Long
    Dim st As Object, d As Variant, b0 As Double, b1 As Double, comps As String, sBatch As String, txt As String, res As String, pu As String, pn As Double
    WmsInit()
    On Error GoTo EH
    If nKits <= 0 Or nKits <> Int(nKits) Then
        KitBuild = "ERR:сколько собрать — целое число больше нуля"
        gUiLastMsg = KitBuild
        Exit Function
    End If
    If sReason = "" Then sReason = "сборка комплекта"
    why = KitParts(sKit)
    If why <> "" Then
        KitBuild = "ERR:" & why
        gUiLastMsg = KitBuild
        Exit Function
    End If
    sShort = KitPickAll(nKits)
    If sShort <> "" Then
        KitBuild = "ERR:комплект «" & sKit & "» ×" & WmsIssue.QtyText(nKits) & " не собрать — не хватает: " & sShort
        gUiLastMsg = KitBuild
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        KitBuild = why
        gUiLastMsg = KitBuild
        Exit Function
    End If
    nAdj0 = CLng(SysNum(SK_NEXT_ADJ))
    For i = 0 To gKpN
        why = WmsAdjust.NewAdjProblem(nAdj0 + i)
        If why <> "" Then
            KitBuild = "ERR-SYS:" & why
            gUiLastMsg = KitBuild
            Exit Function
        End If
    Next i
    m = CLng(SysNum(SK_NEXT_EI))
    why = WmsOrders.NewEIProblem(m)
    If why <> "" Then
        KitBuild = "ERR-SYS:" & why
        gUiLastMsg = KitBuild
        Exit Function
    End If
    canonK = WmsIssue.EiCanon(m)
    st = WmsIssue.StockSheet()
    For i = 0 To gKpN - 1
        d = st.getCellRangeByPosition(0, gKpEi(i), SC_LAST, gKpEi(i)).getDataArray()(0)
        txt = txt & Chr(10) & "  " & d(SC_EI) & " — " & d(SC_NAME) & ": " & WmsPack.QtyWithParts(gKpEi(i), gKpQty(i), CStr(d(SC_UNIT)))
    Next i
    If Not WmsUi.UiConfirm("Собрать «" & sKit & "» ×" & WmsIssue.QtyText(nKits) & " → новый " & canonK & IIf(sPlace <> "", ", место «" & sPlace & "»", "") & "?" _
        & Chr(10) & Chr(10) & "Будет снято:" & txt) Then
        KitBuild = "SKIP:сборка не подтверждена"
        gUiLastMsg = KitBuild
        Exit Function
    End If
    ash = WmsAdjust.AdjustSheet()
    r0 = WmsOrders.LastRow(ash) + 1
    sBatch = "Сборка «" & sKit & "» № " & nAdj0
    PlanBegin("KIT_BUILD", CStr(nAdj0))
    For i = 0 To gKpN - 1
        n = nAdj0 + i
        r = r0 + i
        d = st.getCellRangeByPosition(0, gKpEi(i), SC_LAST, gKpEi(i)).getDataArray()(0)
        b0 = d(SC_QTY)
        b1 = WmsIssue.Round3(b0 - gKpQty(i))
        comps = comps & IIf(comps <> "", ";", "") & d(SC_EI) & ":" & NumStr(gKpQty(i)) & ":" & NumStr(b0)
        PlanSetValue(SH_STOCK, gKpEi(i), SC_QTY, b1, False)
        KtAdjRow(n, CStr(d(SC_EI)), -gKpQty(i), r, CStr(d(SC_PLACE)), "", dt, canonK)
        KtSheetRow(r, n, CStr(d(SC_EI)), CStr(d(SC_NAME)), CStr(d(SC_ART)), CStr(d(SC_UNIT)), gKpQty(i), "", dt, sReason, CStr(d(SC_PLACE)), b0, b1, _
            ST_POSTED & " · в комплект " & canonK, sBatch)
        WmsOrders.PlanStockMirror(gKpEi(i), b1)
    Next i
    ' the kit: its EI and its own row «Сборка» (+N)
    n = nAdj0 + gKpN
    r = r0 + gKpN
    PlanSetValue(SH_STOCK, m, SC_EI, canonK, False)
    PlanSetValue(SH_STOCK, m, SC_NAME, sKit, False)
    PlanSetValue(SH_STOCK, m, SC_UNIT, KIT_UNIT, False)
    PlanSetValue(SH_STOCK, m, SC_QTY, nKits, False)
    PlanSetValue(SH_STOCK, m, SC_PLACE, sPlace, False)
    PlanSetValue(SH_STOCK, m, SC_CAT, "Комплект", False)
    PlanSetValue(SH_STOCK, m, SC_STATE, EI_ST_ACTIVE, False)
    PlanSetValue(SH_STOCK, m, SC_SRC, "Сборка: корректировки № " & nAdj0 & "–" & n, False)
    PlanSetValue(SH_STOCK, m, SC_STYPE, KIT_KIND_NAME, False)
    KtAdjRow(n, canonK, nKits, r, "", sPlace, dt, canonK)
    KtSheetRow(r, n, canonK, sKit, "", KIT_UNIT, nKits, sPlace, dt, sReason, "", 0, nKits, ST_POSTED & " · комплект из " & gKpN & " поз.", sBatch)
    PlanSetValue(SYS_SHEET, SK_NEXT_ADJ, 1, n + 1, False)
    PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, m + 1, False)
    PlanField("KIT", sKit)
    PlanField("N", nKits)
    PlanField("NEW_EI", canonK)
    PlanField("ADJ", nAdj0)
    PlanField("ADJ_KIT", n)
    PlanField("COMPS", comps)
    PlanField("PLACE", sPlace)
    PlanField("DATE", Format(dt, "YYYY-MM-DD"))
    PlanField("REASON", sReason)
    PlanField("BATCH", sBatch)
    res = ApplyOperation(0, 0)
    If Left(res, 3) <> "OK:" Then
        KitBuild = res
        gUiLastMsg = KitBuild
        Exit Function
    End If
    KitBuild = "OK:" & canonK & " — собран «" & sKit & "» ×" & WmsIssue.QtyText(nKits) & " из " & gKpN & " ЕИ (корректировки № " & nAdj0 & "–" & n _
        & " на листе «" & SH_ADJUST & "»)"
    gUiLastMsg = KitBuild
    Exit Function
EH:
    KitBuild = "ERR-SYS:внутренняя ошибка сборки: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = KitBuild
End Function

' ================================================================ 1.0.1 «Собрать из ЕИ…»

' «Собрать из ЕИ…» (on «Наличие»; on «Комплекты» it shows the way): the EIs of the rows selected on «Наличие», a window — the
' name of the kit, how many, how much of each EI goes into one kit, the place, the date, the reason, whether to write the
' composition on «Комплекты» — then the build (KitFromEis)
Sub BtnKitFromEis(Optional oEvent As Variant)
    Dim rows As Variant, i As Long, n As Long, st As Object, d As Variant, eis() As Long, nE As Long, labels() As String, v() As String, lists() As String
    Dim k As Long, j As Long, s As String, nm As String, q() As String, nKits As Double, dt As Double, msg As String, seen As Boolean
    WmsInit()
    rows = WmsPrint.SelRows(SH_STOCK)
    If UBound(rows) < 0 Then
        WmsUi.UiMessage("Выделите на листе «" & SH_STOCK & "» строки ЕИ, из которых собрать комплект (несколько строк — с Ctrl), и нажмите «Собрать из ЕИ…» там же.")
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
        WmsUi.UiMessage("В выделенных строках «" & SH_STOCK & "» нет ЕИ. Выделите строки нужных ЕИ и нажмите «Собрать из ЕИ…» снова.")
        Exit Sub
    End If
    If nE > KIT_FROM_EIS_MAX Then
        WmsUi.UiMessage("Выделено ЕИ: " & nE & ". За один раз — не больше " & KIT_FROM_EIS_MAX & "; для большого комплекта опишите состав на листе «" & SH_KITS _
            & "» (в «Артикул» можно писать номер ЕИ) и нажмите «Собрать комплект…».")
        Exit Sub
    End If
    ReDim Preserve eis(nE - 1)
    ReDim labels(nE + 5)
    ReDim v(nE + 5)
    ReDim lists(nE + 5)
    labels(0) = "Название комплекта"
    labels(1) = "Сколько собрать"
    v(1) = "1"
    For i = 0 To nE - 1
        d = st.getCellRangeByPosition(0, eis(i), SC_STATE, eis(i)).getDataArray()(0)
        nm = CStr(d(SC_NAME))
        If Len(nm) > 16 Then nm = Left(nm, 15) & "…"
        labels(2 + i) = "ЕИ-" & eis(i) & ", " & d(SC_UNIT) & ": " & nm
        v(2 + i) = "1"
        If VarType(d(SC_QTY)) = 5 Then
            If d(SC_QTY) > 0 And d(SC_QTY) < 1 Then v(2 + i) = WmsIssue.QtyText(d(SC_QTY))
        End If
        If i = 0 Then v(nE + 2) = CStr(d(SC_PLACE))
    Next i
    labels(nE + 2) = "Место готовых"
    labels(nE + 3) = "Дата"
    v(nE + 3) = Format(WmsOrders.Today(), "DD.MM.YYYY")
    labels(nE + 4) = "Основание"
    v(nE + 4) = "сборка комплекта"
    labels(nE + 5) = "Записать состав на «" & SH_KITS & "»"
    v(nE + 5) = "да"
    lists(nE + 5) = "да|нет"
    If Not WmsPrimeUi.UiForm("Собрать из ЕИ", "Для каждого ЕИ — сколько его идёт на 1 комплект (в его единице). Позиции снимутся именно с этих ЕИ, " _
        & "готовые комплекты станут новым ЕИ с единицей «" & KIT_UNIT & "».", labels, v, lists, "Собрать") Then Exit Sub
    If Not IsNumeric(Replace(Trim(v(1)), ",", ".")) Then
        WmsUi.UiMessage("Сколько собрать — целое число, например 2.")
        Exit Sub
    End If
    nKits = Val(Replace(Trim(v(1)), ",", "."))
    If Not WmsOrders.ParseDateTextAs(Trim(v(nE + 3)), "дата", dt, msg) Then
        WmsUi.UiMessage(msg)
        Exit Sub
    End If
    ReDim q(nE - 1)
    For i = 0 To nE - 1
        q(i) = Trim(v(2 + i))
    Next i
    s = KitFromEis(eis, q, Trim(v(0)), nKits, Trim(v(nE + 2)), Int(dt), Trim(v(nE + 4)), LCase(Left(Trim(v(nE + 5)), 1)) <> "н")
    If Left(s, 3) = "OK:" Then
        ' the balances and «Пришло» / «Осталось» of «Наличие» (the sheet stays open: no new opening refreshes them)
        WmsItem.StockInfoRefresh(False)
        WmsOrdersUi.UiInfo(Mid(s, 4))
    ElseIf Left(s, 5) <> "SKIP:" Then
        WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    End If
End Sub

' nKits kits sKit from EIs eis (numbers), q(i) of EI eis(i) for one kit (a text: «2», «0,5», «1 кан»): the operation KIT_BUILD
' of «Собрать комплект…» with this composition; bSave — the composition written on «Комплекты» after the build (the EIs in
' «Артикул»: the next «Собрать комплект…» takes them again). OK:… / ERR:… / SKIP:…
Function KitFromEis(eis As Variant, q As Variant, sKit As String, nKits As Double, sPlace As String, dt As Double, sReason As String, bSave As Boolean) As String
    Dim i As Long, j As Long, n As Long, st As Object, d As Variant, x As Double, msg As String, res As String, names As Variant, ks As Object, r As Long
    Dim out() As Variant, b As Double, info As String
    WmsInit()
    On Error GoTo EH
    gKtPreset = False
    If Trim(sKit) = "" Then
        KitFromEis = "ERR:дайте комплекту название — так он будет называться в «" & SH_STOCK & "» (новый ЕИ)"
        gUiLastMsg = KitFromEis
        Exit Function
    End If
    If UBound(eis) < 0 Or UBound(eis) >= KIT_FROM_EIS_MAX Or UBound(q) <> UBound(eis) Then
        KitFromEis = "ERR:ЕИ для комплекта — от 1 до " & KIT_FROM_EIS_MAX
        gUiLastMsg = KitFromEis
        Exit Function
    End If
    If bSave Then
        names = KitNames()
        For i = 0 To UBound(names)
            If LCase(names(i)) = LCase(Trim(sKit)) Then
                KitFromEis = "ERR:комплект «" & sKit & "» уже описан на листе «" & SH_KITS & "» — дайте другое название, соберите его кнопкой «Собрать " _
                    & "комплект…» или ответьте «нет» на «Записать состав»"
                gUiLastMsg = KitFromEis
                Exit Function
            End If
        Next i
    End If
    st = WmsIssue.StockSheet()
    gKtN = 0
    ReDim gKtName(KIT_MAX_PARTS)
    ReDim gKtArt(KIT_MAX_PARTS)
    ReDim gKtUnit(KIT_MAX_PARTS)
    ReDim gKtQty(KIT_MAX_PARTS)
    ReDim gKtEi(KIT_MAX_PARTS)
    For i = 0 To UBound(eis)
        n = eis(i)
        If n >= 1 Then d = st.getCellRangeByPosition(0, n, SC_STATE, n).getDataArray()(0)
        If n < 1 Then
            KitFromEis = "ERR:ЕИ-" & n & " нет в «" & SH_STOCK & "»"
            gUiLastMsg = KitFromEis
            Exit Function
        ElseIf CStr(d(SC_EI)) <> WmsIssue.EiCanon(n) Then
            KitFromEis = "ERR:ЕИ-" & n & " нет в «" & SH_STOCK & "»"
            gUiLastMsg = KitFromEis
            Exit Function
        End If
        For j = 0 To i - 1
            If eis(j) = n Then
                KitFromEis = "ERR:" & d(SC_EI) & " выбран дважды"
                gUiLastMsg = KitFromEis
                Exit Function
            End If
        Next j
        ' the quantity of one kit as an issue of this EI reads it («1 кан» of a packed EI is its part)
        If VarType(d(SC_QTY)) = 5 Then b = d(SC_QTY) Else b = 0
        If WmsPack.QtyEIText(CStr(q(i)), n, b, CStr(d(SC_UNIT)), "OUT", 0, x, info, msg) <> 0 Then
            KitFromEis = "ERR:" & d(SC_EI) & ": на 1 комплект «" & q(i) & "» — " & msg
            gUiLastMsg = KitFromEis
            Exit Function
        End If
        If x <= 0 Then
            KitFromEis = "ERR:" & d(SC_EI) & ": на 1 комплект — больше нуля"
            gUiLastMsg = KitFromEis
            Exit Function
        End If
        gKtName(i) = CStr(d(SC_NAME))
        gKtArt(i) = CStr(d(SC_EI))
        gKtUnit(i) = CStr(d(SC_UNIT))
        gKtQty(i) = x
        gKtEi(i) = n
        gKtN = gKtN + 1
    Next i
    gKtPreset = True
    res = KitBuild(Trim(sKit), nKits, sPlace, dt, sReason)
    gKtPreset = False
    If Left(res, 3) = "OK:" And bSave And gDoc.Sheets.hasByName(SH_KITS) Then
        ' the composition on «Комплекты»: the name in the first row, the EI in «Артикул», the quantity for one kit
        ks = gDoc.Sheets.getByName(SH_KITS)
        r = WmsOrders.LastRow(ks) + 1
        If r < 1 Then r = 1
        ReDim out(gKtN - 1)
        For i = 0 To gKtN - 1
            out(i) = Array(IIf(i = 0, Trim(sKit), ""), gKtName(i), gKtArt(i), gKtQty(i), gKtUnit(i), IIf(i = 0, "собран из ЕИ " & Format(dt, "DD.MM.YYYY"), ""))
        Next i
        ks.getCellRangeByPosition(0, r, KT_LAST, r + gKtN - 1).setDataArray(out())
        res = res & "; состав записан на лист «" & SH_KITS & "» (строки " & (r + 1) & "–" & (r + gKtN) & ")"
    End If
    KitFromEis = res
    gUiLastMsg = KitFromEis
    Exit Function
EH:
    gKtPreset = False
    KitFromEis = "ERR-SYS:внутренняя ошибка сборки из ЕИ: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
    gUiLastMsg = KitFromEis
End Function

' the record n of _ADJ of a build: kind KIT, the EI, the change, LIVE, the row, the places, the date, the EI of the kit
Sub KtAdjRow(n As Long, canon As String, delta As Double, r As Long, sFrom As String, sTo As String, dt As Double, canonK As String)
    PlanSetValue(SH_ADJ, n, AJ_NO, n, False)
    PlanSetValue(SH_ADJ, n, AJ_KIND, "KIT", False)
    PlanSetValue(SH_ADJ, n, AJ_EI, canon, False)
    PlanSetValue(SH_ADJ, n, AJ_QTY, delta, False)
    PlanSetValue(SH_ADJ, n, AJ_STATE, RV_LIVE, False)
    PlanSetValue(SH_ADJ, n, AJ_ROW, r, False)
    If sFrom <> "" Then PlanSetValue(SH_ADJ, n, AJ_FROM, sFrom, False)
    If sTo <> "" Then PlanSetValue(SH_ADJ, n, AJ_TO, sTo, False)
    PlanSetValue(SH_ADJ, n, AJ_DATE, dt, False)
    PlanSetValue(SH_ADJ, n, AJ_NEW, canonK, False)
End Sub

' a posted row «Сборка» of «Корректировки»: A … S as a posted correction has them
Sub KtSheetRow(r As Long, n As Long, canon As String, sName As String, sArt As String, sUnit As String, q As Double, sTo As String, dt As Double, _
    sReason As String, sFrom As String, b0 As Double, b1 As Double, sCtl As String, sBatch As String)
    PlanSetValue(SH_ADJUST, r, AC_NO, n, False)
    PlanSetValue(SH_ADJUST, r, AC_KIND, KIT_KIND_NAME, True)
    PlanSetValue(SH_ADJUST, r, AC_EI, canon, True)
    PlanSetValue(SH_ADJUST, r, AC_NAME, sName, True)
    If sArt <> "" Then PlanSetValue(SH_ADJUST, r, AC_ART, sArt, True)
    PlanSetValue(SH_ADJUST, r, AC_UNIT, sUnit, True)
    PlanSetValue(SH_ADJUST, r, AC_QTY, q, True)
    If sTo <> "" Then PlanSetValue(SH_ADJUST, r, AC_PLACE, sTo, True)
    PlanSetValue(SH_ADJUST, r, AC_DATE, dt, True)
    PlanSetValue(SH_ADJUST, r, AC_REASON, sReason, True)
    If sFrom <> "" Then PlanSetValue(SH_ADJUST, r, AC_FROM, sFrom, True)
    PlanSetValue(SH_ADJUST, r, AC_BEFORE, b0, True)
    PlanSetValue(SH_ADJUST, r, AC_AFTER, b1, True)
    PlanSetValue(SH_ADJUST, r, AC_DIFF, WmsIssue.Round3(b1 - b0), True)
    PlanSetValue(SH_ADJUST, r, AC_CTL, sCtl, True)
    PlanSetValue(SH_ADJUST, r, AC_BATCH, sBatch, True)
    PlanLockBits(SH_ADJUST, r, 0, AC_LAST, ADJUST_LOCKS_POSTED)
End Sub

' ================================================================ the kit taken apart (KIT_DEL)

' the records of _ADJ of the build of kit canonK (live ones), their numbers into nums(); the record of the kit itself (+N) in nKit
Private Function KtGroup(canonK As String, nums() As Long, ByRef nKit As Long) As Long
    Dim sh As Object, last As Long, desc As Object, found As Object, a As Variant, k As Long, r As Long, d As Variant, cnt As Long
    nKit = 0
    ReDim nums(KIT_MAX_PARTS + 1)
    sh = gDoc.Sheets.getByName(SH_ADJ)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    desc = sh.createSearchDescriptor()
    desc.SearchString = "^" & canonK & "$"
    desc.SearchRegularExpression = True
    desc.SearchCaseSensitive = True
    found = sh.getCellRangeByPosition(AJ_NEW, 1, AJ_NEW, last).findAll(desc)
    If IsNull(found) Then Exit Function
    a = found.getRangeAddresses()
    For k = 0 To UBound(a)
        For r = a(k).StartRow To a(k).EndRow
            d = sh.getCellRangeByPosition(0, r, AJ_LAST, r).getDataArray()(0)
            If CStr(d(AJ_KIND)) = "KIT" And CStr(d(AJ_STATE)) = RV_LIVE And VarType(d(AJ_FACT)) <> 5 And cnt <= KIT_MAX_PARTS + 1 Then
                nums(cnt) = r
                cnt = cnt + 1
                If CStr(d(AJ_EI)) = canonK Then nKit = r
            End If
        Next r
    Next k
    KtGroup = cnt
End Function

' "" when the kit of the build record rd may be taken apart now: its EI as the build made it — the whole quantity, no movement
Function KitDeleteProblem(rd As Variant) As String
    Dim canonK As String, m As Long, nums() As Long, nKit As Long, cnt As Long, bk As Double, d As Variant, nIss As Long, nRet As Long, nAdj As Long, nLive As Long
    ' 1.2.0: the build of a part (WmsAsm) is cancelled by its own rules
    If WmsAsm.AsmIsBuild(rd) Then
        KitDeleteProblem = WmsAsm.AsmDeleteProblem(rd)
        Exit Function
    End If
    canonK = CStr(rd(AJ_NEW))
    If Not WmsOrders.StrictEI(canonK, m) Then
        KitDeleteProblem = "ERR-SYS:у сборки № " & rd(AJ_NO) & " неверный ЕИ комплекта «" & canonK & "»"
        Exit Function
    End If
    cnt = KtGroup(canonK, nums, nKit)
    If cnt < 2 Or nKit = 0 Then
        KitDeleteProblem = "ERR-SYS:сборка комплекта " & canonK & " в " & SH_ADJ & " неполная (записей " & cnt & ") — нужна самопроверка"
        Exit Function
    End If
    d = WmsIssue.StockData(m)
    If Not WmsIssue.StockQty(m, bk) Then
        KitDeleteProblem = "ERR:остаток " & canonK & " в «" & SH_STOCK & "» не число — реестр повреждён"
        Exit Function
    End If
    If Abs(bk - gDoc.Sheets.getByName(SH_ADJ).getCellByPosition(AJ_QTY, nKit).getValue()) > 0.0000005 Then
        KitDeleteProblem = "ERR:у комплекта " & canonK & " остаток уже " & WmsIssue.QtyText(bk) & " — по нему были движения; комплект не разбирается. " _
            & "Сначала верните его (сторно выдач)"
        Exit Function
    End If
    If Not WmsReturn.LiveDependents(canonK, nIss, nRet) Or Not WmsAdjust.LiveAdjustments(canonK, nAdj) Or Not WmsSpecial.LiveLinesOfEI(canonK, nLive) Then
        KitDeleteProblem = "ERR-SYS:не удалось проверить движения комплекта " & canonK & " (формула движка Calc) — разборка не выполнена"
        Exit Function
    End If
    If nIss + nRet + (nAdj - 1) + nLive > 0 Then
        KitDeleteProblem = "ERR:у комплекта " & canonK & " есть действующие движения (выдач " & nIss & ", возвратов " & nRet & ", корректировок " & (nAdj - 1) _
            & ") — сначала их сторно"
        Exit Function
    End If
    KitDeleteProblem = ""
End Function

' the kit of the build record rd taken apart: every component back on its EI, the kit EI closed (0, its number not reused),
' every row of the build «Удалено (сторно)». One operation KIT_DEL. The result of ApplyOperation
Function KitDelete(rd As Variant) As String
    Dim canonK As String, m As Long, nums() As Long, nKit As Long, cnt As Long, i As Long, ad As Object, ash As Object, d As Variant, n As Long, e As Long
    Dim b0 As Double, b1 As Double, comps As String, r As Long, s As String, why As String
    If WmsAsm.AsmIsBuild(rd) Then
        KitDelete = WmsAsm.AsmDelete(rd)
        Exit Function
    End If
    why = KitDeleteProblem(rd)
    If why <> "" Then
        KitDelete = why
        Exit Function
    End If
    canonK = CStr(rd(AJ_NEW))
    WmsOrders.StrictEI(canonK, m)
    cnt = KtGroup(canonK, nums, nKit)
    ad = gDoc.Sheets.getByName(SH_ADJ)
    ash = WmsAdjust.AdjustSheet()
    PlanBegin("KIT_DEL", CStr(nums(0)))
    If Trim(gOpReason) <> "" Then PlanField("WHY", Trim(gOpReason))
    For i = 0 To cnt - 1
        n = nums(i)
        d = ad.getCellRangeByPosition(0, n, AJ_LAST, n).getDataArray()(0)
        r = WmsAdjust.AdjustRowOf(n, d(AJ_ROW))
        If r < 1 Then
            PlanBegin("", "")
            KitDelete = "ERR-SYS:строка корректировки № " & n & " (сборка " & canonK & ") не найдена на листе «" & SH_ADJUST & "»"
            Exit Function
        End If
        If n <> nKit Then
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
        If n = nKit Then PlanField("DATE", Format(d(AJ_DATE), "YYYY-MM-DD"))
    Next i
    PlanSetValue(SH_STOCK, m, SC_QTY, 0, False)
    PlanSetValue(SH_STOCK, m, SC_STATE, EI_ST_STORNO, False)
    PlanField("NEW_EI", canonK)
    PlanField("ADJ", nums(0))
    PlanField("ADJ_KIT", nKit)
    PlanField("N", ad.getCellByPosition(AJ_QTY, nKit).getValue())
    PlanField("COMPS", comps)
    KitDelete = ApplyOperation(0, 0)
End Function

' ================================================================ the price of a kit

' the price of one kit of EI canonK: the money of its components (their prices) ÷ the number of kits; −1 when a component has
' no price or canonK is not a kit
Function KitPrice(canonK As String, Optional depth As Variant) As Double
    Dim nums() As Long, nKit As Long, cnt As Long, i As Long, ad As Object, d As Variant, e As Long, p As Double, sum As Double, dp As Integer
    KitPrice = -1
    dp = 0
    If Not IsMissing(depth) Then dp = depth
    If dp > 5 Then Exit Function
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_ADJ) Then Exit Function
    cnt = KtGroup(canonK, nums, nKit)
    If cnt < 2 Or nKit = 0 Then Exit Function
    ad = gDoc.Sheets.getByName(SH_ADJ)
    For i = 0 To cnt - 1
        If nums(i) <> nKit Then
            d = ad.getCellRangeByPosition(0, nums(i), AJ_LAST, nums(i)).getDataArray()(0)
            If Not WmsOrders.StrictEI(CStr(d(AJ_EI)), e) Then Exit Function
            p = WmsPrint.EiPrice(e)
            If p < 0 Then p = KitPrice(CStr(d(AJ_EI)), dp + 1)
            If p < 0 Then Exit Function
            sum = sum - d(AJ_QTY) * p
        End If
    Next i
    If ad.getCellByPosition(AJ_QTY, nKit).getValue() > 0 Then KitPrice = sum / ad.getCellByPosition(AJ_QTY, nKit).getValue()
EH:
End Function
