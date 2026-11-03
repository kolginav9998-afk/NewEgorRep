' WmsCat — подсказки категорий по названию, справочник «Метизы» и список «Без категории» (1.0.2, расширено в 1.1).
' «Категории» — справочник кладовщика (не учёт, не журнал): категория, слова, по которым её узнают в названии («болт, гайка, шайба» → «Метизы»), и слова, при которых
' она не подходит («Не считать»). Вписали название на «Заказах» — пустая «Категория» заполняется подсказкой: сначала из «Номенклатуры» (если там у товара есть
' категория), затем по словам; побеждает категория, у которой нашлось больше слов (целое слово весит больше начала слова); при равенстве — та, что выше. Слово можно
' писать «=слово» (только целое) и «-слово» (если оно есть в названии — категория не подходит). Категория, вписанная вручную, подсказкой не трогается.
' «Метизы» — справочник метизов, собранный из «Наличия» по названию: тип, размер (М8×30), диаметр, длина, стандарт (DIN 933) и его аналоги (ISO 4017, ГОСТ 7805 — по листу
' «Стандарты»), класс прочности (8.8, А2), покрытие, головка и шлиц, остаток, места; поиск как на странице «Поиск», фильтры по каждому столбцу, пометка возможных дублей.
' «Без категории» — названия «Наличия», у которых нет категории ни в реестре, ни в «Номенклатуре», с подсказкой; принятая категория записывается только в «Номенклатуру»
' (в пустую клетку уже существующего товара). Только просмотр: ничего не проводится, в журнал не пишется. Library WmsPlus.
Option Explicit

Public Const SH_CATS = "Категории"
Private Const CT_NAME = 0
Private Const CT_WORDS = 1
Private Const CT_EXCL = 2
Private Const CT_NPOS = 3
Private Const CT_NEI = 4
Private Const CT_MAX = 300
Private Const CT_FASTENERS = "Метизы"

Public Const SH_FAST = "Метизы"
Public Const FS_FIND_ROW = 1
Private Const FS_FILTER_ROW = 2
Private Const FS_HEAD = 3
Private Const FS_ROW0 = 4
Private Const FS_COLS = 16
Private Const FS_KEY_COL = 16
Private Const FS_STAT_COL = 6
Private Const FS_MAX = 5000
' the columns of «Метизы»
Private Const FC_TYPE = 0
Private Const FC_SIZE = 1
Private Const FC_DIAM = 2
Private Const FC_LEN = 3
Private Const FC_STD = 4
Private Const FC_EQ = 5
Private Const FC_CLASS = 6
Private Const FC_COAT = 7
Private Const FC_HEAD = 8
Private Const FC_NAME = 9
Private Const FC_ART = 10
Private Const FC_QTY = 11
Private Const FC_UNIT = 12
Private Const FC_PLACES = 13
Private Const FC_NEI = 14
Private Const FC_NOTE = 15

Public Const SH_NOCAT = "Без категории"
Private Const NK_FIND_ROW = 1
Private Const NK_HEAD = 3
Private Const NK_ROW0 = 4
Private Const NK_COLS = 9
Private Const NK_MAX = 5000

' the categories as the sheet was read (a Global keeps it between the calls; the sheet's own event drops it)
Global gCtCat As Variant
Global gCtWords As Variant
Global gCtExcl As Variant
Global gCtN As Long
Global gCtLoaded As Boolean
Global gCtStatSeq As Long
' the journal position of the last list of «Метизы» and of «Без категории» (a list is made again when the journal moved)
Global gFsSeq As Long
Global gFsBusy As Boolean
Global gFsProf As String
Global gNkSeq As Long
Global gNkBusy As Boolean
' the registry grouped by name and article (StockGroups)
Global gGrp As Variant
Global gGrpN As Long
' «Номенклатура»: «a:article» / «n:name» → «row|category»
Global gGc As Object
' 1.3.0: the keys of gGc in one text («Chr(1) key Chr(1) key …»): a key that is not there is found by one search, not by the error of a Collection
Global gGcKeys As String
' 1.3.0: the category of every group of gGrp, made for all of them at once (GroupCats): gGrpCat(i) — as CategoryOf says it (the registry, «Номенклатура», the hint by the words), gGrpRow(i) — the row of the goods in
' «Номенклатура» (from 0; -1 — none), gGrpHint(i) — the hint by the words ("" when the category came from the registry or from «Номенклатуры»), gGrpNeed(i) — True when the group has no category
' in the registry and none in «Номенклатуре» (it has only the hint, which may be empty)
Global gGrpCat As Variant
Global gGrpRow As Variant
Global gGrpHint As Variant
Global gGrpNeed As Variant
' 1.3.0: the time (ms) of the steps of the last GroupCats (for the test of the speed)
Global gGrpProf As String

' the value of a key of a Collection; False when it has none (an error handler set inside a loop works for the first error only in this Basic: each lookup is a call of its own)
Private Function CtCollGet(coll As Object, key As String, ByRef v As Variant) As Boolean
    On Error GoTo Miss
    v = coll.Item(key)
    CtCollGet = True
    Exit Function
Miss:
    CtCollGet = False
End Function

' ================================================================ the hints of the categories

' a name or a list of words for the comparison: the signs between the words (, ; ( ) / " « » : |) are spaces, then WmsNames.Fold (lower case, «ё» = «е», the Latin letters that
' look like the Russian ones are Russian, «×» = «х»), one space between the words
Private Function CtName(ByVal s As String) As String
    Dim k As Variant
    For Each k In Array(",", ";", "(", ")", "/", Chr(34), "«", "»", ":", "|")
        s = Replace(s, CStr(k), " ")
    Next k
    CtName = WmsNames.Fold(s)
End Function

' the words of a cell «слово, =слово, -слово» → the words to find ("w" — the beginning of a word, "=w" — a whole word) and the words to exclude
Private Sub CtSplit(ByVal sCell As String, ByRef inc As String, ByRef exc As String)
    Dim a As Variant, i As Long, t As String
    a = Split(sCell, ",")
    For i = 0 To UBound(a)
        t = Trim(CStr(a(i)))
        If t <> "" Then
            If Left(t, 1) = "-" Then
                t = WmsNames.Fold(Mid(t, 2))
                If Len(t) >= 2 Then exc = exc & "|" & t
            ElseIf Left(t, 1) = "=" Then
                t = WmsNames.Fold(Mid(t, 2))
                If Len(t) >= 2 Then inc = inc & "|=" & t
            Else
                t = WmsNames.Fold(t)
                If Len(t) >= 2 Then inc = inc & "|" & t
            End If
        End If
    Next i
End Sub

Sub CtLoad()
    Dim sh As Object, last As Long, d As Variant, i As Long, n As Long, inc As String, exc As String, ex2 As String, w() As Variant, x() As Variant, cats() As String, useExcl As Boolean
    gCtN = 0
    gCtCat = Array()
    gCtWords = Array()
    gCtExcl = Array()
    gCtLoaded = True
    If Not gDoc.Sheets.hasByName(SH_CATS) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_CATS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    If last > CT_MAX Then last = CT_MAX
    d = sh.getCellRangeByPosition(0, 1, CT_EXCL, last).getDataArray()
    ' (the column C is the words to leave out only when it is headed so: a sheet of 1.0.2 has something else or nothing there)
    useExcl = (Left(Trim(sh.getCellByPosition(CT_EXCL, 0).getString()), 8) = "Не счита")
    ReDim cats(UBound(d))
    ReDim w(UBound(d))
    ReDim x(UBound(d))
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(CT_NAME))) <> "" Then
            inc = ""
            exc = ""
            ex2 = ""
            CtSplit(CStr(d(i)(CT_WORDS)), inc, exc)
            If useExcl Then CtSplit(CStr(d(i)(CT_EXCL)), ex2, exc)
            ' (a word of the column «Не считать» is a word to leave out whatever its sign)
            If ex2 <> "" Then exc = exc & Replace(ex2, "|=", "|")
            cats(n) = Trim(CStr(d(i)(CT_NAME)))
            w(n) = Split(Mid(inc & "|", 2, Len(inc)), "|")
            x(n) = Split(Mid(exc & "|", 2, Len(exc)), "|")
            n = n + 1
        End If
    Next i
    gCtCat = cats
    gCtWords = w
    gCtExcl = x
    gCtN = n
End Sub

' the sheet «Категории» was edited: read again at the next hint
Sub OnCatsChange(Optional oTarget As Variant)
    gCtLoaded = False
    gCtStatSeq = -1
End Sub

' the sheet is opened: the counting of the columns D, E is made again when the journal moved or the sheet was edited since
Sub OnCatsFocus(Optional oEvent As Variant)
    Dim seq As Long
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gCtStatSeq And gCtStatSeq > 0 Then Exit Sub
    CatsStat()
    Exit Sub
EH:
End Sub

' how well the name nm (" folded name ") fits the category i: 0 — no word, -1 — a word to leave out is there; a whole word 2, the beginning of a word 1
' 1.3.0: the folded text is compared as it is (InStr with the binary mode): the text mode of the language looks for every call through the whole text with the case rules of Unicode (ß = ss, ſ = s) and costs a millisecond on
' a long text; the text is lower case already (Fold), so only those rare signs differ
Private Function CtScore(nm As String, i As Long) As Long
    Dim w As Variant, j As Long, k As String, sc As Long
    w = gCtExcl(i)
    For j = 0 To UBound(w)
        k = CStr(w(j))
        If k <> "" Then
            If InStr(1, nm, " " & k, 0) > 0 Then
                CtScore = -1
                Exit Function
            End If
        End If
    Next j
    w = gCtWords(i)
    For j = 0 To UBound(w)
        k = CStr(w(j))
        If k <> "" Then
            If Left(k, 1) = "=" Then
                If InStr(1, nm, " " & Mid(k, 2) & " ", 0) > 0 Then sc = sc + 2
            ElseIf InStr(1, nm, " " & k & " ", 0) > 0 Then
                sc = sc + 2
            ElseIf InStr(1, nm, " " & k, 0) > 0 Then
                sc = sc + 1
            End If
        End If
    Next j
    CtScore = sc
End Function

' the category the name of a goods points at: the category with the most words found in the name (the first one, in the order of the sheet, when equal); "" — none. A word
' of the sheet is the beginning of a word of the name («болт» finds «болт», «болты», «болторез»; «ключ» does not find «выключатель»), «=слово» — a whole word only, «-слово» —
' the category does not fit when the name has it
Function CategorySuggest(sName As String) As String
    Dim nm As String, i As Long, best As Long, sc As Long
    If Not gCtLoaded Then CtLoad()
    ' 1.1: a name typed in several lines is one line («Винт» / «м3х25» / «Din 912»)
    nm = CtName(WmsNames.NormName(sName))
    If nm = "" Then Exit Function
    nm = " " & nm & " "
    For i = 0 To gCtN - 1
        sc = CtScore(nm, i)
        If sc > best Then
            best = sc
            CategorySuggest = gCtCat(i)
        End If
    Next i
End Function

' ---------------------------------------------------------------- «Номенклатура»: the category of a goods

Private Function CtGoodsKey(sName As String, sArt As String) As String
    If Trim(sArt) <> "" Then
        CtGoodsKey = "a:" & WmsNames.Fold(Trim(sArt))
    ElseIf Trim(sName) <> "" Then
        CtGoodsKey = "n:" & WmsNames.Fold(WmsNames.NormName(sName))
    End If
End Function

' «Номенклатура» read: a goods is its article (or its name when it has none); the value is «row|category» (the row of the sheet from 0); the first row of a key counts
Sub GoodsLoad()
    Dim sh As Object, last As Long, d As Variant, i As Long, k As String, cN As Long, cA As Long, cC As Long, r As Variant, ks() As String, nk As Long
    gGc = New Collection
    gGcKeys = ""
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_GOODS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    cN = GD_NAME
    cA = GD_ART
    cC = GD_CAT
    d = sh.getCellRangeByPosition(0, 1, cC, last).getDataArray()
    ReDim ks(UBound(d))
    For i = 0 To UBound(d)
        r = d(i)
        k = CtGoodsKey(CStr(r(cN)), CStr(r(cA)))
        If k <> "" Then
            On Error Resume Next
            gGc.Add((i + 1) & "|" & Trim(CStr(r(cC))), k)
            On Error GoTo 0
            ks(nk) = k
            nk = nk + 1
        End If
    Next i
    If nk > 0 Then
        ReDim Preserve ks(nk - 1)
        gGcKeys = Chr(1) & Join(ks, Chr(1)) & Chr(1)
    End If
End Sub

' the row of «Номенклатура» (from 0; -1 — none) and the category of the goods of this name and article ("" — none)
Private Function GoodsLookup(sName As String, sArt As String, ByRef nRow As Long) As String
    Dim k As String, v As String, p As Long
    nRow = -1
    k = CtGoodsKey(sName, sArt)
    If k = "" Then Exit Function
    On Error GoTo Miss
    v = gGc.Item(k)
    On Error GoTo 0
    p = InStr(v, "|")
    nRow = CLng(Left(v, p - 1))
    GoodsLookup = Mid(v, p + 1)
    Exit Function
Miss:
    On Error GoTo 0
    nRow = -1
End Function

' 1.3.0: a card of the registry is a part (the one definition for «Остаток деталей» and for the article on the label): its category is «Детали» («Деталь»), or it came by «Иной приход» of the type «Детали»
' (the source or the type of the source of the card says so). cat — the category, src — «Источник», stype — «Тип источника» of the card
Function IsPartCard(cat As String, src As String, stype As String) As Boolean
    Dim c As String
    c = LCase(Trim(cat))
    IsPartCard = (c = "детали" Or c = "деталь" Or LCase(Left(Trim(src), 6)) = "детали" Or LCase(Trim(stype)) = "детали")
End Function

' the category of a goods in the order: the registry (sRegCat), «Номенклатура», the hint by the words. «Номенклатура» must be read (GoodsLoad)
Function CategoryOf(sName As String, sArt As String, sRegCat As String) As String
    Dim nRow As Long, c As String
    If Trim(sRegCat) <> "" Then
        CategoryOf = Trim(sRegCat)
        Exit Function
    End If
    c = GoodsLookup(sName, sArt, nRow)
    If c <> "" Then
        CategoryOf = c
        Exit Function
    End If
    CategoryOf = CategorySuggest(sName)
End Function

' 1.3.0: the hints of the categories for many names at once — for each name the answer of CategorySuggest; found by the words, not by the names: every word of the sheet «Категории» is looked for once in the text
' that joins all the names (a name by itself costs about 2 ms: 4500 names are 10 s, this is a second). names — an array (from 0) of the names, n — how many
Private Function CtHintsAll(names As Variant, n As Long) As Variant
    Dim res() As String, prep() As String, off() As Long, i As Long, c As Long, j As Long, k As Long, w As Variant, sw As String, p As Long, tot As Long, big As String, parts As Variant, jt As String
    Dim sc() As Long, ex() As Boolean, cur As Long, cb As Long, best As Long, hint As String, nc As Long, pat As String, bFall As Boolean, sg As Variant, t0 As Long, tp As String
    t0 = GetSystemTicks()
    ReDim res(IIf(n > 0, n - 1, 0))
    If n <= 0 Then
        CtHintsAll = res
        Exit Function
    End If
    If Not gCtLoaded Then CtLoad()
    nc = gCtN
    If nc = 0 Then
        CtHintsAll = res
        Exit Function
    End If
    ReDim prep(n - 1)
    For i = 0 To n - 1
        prep(i) = WmsNames.NormName(CStr(names(i)))
        If InStr(prep(i), "¦") > 0 Then bFall = True
    Next i
    If Not bFall Then
        big = Join(prep, "¦")
        For Each sg In Array(",", ";", "(", ")", "/", Chr(34), "«", "»", ":", "|")
            big = Replace(big, CStr(sg), " ")
        Next sg
        parts = Split(WmsNames.Fold(big), "¦")
        If UBound(parts) <> n - 1 Then bFall = True
    End If
    If bFall Then
        ' a name holds the sign that joins them: one by one
        For i = 0 To n - 1
            res(i) = CategorySuggest(CStr(names(i)))
        Next i
        CtHintsAll = res
        Exit Function
    End If
    ReDim off(n)
    tot = 1
    For i = 0 To n - 1
        prep(i) = " " & Trim(CStr(parts(i))) & " "
        off(i) = tot
        tot = tot + Len(prep(i))
    Next i
    off(n) = tot
    jt = Join(prep, "")
    ReDim sc(nc * n - 1)
    ReDim ex(nc * n - 1)
    tp = "тексты " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For c = 0 To nc - 1
        cb = c * n
        ' the words to leave out: a word at the beginning of a word of the name
        w = gCtExcl(c)
        For k = 0 To UBound(w)
            sw = CStr(w(k))
            If sw <> "" Then
                pat = " " & sw
                cur = 0
                p = InStr(1, jt, pat, 0)
                Do While p > 0
                    cur = CtNameAt(off, n, p, cur)
                    ex(cb + cur) = True
                    p = InStr(off(cur + 1), jt, pat, 0)
                Loop
            End If
        Next k
        ' the words to find: the beginning of a word 1 and a whole word 1 more; «=word» — a whole word only, 2; one word counts once in a name (the search goes on from the next name)
        w = gCtWords(c)
        For k = 0 To UBound(w)
            sw = CStr(w(k))
            If sw <> "" Then
                If Left(sw, 1) = "=" Then
                    pat = " " & Mid(sw, 2) & " "
                    cur = 0
                    p = InStr(1, jt, pat, 0)
                    Do While p > 0
                        cur = CtNameAt(off, n, p, cur)
                        sc(cb + cur) = sc(cb + cur) + 2
                        p = InStr(off(cur + 1), jt, pat, 0)
                    Loop
                Else
                    pat = " " & sw
                    cur = 0
                    p = InStr(1, jt, pat, 0)
                    Do While p > 0
                        cur = CtNameAt(off, n, p, cur)
                        sc(cb + cur) = sc(cb + cur) + 1
                        p = InStr(off(cur + 1), jt, pat, 0)
                    Loop
                    pat = pat & " "
                    cur = 0
                    p = InStr(1, jt, pat, 0)
                    Do While p > 0
                        cur = CtNameAt(off, n, p, cur)
                        sc(cb + cur) = sc(cb + cur) + 1
                        p = InStr(off(cur + 1), jt, pat, 0)
                    Loop
                End If
            End If
        Next k
    Next c
    tp = tp & ", поиск слов " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For j = 0 To n - 1
        best = 0
        hint = ""
        For c = 0 To nc - 1
            If Not ex(c * n + j) Then
                If sc(c * n + j) > best Then
                    best = sc(c * n + j)
                    hint = gCtCat(c)
                End If
            End If
        Next c
        res(j) = hint
    Next j
    gGrpProf = tp & ", выбор " & (GetSystemTicks() - t0)
    CtHintsAll = res
End Function

' the number of the name (from 0) that holds the position p of the joined text: off(i) — where the name i begins, off(n) — the end. The hits of one search go up, so the search starts from the name of the
' previous hit (cur) and steps forward by 1, 2, 4 … names, then halves the last step
Private Function CtNameAt(off() As Long, n As Long, p As Long, cur As Long) As Long
    Dim lo As Long, hi As Long, st As Long, mid As Long
    If p < off(cur + 1) Then
        CtNameAt = cur
        Exit Function
    End If
    lo = cur + 1
    st = 1
    hi = lo + st
    Do While hi < n
        If off(hi) > p Then Exit Do
        lo = hi
        st = st * 2
        hi = lo + st
    Loop
    If hi > n Then hi = n
    Do While hi - lo > 1
        mid = (lo + hi) \ 2
        If off(mid) <= p Then
            lo = mid
        Else
            hi = mid
        End If
    Loop
    CtNameAt = lo
End Function

' the category of every group of gGrp, for all at once (after StockGroups; CtLoad and GoodsLoad are made here): gGrpCat, gGrpRow, gGrpHint — the same answers as CategoryOf and GoodsLookup give group by group
Sub GroupCats()
    Dim n As Long, i As Long, j As Long, g As Variant, cat() As String, rw() As Long, ht() As String, nm() As String, ar() As String, fa As Variant, fn As Variant, key As String, v As String, p As Long, c As String
    Dim idx() As Long, nNeed As Long, names() As String, hints As Variant, bGoods As Boolean, t0 As Long, tp As String, nd() As Boolean
    t0 = GetSystemTicks()
    gGrpProf = ""
    If Not gCtLoaded Then CtLoad()
    n = gGrpN
    If n = 0 Then
        gGrpCat = Array()
        gGrpRow = Array()
        gGrpHint = Array()
        gGrpNeed = Array()
        Exit Sub
    End If
    ReDim cat(n - 1)
    ReDim rw(n - 1)
    ReDim ht(n - 1)
    ReDim nd(n - 1)
    ReDim idx(n - 1)
    ReDim names(n - 1)
    bGoods = (gGcKeys <> "")
    If bGoods Then
        ReDim nm(n - 1)
        ReDim ar(n - 1)
        For i = 0 To n - 1
            g = gGrp(i)
            ar(i) = Trim(CStr(g(1)))
            ' (the key by the name is used only for a goods with no article, and only when the registry has no category of it)
            If ar(i) = "" And Trim(CStr(g(6))) = "" Then nm(i) = WmsNames.NormName(CStr(g(0)))
        Next i
        fa = WmsNames.FoldAll(ar)
        fn = WmsNames.FoldAll(nm)
    End If
    tp = "ключи " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 0 To n - 1
        g = gGrp(i)
        rw(i) = -1
        c = Trim(CStr(g(6)))
        If c <> "" Then
            cat(i) = c
        Else
            If bGoods Then
                key = ""
                If ar(i) <> "" Then
                    key = "a:" & CStr(fa(i))
                ElseIf Trim(CStr(g(0))) <> "" Then
                    key = "n:" & CStr(fn(i))
                End If
                If key <> "" Then
                    If InStr(gGcKeys, Chr(1) & key & Chr(1)) > 0 Then
                        v = gGc.Item(key)
                        p = InStr(v, "|")
                        rw(i) = CLng(Left(v, p - 1))
                        cat(i) = Mid(v, p + 1)
                    End If
                End If
            End If
            If cat(i) = "" Then
                nd(i) = True
                idx(nNeed) = i
                names(nNeed) = CStr(g(0))
                nNeed = nNeed + 1
            End If
        End If
    Next i
    tp = tp & ", «Номенклатура» " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    If nNeed > 0 Then
        hints = CtHintsAll(names, nNeed)
        For j = 0 To nNeed - 1
            ht(idx(j)) = CStr(hints(j))
            cat(idx(j)) = CStr(hints(j))
        Next j
    End If
    gGrpCat = cat
    gGrpRow = rw
    gGrpHint = ht
    gGrpNeed = nd
    gGrpProf = tp & ", подсказки " & (GetSystemTicks() - t0) & " [" & gGrpProf & "] (без категории " & nNeed & ")"
End Sub

' the category of every row of the registry array d (rows from 0; the columns of «Наличие»): its own, else of «Номенклатура», else the hint — one reading of a name, however many EIs
' have it. An array aligned with d ("" for an empty row). The search «категория:…» uses it
Function EffectiveCats(d As Variant, nRows As Long) As Variant
    Dim res() As String, cache As Object, i As Long, r As Variant, key As String, v As Variant, cEi As Long, cNm As Long, cArt As Long, cCat As Long, nm As String, c As String
    If Not gCtLoaded Then CtLoad()
    GoodsLoad()
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cCat = SC_CAT
    ReDim res(nRows)
    cache = New Collection
    For i = 0 To nRows - 1
        r = d(i)
        If CStr(r(cEi)) <> "" Then
            If Trim(CStr(r(cCat))) <> "" Then
                res(i) = Trim(CStr(r(cCat)))
            Else
                nm = CStr(r(cNm))
                key = LCase(nm) & "|" & LCase(Trim(CStr(r(cArt))))
                If CtCollGet(cache, key, v) Then
                    res(i) = CStr(v)
                Else
                    c = CategoryOf(nm, Trim(CStr(r(cArt))), "")
                    cache.Add(c, key)
                    res(i) = c
                End If
            End If
        End If
    Next i
    EffectiveCats = res
End Function

' the empty category of row r of sheet sSheet gets the hint of the name (nameCol) — the storekeeper may write his own over it. Called by
' the check of the row after a name was typed; never stops it. 1.1: the category of the goods in «Номенклатура» comes before the hint by the words
Sub CategoryHint(sSheet As String, r As Long, nameCol As Integer, catCol As Integer)
    Dim sh As Object, nm As String, c As String, nRow As Long
    On Error Resume Next
    sh = gDoc.Sheets.getByName(sSheet)
    If Trim(sh.getCellByPosition(catCol, r).getString()) <> "" Then Exit Sub
    nm = Trim(sh.getCellByPosition(nameCol, r).getString())
    If nm = "" Then Exit Sub
    GoodsLoad()
    c = GoodsLookup(nm, "", nRow)
    If c = "" Then c = CategorySuggest(nm)
    If c <> "" Then sh.getCellByPosition(catCol, r).setString(c)
End Sub

' 1.1: the empty article / unit / place of the row r of sSheet are filled from the last EI of the registry that has the same name (the letters that look alike not told apart): a hint of
' what was written the last time, never over what the storekeeper typed. artCol / placeCol = -1 — not filled (in «Иной приход» an article is the key of a part: a hint would make a new
' part an addition to the old one). Returns the note for «Контроль» ("" — nothing found or nothing filled)
Function PastHint(sSheet As String, r As Long, nameCol As Integer, artCol As Integer, unitCol As Integer, placeCol As Integer) As String
    Dim sh As Object, st As Object, last As Long, d As Variant, names() As String, i As Long, key As String, found As Long, nm As String, row As Variant, filled As String
    Dim cEi As Long, cNm As Long, cArt As Long, cUn As Long, cPl As Long, folded As Variant, needArt As Boolean, needPlace As Boolean
    On Error GoTo EH
    sh = gDoc.Sheets.getByName(sSheet)
    nm = Trim(sh.getCellByPosition(nameCol, r).getString())
    If nm = "" Then Exit Function
    ' (no short circuit in Basic: a column that is not used is not read)
    needArt = False
    needPlace = False
    If artCol >= 0 Then needArt = (Trim(sh.getCellByPosition(artCol, r).getString()) = "")
    If placeCol >= 0 Then needPlace = (Trim(sh.getCellByPosition(placeCol, r).getString()) = "")
    If Not needArt And Not needPlace And Trim(sh.getCellByPosition(unitCol, r).getString()) <> "" Then Exit Function
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    If last < 1 Then Exit Function
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cUn = SC_UNIT
    cPl = SC_PLACE
    d = st.getCellRangeByPosition(0, 1, SC_PLACE, last).getDataArray()
    ReDim names(UBound(d))
    For i = 0 To UBound(d)
        names(i) = CStr(d(i)(cNm))
    Next i
    folded = WmsNames.FoldAll(names)
    key = WmsNames.Fold(WmsNames.NormName(nm))
    found = -1
    For i = UBound(d) To 0 Step -1
        If CStr(d(i)(cEi)) <> "" Then
            If CStr(folded(i)) = key Then
                found = i
                Exit For
            End If
        End If
    Next i
    If found < 0 Then Exit Function
    row = d(found)
    If artCol >= 0 Then
        If Trim(sh.getCellByPosition(artCol, r).getString()) = "" And Trim(CStr(row(cArt))) <> "" Then
            sh.getCellByPosition(artCol, r).setString(Trim(CStr(row(cArt))))
            filled = filled & IIf(filled <> "", ", ", "") & "артикул"
        End If
    End If
    If Trim(sh.getCellByPosition(unitCol, r).getString()) = "" And Trim(CStr(row(cUn))) <> "" Then
        sh.getCellByPosition(unitCol, r).setString(Trim(CStr(row(cUn))))
        filled = filled & IIf(filled <> "", ", ", "") & "единица"
    End If
    If placeCol >= 0 Then
        If Trim(sh.getCellByPosition(placeCol, r).getString()) = "" And Trim(CStr(row(cPl))) <> "" Then
            sh.getCellByPosition(placeCol, r).setString(Trim(CStr(row(cPl))))
            filled = filled & IIf(filled <> "", ", ", "") & "место"
        End If
    End If
    If filled <> "" Then PastHint = "подставлено по прошлому приходу " & CStr(row(cEi)) & ": " & filled & " (можно изменить)"
    Exit Function
EH:
    PastHint = ""
End Function

' ================================================================ the registry by names

' the registry grouped by name and article: gGrp(i) = (name, article, unit, balance, EIs with a balance, places, category of the registry (the first one that is not empty)), gGrpN
Sub StockGroups()
    Dim st As Object, last As Long, d As Variant, i As Long, r As Variant, cEi As Long, cNm As Long, cArt As Long, cUn As Long, cQty As Long, cPl As Long, cCat As Long, q As Double
    Dim nm As String, art As String, key As String, k As Long, n As Long, coll As Object, pl As String
    Dim gN() As String, gA() As String, gU() As String, gB() As Double, gE() As Long, gP() As String, gC() As String, res() As Variant
    Dim c10 As String, c13 As String, c9 As String, c160 As String, kv As Variant
    gGrpN = 0
    gGrp = Array()
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    If last < 1 Then Exit Sub
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cUn = SC_UNIT
    cQty = SC_QTY
    cPl = SC_PLACE
    cCat = SC_CAT
    c10 = Chr(10)
    c13 = Chr(13)
    c9 = Chr(9)
    c160 = Chr(160)
    d = st.getCellRangeByPosition(0, 1, SC_CAT, last).getDataArray()
    ReDim gN(UBound(d))
    ReDim gA(UBound(d))
    ReDim gU(UBound(d))
    ReDim gB(UBound(d))
    ReDim gE(UBound(d))
    ReDim gP(UBound(d))
    ReDim gC(UBound(d))
    coll = New Collection
    For i = 0 To UBound(d)
        r = d(i)
        If CStr(r(cEi)) <> "" Then
            nm = CStr(r(cNm))
            If InStr(nm, c10) + InStr(nm, c13) + InStr(nm, c9) + InStr(nm, c160) + InStr(nm, "  ") > 0 Then nm = WmsNames.NormName(nm)
            art = Trim(CStr(r(cArt)))
            key = LCase(nm) & "|" & LCase(art)
            If CtCollGet(coll, key, kv) Then
                k = CLng(kv)
            Else
                k = n
                coll.Add(k, key)
                gN(k) = nm
                gA(k) = art
                gU(k) = Trim(CStr(r(cUn)))
                n = n + 1
            End If
            q = 0
            If VarType(r(cQty)) = 5 Then q = r(cQty)
            gB(k) = gB(k) + q
            If q > 0 Then gE(k) = gE(k) + 1
            pl = Trim(CStr(r(cPl)))
            If q > 0 And pl <> "" Then
                If InStr("; " & gP(k) & ";", "; " & pl & ";") = 0 Then gP(k) = gP(k) & IIf(gP(k) <> "", "; ", "") & pl
            End If
            If gC(k) = "" Then gC(k) = Trim(CStr(r(cCat)))
        End If
    Next i
    If n = 0 Then Exit Sub
    ReDim res(n - 1)
    For i = 0 To n - 1
        res(i) = Array(gN(i), gA(i), gU(i), gB(i), gE(i), gP(i), gC(i))
    Next i
    gGrp = res
    gGrpN = n
End Sub

' ================================================================ the statistics of «Категории»

' how many names (positions) and EIs with a balance each category of the sheet has, by the category of the registry, of «Номенклатура» or the hint; the two columns D, E of the sheet
' are written (they are the counting, not the input). OK:<categories>|<positions without a category>
Function CatsStat() As String
    Dim sh As Object, i As Long, j As Long, g As Variant, c As String, nPos() As Long, nEi() As Long, nNone As Long, last As Long, out() As Variant, key As String, names As Variant, keys() As String, fcat As Variant
    On Error GoTo EH
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_CATS) Then
        CatsStat = "ERR:в книге нет листа «" & SH_CATS & "»"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_CATS)
    ' (a sheet of 1.0.2 has no columns of the counting)
    If Trim(sh.getCellByPosition(CT_NPOS, 0).getString()) <> "Позиций" Then
        CatsStat = "OK:0|0"
        Exit Function
    End If
    CtLoad()
    GoodsLoad()
    StockGroups()
    GroupCats()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then last = 1
    If last > CT_MAX Then last = CT_MAX
    names = sh.getCellRangeByPosition(CT_NAME, 1, CT_NAME, last).getDataArray()
    ReDim keys(last)
    ReDim nPos(last)
    ReDim nEi(last)
    For j = 1 To last
        keys(j) = WmsNames.Fold(Trim(CStr(names(j - 1)(0))))
    Next j
    ' (the category of every group is made at once, GroupCats; the categories are folded at once)
    If gGrpN > 0 Then fcat = WmsNames.FoldAll(gGrpCat)
    For i = 0 To gGrpN - 1
        g = gGrp(i)
        c = CStr(gGrpCat(i))
        If c = "" Then
            nNone = nNone + 1
        Else
            key = CStr(fcat(i))
            For j = 1 To last
                If keys(j) = key Then
                    nPos(j) = nPos(j) + 1
                    nEi(j) = nEi(j) + g(4)
                    Exit For
                End If
            Next j
        End If
    Next i
    ReDim out(last - 1)
    For j = 1 To last
        If keys(j) <> "" Then
            out(j - 1) = Array(nPos(j), nEi(j))
        Else
            out(j - 1) = Array("", "")
        End If
    Next j
    sh.getCellRangeByPosition(CT_NPOS, 1, CT_NEI, last).setDataArray(out)
    gCtStatSeq = CLng(SysNum(SK_LAST_SEQ))
    If gCtStatSeq = 0 Then gCtStatSeq = 1
    CatsStat = "OK:" & last & "|" & nNone
    gUiLastMsg = CatsStat
    Exit Function
EH:
    CatsStat = "ERR:подсчёт по категориям не выполнен: " & Error$ & " (строка " & Erl & ")"
End Function

' ================================================================ «Метизы»

Sub BtnFastRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = FastRefresh(FastFind())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the words typed in the cell «Найти» of the sheet
Function FastFind() As String
    If gDoc.Sheets.hasByName(SH_FAST) Then FastFind = Trim(gDoc.Sheets.getByName(SH_FAST).getCellByPosition(1, FS_FIND_ROW).getString())
End Function

' a cell of the search or of the filters was changed: the list is made again
Sub OnFastChange(Optional oTarget As Variant)
    Dim ra As Object
    If gFsBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If Not ((ra.Row = FS_FIND_ROW And (ra.Column = 1 Or ra.Column = 5)) Or (ra.Row = FS_FILTER_ROW And ra.Column < FS_COLS)) Then Exit Sub
    WmsInit()
    FastRefresh(FastFind())
    Exit Sub
EH:
End Sub

' the sheet is opened: the list is made again when the journal moved since it was made (a new book — the first time)
Sub OnFastFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gFsBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gFsSeq And gFsSeq > 0 Then Exit Sub
    FastRefresh(FastFind())
    Exit Sub
EH:
End Sub

' a number filter: «8», «8 10 12» (any of them), «8-12» (from–to), «>20», «>=20», «<30», «<=30», «=0» (with signs all must hold); empty — all
Function NumFilterOk(ByVal sExpr As String, v As Double) As Boolean
    Dim a As Variant, i As Long, t As String, hasOp As Boolean, anyOk As Boolean, allOk As Boolean, ok As Boolean, p As Long, x As Double, y As Double
    sExpr = Trim(Replace(Replace(sExpr, ";", " "), ",", "."))
    If sExpr = "" Then
        NumFilterOk = True
        Exit Function
    End If
    Do While InStr(sExpr, "  ") > 0
        sExpr = Replace(sExpr, "  ", " ")
    Loop
    a = Split(sExpr, " ")
    allOk = True
    For i = 0 To UBound(a)
        t = Trim(CStr(a(i)))
        ok = False
        If t <> "" Then
            If Left(t, 2) = ">=" Then
                hasOp = True
                ok = (v >= Val(Mid(t, 3)) - 0.0000001)
                If Not ok Then allOk = False
            ElseIf Left(t, 2) = "<=" Then
                hasOp = True
                ok = (v <= Val(Mid(t, 3)) + 0.0000001)
                If Not ok Then allOk = False
            ElseIf Left(t, 1) = ">" Then
                hasOp = True
                ok = (v > Val(Mid(t, 2)))
                If Not ok Then allOk = False
            ElseIf Left(t, 1) = "<" Then
                hasOp = True
                ok = (v < Val(Mid(t, 2)))
                If Not ok Then allOk = False
            ElseIf Left(t, 1) = "=" Then
                hasOp = True
                ok = (Abs(v - Val(Mid(t, 2))) < 0.0000001)
                If Not ok Then allOk = False
            Else
                p = InStr(2, t, "-")
                If p > 1 Then
                    x = Val(Left(t, p - 1))
                    y = Val(Mid(t, p + 1))
                    ok = (v >= x - 0.0000001 And v <= y + 0.0000001)
                Else
                    ok = (Abs(v - Val(t)) < 0.0000001)
                End If
                If ok Then anyOk = True
            End If
        End If
    Next i
    If hasOp Then
        NumFilterOk = allOk And (anyOk Or Not HasPlain(a))
    Else
        NumFilterOk = anyOk
    End If
End Function

' are there plain numbers (not signs) among the items of a number filter
Private Function HasPlain(a As Variant) As Boolean
    Dim i As Long, t As String
    For i = 0 To UBound(a)
        t = Trim(CStr(a(i)))
        If t <> "" Then
            If InStr("<>=", Left(t, 1)) = 0 Then HasPlain = True
        End If
    Next i
End Function

' a text filter: every word of sExpr is in the folded text hay
Function TextFilterOk(sExpr As String, hay As String) As Boolean
    Dim w As Variant
    TextFilterOk = True
    For Each w In WmsNames.FoldWords(sExpr)
        If CStr(w) <> "" Then
            If InStr(hay, CStr(w)) = 0 Then
                TextFilterOk = False
                Exit Function
            End If
        End If
    Next w
End Function

' the list of the fasteners from «Наличие»: one row for a name with an article — the balance of all its EIs, the places, how many EIs
' have a balance; sorted by type, diameter, length, name. sFind — the words of the search (as on «Поиск»: the stem, the Latin letters, «din 933» = «ISO 4017», «-слово»);
' under it the filters of the columns; «Скрыть нулевые» leaves out the names with no balance. OK:<rows> / ERR:…
Function FastRefresh(sFind As String) As String
    Dim sh As Object, i As Long, j As Long, k As Long, n As Long, g As Variant, rec As Variant, rows() As Variant, hays() As String, nms() As String, cols As Variant, notes() As String
    Dim canon() As String, filt As Variant, fTxt(15) As String, hide As Boolean, scores As Variant, ok As Boolean, out() As Variant, nShown As Long, nAllF As Long, q As Double
    Dim eqShow As String, hay As String, key As String, pk As Long, ra As Object, descr As String, tIdx As Long, typeName As String, fl As Boolean, dd As Variant
    Dim coll As Object, wordsOld As Variant, cc As String, fcat As Variant, rawH() As String, rawN() As String, rawM() As String, hP As Variant, nP As Variant, nmF As Variant, t0 As Long, tp As String, rawO() As String, oldF As Variant
    On Error GoTo EH
    gFsBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_FAST) Then
        FastRefresh = "ERR:в книге нет листа «" & SH_FAST & "» — книга не обновлена до 1.0.2"
        gFsBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_FAST)
    If Not gCtLoaded Then CtLoad()
    GoodsLoad()
    t0 = GetSystemTicks()
    StockGroups()
    tp = "реестр " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' 1.3.0: the category of every group at once (registry, «Номенклатура», the hint by the words) — not group by group
    GroupCats()
    tp = tp & ", категории " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the fasteners among the groups: their category says so («Метизы», «Крепёж»); a category of something else («Инструмент» for «Болторез») keeps a group out
    n = 0
    ReDim rows(IIf(gGrpN > 0, gGrpN - 1, 0))
    If gGrpN > 0 Then fcat = WmsNames.FoldAll(gGrpCat)
    For i = 0 To gGrpN - 1
        cc = CStr(fcat(i))
        If InStr(cc, "метиз") > 0 Or InStr(cc, "крепеж") > 0 Then
            g = gGrp(i)
            rec = WmsSpec.SpRecordC(CStr(g(0)))
            rows(n) = Array(g, rec, i)
            n = n + 1
        End If
    Next i
    tp = tp & ", разбор " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the search and the filters of the sheet
    filt = sh.getCellRangeByPosition(0, FS_FILTER_ROW, FS_COLS - 1, FS_FILTER_ROW).getDataArray()(0)
    For j = 0 To FS_COLS - 1
        fTxt(j) = Trim(CStr(filt(j)))
    Next j
    hide = (WmsNames.Fold(Trim(sh.getCellByPosition(5, FS_FIND_ROW).getString())) = "да")
    ReDim hays(IIf(n > 0, n - 1, 0))
    ReDim nms(IIf(n > 0, n - 1, 0))
    ReDim notes(IIf(n > 0, n - 1, 0))
    ReDim canon(IIf(n > 0, n - 1, 0))
    ' 1.3.0: the texts of all the rows are folded at once (the engine of the language works on one long text, not on every row)
    If n > 0 Then
        ReDim rawH(n - 1)
        ReDim rawN(n - 1)
        ReDim rawM(n - 1)
        For i = 0 To n - 1
            g = rows(i)(0)
            rec = rows(i)(1)
            typeName = CStr(rec(0))
            If typeName = "" Then typeName = "Прочее"
            rawH(i) = typeName & " " & rec(2) & " " & WmsSpec.SpTokens(rec) & " " & rec(7) & " " & rec(8) & " " & rec(9) & " " & g(0) & " " & g(1) & " " & g(5)
            rawN(i) = CStr(g(0)) & " " & g(1)
            rawM(i) = CStr(g(0))
        Next i
        hP = WmsSearch.PrepAll(rawH)
        nP = WmsSearch.PrepAll(rawN)
        nmF = WmsNames.FoldAll(rawM)
        For i = 0 To n - 1
            hays(i) = " " & Trim(CStr(hP(i))) & " "
            nms(i) = " " & Trim(CStr(nP(i))) & " "
        Next i
    End If
    tp = tp & ", тексты поиска " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the possible doubles: the same type, size, standard (or its equivalent), class, coating, head — under different names; or the same name — under different articles
    coll = New Collection
    For i = 0 To n - 1
        rec = rows(i)(1)
        g = rows(i)(0)
        If CStr(rec(0)) <> "" And CStr(rec(2)) <> "" Then
            ' (a zinc coating is «Цинк» or «Оцинкованный» — one thing)
            cc = CStr(rec(8))
            If cc = "Оцинкованный" Then cc = "Цинк"
            key = "p|" & rec(0) & "|" & rec(2) & "|" & rec(11) & "|" & rec(7) & "|" & cc & "|" & rec(9)
            DupMark(coll, key, i, notes, rows)
        End If
        key = "n|" & CStr(nmF(i))
        DupMark(coll, key, i, notes, rows)
    Next i
    tp = tp & ", дубли " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the search with the understanding of «Поиск»; the filters «место:», «категория:», «артикул:» of the query need the columns of the rows
    cols = Empty
    If Mid(WmsSearch.QueryNeeds(sFind), 1, 4) <> "0000" Then
        Dim aPl() As String, aCt() As String, aAr() As String, aSr() As String
        ReDim aPl(IIf(n > 0, n - 1, 0))
        ReDim aCt(IIf(n > 0, n - 1, 0))
        ReDim aAr(IIf(n > 0, n - 1, 0))
        ReDim aSr(IIf(n > 0, n - 1, 0))
        For i = 0 To n - 1
            g = rows(i)(0)
            aPl(i) = CStr(g(5))
            aAr(i) = CStr(g(1))
            aCt(i) = CStr(gGrpCat(CLng(rows(i)(2))))
        Next i
        cols = Array(WmsNames.FoldAll(aPl), WmsNames.FoldAll(aCt), WmsNames.FoldAll(aAr), WmsNames.FoldAll(aSr))
    End If
    tp = tp & ", столбцы запроса " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    scores = WmsSearch.MatchScores(sFind, hays, nms, n, cols)
    wordsOld = WmsNames.FoldWords(sFind)
    ' the old way too: every word of the search as a part of the text of the row (a part of an article, of a place: «b-8», «5-1»); the texts folded at once, only for a search
    If Trim(sFind) <> "" And n > 0 Then
        ReDim rawO(n - 1)
        For i = 0 To n - 1
            g = rows(i)(0)
            rec = rows(i)(1)
            rawO(i) = CStr(g(0)) & " " & g(1) & " " & g(5) & " " & rec(0) & " " & rec(2) & " " & rec(6) & " " & rec(7) & " " & rec(8)
        Next i
        oldF = WmsNames.FoldAll(rawO)
    End If
    tp = tp & ", поиск " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ReDim out(IIf(n > 0, n - 1, 0))
    k = 0
    For i = 0 To n - 1
        g = rows(i)(0)
        rec = rows(i)(1)
        q = g(3)
        ok = (scores(i) > 0)
        If Not ok And Trim(sFind) <> "" Then ok = WmsNames.HasWords(CStr(oldF(i)), wordsOld)
        If ok And hide And q < 0.0000005 Then ok = False
        If ok Then
            typeName = CStr(rec(0))
            If typeName = "" Then typeName = "Прочее"
            tIdx = IIf(CStr(rec(0)) = "", 99, CLng(rec(1)))
            eqShow = UCase(Replace(CStr(rec(10)), ";", "; "))
            ' the filters of the columns
            If fTxt(FC_TYPE) <> "" Then ok = ok And TextFilterOk(fTxt(FC_TYPE), WmsNames.Fold(typeName))
            If fTxt(FC_SIZE) <> "" Then ok = ok And TextFilterOk(fTxt(FC_SIZE), WmsNames.Fold(CStr(rec(2))))
            If fTxt(FC_DIAM) <> "" Then ok = ok And NumFilterOk(fTxt(FC_DIAM), CDbl(rec(3)))
            If fTxt(FC_LEN) <> "" Then ok = ok And NumFilterOk(fTxt(FC_LEN), CDbl(rec(4)))
            If fTxt(FC_STD) <> "" Then ok = ok And TextFilterOk(fTxt(FC_STD), WmsNames.Fold(rec(6) & " " & Replace(CStr(rec(10)), ";", " ")))
            If fTxt(FC_EQ) <> "" Then ok = ok And TextFilterOk(fTxt(FC_EQ), WmsNames.Fold(eqShow))
            If fTxt(FC_CLASS) <> "" Then ok = ok And TextFilterOk(fTxt(FC_CLASS), WmsNames.Fold(CStr(rec(7))))
            If fTxt(FC_COAT) <> "" Then ok = ok And TextFilterOk(fTxt(FC_COAT), WmsNames.Fold(CStr(rec(8))))
            If fTxt(FC_HEAD) <> "" Then ok = ok And TextFilterOk(fTxt(FC_HEAD), WmsNames.Fold(CStr(rec(9))))
            If fTxt(FC_NAME) <> "" Then ok = ok And TextFilterOk(fTxt(FC_NAME), WmsNames.Fold(CStr(g(0))))
            If fTxt(FC_ART) <> "" Then ok = ok And TextFilterOk(fTxt(FC_ART), WmsNames.Fold(CStr(g(1))))
            If fTxt(FC_QTY) <> "" Then ok = ok And NumFilterOk(fTxt(FC_QTY), q)
            If fTxt(FC_UNIT) <> "" Then ok = ok And TextFilterOk(fTxt(FC_UNIT), WmsNames.Fold(CStr(g(2))))
            If fTxt(FC_PLACES) <> "" Then ok = ok And TextFilterOk(fTxt(FC_PLACES), WmsNames.Fold(CStr(g(5))))
            If fTxt(FC_NEI) <> "" Then ok = ok And NumFilterOk(fTxt(FC_NEI), CDbl(g(4)))
            If fTxt(FC_NOTE) <> "" Then ok = ok And TextFilterOk(fTxt(FC_NOTE), WmsNames.Fold(notes(i)))
        End If
        If ok Then
            out(k) = Array(typeName, rec(2), IIf(rec(3) > 0, rec(3), ""), IIf(rec(4) > 0, rec(4), ""), rec(6), eqShow, rec(7), rec(8), rec(9), g(0), g(1), WmsPack.PkRound(q, 3), g(2), g(5), g(4), notes(i), _
                CDbl(tIdx) * 1000000000000# + CDbl(Int(rec(3) * 100 + 0.5)) * 1000000# + CDbl(Int(rec(4) * 10 + 0.5)))
            k = k + 1
        End If
    Next i
    tp = tp & ", фильтры " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the old list out, the new one in
    sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + FS_MAX).clearContents(1 + 2 + 4 + 16)
    nAllF = n
    If k > FS_MAX Then k = FS_MAX
    If k > 0 Then
        ReDim Preserve out(k - 1)
        sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + k - 1).setDataArray(out)
        FsSort(sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + k - 1))
        sh.getCellRangeByPosition(0, FS_ROW0, FS_COLS - 1, FS_ROW0 + k - 1).CharColor = RGB(38, 50, 56)
    End If
    tp = tp & ", запись " & (GetSystemTicks() - t0)
    gFsProf = tp & " (групп " & gGrpN & ", метизов " & n & ")"
    descr = ""
    If Trim(sFind) <> "" Then descr = " — по словам «" & Trim(sFind) & "»"
    For j = 0 To FS_COLS - 1
        If fTxt(j) <> "" Then
            descr = descr & IIf(descr <> "", "; ", " — ") & "фильтры столбцов"
            Exit For
        End If
    Next j
    sh.getCellByPosition(FS_STAT_COL, FS_FIND_ROW).setString("Метизов (позиций): " & k & IIf(k < nAllF, " из " & nAllF, "") & descr & ". Список из «Наличия» на " & Format(Now(), "DD.MM.YYYY HH:MM") _
        & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)) & BadNote())
    gFsSeq = CLng(SysNum(SK_LAST_SEQ))
    If gFsSeq = 0 Then gFsSeq = 1
    gFsBusy = False
    FastRefresh = "OK:" & k
    gUiLastMsg = FastRefresh
    Exit Function
EH:
    gFsBusy = False
    FastRefresh = "ERR:список метизов не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = FastRefresh
End Function

' 1.2.1: the names of the groups on which a rule of the fasteners failed (WmsSpec.SpRecordC answered them by an empty record): the note for the status line — how many and the first
Private Function BadNote() As String
    Dim i As Long, nBad As Long, sBad As String, pb As Long, qb As Long
    If gSpBadNames = "" Then Exit Function
    For i = 0 To gGrpN - 1
        pb = InStr(gSpBadNames, Chr(1) & CStr(gGrp(i)(0)) & Chr(2))
        If pb > 0 Then
            nBad = nBad + 1
            If sBad = "" Then
                qb = InStr(pb + 1, gSpBadNames, Chr(2)) + 1
                sBad = "«" & Left(CStr(gGrp(i)(0)), 60) & "» — " & Mid(gSpBadNames, qb, InStr(qb, gSpBadNames, Chr(1)) - qb)
            End If
        End If
    Next i
    If nBad > 0 Then BadNote = ". НЕ ПРОЧИТАНО названий: " & nBad & " (первое: " & sBad & ") — их нет в списке"
End Function

' marks a possible double: the key was met in a row before — both rows get a note
Sub DupMark(coll As Object, key As String, i As Long, notes() As String, rows() As Variant)
    Dim j As Long, kind As String, why As String
    On Error GoTo First
    j = coll.Item(key)
    On Error GoTo 0
    If Left(key, 2) = "n|" Then
        why = "одно название — разные артикулы"
        If CStr(rows(i)(0)(1)) = CStr(rows(j)(0)(1)) Then Exit Sub
    Else
        why = "то же самое, что: "
        If CStr(rows(i)(0)(0)) = CStr(rows(j)(0)(0)) And CStr(rows(i)(0)(1)) = CStr(rows(j)(0)(1)) Then Exit Sub
    End If
    If Left(key, 2) = "n|" Then
        CtAddNote(notes, i, "возможный дубль: " & why & " (арт. " & rows(j)(0)(1) & ")")
        CtAddNote(notes, j, "возможный дубль: " & why & " (арт. " & rows(i)(0)(1) & ")")
    Else
        CtAddNote(notes, i, "возможный дубль: " & why & rows(j)(0)(0) & IIf(CStr(rows(j)(0)(1)) <> "", " (арт. " & rows(j)(0)(1) & ")", ""))
        CtAddNote(notes, j, "возможный дубль: " & why & rows(i)(0)(0) & IIf(CStr(rows(i)(0)(1)) <> "", " (арт. " & rows(i)(0)(1) & ")", ""))
    End If
    Exit Sub
First:
    On Error GoTo 0
    coll.Add(i, key)
End Sub

Sub CtAddNote(notes() As String, i As Long, s As String)
    If InStr(notes(i), s) = 0 Then notes(i) = notes(i) & IIf(notes(i) <> "", "; ", "") & s
End Sub

' the list sorted by the key of the type, the diameter and the length (the hidden column), then by the name
Private Sub FsSort(rng As Object)
    Dim f(1) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = FS_KEY_COL
    f(0).IsAscending = True
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    f(1).Field = FC_NAME
    f(1).IsAscending = True
    f(1).IsCaseSensitive = False
    f(1).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(1).CollatorLocale = loc
    pv(0).Name = "SortFields"
    pv(0).Value = f()
    pv(1).Name = "ContainsHeader"
    pv(1).Value = False
    rng.sort(pv())
End Sub

' ================================================================ «Без категории»

Sub BtnNoCatRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = NoCatRefresh(NoCatFind())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Function NoCatFind() As String
    If gDoc.Sheets.hasByName(SH_NOCAT) Then NoCatFind = Trim(gDoc.Sheets.getByName(SH_NOCAT).getCellByPosition(1, NK_FIND_ROW).getString())
End Function

Sub OnNoCatChange(Optional oTarget As Variant)
    Dim ra As Object
    If gNkBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> NK_FIND_ROW Or ra.Column <> 1 Then Exit Sub
    WmsInit()
    NoCatRefresh(NoCatFind())
    Exit Sub
EH:
End Sub

Sub OnNoCatFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gNkBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gNkSeq And gNkSeq > 0 Then Exit Sub
    NoCatRefresh(NoCatFind())
    Exit Sub
EH:
End Sub

' the names of «Наличие» with no category — neither in the registry, nor in «Номенклатуре» — and the hint by the words for each; the column «Категория» is filled with the hint:
' the storekeeper changes it or clears it, «Записать в «Номенклатуру»» takes what stays. OK:<rows> / ERR:…
Function NoCatRefresh(sFind As String) As String
    Dim sh As Object, i As Long, j As Long, k As Long, g As Variant, nRow As Long, hint As String, out() As Variant, words As Variant, w As Variant, ok As Boolean, nNoGoods As Long
    Dim cand() As Long, nc As Long, raw() As String, hayF As Variant, bSearch As Boolean
    On Error GoTo EH
    gNkBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_NOCAT) Then
        NoCatRefresh = "ERR:в книге нет листа «" & SH_NOCAT & "» — книга не обновлена до 1.1"
        gNkBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_NOCAT)
    If Not gCtLoaded Then CtLoad()
    GoodsLoad()
    StockGroups()
    ' 1.3.0: the category of every group at once; the groups with no category in the registry and none in «Номенклатуре» are the candidates, the hint is the one that GroupCats made
    GroupCats()
    words = WmsNames.FoldWords(sFind)
    bSearch = (Trim(sFind) <> "")
    ReDim out(IIf(gGrpN > 0, gGrpN - 1, 0))
    ReDim cand(IIf(gGrpN > 0, gGrpN - 1, 0))
    For i = 0 To gGrpN - 1
        If gGrpNeed(i) Then
            cand(nc) = i
            nc = nc + 1
        End If
    Next i
    If bSearch And nc > 0 Then
        ReDim raw(nc - 1)
        For j = 0 To nc - 1
            g = gGrp(cand(j))
            raw(j) = CStr(g(0)) & " " & g(1)
        Next j
        hayF = WmsNames.FoldAll(raw)
    End If
    For j = 0 To nc - 1
        i = cand(j)
        g = gGrp(i)
        ok = True
        If bSearch Then
            For Each w In words
                If CStr(w) <> "" Then
                    If InStr(CStr(hayF(j)), CStr(w)) = 0 Then ok = False
                End If
            Next w
        End If
        If ok Then
            hint = CStr(gGrpHint(i))
            nRow = CLng(gGrpRow(i))
            If nRow < 0 Then nNoGoods = nNoGoods + 1
            out(k) = Array(g(0), g(1), g(4), WmsPack.PkRound(CDbl(g(3)), 3), g(2), hint, hint, IIf(nRow >= 0, "есть, строка " & (nRow + 1), "нет в «Номенклатуре»"), nRow)
            k = k + 1
        End If
    Next j
    sh.getCellRangeByPosition(0, NK_ROW0, NK_COLS - 1, NK_ROW0 + NK_MAX).clearContents(1 + 2 + 4 + 16)
    If k > NK_MAX Then k = NK_MAX
    If k > 0 Then
        ReDim Preserve out(k - 1)
        sh.getCellRangeByPosition(0, NK_ROW0, NK_COLS - 1, NK_ROW0 + k - 1).setDataArray(out)
        NkSort(sh.getCellRangeByPosition(0, NK_ROW0, NK_COLS - 1, NK_ROW0 + k - 1))
    End If
    sh.getCellByPosition(0, NK_FIND_ROW + 1).setString("Названий без категории: " & k & IIf(Trim(sFind) <> "", " — по словам «" & Trim(sFind) & "»", "") & IIf(nNoGoods > 0, "; из них нет в «Номенклатуре»: " & nNoGoods, "") _
        & ". Список из «Наличия» на " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
    gNkSeq = CLng(SysNum(SK_LAST_SEQ))
    If gNkSeq = 0 Then gNkSeq = 1
    gNkBusy = False
    NoCatRefresh = "OK:" & k
    gUiLastMsg = NoCatRefresh
    Exit Function
EH:
    gNkBusy = False
    NoCatRefresh = "ERR:список «Без категории» не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = NoCatRefresh
End Function

Private Sub NkSort(rng As Object)
    Dim f(0) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = 0
    f(0).IsAscending = True
    f(0).IsCaseSensitive = False
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    pv(0).Name = "SortFields"
    pv(0).Value = f()
    pv(1).Name = "ContainsHeader"
    pv(1).Value = False
    rng.sort(pv())
End Sub

Sub BtnNoCatApply(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = NoCatApply()
    If Left(s, 4) = "ERR:" Then
        WmsUi.UiMessage(Mid(s, 5))
    ElseIf Left(s, 3) = "OK:" Then
        WmsUi.UiMessage(Mid(s, 4))
    End If
End Sub

' the categories of the column «Категория» go into the empty category of the goods in «Номенклатура» (an existing row; nothing that is written there is replaced; goods that
' are not in «Номенклатуре» are only counted). Does not post anything. OK:<text> / ERR:…
Function NoCatApply() As String
    Dim sh As Object, gs As Object, last As Long, d As Variant, i As Long, nWritten As Long, nNoGoods As Long, nSkip As Long, nRow As Long, cat As String, nm As String, art As String
    Dim cur As String, gcN As Long, gcA As Long, gcC As Long, r As Variant
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_NOCAT) Then
        NoCatApply = "ERR:в книге нет листа «" & SH_NOCAT & "»"
        Exit Function
    End If
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then
        NoCatApply = "ERR:в книге нет листа «" & SH_GOODS & "»"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_NOCAT)
    gs = gDoc.Sheets.getByName(SH_GOODS)
    last = WmsOrders.LastRow(sh)
    If last < NK_ROW0 Then
        NoCatApply = "OK:В списке нет названий — записывать нечего."
        Exit Function
    End If
    gcN = GD_NAME
    gcA = GD_ART
    gcC = GD_CAT
    d = sh.getCellRangeByPosition(0, NK_ROW0, NK_COLS - 1, last).getDataArray()
    For i = 0 To UBound(d)
        r = d(i)
        cat = Trim(CStr(r(6)))
        If cat <> "" And Trim(CStr(r(0))) <> "" Then
            If VarType(r(8)) = 5 Then
                nRow = CLng(r(8))
            Else
                nRow = -1
            End If
            If nRow < 0 Then
                nNoGoods = nNoGoods + 1
            Else
                ' the row of «Номенклатуры» must still be the same goods and its category still empty
                nm = Trim(gs.getCellByPosition(gcN, nRow).getString())
                art = Trim(gs.getCellByPosition(gcA, nRow).getString())
                cur = Trim(gs.getCellByPosition(gcC, nRow).getString())
                If CtGoodsKey(nm, art) <> CtGoodsKey(CStr(r(0)), CStr(r(1))) Or cur <> "" Then
                    nSkip = nSkip + 1
                Else
                    gs.getCellByPosition(gcC, nRow).setString(cat)
                    nWritten = nWritten + 1
                End If
            End If
        End If
    Next i
    NoCatApply = "OK:Записано в «Номенклатуру»: " & nWritten & IIf(nNoGoods > 0, "." & Chr(10) & "Нет в «Номенклатуре» (добавьте товар, потом повторите): " & nNoGoods, "") _
        & IIf(nSkip > 0, "." & Chr(10) & "Не записано (строка «Номенклатуры» изменилась или категория уже есть): " & nSkip, "") & "."
    gUiLastMsg = NoCatApply
    If nWritten > 0 Then NoCatRefresh(NoCatFind())
    Exit Function
EH:
    NoCatApply = "ERR:запись в «Номенклатуру» не выполнена: " & Error$ & " (строка " & Erl & ")"
End Function

' ================================================================ test seam

' test seam (1.3.0): the hints made for many names at once and the hint of each name alone must agree: the names (Chr(10) between); how many, how many differ, the first of them («name ¦ alone ¦ at once»)
' 1.3.0: test seam — "groups|groups with no category in the registry and in «Номенклатура»|of them with no hint" after StockGroups and GroupCats
Function TestGroupNeed() As String
    Dim i As Long, nNeed As Long, nNo As Long
    WmsInit()
    If Not gCtLoaded Then CtLoad()
    GoodsLoad()
    StockGroups()
    GroupCats()
    For i = 0 To gGrpN - 1
        If CBool(gGrpNeed(i)) Then
            nNeed = nNeed + 1
            If CStr(gGrpHint(i)) = "" Then nNo = nNo + 1
        End If
    Next i
    TestGroupNeed = gGrpN & "|" & nNeed & "|" & nNo
End Function

' 1.3.0: test seam — GroupCats against CategoryOf, GoodsLookup and CategorySuggest, group by group (the same answers must come). "групп N, различий M …", the first differences after
Function TestGroupCats() As String
    Dim i As Long, g As Variant, c As String, nRow As Long, gc As String, nBad As Long, out As String, bNeed As Boolean, hint As String, why As String
    WmsInit()
    If Not gCtLoaded Then CtLoad()
    GoodsLoad()
    StockGroups()
    GroupCats()
    For i = 0 To gGrpN - 1
        g = gGrp(i)
        why = ""
        c = CategoryOf(CStr(g(0)), CStr(g(1)), CStr(g(6)))
        If c <> CStr(gGrpCat(i)) Then why = "категория «" & c & "» ≠ «" & CStr(gGrpCat(i)) & "»"
        If Trim(CStr(g(6))) = "" Then
            nRow = -1
            gc = GoodsLookup(CStr(g(0)), CStr(g(1)), nRow)
            bNeed = (gc = "")
            hint = ""
            If bNeed Then hint = CategorySuggest(CStr(g(0)))
            If nRow <> CLng(gGrpRow(i)) Then why = why & "; строка номенклатуры " & nRow & " ≠ " & CLng(gGrpRow(i))
            If bNeed <> CBool(gGrpNeed(i)) Then why = why & "; «без категории» " & bNeed & " ≠ " & CBool(gGrpNeed(i))
            If hint <> CStr(gGrpHint(i)) Then why = why & "; подсказка «" & hint & "» ≠ «" & CStr(gGrpHint(i)) & "»"
        ElseIf CBool(gGrpNeed(i)) Then
            why = why & "; группа с категорией реестра помечена «без категории»"
        End If
        If why <> "" Then
            nBad = nBad + 1
            If nBad <= 10 Then out = out & Chr(10) & CStr(g(0)) & " | " & CStr(g(1)) & " — " & why
        End If
    Next i
    TestGroupCats = "групп " & gGrpN & ", различий " & nBad & out
End Function

Function TestHintsAll(sList As String) As String
    Dim a As Variant, i As Long, h As Variant, nBad As Long, out As String, x As String
    If Not gCtLoaded Then CtLoad()
    a = Split(sList, Chr(10))
    h = CtHintsAll(a, UBound(a) + 1)
    For i = 0 To UBound(a)
        x = CategorySuggest(CStr(a(i)))
        If x <> CStr(h(i)) Then
            nBad = nBad + 1
            If nBad <= 15 Then out = out & Chr(10) & CStr(a(i)) & " ¦ " & x & " ¦ " & CStr(h(i))
        End If
    Next i
    TestHintsAll = (UBound(a) + 1) & " имён, различий " & nBad & out
End Function

Function TestCategory(sName As String) As String
    TestCategory = CategorySuggest(sName)
End Function

' the category of a name and an article by the order registry-category (the third argument) → «Номенклатура» → the hint
Function TestCategoryOf(sName As String, sArt As String, sRegCat As String) As String
    GoodsLoad()
    TestCategoryOf = CategoryOf(sName, sArt, sRegCat)
End Function

' «М8х30 …» parsed: size|diameter|length|standard|class|coating|type
' 1.3.0: the time of the steps of the last FastRefresh, ms (for the test of the speed)
Function TestFastProf() As String
    TestFastProf = gFsProf
End Function

Function TestGroupProf() As String
    TestGroupProf = gGrpProf
End Function

Function TestFastParse(sName As String) As String
    Dim sSize As String, dD As Double, dL As Double, dP As Double, idx As Long, t As String
    t = WmsSpec.SpType(sName, idx)
    WmsSpec.SpSize(sName, sSize, dD, dL, dP)
    TestFastParse = sSize & "|" & dD & "|" & dL & "|" & WmsSpec.SpStandard(sName) & "|" & WmsSpec.SpClass(sName) & "|" & WmsSpec.SpCoating(sName) & "|" & t
End Function
