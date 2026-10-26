' WmsCat — 1.0.2: подсказки категорий по названию и справочник «Метизы».
' «Категории» — справочник кладовщика (не учёт, не журнал): категория и слова, по которым её узнают в названии («болт, гайка, шайба» →
' «Метизы»). Вписали название товара на «Заказах» — пустая «Категория» заполняется подсказкой; можно написать свою, подсказка её не
' трогает; список категорий — выпадающий, но допускает любое слово. «Метизы» — справочник метизов, собранный из «Наличия» по названию:
' тип, размер (М8×20), стандарт (DIN 933, ГОСТ 7798-70), класс прочности (8.8, А2), покрытие, остаток, места; поиск по любому слову
' названия. Только просмотр: ничего не проводит, в журнал не пишет. Library WmsPlus.
Option Explicit

Public Const SH_CATS = "Категории"
Public Const CT_NAME = 0
Public Const CT_WORDS = 1
Public Const CT_MAX = 300
Public Const CT_FASTENERS = "Метизы"

Public Const SH_FAST = "Метизы"
Public Const FS_FIND_ROW = 1
Public Const FS_STAT_ROW = 2
Public Const FS_HEAD = 3
Public Const FS_ROW0 = 4
Public Const FS_COLS = 11
Public Const FS_KEY_COL = 11
Public Const FS_MAX = 5000

' the categories as the sheet was read (a Global keeps it between the calls; the sheet's own event drops it)
Global gCtCat() As String
Global gCtWords() As String
Global gCtN As Long
Global gCtLoaded As Boolean
' the journal position of the last list of «Метизы» (the list is made again when it moved)
Global gFsSeq As Long
Global gFsBusy As Boolean

' ================================================================ the hints of the categories

' the name as the words are compared: lower case, «ё» = «е», one space between the words
Private Function CtNorm(s As String) As String
    Dim t As String
    t = LCase(Trim(s))
    t = Replace(t, "ё", "е")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CtNorm = t
End Function

' a name for the comparison with the words: as CtNorm, the signs between the words (, ; ( ) / " « » : |) are spaces
Private Function CtName(s As String) As String
    Dim t As String, k As Variant
    t = s
    For Each k In Array(",", ";", "(", ")", "/", Chr(34), "«", "»", ":", "|")
        t = Replace(t, CStr(k), " ")
    Next k
    CtName = CtNorm(t)
End Function

Private Sub CtLoad()
    Dim sh As Object, last As Long, d As Variant, i As Long, n As Long
    gCtN = 0
    ReDim gCtCat(0)
    ReDim gCtWords(0)
    gCtLoaded = True
    If Not gDoc.Sheets.hasByName(SH_CATS) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_CATS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    If last > CT_MAX Then last = CT_MAX
    d = sh.getCellRangeByPosition(0, 1, 1, last).getDataArray()
    ReDim gCtCat(UBound(d))
    ReDim gCtWords(UBound(d))
    For i = 0 To UBound(d)
        If Trim(CStr(d(i)(CT_NAME))) <> "" Then
            gCtCat(n) = Trim(CStr(d(i)(CT_NAME)))
            gCtWords(n) = "," & Replace(CtNorm(CStr(d(i)(CT_WORDS))), ", ", ",") & ","
            n = n + 1
        End If
    Next i
    gCtN = n
End Sub

' the sheet «Категории» was edited: read again at the next hint
Sub OnCatsChange(Optional oTarget As Variant)
    gCtLoaded = False
End Sub

' the category the name of a goods points at (the first category, in the order of the sheet, one of whose words the name has);
' "" — none. A word of the sheet is the beginning of a word of the name: «болт» finds «болт», «болты», «болторез», but «ключ» does
' not find «выключатель»
Function CategorySuggest(sName As String) As String
    Dim nm As String, i As Long, w As Variant, j As Long, k As String
    If Not gCtLoaded Then CtLoad()
    nm = CtName(sName)
    If nm = "" Then Exit Function
    For i = 0 To gCtN - 1
        w = Split(Mid(gCtWords(i), 2, Len(gCtWords(i)) - 2), ",")
        For j = 0 To UBound(w)
            k = Trim(CStr(w(j)))
            If Len(k) >= 2 Then
                If InStr(" " & nm & " ", " " & k) > 0 Then
                    CategorySuggest = gCtCat(i)
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

' the empty category of row r of sheet sSheet gets the hint of the name (nameCol) — the storekeeper may write his own over it. Called by
' the check of the row after a name was typed; never stops it
Sub CategoryHint(sSheet As String, r As Long, nameCol As Integer, catCol As Integer)
    Dim sh As Object, nm As String, c As String
    On Error Resume Next
    sh = gDoc.Sheets.getByName(sSheet)
    If Trim(sh.getCellByPosition(catCol, r).getString()) <> "" Then Exit Sub
    nm = Trim(sh.getCellByPosition(nameCol, r).getString())
    If nm = "" Then Exit Sub
    c = CategorySuggest(nm)
    If c <> "" Then sh.getCellByPosition(catCol, r).setString(c)
End Sub

' ================================================================ «Метизы»

' the types in the order of the list: (the stem of the word, the name of the type)
Private Function FsTypes() As Variant
    FsTypes = Array("болт|Болт", "винт|Винт", "шуруп|Шуруп", "саморез|Саморез", "шпильк|Шпилька", "гайк|Гайка", "шайб|Шайба", "шплинт|Шплинт", _
        "дюбел|Дюбель", "анкер|Анкер", "заклепк|Заклёпка", "гвозд|Гвоздь", "хомут|Хомут", "стопорн|Стопорное кольцо")
End Function

Private Function FsIsDigit(c As String) As Boolean
    FsIsDigit = (Len(c) = 1 And c >= "0" And c <= "9")
End Function

Private Function FsIsLetter(c As String) As Boolean
    FsIsLetter = (Len(c) = 1) And ((c >= "a" And c <= "z") Or (c >= "а" And c <= "я") Or c = "ё")
End Function

' the number at position p of s (digits, a decimal comma or point followed by a digit): its text; p is moved past it
Private Function FsNumber(s As String, ByRef p As Long) As String
    Dim t As String, c As String
    Do While p <= Len(s)
        c = Mid(s, p, 1)
        If FsIsDigit(c) Then
            t = t & c
            p = p + 1
        ElseIf (c = "," Or c = ".") And t <> "" And p < Len(s) Then
            If FsIsDigit(Mid(s, p + 1, 1)) Then
                t = t & "."
                p = p + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    FsNumber = t
End Function

' the size in a name: «М8х30», «M10 x 40», «4,2×32», «М8» → «М8×30», «4.2×32», «М8»; diameter and length as numbers (0 — none)
Private Sub FsParseSize(sName As String, ByRef sSize As String, ByRef dDiam As Double, ByRef dLen As Double)
    Dim s As String, n As Long, i As Long, p As Long, c As String, prev As String, hasM As Boolean, n1 As String, n2 As String, q As Long, st As Long
    s = LCase(sName)
    n = Len(s)
    sSize = ""
    dDiam = 0
    dLen = 0
    i = 1
    Do While i <= n
        c = Mid(s, i, 1)
        hasM = False
        st = i
        If (c = "м" Or c = "m") And i < n Then
            If FsIsDigit(Mid(s, i + 1, 1)) Then
                prev = IIf(i > 1, Mid(s, i - 1, 1), " ")
                If Not FsIsLetter(prev) And Not FsIsDigit(prev) Then
                    hasM = True
                    i = i + 1
                    c = Mid(s, i, 1)
                End If
            End If
        End If
        If FsIsDigit(c) Then
            prev = IIf(st > 1, Mid(s, st - 1, 1), " ")
            If hasM Or (Not FsIsLetter(prev) And Not FsIsDigit(prev) And prev <> "." And prev <> ",") Then
                p = i
                n1 = FsNumber(s, p)
                n2 = ""
                q = p
                Do While q <= n And Mid(s, q, 1) = " "
                    q = q + 1
                Loop
                If q < n Then
                    c = Mid(s, q, 1)
                    If c = "x" Or c = "х" Or c = "×" Or c = "*" Then
                        q = q + 1
                        Do While q <= n And Mid(s, q, 1) = " "
                            q = q + 1
                        Loop
                        If q <= n Then
                            If FsIsDigit(Mid(s, q, 1)) Then
                                n2 = FsNumber(s, q)
                                p = q
                            End If
                        End If
                    End If
                End If
                If hasM Or n2 <> "" Then
                    sSize = IIf(hasM, "М", "") & Replace(n1, ".", ",") & IIf(n2 <> "", "×" & Replace(n2, ".", ","), "")
                    dDiam = Val(n1)
                    dLen = Val(n2)
                    Exit Sub
                End If
                i = p
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

' the number after a mark («din 933», «гост 7798-70», «iso 4017»): its text; "" — none
Private Function FsAfter(s As String, sKey As String) As String
    Dim p As Long, t As String, c As String
    p = InStr(s, sKey)
    Do While p > 0
        p = p + Len(sKey)
        Do While p <= Len(s) And (Mid(s, p, 1) = " " Or Mid(s, p, 1) = "." Or Mid(s, p, 1) = "№")
            p = p + 1
        Loop
        t = ""
        Do While p <= Len(s)
            c = Mid(s, p, 1)
            If FsIsDigit(c) Then
                t = t & c
                p = p + 1
            ElseIf (c = "-" Or c = "–") And t <> "" And p < Len(s) Then
                If FsIsDigit(Mid(s, p + 1, 1)) Then
                    t = t & "-"
                    p = p + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        If t <> "" Then
            FsAfter = t
            Exit Function
        End If
        p = InStr(p, s, sKey)
    Loop
End Function

' DIN, ГОСТ or ISO of a name: «DIN 933», «ГОСТ 7798-70», «ISO 4017»; "" — none
Private Function FsStandard(sName As String) As String
    Dim s As String, t As String
    s = LCase(sName)
    t = FsAfter(s, "din")
    If t <> "" Then
        FsStandard = "DIN " & t
        Exit Function
    End If
    t = FsAfter(s, "гост")
    If t <> "" Then
        FsStandard = "ГОСТ " & t
        Exit Function
    End If
    t = FsAfter(s, "iso")
    If t <> "" Then FsStandard = "ISO " & t
End Function

' the class of strength («8.8», «10.9») or of stainless steel («А2», «А4») of a name; "" — none
Private Function FsClass(sName As String) As String
    Dim a As Variant, i As Long, t As String, cl As Variant, j As Long, s As String
    s = LCase(sName)
    For i = 1 To Len(s)
        If InStr(" ,;()/", Mid(s, i, 1)) > 0 Then s = Left(s, i - 1) & " " & Mid(s, i + 1)
    Next i
    a = Split(s, " ")
    cl = Array("4.6", "5.6", "5.8", "6.8", "8.8", "9.8", "10.9", "12.9")
    For i = 0 To UBound(a)
        t = Replace(Trim(CStr(a(i))), ",", ".")
        If t <> "" Then
            For j = 0 To UBound(cl)
                If t = cl(j) Then
                    FsClass = t
                    Exit Function
                End If
            Next j
            If t = "а2" Or t = "а4" Or t = "a2" Or t = "a4" Or Left(t, 3) = "а2-" Or Left(t, 3) = "а4-" Or Left(t, 3) = "a2-" Or Left(t, 3) = "a4-" Then
                FsClass = "А" & Mid(t, 2, 1) & IIf(Len(t) > 2, Mid(t, 3), "")
                Exit Function
            End If
        End If
    Next i
End Function

' the coating or the material of a name: «Оцинкованный», «Нержавеющая сталь», «Чёрный» …; "" — none
Private Function FsCoating(sName As String) As String
    Dim s As String, k As Variant, i As Long, p As Variant
    s = Replace(LCase(sName), "ё", "е")
    k = Array("оцинк|Оцинкованный", "цинк|Цинк", "нерж|Нержавеющая сталь", "латун|Латунь", "черн|Чёрный", "хромир|Хромированный", "фосфат|Фосфатированный", _
        "медн|Медь", "алюмин|Алюминий", "нейлон|Пластик", "пластик|Пластик", "полиамид|Пластик")
    For i = 0 To UBound(k)
        p = Split(CStr(k(i)), "|")
        If InStr(s, p(0)) > 0 Then
            FsCoating = p(1)
            Exit Function
        End If
    Next i
End Function

' the type of a fastener by the earliest word of the name that names one (the index in the list of types; -1 — none) and its name
Private Function FsType(sName As String, ByRef idx As Long) As String
    Dim s As String, t As Variant, i As Long, p As Variant, k As Long, best As Long
    s = Replace(LCase(sName), "ё", "е")
    t = FsTypes()
    idx = -1
    best = 100000
    For i = 0 To UBound(t)
        p = Split(CStr(t(i)), "|")
        k = InStr(s, Replace(CStr(p(0)), "ё", "е"))
        If k > 0 And k < best Then
            best = k
            idx = i
            FsType = p(1)
        End If
    Next i
End Function

' True when the EI of this name and category is a fastener: its category says so («Метизы», «Крепёж»), or it has none and the name says
' so (the hints of «Категории»); a category of something else («Инструмент» for «Болторез») keeps it out
Private Function FsIsFastener(sName As String, sCat As String) As Boolean
    Dim c As String
    c = CtNorm(sCat)
    If c = "" Then
        FsIsFastener = (CtNorm(CategorySuggest(sName)) = CtNorm(CT_FASTENERS))
    Else
        FsIsFastener = (InStr(c, "метиз") > 0 Or InStr(c, "крепеж") > 0)
    End If
End Function

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

' the search cell was changed: the list is made again with the words
Sub OnFastChange(Optional oTarget As Variant)
    Dim ra As Object
    If gFsBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> FS_FIND_ROW Or ra.Column <> 1 Then Exit Sub
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

' the list of the fasteners from «Наличие»: one row for a name with an article — the balance of all its EIs, the places, how many EIs
' have a balance; sorted by type, diameter, length, name. sFind — the words every one of which the row must have (any column).
' OK:<rows> / ERR:…
Function FastRefresh(sFind As String) As String
    Dim st As Object, last As Long, d As Variant, i As Long, j As Long, sh As Object, nm As String, art As String, key As String, k As Long, n As Long
    Dim typeName As String, tIdx As Long, sSize As String, dDiam As Double, dLen As Double, std As String, cl As String, cov As String, unit As String
    Dim keys() As String, rows() As Variant, bal() As Double, places() As String, nEI() As Long, cnt() As Long, out() As Variant, words As Variant, hay As String, w As Variant, ok As Boolean
    Dim cEi As Long, cNm As Long, cArt As Long, cUn As Long, cQty As Long, cPl As Long, cCat As Long, q As Double, pl As String, nShown As Long, ra As Object, dbr As Object
    On Error GoTo EH
    gFsBusy = True
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_FAST) Then
        FastRefresh = "ERR:в книге нет листа «" & SH_FAST & "» — книга не обновлена до 1.0.2"
        gFsBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_FAST)
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    If Not gCtLoaded Then CtLoad()
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cUn = SC_UNIT
    cQty = SC_QTY
    cPl = SC_PLACE
    cCat = SC_CAT
    If last >= 1 Then
        d = st.getCellRangeByPosition(0, 1, SC_CAT, last).getDataArray()
        ReDim keys(UBound(d))
        ReDim rows(UBound(d))
        ReDim bal(UBound(d))
        ReDim places(UBound(d))
        ReDim nEI(UBound(d))
        ReDim cnt(UBound(d))
        For i = 0 To UBound(d)
            If CStr(d(i)(cEi)) <> "" Then
                nm = Trim(CStr(d(i)(cNm)))
                If FsIsFastener(nm, CStr(d(i)(cCat))) Then
                    art = Trim(CStr(d(i)(cArt)))
                    key = LCase(nm) & "|" & LCase(art)
                    k = -1
                    For j = 0 To n - 1
                        If keys(j) = key Then
                            k = j
                            Exit For
                        End If
                    Next j
                    If k < 0 Then
                        k = n
                        keys(k) = key
                        rows(k) = Array(nm, art, Trim(CStr(d(i)(cUn))))
                        n = n + 1
                    End If
                    q = 0
                    If VarType(d(i)(cQty)) = 5 Then q = d(i)(cQty)
                    bal(k) = bal(k) + q
                    cnt(k) = cnt(k) + 1
                    If q > 0 Then nEI(k) = nEI(k) + 1
                    pl = Trim(CStr(d(i)(cPl)))
                    If q > 0 And pl <> "" Then
                        If InStr("; " & places(k) & ";", "; " & pl & ";") = 0 Then places(k) = places(k) & IIf(places(k) <> "", "; ", "") & pl
                    End If
                End If
            End If
        Next i
    End If
    ' the words of the search
    words = Split(CtNorm(sFind), " ")
    ReDim out(IIf(n > 0, n - 1, 0))
    k = 0
    For i = 0 To n - 1
        typeName = FsType(CStr(rows(i)(0)), tIdx)
        FsParseSize(CStr(rows(i)(0)), sSize, dDiam, dLen)
        std = FsStandard(CStr(rows(i)(0)))
        cl = FsClass(CStr(rows(i)(0)))
        cov = FsCoating(CStr(rows(i)(0)))
        If typeName = "" Then
            typeName = "Прочее"
            tIdx = 99
        End If
        ok = True
        If Trim(sFind) <> "" Then
            hay = CtNorm(typeName & " " & sSize & " " & std & " " & cl & " " & cov & " " & CStr(rows(i)(0)) & " " & CStr(rows(i)(1)) & " " & places(i))
            For Each w In words
                If CStr(w) <> "" Then
                    If InStr(hay, CStr(w)) = 0 Then ok = False
                End If
            Next w
        End If
        If ok Then
            out(k) = Array(typeName, sSize, std, cl, cov, rows(i)(0), rows(i)(1), WmsPack.PkRound(bal(i), 3), rows(i)(2), places(i), nEI(i), CDbl(tIdx) * 1000000000000# + CDbl(Int(dDiam * 100 + 0.5)) * 1000000# + CDbl(Int(dLen * 10 + 0.5)))
            k = k + 1
        End If
    Next i
    ' the old list out, the new one in
    ra = sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + FS_MAX).getRangeAddress()
    sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + FS_MAX).clearContents(1 + 2 + 4 + 16)
    If k > FS_MAX Then k = FS_MAX
    If k > 0 Then
        ReDim Preserve out(k - 1)
        sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + k - 1).setDataArray(out)
        FsSort(sh.getCellRangeByPosition(0, FS_ROW0, FS_KEY_COL, FS_ROW0 + k - 1))
        ' the zero balances grey, the rest as they are
        sh.getCellRangeByPosition(0, FS_ROW0, FS_COLS - 1, FS_ROW0 + k - 1).CharColor = RGB(38, 50, 56)
    End If
    sh.getCellByPosition(0, FS_STAT_ROW).setString("Метизов (позиций): " & k & IIf(Trim(sFind) <> "", " — по словам «" & Trim(sFind) & "»", "") & ". Список из «Наличия» на " & Format(Now(), "DD.MM.YYYY HH:MM") _
        & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
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

' the list sorted by the key of the type, the diameter and the length (the hidden column L), then by the name
Private Sub FsSort(rng As Object)
    Dim f(1) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = FS_KEY_COL
    f(0).IsAscending = True
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    f(1).Field = 5
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

' ================================================================ test seam

Function TestCategory(sName As String) As String
    TestCategory = CategorySuggest(sName)
End Function

' «М8х30 …» parsed: size|diameter|length|standard|class|coating|type
Function TestFastParse(sName As String) As String
    Dim sSize As String, dD As Double, dL As Double, idx As Long, t As String
    t = FsType(sName, idx)
    FsParseSize(sName, sSize, dD, dL)
    TestFastParse = sSize & "|" & dD & "|" & dL & "|" & FsStandard(sName) & "|" & FsClass(sName) & "|" & FsCoating(sName) & "|" & t
End Function
