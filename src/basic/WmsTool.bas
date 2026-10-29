' WmsTool — 1.2.0: «Инструмент ЧПУ». What a name of a cutting tool or of a consumable of a machine says: the type («Фреза», «Сверло», «Метчик», «Пластина»…), the kind («концевая», «спиральное»), the diameter, the length,
' the number of teeth (Z), the material (Р6М5, ВК8, HSS-Co, твёрдый сплав), the coating (TiN, TiAlN), the standard (DIN 338, ГОСТ 10902-77) and its equivalents (the sheet «Стандарты»), the shank, the thread or the code
' (ER32, BT40, МК2, М8×1,25, CNMG 120408). A name is read as a person reads it — by the rules below, no network, no guessing: what is not found stays empty (the search by the words of the name always works).
' The sheet «Инструмент ЧПУ» is made from «Наличие» like «Метизы» (one row per name with an article: the balance of all its EIs, the places), with the search of «Поиск», the filters of the columns and the
' possible doubles; it only shows, nothing is posted. Library WmsPlus.
Option Explicit

Private Const TL_SHEET = "Инструмент ЧПУ"
Private Const TL_FIND_ROW = 1
Private Const TL_FILTER_ROW = 2
Private Const TL_HEAD = 3
Private Const TL_ROW0 = 4
Private Const TL_COLS = 18
Private Const TL_KEY_COL = 18
Private Const TL_STAT_COL = 7
Private Const TL_MAX = 5000

' the columns of the sheet
Private Const TC_TYPE = 0
Private Const TC_KIND = 1
Private Const TC_SIZE = 2
Private Const TC_DIAM = 3
Private Const TC_LEN = 4
Private Const TC_Z = 5
Private Const TC_MAT = 6
Private Const TC_COAT = 7
Private Const TC_STD = 8
Private Const TC_EQ = 9
Private Const TC_DES = 10
Private Const TC_NAME = 11
Private Const TC_ART = 12
Private Const TC_QTY = 13
Private Const TC_UNIT = 14
Private Const TC_PLACES = 15
Private Const TC_NEI = 16
Private Const TC_NOTE = 17

' the fields of a record (TlRecord)
Private Const RC_TYPE = 0
Private Const RC_IDX = 1
Private Const RC_GRP = 2
Private Const RC_KIND = 3
Private Const RC_SIZE = 4
Private Const RC_DIAM = 5
Private Const RC_LEN = 6
Private Const RC_Z = 7
Private Const RC_MAT = 8
Private Const RC_COAT = 9
Private Const RC_STD = 10
Private Const RC_EQ = 11
Private Const RC_DES = 12
Private Const RC_LAST = 12

' the folded keys of the tables (made at the first use; the sheet of the standards drops the records through gSpCacheOn)
Global gTlInit As Boolean
Global gTlTab As Variant
Global gTlCache As Object
Global gTlCacheOn As Boolean
Global gTlSeq As Long
Global gTlBusy As Boolean
' the constants of the rules in the folded form, made once each: «Chr(1) key Chr(2) folded Chr(1) key Chr(2) folded Chr(1)» (a Fold costs about 1.5 ms: not for every name)
Global gTlKTab As String
' the milliseconds of the steps of the last list (a line for the tests and for a talk about the speed)
Global gTlProf As String

' ================================================================ the tables

' the types in the order of the list: stems (folded words, "," between them; a stem begins a word) | name | group | what else a name must have to be taken ("" — nothing; "#insert" — the code of an
' insert or a word of an insert; "#holder" — the code of a holder or a word of a chuck)
' groups: M milling, D drilling, T threads, R reaming and countersinking, L turning, H holders, A abrasive, S saws and files, C consumables of the machines
Private Function TlTypeDefs() As Variant
    TlTypeDefs = Array( _
        "фрез|Фреза|M|", "борфрез,шарошк|Борфреза|M|", "гравер,штихел|Гравёр|M|", _
        "сверл|Сверло|D|", "коронк|Коронка|D|", "бур|Бур|D|", _
        "метчик|Метчик|T|", "плашк,лерк|Плашка|T|", _
        "развертк|Развёртка|R|", "зенковк|Зенковка|R|", "зенкер|Зенкер|R|", "цековк|Цековка|R|", _
        "резец,резц|Резец|L|", "пластин|Пластина|L|#insert", "державк|Державка|L|", "головка расточн,расточная головк|Расточная головка|L|", _
        "оправк|Оправка|H|", "цанг|Цанга|H|", "патрон|Патрон|H|#holder", "центр вращающ,центр токарн,центр упорн,центр рифлен|Центр|H|", _
        "круг шлиф,шлифкруг,круг отрез,диск отрез,диск шлиф,лепестк,диск алмаз,круг алмаз,шлифлент|Абразив|A|", _
        "диск пильн,пильный диск|Пильный диск|A|", "полотно лент,полотно нож,полотно пильн,пилк|Полотно пильное|S|", "напильник,надфил,рашпил|Напильник|S|", _
        "электрод плазм,сопло плазм,сопло лазер,линза лазер,стекло защитн,защитн стекл|Расходник плазмы / лазера|C|", "сож,эмульсол,охлаждающ жидк|СОЖ|C|")
End Function

' the kinds: stems (folded)
Private Function TlKindDefs() As Variant
    TlKindDefs = Array("концев", "торцев", "шаров", "сферич", "дисков", "пазов", "т-образ", "угловая", "галтел", "фасонн", "червячн", "модульн", "гравиров", "кукуруз", "черновая", "чистов", _
        "радиусн", "конусн", "по алюмин", "по нерж", "по дерев", "по пластик", "по мдф", "по камн", "по стал", "по металл", "по бетон", "по стекл", "по плитк", "по кирпич", "для алюмин", _
        "спиральн", "центровоч", "ступенчат", "перов", "кольцев", "форстнер", "удлинен", "машинно-ручн", "машинн", "ручн", "гаечн", "бесстружеч", "раскатник", "трубн", "коническ", "комплект", _
        "проходн", "отрезн", "расточн", "резьбов", "канавоч", "подрезн", "фасоч", "токарн", "фрезерн", "сменн", "шлифов", "лепестков", "алмазн", "ленточн", "биметалл", "плоск", "круглый", "полукругл", "трехгранн", "квадратн")
End Function

' the materials: key (folded) | name
Private Function TlMatDefs() As Variant
    TlMatDefs = Array("р6м5к5|Р6М5К5", "р6м5ф3|Р6М5Ф3", "р6м5|Р6М5", "р6м3|Р6М3", "р18|Р18", "р12|Р12", "р9|Р9", "р9к5|Р9К5", "9хс|9ХС", "хвг|ХВГ", "ст.у8|Сталь У8", "ст.у10|Сталь У10", "у10а|У10А", "у8а|У8А", _
        "вк8|ВК8", "вк6|ВК6", "вк10|ВК10", "вк15|ВК15", "вк3|ВК3", "вк4|ВК4", "т15к6|Т15К6", "т5к10|Т5К10", "т30к4|Т30К4", "т14к8|Т14К8", "hss-e|HSS-E", "hss-co|HSS-Co", "hssco|HSS-Co", "hss-g|HSS-G", "hss-pm|HSS-PM", "hss|HSS", _
        "vhm|Твёрдый сплав", "solid carbide|Твёрдый сплав", "карбид|Твёрдый сплав", "твердосплав|Твёрдый сплав", "твердый сплав|Твёрдый сплав", "т/с|Твёрдый сплав", "быстрорез|Быстрорез", "керамик|Керамика", _
        "cbn|CBN", "эльбор|CBN", "pcd|PCD", "алмаз|Алмаз", "кобальт|Кобальт", "инструментальн|Инструментальная сталь")
End Function

' the coatings: key (folded) | name
Private Function TlCoatDefs() As Variant
    TlCoatDefs = Array("tialn|TiAlN", "altin|AlTiN", "alcrn|AlCrN", "ticn|TiCN", "tisin|TiSiN", "altisin|AlTiSiN", "tin|TiN", "crn|CrN", "zrn|ZrN", "dlc|DLC", "al2o3|Al2O3", "tialcrn|TiAlCrN", _
        "нитрид титан|TiN", "оксид|Оксидное", "паровая обработка|Паровая обработка", "пароокс|Паровая обработка", "без покрыт|Без покрытия", "bright|Без покрытия")
End Function

' the codes of the holders (the letters before the number): folded prefix | the name written
Private Function TlHolderDefs() As Variant
    TlHolderDefs = Array("er|ER", "bt|BT", "sk|SK", "hsk|HSK", "cat|CAT", "nt|NT", "mk|МК", "км|КМ", "mt|МК", "morse|МК")
End Function

' the threads and the pipe sizes: the prefix (folded) | the name written
Private Function TlThreadDefs() As Variant
    TlThreadDefs = Array("npt|NPT", "unc|UNC", "unf|UNF", "unef|UNEF", "bsp|BSP", "rc|Rc", "rp|Rp", "g|G", "r|R")
End Function

' the English words of the types for the search
Private Function TlEnglish(sType As String) As String
    Select Case sType
    Case "Фреза", "Борфреза": TlEnglish = "mill end mill endmill cutter"
    Case "Сверло": TlEnglish = "drill drill bit"
    Case "Коронка": TlEnglish = "hole saw core bit"
    Case "Метчик": TlEnglish = "tap"
    Case "Плашка": TlEnglish = "die"
    Case "Развёртка": TlEnglish = "reamer"
    Case "Зенковка", "Зенкер", "Цековка": TlEnglish = "countersink counterbore"
    Case "Резец": TlEnglish = "turning tool lathe tool"
    Case "Пластина": TlEnglish = "insert"
    Case "Державка": TlEnglish = "tool holder shank"
    Case "Оправка": TlEnglish = "arbor"
    Case "Цанга": TlEnglish = "collet"
    Case "Патрон": TlEnglish = "chuck"
    Case "Абразив": TlEnglish = "wheel disc abrasive"
    Case "Пильный диск", "Полотно пильное": TlEnglish = "saw blade"
    Case "Напильник": TlEnglish = "file"
    End Select
End Function

' the constant sKey as WmsNames.Fold makes it (read from the memo; the first time it is folded and put there)
Private Function TlK(sKey As String) As String
    Dim p As Long, q As Long, t As String
    If gTlKTab = "" Then gTlKTab = Chr(1)
    t = Chr(1) & sKey & Chr(2)
    p = InStr(gTlKTab, t)
    If p > 0 Then
        p = p + Len(t)
        q = InStr(p, gTlKTab, Chr(1))
        TlK = Mid(gTlKTab, p, q - p)
        Exit Function
    End If
    TlK = WmsNames.Fold(sKey)
    gTlKTab = gTlKTab & sKey & Chr(2) & TlK & Chr(1)
End Function

' the tables folded once
Private Sub TlLoad()
    Dim d As Variant, i As Long, p As Variant, k As Variant, j As Long, a As Variant
    If gTlInit Then Exit Sub
    ' gTlTab: 0 types, 1 kinds, 2 materials, 3 coatings, 4 holders, 5 threads — every item an array of the folded keys and the rest
    Dim types() As Variant, kinds() As String, mats() As Variant, coats() As Variant, hold() As Variant, thr() As Variant
    d = TlTypeDefs()
    ReDim types(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        a = Split(CStr(p(0)), ",")
        For j = 0 To UBound(a)
            a(j) = WmsNames.Fold(CStr(a(j)))
        Next j
        types(i) = Array(a, CStr(p(1)), CStr(p(2)), CStr(p(3)))
    Next i
    d = TlKindDefs()
    ReDim kinds(UBound(d))
    For i = 0 To UBound(d)
        kinds(i) = WmsNames.Fold(CStr(d(i)))
    Next i
    d = TlMatDefs()
    ReDim mats(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        mats(i) = Array(WmsNames.Fold(CStr(p(0))), CStr(p(1)), TlIsStemKey(WmsNames.Fold(CStr(p(0)))))
    Next i
    d = TlCoatDefs()
    ReDim coats(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        coats(i) = Array(WmsNames.Fold(CStr(p(0))), CStr(p(1)), TlIsStemKey(WmsNames.Fold(CStr(p(0)))))
    Next i
    d = TlHolderDefs()
    ReDim hold(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        hold(i) = Array(WmsNames.Fold(CStr(p(0))), CStr(p(1)))
    Next i
    d = TlThreadDefs()
    ReDim thr(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        thr(i) = Array(WmsNames.Fold(CStr(p(0))), CStr(p(1)))
    Next i
    gTlTab = Array(types, kinds, mats, coats, hold, thr)
    gTlInit = True
End Sub

' ================================================================ the characters

Private Function TlIsDigit(c As String) As Boolean
    TlIsDigit = (Len(c) = 1 And c >= "0" And c <= "9")
End Function

Private Function TlIsLetter(c As String) As Boolean
    TlIsLetter = (Len(c) = 1) And ((c >= "a" And c <= "z") Or (c >= "а" And c <= "я") Or c = "ё")
End Function

Private Function TlIsAlnum(c As String) As Boolean
    TlIsAlnum = TlIsDigit(c) Or TlIsLetter(c)
End Function

' the character at position p of s ("" outside)
Private Function TlCh(s As String, p As Long) As String
    If p >= 1 And p <= Len(s) Then TlCh = Mid(s, p, 1)
End Function

' the number at position p (digits, a decimal point followed by a digit — the folded text has a point instead of a comma): its text; p is moved past it
Private Function TlNumAt(s As String, ByRef p As Long) As String
    Dim t As String, c As String
    Do While p <= Len(s)
        c = Mid(s, p, 1)
        If TlIsDigit(c) Then
            t = t & c
            p = p + 1
        ElseIf (c = "." Or c = ",") And t <> "" Then
            If TlIsDigit(TlCh(s, p + 1)) Then
                t = t & "."
                p = p + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    TlNumAt = t
End Function

' the number as it is written for people: a comma for a point, a dash in a range («4–20»)
Private Function TlShow(ByVal t As String) As String
    TlShow = Replace(Replace(t, ".", ","), Chr(2), "–")
End Function

' the first position of the key at the beginning of a word, from start; 0 — none
Private Function TlWord(s As String, key As String, start As Long) As Long
    Dim p As Long
    If key = "" Then Exit Function
    p = InStr(start, s, key)
    Do While p > 0
        If p = 1 Then
            TlWord = p
            Exit Function
        ElseIf Not TlIsAlnum(Mid(s, p - 1, 1)) Then
            TlWord = p
            Exit Function
        End If
        p = InStr(p + 1, s, key)
    Loop
End Function

' the end of the word that holds position p (the last letter or digit)
Private Function TlWordEnd(s As String, p As Long) As Long
    Dim q As Long
    q = p
    Do While q < Len(s)
        If TlIsAlnum(Mid(s, q + 1, 1)) Or (Mid(s, q + 1, 1) = "-" And TlIsAlnum(TlCh(s, q + 2))) Then
            q = q + 1
        Else
            Exit Do
        End If
    Loop
    TlWordEnd = q
End Function

' the span a..b of the mask is taken (b >= a)
Private Sub TlTake(ByRef mask As String, a As Long, b As Long)
    If a < 1 Then a = 1
    If b > Len(mask) Then b = Len(mask)
    If b >= a Then mask = Left(mask, a - 1) & String(b - a + 1, "x") & Mid(mask, b + 1)
End Sub

' no character of the span a..b is taken yet
Private Function TlFree(mask As String, a As Long, b As Long) As Boolean
    If a < 1 Then a = 1
    If b > Len(mask) Then b = Len(mask)
    If b < a Then
        TlFree = True
        Exit Function
    End If
    TlFree = (InStr(Mid(mask, a, b - a + 1), "x") = 0)
End Function

' ================================================================ the parts of a name

' a standard: «DIN 338», «ГОСТ 10902-77», «ISO 1832»; the span is taken
Private Function TlStandard(s As String, orig As String, ByRef mask As String) As String
    Dim k As Variant, keys As Variant, p As Long, q As Long, t As String, c As String, nm As String
    If InStr(s, TlK("din")) + InStr(s, TlK("гост")) + InStr(s, TlK("iso")) + InStr(s, TlK("исо")) + InStr(s, TlK("ту")) + InStr(s, TlK("ansi")) + InStr(s, TlK("jis")) = 0 Then Exit Function
    keys = Array("din|DIN", "гост|ГОСТ", "iso|ISO", "исо|ISO", "ту|ТУ", "ansi|ANSI", "jis|JIS")
    For Each k In keys
        nm = TlK(Split(CStr(k), "|")(0))
        p = 0
        If InStr(s, nm) > 0 Then p = TlWord(s, nm, 1)
        Do While p > 0
            q = p + Len(nm)
            If Not TlIsLetter(TlCh(s, q)) Then
                Do While q <= Len(s) And (Mid(s, q, 1) = " " Or Mid(s, q, 1) = "." Or Mid(s, q, 1) = "№")
                    q = q + 1
                Loop
                t = ""
                Do While q <= Len(s)
                    c = Mid(s, q, 1)
                    If TlIsDigit(c) Then
                        t = t & c
                        q = q + 1
                    ElseIf (c = "-" Or c = "." Or c = "/") And t <> "" And TlIsDigit(TlCh(s, q + 1)) Then
                        t = t & c
                        q = q + 1
                    Else
                        Exit Do
                    End If
                Loop
                If t <> "" Then
                    TlTake(mask, p, q - 1)
                    TlStandard = Split(CStr(k), "|")(1) & " " & t
                    Exit Function
                End If
            End If
            p = TlWord(s, nm, p + 1)
        Loop
    Next k
End Function

' the code of an insert («CNMG 120408», «DCMT11T304») or of a holder («MCLNR 2525M12»): Latin letters (four for an insert, four to six for a holder), a space may follow, two digits, then two to six letters or digits,
' at least four digits in all; a size («фреза 12х50», «125x22») is not a code: no «x», no separator, the letters are Latin. The span is taken; the code as the name wrote it, in capitals
Private Function TlInsert(s As String, orig As String, ByRef mask As String) As String
    Dim p As Long, q As Long, n As Long, a As Long, nd As Long, c As String, nl As Long, i As Long, ch As String, ok As Boolean
    n = Len(s)
    p = 1
    Do While p <= n - 7
        If TlIsLetter(Mid(s, p, 1)) And (p = 1 Or Not TlIsAlnum(TlCh(s, p - 1))) Then
            a = p
            q = p
            Do While q <= n And TlIsLetter(Mid(s, q, 1))
                q = q + 1
            Loop
            nl = q - a
            If nl >= 4 And nl <= 6 Then
                ' the letters of the head are Latin in the name as it was written
                ok = True
                For i = a To q - 1
                    ch = LCase(Mid(orig, i, 1))
                    If Not (ch >= "a" And ch <= "z") Then ok = False
                Next i
                nd = q
                If TlCh(s, nd) = " " Then nd = nd + 1
                If ok And TlIsDigit(TlCh(s, nd)) And TlIsDigit(TlCh(s, nd + 1)) Then
                    q = nd
                    Do While q <= n And TlIsAlnum(Mid(s, q, 1))
                        q = q + 1
                    Loop
                    c = Mid(s, nd, q - nd)
                    ' no separator of a size in the tail (the folded «x» is a Cyrillic «х»)
                    For i = nd To q - 1
                        ch = LCase(Mid(orig, i, 1))
                        If ch = "x" Or ch = "х" Or ch = "×" Then ok = False
                    Next i
                    If ok And q - nd >= 4 And q - nd <= 8 And TlDigits(c) >= 4 And TlFree(mask, a, q - 1) Then
                        TlTake(mask, a, q - 1)
                        TlInsert = UCase(Mid(orig, a, nl)) & " " & UCase(Mid(orig, nd, q - nd))
                        Exit Function
                    End If
                End If
            End If
            p = IIf(q > p, q, p + 1)
        Else
            p = p + 1
        End If
    Loop
End Function

Private Function TlDigits(t As String) As Long
    Dim i As Long
    For i = 1 To Len(t)
        If TlIsDigit(Mid(t, i, 1)) Then TlDigits = TlDigits + 1
    Next i
End Function

' the code of a holder or a taper («ER32», «BT40», «HSK-A63», «МК2», «конус Морзе 2», «SDS-plus»): the codes found in the order of the name, joined by a comma; the spans are taken
Private Function TlHolders(s As String, orig As String, ByRef mask As String) As String
    Dim hd As Variant, i As Long, p As Long, q As Long, a As Long, pre As String, num As String, c As String, letter As String, tail As String, isHsk As Boolean, code As String
    Dim cp(20) As Long, cc(20) As String, n As Long, j As Long, tp As Long, tc As String, out As String, dup As Boolean
    TlLoad()
    hd = gTlTab(4)
    For i = 0 To UBound(hd)
        pre = CStr(hd(i)(0))
        isHsk = (pre = TlK("hsk"))
        p = 0
        If InStr(s, pre) > 0 Then p = TlWord(s, pre, 1)
        Do While p > 0
            q = p + Len(pre)
            ' the prefix must not go on with other letters (only the letter of an HSK after a dash: «hsk-a63»)
            If Not TlIsLetter(TlCh(s, q)) Or isHsk Then
                a = q
                If TlCh(s, a) = "-" Or TlCh(s, a) = " " Then a = a + 1
                letter = ""
                tail = ""
                If isHsk And TlIsLetter(TlCh(s, a)) And TlIsDigit(TlCh(s, a + 1)) Then
                    letter = UCase(Mid(orig, a, 1))
                    a = a + 1
                End If
                If TlIsDigit(TlCh(s, a)) And TlFree(mask, p, a) Then
                    num = TlNumAt(s, a)
                    If isHsk And TlIsLetter(TlCh(s, a)) And Not TlIsAlnum(TlCh(s, a + 1)) Then
                        tail = UCase(Mid(orig, a, 1))
                        a = a + 1
                    End If
                    If Not TlIsLetter(TlCh(s, a)) And n < 20 Then
                        code = CStr(hd(i)(1)) & letter & num & tail
                        dup = False
                        For j = 0 To n - 1
                            If cc(j) = code Then dup = True
                        Next j
                        If Not dup Then
                            cp(n) = p
                            cc(n) = code
                            n = n + 1
                        End If
                        TlTake(mask, p, a - 1)
                    End If
                End If
            End If
            p = TlWord(s, pre, p + 1)
        Loop
    Next i
    ' «конус Морзе 2», «Morse taper 2»
    p = TlWord(s, TlK("морзе"), 1)
    If p = 0 Then p = TlWord(s, TlK("morse"), 1)
    If p > 0 Then
        q = TlWordEnd(s, p) + 1
        Do While q <= Len(s) And InStr(" №", Mid(s, q, 1)) > 0
            q = q + 1
        Loop
        If Mid(s, q, 5) = "taper" Then q = q + 6
        If TlIsDigit(TlCh(s, q)) And n < 20 Then
            num = TlNumAt(s, q)
            cp(n) = p
            cc(n) = "МК" & num
            n = n + 1
            TlTake(mask, p, q - 1)
        End If
    End If
    ' SDS
    p = TlWord(s, TlK("sds"), 1)
    If p > 0 And n < 20 Then
        c = Mid(s, p, 12)
        cp(n) = p
        If InStr(c, TlK("max")) > 0 Then
            cc(n) = "SDS-max"
        ElseIf InStr(c, "+") > 0 Or InStr(c, TlK("plus")) > 0 Or InStr(c, TlK("плюс")) > 0 Then
            cc(n) = "SDS-plus"
        Else
            cc(n) = "SDS"
        End If
        n = n + 1
        q = p + 3
        Do While q <= Len(s) And InStr(" -+", Mid(s, q, 1)) > 0
            q = q + 1
        Loop
        Do While q <= Len(s) And TlIsLetter(Mid(s, q, 1))
            q = q + 1
        Loop
        TlTake(mask, p, q - 1)
    End If
    ' in the order of the name
    For i = 1 To n - 1
        For j = 0 To n - 1 - i
            If cp(j) > cp(j + 1) Then
                tp = cp(j)
                cp(j) = cp(j + 1)
                cp(j + 1) = tp
                tc = cc(j)
                cc(j) = cc(j + 1)
                cc(j + 1) = tc
            End If
        Next j
    Next i
    For i = 0 To n - 1
        out = out & IIf(out <> "", ", ", "") & cc(i)
    Next i
    TlHolders = out
End Function

' a metric thread with the pitch («М8х1,25» → "М8×1,25", the diameter 8; «М8х1,25х30», «М8х30» — the length 30 is returned in dLen), a pipe or an inch thread («G1/4», «NPT 1/4», «UNC 1/4-20»): the text; the span is taken
Private Function TlThread(s As String, orig As String, ByRef mask As String, ByRef dThread As Double, ByRef dLen As Double) As String
    Dim p As Long, q As Long, n1 As String, n2 As String, n3 As String, th As Variant, i As Long, pre As String, t As String, a As Long, mk As String
    TlLoad()
    dThread = 0
    dLen = 0
    mk = TlK("м")
    ' metric: «м» at the beginning of a word, then a digit
    p = TlWord(s, mk, 1)
    Do While p > 0
        If TlIsDigit(TlCh(s, p + 1)) And TlFree(mask, p, p + 1) Then
            q = p + 1
            n1 = TlNumAt(s, q)
            n2 = ""
            n3 = ""
            a = q
            If TlCh(s, a) = " " Then a = a + 1
            If TlCh(s, a) = "х" And TlIsDigit(TlCh(s, a + 1)) Then
                a = a + 1
                n2 = TlNumAt(s, a)
                q = a
                If TlCh(s, a) = "х" And TlIsDigit(TlCh(s, a + 1)) Then
                    a = a + 1
                    n3 = TlNumAt(s, a)
                    q = a
                End If
            End If
            If Not TlIsLetter(TlCh(s, q)) And Val(n1) > 0 And Val(n1) <= 100 Then
                dThread = Val(n1)
                If n3 <> "" Then
                    dLen = Val(n3)
                    TlThread = "М" & TlShow(n1) & "×" & TlShow(n2)
                ElseIf n2 <> "" And Val(n2) < Val(n1) Then
                    TlThread = "М" & TlShow(n1) & "×" & TlShow(n2)
                ElseIf n2 <> "" Then
                    dLen = Val(n2)
                    TlThread = "М" & TlShow(n1)
                Else
                    TlThread = "М" & TlShow(n1)
                End If
                TlTake(mask, p, q - 1)
                Exit Function
            End If
        End If
        p = TlWord(s, mk, p + 1)
    Loop
    ' pipe and inch threads
    th = gTlTab(5)
    For i = 0 To UBound(th)
        pre = CStr(th(i)(0))
        p = 0
        If InStr(s, pre) > 0 Then p = TlWord(s, pre, 1)
        Do While p > 0
            q = p + Len(pre)
            If TlCh(s, q) = " " Then q = q + 1
            If TlIsDigit(TlCh(s, q)) And TlFree(mask, p, q) And Not TlIsLetter(TlCh(s, p + Len(pre))) Then
                t = TlNumAt(s, q)
                If TlCh(s, q) = "/" And TlIsDigit(TlCh(s, q + 1)) Then
                    q = q + 1
                    t = t & "/" & TlNumAt(s, q)
                    If TlCh(s, q) = "-" And TlIsDigit(TlCh(s, q + 1)) Then
                        q = q + 1
                        t = t & "-" & TlNumAt(s, q)
                    End If
                    TlTake(mask, p, q - 1)
                    TlThread = CStr(th(i)(1)) & " " & t
                    Exit Function
                ElseIf pre = TlK("g") And Not TlIsLetter(TlCh(s, q)) Then
                    TlTake(mask, p, q - 1)
                    TlThread = "G " & t
                    Exit Function
                ElseIf pre <> TlK("g") And pre <> TlK("r") And pre <> TlK("rc") And pre <> TlK("rp") And Not TlIsLetter(TlCh(s, q)) Then
                    TlTake(mask, p, q - 1)
                    TlThread = CStr(th(i)(1)) & " " & t
                    Exit Function
                End If
            End If
            p = TlWord(s, pre, p + 1)
        Loop
    Next i
End Function

' the shank in words: «хвостовик 6», «ц/х», «к/х», «цилиндрический хвостовик»; the text ("" — none)
Private Function TlShank(s As String, ByRef mask As String) As String
    Dim p As Long, q As Long, num As String, out As String
    If InStr(s, TlK("хвост")) + InStr(s, TlK("хв.")) + InStr(s, "shank") + InStr(s, TlK("ц/х")) + InStr(s, TlK("к/х")) = 0 Then Exit Function
    p = TlWord(s, TlK("хвост"), 1)
    If p = 0 Then p = TlWord(s, TlK("хв."), 1)
    If p = 0 Then p = TlWord(s, TlK("shank"), 1)
    If p > 0 Then
        q = TlWordEnd(s, p) + 1
        Do While q <= Len(s) And InStr(" =.:", Mid(s, q, 1)) > 0
            q = q + 1
        Loop
        If TlCh(s, q) = "ø" Or TlCh(s, q) = "⌀" Or TlCh(s, q) = "d" Then q = q + 1
        If TlIsDigit(TlCh(s, q)) Then
            num = TlNumAt(s, q)
            TlTake(mask, p, q - 1)
            out = "хв. Ø" & TlShow(num)
        End If
    End If
    If InStr(s, TlK("ц/х")) > 0 Or (InStr(s, TlK("цилиндрическ")) > 0 And InStr(s, TlK("хвост")) > 0) Then out = out & IIf(out <> "", ", ", "") & "цилиндрический хвостовик"
    If InStr(s, TlK("к/х")) > 0 Or (InStr(s, TlK("коническ")) > 0 And InStr(s, TlK("хвост")) > 0) Then out = out & IIf(out <> "", ", ", "") & "конический хвостовик"
    TlShank = out
End Function

' the number of teeth or flutes: «Z4», «z=4», «4z», «4 зуба», «4-х зубая», «3 флейты», «двухзаходная»; 0 — none
Private Function TlZ(s As String, ByRef mask As String) As Long
    Dim p As Long, q As Long, num As String, k As Variant, w As Variant, st As String
    ' «z4», «z=4», «z 4»
    p = TlWord(s, "z", 1)
    Do While p > 0
        q = p + 1
        If TlCh(s, q) = "=" Or TlCh(s, q) = " " Then q = q + 1
        If TlIsDigit(TlCh(s, q)) And TlFree(mask, p, q) Then
            num = TlNumAt(s, q)
            If Val(num) >= 1 And Val(num) <= 60 And Not TlIsLetter(TlCh(s, q)) Then
                TlTake(mask, p, q - 1)
                TlZ = CLng(Val(num))
                Exit Function
            End If
        End If
        p = TlWord(s, "z", p + 1)
    Loop
    ' «4z» (the digit before a lone z)
    p = InStr(s, "z")
    Do While p > 1
        If TlIsDigit(Mid(s, p - 1, 1)) And Not TlIsLetter(TlCh(s, p + 1)) Then
            q = p - 1
            Do While q > 1
                If Not TlIsDigit(TlCh(s, q - 1)) Then Exit Do
                q = q - 1
            Loop
            If q = 1 Or Not TlIsAlnum(Mid(s, q - 1, 1)) Then
                num = Mid(s, q, p - q)
                If Val(num) >= 1 And Val(num) <= 60 And TlFree(mask, q, p) Then
                    TlTake(mask, q, p)
                    TlZ = CLng(Val(num))
                    Exit Function
                End If
            End If
        End If
        p = InStr(p + 1, s, "z")
    Loop
    ' the words «зуб», «флейт», «канав»: the number before them («4 зуба», «4-х зубая», «4зуб»), or a word of a count («двухзубая», «трёхзаходная»)
    For Each w In Array("зуб", "флейт", "flute", "канав")
        st = TlK(CStr(w))
        p = InStr(s, st)
        Do While p > 0
            q = p - 1
            If q >= 1 Then
                If Mid(s, q, 1) = " " Or Mid(s, q, 1) = "-" Then q = q - 1
                If q >= 1 Then
                    If Mid(s, q, 1) = "х" And q > 1 Then
                        If Mid(s, q - 1, 1) = "-" Then q = q - 2
                    End If
                End If
                If q >= 1 Then
                    If TlIsDigit(Mid(s, q, 1)) Then
                        num = ""
                        Do While q >= 1
                            If TlIsDigit(Mid(s, q, 1)) Then
                                num = Mid(s, q, 1) & num
                                q = q - 1
                            Else
                                Exit Do
                            End If
                        Loop
                        If (q = 0 Or Not TlIsAlnum(Mid(s, q, 1))) And Val(num) >= 1 And Val(num) <= 60 Then
                            TlTake(mask, q + 1, p + Len(st) - 1)
                            TlZ = CLng(Val(num))
                            Exit Function
                        End If
                    End If
                End If
            End If
            p = InStr(p + 1, s, st)
        Loop
    Next w
    For Each k In Array("одно|1", "двух|2", "трех|3", "четыр|4", "пяти|5", "шести|6", "восьми|8")
        w = Split(CStr(k), "|")
        p = InStr(s, TlK(CStr(w(0))))
        If p > 0 Then
            q = p + Len(CStr(w(0)))
            If Mid(s, q, 3) = TlK("зуб") Or Mid(s, q, 3) = TlK("заход") Or Mid(s, q, 5) = TlK("флейт") Or Mid(s, q, 4) = TlK("канав") Then
                TlZ = CLng(Val(w(1)))
                Exit Function
            End If
        End If
    Next k
End Function

' the first key of the table (items: folded key, name) found in s: the name; "" — none; the span is taken. A key of a long word (five letters and more, no digits: «твердосплав», «карбид») is a stem
' («твердосплавная» too); a code («р6м5», «вк8», «hss», «tin») has to be a whole token
Private Function TlTable(s As String, tab As Variant, ByRef mask As String) As String
    Dim i As Long, p As Long, k As String, stem As Boolean, e As Long
    For i = 0 To UBound(tab)
        k = CStr(tab(i)(0))
        p = InStr(s, k)
        If p > 0 Then
            stem = tab(i)(2)
            Do While p > 0
                e = p + Len(k)
                If (p = 1 Or Not TlIsAlnum(TlCh(s, p - 1))) And (stem Or Not TlIsAlnum(TlCh(s, e))) Then
                    If stem Then e = TlWordEnd(s, p) + 1
                    TlTake(mask, p, e - 1)
                    TlTable = CStr(tab(i)(1))
                    Exit Function
                End If
                p = InStr(p + 1, s, k)
            Loop
        End If
    Next i
End Function

' a key of a long word (five letters and more, no digits: «твердосплав», «карбид») is a stem; a code («р6м5», «вк8», «hss», «tin») has to be a whole token
Private Function TlIsStemKey(k As String) As Boolean
    TlIsStemKey = (Len(k) >= 5 And TlDigits(k) = 0)
End Function

' ================================================================ one name, all of it

' the record of a name: (type, the index of the type, the group, the kind, the size, the diameter, the length, Z, the material, the coating, the standard, the equivalents of the standard,
' the shank / thread / code). An empty type — a name that is not read as a tool.
Function TlRecord(sName As String) As Variant
    TlRecord = TlRecordF(sName, "")
End Function

' TlRecord with the folded name given (sFold; "" — made here). The folded name is taken as it is when it is as long as the name (nothing was joined or dropped in it: the positions of the two agree)
' and has no «М 8» that NormName joins; otherwise the name is put through NormName and Fold
Function TlRecordF(sName As String, sFold As String) As Variant
    Dim orig As String, s As String, mask As String, i As Long, j As Long, p As Long, best As Long, tIdx As Long, tName As String, tGrp As String, tEnd As Long, need As String
    Dim ty As Variant, stems As Variant, ok As Boolean, ins As String, hold As String, thread As String, std As String, shank As String, des As String, z As Long, mat As String, coat As String
    Dim dThread As Double, dThLen As Double, kinds As Variant, kind As String, q As Long, w As Variant, kEnd As Long, tStart As Long, kw As String
    Dim gs() As Long, ge() As Long, gv() As String, ng As Long, dDiam As Double, dLen As Double, sSize As String, g As Long, vals As Variant, pre As String, post As String, dRad As Double, dSign As String
    TlLoad()
    If sFold <> "" And Len(sFold) = Len(sName) And Not (sFold Like "*м #*") Then
        orig = sName
        s = sFold
    Else
        orig = WmsNames.NormName(sName)
        s = WmsNames.Fold(orig)
    End If
    ' a name with no word of a type and with no room for the code of an insert is not read any further (most of the names of a warehouse)
    If Not TlHasType(s) And Not (s Like "*#*#*#*#*") Then
        TlRecordF = Array("", 99, "O", "", "", 0#, 0#, 0, "", "", "", "", "")
        Exit Function
    End If
    mask = Space(Len(s))
    tIdx = -1
    ' the standard, the code of an insert and of a holder first: their numbers are not sizes
    std = TlStandard(s, orig, mask)
    hold = TlHolders(s, orig, mask)
    ins = TlInsert(s, orig, mask)
    ' the type: the earliest word of the name that names one
    ty = gTlTab(0)
    best = 100000
    For i = 0 To UBound(ty)
        stems = ty(i)(0)
        For j = 0 To UBound(stems)
            p = 0
            If InStr(s, CStr(stems(j))) > 0 Then p = TlWord(s, CStr(stems(j)), 1)
            If p > 0 And p < best Then
                ok = True
                need = CStr(ty(i)(3))
                If need = "#insert" Then
                    ok = (ins <> "")
                    If Not ok Then ok = TlHasAny(s, Array("сменн", "режущ", "токарн", "фрезерн", "твердосплав", "отрезн", "резьбов"))
                ElseIf need = "#holder" Then
                    ok = (hold <> "")
                    If Not ok Then ok = TlHasAny(s, Array("цанг", "сверлил", "токар", "кулачк", "быстроз", "гидро", "термо", "станк", "чпу"))
                End If
                If ok Then
                    best = p
                    tIdx = i
                    tName = CStr(ty(i)(1))
                    tGrp = CStr(ty(i)(2))
                End If
            End If
        Next j
    Next i
    ' an insert that has no word of a type is still an insert
    If tIdx < 0 And ins <> "" Then
        If Len(Split(ins, " ")(0)) = 4 Then
            For i = 0 To UBound(ty)
                If CStr(ty(i)(1)) = "Пластина" Then tIdx = i
            Next i
            tName = "Пластина"
            tGrp = "L"
            best = 0
        End If
    End If
    If tIdx < 0 Then
        TlRecordF = Array("", 99, "O", "", "", 0#, 0#, 0, "", "", "", "", "")
        Exit Function
    End If
    tStart = best
    If best > 0 And best < 100000 Then
        tEnd = TlWordEnd(s, best)
        TlTake(mask, best, tEnd)
    End If
    thread = TlThread(s, orig, mask, dThread, dThLen)
    shank = TlShank(s, mask)
    mat = TlTable(s, gTlTab(2), mask)
    coat = TlTable(s, gTlTab(3), mask)
    z = TlZ(s, mask)
    ' the kind: the earliest of the words that are not the type
    kinds = gTlTab(1)
    best = 100000
    For i = 0 To UBound(kinds)
        p = 0
        If InStr(s, CStr(kinds(i))) > 0 Then p = TlWord(s, CStr(kinds(i)), 1)
        If p > 0 And p < best Then
            If TlFree(mask, p, p + 1) Or (p >= tStart And p <= tEnd) Then
                kEnd = TlWordEnd(s, p + Len(CStr(kinds(i))) - 1)
                kw = LCase(Mid(orig, p, kEnd - p + 1))
                ' the kind that the name of the type says already («Расточная головка» — «расточная») is not shown twice
                If InStr(LCase(tName), Left(kw, 6)) = 0 Then
                    best = p
                    kind = kw
                End If
            End If
        End If
    Next i
    ' «м/р»: a tap for the machine and the hand
    If kind = "" And tGrp = "T" Then
        If InStr(s, TlK("м/р")) > 0 Or TlWord(s, TlK("мр"), 1) > 0 Then kind = "машинно-ручной"
    End If
    dSign = TlSign(tGrp)
    ' the numbers that are left: the marks first (the diameter «ø10», «d=10», «диам. 10»; the length «l=100», «длина 100»; the radius «r3»)
    ng = TlGroups(s, mask, gs, ge, gv)
    For g = 0 To ng - 1
        pre = TlBefore(s, gs(g))
        vals = Split(gv(g), Chr(1))
        If TlIsDiamMark(pre) Then
            If dDiam = 0 Then
                dDiam = Val(vals(0))
                If UBound(vals) >= 1 Then dLen = Val(vals(UBound(vals)))
                sSize = "Ø" & TlShow(CStr(vals(0))) & TlRest(vals)
                TlTake(mask, gs(g), ge(g))
            End If
        ElseIf pre = "l" Or pre = TlK("длина") Or pre = TlK("длин") Or pre = TlK("дл") Then
            If dLen = 0 Then
                dLen = Val(vals(0))
                TlTake(mask, gs(g), ge(g))
            End If
        ElseIf pre = "r" Or pre = TlK("радиус") Then
            If dRad = 0 Then
                dRad = Val(vals(0))
                TlTake(mask, gs(g), ge(g))
            End If
        End If
    Next g
    If dDiam = 0 And dThread > 0 And (tGrp = "T" Or tGrp = "L" Or tGrp = "H" Or tGrp = "O") Then
        dDiam = dThread
        If dLen = 0 Then dLen = dThLen
    End If
    ' then the first group of numbers that is a size: several numbers («10х22х72»), a number with «мм», or the first number of a tool that has a diameter
    If dDiam = 0 Then
        For g = 0 To ng - 1
            If TlFree(mask, gs(g), ge(g)) Then
                post = TlAfter(s, ge(g))
                vals = Split(gv(g), Chr(1))
                If Not TlSkipGroup(s, gs(g), ge(g), post) Then
                    ok = (UBound(vals) >= 1 Or post = TlK("мм") Or post = "mm" Or post = TlK("мм."))
                    If Not ok Then ok = (tGrp = "M" Or tGrp = "D" Or tGrp = "R" Or tGrp = "T" Or tGrp = "L" Or tGrp = "H")
                    If ok And InStr(gv(g), "/") = 0 Then
                        If Val(vals(0)) > 0 And Val(vals(0)) <= 5000 Then
                            If tGrp = "S" Then
                                ' a file, a saw blade: the numbers are the width / thickness / length, there is no diameter
                                dLen = Val(vals(UBound(vals)))
                                sSize = TlShow(CStr(vals(0))) & TlRest(vals)
                            ElseIf Val(vals(0)) <= 500 Then
                                dDiam = Val(vals(0))
                                If UBound(vals) >= 1 Then dLen = Val(vals(UBound(vals)))
                                sSize = dSign & TlShow(CStr(vals(0))) & TlRest(vals)
                            Else
                                ok = False
                            End If
                            If ok Then
                                TlTake(mask, gs(g), ge(g))
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next g
    End If
    ' the size as it is shown: the thread of a tap, else the diameter with the lengths, else the radius
    If thread <> "" And (tGrp = "T" Or (sSize = "" And (tGrp = "L" Or tGrp = "H" Or tGrp = "O"))) Then sSize = thread
    If sSize = "" And dRad > 0 Then sSize = "R" & TlShow(Trim(Str(dRad)))
    If sSize = "" And ins <> "" Then sSize = ins
    If sSize = "" And hold <> "" Then sSize = Split(hold, ", ")(0)
    If tGrp = "A" Then dLen = 0
    ' the code, the thread, the shank in one column
    des = IIf(ins <> sSize, ins, "")
    For Each w In Split(hold, ", ")
        If CStr(w) <> "" And CStr(w) <> sSize Then des = des & IIf(des <> "", ", ", "") & CStr(w)
    Next w
    If thread <> "" And thread <> sSize Then des = des & IIf(des <> "", ", ", "") & thread
    If shank <> "" Then des = des & IIf(des <> "", ", ", "") & shank
    TlRecordF = Array(tName, tIdx, tGrp, kind, sSize, dDiam, dLen, z, mat, coat, std, WmsSpec.SpEquiv(std), des)
End Function

' True when the folded name holds the stem of any type
Private Function TlHasType(s As String) As Boolean
    Dim ty As Variant, stems As Variant, i As Long, j As Long
    TlLoad()
    ty = gTlTab(0)
    For i = 0 To UBound(ty)
        stems = ty(i)(0)
        For j = 0 To UBound(stems)
            If InStr(s, CStr(stems(j))) > 0 Then
                TlHasType = True
                Exit Function
            End If
        Next j
    Next i
End Function

' the sign before a diameter in the size text: a rotating tool has a diameter (Ø10), a holder of a lathe, a wheel, a saw blade has the sizes only (16×16×100)
Private Function TlSign(sGrp As String) As String
    If sGrp = "M" Or sGrp = "D" Or sGrp = "T" Or sGrp = "R" Or sGrp = "H" Then TlSign = "Ø"
End Function

' True when s (folded) has any of the words (folded here)
Private Function TlHasAny(s As String, words As Variant) As Boolean
    Dim w As Variant
    For Each w In words
        If InStr(s, TlK(CStr(w))) > 0 Then
            TlHasAny = True
            Exit Function
        End If
    Next w
End Function

' the mark of a diameter before a number: «ø», «d», «диам», «диаметр», «dia»
Private Function TlIsDiamMark(pre As String) As Boolean
    TlIsDiamMark = (pre = "ø" Or pre = "⌀" Or pre = "∅" Or pre = "d" Or pre = "dia" Or pre = TlK("диам") Or pre = TlK("диаметр"))
End Function

' the text of the numbers of a group after the first: "×22×72" for 10,22,72
Private Function TlRest(vals As Variant) As String
    Dim i As Long
    For i = 1 To UBound(vals)
        TlRest = TlRest & "×" & TlShow(CStr(vals(i)))
    Next i
End Function

' the word that stands right before position p (spaces and «=», «:» skipped): "ø", "d", "диам", "l"; "" — none. A point at the end of an abbreviation is not part of it
Private Function TlBefore(s As String, p As Long) As String
    Dim q As Long, a As Long, t As String
    q = p - 1
    Do While q >= 1 And InStr(" =:", Mid(s, q, 1)) > 0
        q = q - 1
    Loop
    If q < 1 Then Exit Function
    If InStr("ø⌀∅", Mid(s, q, 1)) > 0 Then
        TlBefore = Mid(s, q, 1)
        Exit Function
    End If
    If Mid(s, q, 1) = "." Then q = q - 1
    a = q
    Do While a >= 1
        If TlIsLetter(Mid(s, a, 1)) Then
            a = a - 1
        Else
            Exit Do
        End If
    Loop
    If q > a Then TlBefore = Mid(s, a + 1, q - a)
End Function

' the text that stands right after position p (spaces skipped), up to a space: "мм", "°", "hrc"
Private Function TlAfter(s As String, p As Long) As String
    Dim q As Long, a As Long
    q = p + 1
    Do While q <= Len(s) And Mid(s, q, 1) = " "
        q = q + 1
    Loop
    a = q
    Do While q <= Len(s)
        If InStr(" ,;)", Mid(s, q, 1)) > 0 Then Exit Do
        q = q + 1
    Loop
    TlAfter = Mid(s, a, q - a)
End Function

' the numbers that are not sizes: glued to letters (a hardness «45hrc», a grade), an angle, a quantity of a set, a number of a standard
Private Function TlSkipGroup(s As String, a As Long, b As Long, post As String) As Boolean
    Dim pre As String
    If TlIsLetter(TlCh(s, a - 1)) Then
        TlSkipGroup = True
        Exit Function
    End If
    ' «3-х» («3-х зубая», «3-х кулачковый»): a count, not a size
    If TlCh(s, b + 1) = "-" And TlCh(s, b + 2) = "х" And Not TlIsDigit(TlCh(s, b + 3)) Then
        TlSkipGroup = True
        Exit Function
    End If
    If TlIsLetter(TlCh(s, b + 1)) Then
        TlSkipGroup = True
        Exit Function
    End If
    If Left(post, 1) = "°" Or Left(post, 4) = TlK("град") Or Left(post, 2) = TlK("шт") Or Left(post, 1) = "%" Or Left(post, 2) = TlK("кг") Then
        TlSkipGroup = True
        Exit Function
    End If
    ' a hardness: «45 HRC»; in «8 HRC55» the number after the letters is the hardness, not the one before
    If Left(post, 3) = TlK("hrc") And Not TlIsDigit(Mid(post, 4, 1)) Then
        TlSkipGroup = True
        Exit Function
    End If
    pre = TlBefore(s, a)
    If pre = TlK("модуль") Or pre = TlK("шаг") Or pre = TlK("угол") Or pre = "n" Or pre = "no" Or pre = TlK("набор") Or pre = TlK("комплект") Or pre = TlK("уп") Or pre = TlK("партия") Then TlSkipGroup = True
    If TlCh(s, a - 1) = "№" Then TlSkipGroup = True
End Function

' the groups of numbers outside the taken spans: gs(i)..ge(i) the span, gv(i) the numbers (a separator «х» between them is dropped; Chr(1) joins them). Returns the count
Private Function TlGroups(s As String, mask As String, gs() As Long, ge() As Long, gv() As String) As Long
    Dim i As Long, n As Long, c As String, p As Long, q As Long, v As String, t As String, cnt As Long, fr As String
    n = Len(s)
    ReDim gs(40)
    ReDim ge(40)
    ReDim gv(40)
    i = 1
    Do While i <= n
        c = Mid(s, i, 1)
        If TlIsDigit(c) And TlFree(mask, i, i) Then
            p = i
            v = TlNumAt(s, p)
            If TlCh(s, p) = "/" And TlIsDigit(TlCh(s, p + 1)) Then
                q = p + 1
                fr = TlNumAt(s, q)
                v = v & "/" & fr
                p = q
            End If
            ' a range («4-20»: a stepped drill, a boring head): the numbers joined by Chr(2)
            If TlCh(s, p) = "-" And TlIsDigit(TlCh(s, p + 1)) And InStr(v, "/") = 0 Then
                q = p + 1
                t = TlNumAt(s, q)
                If Not TlIsLetter(TlCh(s, q)) Or (TlCh(s, q) = "х" And TlIsDigit(TlCh(s, q + 1))) Then
                    v = v & Chr(2) & t
                    p = q
                End If
            End If
            Do
                q = p
                Do While q <= n And Mid(s, q, 1) = " "
                    q = q + 1
                Loop
                If TlCh(s, q) = "х" Then
                    q = q + 1
                    Do While q <= n And Mid(s, q, 1) = " "
                        q = q + 1
                    Loop
                    If TlIsDigit(TlCh(s, q)) And TlFree(mask, q, q) Then
                        t = TlNumAt(s, q)
                        v = v & Chr(1) & t
                        p = q
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            If cnt <= UBound(gs) Then
                gs(cnt) = i
                ge(cnt) = p - 1
                gv(cnt) = v
                cnt = cnt + 1
            End If
            i = p
        Else
            i = i + 1
        End If
    Loop
    TlGroups = cnt
End Function

' TlRecord, but a name that was read already is not read again (a Collection of this Basic keeps text, not an array: the record is kept as one text and made again from it). sFold — the
' folded name when the caller has it made (a list of names is folded at once): a name with no word of a type and no room for the code of an insert is answered by an empty record without
' being read; the answer is kept too
Function TlRecordC(sName As String, Optional vFold As Variant) As Variant
    Dim v As Variant, p As Variant, r As Variant, t As String, i As Long, sFold As String
    If Not gTlCacheOn Then
        gTlCache = New Collection
        gTlCacheOn = True
    End If
    If TlCacheGet(sName, v) Then
        p = Split(CStr(v), Chr(1))
        ' (the equivalents of the standard are not kept: the sheet «Стандарты» may be edited any time)
        TlRecordC = Array(p(0), CLng(p(1)), p(2), p(3), p(4), Val(p(5)), Val(p(6)), CLng(p(7)), p(8), p(9), p(10), WmsSpec.SpEquiv(CStr(p(10))), p(11))
        Exit Function
    End If
    If Not IsMissing(vFold) Then sFold = CStr(vFold)
    If sFold <> "" And Not TlHasType(sFold) And Not (sFold Like "*#*#*#*#*") Then
        r = Array("", 99, "O", "", "", 0#, 0#, 0, "", "", "", "", "")
    Else
        r = TlRecordF(sName, sFold)
    End If
    t = r(0) & Chr(1) & r(1) & Chr(1) & r(2) & Chr(1) & r(3) & Chr(1) & r(4) & Chr(1) & Trim(Str(r(5))) & Chr(1) & Trim(Str(r(6))) & Chr(1) & r(7) & Chr(1) & r(8) & Chr(1) & r(9) & Chr(1) & r(10) _
        & Chr(1) & r(12)
    On Error Resume Next
    gTlCache.Add(t, sName)
    On Error GoTo 0
    TlRecordC = r
End Function

Private Function TlCacheGet(sName As String, ByRef v As Variant) As Boolean
    On Error GoTo Miss
    v = gTlCache.Item(sName)
    TlCacheGet = True
    Exit Function
Miss:
    TlCacheGet = False
End Function

' the words that stand for the parts of a record in the search («ø10», «10», «10х22х72», «z4», «4z», «er32», «din 338» with its equivalents, the English words of the type)
Function TlTokens(rec As Variant) As String
    Dim s As String, sz As String, k As Variant, p As Long, std As String, eq As Variant
    s = LCase(CStr(rec(RC_TYPE))) & " " & TlEnglish(CStr(rec(RC_TYPE))) & " " & LCase(CStr(rec(RC_KIND))) & " " & LCase(CStr(rec(RC_MAT))) & " " & LCase(CStr(rec(RC_COAT)))
    If rec(RC_DIAM) > 0 Then s = s & " ø" & Replace(Trim(Str(rec(RC_DIAM))), ".", ",") & " d" & Trim(Str(rec(RC_DIAM))) & " " & Trim(Str(rec(RC_DIAM)))
    If rec(RC_LEN) > 0 Then s = s & " l" & Trim(Str(rec(RC_LEN)))
    If rec(RC_Z) > 0 Then s = s & " z" & rec(RC_Z) & " " & rec(RC_Z) & "z"
    sz = LCase(CStr(rec(RC_SIZE)))
    If sz <> "" Then
        s = s & " " & Replace(Replace(sz, "×", "х"), ",", ".")
        For Each k In Split(Replace(Replace(Replace(sz, "ø", ""), "×", Chr(1)), "–", Chr(1)), Chr(1))
            If CStr(k) <> "" Then s = s & " " & Replace(CStr(k), ",", ".")
        Next k
    End If
    sz = LCase(CStr(rec(RC_DES)))
    If sz <> "" Then s = s & " " & sz & " " & Replace(Replace(sz, " ", ""), ",", " ")
    std = LCase(CStr(rec(RC_STD)))
    If std <> "" Then
        p = InStr(std, " ")
        s = s & " " & std & " " & Replace(std, " ", "") & " " & Mid(std, p + 1)
        p = InStr(std, "-")
        If p > 0 Then
            std = Trim(Left(std, p - 1))
            s = s & " " & std & " " & Replace(std, " ", "")
        End If
    End If
    If CStr(rec(RC_EQ)) <> "" Then
        For Each k In Split(CStr(rec(RC_EQ)), ";")
            p = InStr(CStr(k), " ")
            s = s & " " & CStr(k) & " " & Replace(CStr(k), " ", "") & IIf(p > 0, " " & Mid(CStr(k), p + 1), "")
        Next k
    End If
    TlTokens = s
End Function

' ================================================================ the sheet «Инструмент ЧПУ»

' True when the name of a category says that its goods are tools or consumables of a machine («Инструмент ЧПУ», «Режущий инструмент»)
Private Function TlCatSays(sCat As String) As Boolean
    Dim c As String
    c = LCase(sCat)
    TlCatSays = (InStr(c, "чпу") > 0 Or InStr(c, "режущ") > 0)
End Function

' the keys of the goods of «Номенклатура» whose category says «ЧПУ» or «режущий» («a:article» / «n:name» folded, each between two Chr(1)); "" — there are none, then «Номенклатура» is not looked at
Private Function TlGoodsToolKeys() As String
    Dim sh As Object, last As Long, d As Variant, i As Long, art As String, out As String
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_GOODS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, GD_CAT, last).getDataArray()
    For i = 0 To UBound(d)
        If TlCatSays(Trim(CStr(d(i)(GD_CAT)))) Then
            art = Trim(CStr(d(i)(GD_ART)))
            If art <> "" Then
                out = out & Chr(1) & "a:" & WmsNames.Fold(art) & Chr(1)
            ElseIf Trim(CStr(d(i)(GD_NAME))) <> "" Then
                out = out & Chr(1) & "n:" & WmsNames.Fold(WmsNames.NormName(CStr(d(i)(GD_NAME)))) & Chr(1)
            End If
        End If
    Next i
    TlGoodsToolKeys = out
End Function

' True when the name begins with the word of a fastener («Болт крепления фрезы», «Винт для пластины»): that is a fastener, not a tool
Private Function TlFastenerFirst(sName As String) As Boolean
    Dim h As String, a As Variant, k As Variant
    h = LCase(Trim(sName))
    a = Array("болт", "гайк", "шайб", "винт", "саморез", "шуруп", "шпильк", "дюбел", "анкер", "заклепк", "заклёпк", "гвозд", "хомут", "шплинт", "крепеж", "крепёж")
    For Each k In a
        If Left(h, Len(CStr(k))) = CStr(k) Then
            TlFastenerFirst = True
            Exit Function
        End If
    Next k
End Function

' True when the category (folded here) is the one of the fasteners («Метизы», «Крепёж»): those are the business of «Метизы»
Private Function TlFastenerCat(sCat As String) As Boolean
    Dim c As String
    c = WmsNames.Fold(sCat)
    TlFastenerCat = (InStr(c, "метиз") > 0 Or InStr(c, "крепеж") > 0)
End Function

Sub BtnToolRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = ToolRefresh(ToolFind())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' Ctrl+Shift+C — «Инструмент ЧПУ», the cursor in «Найти»
Sub NavTool(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(TL_SHEET, TL_FIND_ROW, 1)
End Sub

' the words typed in the cell «Найти» of the sheet
Function ToolFind() As String
    If gDoc.Sheets.hasByName(TL_SHEET) Then ToolFind = Trim(gDoc.Sheets.getByName(TL_SHEET).getCellByPosition(1, TL_FIND_ROW).getString())
End Function

' a cell of the search or of the filters was changed: the list is made again
Sub OnToolChange(Optional oTarget As Variant)
    Dim ra As Object
    If gTlBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If Not ((ra.Row = TL_FIND_ROW And (ra.Column = 1 Or ra.Column = 6)) Or (ra.Row = TL_FILTER_ROW And ra.Column < TL_COLS)) Then Exit Sub
    WmsInit()
    ToolRefresh(ToolFind())
    Exit Sub
EH:
End Sub

' the sheet is opened: the list is made again when the journal moved since it was made (a new book — the first time)
Sub OnToolFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gTlBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gTlSeq And gTlSeq > 0 Then Exit Sub
    ToolRefresh(ToolFind())
    Exit Sub
EH:
End Sub

' the list of the tools from «Наличия»: one row for a name with an article — the balance of all its EIs, the places, how many EIs have a balance; sorted by type, diameter, length, name.
' sFind — the words of the search (as on «Поиск»: the stem, the Latin letters, «din 338» = «ГОСТ 10902», «-слово»); under it the filters of the columns; «Скрыть нулевые» leaves out the names
' with no balance. Only shows: nothing is posted, nothing is written but the cells of this sheet. OK:<rows> / ERR:…
Function ToolRefresh(sFind As String) As String
    Dim sh As Object, i As Long, j As Long, k As Long, n As Long, g As Variant, rec As Variant, rows() As Variant, hays() As String, nms() As String, cols As Variant, notes() As String
    Dim filt As Variant, fTxt(17) As String, hide As Boolean, scores As Variant, ok As Boolean, out() As Variant, nAllF As Long, q As Double, cat As String
    Dim eqShow As String, key As String, descr As String, typeName As String, coll As Object, wordsOld As Variant, cc As String
    Dim aPl() As String, aCt() As String, aAr() As String, aSr() As String, t0 As Long, tp As String, nmAll() As String, arAll() As String, folds As Variant, arF As Variant, tgood As String
    Dim sF As String, rawH() As String, rawN() As String, hP As Variant, nP As Variant, nCold As Long, bProg As Boolean, rawO() As String, oldF As Variant
    On Error GoTo EH
    gTlBusy = True
    t0 = GetSystemTicks()
    WmsInit()
    If Not gDoc.Sheets.hasByName(TL_SHEET) Then
        ToolRefresh = "ERR:в книге нет листа «" & TL_SHEET & "» — книга не обновлена до 1.2.0"
        gTlBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(TL_SHEET)
    If Not gCtLoaded Then WmsCat.CtLoad()
    WmsCat.GoodsLoad()
    WmsCat.StockGroups()
    tp = "реестр " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the tools among the groups: the name is read as one (a name with no word of a type and no room for the code of an insert is not read at all), or the category that the storekeeper wrote says so
    ' (in the registry, in «Номенклатура»); the fasteners are the business of «Метизы». The names are folded all at once
    ReDim nmAll(IIf(gGrpN > 0, gGrpN - 1, 0))
    For i = 0 To gGrpN - 1
        nmAll(i) = CStr(gGrp(i)(0))
    Next i
    folds = WmsNames.FoldAll(nmAll)
    tgood = TlGoodsToolKeys()
    If tgood <> "" Then
        ReDim arAll(IIf(gGrpN > 0, gGrpN - 1, 0))
        For i = 0 To gGrpN - 1
            arAll(i) = Trim(CStr(gGrp(i)(1)))
        Next i
        arF = WmsNames.FoldAll(arAll)
    End If
    tp = tp & ", свёртка имён " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the names that were not read yet are read now (about 15 ms a tool name): many — the bar of the status line tells that the sheet is at work
    nCold = gGrpN
    If gTlCacheOn Then nCold = gGrpN - gTlCache.Count
    If nCold > 300 Then
        WmsView.ProgressStart("Инструмент ЧПУ: читаю названия…", gGrpN)
        bProg = True
    End If
    n = 0
    ReDim rows(IIf(gGrpN > 0, gGrpN - 1, 0))
    For i = 0 To gGrpN - 1
        If bProg Then
            If (i Mod 40) = 0 Then WmsView.ProgressSet(i)
        End If
        g = gGrp(i)
        sF = CStr(folds(i))
        rec = TlRecordC(CStr(g(0)), sF)
        cat = Trim(CStr(g(6)))
        ok = False
        If CStr(rec(RC_TYPE)) <> "" Then
            If cat <> "" Then
                ok = Not TlFastenerCat(cat)
            Else
                ok = Not TlFastenerFirst(CStr(g(0)))
            End If
        ElseIf cat <> "" Then
            ok = TlCatSays(cat)
        ElseIf tgood <> "" Then
            key = IIf(CStr(arAll(i)) <> "", "a:" & CStr(arF(i)), "n:" & sF)
            ok = (InStr(tgood, Chr(1) & key & Chr(1)) > 0)
        End If
        If ok Then
            rows(n) = Array(g, rec, sF)
            n = n + 1
        End If
    Next i
    If bProg Then
        WmsView.ProgressEnd()
        bProg = False
    End If
    tp = tp & ", отбор " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the search and the filters of the sheet
    filt = sh.getCellRangeByPosition(0, TL_FILTER_ROW, TL_COLS - 1, TL_FILTER_ROW).getDataArray()(0)
    For j = 0 To TL_COLS - 1
        fTxt(j) = Trim(CStr(filt(j)))
    Next j
    hide = (WmsNames.Fold(Trim(sh.getCellByPosition(6, TL_FIND_ROW).getString())) = "да")
    ReDim hays(IIf(n > 0, n - 1, 0))
    ReDim nms(IIf(n > 0, n - 1, 0))
    ReDim notes(IIf(n > 0, n - 1, 0))
    ReDim rawH(IIf(n > 0, n - 1, 0))
    ReDim rawN(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        g = rows(i)(0)
        rec = rows(i)(1)
        typeName = CStr(rec(RC_TYPE))
        If typeName = "" Then typeName = "Прочее"
        rawH(i) = typeName & " " & TlTokens(rec) & " " & g(0) & " " & g(1) & " " & g(5)
        rawN(i) = CStr(g(0)) & " " & g(1)
    Next i
    If n > 0 Then
        hP = WmsSearch.PrepAll(rawH)
        nP = WmsSearch.PrepAll(rawN)
        For i = 0 To n - 1
            hays(i) = " " & Trim(CStr(hP(i))) & " "
            nms(i) = " " & Trim(CStr(nP(i))) & " "
        Next i
    End If
    tp = tp & ", тексты поиска " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the possible doubles: the same type, kind, size, teeth, material, coating, designation and standard (or its equivalent) — under different names; or the same name — under different articles
    coll = New Collection
    For i = 0 To n - 1
        rec = rows(i)(1)
        g = rows(i)(0)
        If CStr(rec(RC_TYPE)) <> "" And CStr(rec(RC_SIZE)) <> "" Then
            key = "p|" & rec(RC_TYPE) & "|" & rec(RC_KIND) & "|" & rec(RC_SIZE) & "|" & rec(RC_Z) & "|" & rec(RC_MAT) & "|" & rec(RC_COAT) & "|" & rec(RC_DES) & "|" & WmsSpec.SpCanon(CStr(rec(RC_STD)))
            WmsCat.DupMark(coll, key, i, notes, rows)
        End If
        key = "n|" & rows(i)(2)
        WmsCat.DupMark(coll, key, i, notes, rows)
    Next i
    tp = tp & ", дубли " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the search with the understanding of «Поиск»; the filters «место:», «категория:», «артикул:» of the query need the columns of the rows
    cols = Empty
    If Mid(WmsSearch.QueryNeeds(sFind), 1, 4) <> "0000" Then
        ReDim aPl(IIf(n > 0, n - 1, 0))
        ReDim aCt(IIf(n > 0, n - 1, 0))
        ReDim aAr(IIf(n > 0, n - 1, 0))
        ReDim aSr(IIf(n > 0, n - 1, 0))
        For i = 0 To n - 1
            g = rows(i)(0)
            aPl(i) = CStr(g(5))
            aAr(i) = CStr(g(1))
            aCt(i) = WmsCat.CategoryOf(CStr(g(0)), CStr(g(1)), CStr(g(6)))
        Next i
        cols = Array(WmsNames.FoldAll(aPl), WmsNames.FoldAll(aCt), WmsNames.FoldAll(aAr), WmsNames.FoldAll(aSr))
    End If
    scores = WmsSearch.MatchScores(sFind, hays, nms, n, cols)
    wordsOld = WmsNames.FoldWords(sFind)
    ' the old way too: every word of the search as a part of the text of the row (a part of an article, of a place: «b-8», «5-1»); the texts folded at once
    If Trim(sFind) <> "" And n > 0 Then
        ReDim rawO(n - 1)
        For i = 0 To n - 1
            g = rows(i)(0)
            rec = rows(i)(1)
            rawO(i) = CStr(g(0)) & " " & g(1) & " " & g(5) & " " & rec(RC_TYPE) & " " & rec(RC_SIZE) & " " & rec(RC_KIND) & " " & rec(RC_MAT) & " " & rec(RC_COAT)
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
        ' the old way too: every word of the search as a part of the text of the row (a part of an article, of a place: «b-8», «5-1»)
        If Not ok And Trim(sFind) <> "" Then ok = WmsNames.HasWords(CStr(oldF(i)), wordsOld)
        If ok And hide And q < 0.0000005 Then ok = False
        If ok Then
            typeName = CStr(rec(RC_TYPE))
            If typeName = "" Then typeName = "Прочее"
            eqShow = UCase(Replace(CStr(rec(RC_EQ)), ";", "; "))
            ' the filters of the columns
            If fTxt(TC_TYPE) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_TYPE), WmsNames.Fold(typeName))
            If fTxt(TC_KIND) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_KIND), WmsNames.Fold(CStr(rec(RC_KIND))))
            If fTxt(TC_SIZE) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_SIZE), WmsNames.Fold(CStr(rec(RC_SIZE))))
            If fTxt(TC_DIAM) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(TC_DIAM), CDbl(rec(RC_DIAM)))
            If fTxt(TC_LEN) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(TC_LEN), CDbl(rec(RC_LEN)))
            If fTxt(TC_Z) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(TC_Z), CDbl(rec(RC_Z)))
            If fTxt(TC_MAT) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_MAT), WmsNames.Fold(CStr(rec(RC_MAT))))
            If fTxt(TC_COAT) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_COAT), WmsNames.Fold(CStr(rec(RC_COAT))))
            If fTxt(TC_STD) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_STD), WmsNames.Fold(rec(RC_STD) & " " & Replace(CStr(rec(RC_EQ)), ";", " ")))
            If fTxt(TC_EQ) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_EQ), WmsNames.Fold(eqShow))
            If fTxt(TC_DES) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_DES), WmsNames.Fold(CStr(rec(RC_DES))))
            If fTxt(TC_NAME) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_NAME), WmsNames.Fold(CStr(g(0))))
            If fTxt(TC_ART) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_ART), WmsNames.Fold(CStr(g(1))))
            If fTxt(TC_QTY) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(TC_QTY), q)
            If fTxt(TC_UNIT) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_UNIT), WmsNames.Fold(CStr(g(2))))
            If fTxt(TC_PLACES) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_PLACES), WmsNames.Fold(CStr(g(5))))
            If fTxt(TC_NEI) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(TC_NEI), CDbl(g(4)))
            If fTxt(TC_NOTE) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(TC_NOTE), WmsNames.Fold(notes(i)))
        End If
        If ok Then
            out(k) = Array(typeName, rec(RC_KIND), rec(RC_SIZE), IIf(rec(RC_DIAM) > 0, rec(RC_DIAM), ""), IIf(rec(RC_LEN) > 0, rec(RC_LEN), ""), IIf(rec(RC_Z) > 0, rec(RC_Z), ""), rec(RC_MAT), rec(RC_COAT), _
                rec(RC_STD), eqShow, rec(RC_DES), g(0), g(1), WmsPack.PkRound(q, 3), g(2), g(5), g(4), notes(i), _
                CDbl(rec(RC_IDX)) * 1000000000000# + CDbl(Int(rec(RC_DIAM) * 100 + 0.5)) * 1000000# + CDbl(Int(rec(RC_LEN) * 10 + 0.5)))
            k = k + 1
        End If
    Next i
    tp = tp & ", фильтры " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the old list out, the new one in
    sh.getCellRangeByPosition(0, TL_ROW0, TL_KEY_COL, TL_ROW0 + TL_MAX).clearContents(1 + 2 + 4 + 16)
    nAllF = n
    If k > TL_MAX Then k = TL_MAX
    If k > 0 Then
        ReDim Preserve out(k - 1)
        sh.getCellRangeByPosition(0, TL_ROW0, TL_KEY_COL, TL_ROW0 + k - 1).setDataArray(out)
        TlSort(sh.getCellRangeByPosition(0, TL_ROW0, TL_KEY_COL, TL_ROW0 + k - 1))
        sh.getCellRangeByPosition(0, TL_ROW0, TL_COLS - 1, TL_ROW0 + k - 1).CharColor = RGB(38, 50, 56)
    End If
    tp = tp & ", запись " & (GetSystemTicks() - t0)
    gTlProf = tp & " (групп " & gGrpN & ", инструмента " & n & ")"
    descr = ""
    If Trim(sFind) <> "" Then descr = " — по словам «" & Trim(sFind) & "»"
    For j = 0 To TL_COLS - 1
        If fTxt(j) <> "" Then
            descr = descr & IIf(descr <> "", "; ", " — ") & "фильтры столбцов"
            Exit For
        End If
    Next j
    sh.getCellByPosition(TL_STAT_COL, TL_FIND_ROW).setString("Инструмента (позиций): " & k & IIf(k < nAllF, " из " & nAllF, "") & descr & ". Список из «Наличия» на " & Format(Now(), "DD.MM.YYYY HH:MM") _
        & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)))
    gTlSeq = CLng(SysNum(SK_LAST_SEQ))
    If gTlSeq = 0 Then gTlSeq = 1
    gTlBusy = False
    ToolRefresh = "OK:" & k
    gUiLastMsg = ToolRefresh
    Exit Function
EH:
    If bProg Then WmsView.ProgressEnd()
    gTlBusy = False
    ToolRefresh = "ERR:список инструмента не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = ToolRefresh
End Function

' the list sorted by the key of the type, the diameter and the length (the hidden column), then by the name
Private Sub TlSort(rng As Object)
    Dim f(1) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = TL_KEY_COL
    f(0).IsAscending = True
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    f(1).Field = TC_NAME
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

' test seams
Function TestToolProf() As String
    TestToolProf = gTlProf
End Function

Function TestToolParts(sName As String, nRep As Long) As String
    Dim orig As String, s As String, mask As String, i As Long, t0 As Long, out As String, d1 As Double, d2 As Double, ga As Variant, gs() As Long, ge() As Long, gv() As String, ng As Long, x As String
    TlLoad()
    orig = WmsNames.NormName(sName)
    s = WmsNames.Fold(orig)
    mask = Space(Len(s))
    t0 = GetSystemTicks()
    For i = 1 To nRep
        x = WmsNames.NormName(sName)
    Next i
    out = out & "NormName " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        x = WmsNames.Fold(orig)
    Next i
    out = out & ", Fold " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        x = TlStandard(s, orig, mask)
    Next i
    out = out & ", Standard " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        x = TlHolders(s, orig, mask)
    Next i
    out = out & ", Holders " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        x = TlInsert(s, orig, mask)
    Next i
    out = out & ", Insert " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        d1 = 0
        d2 = 0
        mask = Space(Len(s))
        x = TlThread(s, orig, mask, d1, d2)
    Next i
    out = out & ", Thread " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        x = TlShank(s, mask)
    Next i
    out = out & ", Shank " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        x = TlTable(s, gTlTab(2), mask)
        x = TlTable(s, gTlTab(3), mask)
    Next i
    out = out & ", Mat+Coat " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        d1 = TlZ(s, mask)
    Next i
    out = out & ", Z " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        mask = Space(Len(s))
        ng = TlGroups(s, mask, gs, ge, gv)
    Next i
    out = out & ", Groups " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        d1 = TlHasType(s)
    Next i
    out = out & ", HasType " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    For i = 1 To nRep
        ga = TlRecord(sName)
    Next i
    out = out & ", ВСЁ TlRecord " & (GetSystemTicks() - t0) & " мс на " & nRep
    TestToolParts = out
End Function

Function TestToolSteps() As String
    Dim i As Long, g As Variant, rec As Variant, t0 As Long, ta As Long, tb As Long, tc As Long, td As Long, n As Long, c As String, f As String
    WmsInit()
    If Not gCtLoaded Then WmsCat.CtLoad()
    WmsCat.GoodsLoad()
    WmsCat.StockGroups()
    t0 = GetSystemTicks()
    For i = 0 To gGrpN - 1
        g = gGrp(i)
        rec = TlRecordC(CStr(g(0)))
        If CStr(rec(RC_TYPE)) <> "" Then n = n + 1
    Next i
    ta = GetSystemTicks() - t0
    t0 = GetSystemTicks()
    For i = 0 To gGrpN - 1
        g = gGrp(i)
        rec = TlRecordC(CStr(g(0)))
    Next i
    tb = GetSystemTicks() - t0
    t0 = GetSystemTicks()
    For i = 0 To gGrpN - 1 Step 3
        g = gGrp(i)
        c = WmsCat.CategoryOf(CStr(g(0)), CStr(g(1)), CStr(g(6)))
    Next i
    tc = GetSystemTicks() - t0
    t0 = GetSystemTicks()
    For i = 0 To gGrpN - 1
        g = gGrp(i)
        f = WmsNames.Fold(CStr(g(0)))
    Next i
    td = GetSystemTicks() - t0
    TestToolSteps = "TlRecordC холодный " & ta & ", тёплый " & tb & "; CategoryOf (1/3 групп) " & tc & "; Fold всех " & td & "; инструмента " & n
End Function

Function TestToolParse(sName As String) As String
    Dim r As Variant
    TlLoad()
    r = TlRecord(sName)
    TestToolParse = r(RC_TYPE) & "|" & r(RC_KIND) & "|" & r(RC_SIZE) & "|" & Trim(Str(r(RC_DIAM))) & "|" & Trim(Str(r(RC_LEN))) & "|" & r(RC_Z) & "|" & r(RC_MAT) & "|" & r(RC_COAT) & "|" & r(RC_STD) & "|" & r(RC_DES)
End Function
