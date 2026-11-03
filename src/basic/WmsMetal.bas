' WmsMetal — 1.3.0: «Металл». What a name of a rolled metal says: the kind («Лист», «Круг», «Труба профильная», «Уголок», «Швеллер» …), the finish («г/к», «х/к», «оцинк.», «бесшовная»), the material («Сталь», «Нержавеющая
' сталь», «Алюминий», «Медь», «Латунь» …) and the grade («Ст3», «09Г2С», «AISI 304», «АМг6»), the size (thickness, width, Ø or side, length), the standard (ГОСТ 19903-2015) — and the weight that follows from the size.
' The sheet «Металл» is made from «Наличие» like «Метизы» and «Инструмент ЧПУ» (one row per name with an article: the balance of all its EIs, the places, the EIs themselves with their lengths — the pieces
' and the offcuts) and with the same search and filters of the columns. Only shows: nothing is posted. What the name does not tell stays empty — nothing is guessed except the weight: it is counted from the size and
' the density of the material (steel when the material is not told) and is always marked as approximate («≈»). Library WmsPlus.
Option Explicit

Private Const MT_SHEET = "Металл"
Private Const MT_FIND_ROW = 1
Private Const MT_FILTER_ROW = 2
Private Const MT_HEAD = 3
Private Const MT_ROW0 = 4
Private Const MT_COLS = 20
Private Const MT_KEY_COL = 20
Private Const MT_STAT_COL = 7
Private Const MT_MAX = 5000

' the columns of the sheet
Private Const MC_KIND = 0
Private Const MC_FIN = 1
Private Const MC_MAT = 2
Private Const MC_GRADE = 3
Private Const MC_SIZE = 4
Private Const MC_T = 5
Private Const MC_W = 6
Private Const MC_D = 7
Private Const MC_L = 8
Private Const MC_STD = 9
Private Const MC_NAME = 10
Private Const MC_ART = 11
Private Const MC_QTY = 12
Private Const MC_UNIT = 13
Private Const MC_LEN = 14
Private Const MC_MASS = 15
Private Const MC_PLACES = 16
Private Const MC_NEI = 17
Private Const MC_NOTE = 18
Private Const MC_EIS = 19

' the record of a name
Private Const MR_KIND = 0
Private Const MR_IDX = 1
Private Const MR_FIN = 2
Private Const MR_MAT = 3
Private Const MR_GRADE = 4
Private Const MR_SIZE = 5
Private Const MR_T = 6
Private Const MR_W = 7
Private Const MR_D = 8
Private Const MR_L = 9
Private Const MR_STD = 10
Private Const MR_KGM = 11
Private Const MR_KGPC = 12
Private Const MR_LINEAR = 13
Private Const MR_LAST = 13

Global gMtInit As Boolean
Global gMtTab As Variant
Global gMtCache As Object
Global gMtCacheOn As Boolean
Global gMtSeq As Long
Global gMtBusy As Boolean
Global gMtProf As String
Global gMtBadNames As String
Global gMtAisi As String

' the numbers of the name being read (at most 30): value, first and last character, unit (0 none, 1 mm, 2 m, 3 kg, 4 cm), mark (D diameter, T thickness, L length), chain (0 alone) and the place in it
Private mnN As Long
Private mnVal(29) As Double
Private mnPos(29) As Long
Private mnEnd(29) As Long
Private mnUnit(29) As Integer
Private mnMark(29) As String
Private mnChain(29) As Long
Private mnLink(29) As Long
Private mnUsed(29) As Boolean
Private mnChains As Long

' the context of the name being read
Private mcBeam As Boolean
Private mcRebar As Boolean
Private mcSingle As Boolean
Private mcWeakOk As Boolean
Private mcStainless As Boolean
Private mcAlu As Boolean

' ================================================================ the tables

' kinds: stems (folded, at the beginning of a word) | name | order | need: "" always, "e" the name must tell a metal (a grade or a word of a metal), "s" a metal or a chain of three sizes (a sheet), "t" a metal or a chain
' of two sizes (a strip), "w" a whole word, "c" a class of a rebar, "b" a beam, "p" a profile with the sizes
Private Function MtKindDefs() As Variant
    MtKindDefs = Array( _
        "лист|Лист|1|s", "плит|Плита|2|s", "полос|Полоса|3|t", "лент|Лента|4|e", "шина|Шина|5|e", _
        "круг|Круг|6|", "пруток,прутк,прутьев|Пруток|7|", "квадрат|Квадрат|8|w", "шестигран|Шестигранник|9|", "катанк|Катанка|10|", "проволок|Проволока|11|", "арматур|Арматура|12|c", _
        "труб|Труба|13|", "уголок,уголк|Уголок|14|", "швеллер|Швеллер|15|", "двутавр|Двутавр|16|", "балк|Балка|17|b", "профил|Профиль|18|p", _
        "заготовк|Заготовка|19|e", "поковк|Поковка|20|e", "отливк|Отливка|21|e", "слитк|Слиток|22|e", "металлопрокат,прокат|Прочее|30|e")
End Function

' the words that tell that a name is not a rolled metal even when it has the word of a kind («лист ДСП», «труба ПВХ», «круг отрезной», «ключ шестигранный»)
Private Function MtBlockDefs() As Variant
    MtBlockDefs = Array("болт", "винт", "гайк", "шайб", "саморез", "шуруп", "шпильк", "анкер", "заклепк", "дюбел", "гвозд", "отрезн", "шлифов", "лепестк", "зачистн", "алмазн", "абразив", "полировальн", _
        "ключ", "биты", "бита", "головк", "пластик", "пвх", "пнд", "полиэтилен", "полипроп", "фанер", "дсп", "мдф", "двп", "гипс", "стеклян", "оргстекл", "поликарбонат", "картон", "бумаг", "ткан", "резин", "силикон", _
        "пенопл", "поролон", "дерев", "сосн", "березов", "ламинат", "кабел", "электрод", "мебельн", "декоратив", "светодиод", "гкл", "потолочн", "оконн", "дверн", "термоусад", "гофр", "фольгирован", "наждач", "войлок", "osb", "осб", "цсп", "шифер", "стекл")
End Function

' the words that tell a metal: the stems (folded) at the beginning of a word
Private Function MtMetalWords() As Variant
    MtMetalWords = Array("сталь", "стальн", "нерж", "алюмин", "дюрал", "силумин", "медь", "медн", "латун", "бронз", "титан", "чугун", "оцинк", "г/к", "х/к", "горячекат", "холоднокат", "метал", "сплав", "н/ж", "жесть", "aisi", "рифл", "чечевиц", "просечн", "пвл", "бесшов", "электросвар", "водогаз", "вгп", "калибров")
End Function

' the finish: key (folded, anywhere in the name) | name
Private Function MtFinishDefs() As Variant
    MtFinishDefs = Array("горячекат|г/к", "холоднокат|х/к", "г/к|г/к", "х/к|х/к", "=гк|г/к", "=хк|х/к", "оцинк|оцинк.", "рифл|рифлёный", "чечевиц|рифлёный", "просечн|просечно-вытяжной", "=пвл|просечно-вытяжной", "перфор|перфорированный", _
        "бесшов|бесшовная", "=б/ш|бесшовная", "холоднодеф|х/д", "горячедеф|г/д", "электросвар|электросварная", "=э/с|электросварная", "водогаз|ВГП", "=вгп|ВГП", "калибров|калиброванный", "серебрянк|шлифованный", "шлифован|шлифованный", _
        "полиров|полированный", "матов|матовый", "зеркал|зеркальный", "равнополоч|равнополочный", "неравнополоч|неравнополочный", "анодир|анодированный", "окрашен|окрашенный", "в бухт|в бухте")
End Function

' the groups of materials: the density (g/cm³)
Private Function MtDensity(sMat As String) As Double
    Select Case sMat
    Case "Нержавеющая сталь": MtDensity = 7.9
    Case "Алюминий": MtDensity = 2.7
    Case "Медь": MtDensity = 8.94
    Case "Латунь": MtDensity = 8.5
    Case "Бронза": MtDensity = 8.8
    Case "Титан": MtDensity = 4.5
    Case "Чугун": MtDensity = 7.2
    Case Else: MtDensity = 7.85
    End Select
End Function

' the tables folded once
Private Sub MtLoad()
    Dim d As Variant, i As Long, p As Variant, a As Variant, j As Long
    If gMtInit Then Exit Sub
    Dim kinds() As Variant, block() As String, words() As String, fin() As Variant
    d = MtKindDefs()
    ReDim kinds(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        a = Split(CStr(p(0)), ",")
        For j = 0 To UBound(a)
            a(j) = WmsNames.Fold(CStr(a(j)))
        Next j
        kinds(i) = Array(a, CStr(p(1)), CLng(p(2)), CStr(p(3)))
    Next i
    d = MtBlockDefs()
    ReDim block(UBound(d))
    For i = 0 To UBound(d)
        block(i) = WmsNames.Fold(CStr(d(i)))
    Next i
    d = MtMetalWords()
    ReDim words(UBound(d))
    For i = 0 To UBound(d)
        words(i) = WmsNames.Fold(CStr(d(i)))
    Next i
    d = MtFinishDefs()
    ReDim fin(UBound(d))
    For i = 0 To UBound(d)
        p = Split(CStr(d(i)), "|")
        fin(i) = Array(WmsNames.Fold(CStr(p(0))), CStr(p(1)))
    Next i
    gMtAisi = WmsNames.Fold("aisi")
    gMtTab = Array(kinds, block, words, fin)
    gMtInit = True
End Sub

' ================================================================ the characters

Private Function MtIsDigit(c As String) As Boolean
    MtIsDigit = (Len(c) = 1 And c >= "0" And c <= "9")
End Function

Private Function MtIsLetter(c As String) As Boolean
    MtIsLetter = (Len(c) = 1) And ((c >= "a" And c <= "z") Or (c >= "а" And c <= "я") Or c = "ё")
End Function

Private Function MtIsAlnum(c As String) As Boolean
    MtIsAlnum = MtIsDigit(c) Or MtIsLetter(c)
End Function

' the character at position p of s ("" outside)
Private Function MtCh(s As String, p As Long) As String
    If p >= 1 And p <= Len(s) Then MtCh = Mid(s, p, 1)
End Function

' the first position of the key at the beginning of a word, from start; 0 — none
Private Function MtWord(s As String, key As String, start As Long) As Long
    Dim p As Long
    If key = "" Then Exit Function
    p = InStr(start, s, key, 0)
    Do While p > 0
        If p = 1 Then
            MtWord = p
            Exit Function
        ElseIf Not MtIsAlnum(Mid(s, p - 1, 1)) Then
            MtWord = p
            Exit Function
        End If
        p = InStr(p + 1, s, key, 0)
    Loop
End Function

' the number at position p (digits, a point followed by a digit): its text; p is moved past it
Private Function MtNumAt(s As String, ByRef p As Long) As String
    Dim t As String, c As String
    Do While p <= Len(s)
        c = Mid(s, p, 1)
        If MtIsDigit(c) Then
            If Len(t) < 15 Then t = t & c
            p = p + 1
        ElseIf (c = "." Or c = ",") And t <> "" Then
            If MtIsDigit(MtCh(s, p + 1)) Then
                If Len(t) < 15 Then t = t & "."
                p = p + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    MtNumAt = t
End Function

' the number as it is written for people (a comma for a point, no trailing zeros)
Private Function MtShow(x As Double) As String
    Dim t As String
    t = Trim(Str(Int(x * 1000 + 0.5) / 1000))
    MtShow = Replace(t, ".", ",")
End Function

' the span a..b of the mask is taken
Private Sub MtTake(ByRef mask As String, a As Long, b As Long)
    If a < 1 Then a = 1
    If b > Len(mask) Then b = Len(mask)
    If b >= a Then mask = Left(mask, a - 1) & String(b - a + 1, "x") & Mid(mask, b + 1)
End Sub

' the text without the signs round it
Private Function MtTrimTok(ByVal t As String) As String
    Do While Len(t) > 0
        If InStr(".,-/:;=", Right(t, 1)) > 0 Then t = Left(t, Len(t) - 1) Else Exit Do
    Loop
    Do While Len(t) > 0
        If InStr(".,-/:;=", Left(t, 1)) > 0 Then t = Mid(t, 2) Else Exit Do
    Loop
    MtTrimTok = t
End Function

' ================================================================ the kind, the metal words, the finish, the standard

' True when the folded name has a word that excludes it
Private Function MtBlocked(s As String) As Boolean
    Dim b As Variant, i As Long
    b = gMtTab(1)
    For i = 0 To UBound(b)
        If InStr(1, s, CStr(b(i)), 0) > 0 Then
            If MtWord(s, CStr(b(i)), 1) > 0 Then
                MtBlocked = True
                Exit Function
            End If
        End If
    Next i
End Function

' True when the name has a word that tells a metal
Private Function MtHasMetalWord(s As String) As Boolean
    Dim w As Variant, i As Long
    w = gMtTab(2)
    For i = 0 To UBound(w)
        If InStr(1, s, CStr(w(i)), 0) > 0 Then
            If MtWord(s, CStr(w(i)), 1) > 0 Then
                MtHasMetalWord = True
                Exit Function
            End If
        End If
    Next i
End Function

' a stem at p is a whole word (or the word with a short case ending: «квадрата», «квадраты»)
Private Function MtWholeWord(s As String, p As Long, stem As String) As Boolean
    Dim e As Long, c As String
    e = p + Len(stem)
    c = MtCh(s, e)
    If c = "" Or Not MtIsLetter(c) Then
        MtWholeWord = True
    ElseIf InStr("аыуеи", c) > 0 Then
        MtWholeWord = Not MtIsLetter(MtCh(s, e + 1))
    End If
End Function

' the earliest word of a kind: its name, order and position (0 — none); the words of "need" that are not met at the position are skipped
Private Function MtFindKind(s As String, ByRef idx As Long, ByRef pos As Long, ByRef need As String) As String
    Dim defs As Variant, i As Long, j As Long, stems As Variant, p As Long, best As Long, ok As Boolean
    defs = gMtTab(0)
    best = 100000
    For i = 0 To UBound(defs)
        stems = defs(i)(0)
        For j = 0 To UBound(stems)
            p = 0
            If InStr(1, s, CStr(stems(j)), 0) > 0 Then p = MtWord(s, CStr(stems(j)), 1)
            If p > 0 And p < best Then
                ok = True
                If CStr(defs(i)(3)) = "w" Then ok = MtWholeWord(s, p, CStr(stems(j)))
                If CStr(defs(i)(1)) = "Шина" Then ok = MtWholeWord(s, p, CStr(stems(j)))
                If ok Then
                    best = p
                    MtFindKind = CStr(defs(i)(1))
                    idx = CLng(defs(i)(2))
                    need = CStr(defs(i)(3))
                End If
            End If
        Next j
    Next i
    If best < 100000 Then pos = best
End Function

' the finish of the name: up to two names, «г/к, оцинк.»
Private Function MtFinish(s As String) As String
    Dim f As Variant, i As Long, n As Long, out As String, lab As String, key As String, found As Boolean, p As Long
    f = gMtTab(3)
    For i = 0 To UBound(f)
        key = CStr(f(i)(0))
        found = False
        If Left(key, 1) = "=" Then
            ' a short abbreviation counts only as a whole word («гк», «вгп», «пвл»)
            key = Mid(key, 2)
            p = MtWord(s, key, 1)
            If p > 0 Then found = Not MtIsAlnum(MtCh(s, p + Len(key)))
        Else
            found = (InStr(1, s, key, 0) > 0)
        End If
        If found Then
            lab = CStr(f(i)(1))
            If InStr(", " & out & ",", ", " & lab & ",") = 0 Then
                out = out & IIf(out <> "", ", ", "") & lab
                n = n + 1
                If n >= 2 Then Exit For
            End If
        End If
    Next i
    MtFinish = out
End Function

' a standard: «ГОСТ 19903-2015», «ТУ 14-3-377-75», «DIN 1017»; the span is taken
Private Function MtStandard(s As String, ByRef mask As String) As String
    Dim keys As Variant, k As Variant, p As Long, q As Long, a As Long, c As String, t As String, key As String
    keys = Array("гост", "ту", "din", "iso", "en", "astm", "ост", "сто")
    For Each k In keys
        key = WmsNames.Fold(CStr(k))
        If InStr(1, s, key, 0) > 0 Then
            p = MtWord(s, key, 1)
            Do While p > 0
                q = p + Len(key)
                ' «гост р 8732» — the letter of the national standard; the signs between the word and the number
                Do While q <= Len(s)
                    c = Mid(s, q, 1)
                    If c = " " Or c = "." Or c = "№" Or c = "-" Then
                        q = q + 1
                    ElseIf (c = "р" Or c = "r") And Not MtIsAlnum(MtCh(s, q + 1)) Then
                        q = q + 1
                    Else
                        Exit Do
                    End If
                Loop
                If MtIsDigit(MtCh(s, q)) Then
                    a = q
                    Do While q <= Len(s)
                        c = Mid(s, q, 1)
                        If MtIsDigit(c) Or c = "-" Or c = "." Or c = "/" Then
                            q = q + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    t = MtTrimTok(Mid(s, a, q - a))
                    If Len(t) >= 2 Then
                        MtTake(mask, p, a + Len(t) - 1)
                        MtStandard = UCase(CStr(k)) & " " & t
                        Exit Function
                    End If
                End If
                p = MtWord(s, key, p + 1)
            Loop
        End If
    Next k
End Function

' ================================================================ the material and the grade

' a separator between two words of a name
Private Function MtIsSep(c As String) As Boolean
    MtIsSep = (c = " " Or c = "," Or c = ";" Or c = "(" Or c = ")" Or c = "[" Or c = "]" Or c = Chr(34) Or c = "«" Or c = "»" Or c = "|")
End Function

' the text is only «х», digits and points («40х20х2», «20х13»)
Private Function MtOnlyXDigits(t As String) As Boolean
    Dim i As Long, c As String, nd As Long
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If MtIsDigit(c) Then
            nd = nd + 1
        ElseIf c <> "х" And c <> "." Then
            Exit Function
        End If
    Next i
    MtOnlyXDigits = (nd > 0)
End Function

' the letters after the number of a beam or of a channel («20б1», «25ш1», «30к2», «12п», «10у», «14э»)
Private Function MtBeamSeries(rest As String) As Boolean
    Dim c As String
    c = Left(rest, 1)
    If c = "б" Or c = "ш" Or c = "к" Or c = "ц" Or c = "п" Or c = "у" Or c = "э" Or c = "л" Or c = "в" Then
        MtBeamSeries = (Len(rest) <= 3)
    End If
End Function

' «х18н10т»: the chromium of two digits from 10 to 29 — a stainless steel
Private Function MtStainlessRest(rest As String) As Boolean
    Dim v As Long
    If Left(rest, 1) = "х" And MtIsDigit(MtCh(rest, 2)) And MtIsDigit(MtCh(rest, 3)) Then
        v = CLng(Val(Mid(rest, 2, 2)))
        MtStainlessRest = (v >= 10 And v <= 29)
    End If
End Function

' the digits at the beginning of t from position k (at most 4): their text; k is moved past them
Private Function MtDigitsAt(t As String, ByRef k As Long) As String
    Dim d As String
    Do While MtIsDigit(MtCh(t, k)) And Len(d) < 4
        d = d & MtCh(t, k)
        k = k + 1
    Loop
    MtDigitsAt = d
End Function

' one word of the name: is it the mark of a metal? mat — the group of the material, grade — the mark as it is written
Private Function MtGradeTok(ByVal t As String, ByRef mat As String, ByRef grade As String) As Boolean
    Dim n As Long, c1 As String, i As Long, d As String, rest As String, ch As String, ok As Boolean, hasMain As Boolean, k As Long, sfx As String
    t = MtTrimTok(t)
    n = Len(t)
    If n < 2 Or n > 24 Then Exit Function
    c1 = Left(t, 1)
    ' the western marks: 304, 316l (stainless: when the name says «нерж» or «aisi»), 6061, 5083 (aluminium: when it says «алюмин»)
    If mcStainless Then
        If t = "201" Or t = "202" Or t = "301" Or t = "303" Or t = "304" Or t = "304l" Or t = "309" Or t = "310" Or t = "316" Or t = "316l" Or t = "316ti" Or t = "321" Or t = "409" Or t = "410" Or t = "420" Or t = "430" Or t = "431" Then
            grade = "AISI " & UCase(t)
            mat = "Нержавеющая сталь"
            MtGradeTok = True
            Exit Function
        End If
    End If
    If mcAlu Then
        If t = "6061" Or t = "6063" Or t = "6082" Or t = "5052" Or t = "5083" Or t = "5754" Or t = "3003" Or t = "2024" Or t = "7075" Then
            grade = t
            mat = "Алюминий"
            MtGradeTok = True
            Exit Function
        End If
    End If
    If MtIsDigit(c1) Then
        ' 09г2с, 40х, 12х18н10т, 65г, 08кп: the number of the steel (1–3 digits) and the letters of the alloy
        i = 1
        Do While i <= n
            If MtIsDigit(Mid(t, i, 1)) Then i = i + 1 Else Exit Do
        Loop
        d = Left(t, i - 1)
        rest = Mid(t, i)
        If Len(d) > 3 Or rest = "" Then Exit Function
        If Not MtIsLetter(Left(rest, 1)) Then Exit Function
        If rest = "мм" Or rest = "м" Or rest = "см" Or rest = "кг" Or rest = "шт" Or rest = "т" Or rest = "л" Or rest = "мкм" Or rest = "мин" Or rest = "ч" Or rest = "г" Or rest = "мп" Or rest = "пм" Or rest = "тн" Or rest = "гр" Or rest = "дм" Or rest = "мг" Or rest = "мс" Or rest = "нм" Then Exit Function
        If MtOnlyXDigits(rest) Then Exit Function
        If mcBeam Then
            If MtBeamSeries(rest) Then Exit Function
        End If
        If rest = "кп" Or rest = "пс" Or rest = "сп" Or rest = "к" Then
            grade = UCase(t)
            mat = "Сталь"
            MtGradeTok = True
            Exit Function
        End If
        ok = True
        hasMain = False
        For i = 1 To Len(rest)
            ch = Mid(rest, i, 1)
            If MtIsDigit(ch) Or ch = "-" Or ch = "." Then
                ok = ok
            ElseIf InStr(1, "хнгсмтфбдкавюрцчп", ch, 0) > 0 Then
                If InStr(1, "хнгсмтфб", ch, 0) > 0 Then hasMain = True
            Else
                ok = False
            End If
        Next i
        If Not (ok And hasMain) Then Exit Function
        grade = UCase(t)
        If MtStainlessRest(rest) Then mat = "Нержавеющая сталь" Else mat = "Сталь"
        MtGradeTok = True
        Exit Function
    End If
    Select Case c1
    Case "с"
        If Left(t, 2) = "ст" Then
            k = 3
            If MtCh(t, 3) = "." Then k = 4
            d = MtDigitsAt(t, k)
            If d <> "" And Len(d) <= 3 Then
                sfx = Mid(t, k)
                ok = (sfx = "" Or sfx = "сп" Or sfx = "пс" Or sfx = "кп" Or sfx = "г" Or sfx = "гпс" Or sfx = "гсп")
                If Not ok And Len(sfx) = 3 Then
                    ' «сп5», «пс2»: the number of the group of the steel
                    If MtIsDigit(Right(sfx, 1)) Then
                        ch = Left(sfx, 2)
                        ok = (ch = "сп" Or ch = "пс" Or ch = "кп")
                    End If
                End If
                If ok Then
                    grade = "Ст" & d & sfx
                    mat = "Сталь"
                    MtGradeTok = True
                End If
            End If
        ElseIf Left(t, 2) = "св" Then
            k = 3
            If MtCh(t, 3) = "-" Then k = 4
            If MtIsDigit(MtCh(t, k)) Then
                grade = "Св-" & UCase(Mid(t, k))
                mat = "Сталь"
                MtGradeTok = True
            End If
        ElseIf Left(t, 2) = "сч" Then
            If MtIsDigit(MtCh(t, 3)) Then
                grade = UCase(t)
                mat = "Чугун"
                MtGradeTok = True
            End If
        End If
    Case "у"
        k = 2
        d = MtDigitsAt(t, k)
        If d <> "" And Len(d) <= 2 Then
            sfx = Mid(t, k)
            If sfx = "" Or sfx = "а" Then
                grade = UCase(t)
                mat = "Сталь"
                MtGradeTok = True
            End If
        End If
    Case "р"
        If t = "р6м5" Or t = "р6м5к5" Or t = "р6м5ф3" Or t = "р6м3" Or t = "р18" Or t = "р12" Or t = "р9" Or t = "р9к5" Or t = "р10к5ф5" Then
            grade = UCase(t)
            mat = "Сталь"
            MtGradeTok = True
        End If
    Case "х"
        If t = "х12" Or t = "х12м" Or t = "х12мф" Or t = "х12ф1" Or t = "х6вф" Or t = "хвг" Or t = "х12вмф" Then
            grade = UCase(t)
            mat = "Сталь"
            MtGradeTok = True
        End If
    Case "ш"
        If Left(t, 2) = "шх" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Сталь"
            MtGradeTok = True
        End If
    Case "а"
        If Left(t, 2) = "ад" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Алюминий"
            MtGradeTok = True
        ElseIf Left(t, 3) = "амг" And MtIsDigit(MtCh(t, 4)) Then
            grade = "АМг" & UCase(Mid(t, 4))
            mat = "Алюминий"
            MtGradeTok = True
        ElseIf t = "амц" Then
            grade = "АМц"
            mat = "Алюминий"
            MtGradeTok = True
        ElseIf Left(t, 2) = "ак" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Алюминий"
            MtGradeTok = True
        ElseIf t = "ав" Or t = "авт" Or t = "авт1" Then
            grade = UCase(t)
            mat = "Алюминий"
            MtGradeTok = True
        ElseIf mcRebar Then
            ' the classes of a rebar: а240, а400, а500с, а1000, ат800, а1…а4, а-iii
            If Left(t, 2) = "ат" And MtIsDigit(MtCh(t, 3)) Then
                grade = UCase(t)
                mat = "Сталь"
                MtGradeTok = True
            ElseIf MtIsDigit(MtCh(t, 2)) Then
                k = 2
                d = MtDigitsAt(t, k)
                sfx = Mid(t, k)
                If sfx = "" Or sfx = "с" Or sfx = "к" Then
                    grade = UCase(t)
                    mat = "Сталь"
                    MtGradeTok = True
                End If
            ElseIf Left(t, 2) = "а-" And Len(t) <= 6 Then
                grade = "А-" & UCase(Mid(t, 3))
                mat = "Сталь"
                MtGradeTok = True
            End If
        End If
    Case "д"
        If MtIsDigit(MtCh(t, 2)) And n <= 6 Then
            k = 2
            d = MtDigitsAt(t, k)
            sfx = Mid(t, k)
            If Len(d) <= 2 And ((sfx = "" And (d = "1" Or d = "16" Or d = "18" Or d = "19")) Or sfx = "т" Or sfx = "ат" Or sfx = "чт" Or sfx = "н" Or sfx = "б" Or sfx = "ч" Or sfx = "т1") Then
                grade = UCase(t)
                mat = "Алюминий"
                MtGradeTok = True
            End If
        End If
    Case "в"
        If Left(t, 2) = "вт" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Титан"
            MtGradeTok = True
        ElseIf Left(t, 2) = "вч" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Чугун"
            MtGradeTok = True
        ElseIf Left(t, 2) = "в9" And MtIsDigit(MtCh(t, 3)) And n <= 5 Then
            grade = UCase(t)
            mat = "Алюминий"
            MtGradeTok = True
        End If
    Case "о"
        If Left(t, 2) = "от" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Титан"
            MtGradeTok = True
        End If
    Case "п"
        If Left(t, 2) = "пт" And (MtIsDigit(MtCh(t, 3)) Or MtCh(t, 3) = "-") Then
            grade = UCase(t)
            mat = "Титан"
            MtGradeTok = True
        End If
    Case "к"
        If Left(t, 2) = "кч" And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Чугун"
            MtGradeTok = True
        End If
    Case "м"
        If mcWeakOk Then
            If t = "м0" Or t = "м00" Or t = "м1" Or t = "м1р" Or t = "м1ф" Or t = "м2" Or t = "м2р" Or t = "м3" Or t = "м3р" Or t = "мб" Then
                grade = UCase(t)
                mat = "Медь"
                MtGradeTok = True
            End If
        End If
        If Not MtGradeTok Then
            If Left(t, 2) = "мн" And MtIsDigit(MtCh(t, 3)) Then
                grade = UCase(t)
                mat = "Медь"
                MtGradeTok = True
            End If
        End If
    Case "л"
        If MtIsDigit(MtCh(t, 2)) And MtIsDigit(MtCh(t, 3)) And n <= 3 Then
            grade = UCase(t)
            mat = "Латунь"
            MtGradeTok = True
        ElseIf (Left(t, 2) = "лс" Or Left(t, 2) = "ло" Or Left(t, 2) = "лн" Or Left(t, 2) = "лж") And MtIsDigit(MtCh(t, 3)) Then
            grade = UCase(t)
            mat = "Латунь"
            MtGradeTok = True
        ElseIf (Left(t, 3) = "лмц" Or Left(t, 4) = "лжмц" Or Left(t, 4) = "ламц") And n >= 5 Then
            If MtIsDigit(MtCh(t, n)) Then
                grade = UCase(t)
                mat = "Латунь"
                MtGradeTok = True
            End If
        End If
    Case "б"
        If Left(t, 2) = "бр" And n >= 4 Then
            If Left(t, 4) <> "брус" And Left(t, 5) <> "бронз" And Left(t, 4) <> "брыз" And Left(t, 4) <> "брит" And Left(t, 4) <> "брюк" And Left(t, 4) <> "брел" And Left(t, 4) <> "брак" And Left(t, 4) <> "брош" Then
                ok = False
                For i = 3 To n
                    If MtIsDigit(Mid(t, i, 1)) Then ok = True
                Next i
                If ok Then
                    grade = "Бр" & UCase(Mid(t, 3))
                    mat = "Бронза"
                    MtGradeTok = True
                End If
            End If
        End If
    End Select
End Function

' the next word of the name after position p: its text; sFrom, sTo — its place (0 — none)
Private Function MtNextTok(s As String, p As Long, ByRef sFrom As Long, ByRef sTo As Long) As String
    Dim q As Long
    q = p
    sFrom = 0
    sTo = 0
    Do While q <= Len(s)
        If MtIsSep(Mid(s, q, 1)) Then q = q + 1 Else Exit Do
    Loop
    If q > Len(s) Then Exit Function
    sFrom = q
    Do While q <= Len(s)
        If MtIsSep(Mid(s, q, 1)) Then Exit Do
        q = q + 1
    Loop
    sTo = q - 1
    MtNextTok = Mid(s, sFrom, sTo - sFrom + 1)
End Function

' the material and the grade of the name: the first word that is a mark; else the words («нержавеющая», «алюминиевый», «латунь» …). The spans of the marks are taken
Private Sub MtMaterial(s As String, ByRef mask As String, ByRef mat As String, ByRef grade As String)
    Dim p As Long, st As Long, en As Long, t As String, tt As String, f2 As Long, e2 As Long, t2 As String, m As String, g As String, d As String, k As Long, t3 As String, f3 As Long, e3 As Long
    mat = ""
    grade = ""
    mcStainless = (InStr(1, s, "нерж", 0) > 0 Or InStr(1, s, gMtAisi, 0) > 0 Or InStr(1, s, "inox", 0) > 0 Or MtWord(s, "н/ж", 1) > 0)
    mcAlu = (InStr(1, s, "алюмин", 0) > 0 Or InStr(1, s, "дюрал", 0) > 0)
    p = 1
    Do While p <= Len(s)
        t = MtNextTok(s, p, st, en)
        If st = 0 Then Exit Do
        p = en + 1
        tt = MtTrimTok(t)
        If tt = "сталь" Or tt = "стали" Or tt = "ст" Then
            t2 = MtNextTok(s, p, f2, e2)
            If f2 > 0 Then
                t2 = MtTrimTok(t2)
                k = 1
                d = MtDigitsAt(t2, k)
                If d <> "" And k > Len(t2) And Len(d) <= 3 Then
                    ' («сталь 3 мм» is a thickness, «сталь 45» is a grade)
                    t3 = MtNextTok(s, e2 + 1, f3, e3)
                    If t3 <> "мм" And t3 <> "м" And t3 <> "см" And Left(t3, 2) <> "мм" Then
                        If grade = "" Then
                            grade = "Ст" & d
                            mat = "Сталь"
                        End If
                        MtTake(mask, st, e2)
                        p = e2 + 1
                    End If
                ElseIf MtGradeTok(t2, m, g) Then
                    If grade = "" Then
                        grade = g
                        mat = m
                    End If
                    MtTake(mask, st, e2)
                    p = e2 + 1
                End If
            End If
            If mat = "" And tt <> "ст" Then mat = "Сталь"
        ElseIf Left(tt, Len(gMtAisi)) = gMtAisi And Len(tt) > Len(gMtAisi) And tt <> gMtAisi Then
            If MtIsDigit(MtCh(tt, Len(gMtAisi) + 1)) Then
                If grade = "" Then grade = "AISI " & UCase(Mid(tt, Len(gMtAisi) + 1))
                mat = "Нержавеющая сталь"
                MtTake(mask, st, en)
            End If
        ElseIf tt = gMtAisi Then
            t2 = MtNextTok(s, p, f2, e2)
            If f2 > 0 Then
                t2 = MtTrimTok(t2)
                If MtIsDigit(Left(t2, 1)) Then
                    If grade = "" Then grade = "AISI " & UCase(t2)
                    mat = "Нержавеющая сталь"
                    MtTake(mask, st, e2)
                    p = e2 + 1
                End If
            End If
        ElseIf grade = "" Then
            If MtGradeTok(tt, m, g) Then
                grade = g
                mat = m
                MtTake(mask, st, en)
            End If
        End If
    Loop
    If mcStainless And (mat = "" Or mat = "Сталь") Then mat = "Нержавеющая сталь"
    If mat = "" Then
        If mcAlu Then
            mat = "Алюминий"
        ElseIf InStr(1, s, "медн", 0) > 0 Or MtWord(s, "медь", 1) > 0 Or MtWord(s, "меди", 1) > 0 Then
            mat = "Медь"
        ElseIf InStr(1, s, "латун", 0) > 0 Then
            mat = "Латунь"
        ElseIf InStr(1, s, "бронз", 0) > 0 Then
            mat = "Бронза"
        ElseIf InStr(1, s, "титан", 0) > 0 Then
            mat = "Титан"
        ElseIf InStr(1, s, "чугун", 0) > 0 Then
            mat = "Чугун"
        ElseIf InStr(1, s, "стал", 0) > 0 Or InStr(1, s, "оцинк", 0) > 0 Then
            mat = "Сталь"
        End If
    End If
End Sub

' ================================================================ the numbers of a name

' the marker before the digits at st: ø ∅ ⌀ ф d (a diameter), s t (a thickness), l (a length), also with «=»; "" when none. ok = False when the digits are a part of a word or of a code («ст3», «лс59-1»)
Private Function MtMarkBefore(s As String, st As Long, ByRef ok As Boolean) As String
    Dim pc As String, q As Long
    ok = True
    pc = MtCh(s, st - 1)
    If pc = "" Then Exit Function
    q = st - 1
    If pc = "=" Then
        q = st - 2
        pc = MtCh(s, q)
        If pc = "" Then Exit Function
    End If
    ' a marker and the number with a space between them: «L 3000», «d 20»
    If pc = " " Then
        q = q - 1
        pc = MtCh(s, q)
        If pc = "ø" Or pc = "∅" Or pc = "⌀" Then
            q = q
        ElseIf (pc = "ф" Or pc = "d" Or pc = "д" Or pc = "l" Or pc = "s") And Not MtIsAlnum(MtCh(s, q - 1)) Then
            q = q
        Else
            ' a space and no marker of its own before it («проф 40»): the digits are a number by themselves
            Exit Function
        End If
    End If
    If pc = "ø" Or pc = "∅" Or pc = "⌀" Then
        MtMarkBefore = "D"
    ElseIf pc = "ф" Or pc = "d" Or pc = "д" Or pc = "s" Or pc = "t" Or pc = "т" Or pc = "l" Then
        If MtIsAlnum(MtCh(s, q - 1)) Then
            ok = False
        ElseIf pc = "s" Or pc = "t" Or pc = "т" Then
            MtMarkBefore = "T"
        ElseIf pc = "l" Then
            MtMarkBefore = "L"
        Else
            MtMarkBefore = "D"
        End If
    ElseIf pc = "х" And MtIsDigit(MtCh(s, q - 1)) Then
        ok = True
    ElseIf MtIsLetter(pc) Then
        ok = False
    ElseIf pc = "-" Or pc = "/" Then
        If MtIsAlnum(MtCh(s, st - 2)) Then ok = False
    End If
End Function

' what follows the digits (e is the place right after them): a unit («мм», «м», «см», «кг») and its last place, or the end of a code; ok = False when letters stick to the digits («20б1», «40х»)
Private Function MtTail(s As String, e As Long, ByRef unit As Integer, ByRef endp As Long) As Boolean
    Dim q As Long, c As String, c2 As String
    unit = 0
    endp = e - 1
    MtTail = True
    c = MtCh(s, e)
    q = e
    If c = " " Then
        q = e + 1
        c = MtCh(s, q)
    End If
    If c = "м" Then
        c2 = MtCh(s, q + 1)
        If c2 = "м" Then
            If Not MtIsLetter(MtCh(s, q + 2)) Then
                unit = 1
                endp = q + 1
            End If
        ElseIf Not MtIsLetter(c2) Then
            unit = 2
            endp = q
        End If
    ElseIf c = "с" Then
        If MtCh(s, q + 1) = "м" And Not MtIsLetter(MtCh(s, q + 2)) Then
            unit = 4
            endp = q + 1
        End If
    ElseIf c = "к" Then
        If MtCh(s, q + 1) = "г" And Not MtIsLetter(MtCh(s, q + 2)) Then
            unit = 3
            endp = q + 1
        End If
    End If
    If unit = 0 Then
        c = MtCh(s, e)
        If MtIsLetter(c) Then
            If c = "х" Then
                q = e + 1
                If MtCh(s, q) = " " Then q = q + 1
                If Not MtIsDigit(MtCh(s, q)) Then
                    If MtCh(s, q) <> "ø" And MtCh(s, q) <> "∅" And MtCh(s, q) <> "⌀" Then MtTail = False
                End If
            Else
                MtTail = False
            End If
        End If
    End If
End Function

' the numbers of the folded name (not those inside a taken span): value, place, unit, marker
Private Sub MtScanNumbers(s As String, mask As String)
    Dim p As Long, st As Long, e As Long, t As String, mark As String, ok As Boolean, unit As Integer, endp As Long, v As Double
    mnN = 0
    p = 1
    Do While p <= Len(s) And mnN < 30
        If Not MtIsDigit(Mid(s, p, 1)) Then
            p = p + 1
        Else
            st = p
            e = p
            t = MtNumAt(s, e)
            ok = True
            mark = MtMarkBefore(s, st, ok)
            If Mid(mask, st, 1) = "x" Then ok = False
            If ok Then ok = MtTail(s, e, unit, endp)
            If ok And t <> "" Then
                v = Val(t)
                If unit = 4 Then
                    v = v * 10
                    unit = 1
                End If
                mnVal(mnN) = v
                mnPos(mnN) = st
                mnEnd(mnN) = endp
                mnUnit(mnN) = unit
                mnMark(mnN) = mark
                mnUsed(mnN) = False
                mnN = mnN + 1
            End If
            If endp >= e Then p = endp + 1 Else p = e
            If p <= st Then p = st + 1
        End If
    Loop
End Sub

' the numbers joined by «х» are a chain: 3х1250х2500, 57х3,5, 40 х 20 х 2
Private Sub MtLinkChains(s As String)
    Dim i As Long, sep As String, n As Long
    mnChains = 0
    For i = 0 To mnN - 1
        mnChain(i) = 0
        mnLink(i) = 0
    Next i
    For i = 0 To mnN - 2
        n = mnPos(i + 1) - mnEnd(i) - 1
        sep = ""
        If n > 0 And n <= 4 Then sep = Trim(Mid(s, mnEnd(i) + 1, n))
        If sep = "х" Or sep = "хø" Or sep = "х∅" Or sep = "х⌀" Then
            If mnChain(i) = 0 Then
                mnChains = mnChains + 1
                mnChain(i) = mnChains
                mnLink(i) = 1
            End If
            mnChain(i + 1) = mnChain(i)
            mnLink(i + 1) = mnLink(i) + 1
        End If
    Next i
End Sub

' how many numbers a chain has
Private Function MtChainLen(c As Long) As Long
    Dim i As Long, n As Long
    If c <= 0 Then Exit Function
    For i = 0 To mnN - 1
        If mnChain(i) = c Then n = n + 1
    Next i
    MtChainLen = n
End Function

' the longest chain (the first of the equal ones); 0 — none
Private Function MtBestChain() As Long
    Dim c As Long, best As Long, bl As Long, n As Long
    For c = 1 To mnChains
        n = MtChainLen(c)
        If n > bl Then
            bl = n
            best = c
        End If
    Next c
    MtBestChain = best
End Function

' the numbers of chain c in mm in the order written (meters turned into millimeters when the chain ends with «м»); at most 10; their count
Private Function MtChainVals(c As Long, v() As Double) As Long
    Dim i As Long, n As Long, k As Double, allSmall As Boolean, lastM As Boolean
    k = 1
    allSmall = True
    For i = 0 To mnN - 1
        If mnChain(i) = c Then
            If mnVal(i) > 20 Then allSmall = False
            lastM = (mnUnit(i) = 2)
        End If
    Next i
    If lastM And allSmall Then k = 1000
    For i = 0 To mnN - 1
        If mnChain(i) = c And n < 10 Then
            v(n) = mnVal(i) * k
            mnUsed(i) = True
            n = n + 1
        End If
    Next i
    MtChainVals = n
End Function

' a number written in mm (a length in meters turned into mm)
Private Function MtEntryMm(i As Long) As Double
    If mnUnit(i) = 2 Then MtEntryMm = mnVal(i) * 1000 Else MtEntryMm = mnVal(i)
End Function

' the diameter: a number with the sign Ø; else the first number of its own (not of a chain, no unit of meters), up to 700; else the first of the chain
Private Function MtDiam(c As Long) As Double
    Dim i As Long
    For i = 0 To mnN - 1
        If mnMark(i) = "D" Then
            mnUsed(i) = True
            MtDiam = mnVal(i)
            Exit Function
        End If
    Next i
    For i = 0 To mnN - 1
        If mnChain(i) = 0 And mnMark(i) = "" And mnUnit(i) <> 2 And mnUnit(i) <> 3 And mnVal(i) <= 700 And mnVal(i) > 0 And Not mnUsed(i) Then
            mnUsed(i) = True
            MtDiam = mnVal(i)
            Exit Function
        End If
    Next i
    If c > 0 Then
        For i = 0 To mnN - 1
            If mnChain(i) = c And mnLink(i) = 1 Then
                MtDiam = mnVal(i)
                Exit Function
            End If
        Next i
    End If
End Function

' the thickness outside of the chain: a number with «s», «t»; a number with «мм» up to 100; a single small number
Private Function MtThickness(c As Long) As Double
    Dim i As Long
    For i = 0 To mnN - 1
        If mnChain(i) <> c Or c = 0 Then
            If mnMark(i) = "T" Then
                mnUsed(i) = True
                MtThickness = mnVal(i)
                Exit Function
            End If
        End If
    Next i
    For i = 0 To mnN - 1
        If (mnChain(i) <> c Or c = 0) And mnUnit(i) = 1 And mnVal(i) <= 100 And mnVal(i) > 0 And mnMark(i) = "" And Not mnUsed(i) Then
            mnUsed(i) = True
            MtThickness = mnVal(i)
            Exit Function
        End If
    Next i
    For i = 0 To mnN - 1
        If mnChain(i) = 0 And mnUnit(i) = 0 And mnMark(i) = "" And mnVal(i) <= 60 And mnVal(i) > 0 And Not mnUsed(i) Then
            mnUsed(i) = True
            MtThickness = mnVal(i)
            Exit Function
        End If
    Next i
End Function

' the length of a piece: a number with «L»; a number with «м» (0,2…20 m); a number of its own with «мм» from 300 up
Private Function MtLength() As Double
    Dim i As Long
    For i = 0 To mnN - 1
        If mnMark(i) = "L" And Not mnUsed(i) Then
            If mnVal(i) <= 20 And mnUnit(i) <> 1 Then MtLength = mnVal(i) * 1000 Else MtLength = mnVal(i)
            mnUsed(i) = True
            Exit Function
        End If
    Next i
    For i = 0 To mnN - 1
        If mnUnit(i) = 2 And Not mnUsed(i) And mnVal(i) >= 0.2 And mnVal(i) <= 20 Then
            mnUsed(i) = True
            MtLength = mnVal(i) * 1000
            Exit Function
        End If
    Next i
    For i = 0 To mnN - 1
        If mnUnit(i) = 1 And mnChain(i) = 0 And Not mnUsed(i) And mnVal(i) >= 300 And mnMark(i) = "" Then
            mnUsed(i) = True
            MtLength = mnVal(i)
            Exit Function
        End If
    Next i
End Function

' the series of a beam or of a channel: «10У», «12П», «20Б1», «25Ш1», «№14»; "" — none. The words of the name are taken one by one
Private Function MtBeam(s As String, mask As String) As String
    Dim p As Long, st As Long, en As Long, t As String, tt As String, k As Long, d As String, rest As String, f2 As Long, e2 As Long, t2 As String
    p = 1
    Do While p <= Len(s)
        t = MtNextTok(s, p, st, en)
        If st = 0 Then Exit Do
        p = en + 1
        If Mid(mask, st, 1) <> "x" Then
            tt = MtTrimTok(t)
            If Left(tt, 1) = "№" Then tt = Mid(tt, 2)
            k = 1
            d = MtDigitsAt(tt, k)
            If d <> "" And Len(d) <= 2 Then
                rest = Mid(tt, k)
                If rest = "" Or MtBeamSeries(rest) Then
                    t2 = MtNextTok(s, p, f2, e2)
                    If rest = "" And f2 > 0 Then
                        ' «10 У», «20 Б1»: the series is a word of its own
                        k = 1
                        tt = MtTrimTok(t2)
                        If MtBeamSeries(tt) And Not MtIsDigit(Left(tt, 1)) And Len(tt) <= 3 Then
                            If Left(tt, 1) = "в" Then tt = "б" & Mid(tt, 2)
                            MtBeam = d & UCase(tt)
                            MtTake(mask, st, e2)
                            Exit Function
                        End If
                    End If
                    If Left(t2, 2) <> "мм" And t2 <> "м" Then
                        If Left(rest, 1) = "в" Then rest = "б" & Mid(rest, 2)
                        MtBeam = d & UCase(rest)
                        MtTake(mask, st, en)
                        Exit Function
                    End If
                End If
            End If
        End If
    Loop
End Function

' ================================================================ the sizes of a kind

' the numbers of the chain in order into text: «3×1250×2500»
Private Function MtJoin(a As Double, b As Double, c As Double) As String
    Dim t As String
    If a > 0 Then t = MtShow(a)
    If b > 0 Then t = t & IIf(t <> "", "×", "") & MtShow(b)
    If c > 0 Then t = t & IIf(t <> "", "×", "") & MtShow(c)
    MtJoin = t
End Function

' the thickness, width, diameter (or the side) and length that the name tells for a kind; sz — the size as it is written for people. The kind of a pipe becomes «Труба профильная» when it is a profile
Private Sub MtSizes(s As String, mask As String, ByRef kind As String, ByRef t As Double, ByRef w As Double, ByRef d As Double, ByRef l As Double, ByRef sz As String)
    Dim c As Long, n As Long, v(9) As Double, lo As Double, hi As Double, md As Double, profile As Boolean, bm As String
    t = 0
    w = 0
    d = 0
    l = 0
    sz = ""
    c = MtBestChain()
    n = 0
    If c > 0 Then n = MtChainVals(c, v)
    Select Case kind
    Case "Лист", "Плита"
        If n >= 3 Then
            MtSort3(v(0), v(1), v(2), lo, md, hi)
            t = lo
            w = md
            l = hi
        ElseIf n = 2 Then
            lo = v(0)
            hi = v(1)
            If lo > hi Then
                lo = v(1)
                hi = v(0)
            End If
            If lo >= 100 Then
                w = lo
                l = hi
            Else
                t = lo
                w = hi
            End If
        End If
        If t = 0 Then t = MtThickness(c)
        sz = MtJoin(t, w, l)
    Case "Полоса", "Лента", "Шина"
        If n >= 3 Then
            MtSort3(v(0), v(1), v(2), lo, md, hi)
            t = lo
            w = md
            l = hi
        ElseIf n = 2 Then
            t = v(0)
            w = v(1)
            If t > w Then
                t = v(1)
                w = v(0)
            End If
        Else
            w = MtDiam(0)
            t = MtThickness(c)
        End If
        sz = MtJoin(t, w, l)
    Case "Круг", "Пруток", "Катанка", "Проволока", "Арматура"
        d = MtDiam(c)
        If c > 0 And n >= 2 And d = v(0) Then
            If v(n - 1) >= 300 Then l = v(n - 1)
        End If
        If d > 0 Then sz = "Ø" & MtShow(d)
    Case "Квадрат"
        If n >= 1 Then
            d = v(0)
        Else
            d = MtDiam(0)
        End If
        If d > 0 Then
            w = d
            sz = MtShow(d) & "×" & MtShow(d)
        End If
    Case "Шестигранник"
        d = MtDiam(c)
        If d = 0 Then d = MtThickness(0)
        If d > 0 Then sz = "S" & MtShow(d)
    Case "Труба", "Профиль"
        profile = (kind = "Профиль")
        If InStr(1, s, "проф", 0) > 0 Or InStr(1, s, "квадратн", 0) > 0 Or InStr(1, s, "прямоугол", 0) > 0 Then profile = True
        If n >= 3 Then
            profile = True
        ElseIf n = 2 Then
            If v(1) > v(0) * 0.35 Then profile = True
        End If
        If profile Then
            If kind = "Труба" Then kind = "Труба профильная"
            If n >= 4 Then
                ' a C- or Z-profile: height, width, the lip, the wall (the last)
                d = v(0)
                w = v(1)
                t = v(n - 1)
            ElseIf n = 3 Then
                d = v(0)
                w = v(1)
                t = v(2)
            ElseIf n = 2 Then
                d = v(0)
                w = v(1)
            Else
                d = MtDiam(0)
            End If
            If t = 0 Then t = MtThickness(c)
            sz = MtJoin(d, w, t)
            If n >= 4 Then sz = MtShow(v(0)) & "×" & MtShow(v(1)) & "×" & MtShow(v(2)) & "×" & MtShow(v(n - 1))
        Else
            If n = 2 Then
                d = v(0)
                t = v(1)
            Else
                d = MtDiam(c)
                t = MtThickness(c)
            End If
            If d > 0 Then sz = "Ø" & MtShow(d) & IIf(t > 0, "×" & MtShow(t), "")
        End If
    Case "Уголок"
        If n >= 3 Then
            d = v(0)
            w = v(1)
            t = v(2)
        ElseIf n = 2 Then
            If v(1) <= v(0) * 0.25 Then
                d = v(0)
                w = v(0)
                t = v(1)
            Else
                d = v(0)
                w = v(1)
            End If
        Else
            d = MtDiam(0)
            If d > 0 Then w = d
        End If
        If t = 0 Then t = MtThickness(c)
        sz = MtJoin(d, w, t)
    Case "Швеллер", "Двутавр", "Балка"
        bm = MtBeam(s, mask)
        sz = bm
    Case Else
        If n >= 3 Then
            d = v(0)
            w = v(1)
            t = v(2)
            sz = MtJoin(d, w, t)
        ElseIf n = 2 Then
            d = v(0)
            If MtFirstMark(c) = "D" Then
                l = v(1)
                sz = "Ø" & MtShow(d) & "×" & MtShow(l)
            Else
                w = v(1)
                sz = MtJoin(d, w, 0)
            End If
        Else
            d = MtDiam(0)
            If d > 0 Then sz = "Ø" & MtShow(d)
        End If
    End Select
    If l = 0 And kind <> "Лист" And kind <> "Плита" Then l = MtLength()
End Sub

' the marker of the first number of chain c
Private Function MtFirstMark(c As Long) As String
    Dim i As Long
    For i = 0 To mnN - 1
        If mnChain(i) = c And mnLink(i) = 1 Then
            MtFirstMark = mnMark(i)
            Exit Function
        End If
    Next i
End Function

' the three numbers in order: lo <= md <= hi
Private Sub MtSort3(a As Double, b As Double, c As Double, ByRef lo As Double, ByRef md As Double, ByRef hi As Double)
    lo = a
    md = b
    hi = c
    Dim x As Double
    If lo > md Then
        x = lo
        lo = md
        md = x
    End If
    If md > hi Then
        x = md
        md = hi
        hi = x
    End If
    If lo > md Then
        x = lo
        lo = md
        md = x
    End If
End Sub

' ================================================================ the weight

' kg per metre of the section and kg per piece (0 — not told) from the size and the density of the material (steel when it is not told)
Private Sub MtWeight(kind As String, mat As String, t As Double, w As Double, d As Double, l As Double, ByRef kgm As Double, ByRef kgpc As Double)
    Dim rho As Double, area As Double, pi As Double
    kgm = 0
    kgpc = 0
    rho = MtDensity(mat)
    pi = 3.14159265358979
    Select Case kind
    Case "Круг", "Пруток", "Катанка", "Проволока", "Арматура"
        If d > 0 Then area = pi / 4 * d * d
    Case "Квадрат"
        If d > 0 Then area = d * d
    Case "Шестигранник"
        If d > 0 Then area = 0.866025403784439 * d * d
    Case "Труба"
        If d > 0 And t > 0 And t < d / 2 Then area = pi * t * (d - t)
    Case "Труба профильная", "Профиль"
        If d > 0 And w > 0 And t > 0 And t < w / 2 And t < d / 2 Then area = t * (2 * (d + w) - 4 * t)
    Case "Уголок"
        If d > 0 And w > 0 And t > 0 And t < w And t < d Then area = t * (d + w - t)
    Case "Полоса", "Лента", "Шина"
        If w > 0 And t > 0 Then area = w * t
    Case "Лист", "Плита"
        If t > 0 And w > 0 And l > 0 Then kgpc = t * w * l * rho * 0.000001
    Case "Заготовка", "Поковка", "Отливка", "Слиток"
        If d > 0 And w > 0 And t > 0 Then
            kgpc = d * w * t * rho * 0.000001
        ElseIf d > 0 And l > 0 And w = 0 Then
            kgpc = pi / 4 * d * d * l * rho * 0.000001
        End If
    End Select
    If area > 0 Then
        kgm = area * rho * 0.001
        If l > 0 Then kgpc = kgm * l / 1000
    End If
End Sub

' ================================================================ the record of a name

Private Function MtEmpty() As Variant
    MtEmpty = Array("", 99, "", "", "", "", 0#, 0#, 0#, 0#, "", 0#, 0#, 0)
End Function

' the record of a name: (kind, order, finish, material, grade, size, thickness, width, diameter or side, length, standard, kg per metre, kg per piece, 1 — sold by the length). sFold — the folded name ("" — made here).
' A name that is not a rolled metal gives an empty kind
Function MtRecordF(sName As String, sFold As String) As Variant
    Dim s As String, kind As String, idx As Long, pos As Long, need As String, mask As String, std As String, mat As String, grade As String, fin As String
    Dim t As Double, w As Double, d As Double, l As Double, sz As String, kgm As Double, kgpc As Double, ev As Boolean, lin As Long, nCh As Long, c As Long
    MtLoad()
    If sFold <> "" Then s = sFold Else s = WmsNames.Fold(WmsNames.NormName(sName))
    kind = MtFindKind(s, idx, pos, need)
    If kind = "" Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If MtBlocked(s) Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If kind = "Профиль" And InStr(1, s, "труб", 0) > 0 Then
        kind = "Труба"
        idx = 13
    ElseIf kind = "Балка" And InStr(1, s, "двутавр", 0) > 0 Then
        kind = "Двутавр"
        idx = 16
    End If
    mask = Space(Len(s))
    std = MtStandard(s, mask)
    mcBeam = (kind = "Швеллер" Or kind = "Двутавр" Or kind = "Балка")
    mcRebar = (kind = "Арматура")
    mcWeakOk = (InStr(1, s, "мед", 0) > 0 Or kind = "Шина" Or kind = "Лента" Or kind = "Пруток" Or kind = "Лист" Or kind = "Круг" Or kind = "Проволока" Or kind = "Катанка" Or kind = "Труба")
    MtMaterial(s, mask, mat, grade)
    ev = (grade <> "" Or mat <> "" Or MtHasMetalWord(s))
    If need = "e" And Not ev Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If need = "c" Then
        ' a rebar: the class (А500С) or the standard of a rebar
        If grade = "" And InStr(std, "5781") + InStr(std, "34028") + InStr(std, "10884") = 0 Then
            MtRecordF = MtEmpty()
            Exit Function
        End If
    End If
    MtScanNumbers(s, mask)
    MtLinkChains(s)
    c = MtBestChain()
    nCh = MtChainLen(c)
    If need = "p" And nCh < 2 Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If need = "s" And Not ev And nCh < 3 Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If need = "t" And Not ev And nCh < 2 Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If (kind = "Труба" Or kind = "Круг" Or kind = "Квадрат" Or kind = "Шестигранник") And mnN = 0 And Not ev Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    If kind = "Уголок" And nCh < 2 And Not ev Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    MtSizes(s, mask, kind, t, w, d, l, sz)
    If kind = "Балка" And sz = "" Then
        MtRecordF = MtEmpty()
        Exit Function
    End If
    fin = MtFinish(s)
    MtWeight(kind, mat, t, w, d, l, kgm, kgpc)
    lin = 1
    If kind = "Лист" Or kind = "Плита" Or kind = "Заготовка" Or kind = "Поковка" Or kind = "Отливка" Or kind = "Слиток" Or kind = "Прочее" Then lin = 0
    MtRecordF = Array(kind, idx, fin, mat, grade, sz, t, w, d, l, std, kgm, kgpc, lin)
End Function

' MtRecordF made once for a name (kept while the book is open); a rule that fails on a name is the failure of that name, not of the list: the answer is an empty record and the text of the error is kept
Function MtRecordC(sName As String, Optional vFold As Variant) As Variant
    Dim v As Variant, p As Variant, r As Variant, sFold As String, sErr As String, t As String
    If Not gMtCacheOn Then
        gMtCache = New Collection
        gMtCacheOn = True
        gMtBadNames = ""
    End If
    If MtCacheGet(sName, v) Then
        p = Split(CStr(v), Chr(1))
        MtRecordC = Array(p(0), CLng(p(1)), p(2), p(3), p(4), p(5), Val(p(6)), Val(p(7)), Val(p(8)), Val(p(9)), p(10), Val(p(11)), Val(p(12)), CLng(p(13)))
        Exit Function
    End If
    If Not IsMissing(vFold) Then sFold = CStr(vFold)
    r = MtTry(sName, sFold, sErr)
    If sErr <> "" Then gMtBadNames = gMtBadNames & Chr(1) & sName & Chr(2) & sErr & Chr(1)
    t = r(0) & Chr(1) & r(1) & Chr(1) & r(2) & Chr(1) & r(3) & Chr(1) & r(4) & Chr(1) & r(5) & Chr(1) & Trim(Str(r(6))) & Chr(1) & Trim(Str(r(7))) & Chr(1) & Trim(Str(r(8))) & Chr(1) & Trim(Str(r(9))) _
        & Chr(1) & r(10) & Chr(1) & Trim(Str(r(11))) & Chr(1) & Trim(Str(r(12))) & Chr(1) & r(13)
    On Error Resume Next
    gMtCache.Add(t, sName)
    On Error GoTo 0
    MtRecordC = r
End Function

Private Function MtTry(sName As String, sFold As String, ByRef sErr As String) As Variant
    On Error GoTo EH
    MtTry = MtRecordF(sName, sFold)
    Exit Function
EH:
    sErr = Error$ & " (код " & Err & ", строка " & Erl & ")"
    MtTry = MtEmpty()
End Function

Private Function MtCacheGet(sName As String, ByRef v As Variant) As Boolean
    On Error GoTo Miss
    v = gMtCache.Item(sName)
    MtCacheGet = True
    Exit Function
Miss:
    MtCacheGet = False
End Function

' the words that stand for the parts of a record in the search («лист», «sheet», «нержавейка», «ø20», «3х1250х2500», «ст3», «гост 19903»)
Function MtTokens(rec As Variant) As String
    Dim s As String, sz As String, k As Variant, std As String, p As Long
    s = LCase(CStr(rec(MR_KIND))) & " " & LCase(CStr(rec(MR_FIN))) & " " & LCase(CStr(rec(MR_MAT))) & " " & LCase(CStr(rec(MR_GRADE)))
    Select Case CStr(rec(MR_KIND))
    Case "Лист", "Плита": s = s & " sheet plate"
    Case "Круг", "Пруток", "Катанка", "Проволока": s = s & " bar rod round wire"
    Case "Квадрат": s = s & " square bar"
    Case "Шестигранник": s = s & " hex bar hexagon"
    Case "Труба", "Труба профильная": s = s & " pipe tube"
    Case "Уголок": s = s & " angle"
    Case "Швеллер": s = s & " channel"
    Case "Двутавр", "Балка": s = s & " beam ibeam"
    Case "Полоса", "Лента", "Шина": s = s & " strip flat bar"
    Case "Арматура": s = s & " rebar"
    End Select
    Select Case CStr(rec(MR_MAT))
    Case "Нержавеющая сталь": s = s & " нержавейка нержавеющая stainless inox"
    Case "Алюминий": s = s & " алюминиевый aluminium aluminum"
    Case "Медь": s = s & " медный copper"
    Case "Латунь": s = s & " латунный brass"
    Case "Бронза": s = s & " бронзовый bronze"
    Case "Титан": s = s & " титановый titanium"
    Case "Чугун": s = s & " чугунный"
    Case "Сталь": s = s & " стальной steel"
    End Select
    If rec(MR_D) > 0 Then s = s & " ø" & Replace(Trim(Str(rec(MR_D))), ".", ",") & " d" & Trim(Str(rec(MR_D))) & " " & Trim(Str(rec(MR_D)))
    If rec(MR_T) > 0 Then s = s & " s" & Trim(Str(rec(MR_T))) & " " & Trim(Str(rec(MR_T)))
    If rec(MR_L) > 0 Then s = s & " l" & Trim(Str(rec(MR_L)))
    sz = LCase(CStr(rec(MR_SIZE)))
    If sz <> "" Then
        s = s & " " & Replace(Replace(sz, "×", "х"), ",", ".")
        For Each k In Split(Replace(Replace(Replace(Replace(sz, "ø", ""), "s", ""), "×", Chr(1)), "–", Chr(1)), Chr(1))
            If CStr(k) <> "" Then s = s & " " & Replace(CStr(k), ",", ".")
        Next k
    End If
    std = LCase(CStr(rec(MR_STD)))
    If std <> "" Then
        p = InStr(std, " ")
        s = s & " " & std & " " & Replace(std, " ", "") & " " & Mid(std, p + 1)
        p = InStr(std, "-")
        If p > 0 Then
            std = Trim(Left(std, p - 1))
            s = s & " " & std & " " & Replace(std, " ", "")
        End If
    End If
    MtTokens = s
End Function

' ================================================================ the categories

' True when the name of a category says that its goods are rolled metal («Металл», «Металлопрокат», «Сталь»)
Private Function MtCatSays(sCat As String) As Boolean
    Dim c As String
    c = WmsNames.Fold(sCat)
    MtCatSays = (InStr(c, "металл") > 0 Or InStr(c, "прокат") > 0 Or InStr(c, "сталь") > 0 Or InStr(c, "нерж") > 0 Or InStr(c, "алюмин") > 0 Or InStr(c, "цветмет") > 0 Or InStr(c, "чермет") > 0)
End Function

' True when the category is the one of the fasteners, of the tools or of the parts: those are the business of «Метизы», «Инструмент ЧПУ», «Остаток деталей»
Private Function MtCatElsewhere(sCat As String) As Boolean
    Dim c As String
    c = WmsNames.Fold(sCat)
    MtCatElsewhere = (InStr(c, "метиз") > 0 Or InStr(c, "крепеж") > 0 Or InStr(c, "чпу") > 0 Or InStr(c, "режущ") > 0 Or InStr(c, "детал") > 0)
End Function

' the keys of the goods of «Номенклатура» whose category says metal («a:article» / «n:name» folded, each between two Chr(1)); "" — there are none
Private Function MtGoodsKeys() As String
    Dim sh As Object, last As Long, d As Variant, i As Long, art As String, out As String
    If Not gDoc.Sheets.hasByName(SH_GOODS) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_GOODS)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, GD_CAT, last).getDataArray()
    For i = 0 To UBound(d)
        If MtCatSays(Trim(CStr(d(i)(GD_CAT)))) Then
            art = Trim(CStr(d(i)(GD_ART)))
            If art <> "" Then
                out = out & Chr(1) & "a:" & WmsNames.Fold(art) & Chr(1)
            ElseIf Trim(CStr(d(i)(GD_NAME))) <> "" Then
                out = out & Chr(1) & "n:" & WmsNames.Fold(WmsNames.NormName(CStr(d(i)(GD_NAME)))) & Chr(1)
            End If
        End If
    Next i
    MtGoodsKeys = out
End Function

' ================================================================ the unit of an EI

' what a unit of the registry measures: "m" metres, "mm", "kg", "t", "pc" pieces (шт, лист, пруток …; no unit is a piece), "" — something else
Private Function MtUnitKind(sUnit As String) As String
    Dim u As String
    u = LCase(Trim(Replace(Replace(Replace(sUnit, ".", ""), " ", ""), Chr(160), "")))
    Select Case u
    Case "м", "метр", "метра", "метры", "метров", "пм", "погм", "мп": MtUnitKind = "m"
    Case "мм": MtUnitKind = "mm"
    Case "кг", "килограмм": MtUnitKind = "kg"
    Case "т", "тн", "тонна", "тонн", "тонны": MtUnitKind = "t"
    Case "", "шт", "штук", "штука", "лист", "листов", "прут", "пруток", "хлыст", "штанга", "заготовка", "пластина", "отрезок", "кусок", "обрезок", "лента", "труба", "уголок", "швеллер", "балка", "полоса", "круг": MtUnitKind = "pc"
    End Select
End Function

' the length (m) and the weight (kg) of an EI: q in its unit, the packing (pu, pn: «в 1 шт — 6 м»); -1 — not known. rec — the record of the name
Private Sub MtEiAmounts(rec As Variant, q As Double, sUnit As String, pu As String, pn As Double, ByRef lenM As Double, ByRef mass As Double)
    Dim uk As String, pk As String
    lenM = -1
    mass = -1
    uk = MtUnitKind(sUnit)
    pk = ""
    If pu <> "" Then pk = MtUnitKind(pu)
    If pk = "m" Then
        lenM = q * pn
    ElseIf pk = "mm" Then
        lenM = q * pn / 1000
    ElseIf uk = "m" Then
        lenM = q
    ElseIf uk = "mm" Then
        lenM = q / 1000
    ElseIf uk = "pc" And rec(MR_LINEAR) = 1 And rec(MR_L) > 0 Then
        lenM = q * rec(MR_L) / 1000
    End If
    If pk = "kg" Then
        mass = q * pn
    ElseIf uk = "kg" Then
        mass = q
    ElseIf uk = "t" Then
        mass = q * 1000
    ElseIf lenM >= 0 And rec(MR_KGM) > 0 Then
        mass = lenM * rec(MR_KGM)
    ElseIf uk = "pc" And pk = "" And rec(MR_KGPC) > 0 Then
        mass = q * rec(MR_KGPC)
    End If
End Sub

' ================================================================ the sheet «Металл»

Sub BtnMetalRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = MetalRefresh(MetalFind())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' Ctrl+Shift+W — «Металл», the cursor in «Найти»
Sub NavMetal(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(MT_SHEET, MT_FIND_ROW, 1)
End Sub

' the words typed in the cell «Найти» of the sheet
Function MetalFind() As String
    If gDoc.Sheets.hasByName(MT_SHEET) Then MetalFind = Trim(gDoc.Sheets.getByName(MT_SHEET).getCellByPosition(1, MT_FIND_ROW).getString())
End Function

' a cell of the search or of the filters was changed: the list is made again
Sub OnMetalChange(Optional oTarget As Variant)
    Dim ra As Object
    If gMtBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If Not ((ra.Row = MT_FIND_ROW And (ra.Column = 1 Or ra.Column = 6)) Or (ra.Row = MT_FILTER_ROW And ra.Column < MT_COLS)) Then Exit Sub
    WmsInit()
    MetalRefresh(MetalFind())
    Exit Sub
EH:
End Sub

' the sheet is opened: the list is made again when the journal moved since it was made (a new book — the first time)
Sub OnMetalFocus(Optional oEvent As Variant)
    Dim seq As Long
    If gMtBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    seq = CLng(SysNum(SK_LAST_SEQ))
    If seq = gMtSeq And gMtSeq > 0 Then Exit Sub
    MetalRefresh(MetalFind())
    Exit Sub
EH:
End Sub

' the list of the rolled metal from «Наличия»: one row per name with an article — the balance of all its EIs, the places, how many EIs have a balance, the length (m) and the weight (kg) of what is left and
' the EIs themselves with their lengths (the pieces and the offcuts); sorted by kind, size, name. sFind — the words of the search (as on «Поиск»); under it the filters of the columns; «Скрыть нулевые» leaves
' out the names with no balance. Only shows: nothing is posted, nothing is written but the cells of this sheet. OK:<rows> / ERR:…
Function MetalRefresh(sFind As String) As String
    Dim sh As Object, i As Long, j As Long, k As Long, n As Long, g As Variant, rec As Variant, rows() As Variant, hays() As String, nms() As String, cols As Variant, notes() As String
    Dim filt As Variant, fTxt(19) As String, hide As Boolean, scores As Variant, ok As Boolean, out() As Variant, nAllF As Long, q As Double, cat As String
    Dim key As String, descr As String, kindName As String, coll As Object, wordsOld As Variant, mgood As String, folds As Variant, arF As Variant
    Dim aPl() As String, aCt() As String, aAr() As String, aSr() As String, t0 As Long, tp As String, nmAll() As String, arAll() As String
    Dim sF As String, rawH() As String, rawN() As String, hP As Variant, nP As Variant, nCold As Long, bProg As Boolean, rawO() As String, oldF As Variant
    Dim sErr As String, nErr As Long, nErl As Long, sStep As String, sAt As String, nBad As Long, sBad As String, pb As Long, qb As Long, bCatQ As Boolean
    Dim mapIdx() As Long, gi As Long, m As Long, r As Variant, pd As Variant, pk As Variant, nReg As Long, cEi As Long, cNm As Long, cUn As Long, cQ As Long, cPl As Long
    Dim totLen() As Double, totMass() As Double, badLen() As Boolean, badMass() As Boolean, eiTxt() As String, eiCnt() As Long, lenM As Double, mass As Double, pu As String, pn As Double, pkEi As Long, pkUn As Long, pkN As Long
    Dim sUn As String, sEx As String, psh As Object, sLen As Variant, sMass As Variant, lenV As Double, massV As Double, sPl As String, main As Double, second As Double, eiAll As Long
    On Error GoTo EH
    gMtBusy = True
    sStep = "начало"
    t0 = GetSystemTicks()
    WmsInit()
    If Not gDoc.Sheets.hasByName(MT_SHEET) Then
        MetalRefresh = "ERR:в книге нет листа «" & MT_SHEET & "» — книга не обновлена до 1.3.0"
        gMtBusy = False
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(MT_SHEET)
    sStep = "справочник категорий"
    If Not gCtLoaded Then WmsCat.CtLoad()
    sStep = "номенклатура"
    WmsCat.GoodsLoad()
    sStep = "наличие"
    WmsCat.StockGroups()
    MtLoad()
    tp = "реестр " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    sStep = "свёртка названий"
    ReDim nmAll(IIf(gGrpN > 0, gGrpN - 1, 0))
    For i = 0 To gGrpN - 1
        nmAll(i) = CStr(gGrp(i)(0))
    Next i
    folds = WmsNames.FoldAll(nmAll)
    sStep = "категории номенклатуры"
    mgood = MtGoodsKeys()
    If mgood <> "" Then
        ReDim arAll(IIf(gGrpN > 0, gGrpN - 1, 0))
        For i = 0 To gGrpN - 1
            arAll(i) = Trim(CStr(gGrp(i)(1)))
        Next i
        arF = WmsNames.FoldAll(arAll)
    End If
    tp = tp & ", свёртка имён " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    nCold = gGrpN
    If gMtCacheOn Then nCold = gGrpN - gMtCache.Count
    If nCold > 300 Then
        WmsView.ProgressStart("Металл: читаю названия…", gGrpN)
        bProg = True
    End If
    n = 0
    sStep = "разбор названий"
    ReDim rows(IIf(gGrpN > 0, gGrpN - 1, 0))
    ReDim mapIdx(IIf(gGrpN > 0, gGrpN - 1, 0))
    For i = 0 To gGrpN - 1
        mapIdx(i) = -1
        If bProg Then
            If (i Mod 40) = 0 Then WmsView.ProgressSet(i)
        End If
        g = gGrp(i)
        sAt = CStr(g(0))
        sF = CStr(folds(i))
        rec = MtRecordC(CStr(g(0)), sF)
        cat = Trim(CStr(g(6)))
        ok = False
        If CStr(rec(MR_KIND)) <> "" Then
            ok = Not MtCatElsewhere(cat)
        ElseIf cat <> "" Then
            ok = MtCatSays(cat)
        ElseIf mgood <> "" Then
            key = IIf(CStr(arAll(i)) <> "", "a:" & CStr(arF(i)), "n:" & sF)
            ok = (InStr(mgood, Chr(1) & key & Chr(1)) > 0)
        End If
        If ok Then
            If CStr(rec(MR_KIND)) = "" Then rec = Array("Прочее", 30, "", "", "", "", 0#, 0#, 0#, 0#, "", 0#, 0#, 0)
            rows(n) = Array(g, rec, sF, i)
            mapIdx(i) = n
            n = n + 1
        End If
    Next i
    sAt = ""
    If bProg Then
        WmsView.ProgressEnd()
        bProg = False
    End If
    If gMtBadNames <> "" Then
        For i = 0 To gGrpN - 1
            pb = InStr(gMtBadNames, Chr(1) & CStr(gGrp(i)(0)) & Chr(2))
            If pb > 0 Then
                nBad = nBad + 1
                If sBad = "" Then
                    qb = InStr(pb + 1, gMtBadNames, Chr(2)) + 1
                    sBad = "«" & Left(CStr(gGrp(i)(0)), 60) & "» — " & Mid(gMtBadNames, qb, InStr(qb, gMtBadNames, Chr(1)) - qb)
                End If
            End If
        Next i
    End If
    tp = tp & ", отбор " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the EIs of the metal: the length and the weight of what is left, the EIs as a text (the first eight)
    sStep = "ЕИ металла"
    ReDim totLen(IIf(n > 0, n - 1, 0))
    ReDim totMass(IIf(n > 0, n - 1, 0))
    ReDim badLen(IIf(n > 0, n - 1, 0))
    ReDim badMass(IIf(n > 0, n - 1, 0))
    ReDim eiTxt(IIf(n > 0, n - 1, 0))
    ReDim eiCnt(IIf(n > 0, n - 1, 0))
    If n > 0 And gGrpN > 0 Then
        nReg = UBound(gGrpReg) + 1
        psh = WmsPack.PackSheet()
        If Not IsNull(psh) Then pd = psh.getCellRangeByPosition(0, 1, PK_N, nReg).getDataArray()
        cEi = SC_EI
        cNm = SC_NAME
        cUn = SC_UNIT
        cQ = SC_QTY
        cPl = SC_PLACE
        pkEi = PK_EI
        pkUn = PK_UNIT
        pkN = PK_N
        For i = 0 To nReg - 1
            gi = gGrpOfRow(i)
            If gi >= 0 Then
                m = mapIdx(gi)
                If m >= 0 Then
                    r = gGrpReg(i)
                    q = 0
                    If VarType(r(cQ)) = 5 Then q = r(cQ)
                    If q > 0.0000005 Then
                        sUn = Trim(CStr(r(cUn)))
                        pu = ""
                        pn = 0
                        If IsArray(pd) Then
                            pk = pd(i)
                            If CStr(pk(pkEi)) = CStr(r(cEi)) And VarType(pk(pkN)) = 5 Then
                                If pk(pkN) > 0 And Trim(CStr(pk(pkUn))) <> "" Then
                                    pu = Trim(CStr(pk(pkUn)))
                                    pn = pk(pkN)
                                End If
                            End If
                        End If
                        rec = rows(m)(1)
                        MtEiAmounts(rec, q, sUn, pu, pn, lenM, mass)
                        ' (a sheet, a plate and a blank have no length to count; a section whose size does not give a weight — a channel, a beam — has no weight: neither is a lack to tell about)
                        If lenM >= 0 Then
                            totLen(m) = totLen(m) + lenM
                        ElseIf rec(MR_LINEAR) = 1 Then
                            badLen(m) = True
                        End If
                        If mass >= 0 Then
                            totMass(m) = totMass(m) + mass
                        ElseIf rec(MR_KGM) > 0 Or rec(MR_KGPC) > 0 Then
                            badMass(m) = True
                        End If
                        eiCnt(m) = eiCnt(m) + 1
                        If eiCnt(m) <= 8 Then
                            sEx = ""
                            If pu <> "" Then sEx = WmsIssue.QtyText(WmsPack.PartsOf(q, pn)) & " " & pu
                            If lenM >= 0 And MtUnitKind(pu) <> "m" And MtUnitKind(sUn) <> "m" Then sEx = sEx & IIf(sEx <> "", "; ", "") & WmsIssue.QtyText(WmsPack.PkRound(lenM, 2)) & " м"
                            If mass >= 0 Then sEx = sEx & IIf(sEx <> "", "; ", "") & "≈" & WmsIssue.QtyText(WmsPack.PkRound(mass, 1)) & " кг"
                            sPl = Trim(CStr(r(cPl)))
                            eiTxt(m) = eiTxt(m) & IIf(eiTxt(m) <> "", "; ", "") & CStr(r(cEi)) & " — " & WmsIssue.QtyText(q) & " " & sUn & IIf(sEx <> "", " (" & sEx & ")", "") & IIf(sPl <> "", " [" & sPl & "]", "")
                        End If
                    End If
                End If
            End If
        Next i
    End If
    tp = tp & ", ЕИ " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    ' the search and the filters of the sheet
    sStep = "поиск и фильтры листа"
    filt = sh.getCellRangeByPosition(0, MT_FILTER_ROW, MT_COLS - 1, MT_FILTER_ROW).getDataArray()(0)
    For j = 0 To MT_COLS - 1
        fTxt(j) = Trim(CStr(filt(j)))
    Next j
    hide = (WmsNames.Fold(Trim(sh.getCellByPosition(6, MT_FIND_ROW).getString())) = "да")
    ReDim hays(IIf(n > 0, n - 1, 0))
    ReDim nms(IIf(n > 0, n - 1, 0))
    ReDim notes(IIf(n > 0, n - 1, 0))
    ReDim rawH(IIf(n > 0, n - 1, 0))
    ReDim rawN(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        g = rows(i)(0)
        rec = rows(i)(1)
        rawH(i) = MtTokens(rec) & " " & g(0) & " " & g(1) & " " & g(5)
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
    ' the possible doubles: the same kind, size, grade, finish and standard under different names; or the same name under different articles
    sStep = "возможные дубли"
    coll = New Collection
    For i = 0 To n - 1
        rec = rows(i)(1)
        g = rows(i)(0)
        If CStr(rec(MR_KIND)) <> "Прочее" And CStr(rec(MR_SIZE)) <> "" Then
            key = "p|" & rec(MR_KIND) & "|" & rec(MR_SIZE) & "|" & rec(MR_GRADE) & "|" & rec(MR_FIN) & "|" & WmsSpec.SpCanon(CStr(rec(MR_STD)))
            WmsCat.DupMark(coll, key, i, notes, rows)
        End If
        key = "n|" & rows(i)(2)
        WmsCat.DupMark(coll, key, i, notes, rows)
    Next i
    For i = 0 To n - 1
        If (badLen(i) Or badMass(i)) And eiCnt(i) > 0 Then
            notes(i) = notes(i) & IIf(notes(i) <> "", "; ", "") & "длина или масса не по всем ЕИ — нужна длина в названии, фасовка «в 1 шт — N м» или единица «м»"
        End If
    Next i
    tp = tp & ", дубли " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    sStep = "поиск по словам"
    cols = Empty
    If Mid(WmsSearch.QueryNeeds(sFind), 1, 4) <> "0000" Then
        bCatQ = (Mid(WmsSearch.QueryNeeds(sFind), 2, 1) = "1")
        If bCatQ Then WmsCat.GroupCats()
        ReDim aPl(IIf(n > 0, n - 1, 0))
        ReDim aCt(IIf(n > 0, n - 1, 0))
        ReDim aAr(IIf(n > 0, n - 1, 0))
        ReDim aSr(IIf(n > 0, n - 1, 0))
        For i = 0 To n - 1
            g = rows(i)(0)
            aPl(i) = CStr(g(5))
            aAr(i) = CStr(g(1))
            If bCatQ Then aCt(i) = CStr(gGrpCat(CLng(rows(i)(3))))
        Next i
        cols = Array(WmsNames.FoldAll(aPl), WmsNames.FoldAll(aCt), WmsNames.FoldAll(aAr), WmsNames.FoldAll(aSr))
    End If
    scores = WmsSearch.MatchScores(sFind, hays, nms, n, cols)
    wordsOld = WmsNames.FoldWords(sFind)
    If Trim(sFind) <> "" And n > 0 Then
        ReDim rawO(n - 1)
        For i = 0 To n - 1
            g = rows(i)(0)
            rec = rows(i)(1)
            rawO(i) = CStr(g(0)) & " " & g(1) & " " & g(5) & " " & rec(MR_KIND) & " " & rec(MR_SIZE) & " " & rec(MR_MAT) & " " & rec(MR_GRADE)
        Next i
        oldF = WmsNames.FoldAll(rawO)
    End If
    tp = tp & ", поиск " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    sStep = "фильтры столбцов"
    ReDim out(IIf(n > 0, n - 1, 0))
    k = 0
    For i = 0 To n - 1
        g = rows(i)(0)
        rec = rows(i)(1)
        sAt = CStr(g(0))
        q = g(3)
        ok = (scores(i) > 0)
        If Not ok And Trim(sFind) <> "" Then ok = WmsNames.HasWords(CStr(oldF(i)), wordsOld)
        If ok And hide And q < 0.0000005 Then ok = False
        If ok Then
            kindName = CStr(rec(MR_KIND))
            sLen = ""
            sMass = ""
            lenV = 0
            massV = 0
            If eiCnt(i) > 0 Then
                If Not badLen(i) And totLen(i) > 0 Then
                    lenV = WmsPack.PkRound(totLen(i), 2)
                    sLen = lenV
                End If
                If Not badMass(i) And totMass(i) > 0 Then
                    massV = WmsPack.PkRound(totMass(i), 1)
                    sMass = massV
                End If
            End If
            If fTxt(MC_KIND) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_KIND), WmsNames.Fold(kindName))
            If fTxt(MC_FIN) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_FIN), WmsNames.Fold(CStr(rec(MR_FIN))))
            If fTxt(MC_MAT) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_MAT), WmsNames.Fold(CStr(rec(MR_MAT))))
            If fTxt(MC_GRADE) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_GRADE), WmsNames.Fold(CStr(rec(MR_GRADE))))
            If fTxt(MC_SIZE) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_SIZE), WmsNames.Fold(CStr(rec(MR_SIZE))))
            If fTxt(MC_T) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_T), CDbl(rec(MR_T)))
            If fTxt(MC_W) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_W), CDbl(rec(MR_W)))
            If fTxt(MC_D) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_D), CDbl(rec(MR_D)))
            If fTxt(MC_L) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_L), CDbl(rec(MR_L)))
            If fTxt(MC_STD) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_STD), WmsNames.Fold(CStr(rec(MR_STD))))
            If fTxt(MC_NAME) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_NAME), WmsNames.Fold(CStr(g(0))))
            If fTxt(MC_ART) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_ART), WmsNames.Fold(CStr(g(1))))
            If fTxt(MC_QTY) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_QTY), q)
            If fTxt(MC_UNIT) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_UNIT), WmsNames.Fold(CStr(g(2))))
            If fTxt(MC_LEN) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_LEN), lenV)
            If fTxt(MC_MASS) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_MASS), massV)
            If fTxt(MC_PLACES) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_PLACES), WmsNames.Fold(CStr(g(5))))
            If fTxt(MC_NEI) <> "" Then ok = ok And WmsCat.NumFilterOk(fTxt(MC_NEI), CDbl(g(4)))
            If fTxt(MC_NOTE) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_NOTE), WmsNames.Fold(notes(i)))
            If fTxt(MC_EIS) <> "" Then ok = ok And WmsCat.TextFilterOk(fTxt(MC_EIS), WmsNames.Fold(eiTxt(i)))
        End If
        If ok Then
            If rec(MR_D) > 0 Then main = rec(MR_D) Else main = 0
            If main = 0 And rec(MR_T) > 0 And (kindName = "Лист" Or kindName = "Плита") Then main = rec(MR_T)
            If main = 0 Then main = rec(MR_W)
            second = rec(MR_T)
            If kindName = "Лист" Or kindName = "Плита" Then second = rec(MR_W)
            out(k) = Array(kindName, rec(MR_FIN), rec(MR_MAT), rec(MR_GRADE), rec(MR_SIZE), IIf(rec(MR_T) > 0, rec(MR_T), ""), IIf(rec(MR_W) > 0, rec(MR_W), ""), IIf(rec(MR_D) > 0, rec(MR_D), ""), _
                IIf(rec(MR_L) > 0, rec(MR_L), ""), rec(MR_STD), g(0), g(1), WmsPack.PkRound(q, 3), g(2), sLen, sMass, g(5), g(4), notes(i), eiTxt(i) & IIf(eiCnt(i) > 8, " … и ещё " & (eiCnt(i) - 8), ""), _
                CDbl(rec(MR_IDX)) * 1000000000000# + CDbl(Int(main * 10 + 0.5)) * 1000000# + CDbl(Int(second * 10 + 0.5)))
            k = k + 1
        End If
    Next i
    tp = tp & ", фильтры " & (GetSystemTicks() - t0)
    t0 = GetSystemTicks()
    sStep = "запись листа"
    sAt = ""
    sh.getCellRangeByPosition(0, MT_ROW0, MT_KEY_COL, MT_ROW0 + MT_MAX).clearContents(1 + 2 + 4 + 16)
    nAllF = n
    If k > MT_MAX Then k = MT_MAX
    If k > 0 Then
        ReDim Preserve out(k - 1)
        sh.getCellRangeByPosition(0, MT_ROW0, MT_KEY_COL, MT_ROW0 + k - 1).setDataArray(out)
        MtSort(sh.getCellRangeByPosition(0, MT_ROW0, MT_KEY_COL, MT_ROW0 + k - 1))
        sh.getCellRangeByPosition(0, MT_ROW0, MT_COLS - 1, MT_ROW0 + k - 1).CharColor = RGB(38, 50, 56)
    End If
    tp = tp & ", запись " & (GetSystemTicks() - t0)
    gMtProf = tp & " (групп " & gGrpN & ", металла " & n & ")"
    descr = ""
    If Trim(sFind) <> "" Then descr = " — по словам «" & Trim(sFind) & "»"
    For j = 0 To MT_COLS - 1
        If fTxt(j) <> "" Then
            descr = descr & IIf(descr <> "", "; ", " — ") & "фильтры столбцов"
            Exit For
        End If
    Next j
    sh.getCellByPosition(MT_STAT_COL, MT_FIND_ROW).setString("Металла (позиций): " & k & IIf(k < nAllF, " из " & nAllF, "") & descr & ". Список из «Наличия» на " & Format(Now(), "DD.MM.YYYY HH:MM") _
        & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)) & IIf(nBad > 0, ". НЕ ПРОЧИТАНО названий: " & nBad & " (первое: " & sBad & ") — их нет в списке", ""))
    gMtSeq = CLng(SysNum(SK_LAST_SEQ))
    If gMtSeq = 0 Then gMtSeq = 1
    gMtBusy = False
    MetalRefresh = "OK:" & k
    gUiLastMsg = MetalRefresh
    Exit Function
EH:
    ' the error is read first: the closing of the bar (WmsView.ProgressEnd) switches the error handling off and the text of the error is lost
    sErr = Error$
    nErr = Err
    nErl = Erl
    If bProg Then WmsView.ProgressEnd()
    gMtBusy = False
    MetalRefresh = "ERR:список металла не построен: " & sErr & " (код " & nErr & ", строка " & nErl & ", шаг «" & sStep & "»" & IIf(sAt <> "", ", название «" & Left(sAt, 80) & "»", "") & ")"
    gUiLastMsg = MetalRefresh
    On Error Resume Next
    gDoc.Sheets.getByName(MT_SHEET).getCellByPosition(MT_STAT_COL, MT_FIND_ROW).setString(Mid(MetalRefresh, 5) & ". Нажмите «Обновить» ещё раз; если повторится — покажите это сообщение.")
    On Error GoTo 0
End Function

' the list sorted by the key of the kind and the size (the hidden column), then by the name
Private Sub MtSort(rng As Object)
    Dim f(1) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = MT_KEY_COL
    f(0).IsAscending = True
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    f(1).Field = MC_NAME
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

' ================================================================ test seams

Function TestMetalProf() As String
    TestMetalProf = gMtProf
End Function

' the record of a name as "kind|finish|material|grade|size|thickness|width|diameter|length|standard" (numbers: a point, 0 — none)
Function TestMetalParse(sName As String) As String
    Dim r As Variant
    MtLoad()
    r = MtRecordF(sName, "")
    TestMetalParse = r(MR_KIND) & "|" & r(MR_FIN) & "|" & r(MR_MAT) & "|" & r(MR_GRADE) & "|" & r(MR_SIZE) & "|" & Trim(Str(r(MR_T))) & "|" & Trim(Str(r(MR_W))) & "|" & Trim(Str(r(MR_D))) & "|" & Trim(Str(r(MR_L))) & "|" & r(MR_STD)
End Function

' the weight of a name: "kg per metre|kg per piece" (rounded to 0,001)
Function TestMetalWeight(sName As String) As String
    Dim r As Variant
    MtLoad()
    r = MtRecordF(sName, "")
    TestMetalWeight = Trim(Str(Int(r(MR_KGM) * 1000 + 0.5) / 1000)) & "|" & Trim(Str(Int(r(MR_KGPC) * 1000 + 0.5) / 1000))
End Function

' the amounts of an EI: "length m|weight kg" for the name, the quantity, the unit and the packing (-1 — not known)
Function TestMetalAmounts(sName As String, q As Double, sUnit As String, pu As String, pn As Double) As String
    Dim r As Variant, lenM As Double, mass As Double
    MtLoad()
    r = MtRecordF(sName, "")
    MtEiAmounts(r, q, sUnit, pu, pn, lenM, mass)
    TestMetalAmounts = Trim(Str(Int(lenM * 1000 + 0.5) / 1000)) & "|" & Trim(Str(Int(mass * 1000 + 0.5) / 1000))
End Function

' every name of the list (Chr(10) between them) is read by the rules, twice: by itself and with the folded name given (as the list does it); the answer: how many names, how many failures and the first of them
Function TestMetalFuzz(sList As String) As String
    Dim a As Variant, i As Long, nBad As Long, out As String, sErr As String, sNm As String, t0 As Long, nSlow As Long
    MtLoad()
    a = Split(sList, Chr(10))
    For i = 0 To UBound(a)
        sNm = CStr(a(i))
        t0 = GetSystemTicks()
        sErr = ""
        MtTry(sNm, "", sErr)
        If sErr = "" Then MtTry(sNm, WmsNames.Fold(sNm), sErr)
        If sErr <> "" Then
            nBad = nBad + 1
            If nBad <= 30 Then out = out & Chr(10) & sNm & " ¦ " & sErr
        End If
        If GetSystemTicks() - t0 > 2000 Then
            nSlow = nSlow + 1
            If nSlow <= 5 Then out = out & Chr(10) & "МЕДЛЕННО " & (GetSystemTicks() - t0) & " мс: " & sNm
        End If
    Next i
    TestMetalFuzz = (UBound(a) + 1) & " имён, ошибок " & nBad & ", медленных " & nSlow & out
End Function

' the mark that a word is: "material|mark" ("" — it is not a mark); the context of the name given by sCtx (the folded name)
Function TestMetalGrade(sTok As String, sCtx As String) As String
    Dim m As String, g As String
    MtLoad()
    mcStainless = (InStr(1, sCtx, "нерж", 0) > 0)
    mcAlu = (InStr(1, sCtx, "алюмин", 0) > 0)
    mcBeam = (InStr(1, sCtx, "швеллер", 0) > 0 Or InStr(1, sCtx, "двутавр", 0) > 0)
    mcRebar = (InStr(1, sCtx, "арматур", 0) > 0)
    mcWeakOk = (InStr(1, sCtx, "мед", 0) > 0)
    If MtGradeTok(sTok, m, g) Then TestMetalGrade = m & "|" & g
End Function

' the key by which two names are taken for a possible double (kind, size, grade, finish, standard)
Function TestMetalDupKey(sName As String) As String
    Dim rec As Variant
    MtLoad()
    rec = MtRecordF(sName, "")
    TestMetalDupKey = "p|" & rec(MR_KIND) & "|" & rec(MR_SIZE) & "|" & rec(MR_GRADE) & "|" & rec(MR_FIN) & "|" & WmsSpec.SpCanon(CStr(rec(MR_STD)))
End Function
