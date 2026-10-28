' WmsSpec — 1.1: what a name of a fastener says. The parsing that «Метизы» and the search use: the type («Болт»), the size («М8×30», with the pitch «М8×1,25×30»,
' «4,2×32»), the standard («DIN 933», «ГОСТ 7798-70», «ISO 4017») and its equivalents (the sheet «Стандарты»: DIN 933 = ISO 4017 = …), the class of strength or
' of the stainless steel («8.8», «А2»), the coating or the material, the head and the drive («потайная головка», «Torx»). A name is read as a person reads it —
' no network, no guessing beyond the rules below. Library WmsPlus.
Option Explicit

Public Const SH_STD = "Стандарты"
Public Const ST_MAX = 400

' the equivalents of the standards: «din 933» → «iso 4017;гост 7798» (read from the sheet «Стандарты» by the first use; the event of the sheet drops it)
Global gSpEq As Object
Global gSpEqLoaded As Boolean

' the types in the order of the list: (the stem of the word, the name of the type)
Function SpTypes() As Variant
    SpTypes = Array("болт|Болт", "винт|Винт", "шуруп|Шуруп", "саморез|Саморез", "шпильк|Шпилька", "гайк|Гайка", "шайб|Шайба", "шплинт|Шплинт", _
        "дюбел|Дюбель", "анкер|Анкер", "заклепк|Заклёпка", "гвозд|Гвоздь", "хомут|Хомут", "стопорн|Стопорное кольцо", "штифт|Штифт")
End Function

Private Function IsDigit(c As String) As Boolean
    IsDigit = (Len(c) = 1 And c >= "0" And c <= "9")
End Function

Private Function IsLetter(c As String) As Boolean
    IsLetter = (Len(c) = 1) And ((c >= "a" And c <= "z") Or (c >= "а" And c <= "я") Or c = "ё")
End Function

' the number at position p of s (digits, a decimal comma or point followed by a digit): its text; p is moved past it
Private Function NumberAt(s As String, ByRef p As Long) As String
    Dim t As String, c As String
    Do While p <= Len(s)
        c = Mid(s, p, 1)
        If IsDigit(c) Then
            t = t & c
            p = p + 1
        ElseIf (c = "," Or c = ".") And t <> "" And p < Len(s) Then
            If IsDigit(Mid(s, p + 1, 1)) Then
                t = t & "."
                p = p + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    NumberAt = t
End Function

Private Function IsSep(c As String) As Boolean
    IsSep = (c = "x" Or c = "х" Or c = "×" Or c = "*")
End Function

' after the number that ended at position q: «х 30» (a separator and a number) → the number, q moved past it; "" — there is none
Private Function NextAfterSep(s As String, ByRef q As Long) As String
    Dim n As Long, p As Long
    n = Len(s)
    p = q
    Do While p <= n And Mid(s, p, 1) = " "
        p = p + 1
    Loop
    If p < n Then
        If IsSep(Mid(s, p, 1)) Then
            p = p + 1
            Do While p <= n And Mid(s, p, 1) = " "
                p = p + 1
            Loop
            If p <= n Then
                If IsDigit(Mid(s, p, 1)) Then
                    NextAfterSep = NumberAt(s, p)
                    q = p
                End If
            End If
        End If
    End If
End Function

' the size in a name: «М8х30», «M10 x 40», «4,2×32», «М8х1,25х30», «М8» → «М8×30», «4.2×32», «М8×1,25×30», «М8»; the diameter, the pitch (0 — none) and the length as numbers
Sub SpSize(sName As String, ByRef sSize As String, ByRef dDiam As Double, ByRef dLen As Double, ByRef dPitch As Double)
    Dim s As String, n As Long, i As Long, p As Long, c As String, prev As String, hasM As Boolean, n1 As String, n2 As String, n3 As String, q As Long, st As Long
    s = LCase(sName)
    n = Len(s)
    sSize = ""
    dDiam = 0
    dLen = 0
    dPitch = 0
    i = 1
    Do While i <= n
        c = Mid(s, i, 1)
        hasM = False
        st = i
        If (c = "м" Or c = "m") And i < n Then
            If IsDigit(Mid(s, i + 1, 1)) Then
                prev = IIf(i > 1, Mid(s, i - 1, 1), " ")
                If Not IsLetter(prev) And Not IsDigit(prev) Then
                    hasM = True
                    i = i + 1
                    c = Mid(s, i, 1)
                End If
            End If
        End If
        If IsDigit(c) Then
            prev = IIf(st > 1, Mid(s, st - 1, 1), " ")
            If hasM Or (Not IsLetter(prev) And Not IsDigit(prev) And prev <> "." And prev <> ",") Then
                p = i
                n1 = NumberAt(s, p)
                q = p
                n2 = NextAfterSep(s, q)
                n3 = ""
                If n2 <> "" Then
                    p = q
                    n3 = NextAfterSep(s, q)
                    If n3 <> "" Then p = q
                End If
                If hasM Or n2 <> "" Then
                    If n3 <> "" Then
                        sSize = IIf(hasM, "М", "") & Replace(n1, ".", ",") & "×" & Replace(n2, ".", ",") & "×" & Replace(n3, ".", ",")
                        dDiam = Val(n1)
                        dPitch = Val(n2)
                        dLen = Val(n3)
                    Else
                        sSize = IIf(hasM, "М", "") & Replace(n1, ".", ",") & IIf(n2 <> "", "×" & Replace(n2, ".", ","), "")
                        dDiam = Val(n1)
                        dLen = Val(n2)
                    End If
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
Private Function After(s As String, sKey As String) As String
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
            If IsDigit(c) Then
                t = t & c
                p = p + 1
            ElseIf (c = "-" Or c = "–") And t <> "" And p < Len(s) Then
                If IsDigit(Mid(s, p + 1, 1)) Then
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
            After = t
            Exit Function
        End If
        p = InStr(p, s, sKey)
    Loop
End Function

' DIN, ГОСТ or ISO of a name: «DIN 933», «ГОСТ 7798-70», «ISO 4017»; "" — none
Function SpStandard(sName As String) As String
    Dim s As String, t As String
    s = LCase(sName)
    t = After(s, "din")
    If t <> "" Then
        SpStandard = "DIN " & t
        Exit Function
    End If
    t = After(s, "гост")
    If t <> "" Then
        SpStandard = "ГОСТ " & t
        Exit Function
    End If
    t = After(s, "iso")
    If t <> "" Then
        SpStandard = "ISO " & t
        Exit Function
    End If
    t = After(s, "исо")
    If t <> "" Then SpStandard = "ISO " & t
End Function

' the class of strength («8.8», «10.9») or of stainless steel («А2», «А4», «А2-70») of a name; "" — none
Function SpClass(sName As String) As String
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
                    SpClass = t
                    Exit Function
                End If
            Next j
            If t = "а2" Or t = "а4" Or t = "a2" Or t = "a4" Or Left(t, 3) = "а2-" Or Left(t, 3) = "а4-" Or Left(t, 3) = "a2-" Or Left(t, 3) = "a4-" Then
                SpClass = "А" & Mid(t, 2, 1) & IIf(Len(t) > 2, Mid(t, 3), "")
                Exit Function
            End If
        End If
    Next i
End Function

' the coating or the material of a name: «Оцинкованный», «Нержавеющая сталь», «Чёрный» …; "" — none
Function SpCoating(sName As String) As String
    Dim s As String, k As Variant, i As Long, p As Variant
    s = Replace(LCase(sName), "ё", "е")
    k = Array("горячеоцинк|Горячее цинкование", "горячий цинк|Горячее цинкование", "оцинк|Оцинкованный", "дакромет|Дакромет", "цинк|Цинк", "нерж|Нержавеющая сталь", "латун|Латунь", _
        "черн|Чёрный", "хромир|Хромированный", "никел|Никелированный", "фосфат|Фосфатированный", "медн|Медь", "алюмин|Алюминий", "нейлон|Пластик", "пластик|Пластик", _
        "полиамид|Пластик", "полимер|Пластик")
    For i = 0 To UBound(k)
        p = Split(CStr(k(i)), "|")
        If InStr(s, p(0)) > 0 Then
            SpCoating = p(1)
            Exit Function
        End If
    Next i
End Function

' the type of a fastener by the earliest word of the name that names one (the index in the list of types; -1 — none) and its name
Function SpType(sName As String, ByRef idx As Long) As String
    Dim s As String, t As Variant, i As Long, p As Variant, k As Long, best As Long
    s = Replace(LCase(sName), "ё", "е")
    t = SpTypes()
    idx = -1
    best = 100000
    For i = 0 To UBound(t)
        p = Split(CStr(t(i)), "|")
        k = InStr(s, Replace(CStr(p(0)), "ё", "е"))
        If k > 0 And k < best Then
            best = k
            idx = i
            SpType = p(1)
        End If
    Next i
End Function

' the head and the drive of a name («потайная головка, Torx»); "" — none
Function SpHead(sName As String) As String
    Dim s As String, out As String
    s = Replace(LCase(sName), "ё", "е")
    If InStr(s, "потай") > 0 Then
        out = "потайная головка"
    ElseIf InStr(s, "полукруг") > 0 Then
        out = "полукруглая головка"
    ElseIf InStr(s, "цилиндр") > 0 Then
        out = "цилиндрическая головка"
    End If
    If InStr(s, "внутр") > 0 And InStr(s, "шестигран") > 0 Then
        out = out & IIf(out <> "", ", ", "") & "внутренний шестигранник"
    ElseIf InStr(s, "шестигран") > 0 Then
        out = out & IIf(out <> "", ", ", "") & "шестигранная"
    End If
    If InStr(s, "torx") > 0 Or InStr(s, "торкс") > 0 Then
        out = out & IIf(out <> "", ", ", "") & "Torx"
    ElseIf InStr(s, "pozidriv") > 0 Or InStr(s, "позидрив") > 0 Then
        out = out & IIf(out <> "", ", ", "") & "Pozidriv"
    ElseIf InStr(s, "крест") > 0 Or InStr(s, "phillips") > 0 Then
        out = out & IIf(out <> "", ", ", "") & "крестовой шлиц"
    ElseIf InStr(s, "шлиц") > 0 Then
        out = out & IIf(out <> "", ", ", "") & "прямой шлиц"
    End If
    SpHead = out
End Function

' ================================================================ the equivalents of the standards

' «ГОСТ 7798-70» → «гост 7798» (the year goes), «DIN 933» → «din 933»
Private Function StdKey(ByVal s As String) As String
    Dim p As Long, t As String
    t = LCase(Trim(s))
    t = Replace(t, "исо", "iso")
    t = Replace(t, "ё", "е")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    p = InStr(t, "-")
    If p > 0 Then t = Left(t, p - 1)
    p = InStr(t, "–")
    If p > 0 Then t = Left(t, p - 1)
    StdKey = Trim(t)
End Function

Private Function EqAdd(sKey As String, sVal As String) As Boolean
    On Error GoTo Taken
    gSpEq.Add(sVal, sKey)
    EqAdd = True
    Exit Function
Taken:
    EqAdd = False
End Function

Private Sub EqLoad()
    Dim sh As Object, last As Long, d As Variant, i As Long, j As Long, k As Long, keys(2) As String, n As Long, cur As String, all As String, v As String
    gSpEq = New Collection
    gSpEqLoaded = True
    If Not gDoc.Sheets.hasByName(SH_STD) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_STD)
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Sub
    If last > ST_MAX Then last = ST_MAX
    d = sh.getCellRangeByPosition(0, 1, 2, last).getDataArray()
    For i = 0 To UBound(d)
        n = 0
        all = ""
        For j = 0 To 2
            v = Trim(CStr(d(i)(j)))
            If v <> "" Then
                keys(n) = StdKey(v)
                all = all & IIf(all <> "", ";", "") & keys(n)
                n = n + 1
            End If
        Next j
        If n >= 2 Then
            For j = 0 To n - 1
                cur = ""
                For k = 0 To n - 1
                    If k <> j Then cur = cur & IIf(cur <> "", ";", "") & keys(k)
                Next k
                If Not EqAdd(keys(j), cur) Then
                    ' the standard is in a row already: the equivalents of the new row come in too
                    gSpEq.Remove(keys(j))
                    gSpEq.Add(cur, keys(j))
                End If
            Next j
        End If
    Next i
End Sub

' the sheet «Стандарты» was edited: read again at the next use
Sub OnStdChange(Optional oTarget As Variant)
    gSpEqLoaded = False
End Sub

' the equivalents of a standard («DIN 933» → «iso 4017;гост 7798»; the keys as the search compares them); "" — none known
Function SpEquiv(sStd As String) As String
    Dim k As String
    If sStd = "" Then Exit Function
    If Not gSpEqLoaded Then EqLoad()
    k = StdKey(sStd)
    On Error GoTo None
    SpEquiv = gSpEq.Item(k)
    Exit Function
None:
    SpEquiv = ""
End Function

' ================================================================ one name, all of it

' the record of a name: (type, the index of the type, the size, the diameter, the length, the pitch, the standard, the class, the coating, the head, the equivalents of the standard)
Function SpRecord(sName As String) As Variant
    Dim t As String, idx As Long, sSize As String, dD As Double, dL As Double, dP As Double, std As String
    t = SpType(sName, idx)
    SpSize(sName, sSize, dD, dL, dP)
    std = SpStandard(sName)
    SpRecord = Array(t, idx, sSize, dD, dL, dP, std, SpClass(sName), SpCoating(sName), SpHead(sName), SpEquiv(std))
End Function

' the words that stand for the parts of a record in the search («м8х30» also as «м8», «30», «8х30»; «din 933» also as «din933», «933» and its equivalents)
Function SpTokens(rec As Variant) As String
    Dim s As String, sz As String, std As String, k As Variant, p As Long, eq As Variant
    sz = LCase(CStr(rec(2)))
    If sz <> "" Then
        sz = Replace(Replace(sz, "м", "м"), ",", ".")
        s = s & " " & Replace(sz, "×", "х")
        eq = Split(sz, "×")
        For Each k In eq
            If CStr(k) <> "" Then s = s & " " & CStr(k)
        Next k
        If Left(sz, 1) = "м" Then s = s & " " & Mid(Replace(sz, "×", "х"), 2)
    End If
    std = LCase(CStr(rec(6)))
    If std <> "" Then
        p = InStr(std, " ")
        s = s & " " & std & " " & Replace(std, " ", "") & " " & Mid(std, p + 1)
        ' the standard without its year («гост 7805-70» is «гост 7805» too)
        p = InStr(std, "-")
        If p > 0 Then
            std = Trim(Left(std, p - 1))
            s = s & " " & std & " " & Replace(std, " ", "") & " " & Mid(std, InStr(std, " ") + 1)
        End If
    End If
    If CStr(rec(10)) <> "" Then
        eq = Split(CStr(rec(10)), ";")
        For Each k In eq
            p = InStr(CStr(k), " ")
            s = s & " " & CStr(k) & " " & Replace(CStr(k), " ", "") & IIf(p > 0, " " & Mid(CStr(k), p + 1), "")
        Next k
    End If
    SpTokens = s
End Function

' test seams
Function TestSpec(sName As String) As String
    Dim r As Variant
    r = SpRecord(sName)
    TestSpec = r(0) & "|" & r(2) & "|" & r(3) & "|" & r(4) & "|" & r(5) & "|" & r(6) & "|" & r(7) & "|" & r(8) & "|" & r(9) & "|" & r(10)
End Function

Function TestSpTokens(sName As String) As String
    TestSpTokens = SpTokens(SpRecord(sName))
End Function
