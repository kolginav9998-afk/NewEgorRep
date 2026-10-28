' WmsNames — 1.1: names of the goods. (1) NormName — the name as it should be written: a name that was typed (or pasted) in several lines
' («Винт» / «м3х25» / «Din 912» / «A» / «2») is one line («Винт м3х25 DIN 912 A2»): the signs of a line break and the unusual spaces are a space,
' the spaces are not doubled, the class of the stainless steel («A 2», «А 4») and the size («М 8») that fell apart are put together, the standards
' («din», «гост») are written in capitals. (2) NameIssues — what is wrong with a name, in words. (3) Fold — the text as it is compared when
' searching: no difference of the case, «ё» and «е», the Latin letters that look like the Russian ones («M8x30» is «М8х30», «×» and «*» are «х»,
' «4,2» is «4.2»). Library WmsPlus.
Option Explicit

' the sign of a line break, a tab, a non-breaking space and the like (code) — a space in a name
Private Function IsBlankCode(c As Long) As Boolean
    IsBlankCode = (c < 32 Or c = 127 Or c = 160 Or (c >= 8192 And c <= 8207) Or c = 8232 Or c = 8233 Or c = 8239 Or c = 65279)
End Function

' the text with the signs of a line break and the like turned into a space (not yet put in order otherwise)
Private Function BlankToSpace(ByVal s As String) As String
    Dim i As Long, ch As String, bad As Boolean
    For i = 1 To Len(s)
        If IsBlankCode(Asc(Mid(s, i, 1))) Then
            bad = True
            Exit For
        End If
    Next i
    If Not bad Then
        BlankToSpace = s
        Exit Function
    End If
    BlankToSpace = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If IsBlankCode(Asc(ch)) Then ch = " "
        BlankToSpace = BlankToSpace & ch
    Next i
End Function

Private Function OneSpace(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    OneSpace = Trim(s)
End Function

Private Function IsClassHead(t As String) As Boolean
    IsClassHead = (t = "A" Or t = "a" Or t = "А" Or t = "а")
End Function

Private Function IsClassTail(t As String) As Boolean
    IsClassTail = (t = "2" Or t = "4" Or Left(t, 2) = "2-" Or Left(t, 2) = "4-")
End Function

Private Function IsSizeHead(t As String) As Boolean
    IsSizeHead = (t = "M" Or t = "m" Or t = "М" Or t = "м")
End Function

Private Function IsStandard(t As String) As Boolean
    Dim k As String
    k = LCase(t)
    IsStandard = (k = "din" Or k = "iso" Or k = "гост" Or k = "en" Or k = "дин")
End Function

' the name as it should be written (see the head of the module); "" stays ""
Function NormName(ByVal s As String) As String
    Dim a As Variant, out As String, i As Long, t As String, n As String, k As String
    s = OneSpace(BlankToSpace(s))
    If s = "" Then Exit Function
    a = Split(s, " ")
    i = 0
    Do While i <= UBound(a)
        t = CStr(a(i))
        n = ""
        If i < UBound(a) Then n = CStr(a(i + 1))
        If n <> "" And IsClassHead(t) And IsClassTail(n) Then
            ' «A 2» → «A2» (the class of the stainless steel fell apart)
            t = t & n
            i = i + 1
        ElseIf n <> "" And IsSizeHead(t) And Asc(Left(n, 1)) >= 48 And Asc(Left(n, 1)) <= 57 Then
            ' «М 8х30» → «М8х30»
            t = t & n
            i = i + 1
        ElseIf IsStandard(t) Then
            k = UCase(t)
            If k = "ДИН" Then k = "DIN"
            t = k
        End If
        out = out & IIf(out <> "", " ", "") & t
        i = i + 1
    Loop
    NormName = out
End Function

' what is wrong with a name, in words ("" — nothing); the same rules as NormName
Function NameIssues(ByVal s As String) As String
    Dim why As String, i As Long, c As Long, lb As Boolean, w As String, t As String
    For i = 1 To Len(s)
        c = Asc(Mid(s, i, 1))
        If c = 10 Or c = 13 Or c = 11 Or c = 12 Or c = 8232 Or c = 8233 Then lb = True
    Next i
    If lb Then why = "название разбито на несколько строк"
    t = BlankToSpace(s)
    If Trim(t) <> t Or InStr(t, "  ") > 0 Then why = why & IIf(why <> "", "; ", "") & "лишние пробелы"
    If NormName(s) <> OneSpace(t) Then why = why & IIf(why <> "", "; ", "") & "разорванный размер, класс или стандарт («A 2», «М 8», «din»)"
    NameIssues = why
End Function

' the text as it is compared when searching (see the head of the module)
Function Fold(ByVal s As String) As String
    s = LCase(BlankToSpace(s))
    s = Replace(s, "ё", "е")
    s = Replace(s, "a", "а")
    s = Replace(s, "b", "в")
    s = Replace(s, "c", "с")
    s = Replace(s, "e", "е")
    s = Replace(s, "h", "н")
    s = Replace(s, "k", "к")
    s = Replace(s, "m", "м")
    s = Replace(s, "o", "о")
    s = Replace(s, "p", "р")
    s = Replace(s, "t", "т")
    s = Replace(s, "x", "х")
    s = Replace(s, "y", "у")
    s = Replace(s, "×", "х")
    s = Replace(s, "*", "х")
    s = Replace(s, ",", ".")
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    Fold = OneSpace(s)
End Function

' the words of a search text (folded), an empty array when there are none
Function FoldWords(ByVal s As String) As Variant
    s = Fold(s)
    If s = "" Then
        FoldWords = Array()
    Else
        FoldWords = Split(s, " ")
    End If
End Function

' True when every word is in hay (hay — already folded)
Function HasWords(hay As String, words As Variant) As Boolean
    Dim i As Long
    HasWords = True
    For i = 0 To UBound(words)
        If InStr(hay, CStr(words(i))) = 0 Then
            HasWords = False
            Exit Function
        End If
    Next i
End Function

' test seams
Function TestNorm(s As String) As String
    TestNorm = NormName(s)
End Function

Function TestIssues(s As String) As String
    TestIssues = NameIssues(s)
End Function

Function TestFold(s As String) As String
    TestFold = Fold(s)
End Function
