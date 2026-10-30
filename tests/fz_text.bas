Option Explicit
' ad-hoc fuzz of the text functions of WMS: every name of sList (Chr(10) between) goes through every function of sWhat (comma-separated); the answer: the failures grouped by function
Function Run(sList As String, sWhat As String) As String
    Dim a As Variant, f As Variant, i As Long, nm As String, e As String, out As String, nBad As Long, w As Variant
    a = Split(sList, Chr(10))
    w = Split(sWhat, ",")
    For i = 0 To UBound(a)
        nm = CStr(a(i))
        For Each f In w
            e = Call1(Trim(CStr(f)), nm)
            If e <> "" Then
                nBad = nBad + 1
                If nBad <= 60 Then out = out & Chr(10) & nm & " ¦ " & e
            End If
        Next f
    Next i
    Run = (UBound(a) + 1) & " имён, ошибок " & nBad & out
End Function

Private Function Call1(sFn As String, nm As String) As String
    Dim v As Variant, r As Variant, s1 As String, d1 As Double, d2 As Double, d3 As Double, d4 As Double, hays(2) As String, nms(2) As String, b As Boolean, msg As String, idx As Long
    On Error GoTo EH
    Select Case sFn
    Case "NormName": v = WmsNames.NormName(nm)
    Case "NameIssues": v = WmsNames.NameIssues(nm)
    Case "Fold": v = WmsNames.Fold(nm)
    Case "FoldWords": v = WmsNames.FoldWords(nm)
    Case "NormNameFast": v = WmsNames.NormNameFast(nm)
    Case "Prep": v = WmsSearch.Prep(nm)
    Case "Stem": v = WmsSearch.Stem(nm)
    Case "QueryNeeds": v = WmsSearch.QueryNeeds(nm)
    Case "MatchScores"
        hays(0) = " болт м8х30 din 933 "
        hays(1) = " фреза концевая ø10 "
        hays(2) = " 4 зуба "
        nms(0) = hays(0)
        nms(1) = hays(1)
        nms(2) = hays(2)
        v = WmsSearch.MatchScores(nm, hays, nms, 3, Empty)
    Case "SpRecord"
        r = WmsSpec.SpRecord(nm)
        v = WmsSpec.SpTokens(r)
    Case "SpRecordC": v = WmsSpec.SpRecordC(nm)
    Case "SpSize": WmsSpec.SpSize(nm, s1, d1, d2, d3)
    Case "SpStandard": v = WmsSpec.SpStandard(nm)
    Case "SpEquiv": v = WmsSpec.SpEquiv(nm)
    Case "SpCanon": v = WmsSpec.SpCanon(nm)
    Case "SpType": v = WmsSpec.SpType(nm, idx)
    Case "SpClass": v = WmsSpec.SpClass(nm)
    Case "SpCoating": v = WmsSpec.SpCoating(nm)
    Case "SpHead": v = WmsSpec.SpHead(nm)
    Case "CategorySuggest": v = WmsCat.CategorySuggest(nm)
    Case "CategoryOf": v = WmsCat.CategoryOf(nm, "", "")
    Case "PkWords": v = WmsPack.PkWords(nm)
    Case "PackGuess": WmsPack.PackGuess(nm, d1, s1)
    Case "PkWeight": b = WmsPack.PkWeight(nm, "шт", d1, s1, msg)
    Case "UnitMatch": b = WmsPack.UnitMatch(nm, "шт")
    Case "FirstPart": v = WmsUi.FirstPart(nm)
    Case "TlRecord": v = WmsTool.TlRecord(nm)
    Case "TextFilterOk": b = WmsCat.TextFilterOk(nm, " болт м8 ")
    Case "NumFilterOk": b = WmsCat.NumFilterOk(nm, 8)
    Case Else: Call1 = "нет функции " & sFn
    End Select
    Exit Function
EH:
    Call1 = sFn & ": " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function
