' WmsTestOps — TEST ONLY. Installed by tools/build_ods.py --test, never into a production book.
' A synthetic operation on the _TST sheet exercises the Phase 1 engine exactly like a Phase 2 operation will:
' phase 1 (validate/normalize/calculate, no mutation) → plan → WmsCore.ApplyOperation.
' _TST: A key (from NEXT_NO), B text (input), C quantity (text input), D control/status.
Option Explicit

Private Const TST = "_TST"

Function TestPost(r As Long) As String
    Dim why As String, sh As Object, txt As String, q As Double, st As Integer, key As Long
    why = PostingBlockReason()
    If why <> "" Then
        TestPost = why
        Exit Function
    End If
    sh = ThisComponent.Sheets.getByName(TST)
    ' phase 1: no mutation
    If Left(sh.getCellByPosition(3, r).getString(), 5) = "КОПИЯ" Then
        TestPost = "SKIP:строка помечена как КОПИЯ"
        Exit Function
    End If
    If sh.getCellByPosition(0, r).getString() <> "" Then
        TestPost = "SKIP:уже проведено №" & sh.getCellByPosition(0, r).getString()
        Exit Function
    End If
    txt = sh.getCellByPosition(1, r).getString()
    If txt = "" Then
        TestPost = "ERR:не указан текст"
        Exit Function
    End If
    st = ParseQtyCell(sh.getCellByPosition(2, r), q, why)
    If st <> 0 Then
        TestPost = "ERR:" & why
        Exit Function
    End If
    If q > SysNum(SK_MAX_QTY) Then
        TestPost = "ERR:количество больше порога MAX_QTY"
        Exit Function
    End If
    key = CLng(SysNum(SK_NEXT_NO))
    ' plan: every change of the operation, nothing written yet
    PlanBegin("TEST", CStr(key))
    PlanField("NO", key)
    PlanField("ROW", r)
    PlanField("TEXT", txt)
    PlanField("QTY", q)
    PlanInput(TST, r, 1)
    PlanSetValue(TST, r, 2, q, True)
    PlanSetValue(TST, r, 0, key, False)
    PlanSetValue(TST, r, 3, "Проведено", False)
    PlanSetValue(SYS_SHEET, SK_NEXT_NO, 1, key + 1, False)
    PlanLock(TST, r, 0, 3, True)
    TestPost = ApplyOperation(0, 0)
End Function

' posts rows r0..r1 one by one; returns "OK=n;ERR=m;SKIP=k;last=<last result>"
Function TestPostRange(r0 As Long, r1 As Long) As String
    Dim r As Long, res As String, nOk As Long, nErr As Long, nSkip As Long
    For r = r0 To r1
        res = TestPost(r)
        If Left(res, 3) = "OK:" Then
            nOk = nOk + 1
        ElseIf Left(res, 5) = "SKIP:" Then
            nSkip = nSkip + 1
        Else
            nErr = nErr + 1
        End If
    Next r
    TestPostRange = "OK=" & nOk & ";ERR=" & nErr & ";SKIP=" & nSkip & ";last=" & res
End Function

' a plan whose second write carries an unknown value encoding: the first write is already applied when the
' second one fails — a natural error before the journal, the engine must roll the first write back
Function TestPostBrokenPlan(r As Long) As String
    Dim why As String
    why = PostingBlockReason()
    If why <> "" Then
        TestPostBrokenPlan = why
        Exit Function
    End If
    PlanBegin("TEST", "broken")
    PlanField("NO", -1)
    PlanSetValue(TST, r, 3, "частично", False)
    PlanAdd2("V", TST, r, "1", "E", "Q?")
    TestPostBrokenPlan = ApplyOperation(0, 0)
End Function

Private Sub PlanAdd2(k As String, sh As String, r As Long, c As String, b As String, a As String)
    ' extends the plan with a raw record (PlanAdd is private to WmsCore)
    If gPlanN > UBound(gPW_Kind) Then
        ReDim Preserve gPW_Kind(2 * gPlanN + 1)
        ReDim Preserve gPW_Sheet(2 * gPlanN + 1)
        ReDim Preserve gPW_Row(2 * gPlanN + 1)
        ReDim Preserve gPW_Col(2 * gPlanN + 1)
        ReDim Preserve gPW_Before(2 * gPlanN + 1)
        ReDim Preserve gPW_After(2 * gPlanN + 1)
        ReDim Preserve gPW_Restorable(2 * gPlanN + 1)
    End If
    gPW_Kind(gPlanN) = k
    gPW_Sheet(gPlanN) = sh
    gPW_Row(gPlanN) = r
    gPW_Col(gPlanN) = c
    gPW_Before(gPlanN) = b
    gPW_After(gPlanN) = a
    gPW_Restorable(gPlanN) = False
    gPlanN = gPlanN + 1
End Sub

' writes service-like text with every special character into row r (before-image / journal round trip)
Function TestPostSpecial(r As Long, sText As String) As String
    Dim why As String, key As Long
    why = PostingBlockReason()
    If why <> "" Then
        TestPostSpecial = why
        Exit Function
    End If
    key = CLng(SysNum(SK_NEXT_NO))
    PlanBegin("TEST", CStr(key))
    PlanField("NO", key)
    PlanField("TEXT", sText)
    PlanSetValue(TST, r, 1, sText, False)
    PlanSetValue(TST, r, 2, 1, False)
    PlanSetValue(TST, r, 0, key, False)
    PlanSetValue(TST, r, 3, "Проведено", False)
    PlanSetValue(SYS_SHEET, SK_NEXT_NO, 1, key + 1, False)
    TestPostSpecial = ApplyOperation(0, 0)
End Function

' parser vectors for the Python test: "status|value|message"
Function TestParse(s As String) As String
    Dim v As Double, msg As String, st As Integer
    st = ParseQtyText(s, v, msg)
    TestParse = st & "|" & NumStr(v) & "|" & msg
End Function

Function TestParseCell(r As Long) As String
    Dim v As Double, msg As String, st As Integer
    st = ParseQtyCell(ThisComponent.Sheets.getByName(TST).getCellByPosition(2, r), v, msg)
    TestParseCell = st & "|" & NumStr(v) & "|" & msg
End Function

Function TestEscRoundTrip(s As String) As Boolean
    TestEscRoundTrip = (Unesc(Esc(s)) = s)
End Function
