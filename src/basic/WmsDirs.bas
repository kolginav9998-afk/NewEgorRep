' WmsDirs — 1.1: справочники кладовщика из другой книги. The directories («Получатели», «Номенклатура», «Акты», «Категории» …) are the storekeeper's own sheets: they are not in the
' journal, so a return to a copy («Вернуться к копии…») brings back the directories as they were when the copy was made. Every copy, and the book put aside by the return, holds
' the directories as they were at its own time: WMS reads them from there (the book is opened for reading, hidden, without macros), says what differs and, when the storekeeper agrees,
' writes the values of the differing directories into the open book. The journal, the registry and every journaled sheet are not touched; what the open book had in a directory is
' still in the copy it came from. Library WmsPlus.
Option Explicit

' the directories: sheet | first column | last column | columns not copied from | to (the counting of WMS, not the input) ; -1 — none
Private Function Specs() As Variant
    Specs = Array(Array(SH_RCPT, 0, 12, -1, -1), Array(SH_TOPO, 0, 40, -1, -1), Array(SH_PLACES, 0, 12, -1, -1), Array(SH_SUPPLIERS, 0, 20, -1, -1), _
        Array(SH_GOODS, 0, 24, GD_STOCK, GD_UPD), Array(SH_KITS, 0, 12, -1, -1), Array("Сборка деталей", 0, 12, -1, -1), Array(SH_RESERVES, 0, 12, -1, -1), Array(WmsDocs.SH_ACTS, 0, 12, -1, -1), _
        Array(WmsDocs.SH_DAYREP, 0, 12, -1, -1), Array(WmsCat.SH_CATS, 0, 2, -1, -1), Array(WmsSpec.SH_STD, 0, 4, -1, -1), Array(WmsSearch.SH_SYN, 0, 2, -1, -1), _
        Array(WmsSettings.SH_SET, 0, 3, -1, -1))
End Function

' the last used row and column of a sheet
Private Sub UsedEnd(sh As Object, ByRef er As Long, ByRef ec As Long)
    Dim cur As Object
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    er = cur.getRangeAddress().EndRow
    ec = cur.getRangeAddress().EndColumn
End Sub

' the headers of a sheet: its short texts of the first row (a change of the structure of a sheet shows in them)
Private Function HeadText(sh As Object, ec As Long) As String
    Dim d As Variant, i As Long, t As String
    d = sh.getCellRangeByPosition(0, 0, ec, 0).getDataArray()(0)
    For i = 0 To UBound(d)
        If VarType(d(i)) = 8 Then
            If Len(d(i)) > 0 And Len(d(i)) <= 80 Then t = t & "|" & i & ":" & d(i)
        End If
    Next i
    HeadText = t
End Function

' the data of the columns c0…c1 of the rows 0…er of a sheet as one text (the numbers and the texts as they are; the counting columns left out)
Private Function DataText(sh As Object, er As Long, c0 As Long, c1 As Long, s0 As Long, s1 As Long) As String
    Dim d As Variant, i As Long, j As Long, t() As String, ln As String
    If er < 0 Then Exit Function
    d = sh.getCellRangeByPosition(c0, 0, c1, er).getDataArray()
    ReDim t(UBound(d))
    For i = 0 To UBound(d)
        ln = ""
        For j = 0 To UBound(d(i))
            If s0 < 0 Or j + c0 < s0 Or j + c0 > s1 Then ln = ln & CStr(d(i)(j)) & Chr(1)
        Next j
        t(i) = ln
    Next i
    DataText = Join(t, Chr(2))
End Function

' How the directories of the book src differ from the open one: "" — nothing; otherwise lines «sheet|status|text»; status: DIFF (different), SKIP (the structure of the sheet is another).
' The open book is only read.
Function DirsCompare(src As Object, ByRef nDiff As Long) As String
    Dim sp As Variant, i As Long, nm As String, so As Object, sd As Object, er1 As Long, ec1 As Long, er2 As Long, ec2 As Long, c0 As Long, c1 As Long, s0 As Long, s1 As Long, out As String
    Dim t1 As String, t2 As String
    sp = Specs()
    nDiff = 0
    For i = 0 To UBound(sp)
        nm = CStr(sp(i)(0))
        c0 = sp(i)(1)
        c1 = sp(i)(2)
        s0 = sp(i)(3)
        s1 = sp(i)(4)
        If gDoc.Sheets.hasByName(nm) And src.Sheets.hasByName(nm) Then
            sd = gDoc.Sheets.getByName(nm)
            so = src.Sheets.getByName(nm)
            If Not sd.isProtected() Then
                UsedEnd(sd, er1, ec1)
                UsedEnd(so, er2, ec2)
                If HeadText(sd, IIf(ec1 > c1, c1, ec1)) <> HeadText(so, IIf(ec2 > c1, c1, ec2)) Then
                    out = out & nm & "|SKIP|структура листа другая" & Chr(10)
                Else
                    t1 = DataText(sd, IIf(er1 > er2, er1, er2), c0, c1, s0, s1)
                    t2 = DataText(so, IIf(er1 > er2, er1, er2), c0, c1, s0, s1)
                    If t1 <> t2 Then
                        nDiff = nDiff + 1
                        out = out & nm & "|DIFF|строк было " & (er1 + 1) & ", в той книге " & (er2 + 1) & Chr(10)
                    End If
                End If
            End If
        End If
    Next i
    DirsCompare = out
End Function

' The values of the directory nm from src written into the open book (the columns not copied and the formats stay); the formulas of the source are written as formulas.
' The rows below the source are cleared. True when written
Private Function CopyDir(src As Object, nm As String, c0 As Long, c1 As Long, s0 As Long, s1 As Long) As Boolean
    Dim sd As Object, so As Object, er1 As Long, ec1 As Long, er2 As Long, ec2 As Long, d As Variant, f As Variant, i As Long, j As Long, rg As Object, cc As Long, ce As Long
    sd = gDoc.Sheets.getByName(nm)
    so = src.Sheets.getByName(nm)
    UsedEnd(sd, er1, ec1)
    UsedEnd(so, er2, ec2)
    ce = IIf(ec2 > c1, c1, ec2)
    If er2 < 0 Then Exit Function
    ' what the open sheet had in its columns is cleared first (values, texts, formulas — not the formats), the columns not copied are not touched
    For cc = c0 To IIf(ec1 > c1, c1, ec1)
        If s0 < 0 Or cc < s0 Or cc > s1 Then sd.getCellRangeByPosition(cc, 0, cc, IIf(er1 > er2, er1, er2)).clearContents(1 + 2 + 4 + 16)
    Next cc
    d = so.getCellRangeByPosition(c0, 0, ce, er2).getDataArray()
    f = so.getCellRangeByPosition(c0, 0, ce, er2).getFormulaArray()
    For i = 0 To UBound(d)
        For j = 0 To UBound(d(i))
            cc = c0 + j
            If s0 < 0 Or cc < s0 Or cc > s1 Then
                If Left(CStr(f(i)(j)), 1) = "=" Then
                    sd.getCellByPosition(cc, i).setFormula(CStr(f(i)(j)))
                ElseIf VarType(d(i)(j)) = 5 Then
                    sd.getCellByPosition(cc, i).setValue(d(i)(j))
                ElseIf CStr(d(i)(j)) <> "" Then
                    sd.getCellByPosition(cc, i).setString(CStr(d(i)(j)))
                End If
            End If
        Next j
    Next i
    CopyDir = True
End Function

' The directories of the book url (a file .ods: a copy, the book put aside by a return, a book from a flash) put into the open book. bAsk — the storekeeper is asked what will be done
' (the tests answer by the seam of WmsPrimeUi). OK:<text> / SKIP:<why> / ERR:<why>
Function DirsFromBook(url As String, bAsk As Boolean) As String
    Dim args(2) As New com.sun.star.beans.PropertyValue, src As Object, cmp As String, nDiff As Long, a As Variant, i As Long, k As Integer, sp As Variant, nm As String, j As Long
    Dim list As String, nDone As Long, done As String, skipped As String, p As Variant
    On Error GoTo EH
    WmsInit()
    If Not gSfa.exists(url) Then
        DirsFromBook = "ERR:книги «" & ConvertFromURL(url) & "» нет"
        Exit Function
    End If
    If url = gDoc.getURL() Then
        DirsFromBook = "ERR:это та же книга"
        Exit Function
    End If
    args(0).Name = "Hidden"
    args(0).Value = True
    args(1).Name = "ReadOnly"
    args(1).Value = True
    args(2).Name = "MacroExecutionMode"
    args(2).Value = 0
    src = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    If IsNull(src) Then
        DirsFromBook = "ERR:книга «" & ConvertFromURL(url) & "» не открылась"
        Exit Function
    End If
    cmp = DirsCompare(src, nDiff)
    a = Split(cmp, Chr(10))
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            p = Split(a(i), "|")
            If p(1) = "DIFF" Then
                list = list & "  • " & p(0) & " (" & p(2) & ")" & Chr(10)
            Else
                skipped = skipped & IIf(skipped <> "", ", ", "") & p(0)
            End If
        End If
    Next i
    If nDiff = 0 Then
        src.close(True)
        DirsFromBook = "OK:справочники те же" & IIf(skipped <> "", " (не сравнивались — структура другая: " & skipped & ")", "")
        gUiLastMsg = DirsFromBook
        Exit Function
    End If
    If bAsk Then
        k = WmsPrimeUi.UiButtons("Справочники", "Справочники этой книги и книги «" & FileNameOf(url) & "» отличаются:" & Chr(10) & Chr(10) & list & Chr(10) _
            & "Подставить их значения из «" & FileNameOf(url) & "» в эту книгу?" & Chr(10) _
            & "Журнал, остатки и операции не меняются. Что было в справочниках сейчас — останется в той книге, из которой эта книга поднята." & IIf(skipped <> "", Chr(10) & "Не сравнивались (структура другая): " & skipped & ".", ""), _
            Array("Подставить справочники", "Оставить как есть"), Array(1, 0))
        If k <> 1 Then
            src.close(True)
            DirsFromBook = "SKIP:справочники оставлены как есть"
            gUiLastMsg = DirsFromBook
            Exit Function
        End If
    End If
    sp = Specs()
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            p = Split(a(i), "|")
            If p(1) = "DIFF" Then
                For j = 0 To UBound(sp)
                    If CStr(sp(j)(0)) = CStr(p(0)) Then
                        If CopyDir(src, CStr(p(0)), sp(j)(1), sp(j)(2), sp(j)(3), sp(j)(4)) Then
                            nDone = nDone + 1
                            done = done & IIf(done <> "", ", ", "") & p(0)
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    src.close(True)
    DirsFromBook = "OK:справочники подставлены из «" & FileNameOf(url) & "»: " & done
    gUiLastMsg = DirsFromBook
    Exit Function
EH:
    DirsFromBook = "ERR:справочники не подставлены: " & Error$ & " (строка " & Erl & ")"
    On Error Resume Next
    src.close(True)
End Function

' the button «Справочники из копии…» of «Обслуживание»: the storekeeper picks a book (a copy of WMS_Backups, the book put aside by a return, a flash copy)
Sub BtnDirsFromBook(Optional oEvent As Variant)
    Dim fp As Object, s As String
    WmsInit()
    On Error GoTo EH
    fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    fp.setTitle("Книга, из которой взять справочники")
    fp.appendFilter("Книги WMS (*.ods)", "*.ods")
    fp.setCurrentFilter("Книги WMS (*.ods)")
    If gBDir <> "" Then fp.setDisplayDirectory(gBDir)
    If fp.execute() <> 1 Then Exit Sub
    s = DirsFromBook(fp.getFiles()(0), True)
    WmsUi.UiMessage(Mid(s, InStr(s, ":") + 1))
    Exit Sub
EH:
    WmsUi.UiMessage("Справочники не подставлены: " & Error$)
End Sub

' test seam: the directories of the book at url compared with the open one: "n|lines"
Function TestDirsCompare(url As String) As String
    Dim args(2) As New com.sun.star.beans.PropertyValue, src As Object, n As Long, t As String
    args(0).Name = "Hidden"
    args(0).Value = True
    args(1).Name = "ReadOnly"
    args(1).Value = True
    args(2).Name = "MacroExecutionMode"
    args(2).Value = 0
    src = StarDesktop.loadComponentFromURL(url, "_blank", 0, args())
    t = DirsCompare(src, n)
    src.close(True)
    TestDirsCompare = n & "|" & Replace(t, Chr(10), ";")
End Function
