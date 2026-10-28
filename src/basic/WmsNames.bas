' WmsNames — 1.1: names of the goods. (1) NormName — the name as it should be written: a name that was typed (or pasted) in several lines
' («Винт» / «м3х25» / «Din 912» / «A» / «2») is one line («Винт м3х25 DIN 912 A2»): the signs of a line break and the unusual spaces are a space,
' the spaces are not doubled, the class of the stainless steel («A 2», «А 4») and the size («М 8») that fell apart are put together, the standards
' («din», «гост») are written in capitals. (2) NameIssues — what is wrong with a name, in words. (3) Fold — the text as it is compared when
' searching: no difference of the case, «ё» and «е», the Latin letters that look like the Russian ones («M8x30» is «М8х30», «×» and «*» are «х»,
' «4,2» is «4.2»). Library WmsPlus.
Option Explicit

' the characters that are a blank in a name (codes): the controls, the non-breaking and the unusual spaces
Global gNmBl() As String
Global gNmBlN As Long

Private Sub NmInit()
    Dim i As Long, n As Long
    If gNmBlN > 0 Then Exit Sub
    ReDim gNmBl(70)
    For i = 1 To 31
        gNmBl(n) = Chr(i)
        n = n + 1
    Next i
    gNmBl(n) = Chr(127)
    n = n + 1
    gNmBl(n) = Chr(160)
    n = n + 1
    For i = 8192 To 8207
        gNmBl(n) = Chr(i)
        n = n + 1
    Next i
    gNmBl(n) = Chr(8232)
    n = n + 1
    gNmBl(n) = Chr(8233)
    n = n + 1
    gNmBl(n) = Chr(8239)
    n = n + 1
    gNmBl(n) = Chr(65279)
    n = n + 1
    gNmBlN = n
End Sub

' the text with the signs of a line break and the like turned into a space (not yet put in order otherwise)
Private Function BlankToSpace(ByVal s As String) As String
    Dim i As Long
    NmInit()
    For i = 0 To gNmBlN - 1
        If InStr(s, gNmBl(i)) > 0 Then s = Replace(s, gNmBl(i), " ")
    Next i
    BlankToSpace = s
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

Private Function NmIsSizeHead(t As String) As Boolean
    NmIsSizeHead = (t = "M" Or t = "m" Or t = "М" Or t = "м")
End Function

Private Function IsStandard(t As String) As Boolean
    Dim k As String
    k = LCase(t)
    IsStandard = (k = "din" Or k = "iso" Or k = "гост" Or k = "en" Or k = "дин")
End Function

' the name as it should be written (see the head of the module); "" stays ""
Function NormName(ByVal s As String) As String
    Dim a As Variant, out As String, i As Long, t As String, n As String, k As String, bJoin As Boolean, c As Long
    s = OneSpace(BlankToSpace(s))
    If s = "" Then Exit Function
    a = Split(s, " ")
    i = 0
    Do While i <= UBound(a)
        t = CStr(a(i))
        n = ""
        If i < UBound(a) Then n = CStr(a(i + 1))
        ' (And and Or evaluate both parts: the next word is looked at only when there is one)
        bJoin = False
        If n <> "" Then
            If IsClassHead(t) Then
                If IsClassTail(n) Then bJoin = True
            ElseIf NmIsSizeHead(t) Then
                c = Asc(Left(n, 1))
                If c >= 48 And c <= 57 Then bJoin = True
            End If
        End If
        If bJoin Then
            ' «A 2» → «A2» (the class of the stainless steel fell apart), «М 8х30» → «М8х30»
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

' the texts of an array (0-based, Strings) as Fold makes them, all at once: the engine of the language works on one long text, not on every row
Function FoldAll(arr As Variant) As Variant
    FoldAll = Split(Fold(Join(arr, "¦")), "¦")
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

' ================================================================ the name typed in several lines

' NormName for the names that need it at once (a line break, a tab, a non-breaking space, a doubled space); any other name is returned as it is — a list of
' thousands of names is not put through the whole of the rules
Function NormNameFast(ByVal s As String) As String
    If InStr(s, Chr(10)) + InStr(s, Chr(13)) + InStr(s, Chr(9)) + InStr(s, Chr(160)) + InStr(s, "  ") > 0 Then
        NormNameFast = NormName(s)
    Else
        NormNameFast = s
    End If
End Function

' the cell of a name of row r of sheet sSheet is put in order (one line): True when it was changed. Used when a name was just typed or pasted
Function FixNameCell(sSheet As String, r As Long, c As Integer) As Boolean
    Dim cell As Object, t As String, u As String
    On Error GoTo EH
    cell = gDoc.Sheets.getByName(sSheet).getCellByPosition(c, r)
    If cell.getType() <> com.sun.star.table.CellContentType.TEXT Then Exit Function
    t = cell.getString()
    If InStr(t, Chr(10)) + InStr(t, Chr(13)) + InStr(t, Chr(9)) + InStr(t, Chr(160)) + InStr(t, "  ") = 0 Then Exit Function
    u = NormName(t)
    If u <> t Then
        cell.setString(u)
        FixNameCell = True
    End If
    Exit Function
EH:
End Function

Public Const SH_NAMECHK = "Проверка названий"
Public Const NC_FIND_ROW = 1
Public Const NC_STAT_ROW = 2
Public Const NC_HEAD = 3
Public Const NC_ROW0 = 4
Public Const NC_COLS = 7
Public Const NC_MAX = 2000

' the name as the list shows it: a line break is «¶»
Private Function Shown(s As String) As String
    Shown = Replace(Replace(Replace(s, Chr(13) & Chr(10), "¶"), Chr(10), "¶"), Chr(13), "¶")
End Function

Sub BtnNamesCheck(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = NamesCheck()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub BtnNamesFix(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = NamesFixOpen()
    If Left(s, 4) = "ERR:" Then
        WmsUi.UiMessage(Mid(s, 5))
    Else
        WmsUi.UiMessage("Названия приведены к одной строке: " & Mid(s, 4) & ". Названия уже проведённых позиций WMS не меняет сама — у них в колонке «Что делать» написано как.")
        NamesCheck()
    End If
End Sub

' one row of the list of the names to look at
Private Function IssueRow(sWhere As String, r As Long, ei As String, nm As String, why As String, todo As String) As Variant
    IssueRow = Array(sWhere, r + 1, ei, Shown(nm), Shown(NormName(nm)), why, todo)
End Function

' the names of the registry, of «Заказы», «Иной приход» and «Номенклатура» that were typed in several lines or have a doubled space, a broken size or class;
' the list under the header of the sheet. OK:<how many> / ERR:…
' a value of a Collection by its key ("" — none); an error handler set inside a loop works for the first error only in this Basic, so each lookup is a call of its own
Private Function CollText(coll As Object, key As String) As String
    On Error GoTo Miss
    CollText = CStr(coll.Item(key))
    Exit Function
Miss:
    CollText = ""
End Function

' The possible doubles in «Номенклатура»: the same article in two rows; the same name (any case, the letters that look alike not told apart) in two rows; the same fastener (type, size,
' standard or its equivalent, class, coating, head) under two names. A row is added to out(n…) for the second row of each pair, with the first one named. Returns how many.
Private Function GoodsDoubles(out() As Variant, ByRef n As Long) As Long
    Dim sh As Object, last As Long, d As Variant, i As Long, nm As String, art As String, kA As String, kN As String, kP As String, first As String, rec As Variant, cc As String, why As String
    Dim collA As Object, collN As Object, collP As Object, cnt As Long, r As Variant
    If Not gDoc.Sheets.hasByName("Номенклатура") Then Exit Function
    sh = gDoc.Sheets.getByName("Номенклатура")
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, 2, last).getDataArray()
    collA = New Collection
    collN = New Collection
    collP = New Collection
    For i = 0 To UBound(d)
        r = d(i)
        nm = Trim(CStr(r(1)))
        art = Trim(CStr(r(2)))
        If nm <> "" Then
            why = ""
            first = ""
            If art <> "" Then
                kA = Fold(art)
                first = CollText(collA, kA)
                If first = "" Then
                    collA.Add(CStr(i + 2), kA)
                Else
                    why = "тот же артикул «" & art & "», что в строке " & first
                End If
            End If
            If why = "" Then
                kN = Fold(NormName(nm))
                first = CollText(collN, kN)
                If first = "" Then
                    collN.Add(CStr(i + 2), kN)
                Else
                    why = "то же название, что в строке " & first
                End If
            End If
            If why = "" Then
                rec = WmsSpec.SpRecordC(NormName(nm))
                If CStr(rec(0)) <> "" And CStr(rec(2)) <> "" Then
                    cc = CStr(rec(8))
                    If cc = "Оцинкованный" Then cc = "Цинк"
                    kP = rec(0) & "|" & rec(2) & "|" & rec(11) & "|" & rec(7) & "|" & cc & "|" & rec(9)
                    first = CollText(collP, kP)
                    If first = "" Then
                        collP.Add(CStr(i + 2), kP)
                    Else
                        why = "похоже на ту же деталь, что в строке " & first & " (тип, размер, стандарт и его аналоги, класс, покрытие совпадают)"
                    End If
                End If
            End If
            If why <> "" Then
                cnt = cnt + 1
                If n < NC_MAX Then
                    out(n) = IssueRow("Номенклатура", i + 1, "", nm, "возможный дубль: " & why, "справочник: объедините строки или поправьте артикул/название — WMS ничего не меняет сама")
                    n = n + 1
                End If
            End If
        End If
    Next i
    GoodsDoubles = cnt
End Function

Function NamesCheck() As String
    Dim sh As Object, d As Variant, i As Long, n As Long, out() As Variant, last As Long, nm As String, why As String, kind As String, cName As Long, cEi As Long, nAll As Long
    Dim sheetName As Variant, cur As Object, e As Long, cSt As Long, cSrc As Long, where As String, ei As String, nDup As Long
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_NAMECHK) Then
        NamesCheck = "ERR:в книге нет листа «" & SH_NAMECHK & "» — книга не обновлена до 1.1"
        Exit Function
    End If
    ReDim out(NC_MAX - 1)
    ' the registry
    sh = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(sh)
    cName = SC_NAME
    cSt = SC_SRC
    If last >= 1 Then
        d = sh.getCellRangeByPosition(0, 1, SC_SRC, last).getDataArray()
        For i = 0 To UBound(d)
            If CStr(d(i)(0)) <> "" Then
                nm = CStr(d(i)(cName))
                why = NameIssues(nm)
                If why <> "" Then
                    nAll = nAll + 1
                    kind = CStr(d(i)(cSt))
                    If n < NC_MAX Then
                        out(n) = IssueRow(SH_STOCK, i + 1, CStr(d(i)(0)), nm, why, "проведённая позиция: «Исправить» на строке-источнике (" & IIf(kind <> "", kind, "Заказы или Иной приход") & ")")
                        n = n + 1
                    End If
                End If
            End If
        Next i
    End If
    ' the rows of the orders and of «Иной приход»
    For Each sheetName In Array(SH_ORDERS, SH_SPECIAL, "Номенклатура")
        If gDoc.Sheets.hasByName(CStr(sheetName)) Then
            sh = gDoc.Sheets.getByName(CStr(sheetName))
            last = WmsOrders.LastRow(sh)
            If last >= 1 Then
                If sheetName = SH_ORDERS Then
                    cName = OC_NAME
                    cEi = OC_EI
                ElseIf sheetName = SH_SPECIAL Then
                    cName = XC_NAME
                    cEi = XC_EI
                Else
                    cName = 1
                    cEi = -1
                End If
                d = sh.getCellRangeByPosition(0, 1, IIf(cEi > cName, cEi, cName), last).getDataArray()
                For i = 0 To UBound(d)
                    nm = CStr(d(i)(cName))
                    If nm <> "" Then
                        why = NameIssues(nm)
                        If why <> "" Then
                            nAll = nAll + 1
                            If cEi < 0 Then
                                where = "справочник: кнопка «Исправить открытые строки»"
                            ElseIf CStr(d(i)(cEi)) <> "" Then
                                where = "проведено: «Исправить» на этой строке"
                            Else
                                where = "ещё не проведено: кнопка «Исправить открытые строки»"
                            End If
                            ei = ""
                            If cEi >= 0 Then ei = CStr(d(i)(cEi))
                            If n < NC_MAX Then
                                out(n) = IssueRow(CStr(sheetName), i + 1, ei, nm, why, where)
                                n = n + 1
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next sheetName
    ' 1.1: doubles in the directory «Номенклатура» (the one place where a double is a mistake: the registry holds one line per receipt on purpose)
    nDup = GoodsDoubles(out, n)
    nAll = nAll + nDup
    sh = gDoc.Sheets.getByName(SH_NAMECHK)
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    e = cur.getRangeAddress().EndRow
    If e >= NC_ROW0 Then sh.getCellRangeByPosition(0, NC_ROW0, NC_COLS - 1, e).clearContents(1 + 2 + 4 + 16)
    If n > 0 Then
        ReDim Preserve out(n - 1)
        sh.getCellRangeByPosition(0, NC_ROW0, NC_COLS - 1, NC_ROW0 + n - 1).setDataArray(out)
    End If
    sh.getCellByPosition(0, NC_STAT_ROW).setString(IIf(nAll = 0, "Названия в порядке: разбитых на строки, с лишними пробелами и разорванным размером нет.", _
        "Названий, которые стоит исправить: " & nAll & IIf(nAll > n, " (показаны первые " & n & ")", "") & IIf(nDup > 0, ", из них возможных дублей в «Номенклатуре»: " & nDup, "") & ". Проверено " & Format(Now(), "DD.MM.YYYY HH:MM") & ". «¶» — место, где название разбито на строки."))
    NamesCheck = "OK:" & nAll
    gUiLastMsg = NamesCheck
    Exit Function
EH:
    NamesCheck = "ERR:названия не проверены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = NamesCheck
End Function

' the names of the rows that are not posted yet («Заказы», «Иной приход») and of «Номенклатура» put in order (one line); OK:<how many>
Function NamesFixOpen() As String
    Dim sh As Object, i As Long, last As Long, n As Long, nm As String, u As String, k As String, c As Object
    On Error GoTo EH
    If gState <> "CLEAN" Then
        NamesFixOpen = "ERR:" & IIf(gBlockText <> "", gBlockText, "WMS не готова (" & gState & ")")
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    last = WmsOrders.LastRow(sh)
    For i = 1 To last
        k = WmsOrders.OrderRowKind(i)
        If k = "OPEN" Then
            c = sh.getCellByPosition(OC_NAME, i)
            nm = c.getString()
            u = NormNameFast(nm)
            If nm <> "" And u <> nm Then
                c.setString(u)
                n = n + 1
            End If
        End If
    Next i
    sh = gDoc.Sheets.getByName(SH_SPECIAL)
    last = WmsOrders.LastRow(sh)
    For i = 1 To last
        k = WmsSpecial.SpecialRowKind(i)
        If k = "OPEN" Then
            c = sh.getCellByPosition(XC_NAME, i)
            nm = c.getString()
            u = NormNameFast(nm)
            If nm <> "" And u <> nm Then
                c.setString(u)
                n = n + 1
            End If
        End If
    Next i
    If gDoc.Sheets.hasByName("Номенклатура") Then
        sh = gDoc.Sheets.getByName("Номенклатура")
        last = WmsOrders.LastRow(sh)
        For i = 1 To last
            c = sh.getCellByPosition(1, i)
            nm = c.getString()
            u = NormNameFast(nm)
            If nm <> "" And u <> nm Then
                c.setString(u)
                n = n + 1
            End If
        Next i
    End If
    NamesFixOpen = "OK:" & n
    gUiLastMsg = NamesFixOpen
    Exit Function
EH:
    NamesFixOpen = "ERR:названия не исправлены: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = NamesFixOpen
End Function
