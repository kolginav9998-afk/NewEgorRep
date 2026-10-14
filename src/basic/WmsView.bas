' WmsView — the common part of the view sheets of PRIME 0.8 («Дашборд», «Поиск», «Карта склада»; D-094, D-095): they are
' built from the sheets of the book when asked (a button, the sheet shown), never journaled, never read by an operation.
' A view sheet is protected; writing it does not mark the book changed (a view is not data).
Option Explicit

Global gVwWasMod As Boolean
Global gVwT0 As Long

' the view sheet sName open for writing (unprotected); ViewDone protects it again
Function ViewOpen(sName As String) As Object
    Dim sh As Object
    WmsInit()
    gVwWasMod = gDoc.isModified()
    gVwT0 = GetSystemTicks()
    sh = gDoc.Sheets.getByName(sName)
    If sh.isProtected() Then sh.unprotect(PROTECT_PWD)
    ViewOpen = sh
End Function

Sub ViewDone(sh As Object)
    On Error Resume Next
    If Not sh.isProtected() Then sh.protect(PROTECT_PWD)
    If Not gVwWasMod Then gDoc.setModified(False)
End Sub

Function ViewMs() As Long
    ViewMs = GetSystemTicks() - gVwT0
End Function

' today as the serial number of a date cell
Function Today0() As Double
    Today0 = CDbl(DateSerial(Year(Now()), Month(Now()), Day(Now())))
End Function

Function ViewDay(d As Double) As String
    If d <= 0 Then
        ViewDay = ""
    Else
        ViewDay = Format(CDate(d), "DD.MM.YYYY")
    End If
End Function

' the date of a cell value: a number (a date cell) or a text «ДД.ММ.ГГГГ» / «ГГГГ-ММ-ДД»; 0 when there is none
Function DateOf(v As Variant) As Double
    Dim s As String, a As Variant
    If VarType(v) = 5 Or VarType(v) = 7 Then
        If v > 1 Then DateOf = Int(CDbl(v))
        Exit Function
    End If
    s = Trim(CStr(v))
    If Len(s) < 8 Then Exit Function
    On Error GoTo Bad
    If Mid(s, 5, 1) = "-" Then
        DateOf = CDbl(DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 6, 2)), CInt(Mid(s, 9, 2))))
    ElseIf Mid(s, 3, 1) = "." Then
        a = Split(Left(s, 10), ".")
        If UBound(a) = 2 Then DateOf = CDbl(DateSerial(CInt(a(2)), CInt(a(1)), CInt(a(0))))
    End If
    Exit Function
Bad:
    DateOf = 0
End Function

' the number of an EI text: «ЕИ-00000427», «ЕИ-427», «еи 427», «427»; 0 when it is not an EI below nextEI
Function EiNo(ByVal s As String, nextEI As Long) As Long
    Dim t As String, i As Integer, ch As String, x As Double
    s = UCase(Trim(s))
    If Left(s, 3) = "ЕИ-" Or Left(s, 3) = "ЕИ " Then
        s = Trim(Mid(s, 4))
    ElseIf Left(s, 2) = "ЕИ" Then
        s = Trim(Mid(s, 3))
    End If
    If s = "" Or Len(s) > 9 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    x = CDbl(s)
    If x >= 1 And x < nextEI Then EiNo = CLng(x)
End Function

' a place as the storekeeper typed it → «ряд-секция-ярус»: 5-12-3 for «5.12.3», «05/12/03», «5 12 3», «Р5 С12 Я3», «R5-S12-L3»;
' letters of an address stay (A-01-1 → A-1-1, not the same place as B-01-1); two parts → «ряд-секция» (no level); "" for one part
' or more than three (a word, a number alone)
Function PlaceKey(ByVal s As String) As String
    Dim i As Integer, ch As String, cur As String, n As Integer, parts(3) As String
    s = UCase(Trim(s))
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then ch = Mid(s, i, 1) Else ch = " "
        If ch = "-" Or ch = "." Or ch = "/" Or ch = " " Or ch = "," Or ch = ":" Or ch = "_" Or ch = "\" Then
            If cur <> "" Then
                If n > 2 Then
                    PlaceKey = ""
                    Exit Function
                End If
                parts(n) = PlacePart(cur)
                n = n + 1
                cur = ""
            End If
        Else
            cur = cur & ch
        End If
    Next i
    If n = 3 Then
        PlaceKey = parts(0) & "-" & parts(1) & "-" & parts(2)
    ElseIf n = 2 Then
        PlaceKey = parts(0) & "-" & parts(1)
    End If
End Function

' one part of a place: a number without leading zeros; «Р5», «С12», «Я3», «R5», «S12», «L3» (ряд, секция, ярус) → the number;
' anything else as written (upper case)
Private Function PlacePart(ByVal t As String) As String
    Dim k As Integer, lead As String, num As String
    For k = 1 To Len(t)
        If Mid(t, k, 1) >= "0" And Mid(t, k, 1) <= "9" Then Exit For
    Next k
    lead = Left(t, k - 1)
    num = Mid(t, k)
    If num <> "" And IsAllDigits(num) Then
        If lead = "" Or lead = "Р" Or lead = "С" Or lead = "Я" Or lead = "R" Or lead = "S" Or lead = "L" Then
            PlacePart = CStr(CLng(Right(num, 9)))
            Exit Function
        End If
    End If
    PlacePart = t
End Function

Private Function IsAllDigits(t As String) As Boolean
    Dim k As Integer
    For k = 1 To Len(t)
        If Mid(t, k, 1) < "0" Or Mid(t, k, 1) > "9" Then Exit Function
    Next k
    IsAllDigits = (Len(t) > 0)
End Function

' a bar of width cells «█» for x of mx (at least one block for x > 0)
Function Bar(x As Double, mx As Double, w As Integer) As String
    Dim k As Integer
    If x <= 0 Or mx <= 0 Then Exit Function
    k = CInt(x / mx * w)
    If k < 1 Then k = 1
    If k > w Then k = w
    Bar = String(k, "█")
End Function

' the order of idx(0 … n-1) by keys(idx(i)), ascending (or descending): a shell sort, stable enough for the views
Sub SortIdx(keys() As Double, idx() As Long, n As Long, bDesc As Boolean)
    Dim gap As Long, i As Long, j As Long, t As Long, a As Double, b As Double
    gap = n \ 2
    Do While gap > 0
        For i = gap To n - 1
            t = idx(i)
            j = i
            Do While j >= gap
                a = keys(idx(j - gap))
                b = keys(t)
                If bDesc Then
                    If a >= b Then Exit Do
                Else
                    If a <= b Then Exit Do
                End If
                idx(j) = idx(j - gap)
                j = j - gap
            Loop
            idx(j) = t
        Next i
        gap = gap \ 2
    Loop
End Sub

' the whole number with thin spaces between thousands (1 234 567)
Function NumText(x As Double) As String
    Dim s As String, r As String, i As Integer, neg As Boolean
    If Abs(x - Int(x)) > 0.0000005 Then
        NumText = WmsIssue.QtyText(x)
        Exit Function
    End If
    neg = (x < 0)
    s = CStr(CLng(Abs(x)))
    For i = Len(s) To 1 Step -1
        r = Mid(s, i, 1) & r
        If (Len(s) - i) Mod 3 = 2 And i > 1 Then r = " " & r
    Next i
    NumText = IIf(neg, "−", "") & r
End Function

' a word with the number: 1 день, 2 дня, 5 дней
Function Plural(n As Long, one As String, few As String, many As String) As String
    Dim m As Long
    m = Abs(n) Mod 100
    If m >= 11 And m <= 14 Then
        Plural = many
    ElseIf m Mod 10 = 1 Then
        Plural = one
    ElseIf m Mod 10 >= 2 And m Mod 10 <= 4 Then
        Plural = few
    Else
        Plural = many
    End If
End Function

' the name of an operation of the journal for a person
Function OpName(sType As String) As String
    Select Case sType
    Case "ISSUE": OpName = "Выдача"
    Case "ISSUE_FIX": OpName = "Исправлена выдача"
    Case "ISSUE_DEL": OpName = "Выдача удалена"
    Case "RECEIPT": OpName = "Приход по заказу"
    Case "RECEIPT_ADD": OpName = "Ещё поступление"
    Case "RECEIPT_FIX": OpName = "Исправлен приход"
    Case "RECEIPT_DEL": OpName = "Приход удалён"
    Case "ORDER_CANCEL": OpName = "Заказ отменён"
    Case "ORDER_CANCEL_REST": OpName = "Остаток заказа отменён"
    Case "ORDER_UNCANCEL": OpName = "Отмена снята"
    Case "POSITION_FIX": OpName = "Исправлена позиция"
    Case "RETURN": OpName = "Возврат"
    Case "RETURN_FIX": OpName = "Исправлен возврат"
    Case "RETURN_DEL": OpName = "Возврат удалён"
    Case "MOVE": OpName = "Перемещение"
    Case "WRITE_OFF": OpName = "Списание"
    Case "INV_ADJ": OpName = "Инвентаризация"
    Case "MOVE_FIX", "WRITE_OFF_FIX", "INV_ADJ_FIX": OpName = "Исправлена корректировка"
    Case "MOVE_DEL", "WRITE_OFF_DEL", "INV_ADJ_DEL": OpName = "Корректировка удалена"
    Case "CAR_ARRIVE": OpName = "Машина приехала"
    Case "CAR_DEPART": OpName = "Машина уехала"
    Case "CAR_FIX": OpName = "Исправлена машина"
    Case "CAR_CANCEL": OpName = "Визит машины отменён"
    Case "MIGRATE": OpName = "Перенос ЕИ"
    Case "MIGRATE_FIX": OpName = "Исправлен ЕИ переноса"
    Case "LEGACY_RECEIPT", "LEGACY_RECEIPT_ADD": OpName = "Перенос прихода"
    Case "LEGACY_FIX": OpName = "Коррекция переноса"
    Case "LEGACY_BEGIN": OpName = "Начало переноса"
    Case "LEGACY_END": OpName = "Конец переноса"
    Case "SP_FIX": OpName = "Исправлен иной приход"
    Case "SP_DEL": OpName = "Иной приход удалён"
    Case "SP_IDENTIFY": OpName = "Разобран иной приход"
    Case Else
        If Left(sType, 3) = "SP_" Then
            OpName = "Иной приход"
        Else
            OpName = sType
        End If
    End Select
End Function

' a field «KEY=value» of a parsed journal line (the value unescaped), "" when absent
Function JField(fields As Variant, sKey As String) As String
    Dim i As Long
    For i = 0 To UBound(fields)
        If Left(fields(i), Len(sKey) + 1) = sKey & "=" Then
            JField = WmsCore.Unesc(Mid(fields(i), Len(sKey) + 2))
            Exit Function
        End If
    Next i
End Function

' go to a sheet and select a cell (the view sheets lead to the rows of the data sheets)
Sub GoCell(sName As String, r As Long, c As Integer)
    Dim sh As Object
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(sName) Then Exit Sub
    sh = gDoc.Sheets.getByName(sName)
    gDoc.getCurrentController().setActiveSheet(sh)
    gDoc.getCurrentController().select(sh.getCellByPosition(c, r))
    Exit Sub
EH:
End Sub

' the cell under the cursor of a sheet: its row (-1 when the selection is not on sName)
Function ViewRow(sName As String) As Long
    Dim sel As Object
    ViewRow = -1
    On Error GoTo EH
    If gDoc.getCurrentController().getActiveSheet().Name <> sName Then Exit Function
    sel = gDoc.getCurrentController().getSelection()
    ViewRow = sel.getRangeAddress().StartRow
    Exit Function
EH:
End Function

Function ViewCol(sName As String) As Long
    Dim sel As Object
    ViewCol = -1
    On Error GoTo EH
    If gDoc.getCurrentController().getActiveSheet().Name <> sName Then Exit Function
    sel = gDoc.getCurrentController().getSelection()
    ViewCol = sel.getRangeAddress().StartColumn
    Exit Function
EH:
End Function

' a Collection used as a dictionary of numbers: the item of key, -1 when absent. (Its own error handler: an «On Error
' Resume Next» in the caller would stay in force for the rest of the caller in LibreOffice Basic.)
Function CollGet(c As Object, key As String) As Long
    On Error GoTo Absent
    CollGet = c.Item(key)
    Exit Function
Absent:
    CollGet = -1
End Function

' adds v under key; False when the key is already there
Function CollPut(c As Object, v As Long, key As String) As Boolean
    On Error GoTo Taken
    c.Add(v, key)
    CollPut = True
    Exit Function
Taken:
    CollPut = False
End Function

' ================================================================ formulas of the Calc engine over the sheets (spec §25)
' The pages count the sheets of movements by array formulas evaluated at once in the scratch cell of _IDX
' (WmsOrders.ScratchEval): the engine filters and counts 100 000 rows in tens of milliseconds, a Basic loop takes
' 30–300 µs a row. Only what the formula found (a few rows, a list of numbers) is read by Basic. Inside IF the branches
' are plain references or arithmetic: an IF over whole columns with a lookup in its branch is evaluated element by element.

Global gVwEngFail As String
Global gVwEngBody As String

' the absolute address of column c (0-based) of sheet sName, rows r0 … r1 (0-based)
Function ColRng(sName As String, c As Integer, r0 As Long, r1 As Long) As String
    Dim L As String
    L = ColLetters(c)
    ColRng = "$'" & sName & "'.$" & L & "$" & (r0 + 1) & ":$" & L & "$" & (r1 + 1)
End Function

Function ColLetters(ByVal c As Long) As String
    Dim s As String
    c = c + 1
    Do While c > 0
        s = Chr(65 + (c - 1) Mod 26) & s
        c = (c - 1) \ 26
    Loop
    ColLetters = s
End Function

' «the row was posted by WMS» for the control column rng (Проведено, Проведено (исправлено))
Function PostedExpr(rng As String) As String
    PostedExpr = "(LEFT(" & rng & ";" & Len(ST_POSTED) & ")=""" & ST_POSTED & """)"
End Function

' a date column rng within the days [d0, d1] (whole days; a text or an empty cell is not a date)
Function DaysExpr(rng As String, d0 As Double, d1 As Double) As String
    DaysExpr = "ISNUMBER(" & rng & ")*(" & rng & ">=" & Format(d0, "0") & ")*(" & rng & "<" & Format(d1 + 1, "0") & ")"
End Function

' the text result of a formula; "" and gVwEngFail set when the engine could not be used
Function Eng(body As String) As String
    Dim ok As Boolean, s As String
    s = WmsOrders.ScratchEval(body, ok)
    If Not ok Then
        If gVwEngFail = "" Then
            gVwEngFail = "формула Calc не вычислена (лист " & SH_IDX & ")"
            gVwEngBody = Left(body, 400)
        End If
        Exit Function
    End If
    Eng = s
End Function

' the numbers of a TEXTJOIN list as a Split array; nItems — how many ("" → 0)
Function EngList(body As String, ByRef nItems As Long) As Variant
    Dim s As String
    s = Eng(body)
    If s = "" Then
        nItems = 0
        EngList = Array()
    Else
        EngList = Split(s, ";")
        nItems = UBound(EngList) + 1
    End If
End Function

' the field i of a list of numbers «a;b;c» (0 when absent)
Function ListNum(a As Variant, i As Long) As Double
    If i <= UBound(a) Then
        If a(i) <> "" Then ListNum = CDbl(a(i))
    End If
End Function

' the number of a canonical EI text («ЕИ-00000427» — what WMS writes into a posted row), otherwise EiNo
Function EiFast(e As String, nextEI As Long) As Long
    Dim n As Long
    If Len(e) = EI_DIGITS + 3 And Left(e, 3) = EI_PREFIX Then
        n = Val(Mid(e, 4))
        If n >= 1 And n < nextEI Then EiFast = n
    Else
        EiFast = EiNo(e, nextEI)
    End If
End Function

' The sheets of movements: name | EI column | date column | control column (-1: every row with an EI is a movement — the
' receipts of «Заказы») — the same as the card of an EI reads them
Function MoveSpecs() As Variant
    MoveSpecs = Array(Array(SH_ORDERS, OC_EI, OC_RDATE, -1), Array(SH_SPECIAL, XC_EI, XC_DATE, XC_STATUS), Array(SH_ISSUES, IC_EI, IC_DATE, IC_CTL), _
        Array(SH_RETURNS, RC_EI, RC_DATE, RC_CTL), Array(SH_ADJUST, AC_EI, AC_DATE, AC_CTL))
End Function

' the movement condition of a spec over its data rows 1 … last
Function MoveCond(sp As Variant, last As Long) As String
    If sp(3) < 0 Then
        MoveCond = "(" & ColRng(sp(0), sp(1), 1, last) & "<>"""")"
    Else
        MoveCond = PostedExpr(ColRng(sp(0), sp(3), 1, last))
    End If
End Function

' recent(n) = True for every EI that moved on day dFrom or later (the EI texts of those rows, one formula a sheet)
Sub MarkRecent(recent() As Boolean, nextEI As Long, dFrom As Double)
    Dim specs As Variant, i As Integer, sh As Object, last As Long, a As Variant, k As Long, nItems As Long, n As Long, dR As String, e As String
    Dim pfx As String, ln As Integer
    ' the constants of WmsConfig copied once: a name of another module costs 15–25 µs on each use inside a loop
    pfx = EI_PREFIX
    ln = EI_DIGITS + 3
    specs = MoveSpecs()
    For i = 0 To UBound(specs)
        If gDoc.Sheets.hasByName(specs(i)(0)) Then
            sh = gDoc.Sheets.getByName(specs(i)(0))
            last = WmsOrders.LastRow(sh)
            If last >= 1 Then
                dR = ColRng(specs(i)(0), specs(i)(2), 1, last)
                a = EngList("TEXTJOIN("";"";1;IF(" & MoveCond(specs(i), last) & "*ISNUMBER(" & dR & ")*(" & dR & ">=" & Format(dFrom, "0") & ");" _
                    & ColRng(specs(i)(0), specs(i)(1), 1, last) & ";""""))", nItems)
                ' a posted row holds the canonical EI (ЕИ-00000427): its number is read directly
                For k = 0 To nItems - 1
                    e = a(k)
                    If Len(e) = ln And Left(e, 3) = pfx Then n = Val(Mid(e, 4)) Else n = EiNo(e, nextEI)
                    If n > 0 And n < nextEI Then recent(n) = True
                Next k
            End If
        End If
    Next i
End Sub

' the day of the last movement of each EI of eis(0 … n − 1) on the sheets (0 — no movement on the sheets): one formula a
' sheet finds the movement rows of these EIs, Basic takes the latest day of each
Function LastMoveDays(eis() As Long, n As Long) As Variant
    Dim specs As Variant, i As Integer, j As Long, sh As Object, last As Long, f As String, eR As String, out() As Double, a As Variant, nItems As Long
    Dim d As Variant, k As Long, m As Long, dt As Double, canon() As String
    ReDim out(IIf(n > 0, n - 1, 0))
    LastMoveDays = out
    If n = 0 Then Exit Function
    ReDim canon(n - 1)
    For j = 0 To n - 1
        canon(j) = WmsIssue.EiCanon(eis(j))
    Next j
    specs = MoveSpecs()
    For i = 0 To UBound(specs)
        If gDoc.Sheets.hasByName(specs(i)(0)) Then
            sh = gDoc.Sheets.getByName(specs(i)(0))
            last = WmsOrders.LastRow(sh)
            If last >= 1 Then
                eR = ColRng(specs(i)(0), specs(i)(1), 1, last)
                f = ""
                For j = 0 To n - 1
                    f = f & IIf(f <> "", "+", "") & "(" & eR & "=""" & canon(j) & """)"
                Next j
                a = EngList("TEXTJOIN("";"";1;IF(" & MoveCond(specs(i), last) & "*(" & f & ");ROW(" & eR & ")-1;""""))", nItems)
                m = specs(i)(1)
                If specs(i)(2) > m Then m = specs(i)(2)
                d = RowsData(sh, a, nItems, m)
                For k = 0 To nItems - 1
                    If VarType(d(k)(specs(i)(2))) = 5 Then
                        dt = Int(d(k)(specs(i)(2)))
                        For j = 0 To n - 1
                            If CStr(d(k)(specs(i)(1))) = canon(j) Then
                                If dt > out(j) Then out(j) = dt
                            End If
                        Next j
                    End If
                Next k
            End If
        End If
    Next i
    LastMoveDays = out
End Function

' x into the kMax smallest keys kept sorted in keys(0 … n − 1)
Sub TopPush(keys() As Double, ByRef n As Long, kMax As Long, x As Double)
    Dim i As Long
    If n >= kMax Then
        If x >= keys(n - 1) Then Exit Sub
        i = n - 1
    Else
        i = n
        n = n + 1
    End If
    Do While i > 0
        If keys(i - 1) <= x Then Exit Do
        keys(i) = keys(i - 1)
        i = i - 1
    Loop
    keys(i) = x
End Sub

' the values of column c of the rows rows(0 … n − 1) (numbers of the data rows, ascending) of sheet sh: read as one block
' when they lie close together, one by one otherwise
Function RowsCol(sh As Object, rows As Variant, n As Long, c As Integer) As Variant
    Dim out() As Variant, i As Long, r0 As Long, r1 As Long, d As Variant, k As Long, b0 As Long, b1 As Long
    If n = 0 Then
        RowsCol = Array()
        Exit Function
    End If
    ReDim out(n - 1)
    r0 = CLng(rows(0))
    r1 = CLng(rows(n - 1))
    If r1 - r0 + 1 <= 20 * n + 200 Then
        k = 0
        For b0 = r0 To r1 Step CHECK_CHUNK_ROWS
            b1 = b0 + CHECK_CHUNK_ROWS - 1
            If b1 > r1 Then b1 = r1
            d = sh.getCellRangeByPosition(c, b0, c, b1).getDataArray()
            Do While k < n
                i = CLng(rows(k))
                If i > b1 Then Exit Do
                out(k) = d(i - b0)(0)
                k = k + 1
            Loop
        Next b0
    Else
        For k = 0 To n - 1
            i = CLng(rows(k))
            out(k) = sh.getCellRangeByPosition(c, i, c, i).getDataArray()(0)(0)
        Next k
    End If
    RowsCol = out
End Function

' the data of the rows rows(0 … n − 1) (numbers of the data rows, ascending) of sheet sh, columns 0 … c1: read as one block
' when they lie close together, one by one otherwise
Function RowsData(sh As Object, rows As Variant, n As Long, c1 As Integer) As Variant
    Dim out() As Variant, i As Long, r0 As Long, r1 As Long, d As Variant, k As Long, b0 As Long, b1 As Long
    If n = 0 Then
        RowsData = Array()
        Exit Function
    End If
    ReDim out(n - 1)
    r0 = CLng(rows(0))
    r1 = CLng(rows(n - 1))
    If r1 - r0 + 1 <= 4 * n + 200 Then
        k = 0
        For b0 = r0 To r1 Step CHECK_CHUNK_ROWS
            b1 = b0 + CHECK_CHUNK_ROWS - 1
            If b1 > r1 Then b1 = r1
            d = sh.getCellRangeByPosition(0, b0, c1, b1).getDataArray()
            Do While k < n
                i = CLng(rows(k))
                If i > b1 Then Exit Do
                out(k) = d(i - b0)
                k = k + 1
            Loop
        Next b0
    Else
        For k = 0 To n - 1
            i = CLng(rows(k))
            out(k) = sh.getCellRangeByPosition(0, i, c1, i).getDataArray()(0)
        Next k
    End If
    RowsData = out
End Function

' the start of the formula that could not be evaluated (diagnostics of the tests)
Function TestEngLast() As String
    TestEngLast = gVwEngBody
End Function
