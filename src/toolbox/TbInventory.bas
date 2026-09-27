' TbInventory — WMS_INVENTORY: инвентаризация по снимку (задание «FINAL WMS MARATHON», §6).
' Снимок текущего склада → лист «Пересчёт» (ЕИ, наименование, артикул, единица, место, категория, учётный остаток) →
' кладовщик вносит фактический остаток → разница → пакет INV_ADJ (WMS-BATCH-1) для листа «Корректировки» WMS. Книгу WMS
' инструмент не открывает; учётный остаток снимка уходит в пакет (колонка book): если к проведению остаток ЕИ изменился,
' WMS эту строку не проведёт (нужен пересчёт).
Option Explicit

Private Const SH_COUNT = "Пересчёт"
Private Const FIRST_ROW = 2          ' 0-based: row 3 of the sheet (rows 1–2: the snapshot and the headers)

' «Обновить из снимка»: the rows of the count from the latest snapshot (filters of «Настройки»: B6 place starts with,
' B7 category); a count already started is kept unless confirmed
Sub BtnInvRefresh(Optional oEvent As Variant)
    TbMsg(InvRefresh(False))
End Sub

Function InvRefresh(bForce As Boolean) As String
    Dim why As String, snap As String, n As Long, d As Variant, sh As Object, i As Long, out() As Variant, k As Long, st As String
    Dim fPlace As String, fCat As String, last As Long, started As Long
    On Error GoTo EH
    snap = TbPickSnapshot(why)
    If snap = "" Then
        InvRefresh = "ERR:" & why
        Exit Function
    End If
    sh = ThisComponent.Sheets.getByName(SH_COUNT)
    last = TbLastRow(sh)
    If last >= FIRST_ROW And Not bForce Then
        started = CountFilled(sh, 7, FIRST_ROW, last)
        If started > 0 And Not gTbAuto Then
            If MsgBox("Пересчёт уже начат (фактических остатков: " & started & "). Заменить его строками нового снимка?", 4 + 32, "WMS_INVENTORY") <> 6 Then
                InvRefresh = "SKIP:пересчёт оставлен"
                Exit Function
            End If
        End If
    End If
    n = TbLoadTable(snap, "stock.csv", "_stock", True)
    fPlace = LCase(Trim(TbSetting(5)))
    fCat = LCase(Trim(TbSetting(6)))
    If n > 0 Then d = ThisComponent.Sheets.getByName("_stock").getCellRangeByPosition(0, 1, 9, n).getDataArray()
    ReDim out(IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        st = CStr(d(i)(7))
        If CStr(d(i)(0)) <> "" And st <> "Приход удалён (сторно)" Then
            If (fPlace = "" Or Left(LCase(CStr(d(i)(5))), Len(fPlace)) = fPlace) And (fCat = "" Or LCase(CStr(d(i)(6))) = fCat) Then
                out(k) = Array(d(i)(0), d(i)(1), d(i)(2), d(i)(3), d(i)(5), d(i)(6), d(i)(4))
                k = k + 1
            End If
        End If
    Next i
    If last >= FIRST_ROW Then sh.getCellRangeByPosition(0, FIRST_ROW, 9, last).clearContents(1023)
    If k > 0 Then
        Dim rows() As Variant, j As Long
        ReDim rows(k - 1)
        For j = 0 To k - 1
            rows(j) = out(j)
        Next j
        sh.getCellRangeByPosition(0, FIRST_ROW, 6, FIRST_ROW + k - 1).setDataArray(rows)
        For j = 0 To k - 1
            sh.getCellByPosition(8, FIRST_ROW + j).setFormula("=IF(H" & (FIRST_ROW + j + 1) & "="""";"""";H" & (FIRST_ROW + j + 1) & "-G" & (FIRST_ROW + j + 1) & ")")
        Next j
    End If
    sh.getCellByPosition(0, 0).setString("Снимок: " & Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1) & " (LAST_SEQ " _
        & TbManifestValue(snap, "last_seq") & ", " & TbManifestValue(snap, "created") & ")")
    TbPutSetting(3, snap)
    InvRefresh = "OK:строк пересчёта " & k & " из снимка " & ConvertFromURL(snap)
    Exit Function
EH:
    InvRefresh = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

Private Function CountFilled(sh As Object, col As Integer, r0 As Long, r1 As Long) As Long
    Dim d As Variant, i As Long
    d = sh.getCellRangeByPosition(col, r0, col, r1).getDataArray()
    For i = 0 To UBound(d)
        If CStr(d(i)(0)) <> "" Then CountFilled = CountFilled + 1
    Next i
End Function

' «Сформировать пакет»: the rows with a fact (H) that differs from the book balance (G) → one INV_ADJ batch
Sub BtnInvBatch(Optional oEvent As Variant)
    TbMsg(InvBatch(""))
End Sub

Function InvBatch(sDate As String) As String
    Dim sh As Object, last As Long, d As Variant, i As Long, rows() As Variant, k As Long, bad As String, nBad As Long, id As String, path As String
    Dim snap As String, dt As String, x As Double
    On Error GoTo EH
    sh = ThisComponent.Sheets.getByName(SH_COUNT)
    last = TbLastRow(sh)
    snap = TbSetting(3)
    If last < FIRST_ROW Or snap = "" Then
        InvBatch = "ERR:пересчёта нет — сначала «Обновить из снимка»"
        Exit Function
    End If
    dt = IIf(sDate <> "", sDate, Format(Now(), "DD.MM.YYYY"))
    d = sh.getCellRangeByPosition(0, FIRST_ROW, 9, last).getDataArray()
    ReDim rows(UBound(d))
    For i = 0 To UBound(d)
        If CStr(d(i)(7)) <> "" Then
            If VarType(d(i)(7)) <> 5 Then
                nBad = nBad + 1
                If bad = "" Then bad = "строка " & (FIRST_ROW + i + 1) & ": фактический остаток «" & d(i)(7) & "» не число"
            ElseIf d(i)(7) < 0 Then
                nBad = nBad + 1
                If bad = "" Then bad = "строка " & (FIRST_ROW + i + 1) & ": отрицательный фактический остаток"
            ElseIf Abs(d(i)(7) - d(i)(6)) > 0.0000001 Then
                rows(k) = Array("Инвентаризация", d(i)(0), NumText(d(i)(7)), NumText(d(i)(6)), dt, _
                    "Инвентаризация " & dt & IIf(CStr(d(i)(9)) <> "", ": " & d(i)(9), ""), "")
                k = k + 1
            End If
        End If
    Next i
    If nBad > 0 Then
        InvBatch = "ERR:ошибок во вводе: " & nBad & " (" & bad & ") — пакет не создан"
        Exit Function
    End If
    If k = 0 Then
        InvBatch = "SKIP:расхождений нет — пакет не нужен"
        Exit Function
    End If
    Dim outRows() As Variant
    ReDim outRows(k - 1)
    For i = 0 To k - 1
        outRows(i) = rows(i)
    Next i
    id = "INV-" & Format(Now(), "YYYYMMDD-HHMMSS")
    path = TbWriteBatch("Корректировки", id, "WMS_INVENTORY", Mid(snap, TbInStrRev(Left(snap, Len(snap) - 1), "/") + 1, Len(snap) - TbInStrRev(Left(snap, Len(snap) - 1), "/") - 1), _
        Array("kind", "ei", "fact", "book", "date", "reason", "note"), outRows, "")
    InvBatch = "OK:пакет " & id & ": расхождений " & k & " — " & ConvertFromURL(path) & ". В WMS: «Главная» → «Загрузить пакет»"
    Exit Function
EH:
    InvBatch = "ERR:внутренняя ошибка инструмента: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function

' a quantity as WMS parses it: comma, no trailing zeros
Function NumText(x As Variant) As String
    Dim s As String
    ' three decimals at most (the quantity rule of WMS); Round of Basic is avoided (it is not the same in every version)
    s = Trim(Str(Int(CDbl(x) * 1000 + 0.5) / 1000))
    NumText = Replace(s, ".", ",")
End Function
