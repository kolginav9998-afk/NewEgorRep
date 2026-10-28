' WmsStockView — 1.1: two sheets made from the registry «Наличие» for looking, searching and sorting. The registry itself is addressed by the row
' (row n — ЕИ-n) and is never sorted; these are copies that are written anew (when the sheet is opened after a posting, and by the button), so
' the storekeeper sorts and filters them as he likes — nothing can be spoilt. «Наличие по названию»: one row per EI, by default the ones with a balance,
' in the order of the name. «Остаток деталей»: one row per part (the article, or the name when there is none): the balance of all its EIs, how many
' EIs, the places, the serial numbers. A view posts nothing and writes nothing to the journal. Library WmsPlus.
Option Explicit

Private Const SH_BYNAME = "Наличие по названию"
Private Const SH_PARTS = "Остаток деталей"
Private Const SV_FIND_ROW = 1
Private Const SV_ZERO_COL = 5
Private Const SV_STAT_ROW = 2
Private Const SV_HEAD = 3
Private Const SV_ROW0 = 4
Private Const SV_MAX = 60000
Private Const SV_BYNAME_COLS = 11
Private Const SV_PARTS_COLS = 8

' the journal position of the last build of each view (a view is built again when the journal moved)
Global gSvSeqName As Long
Global gSvSeqPart As Long
Global gSvBusy As Boolean
Global gSvTiming As String

Private Function SvSeq() As Long
    SvSeq = CLng(SysNum(SK_LAST_SEQ))
    If SvSeq = 0 Then SvSeq = 1
End Function

Private Function ViewFind(sSheet As String) As String
    If gDoc.Sheets.hasByName(sSheet) Then ViewFind = Trim(gDoc.Sheets.getByName(sSheet).getCellByPosition(1, SV_FIND_ROW).getString())
End Function

' «да» in the cell next to «С нулевым остатком»
Private Function ViewZero(sSheet As String) As Boolean
    Dim t As String
    If gDoc.Sheets.hasByName(sSheet) Then
        t = LCase(Trim(gDoc.Sheets.getByName(sSheet).getCellByPosition(SV_ZERO_COL, SV_FIND_ROW).getString()))
        ViewZero = (t = "да" Or t = "yes" Or t = "1")
    End If
End Function

' the data of the registry rows 1 … last: columns A … «Резерв» (the information of «Наличие»)
Private Function RegistryData(ByRef nRows As Long) As Variant
    Dim st As Object, cur As Object, last As Long
    st = WmsIssue.StockSheet()
    cur = st.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    nRows = 0
    If last < 1 Then
        RegistryData = Array()
        Exit Function
    End If
    nRows = last
    RegistryData = st.getCellRangeByPosition(0, 1, SI_COL0 + SI_COLS - 1, last).getDataArray()
End Function

Private Function QtyOf(v As Variant) As Double
    If VarType(v) = 5 Then QtyOf = v
End Function

' the rows of the list (an array of arrays) written under the header of sheet sh, the old rows out; the sort by the columns c1, c2, c3 (the name, the
' article, the code) in the order of the Russian alphabet
Private Sub WriteRows(sh As Object, rowsOut As Variant, n As Long, nCols As Integer, c1 As Integer, c2 As Integer)
    Dim cur As Object, prevEnd As Long, e As Long, rng As Object, out() As Variant, i As Long
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    prevEnd = cur.getRangeAddress().EndRow
    e = prevEnd
    If SV_ROW0 + n - 1 > e Then e = SV_ROW0 + n - 1
    If e >= SV_ROW0 Then sh.getCellRangeByPosition(0, SV_ROW0, nCols - 1, e).clearContents(1 + 2 + 4 + 16)
    If n = 0 Then Exit Sub
    ReDim out(n - 1)
    For i = 0 To n - 1
        out(i) = rowsOut(i)
    Next i
    rng = sh.getCellRangeByPosition(0, SV_ROW0, nCols - 1, SV_ROW0 + n - 1)
    rng.setDataArray(out)
    SvSort(rng, c1, c2)
End Sub

Function TestTiming() As String
    TestTiming = gSvTiming
End Function

Private Sub SvSort(rng As Object, c1 As Integer, c2 As Integer)
    Dim f(1) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = c1
    f(0).IsAscending = True
    f(0).IsCaseSensitive = False
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    f(1).Field = c2
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

Private Sub PutStat(sh As Object, s As String)
    sh.getCellByPosition(0, SV_STAT_ROW).setString(s)
End Sub

' ================================================================ «Наличие по названию»

Sub BtnByNameRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = ByNameRefresh(ViewFind(SH_BYNAME), ViewZero(SH_BYNAME))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the sheet is opened: built again when the journal moved since it was built (the first time — when it is empty)
Sub OnByNameFocus(Optional oEvent As Variant)
    If gSvBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    If SvSeq() = gSvSeqName And gSvSeqName > 0 Then Exit Sub
    ByNameRefresh(ViewFind(SH_BYNAME), ViewZero(SH_BYNAME))
    Exit Sub
EH:
End Sub

Sub OnByNameChange(Optional oTarget As Variant)
    Dim ra As Object
    If gSvBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> SV_FIND_ROW Or (ra.Column <> 1 And ra.Column <> SV_ZERO_COL) Then Exit Sub
    WmsInit()
    ByNameRefresh(ViewFind(SH_BYNAME), ViewZero(SH_BYNAME))
    Exit Sub
EH:
End Sub

' the list: one row per EI (with a balance, or all of them when bZero) whose name, article, place, category, code or serial number has every word of sFind.
' OK:<rows> / ERR:…
Function ByNameRefresh(sFind As String, bZero As Boolean) As String
    Dim sh As Object, d As Variant, nRows As Long, i As Long, n As Long, q As Double, words As Variant, hay As String, nm As String, out() As Variant
    Dim ser As String, dt As Variant, ei As String, nAll As Long, t0 As Long, hays() As String, pick() As Long, nPick As Long, folded As Variant, bWords As Boolean, k As Long
    Dim cEi As Long, cNm As Long, cArt As Long, cUn As Long, cQ As Long, cPl As Long, cCat As Long, cSt As Long, cSrc As Long, cSer As Long, cDt As Long
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SH_BYNAME) Then
        ByNameRefresh = "ERR:в книге нет листа «" & SH_BYNAME & "» — книга не обновлена до 1.1"
        Exit Function
    End If
    gSvBusy = True
    WmsInit()
    t0 = GetSystemTicks()
    sh = gDoc.Sheets.getByName(SH_BYNAME)
    d = RegistryData(nRows)
    gSvTiming = "read " & (GetSystemTicks() - t0)
    ' the columns of the registry as local numbers: a constant of another library costs 30 µs at every use
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cUn = SC_UNIT
    cQ = SC_QTY
    cPl = SC_PLACE
    cCat = SC_CAT
    cSt = SC_STATE
    cSrc = SC_SRC
    cSer = SI_COL0 + 8
    cDt = SI_COL0 + 4
    words = WmsNames.FoldWords(sFind)
    bWords = (UBound(words) >= 0)
    If nRows > 0 Then
        ReDim out(UBound(d))
        ReDim hays(UBound(d))
        ReDim pick(UBound(d))
    End If
    ' first the rows that come in (a balance, or all of them), and for a search the text of each
    For i = 0 To nRows - 1
        If CStr(d(i)(cEi)) <> "" Then
            nAll = nAll + 1
            q = QtyOf(d(i)(cQ))
            If bZero Or q > 0.0000005 Then
                pick(nPick) = i
                If bWords Then hays(nPick) = d(i)(cEi) & " " & WmsNames.NormNameFast(CStr(d(i)(cNm))) & " " & d(i)(cArt) & " " & d(i)(cPl) & " " & d(i)(cCat) & " " & d(i)(cSer)
                nPick = nPick + 1
            End If
        End If
    Next i
    If bWords And nPick > 0 Then
        ReDim Preserve hays(nPick - 1)
        folded = WmsNames.FoldAll(hays)
    End If
    For k = 0 To nPick - 1
        If (Not bWords) Or WmsNames.HasWords(CStr(folded(k)), words) Then
            i = pick(k)
            dt = d(i)(cDt)
            If VarType(dt) <> 5 Then dt = ""
            If n < SV_MAX Then
                out(n) = Array(CStr(d(i)(cEi)), WmsNames.NormNameFast(CStr(d(i)(cNm))), CStr(d(i)(cArt)), CStr(d(i)(cUn)), QtyOf(d(i)(cQ)), CStr(d(i)(cPl)), CStr(d(i)(cCat)), CStr(d(i)(cSt)), _
                    CStr(d(i)(cSrc)), Trim(CStr(d(i)(cSer))), dt)
                n = n + 1
            End If
        End If
    Next k
    gSvTiming = gSvTiming & "; loop " & (GetSystemTicks() - t0)
    WriteRows(sh, out, n, SV_BYNAME_COLS, 1, 2)
    gSvTiming = gSvTiming & "; write+sort " & (GetSystemTicks() - t0)
    PutStat(sh, "ЕИ в списке: " & n & " из " & nAll & IIf(Trim(sFind) <> "", " — по словам «" & Trim(sFind) & "»", "") & IIf(bZero, " (с нулевым остатком)", " (только с остатком)") _
        & ". Построено " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)) & ". Это копия «Наличия»: сортируйте и фильтруйте как удобно.")
    gSvSeqName = SvSeq()
    gSvBusy = False
    ByNameRefresh = "OK:" & n
    gUiLastMsg = ByNameRefresh
    Exit Function
EH:
    gSvBusy = False
    ByNameRefresh = "ERR:список «" & SH_BYNAME & "» не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = ByNameRefresh
End Function

' ================================================================ «Остаток деталей»

Sub BtnPartsRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = PartsRefresh(ViewFind(SH_PARTS), ViewZero(SH_PARTS))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub OnPartsFocus(Optional oEvent As Variant)
    If gSvBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    If SvSeq() = gSvSeqPart And gSvSeqPart > 0 Then Exit Sub
    PartsRefresh(ViewFind(SH_PARTS), ViewZero(SH_PARTS))
    Exit Sub
EH:
End Sub

Sub OnPartsChange(Optional oTarget As Variant)
    Dim ra As Object
    If gSvBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> SV_FIND_ROW Or (ra.Column <> 1 And ra.Column <> SV_ZERO_COL) Then Exit Sub
    WmsInit()
    PartsRefresh(ViewFind(SH_PARTS), ViewZero(SH_PARTS))
    Exit Sub
EH:
End Sub

' a part: its category is «Детали», or it came as a part («Детали DET-…» in the source)
Private Function IsPart(cat As String, src As String) As Boolean
    IsPart = (LCase(Trim(cat)) = "детали" Or LCase(Left(Trim(src), 6)) = "детали")
End Function

' the list: one row per part — the article (the name when there is none) — with the balance of all its EIs; OK:<rows> / ERR:…
Function PartsRefresh(sFind As String, bZero As Boolean) As String
    Dim sh As Object, d As Variant, nRows As Long, i As Long, j As Long, n As Long, q As Double, words As Variant, hay As String, out() As Variant
    Dim keys() As String, nm() As String, art() As String, un() As String, bal() As Double, cnt() As Long, pl() As String, sn() As String, ns() As Long, last() As Double
    Dim key As String, k As Long, place As String, ser As String, dt As Variant, nOut As Long, sSer As String, idx As Object, hays() As String, folded As Variant, bWords As Boolean
    Dim cEi As Long, cNm As Long, cArt As Long, cUn As Long, cQ As Long, cPl As Long, cCat As Long, cSt As Long, cSrc As Long, cSer As Long, cDt As Long
    On Error GoTo EH
    idx = New Collection
    If Not gDoc.Sheets.hasByName(SH_PARTS) Then
        PartsRefresh = "ERR:в книге нет листа «" & SH_PARTS & "» — книга не обновлена до 1.1"
        Exit Function
    End If
    gSvBusy = True
    WmsInit()
    sh = gDoc.Sheets.getByName(SH_PARTS)
    d = RegistryData(nRows)
    ' the columns of the registry as local numbers: a constant of another library costs 30 µs at every use
    cEi = SC_EI
    cNm = SC_NAME
    cArt = SC_ART
    cUn = SC_UNIT
    cQ = SC_QTY
    cPl = SC_PLACE
    cCat = SC_CAT
    cSt = SC_STATE
    cSrc = SC_SRC
    cSer = SI_COL0 + 8
    cDt = SI_COL0 + 4
    words = WmsNames.FoldWords(sFind)
    If nRows > 0 Then
        ReDim keys(UBound(d))
        ReDim nm(UBound(d))
        ReDim art(UBound(d))
        ReDim un(UBound(d))
        ReDim bal(UBound(d))
        ReDim cnt(UBound(d))
        ReDim pl(UBound(d))
        ReDim sn(UBound(d))
        ReDim ns(UBound(d))
        ReDim last(UBound(d))
    End If
    For i = 0 To nRows - 1
        If CStr(d(i)(cEi)) <> "" Then
            If IsPart(CStr(d(i)(cCat)), CStr(d(i)(cSrc))) Then
                key = LCase(Trim(CStr(d(i)(cArt))))
                If key = "" Then key = "n:" & LCase(Trim(CStr(d(i)(cNm)))) Else key = "a:" & key
                k = WmsView.CollGet(idx, key)
                If k < 0 Then
                    k = n
                    WmsView.CollPut(idx, k, key)
                    keys(k) = key
                    nm(k) = WmsNames.NormNameFast(CStr(d(i)(cNm)))
                    art(k) = Trim(CStr(d(i)(cArt)))
                    un(k) = CStr(d(i)(cUn))
                    n = n + 1
                End If
                q = QtyOf(d(i)(cQ))
                bal(k) = bal(k) + q
                If q > 0.0000005 Then
                    cnt(k) = cnt(k) + 1
                    place = Trim(CStr(d(i)(cPl)))
                    If place <> "" Then
                        If InStr("; " & pl(k) & ";", "; " & place & ";") = 0 Then pl(k) = pl(k) & IIf(pl(k) <> "", "; ", "") & place
                    End If
                    ser = Trim(CStr(d(i)(cSer)))
                    If ser <> "" Then
                        ns(k) = ns(k) + 1
                        If ns(k) <= 12 Then sn(k) = sn(k) & IIf(sn(k) <> "", "; ", "") & ser
                    End If
                End If
                dt = d(i)(cDt)
                If VarType(dt) = 5 Then
                    If dt > last(k) Then last(k) = dt
                End If
            End If
        End If
    Next i
    If n > 0 Then
        ReDim out(n - 1)
        ReDim hays(n - 1)
    End If
    bWords = (UBound(words) >= 0)
    If bWords And n > 0 Then
        For k = 0 To n - 1
            hays(k) = art(k) & " " & nm(k) & " " & pl(k) & " " & sn(k)
        Next k
        folded = WmsNames.FoldAll(hays)
    End If
    For k = 0 To n - 1
        If bZero Or bal(k) > 0.0000005 Then
            If (Not bWords) Or WmsNames.HasWords(CStr(folded(k)), words) Then
                sSer = sn(k) & IIf(ns(k) > 12, " … и ещё " & (ns(k) - 12), "")
                out(nOut) = Array(art(k), nm(k), bal(k), un(k), cnt(k), pl(k), sSer, IIf(last(k) > 0, last(k), ""))
                nOut = nOut + 1
            End If
        End If
    Next k
    WriteRows(sh, out, nOut, SV_PARTS_COLS, 1, 0)
    PutStat(sh, "Деталей (по артикулу) в списке: " & nOut & " из " & n & IIf(Trim(sFind) <> "", " — по словам «" & Trim(sFind) & "»", "") & IIf(bZero, " (и без остатка)", " (только с остатком)") _
        & ". Построено " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)) & ". Деталь — категория «Детали» или приход как детали.")
    gSvSeqPart = SvSeq()
    gSvBusy = False
    PartsRefresh = "OK:" & nOut
    gUiLastMsg = PartsRefresh
    Exit Function
EH:
    gSvBusy = False
    PartsRefresh = "ERR:список «" & SH_PARTS & "» не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = PartsRefresh
End Function
