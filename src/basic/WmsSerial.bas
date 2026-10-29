' WmsSerial — 1.1.1: parts with a serial number. (1) «Серийные детали…» on «Иной приход» and «Заказы»: the receipt of the parts and their serial numbers in one action — the
' storekeeper fills the row (article, name, place, document …), presses the button and types or pastes the serial numbers, one to a line; the quantity is the number of the numbers,
' the row is posted by its own button's macro, and the EI that came is taken apart by the numbers (WmsParts.SerialSplit: every part with a number is an EI of its own with the same
' article, its own place, number and label «Арт.: … · S/N: …»). No new operation: the receipt and the chain of «Разделение» and «Свойства ЕИ» that the buttons «Провести» and
' «Серийные…» already make. (2) The sheet «Серийные детали» — the directory of such parts, one row per serial number: article, name, EI, place, balance, state, to whom and when it was
' given out, where from; a search by any word; the labels of the selected rows. The sheet only shows (built from «Наличие», its properties and «Выдачи»), it posts nothing.
' Library WmsPlus.
Option Explicit

Private Const SZ_SHEET = "Серийные детали"
Private Const SZ_FIND_ROW = 1
Private Const SZ_ALL_COL = 5
Private Const SZ_STAT_ROW = 2
Private Const SZ_ROW0 = 4
Private Const SZ_COLS = 10
Private Const SZ_MAX = 200
Private Const SZ_PROP_COLS = 3
Private Const SZ_AGAIN = "ранее № "

' the text typed last (kept when the receipt was refused: the window opens with it again), the serial number a base EI lost, the journal position of the last build of the sheet
Global gSzText As String
Global gSzReplaced As String
Global gSzSeq As Long
Global gSzBusy As Boolean

' ================================================================ the serial numbers of a text

' the numbers of the text sText (one to a line; a comma, «;» and a Tab separate too — a column pasted from a table): True and the array sers, or False and why
Function SerialsParse(ByVal sText As String, ByRef sers As Variant, ByRef why As String) As Boolean
    Dim a As Variant, i As Long, j As Long, n As Long, s As String, out() As String, other As Long
    sText = Replace(Replace(Replace(Replace(Replace(sText, Chr(13), Chr(10)), Chr(9), Chr(10)), ";", Chr(10)), ",", Chr(10)), Chr(160), " ")
    If Trim(Replace(sText, Chr(10), "")) = "" Then
        why = "не указан ни один серийный номер"
        Exit Function
    End If
    a = Split(sText, Chr(10))
    ReDim out(UBound(a))
    For i = 0 To UBound(a)
        s = Trim(CStr(a(i)))
        If s <> "" Then
            If Len(s) > 60 Then
                why = "серийный номер «" & Left(s, 20) & "…» длиннее 60 знаков"
                Exit Function
            End If
            For j = 0 To n - 1
                If LCase(out(j)) = LCase(s) Then
                    why = "серийный номер «" & s & "» записан дважды (строки " & (j + 1) & " и " & (n + 1) & " списка)"
                    Exit Function
                End If
            Next j
            out(n) = s
            n = n + 1
        End If
    Next i
    If n = 0 Then
        why = "не указан ни один серийный номер"
        Exit Function
    End If
    If n > SZ_MAX Then
        why = "номеров " & n & ", а за один раз не больше " & SZ_MAX & " — разделите список на части"
        Exit Function
    End If
    For j = 0 To n - 1
        other = WmsItem.SerialEI(out(j))
        If other > 0 Then
            why = "серийный номер «" & out(j) & "» уже записан у " & WmsIssue.EiCanon(other) & " — проверьте номер"
            Exit Function
        End If
    Next j
    ReDim Preserve out(n - 1)
    sers = out
    SerialsParse = True
End Function

' ================================================================ the receipt with the numbers

' the row r of «Заказы» (bOrders) or «Иной приход», unposted, with the article and the name filled, posted with the serial numbers sers: the quantity of the row becomes their
' number (on «Заказы» the quantity of the fact F may be bigger — the parts left stay in the EI without a number), the category «Детали» is put when it is empty. The result:
' OK:<how many>|ЕИ-…,ЕИ-… / PARTIAL:<EI>|<why the numbers are not written> (the receipt is posted) / any refusal of the posting as it is (ERR:…, SKIP:…, BLOCKED:…) —
' then the cells that were changed are put back
Function ReceiveSerial(bOrders As Boolean, r As Long, sers As Variant) As String
    Dim sh As Object, qc As Long, ec As Long, cc As Long, ac As Long, nc As Long, n As Long, kind As String, qOld As String, qNow As Double, catOld As String
    Dim res As String, r2 As String, canon As String, nEI As Long, nextEI0 As Long, bal0 As Double, old As Variant, p As Variant, sNote As String, sExp As String
    WmsInit()
    On Error GoTo EH
    gSzReplaced = ""
    n = UBound(sers) + 1
    If bOrders Then
        sh = WmsOrders.OrdersSheet()
        qc = OC_FACT
        ec = OC_EI
        cc = OC_CAT
        ac = OC_ART
        nc = OC_NAME
        kind = WmsOrders.OrderRowKind(r)
    Else
        sh = WmsSpecial.SpecialSheet()
        qc = XC_QTY
        ec = XC_EI
        cc = XC_CAT
        ac = XC_ART
        nc = XC_NAME
        kind = WmsSpecial.SpecialRowKind(r)
    End If
    If kind <> "OPEN" Then
        ReceiveSerial = "ERR:строка не открыта для прихода (" & kind & ") — встаньте на непроведённую строку с данными детали"
        Exit Function
    End If
    If Trim(sh.getCellByPosition(ac, r).getString()) = "" Then
        ReceiveSerial = "ERR:впишите артикул детали (постоянный код) — по нему WMS узнаёт деталь"
        Exit Function
    End If
    If Trim(sh.getCellByPosition(nc, r).getString()) = "" Then
        ReceiveSerial = "ERR:впишите наименование детали"
        Exit Function
    End If
    qOld = sh.getCellByPosition(qc, r).getString()
    If qOld <> "" Then
        qNow = Val(Replace(qOld, ",", "."))
        If bOrders Then
            If qNow < n - 0.0000001 Then
                ReceiveSerial = "ERR:серийных номеров " & n & ", а по факту (F) принято только " & qOld & " — номеров не может быть больше деталей"
                Exit Function
            End If
        ElseIf Abs(qNow - n) > 0.0000001 Then
            If Not WmsUi.UiConfirm("В строке количество " & qOld & ", а серийных номеров " & n & ". Принять " & n & "?") Then
                ReceiveSerial = "SKIP:количество не согласовано с числом номеров — строка не проведена"
                Exit Function
            End If
        End If
    End If
    catOld = sh.getCellByPosition(cc, r).getString()
    If qOld = "" Or Not bOrders Then sh.getCellByPosition(qc, r).setString(CStr(n))
    If catOld = "" Then sh.getCellByPosition(cc, r).setString("Детали")
    nextEI0 = CLng(SysNum(SK_NEXT_EI))
    If bOrders Then
        res = WmsReceipt.ReceiptPostRow(r)
    Else
        res = WmsSpecial.SpecialPostRow(r, "", False)
    End If
    If Left(res, 3) <> "OK:" Then
        sh.getCellByPosition(qc, r).setString(qOld)
        sh.getCellByPosition(cc, r).setString(catOld)
        ReceiveSerial = res
        Exit Function
    End If
    canon = sh.getCellByPosition(ec, r).getString()
    nEI = WmsView.EiNo(canon, CLng(SysNum(SK_NEXT_EI)))
    If nEI < 1 Then
        ReceiveSerial = "PARTIAL:" & canon & "|приход проведён, но ЕИ строки не определён"
        Exit Function
    End If
    ' the article came to an EI that was there already (a receipt of «Детали» goes to the EI of the article): what it held before and the number it carried
    If Not bOrders And nEI < nextEI0 Then bal0 = Val(Replace(sh.getCellByPosition(XC_BEFORE, r).getString(), ",", "."))
    old = WmsItem.PropOf(nEI)
    r2 = WmsParts.SerialSplit(canon, sers)
    If Left(r2, 3) <> "OK:" Then
        ReceiveSerial = "PARTIAL:" & canon & "|" & Mid(r2, InStr(r2, ":") + 1)
        Exit Function
    End If
    ' the EI was empty and carried the number of a part given out: the last part stays in it with its own number — the old number is not lost, it goes to the note of the EI
    p = WmsItem.PropOf(nEI)
    If CStr(old(0)) <> "" And CStr(p(0)) <> CStr(old(0)) Then
        sNote = CStr(p(2))
        sNote = Left(sNote & IIf(sNote <> "", " · ", "") & SZ_AGAIN & CStr(old(0)), 200)
        sExp = IIf(p(1) > 0, Format(p(1), "DD.MM.YYYY"), "")
        WmsItem.PropSet(canon, CStr(p(0)), sExp, sNote)
        gSzReplaced = canon & "|" & CStr(old(0))
    End If
    ReceiveSerial = r2
    gUiLastMsg = ReceiveSerial
    Exit Function
EH:
    ReceiveSerial = "ERR-SYS:серийный приход не выполнен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = ReceiveSerial
End Function

Sub BtnSerialSpecial(Optional oEvent As Variant)
    SerialUi(False)
End Sub

Sub BtnSerialOrder(Optional oEvent As Variant)
    SerialUi(True)
End Sub

' the window of the button: the row under the cursor, the numbers (a big field), the count to be checked, the receipt, the labels
Private Sub SerialUi(bOrders As Boolean)
    Dim r As Long, why As String, ok As Boolean, sh As Object, ac As Long, nc As Long, pc As Long, info As String, sers As Variant, k As Integer, res As String
    Dim eis As String, a As Variant, msg As String, shown As Boolean
    WmsInit()
    On Error GoTo EH
    If bOrders Then
        ok = WmsOrdersUi.ActiveOrderRow(r, why)
        sh = WmsOrders.OrdersSheet()
        ac = OC_ART
        nc = OC_NAME
        pc = OC_PLACE
    Else
        ok = WmsSpecialUi.ActiveSpecialRow(r, why)
        sh = WmsSpecial.SpecialSheet()
        ac = XC_ART
        nc = XC_NAME
        pc = XC_PLACE
    End If
    If Not ok Then
        WmsUi.UiMessage(why)
        Exit Sub
    End If
    If Trim(sh.getCellByPosition(ac, r).getString()) = "" Or Trim(sh.getCellByPosition(nc, r).getString()) = "" Then
        WmsUi.UiMessage("Сначала заполните в строке артикул (постоянный код детали) и наименование, место и документ — потом нажмите «Серийные детали…». Количество WMS поставит сама по числу номеров.")
        Exit Sub
    End If
    info = "Деталь: " & Trim(sh.getCellByPosition(ac, r).getString()) & " · " & Trim(sh.getCellByPosition(nc, r).getString()) _
        & IIf(Trim(sh.getCellByPosition(pc, r).getString()) <> "", " · место " & Trim(sh.getCellByPosition(pc, r).getString()), "") & "." & Chr(10) _
        & "Впишите серийные номера, по одному на строку (можно вставить столбец из таблицы; запятая и «;» тоже разделяют). Деталей будет столько, сколько номеров: " _
        & "каждая станет своим ЕИ с тем же артикулом, своим номером и этикеткой."
    Do
        If Not WmsPrimeUi.UiMemo("Серийные детали — приход", info, gSzText, "Дальше") Then Exit Sub
        If SerialsParse(gSzText, sers, why) Then
            k = WmsPrimeUi.UiButtons("Серийные детали", "Деталей: " & (UBound(sers) + 1) & ". Первый номер: " & sers(0) & ", последний: " & sers(UBound(sers)) & "." & Chr(10) & Chr(10) _
                & "Приход будет проведён на " & (UBound(sers) + 1) & " шт и разделён на отдельные ЕИ — по одной детали с каждым номером.", _
                Array("Провести приход и записать номера", "Назад — поправить список"), Array(1, 0))
            If k = 1 Then Exit Do
            If k < 0 Then Exit Sub
        Else
            WmsUi.UiMessage("Список номеров: " & why & ".")
        End If
    Loop
    res = ReceiveSerial(bOrders, r, sers)
    gUiLastMsg = res
    If Left(res, 3) = "OK:" Then
        gSzText = ""
        a = Split(Mid(res, 4), "|")
        eis = CStr(a(1))
        gSzSeq = 0
        msg = "Готово: приход проведён, ЕИ с серийными номерами — " & a(0) & " (" & eis & ")."
        If gSzReplaced <> "" Then msg = msg & Chr(10) & "Прежний ЕИ " & Left(gSzReplaced, InStr(gSzReplaced, "|") - 1) & " был пуст и нёс номер " & Mid(gSzReplaced, InStr(gSzReplaced, "|") + 1) _
            & ": этот номер записан в его примечание («" & SZ_AGAIN & "…»), лист «" & SZ_SHEET & "» покажет его как заменённый."
        k = WmsPrimeUi.UiButtons("Серийные детали", msg & Chr(10) & Chr(10) & "Напечатать этикетки (ЕИ, артикул, серийный номер)?", Array("Напечатать этикетки", "Не сейчас"), Array(1, 0))
        If k = 1 Then LabelsOf(eis)
    ElseIf Left(res, 8) = "PARTIAL:" Then
        a = Split(Mid(res, 9), "|")
        WmsUi.UiMessage("Приход проведён (" & a(0) & "), но серийные номера записаны не все: " & a(1) & Chr(10) & Chr(10) & "Номера остались в окне: встаньте на " & a(0) & " на «Наличии», нажмите «Серийные…» и впишите недостающие.")
    Else
        WmsUi.UiMessage(Mid(res, InStr(res, ":") + 1))
    End If
    Exit Sub
EH:
    WmsUi.UiMessage("Серийный приход не выполнен: " & Error$ & " (строка " & Erl & ")")
End Sub

' the labels of the EIs "ЕИ-…,ЕИ-…", one of each (the page «Печать»)
Private Sub LabelsOf(sEis As String)
    Dim a As Variant, eis() As Long, counts() As Long, i As Long, msg As String
    a = Split(sEis, ",")
    ReDim eis(UBound(a))
    ReDim counts(UBound(a))
    For i = 0 To UBound(a)
        eis(i) = WmsView.EiNo(Trim(CStr(a(i))), CLng(SysNum(SK_NEXT_EI)))
        counts(i) = 1
    Next i
    msg = WmsItem.LabelPrintN(eis, counts, WmsItem.LabelSize())
    If Left(msg, 4) = "ERR:" Then WmsUi.UiMessage(Mid(msg, 5))
End Sub

' ================================================================ the sheet «Серийные детали»

Private Function SzSeq() As Long
    SzSeq = CLng(SysNum(SK_LAST_SEQ))
    If SzSeq = 0 Then SzSeq = 1
End Function

Private Function SzFind() As String
    If gDoc.Sheets.hasByName(SZ_SHEET) Then SzFind = Trim(gDoc.Sheets.getByName(SZ_SHEET).getCellByPosition(1, SZ_FIND_ROW).getString())
End Function

' «да» in the cell next to «Показывать выданные»: the default is to show everything
Private Function SzAll() As Boolean
    Dim t As String
    SzAll = True
    If gDoc.Sheets.hasByName(SZ_SHEET) Then
        t = LCase(Trim(gDoc.Sheets.getByName(SZ_SHEET).getCellByPosition(SZ_ALL_COL, SZ_FIND_ROW).getString()))
        If t = "нет" Or t = "no" Or t = "0" Then SzAll = False
    End If
End Function

Sub NavSerials(Optional oEvent As Variant)
    WmsInit()
    WmsView.GoCell(SZ_SHEET, SZ_FIND_ROW, 1)
End Sub

Sub BtnSerialsRefresh(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SerialsRefresh(SzFind(), SzAll())
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Sub OnSerialsFocus(Optional oEvent As Variant)
    If gSzBusy Then Exit Sub
    On Error GoTo EH
    WmsInit()
    If SzSeq() = gSzSeq And gSzSeq > 0 Then Exit Sub
    SerialsRefresh(SzFind(), SzAll())
    Exit Sub
EH:
End Sub

Sub OnSerialsChange(Optional oTarget As Variant)
    Dim ra As Object
    If gSzBusy Then Exit Sub
    If IsMissing(oTarget) Then Exit Sub
    On Error GoTo EH
    If Not oTarget.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    ra = oTarget.getCellAddress()
    If ra.Row <> SZ_FIND_ROW Or (ra.Column <> 1 And ra.Column <> SZ_ALL_COL) Then Exit Sub
    WmsInit()
    SerialsRefresh(SzFind(), SzAll())
    Exit Sub
EH:
End Sub

Private Function QtyOfV(v As Variant) As Double
    If VarType(v) = 5 Then QtyOfV = v
End Function

' the EI → the last posted issue of it: the index of the row of the list is kept in a Collection by the code of the EI
Private Function LastIssues(idx As Object, ByRef d As Variant) As Boolean
    Dim sh As Object, last As Long, i As Long, ei As String, ctl As String, cEi As Long, cCtl As Long
    sh = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, IC_LAST, last).getDataArray()
    cEi = IC_EI
    cCtl = IC_CTL
    ' from the end: the first one put under an EI stays, so it is the last posted issue
    For i = UBound(d) To 0 Step -1
        ei = Trim(CStr(d(i)(cEi)))
        If ei <> "" Then
            ctl = CStr(d(i)(cCtl))
            If Left(ctl, 9) = "Проведено" Then WmsView.CollPut(idx, i, UCase(ei))
        End If
    Next i
    LastIssues = True
End Function

' the list: one row per serial number of «Наличия» (the property «серийный №») and per number that an EI carried before (the note «ранее № …»); whose text has every word of sFind;
' bAll = False — only the parts on the stock. OK:<rows> / ERR:…
Function SerialsRefresh(sFind As String, bAll As Boolean) As String
    Dim sh As Object, st As Object, ps As Object, last As Long, d As Variant, pd As Variant, iss As Variant, idx As Object, i As Long, n As Long, nOut As Long, q As Double
    Dim out() As Variant, hays() As String, folded As Variant, words As Variant, bWords As Boolean, ser As String, note As String, canon As String, state As String, who As String
    Dim dt As Variant, k As Long, place As String, src As String, p As Long, q2 As Long, tail As String, again As String, hasIss As Boolean, nStock As Long, nAll As Long
    Dim cEi As Long, cNm As Long, cArt As Long, cQ As Long, cPl As Long, cSrc As Long, cWho As Long, cDt As Long
    On Error GoTo EH
    If Not gDoc.Sheets.hasByName(SZ_SHEET) Then
        SerialsRefresh = "ERR:в книге нет листа «" & SZ_SHEET & "» — книга не обновлена до 1.1.1"
        Exit Function
    End If
    gSzBusy = True
    WmsInit()
    sh = gDoc.Sheets.getByName(SZ_SHEET)
    idx = New Collection
    st = WmsIssue.StockSheet()
    last = CLng(SysNum(SK_NEXT_EI)) - 1
    If last >= 1 And gDoc.Sheets.hasByName(SH_PROP) Then
        d = st.getCellRangeByPosition(0, 1, SC_LAST, last).getDataArray()
        ps = gDoc.Sheets.getByName(SH_PROP)
        pd = ps.getCellRangeByPosition(0, 1, SZ_PROP_COLS, last).getDataArray()
        hasIss = LastIssues(idx, iss)
        cEi = SC_EI
        cNm = SC_NAME
        cArt = SC_ART
        cQ = SC_QTY
        cPl = SC_PLACE
        cSrc = SC_SRC
        cWho = IC_WHO
        cDt = IC_DATE
        words = WmsNames.FoldWords(sFind)
        bWords = (UBound(words) >= 0)
        ReDim out(4 * (UBound(pd) + 1) + 4)
        ReDim hays(4 * (UBound(pd) + 1) + 4)
        For i = 0 To UBound(pd)
            ser = Trim(CStr(pd(i)(1)))
            note = CStr(pd(i)(3))
            If (ser <> "" Or InStr(note, SZ_AGAIN) > 0) And i <= UBound(d) Then
                canon = CStr(d(i)(cEi))
                If canon <> "" Then
                    q = QtyOfV(d(i)(cQ))
                    place = Trim(CStr(d(i)(cPl)))
                    src = Trim(CStr(d(i)(cSrc)))
                    who = ""
                    dt = ""
                    state = ""
                    If ser <> "" Then
                        nAll = nAll + 1
                        If q > 0.0000005 Then
                            state = "на складе"
                            nStock = nStock + 1
                        Else
                            state = "нет на складе"
                            If hasIss Then
                                k = WmsView.CollGet(idx, UCase(canon))
                                If k >= 0 Then
                                    state = "выдана"
                                    who = Trim(CStr(iss(k)(cWho)))
                                    dt = iss(k)(cDt)
                                    If VarType(dt) <> 5 Then dt = ""
                                End If
                            End If
                        End If
                        If (bAll Or q > 0.0000005) And nOut < UBound(out) Then
                            out(nOut) = Array(ser, Trim(CStr(d(i)(cArt))), WmsNames.NormNameFast(CStr(d(i)(cNm))), canon, place, q, state, who, dt, SrcNote(src, note))
                            hays(nOut) = ser & " " & CStr(d(i)(cArt)) & " " & CStr(d(i)(cNm)) & " " & canon & " " & place & " " & who
                            nOut = nOut + 1
                        End If
                    End If
                    ' the numbers this EI carried before ("ранее № A · ранее № B"): shown as replaced
                    p = InStr(note, SZ_AGAIN)
                    Do While p > 0
                        tail = Mid(note, p + Len(SZ_AGAIN))
                        q2 = InStr(tail, " · ")
                        If q2 > 0 Then again = Trim(Left(tail, q2 - 1)) Else again = Trim(tail)
                        If again <> "" And bAll And nOut < UBound(out) Then
                            out(nOut) = Array(again, Trim(CStr(d(i)(cArt))), WmsNames.NormNameFast(CStr(d(i)(cNm))), canon, place, 0, "заменён в ЕИ", "", "", "номер деталь уже не несёт: ЕИ получил новую партию (сейчас № " & ser & ")")
                            hays(nOut) = again & " " & CStr(d(i)(cArt)) & " " & CStr(d(i)(cNm)) & " " & canon
                            nOut = nOut + 1
                        End If
                        p = InStr(p + Len(SZ_AGAIN), note, SZ_AGAIN)
                    Loop
                End If
            End If
        Next i
        n = nOut
        If bWords And n > 0 Then
            ReDim Preserve hays(n - 1)
            folded = WmsNames.FoldAll(hays)
            nOut = 0
            For i = 0 To n - 1
                If WmsNames.HasWords(CStr(folded(i)), words) Then
                    out(nOut) = out(i)
                    nOut = nOut + 1
                End If
            Next i
        End If
    End If
    SzWrite(sh, out, nOut)
    sh.getCellByPosition(0, SZ_STAT_ROW).setString("Серийных деталей: " & nAll & ", из них на складе " & nStock & IIf(Trim(sFind) <> "", "; по словам «" & Trim(sFind) & "» показано " & nOut, "") _
        & IIf(bAll, "", " (выданные скрыты)") & ". Построено " & Format(Now(), "DD.MM.YYYY HH:MM") & "; журнал № " & CLng(SysNum(SK_LAST_SEQ)) & ".")
    gSzSeq = SzSeq()
    gSzBusy = False
    SerialsRefresh = "OK:" & nOut
    gUiLastMsg = SerialsRefresh
    Exit Function
EH:
    gSzBusy = False
    SerialsRefresh = "ERR:список «" & SZ_SHEET & "» не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = SerialsRefresh
End Function

Private Function SrcNote(src As String, note As String) As String
    SrcNote = src & IIf(src <> "" And note <> "", " · ", "") & note
End Function

' the rows written under the header, the old ones out; sorted by the serial number (the Russian alphabet, the case does not matter)
Private Sub SzWrite(sh As Object, rowsOut As Variant, n As Long)
    Dim cur As Object, e As Long, rng As Object, out() As Variant, i As Long
    Dim f(0) As New com.sun.star.table.TableSortField, pv(1) As New com.sun.star.beans.PropertyValue, loc As New com.sun.star.lang.Locale
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    e = cur.getRangeAddress().EndRow
    If SZ_ROW0 + n - 1 > e Then e = SZ_ROW0 + n - 1
    If e >= SZ_ROW0 Then sh.getCellRangeByPosition(0, SZ_ROW0, SZ_COLS - 1, e).clearContents(1 + 2 + 4 + 16)
    If n = 0 Then Exit Sub
    ReDim out(n - 1)
    For i = 0 To n - 1
        out(i) = rowsOut(i)
    Next i
    rng = sh.getCellRangeByPosition(0, SZ_ROW0, SZ_COLS - 1, SZ_ROW0 + n - 1)
    rng.setDataArray(out)
    loc.Language = "ru"
    loc.Country = "RU"
    f(0).Field = 0
    f(0).IsAscending = True
    f(0).IsCaseSensitive = False
    f(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC
    f(0).CollatorLocale = loc
    pv(0).Name = "SortFields"
    pv(0).Value = f()
    pv(1).Name = "ContainsHeader"
    pv(1).Value = False
    rng.sort(pv())
End Sub

' «Этикетки выбранных»: the labels of the EIs of the rows selected on the sheet (one of each, with the article and the serial number); a number replaced earlier has no label
Sub BtnSerialsLabels(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = SerialsLabels()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

Function SerialsLabels() As String
    Dim sel As Object, ra As Object, ranges As Variant, i As Long, r As Long, v As String, lst As String, a As Variant, eis() As Long, counts() As Long, msg As String
    Dim sh As Object, n As Long, nextEI As Long, rr As Variant
    On Error GoTo EH
    If gDoc.getCurrentController().getActiveSheet().Name <> SZ_SHEET Then
        SerialsLabels = "ERR:выделите строки на листе «" & SZ_SHEET & "»"
        Exit Function
    End If
    sh = gDoc.Sheets.getByName(SZ_SHEET)
    sel = gDoc.getCurrentController().getSelection()
    nextEI = CLng(SysNum(SK_NEXT_EI))
    If sel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        ranges = sel.getRangeAddresses()
    Else
        ranges = Array(sel.getRangeAddress())
    End If
    For i = 0 To UBound(ranges)
        For r = ranges(i).StartRow To ranges(i).EndRow
            If r >= SZ_ROW0 Then
                v = Trim(sh.getCellByPosition(3, r).getString())
                If v <> "" And InStr("," & lst & ",", "," & v & ",") = 0 Then
                    If Trim(sh.getCellByPosition(6, r).getString()) <> "заменён в ЕИ" Then lst = lst & IIf(lst <> "", ",", "") & v
                End If
            End If
        Next r
    Next i
    If lst = "" Then
        SerialsLabels = "ERR:выделите строки с серийными номерами (ниже заголовка) — этикетки печатаются для их ЕИ"
        Exit Function
    End If
    a = Split(lst, ",")
    ReDim eis(UBound(a))
    ReDim counts(UBound(a))
    For n = 0 To UBound(a)
        eis(n) = WmsView.EiNo(CStr(a(n)), nextEI)
        counts(n) = 1
    Next n
    msg = WmsItem.LabelPrintN(eis, counts, WmsItem.LabelSize())
    SerialsLabels = msg
    gUiLastMsg = msg
    Exit Function
EH:
    SerialsLabels = "ERR:этикетки не собраны: " & Error$ & " (строка " & Erl & ")"
End Function

' test seams: the text of the serial numbers as the window gave it; the text the window opens with (the tests of the real window press its buttons only)
Function TestSerialText() As String
    TestSerialText = gSzText
End Function

Function TestSerialSetText(s As String) As String
    WmsInit()
    If SysStr(SK_MODE) <> "TEST" Then
        TestSerialSetText = "REFUSED:не тестовая книга"
        Exit Function
    End If
    gSzText = s
    TestSerialSetText = "OK"
End Function
