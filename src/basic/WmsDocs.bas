' WmsDocs — 1.0.2: документы в Writer по образцам пользователя. Акт приёма-передачи (в офис) и акт выдачи (получателю) по выделенным
' строкам «Выдач» — с номером; реестр актов — лист «Акты». Отчёт дня — по шаблону «Завершено сегодня / Задачи на … / В процессе»
' с цифрами из журнала. Документ — обычный новый документ Writer: его можно править, печатать и сохранять как угодно; книга, реестр и
' журнал им не меняются (кроме номера акта в «№ документа выдачи» выдач, где он был пуст, — прежней операцией ISSUE_DOC). Library WmsPlus.
Option Explicit

' «Акты»: the register (A № акта, B дата, C вид, D получатель, E позиций, F выдачи (их №), G заметка) and the settings on the right (J)
Public Const SH_ACTS = "Акты"
Public Const AK_NO = 0
Public Const AK_DATE = 1
Public Const AK_KIND = 2
Public Const AK_WHO = 3
Public Const AK_N = 4
Public Const AK_ISS = 5
Public Const AK_NOTE = 6
Public Const AK_LAST = 6
Public Const AK_SET_COL = 9
Public Const AK_SET_CITY = 1
Public Const AK_SET_OFFICE = 2
Public Const AK_SET_SIGNER = 3
Public Const AK_KIND1 = "Приёма-передачи (в офис)"
Public Const AK_KIND2 = "Выдачи (получателю)"
' the most issues in one act
Public Const AK_MAX = 60

' «Отчёт дня»: B2 the date (empty — today), B3 the signature; three lists from row 5: done by hand (A), tasks for the next working
' day (C), in progress (E)
Public Const SH_DAYREP = "Отчёт дня"
Public Const DR_DATE_ROW = 1
Public Const DR_SIGN_ROW = 2
Public Const DR_END_ROW = 3
Public Const DR_ROW0 = 5
Public Const DR_ROWS = 20
Public Const DR_COL_DONE = 0
Public Const DR_COL_TASKS = 2
Public Const DR_COL_PROG = 4

' the last document built: the tests read it and close it, a person keeps it open
Global gDcDoc As Object
' the lines of the report being built: «D|…» done today, «T|…» tasks, «P|…» in progress, «C|…» control
Global gDrOut As String
' 1.1: a report of a period is being built (the day lists written by hand and the state of today are left out of a past day); the operations of the day read last; the month of the journal read last
Global gDrPeriod As Boolean
Global gDrOps As Long
Global gDrUrl As String
Global gDrLines As Variant

' ================================================================ Writer: the building blocks

Private Function DcNewDoc(bShow As Boolean) As Object
    Dim a(0) As New com.sun.star.beans.PropertyValue, doc As Object, ps As Object, st As Object
    If bShow Then
        doc = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, Array())
    Else
        a(0).Name = "Hidden"
        a(0).Value = True
        doc = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, a())
    End If
    ' A4, the margins of the samples (2 cm left, 1,5 cm right and above), a serif font of 12 pt
    ps = doc.StyleFamilies.getByName("PageStyles").getByName("Standard")
    ps.LeftMargin = 2000
    ps.RightMargin = 1500
    ps.TopMargin = 1500
    ps.BottomMargin = 1500
    st = doc.StyleFamilies.getByName("ParagraphStyles").getByName("Standard")
    st.CharFontName = "Liberation Serif"
    st.CharHeight = 12
    DcNewDoc = doc
End Function

' a paragraph at the end of the text: nAlign 0 left, 1 right, 2 justified, 3 centred; nPt the size; nBelow the space under it
' (1/100 mm); nFirst the indent of the first line, nLeft the left indent (a hanging indent of a list: nFirst < 0); nTab the position of
' one tab stop, bTabRight — it is a right one (the date at the right margin), otherwise a left one
Private Sub DcPara(doc As Object, s As String, nAlign As Integer, nPt As Double, bBold As Boolean, nBelow As Long, Optional nFirst As Variant, Optional nLeft As Variant, Optional nTab As Variant, Optional bTabRight As Variant)
    Dim cur As Object, ts(0) As New com.sun.star.style.TabStop
    cur = doc.Text.createTextCursor()
    cur.gotoEnd(False)
    cur.ParaAdjust = nAlign
    cur.CharHeight = nPt
    cur.CharWeight = IIf(bBold, 150, 100)
    cur.ParaBottomMargin = nBelow
    cur.ParaTopMargin = 0
    cur.ParaFirstLineIndent = 0
    cur.ParaLeftMargin = 0
    If Not IsMissing(nFirst) Then cur.ParaFirstLineIndent = nFirst
    If Not IsMissing(nLeft) Then cur.ParaLeftMargin = nLeft
    If Not IsMissing(nTab) Then
        ts(0).Position = nTab
        If Not IsMissing(bTabRight) Then
            If bTabRight Then ts(0).Alignment = com.sun.star.style.TabAlign.RIGHT Else ts(0).Alignment = com.sun.star.style.TabAlign.LEFT
        Else
            ts(0).Alignment = com.sun.star.style.TabAlign.LEFT
        End If
        cur.ParaTabStops = ts()
    End If
    doc.Text.insertString(cur, s, False)
    doc.Text.insertControlCharacter(cur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
End Sub

' a table at the end of the text: data — an array of rows (the first one is the header); relative positions of the column
' separators (of 10000, one fewer than the columns); the columns of alignCtr (a comma list of letters) centred
Private Sub DcTable(doc As Object, data As Variant, seps As Variant, sCenter As String)
    Dim cur As Object, tbl As Object, n As Long, nc As Integer, i As Integer, cs As Variant, a As Variant, c As Integer, nm As String, rows() As Variant, j As Integer, row() As Variant
    cur = doc.Text.createTextCursor()
    cur.gotoEnd(False)
    n = UBound(data) + 1
    nc = UBound(data(0)) + 1
    tbl = doc.createInstance("com.sun.star.text.TextTable")
    tbl.initialize(n, nc)
    doc.Text.insertTextContent(cur, tbl, False)
    ReDim rows(n - 1)
    For i = 0 To n - 1
        ReDim row(nc - 1)
        For j = 0 To nc - 1
            row(j) = CStr(data(i)(j))
        Next j
        rows(i) = row
    Next i
    tbl.setDataArray(rows)
    cs = tbl.TableColumnSeparators
    For i = 0 To UBound(cs)
        cs(i).Position = seps(i)
    Next i
    tbl.TableColumnSeparators = cs
    tbl.getCellRangeByName("A1:" & Chr(64 + nc) & "1").ParaAdjust = 3
    tbl.getCellRangeByName("A1:" & Chr(64 + nc) & "1").CharWeight = 100
    a = Split(sCenter, ",")
    For i = 0 To UBound(a)
        If a(i) <> "" And n > 1 Then tbl.getCellRangeByName(a(i) & "2:" & a(i) & n).ParaAdjust = 3
    Next i
End Sub

' ================================================================ the register «Акты» and its settings

Private Function AkSheet() As Object
    If gDoc.Sheets.hasByName(SH_ACTS) Then AkSheet = gDoc.Sheets.getByName(SH_ACTS)
End Function

' the setting k (AK_SET_*) of the register: the text of J, def when it is empty
Private Function AkSetting(k As Integer, def As String) As String
    Dim sh As Object, s As String
    sh = AkSheet()
    If IsNull(sh) Then
        AkSetting = def
        Exit Function
    End If
    s = Trim(sh.getCellByPosition(AK_SET_COL, k).getString())
    AkSetting = IIf(s = "", def, s)
End Function

' True when the recipient `who` is the office (a name of the list in the setting «Офис»: «Офис, Бухгалтерия»)
Function AkIsOffice(who As String) As Boolean
    Dim a As Variant, i As Integer, k As String
    k = WmsAsk.WhoKey(who)
    If k = "" Then Exit Function
    a = Split(AkSetting(AK_SET_OFFICE, "Офис"), ",")
    For i = 0 To UBound(a)
        If WmsAsk.WhoKey(CStr(a(i))) = k Then AkIsOffice = True
    Next i
End Function

' the last row of the register (the settings stand right of it, so the used area of the sheet is not its length): the last filled A; 0 — none
Private Function AkLastRow(sh As Object) As Long
    Dim last As Long, d As Variant, i As Long
    last = WmsOrders.LastRow(sh)
    If last < 1 Then Exit Function
    d = sh.getCellRangeByPosition(0, 1, 0, last).getDataArray()
    For i = UBound(d) To 0 Step -1
        If CStr(d(i)(0)) <> "" Then
            AkLastRow = i + 1
            Exit Function
        End If
    Next i
End Function

' the number of an act in a text «А-12» (a Cyrillic or a Latin A, a dash, digits only); 0 — it is not one
Private Function ActNoOf(ByVal s As String) As Long
    Dim t As String, i As Long
    t = Trim(s)
    If Len(t) < 3 Or Len(t) > 12 Then Exit Function
    If Left(t, 2) <> "А-" And Left(t, 2) <> "A-" Then Exit Function
    t = Mid(t, 3)
    For i = 1 To Len(t)
        If Mid(t, i, 1) < "0" Or Mid(t, i, 1) > "9" Then Exit Function
    Next i
    ActNoOf = CLng(t)
End Function

' the largest number of an act the journal knows (the operations ISSUE_DOC, which wrote «А-<№>» into the issues); 0 — none
Private Function AkJournalMax() As Long
    Dim fl As Variant, i As Long, j As Long, txt As String, lines As Variant, seq As Long, ts As String, inst As String, sType As String, f As Variant, m As Long, k As Long
    On Error GoTo EH
    If gJDir = "" Then Exit Function
    If Not gSfa.exists(gJDir) Then Exit Function
    fl = gSfa.getFolderContents(gJDir, False)
    For i = 0 To UBound(fl)
        If LCase(Right(CStr(fl(i)), 4)) = ".csv" Then
            txt = WmsAudit.ReadText(CStr(fl(i)))
            If InStr(txt, "ISSUE_DOC") > 0 Then
                lines = Split(txt, Chr(10))
                For j = 0 To UBound(lines)
                    If InStr(lines(j), "ISSUE_DOC") > 0 Then
                        If WmsJournal.ParseLine(CStr(lines(j)), seq, ts, inst, sType, f) Then
                            If sType = "ISSUE_DOC" Then
                                k = ActNoOf(WmsView.JField(f, "DOC"))
                                If k > m Then m = k
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    AkJournalMax = m
    Exit Function
EH:
    AkJournalMax = m
End Function

' the next number of an act: the largest of the register, of the «№ документа выдачи» of the issues and of the journal, plus one. The register is the storekeeper's own sheet
' (a return to a copy brings back an older one; a row can be cleared), so the number is not taken from it alone: an act number never repeats
Private Function AkNextNo() As Long
    Dim sh As Object, last As Long, d As Variant, i As Long, m As Double, ish As Object, cDoc As Long, k As Long
    sh = AkSheet()
    If IsNull(sh) Then
        AkNextNo = 1
        Exit Function
    End If
    last = WmsOrders.LastRow(sh)
    If last >= 1 Then
        d = sh.getCellRangeByPosition(AK_NO, 1, AK_NO, last).getDataArray()
        For i = 0 To UBound(d)
            If VarType(d(i)(0)) = 5 Then
                If d(i)(0) > m Then m = d(i)(0)
            End If
        Next i
    End If
    ' the numbers written into the issues
    ish = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(ish)
    cDoc = IC_DOC
    If last >= 1 Then
        d = ish.getCellRangeByPosition(cDoc, 1, cDoc, last).getDataArray()
        For i = 0 To UBound(d)
            k = ActNoOf(CStr(d(i)(0)))
            If k > m Then m = k
        Next i
    End If
    k = AkJournalMax()
    If k > m Then m = k
    AkNextNo = CLng(m) + 1
End Function

' ================================================================ the act

Sub BtnAct(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = ActFromSelection()
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' «Акт» on «Выдачи»: the posted issues of the selection (one row — that row or «всё, что выдано <кому> <дата>») → an act of the kind
' the recipient suggests (the office — «приёма-передачи», a person — «выдачи»), asked once. OK:… / ERR:… / SKIP:…
Function ActFromSelection() As String
    Dim ish As Object, sel As Variant, rs() As Long, nR As Long, i As Long, j As Long, d As Variant, who As String, dt As Double, k As Integer, last As Long, a As Variant
    Dim whoKey As String, n As Long, kind As Integer, office As Boolean, bad As String, lb As Variant, cd As Variant
    On Error GoTo EH
    ish = WmsIssue.IssuesSheet()
    sel = WmsPrint.SelRows(SH_ISSUES)
    If UBound(sel) < 0 Then
        ActFromSelection = "ERR:выделите проведённые выдачи на листе «" & SH_ISSUES & "» (или встаньте на одну) и нажмите «Акт»"
        gUiLastMsg = ActFromSelection
        Exit Function
    End If
    ReDim rs(UBound(sel))
    For i = 0 To UBound(sel)
        rs(i) = sel(i)
    Next i
    nR = UBound(sel) + 1
    If nR = 1 Then
        d = ish.getCellRangeByPosition(0, rs(0), IC_LAST, rs(0)).getDataArray()(0)
        who = Trim(CStr(d(IC_WHO)))
        If VarType(d(IC_DATE)) = 5 Then dt = Int(d(IC_DATE))
        If who <> "" And dt > 0 Then
            k = WmsPrimeUi.UiButtons("Акт", "Выдача № " & d(IC_NO) & " — " & who & ", " & Format(dt, "DD.MM.YYYY") & ".", _
                Array("Только эта выдача", "Всё, что выдано «" & who & "» " & Format(dt, "DD.MM.YYYY"), "Отмена"), Array(1, 2, 0))
            If k = 0 Or k = -1 Then
                ActFromSelection = "SKIP:окно закрыто"
                gUiLastMsg = ActFromSelection
                Exit Function
            End If
            If k = 2 Then
                whoKey = WmsAsk.WhoKey(who)
                last = WmsOrders.LastRow(ish)
                a = ish.getCellRangeByPosition(0, 1, IC_LAST, last).getDataArray()
                ReDim rs(last)
                nR = 0
                For j = 0 To last - 1
                    If VarType(a(j)(IC_DATE)) = 5 Then
                        If Int(a(j)(IC_DATE)) = dt And WmsAsk.WhoKey(CStr(a(j)(IC_WHO))) = whoKey And Left(CStr(a(j)(IC_CTL)), Len(ST_POSTED)) = ST_POSTED Then
                            rs(nR) = j + 1
                            nR = nR + 1
                        End If
                    End If
                Next j
            End If
        End If
    End If
    ' only posted issues of one recipient
    n = 0
    who = ""
    For i = 0 To nR - 1
        d = ish.getCellRangeByPosition(0, rs(i), IC_LAST, rs(i)).getDataArray()(0)
        If Left(CStr(d(IC_CTL)), Len(ST_POSTED)) = ST_POSTED And VarType(d(IC_QTY)) = 5 Then
            If who = "" Then
                who = Trim(CStr(d(IC_WHO)))
                whoKey = WmsAsk.WhoKey(who)
            ElseIf WmsAsk.WhoKey(CStr(d(IC_WHO))) <> whoKey Then
                bad = Trim(CStr(d(IC_WHO)))
            End If
            rs(n) = rs(i)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ActFromSelection = "ERR:среди выделенных строк нет проведённых выдач"
        gUiLastMsg = ActFromSelection
        Exit Function
    End If
    If bad <> "" Then
        ActFromSelection = "ERR:выделены выдачи разным получателям («" & who & "», «" & bad & "»). Акт делается на одного получателя — выделите его выдачи."
        gUiLastMsg = ActFromSelection
        Exit Function
    End If
    If who = "" Then
        ActFromSelection = "ERR:в выделенных выдачах не указан получатель («Кому выдано») — укажите его (исправьте выдачу) и повторите"
        gUiLastMsg = ActFromSelection
        Exit Function
    End If
    If n > AK_MAX Then
        ActFromSelection = "ERR:выдач в акте " & n & " — за один раз не больше " & AK_MAX & ". Выделите меньше строк."
        gUiLastMsg = ActFromSelection
        Exit Function
    End If
    office = AkIsOffice(who)
    ' the kind the recipient suggests is the first button (an array is not taken from IIf: it is not kept as one)
    If office Then
        lb = Array("Акт приёма-передачи (в офис)", "Акт выдачи (получателю)", "Отмена")
        cd = Array(1, 2, 0)
    Else
        lb = Array("Акт выдачи (получателю)", "Акт приёма-передачи (в офис)", "Отмена")
        cd = Array(2, 1, 0)
    End If
    k = WmsPrimeUi.UiButtons("Какой акт", "Получатель «" & who & "», выдач: " & n & ".", lb, cd)
    If k <= 0 Then
        ActFromSelection = "SKIP:окно закрыто"
        gUiLastMsg = ActFromSelection
        Exit Function
    End If
    kind = k
    ReDim Preserve rs(n - 1)
    ActFromSelection = ActMake(rs, kind, gUiAuto = 0)
    gUiLastMsg = ActFromSelection
    Exit Function
EH:
    ActFromSelection = "ERR:акт не составлен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = ActFromSelection
End Function

' the act of the posted issues in rows rs (of «Выдачи»), kind 1 — приёма-передачи (в офис), 2 — выдачи: the number is taken, written
' into «№ документа выдачи» of the issues that had none («А-<№>», ISSUE_DOC), the register gets a row, the document is built.
' OK:<№>|<kind>|<recipient>|<positions> / ERR:…
Function ActMake(rs As Variant, kind As Integer, bShow As Boolean) As String
    Dim ish As Object, n As Long, no As Long, i As Long, d As Variant, who As String, dt As Double, nos As String, res As String, sh As Object, r As Long
    Dim rr() As Long
    On Error GoTo EH
    WmsInit()
    If IsNull(AkSheet()) Then
        ActMake = "ERR:в книге нет листа «" & SH_ACTS & "» — книга не обновлена до 1.0.2"
        Exit Function
    End If
    ish = WmsIssue.IssuesSheet()
    n = UBound(rs) + 1
    If n < 1 Or n > AK_MAX Then
        ActMake = "ERR:выдач в акте — от 1 до " & AK_MAX
        Exit Function
    End If
    ReDim rr(n - 1)
    For i = 0 To n - 1
        rr(i) = rs(i)
        d = ish.getCellRangeByPosition(0, rr(i), IC_LAST, rr(i)).getDataArray()(0)
        If Left(CStr(d(IC_CTL)), Len(ST_POSTED)) <> ST_POSTED Or VarType(d(IC_QTY)) <> 5 Then
            ActMake = "ERR:строка " & (rr(i) + 1) & " «" & SH_ISSUES & "» — не проведённая выдача"
            Exit Function
        End If
        If who = "" Then who = Trim(CStr(d(IC_WHO)))
        If dt = 0 And VarType(d(IC_DATE)) = 5 Then dt = Int(d(IC_DATE))
        nos = nos & IIf(nos <> "", ",", "") & d(IC_NO)
    Next i
    no = AkNextNo()
    ' the number into the issues that had none (the operation of the journal ISSUE_DOC: recovery and the self-check know it)
    res = WmsPrint.NoteNumber(ish, rr, n, "А-" & no)
    If Left(res, 4) = "ERR:" Or Left(res, 4) = "ERR-" Then
        ActMake = "ERR:номер акта не записан в выдачи — " & Mid(res, InStr(res, ":") + 1)
        Exit Function
    End If
    sh = AkSheet()
    r = AkLastRow(sh) + 1
    sh.getCellRangeByPosition(AK_NO, r, AK_LAST, r).setDataArray(Array(Array(CDbl(no), CDbl(WmsOrders.Today()), IIf(kind = 1, AK_KIND1, AK_KIND2), who, CDbl(n), nos, "")))
    res = ActDoc(rr, n, kind, no, WmsOrders.Today(), who, bShow)
    If Left(res, 4) = "ERR:" Then
        sh.getCellByPosition(AK_NOTE, r).setString("документ не построен: " & Mid(res, 5))
        ActMake = res
        Exit Function
    End If
    ActMake = "OK:" & no & "|" & kind & "|" & who & "|" & n
    gUiLastMsg = ActMake
    Exit Function
EH:
    ActMake = "ERR:акт не составлен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = ActMake
End Function

' the document of an act: the head as in the sample of the storekeeper («АКТ № …», «приема-передачи оборудования», the city and the
' date, the sentence, the table № п/п · Наименование · Артикул · Количество · Ед изм, the signatures)
Private Function ActDoc(rr() As Long, n As Long, kind As Integer, no As Long, dt As Double, who As String, bShow As Boolean) As String
    Dim ish As Object, doc As Object, data() As Variant, i As Long, d As Variant, en As Long, pu As String, pn As Double, nm As String, q As Double, qtxt As String, unit As String
    Dim p As Variant, nextEI As Long, signer As String
    On Error GoTo EH
    ish = WmsIssue.IssuesSheet()
    nextEI = CLng(SysNum(SK_NEXT_EI))
    ReDim data(n)
    data(0) = Array("№ п/п", "Наименование", "Артикул", "Количество", "Ед изм")
    For i = 0 To n - 1
        d = ish.getCellRangeByPosition(0, rr(i), IC_LAST, rr(i)).getDataArray()(0)
        en = WmsView.EiNo(CStr(d(IC_EI)), nextEI)
        nm = Trim(CStr(d(IC_NAME)))
        If en > 0 Then
            p = WmsItem.PropOf(en)
            If CStr(p(0)) <> "" Then nm = nm & ", S/N " & p(0)
        End If
        q = CDbl(d(IC_QTY))
        unit = Trim(CStr(d(IC_UNIT)))
        qtxt = WmsIssue.QtyText(q)
        ' a part of a unit with a packing («1 кан» of «1 шт = 4 кан»): the parts and their name
        If en > 0 And Abs(q - Int(q + 0.0000001)) > 0.0000001 Then
            If WmsPack.PackOf(en, pu, pn) Then
                qtxt = WmsIssue.QtyText(WmsPack.PkRound(q * pn, 3))
                unit = pu
            End If
        End If
        data(i + 1) = Array((i + 1) & ".", nm, CStr(d(IC_ART)), qtxt, unit)
    Next i
    doc = DcNewDoc(bShow)
    DcPara(doc, "АКТ № " & no, 3, 14, False, 100)
    DcPara(doc, IIf(kind = 1, "приема-передачи оборудования", "выдачи оборудования"), 3, 12, False, 500)
    DcPara(doc, AkSetting(AK_SET_CITY, "г Москва") & Chr(9) & Format(CDate(dt), "DD.MM.YYYY") & " г.", 0, 12, False, 500, 0, 0, 17500, True)
    If kind = 1 Then
        DcPara(doc, "Мы, нижеподписавшиеся, составили Акт о том, что со стороны склада было передано, а офисом принято следующее оборудование:", 2, 12, False, 400, 1000)
    Else
        DcPara(doc, "Мы, нижеподписавшиеся, составили Акт о том, что со стороны склада было выдано, а получателем " & who & " принято следующее оборудование:", 2, 12, False, 400, 1000)
    End If
    DcTable(doc, data, Array(750, 5700, 7500, 9000), "A,D,E")
    signer = AkSetting(AK_SET_SIGNER, "")
    DcPara(doc, "", 0, 12, False, 1500)
    DcPara(doc, IIf(kind = 1, "Оборудование сдал:" & Chr(9) & "Оборудование принял:", "Оборудование выдал:" & Chr(9) & "Оборудование получил:"), 0, 12, False, 900, 0, 0, 8800, False)
    DcPara(doc, "______________________" & IIf(signer <> "", " / " & signer & " /", "") & Chr(9) & "______________________" & IIf(kind = 2, " / " & who & " /", ""), 0, 12, False, 0, 0, 0, 8800, False)
    gDcDoc = doc
    ActDoc = "OK:"
    Exit Function
EH:
    ActDoc = "ERR:" & Error$ & " (строка " & Erl & ")"
End Function

Sub BtnActRepeat(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = ActRepeat(WmsView.ViewRow(SH_ACTS), gUiAuto = 0)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the act of row r of the register built again (the same number, date, kind and issues) — a lost or spoilt paper is printed again
Function ActRepeat(r As Long, bShow As Boolean) As String
    Dim sh As Object, d As Variant, ish As Object, last As Long, a As Variant, nos As Variant, i As Long, j As Long, rr() As Long, n As Long, kind As Integer, res As String
    On Error GoTo EH
    WmsInit()
    sh = AkSheet()
    If IsNull(sh) Or r < 1 Then
        ActRepeat = "ERR:встаньте на строку акта в реестре «" & SH_ACTS & "» и нажмите «Повторить акт»"
        gUiLastMsg = ActRepeat
        Exit Function
    End If
    d = sh.getCellRangeByPosition(0, r, AK_LAST, r).getDataArray()(0)
    If VarType(d(AK_NO)) <> 5 Then
        ActRepeat = "ERR:в строке нет номера акта"
        gUiLastMsg = ActRepeat
        Exit Function
    End If
    kind = IIf(CStr(d(AK_KIND)) = AK_KIND1, 1, 2)
    nos = Split(CStr(d(AK_ISS)), ",")
    ish = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(ish)
    a = ish.getCellRangeByPosition(0, 1, IC_NO, last).getDataArray()
    ReDim rr(UBound(nos))
    For i = 0 To UBound(nos)
        For j = 0 To UBound(a)
            If VarType(a(j)(IC_NO)) = 5 Then
                If CStr(CLng(a(j)(IC_NO))) = Trim(CStr(nos(i))) Then
                    rr(n) = j + 1
                    n = n + 1
                    Exit For
                End If
            End If
        Next j
    Next i
    If n = 0 Then
        ActRepeat = "ERR:выдачи акта № " & d(AK_NO) & " в «" & SH_ISSUES & "» не найдены"
        gUiLastMsg = ActRepeat
        Exit Function
    End If
    ReDim Preserve rr(n - 1)
    res = ActDoc(rr, n, kind, CLng(d(AK_NO)), CDbl(d(AK_DATE)), CStr(d(AK_WHO)), bShow)
    ActRepeat = IIf(Left(res, 3) = "OK:", "OK:" & CLng(d(AK_NO)) & "|" & kind & "|" & d(AK_WHO) & "|" & n, res)
    gUiLastMsg = ActRepeat
    Exit Function
EH:
    ActRepeat = "ERR:акт не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = ActRepeat
End Function

' ================================================================ the daily report

Sub BtnDayReport(Optional oEvent As Variant)
    Dim s As String
    WmsInit()
    s = DayReport(DrDate(), gUiAuto = 0)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the day of the report: B2 of «Отчёт дня» (a date), empty — today
Function DrDate() As Double
    Dim sh As Object, c As Object, msg As String, d As Double
    DrDate = WmsOrders.Today()
    If Not gDoc.Sheets.hasByName(SH_DAYREP) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_DAYREP)
    c = sh.getCellByPosition(1, DR_DATE_ROW)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then
        If c.getValue() > 1 Then DrDate = Int(c.getValue())
    ElseIf Trim(c.getString()) <> "" Then
        If WmsOrders.ParseDateTextAs(Trim(c.getString()), "дата отчёта", d, msg) Then DrDate = Int(d)
    End If
End Function

' a working day: Monday … Friday (the schedule 5/2) except the days off of «Календарь», and the working days of «Календарь»
Function DrIsWork(d As Double) As Boolean
    ' 1.1: the schedule 5/2 corrected by the sheet «Календарь» (a holiday on a weekday, a working Saturday)
    DrIsWork = WmsPeriod.IsWorkDay(d)
End Function

' the next working day after d
Function DrNextWork(d As Double) As Double
    DrNextWork = d + 1
    Do While Not DrIsWork(DrNextWork)
        DrNextWork = DrNextWork + 1
    Loop
End Function

' «на понедельник», «на среду», «на пятницу» — the day as the title of the tasks says it
Private Function DrOnDay(d As Double) As String
    Dim a As Variant
    a = Array("понедельник", "вторник", "среду", "четверг", "пятницу", "субботу", "воскресенье")
    DrOnDay = a(Weekday(d, 2) - 1)
End Function

Private Function DrDayName(d As Double) As String
    Dim a As Variant
    a = Array("понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье")
    DrDayName = a(Weekday(d, 2) - 1)
End Function

' the sums by unit as a text «шт=105;м=3» (an array does not come back from a procedure here): q of unit un added
Private Function DrUnitAdd(acc As String, un As String, q As Double) As String
    Dim a As Variant, i As Integer, found As Boolean, r As String, p As Integer, v As Double
    If acc <> "" Then
        a = Split(acc, ";")
        For i = 0 To UBound(a)
            p = InStr(a(i), "=")
            If Left(a(i), p - 1) = un Then
                v = Val(Mid(a(i), p + 1)) + q
                r = r & IIf(r <> "", ";", "") & un & "=" & Str(v)
                found = True
            Else
                r = r & IIf(r <> "", ";", "") & a(i)
            End If
        Next i
    End If
    If Not found Then r = r & IIf(r <> "", ";", "") & un & "=" & Str(q)
    DrUnitAdd = r
End Function

' «105 шт, 3 м» from the sums by unit
Private Function DrUnitText(acc As String) As String
    Dim a As Variant, i As Integer, p As Integer, s As String, un As String
    If acc = "" Then Exit Function
    a = Split(acc, ";")
    For i = 0 To UBound(a)
        p = InStr(a(i), "=")
        un = Left(a(i), p - 1)
        s = s & IIf(s <> "", ", ", "") & WmsIssue.QtyText(WmsPack.PkRound(Val(Mid(a(i), p + 1)), 3)) & IIf(un <> "", " " & un, "")
    Next i
    DrUnitText = s
End Function

' the day of a journal line («J1;seq;2026-10-26T11:36:06;…»): its first 10 characters of the time; "" when the line has none
Private Function DrLineDay(ln As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(ln, ";")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, ln, ";")
    If p2 = 0 Then Exit Function
    DrLineDay = Mid(ln, p2 + 1, 10)
End Function

' the report of day dt: the document in the template of the storekeeper — «1. Завершено сегодня», «2. Задачи на <следующий рабочий
' день>», «3. В процессе», «4. Контроль данных» — the numbers from the journal (what was posted on that day) and from the sheets;
' the lists written by hand — on «Отчёт дня». OK:<lines of the report> / ERR:…
Function DayReport(dt As Double, bShow As Boolean) As String
    Dim doc As Object, i As Long, sign As String, nwd As Double, sh As Object, secs As Variant, titles As Variant, none As Variant, k As Integer, a As Variant, n As Long, cnt(3) As Long
    On Error GoTo EH
    WmsInit()
    gDrOut = ""
    gDrPeriod = False
    DrFill(dt)
    a = Split(gDrOut, Chr(10))
    nwd = DrNextWork(dt)
    sign = "Кладовщик."
    If gDoc.Sheets.hasByName(SH_DAYREP) Then
        sh = gDoc.Sheets.getByName(SH_DAYREP)
        If Trim(sh.getCellByPosition(1, DR_SIGN_ROW).getString()) <> "" Then sign = Trim(sh.getCellByPosition(1, DR_SIGN_ROW).getString())
    End If
    secs = Array("D", "T", "P", "C")
    titles = Array("1. Завершено сегодня:", "2. Задачи на " & DrOnDay(nwd) & " (" & Format(CDate(nwd), "DD.MM.YYYY") & "):", "3. В процессе:", "4. Контроль данных (по журналу WMS):")
    none = Array("операций в WMS за этот день не проведено", "задач не записано", "ничего не записано", "")
    doc = DcNewDoc(bShow)
    DcPara(doc, Format(CDate(dt), "DD.MM.YYYY") & " г.", 0, 12, False, 500)
    DcPara(doc, sign, 0, 12, False, 500)
    For k = 0 To 3
        DcPara(doc, CStr(titles(k)), 0, 12, True, 100)
        n = 0
        For i = 0 To UBound(a)
            If Left(a(i), 2) = secs(k) & "|" Then
                DcPara(doc, "• " & Mid(a(i), 3), 0, 12, False, 60, -500, 500)
                n = n + 1
            End If
        Next i
        If n = 0 And none(k) <> "" Then DcPara(doc, "• " & none(k), 0, 12, False, 60, -500, 500)
        cnt(k) = n
        If k < 3 Then DcPara(doc, "", 0, 6, False, 0)
    Next k
    gDcDoc = doc
    DayReport = "OK:" & cnt(0) & "|" & cnt(1) & "|" & cnt(2) & "|" & cnt(3)
    gUiLastMsg = DayReport
    Exit Function
EH:
    DayReport = "ERR:отчёт не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = DayReport
End Function

' ---------------------------------------------------------------- 1.1: the report of a period

' the last day of a period: B4 of «Отчёт дня» (a date); empty — 0
Function DrEndDate() As Double
    Dim sh As Object, c As Object, msg As String, d As Double
    If Not gDoc.Sheets.hasByName(SH_DAYREP) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_DAYREP)
    c = sh.getCellByPosition(1, DR_END_ROW)
    If c.getType() = com.sun.star.table.CellContentType.VALUE Then
        If c.getValue() > 1 Then DrEndDate = Int(c.getValue())
    ElseIf Trim(c.getString()) <> "" Then
        If WmsOrders.ParseDateTextAs(Trim(c.getString()), "по дату", d, msg) Then DrEndDate = Int(d)
    End If
End Function

Sub BtnPeriodReport(Optional oEvent As Variant)
    Dim s As String, d1 As Double
    WmsInit()
    d1 = DrEndDate()
    If d1 = 0 Then
        WmsUi.UiMessage("Впишите на листе «" & SH_DAYREP & "» даты периода: «Дата отчёта» (с какой) и «по дату» (по какую), потом нажмите «Отчёт за период».")
        Exit Sub
    End If
    s = PeriodReport(DrDate(), d1, gUiAuto = 0)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the week of the date of the report: from Monday to that date
Sub BtnWeekReport(Optional oEvent As Variant)
    Dim s As String, d As Double
    WmsInit()
    d = DrDate()
    s = PeriodReport(d - (Weekday(d, 2) - 1), d, gUiAuto = 0)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the month of the date of the report: from the 1st to that date
Sub BtnMonthReport(Optional oEvent As Variant)
    Dim s As String, d As Double
    WmsInit()
    d = DrDate()
    s = PeriodReport(d - Day(CDate(d)) + 1, d, gUiAuto = 0)
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' The reports of the days from d0 to d1 one after another in one document: for every working day (and a day off that had operations) what was done and how many operations were
' posted, then the totals of the period. The lists written by hand and the state of today belong to the report of the day and are not in it. At most 62 days.
' OK:<days shown>|<days with operations> / ERR:…
Function PeriodReport(d0 As Double, d1 As Double, bShow As Boolean) As String
    Dim doc As Object, d As Double, i As Long, a As Variant, n As Long, sign As String, sh As Object, nShown As Long, nDays As Long, tot As String, t As Double, b As Variant, nW As Long
    On Error GoTo EH
    WmsInit()
    If d1 < d0 Then
        t = d0
        d0 = d1
        d1 = t
    End If
    If d1 - d0 > 61 Then
        PeriodReport = "ERR:период — не больше 62 дней (два месяца); сделайте два отчёта"
        gUiLastMsg = PeriodReport
        Exit Function
    End If
    sign = "Кладовщик."
    If gDoc.Sheets.hasByName(SH_DAYREP) Then
        sh = gDoc.Sheets.getByName(SH_DAYREP)
        If Trim(sh.getCellByPosition(1, DR_SIGN_ROW).getString()) <> "" Then sign = Trim(sh.getCellByPosition(1, DR_SIGN_ROW).getString())
    End If
    doc = DcNewDoc(bShow)
    DcPara(doc, "Отчёт за период " & Format(CDate(d0), "DD.MM.YYYY") & " — " & Format(CDate(d1), "DD.MM.YYYY") & " г.", 0, 12, True, 100)
    DcPara(doc, sign, 0, 12, False, 400)
    gDrPeriod = True
    gDrUrl = ""
    d = d0
    Do While d <= d1
        gDrOut = ""
        DrFill(d)
        nW = IIf(DrIsWork(d), 1, 0)
        If nW = 1 Or gDrOps > 0 Then
            a = Split(gDrOut, Chr(10))
            DcPara(doc, Format(CDate(d), "DD.MM.YYYY") & ", " & DrDayName(d) & IIf(nW = 0, " (нерабочий день)", ""), 0, 12, True, 100)
            n = 0
            For i = 0 To UBound(a)
                If Left(a(i), 2) = "D|" Then
                    DcPara(doc, "• " & Mid(a(i), 3), 0, 12, False, 60, -500, 500)
                    n = n + 1
                End If
            Next i
            If n = 0 Then DcPara(doc, "• операций в WMS за этот день не проведено", 0, 12, False, 60, -500, 500)
            For i = 0 To UBound(a)
                If Left(a(i), 2) = "C|" Then
                    DcPara(doc, "• " & Mid(a(i), 3), 0, 12, False, 60, -500, 500)
                    Exit For
                End If
            Next i
            DcPara(doc, "", 0, 6, False, 0)
            nShown = nShown + 1
            If gDrOps > 0 Then nDays = nDays + 1
        End If
        d = d + 1
    Loop
    gDrPeriod = False
    gDrUrl = ""
    DcPara(doc, "Итого за период:", 0, 12, True, 100)
    tot = WmsPeriod.PeriodTotals(d0, d1, n)
    If tot = "" Then
        DcPara(doc, "• операций в WMS за период не проведено", 0, 12, False, 60, -500, 500)
    Else
        b = Split(tot, Chr(10))
        For i = 0 To UBound(b)
            DcPara(doc, CStr(b(i)), 0, 12, False, 60, -500, 500)
        Next i
    End If
    gDcDoc = doc
    PeriodReport = "OK:" & nShown & "|" & nDays
    gUiLastMsg = PeriodReport
    Exit Function
EH:
    gDrPeriod = False
    gDrUrl = ""
    PeriodReport = "ERR:отчёт за период не построен: " & Error$ & " (строка " & Erl & ")"
    gUiLastMsg = PeriodReport
End Function

' «Очистить «Сделано вручную»»: the list of the day is emptied for the next day (the tasks and «в процессе» stay)
Sub BtnDrClear(Optional oEvent As Variant)
    Dim sh As Object
    WmsInit()
    If Not gDoc.Sheets.hasByName(SH_DAYREP) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_DAYREP)
    sh.getCellRangeByPosition(DR_COL_DONE, DR_ROW0, DR_COL_DONE, DR_ROW0 + DR_ROWS - 1).clearContents(1 + 4)
    gUiLastMsg = "OK:список «Сделано вручную» очищен"
End Sub

' the lines of the four parts of the report of day dt, into gDrOut («D|», «T|», «P|», «C|» — a line each)
Private Sub DrFill(dt As Double)
    Dim url As String, xs As Object, tis As Object, txt As String, a As Variant, i As Long, seq As Long, ts As String, inst As String, sType As String, fields As Variant
    Dim day As String, t As String, q As Double, un As String, u As String, w As String
    Dim nOps As Long, seqFirst As Long, seqLast As Long, nFix As Long, nDel As Long
    Dim nRcv As Long, nPlaced As Long, rcvU As String, eiMin As Long, eiMax As Long, nEiNew As Long, k As Long
    Dim nCars As Long, carSup As String, nSp(4) As Long, nSpAll As Long, codes As Variant, names As Variant
    Dim iDate() As Double, iWho() As String, iQty() As Double, iUnit() As String, iNo() As String, nIss As Long, delNos As String
    Dim nRet As Long, retU As String, nMove As Long, nWo As Long, nInv As Long, nSplit As Long, nSup As Long, nKit As Long
    Dim nSerial As Long, nPack As Long, nDocNo As Long, docNos As String, nCancel As Long, iLo As Long, iHi As Long
    On Error GoTo EH
    ReDim iDate(200)
    ReDim iWho(200)
    ReDim iQty(200)
    ReDim iUnit(200)
    ReDim iNo(200)
    codes = Split(SRC_CODES, "|")
    names = Split(SRC_NAMES, "|")
    day = Format(CDate(dt), "YYYY-MM-DD")
    gDrOps = 0
    ' ---- the journal of the month of the day: the lines of that day, from the end
    url = gJDir & WmsJournal.JournalFileName(CDate(dt))
    If gSfa.exists(url) Then
        ' (a report of a period reads each month of the journal once)
        If gDrPeriod And url = gDrUrl Then
            a = gDrLines
        Else
            xs = gSfa.openFileRead(url)
            tis = CreateUnoService("com.sun.star.io.TextInputStream")
            tis.setInputStream(xs)
            tis.setEncoding("UTF-8")
            txt = tis.readString(Array(), False)
            tis.closeInput()
            a = Split(Replace(txt, Chr(13), ""), Chr(10))
            txt = ""
            gDrUrl = IIf(gDrPeriod, url, "")
            If gDrPeriod Then gDrLines = a
        End If
        ' from the end back to the first line of the day, then the lines of the day in the order they were posted
        iHi = -1
        iLo = 0
        For i = UBound(a) To 0 Step -1
            t = DrLineDay(CStr(a(i)))
            If t <> "" Then
                If t < day Then
                    iLo = i + 1
                    Exit For
                End If
                If t = day And iHi < 0 Then iHi = i
            End If
        Next i
        For i = iLo To iHi
            t = DrLineDay(CStr(a(i)))
            If t <> "" Then
                If t = day Then
                    If WmsJournal.ParseLine(CStr(a(i)), seq, ts, inst, sType, fields) Then
                        nOps = nOps + 1
                        If seqLast = 0 Or seq > seqLast Then seqLast = seq
                        If seqFirst = 0 Or seq < seqFirst Then seqFirst = seq
                        q = Val(WmsView.JField(fields, "QTY"))
                        un = WmsView.JField(fields, "UNIT")
                        Select Case sType
                        Case "RECEIPT", "RECEIPT_ADD"
                            nRcv = nRcv + 1
                            rcvU = DrUnitAdd(rcvU, un, q)
                            If WmsView.JField(fields, "PLACE") <> "" Then nPlaced = nPlaced + 1
                            DrEiRange(WmsView.JField(fields, "EI"), eiMin, eiMax, nEiNew)
                        Case "SP_RECEIPT"
                            k = DrIndexOf(codes, WmsView.JField(fields, "SRC"))
                            If k >= 0 Then nSp(k) = nSp(k) + 1
                            nSpAll = nSpAll + 1
                            If WmsView.JField(fields, "MODE") = "NEW" Then DrEiRange(WmsView.JField(fields, "EI"), eiMin, eiMax, nEiNew)
                            If WmsView.JField(fields, "PLACE") <> "" Then nPlaced = nPlaced + 1
                        Case "SP_REFILL"
                            k = DrIndexOf(codes, WmsView.JField(fields, "SRC"))
                            If k >= 0 Then nSp(k) = nSp(k) + 1
                            nSpAll = nSpAll + 1
                        Case "CAR_ARRIVE"
                            nCars = nCars + 1
                            w = WmsView.JField(fields, "SUPPLIER")
                            If w <> "" And InStr(";" & carSup & ";", ";" & w & ";") = 0 Then carSup = carSup & IIf(carSup <> "", ";", "") & w
                        Case "ISSUE"
                            If nIss > UBound(iDate) Then
                                ReDim Preserve iDate(2 * nIss)
                                ReDim Preserve iWho(2 * nIss)
                                ReDim Preserve iQty(2 * nIss)
                                ReDim Preserve iUnit(2 * nIss)
                                ReDim Preserve iNo(2 * nIss)
                            End If
                            iDate(nIss) = DrIsoDate(WmsView.JField(fields, "DATE"))
                            iWho(nIss) = WmsView.JField(fields, "WHO")
                            iQty(nIss) = q
                            iUnit(nIss) = un
                            iNo(nIss) = WmsView.JField(fields, "NO")
                            nIss = nIss + 1
                        Case "ISSUE_DEL"
                            delNos = delNos & ";" & WmsView.JField(fields, "NO") & ";"
                        Case "ISSUE_DOC"
                            nDocNo = nDocNo + 1
                            w = WmsView.JField(fields, "DOC")
                            If w <> "" And InStr(";" & docNos & ";", ";" & w & ";") = 0 Then docNos = docNos & IIf(docNos <> "", ";", "") & w
                        Case "RETURN"
                            nRet = nRet + 1
                            retU = DrUnitAdd(retU, un, q)
                        Case "MOVE"
                            nMove = nMove + 1
                        Case "WRITE_OFF"
                            nWo = nWo + 1
                        Case "INV_ADJ"
                            nInv = nInv + 1
                        Case "SPLIT"
                            nSplit = nSplit + 1
                        Case "SUP_RET"
                            nSup = nSup + 1
                        Case "KIT_BUILD"
                            nKit = nKit + 1
                        Case "EI_PROP"
                            If WmsView.JField(fields, "SERIAL") <> "" Then nSerial = nSerial + 1
                        Case "EI_PACK"
                            nPack = nPack + 1
                        Case "ORDER_CANCEL", "ORDER_CANCEL_REST"
                            nCancel = nCancel + 1
                        End Select
                        If Right(sType, 4) = "_FIX" Then nFix = nFix + 1
                        If Right(sType, 4) = "_DEL" Then nDel = nDel + 1
                    End If
                End If
            End If
        Next i
    End If
    ' ---- 1. done today
    If nCars > 0 Then DrAdd("D", "Принято машин с поставками: " & nCars & IIf(carSup <> "", " (" & Replace(carSup, ";", ", ") & ")", ""))
    If nRcv > 0 Then DrAdd("D", "Приёмка по заказам: принято позиций — " & nRcv & "; количество: " & DrUnitText(rcvU) & ". Товар пересчитан и принят в WMS")
    If nSpAll > 0 Then
        u = ""
        For i = 0 To 4
            If nSp(i) > 0 Then u = u & IIf(u <> "", ", ", "") & names(i) & " — " & nSp(i)
        Next i
        DrAdd("D", "Иной приход (инвентаризация текущих остатков, принятое от производства и др.): записей — " & nSpAll & " (" & u & ")")
    End If
    If nEiNew > 0 Then DrAdd("D", "Присвоены внутренние коды (ЕИ) для складского учёта: " & nEiNew & IIf(eiMax > eiMin, " — от ЕИ-" & Format(eiMin, "00000000") & " до ЕИ-" & Format(eiMax, "00000000"), IIf(eiMin > 0, " — ЕИ-" & Format(eiMin, "00000000"), "")))
    If nPlaced > 0 Then DrAdd("D", "Выставлено на адресное хранение (указано место): " & nPlaced)
    ' an issue posted and deleted (storno) on the same day is not an issue of the day: the deletion is counted below
    k = 0
    For i = 0 To nIss - 1
        If InStr(delNos, ";" & iNo(i) & ";") = 0 Or iNo(i) = "" Then
            iDate(k) = iDate(i)
            iWho(k) = iWho(i)
            iQty(k) = iQty(i)
            iUnit(k) = iUnit(i)
            iNo(k) = iNo(i)
            k = k + 1
        End If
    Next i
    nIss = k
    DrIssues(dt, iDate, iWho, iQty, iUnit, nIss)
    If nRet > 0 Then DrAdd("D", "Принято возвратов: " & nRet & " (" & DrUnitText(retU) & ")")
    u = ""
    If nMove > 0 Then u = u & IIf(u <> "", ", ", "") & "перемещений — " & nMove
    If nWo > 0 Then u = u & IIf(u <> "", ", ", "") & "списаний — " & nWo
    If nInv > 0 Then u = u & IIf(u <> "", ", ", "") & "позиций инвентаризации — " & nInv
    If nSplit > 0 Then u = u & IIf(u <> "", ", ", "") & "разделений ЕИ — " & nSplit
    If nSup > 0 Then u = u & IIf(u <> "", ", ", "") & "возвратов поставщику — " & nSup
    If nKit > 0 Then u = u & IIf(u <> "", ", ", "") & "сборок комплектов — " & nKit
    If u <> "" Then DrAdd("D", "Корректировки склада: " & u)
    If nSerial > 0 Then DrAdd("D", "Присвоены серийные номера: " & nSerial)
    If nPack > 0 Then DrAdd("D", "Записана фасовка ЕИ: " & nPack)
    DrActsOfDay(dt)
    If nDocNo > 0 Then DrAdd("D", "Оформлены накладные / документы выдачи: " & docNos)
    If nFix + nDel + nCancel > 0 Then DrAdd("D", "Исправления учёта: исправлено операций — " & nFix & ", удалено (сторно) — " & nDel & IIf(nCancel > 0, ", отменено заказов — " & nCancel, ""))
    If Not gDrPeriod Then DrManual("D", DR_COL_DONE)
    ' ---- 2. tasks for the next working day
    If Not gDrPeriod Then
        DrBacklog(dt)
        DrManual("T", DR_COL_TASKS)
        DrOpenWork()
        ' ---- 3. in progress
        DrManual("P", DR_COL_PROG)
        DrStockState()
    End If
    gDrOps = nOps
    ' ---- 4. control
    If nOps > 0 Then
        DrAdd("C", "Операций проведено за день: " & nOps & " (журнал № " & seqFirst & "–" & seqLast & ")")
    Else
        DrAdd("C", "Операций в журнале за этот день нет")
    End If
    u = WmsSafe.StGet("FLASH")
    If u <> "" And Not gDrPeriod Then DrAdd("C", "Последняя копия на флешку: " & Left(u, InStr(u & "|", "|") - 1))
    Exit Sub
EH:
    DrAdd("C", "Часть данных не прочитана: " & Error$ & " (строка " & Erl & ")")
End Sub

' a line of the report: sec «D», «T», «P» or «C»
Private Sub DrAdd(sec As String, s As String)
    gDrOut = gDrOut & sec & "|" & Replace(s, Chr(10), " ") & Chr(10)
End Sub

Private Function DrIndexOf(a As Variant, s As String) As Integer
    Dim i As Integer
    DrIndexOf = -1
    For i = 0 To UBound(a)
        If a(i) = s Then
            DrIndexOf = i
            Exit Function
        End If
    Next i
End Function

' «ЕИ-00000201» into the range of the EIs of the day (the smallest and the largest number, how many)
Private Sub DrEiRange(canon As String, ByRef lo As Long, ByRef hi As Long, ByRef n As Long)
    Dim k As Long
    k = WmsView.EiNo(canon, 99999999)
    If k < 1 Then Exit Sub
    If lo = 0 Or k < lo Then lo = k
    If k > hi Then hi = k
    n = n + 1
End Sub

' «2026-10-22» into a date number; 0 when it is not one
Private Function DrIsoDate(s As String) As Double
    If Len(s) < 10 Then Exit Function
    If Mid(s, 5, 1) <> "-" Or Mid(s, 8, 1) <> "-" Then Exit Function
    DrIsoDate = CDbl(DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 6, 2)), CInt(Mid(s, 9, 2))))
End Function

' the issues posted on the day, by the date of the issue: those of earlier days are the digitising of the previous working day(s)
Private Sub DrIssues(dt As Double, iDate() As Double, iWho() As String, iQty() As Double, iUnit() As String, nIss As Long)
    Dim dates() As Double, nDt As Long, i As Long, j As Long, g As Long, found As Boolean, tmp As Double
    If nIss = 0 Then Exit Sub
    ReDim dates(nIss)
    For i = 0 To nIss - 1
        found = False
        For j = 0 To nDt - 1
            If dates(j) = iDate(i) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            dates(nDt) = iDate(i)
            nDt = nDt + 1
        End If
    Next i
    For i = 0 To nDt - 2
        For j = i + 1 To nDt - 1
            If dates(j) < dates(i) Then
                tmp = dates(i)
                dates(i) = dates(j)
                dates(j) = tmp
            End If
        Next j
    Next i
    ' earlier days first (the digitising), the day itself last
    For g = 0 To nDt - 1
        If dates(g) <> dt Then DrIssueDay(dt, dates(g), iDate, iWho, iQty, iUnit, nIss)
    Next g
    For g = 0 To nDt - 1
        If dates(g) = dt Then DrIssueDay(dt, dates(g), iDate, iWho, iQty, iUnit, nIss)
    Next g
End Sub

Private Sub DrIssueDay(dt As Double, d As Double, iDate() As Double, iWho() As String, iQty() As Double, iUnit() As String, nIss As Long)
    Dim i As Long, j As Long, cnt As Long, un As String, who() As String, wc() As Long, nW As Long, s As String, found As Boolean
    Dim tw As String, tc As Long, label As String, shown As Long
    ReDim who(nIss)
    ReDim wc(nIss)
    For i = 0 To nIss - 1
        If iDate(i) = d Then
            cnt = cnt + 1
            un = DrUnitAdd(un, iUnit(i), iQty(i))
            found = False
            For j = 0 To nW - 1
                If WmsAsk.WhoKey(who(j)) = WmsAsk.WhoKey(iWho(i)) Then
                    wc(j) = wc(j) + 1
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                who(nW) = iWho(i)
                wc(nW) = 1
                nW = nW + 1
            End If
        End If
    Next i
    ' the recipients, the most frequent first
    For i = 0 To nW - 2
        For j = i + 1 To nW - 1
            If wc(j) > wc(i) Then
                tw = who(i)
                who(i) = who(j)
                who(j) = tw
                tc = wc(i)
                wc(i) = wc(j)
                wc(j) = tc
            End If
        Next j
    Next i
    For i = 0 To nW - 1
        If shown >= 6 Then Exit For
        s = s & IIf(s <> "", ", ", "") & IIf(who(i) <> "", who(i), "без получателя") & " — " & wc(i)
        shown = shown + 1
    Next i
    If nW > shown Then s = s & " и ещё получателей: " & (nW - shown)
    If d = 0 Then
        label = "Выдачи без даты выдачи"
    ElseIf d = dt Then
        label = "Выдачи сегодняшним числом (" & Format(CDate(d), "DD.MM.YYYY") & ")"
    Else
        label = "Оцифровка выдач за " & Format(CDate(d), "DD.MM.YYYY") & " (" & DrDayName(d) & ")"
    End If
    DrAdd("D", label & ": проведено выдач — " & cnt & "; выдано: " & DrUnitText(un) & "; получатели: " & s)
End Sub

' the acts of the day from the register
Private Sub DrActsOfDay(dt As Double)
    Dim sh As Object, last As Long, d As Variant, i As Long, s As String, n As Long
    If Not gDoc.Sheets.hasByName(SH_ACTS) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_ACTS)
    last = AkLastRow(sh)
    If last < 1 Then Exit Sub
    d = sh.getCellRangeByPosition(0, 1, AK_LAST, last).getDataArray()
    For i = 0 To UBound(d)
        If VarType(d(i)(AK_DATE)) = 5 And VarType(d(i)(AK_NO)) = 5 Then
            If Int(d(i)(AK_DATE)) = dt Then
                n = n + 1
                s = s & IIf(s <> "", "; ", "") & "№ " & CLng(d(i)(AK_NO)) & " (" & LCase(Left(CStr(d(i)(AK_KIND)), 1)) & Mid(CStr(d(i)(AK_KIND)), 2) & ", " & d(i)(AK_WHO) & ", позиций — " & CLng(d(i)(AK_N)) & ")"
            End If
        End If
    Next i
    If n > 0 Then DrAdd("D", "Оформлены акты (документы и товар к отправке): " & n & " — " & s)
End Sub

' the lines written by hand in a list of «Отчёт дня» (column c, from row DR_ROW0)
Private Sub DrManual(sec As String, c As Integer)
    Dim sh As Object, d As Variant, i As Long, s As String
    If Not gDoc.Sheets.hasByName(SH_DAYREP) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_DAYREP)
    d = sh.getCellRangeByPosition(c, DR_ROW0, c, DR_ROW0 + DR_ROWS - 1).getDataArray()
    For i = 0 To UBound(d)
        s = Trim(CStr(d(i)(0)))
        If s <> "" Then DrAdd(sec, s)
    Next i
End Sub

' the issues not yet digitised: the working days after the last date of an issue in «Выдачи» up to the day of the report (the
' issues of the day itself are digitised the next working day)
Private Sub DrBacklog(dt As Double)
    Dim ish As Object, last As Long, dd As Variant, cc As Variant, i As Long, mx As Double, d As Double, s As String, n As Long
    On Error GoTo EH
    ish = WmsIssue.IssuesSheet()
    last = WmsOrders.LastRow(ish)
    If last < 1 Then Exit Sub
    dd = ish.getCellRangeByPosition(IC_DATE, 1, IC_DATE, last).getDataArray()
    cc = ish.getCellRangeByPosition(IC_CTL, 1, IC_CTL, last).getDataArray()
    For i = 0 To UBound(dd)
        If VarType(dd(i)(0)) = 5 Then
            If Left(CStr(cc(i)(0)), Len(ST_POSTED)) = ST_POSTED And dd(i)(0) > mx Then mx = Int(dd(i)(0))
        End If
    Next i
    If mx = 0 Or mx >= dt Then Exit Sub
    d = mx
    Do
        d = DrNextWork(d)
        If d > dt Then Exit Do
        s = s & IIf(s <> "", ", ", "") & Format(CDate(d), "DD.MM.YYYY") & " (" & DrDayName(d) & ")"
        n = n + 1
    Loop While n < 10
    If s <> "" Then DrAdd("T", "Оцифровка выдач за " & s & " (последняя оцифрованная дата выдачи — " & Format(CDate(mx), "DD.MM.YYYY") & ")")
    Exit Sub
EH:
End Sub

' what waits in the orders and in the stock: the positions expected, the goods below the minimum
Private Sub DrOpenWork()
    Dim res As String
    On Error Resume Next
    res = WmsSupply.OpenPositions("")
    If Left(res, 4) <> "ERR:" And gSpN > 0 Then DrAdd("T", "Принять поставки: ожидаются позиции заказов — " & gSpN)
    res = WmsSupply.NeedCompute("")
    If Left(res, 3) = "OK:" And gNdN > 0 Then DrAdd("T", "Товаров ниже минимума (нужно заказать): " & gNdN)
End Sub

' the registry: the EIs with a balance, without a place, waiting for a decision — into «В процессе» and the control
Private Sub DrStockState()
    Dim st As Object, last As Long, d As Variant, i As Long, nBal As Long, nNoPlace As Long, nReview As Long, nAll As Long, cEi As Long, cQty As Long, cPl As Long, cSt As Long
    On Error GoTo EH
    st = WmsIssue.StockSheet()
    last = WmsOrders.LastRow(st)
    If last < 1 Then Exit Sub
    d = st.getCellRangeByPosition(0, 1, SC_STATE, last).getDataArray()
    cEi = SC_EI
    cQty = SC_QTY
    cPl = SC_PLACE
    cSt = SC_STATE
    For i = 0 To UBound(d)
        If CStr(d(i)(cEi)) <> "" Then
            nAll = nAll + 1
            If VarType(d(i)(cQty)) = 5 Then
                If d(i)(cQty) > 0 Then
                    nBal = nBal + 1
                    If Trim(CStr(d(i)(cPl))) = "" Then nNoPlace = nNoPlace + 1
                End If
            End If
            If InStr(CStr(d(i)(cSt)), "разбора") > 0 Then nReview = nReview + 1
        End If
    Next i
    If nReview > 0 Then DrAdd("P", "ЕИ, требующих разбора: " & nReview)
    If nNoPlace > 0 Then DrAdd("P", "ЕИ с остатком без указанного места хранения: " & nNoPlace)
    DrAdd("C", "ЕИ в учёте: " & nAll & ", из них с остатком — " & nBal)
    Exit Sub
EH:
End Sub

' ================================================================ test seam (reads and closes the document built last)

Function DcDocText() As String
    DcDocText = gDcDoc.Text.getString()
End Function

' the cells of the first table of the document, row by row: «a|b|c» lines
Function DcDocTable() As String
    Dim tbl As Object, d As Variant, i As Long, j As Long, s As String, ln As String
    tbl = gDcDoc.getTextTables().getByIndex(0)
    d = tbl.getDataArray()
    For i = 0 To UBound(d)
        ln = ""
        For j = 0 To UBound(d(i))
            ln = ln & IIf(j > 0, "|", "") & CStr(d(i)(j))
        Next j
        s = s & IIf(i > 0, Chr(10), "") & ln
    Next i
    DcDocTable = s
End Function

Sub DcDocClose()
    On Error Resume Next
    gDcDoc.close(True)
    gDcDoc = Nothing
End Sub
