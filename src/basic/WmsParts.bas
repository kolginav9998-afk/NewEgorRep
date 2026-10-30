' WmsParts — 1.0.2: детали с серийным номером. Артикул у них постоянный, серийные номера разные: из ЕИ в N деталей получается
' N отдельных ЕИ (по одной детали) — тот же артикул, своё место, свой серийный номер, своя этикетка («Арт.: … · S/N: …»).
' Это не новая операция учёта, а цепочка прежних: «Разделение» (SPLIT) и «Свойства ЕИ» (EI_PROP) — журнал, проверки и
' восстановление те же, что у каждой из них. Library WmsPlus.
Option Explicit

' the serials asked in one window, the most one EI is taken apart into at a time
Private Const SRL_PAGE = 18
Private Const SRL_MAX = 200

Sub BtnSerials(Optional oEvent As Variant)
    Dim eis As Variant, s As String
    WmsInit()
    eis = WmsItem.CursorEIs()
    If UBound(eis) < 0 Then
        WmsUi.UiMessage("Встаньте на строку ЕИ («Наличие», «Заказы», «Иной приход») или откройте его карточку на «Поиске» и нажмите «Серийные…»: " _
            & "WMS разделит ЕИ на отдельные детали — каждая со своим серийным номером и этикеткой.")
        Exit Sub
    End If
    s = SerialWindow(CLng(eis(0)))
    If Left(s, 4) = "ERR:" Then WmsUi.UiMessage(Mid(s, 5))
End Sub

' the window of the serial numbers of EI n: a field for every part (page by page), «Записать» → SerialSplit, then the labels
Function SerialWindow(n As Long) As String
    Dim d As Variant, nParts As Long, sers() As String, p0 As Long, p1 As Long, i As Long, nF As Long, labels() As String, v() As String, lists() As String
    Dim msg As String, canon As String, p As Variant, res As String, info As String, k As Integer, eis As Variant, counts() As Long, a As Variant
    d = WmsIssue.StockData(n)
    canon = CStr(d(SC_EI))
    If Abs(CDbl(d(SC_QTY)) - Int(CDbl(d(SC_QTY)) + 0.5)) > 0.0000001 Or CDbl(d(SC_QTY)) < 1 Then
        SerialWindow = "ERR:остаток " & canon & " — " & WmsIssue.QtyText(CDbl(d(SC_QTY))) & " " & d(SC_UNIT) & ": серийные номера у целых деталей (1, 2, 3 …)"
        gUiLastMsg = SerialWindow
        Exit Function
    End If
    nParts = CLng(CDbl(d(SC_QTY)) + 0.1)
    If nParts > SRL_MAX Then
        SerialWindow = "ERR:деталей в " & canon & " — " & nParts & ", за один раз не больше " & SRL_MAX & ". Сначала отделите часть («Разделить ЕИ…»)."
        gUiLastMsg = SerialWindow
        Exit Function
    End If
    ReDim sers(nParts - 1)
    p0 = 0
    Do While p0 < nParts
        p1 = p0 + SRL_PAGE - 1
        If p1 > nParts - 1 Then p1 = nParts - 1
        nF = p1 - p0 + 1
        ReDim labels(nF - 1)
        ReDim v(nF - 1)
        ReDim lists(nF - 1)
        For i = p0 To p1
            labels(i - p0) = "Деталь " & (i + 1) & " — серийный №"
        Next i
        info = canon & " · " & d(SC_NAME) & IIf(Trim(CStr(d(SC_ART))) <> "", " (" & d(SC_ART) & ")", "") & " · " & nParts & " " & d(SC_UNIT) _
            & IIf(nParts > SRL_PAGE, ". Страница " & (p0 \ SRL_PAGE + 1) & " из " & ((nParts + SRL_PAGE - 1) \ SRL_PAGE), "") _
            & ". Впишите серийный номер каждой детали; пустое — деталь останется в общем ЕИ без номера. Каждая деталь с номером станет своим ЕИ " _
            & "(тот же артикул и место) со своей этикеткой."
        If Not WmsPrimeUi.UiForm("Серийные номера", info, labels, v, lists, IIf(p1 >= nParts - 1, "Записать", "Дальше")) Then
            SerialWindow = "SKIP:окно закрыто"
            gUiLastMsg = SerialWindow
            Exit Function
        End If
        For i = p0 To p1
            sers(i) = Trim(v(i - p0))
        Next i
        p0 = p1 + 1
    Loop
    res = SerialSplit(canon, sers)
    gUiLastMsg = res
    If Left(res, 3) <> "OK:" Then
        SerialWindow = res
        Exit Function
    End If
    ' OK:<how many>|ЕИ-…,ЕИ-…
    a = Split(Mid(res, 4), "|")
    k = WmsPrimeUi.UiButtons("Серийные детали", "Готово: ЕИ с серийными номерами — " & a(0) & ". Напечатать этикетки (номер ЕИ, артикул, серийный номер, штрихкод)?", _
        Array("Напечатать этикетки", "Не сейчас"), Array(1, 0))
    SerialWindow = res
    If k <> 1 Then Exit Function
    a = Split(CStr(a(1)), ",")
    ReDim eis(UBound(a))
    ReDim counts(UBound(a))
    For i = 0 To UBound(a)
        eis(i) = WmsView.EiNo(CStr(a(i)), CLng(SysNum(SK_NEXT_EI)))
        counts(i) = 1
    Next i
    msg = WmsItem.LabelPrintN(eis, counts, WmsItem.LabelSize())
    If Left(msg, 4) = "ERR:" Then WmsUi.UiMessage(Mid(msg, 5))
End Function

' EI sEI taken apart by the serial numbers sers (an array, an empty one — a part that stays): a part with a number becomes an EI of
' its own (SPLIT of one, then EI_PROP with the number; the «годен до» and the note of the EI go along); when every part has
' a number the last one stays in the EI itself and gets its number. A number written twice, or one that another EI has — refused
' before anything is posted. OK:<EIs with a number>|ЕИ-…,ЕИ-… / ERR:…
Function SerialSplit(sEI As String, sers As Variant) As String
    Dim n As Long, canon As String, msg As String, d As Variant, nParts As Long, k As Long, i As Long, j As Long, other As Long, why As String
    Dim p As Variant, sExp As String, sNote As String, place As String, last As Long, done As Long, r As String, newEI As String, eis As String, ser As String
    Dim pu As String, pn As Double, sErrTxt As String, nErrLine As Long
    WmsInit()
    On Error GoTo EH
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        SerialSplit = "ERR:" & msg
        Exit Function
    End If
    why = WmsIssue.StockRowProblem(n, canon)
    If why <> "" Then
        SerialSplit = "ERR:" & why
        Exit Function
    End If
    d = WmsIssue.StockData(n)
    If Abs(CDbl(d(SC_QTY)) - Int(CDbl(d(SC_QTY)) + 0.5)) > 0.0000001 Or CDbl(d(SC_QTY)) < 1 Then
        SerialSplit = "ERR:остаток " & canon & " — " & WmsIssue.QtyText(CDbl(d(SC_QTY))) & ": серийные номера у целых деталей"
        Exit Function
    End If
    nParts = CLng(CDbl(d(SC_QTY)) + 0.1)
    If UBound(sers) + 1 > nParts Then
        SerialSplit = "ERR:серийных номеров " & (UBound(sers) + 1) & ", а деталей в " & canon & " — " & nParts
        Exit Function
    End If
    If WmsPack.PackOf(n, pu, pn) Then
        SerialSplit = "ERR:у " & canon & " есть фасовка — серийные номера записывают штучным деталям; снимите фасовку («Фасовка ЕИ»)"
        Exit Function
    End If
    For i = 0 To UBound(sers)
        ser = Trim(CStr(sers(i)))
        If ser <> "" Then
            If Len(ser) > 60 Then
                SerialSplit = "ERR:серийный номер «" & Left(ser, 20) & "…» длиннее 60 знаков"
                Exit Function
            End If
            k = k + 1
            last = i
            For j = 0 To i - 1
                If LCase(Trim(CStr(sers(j)))) = LCase(ser) Then
                    SerialSplit = "ERR:серийный номер «" & ser & "» записан дважды (детали " & (j + 1) & " и " & (i + 1) & ")"
                    Exit Function
                End If
            Next j
            other = WmsItem.SerialEI(ser)
            If other > 0 Then
                SerialSplit = "ERR:серийный номер «" & ser & "» уже записан у " & WmsIssue.EiCanon(other) & " — проверьте номер"
                Exit Function
            End If
        End If
    Next i
    If k = 0 Then
        SerialSplit = "ERR:не указан ни один серийный номер"
        Exit Function
    End If
    why = PostingBlockReason()
    If why <> "" Then
        SerialSplit = why
        Exit Function
    End If
    p = WmsItem.PropOf(n)
    sExp = IIf(p(1) > 0, Format(p(1), "DD.MM.YYYY"), "")
    sNote = CStr(p(2))
    place = CStr(d(SC_PLACE))
    WmsView.ProgressStart("Серийные детали " & canon, k)
    For i = 0 To UBound(sers)
        ser = Trim(CStr(sers(i)))
        If ser <> "" Then
            If k = nParts And i = last Then
                ' every part has a number: the last one is the EI itself
                r = WmsItem.PropSet(canon, ser, sExp, sNote)
                newEI = canon
            Else
                r = WmsDaily.SplitPost(canon, "1", place, "серийная деталь " & ser, True)
                If Left(r, 3) <> "OK:" Then
                    WmsView.ProgressEnd()
                    SerialSplit = "ERR:остановлено на детали " & (i + 1) & " («" & ser & "»): " & Mid(r, InStr(r, ":") + 1) & ". Готово: " & done & " из " & k & _
                        IIf(eis <> "", " (" & eis & ")", "") & ". Уберите готовые номера и повторите для остальных."
                    Exit Function
                End If
                newEI = Mid(r, 4)
                newEI = Left(newEI, InStr(newEI, "|") - 1)
                r = WmsItem.PropSet(newEI, ser, sExp, sNote)
            End If
            If Left(r, 3) <> "OK:" Then
                WmsView.ProgressEnd()
                SerialSplit = "ERR:номер «" & ser & "» не записан у " & newEI & ": " & Mid(r, InStr(r, ":") + 1) & ". Готово: " & done & " из " & k & _
                    IIf(eis <> "", " (" & eis & ")", "") & ". " & newEI & " уже отделён — впишите его номер через «Свойства ЕИ…»."
                Exit Function
            End If
            done = done + 1
            eis = eis & IIf(eis <> "", ",", "") & newEI
            WmsView.ProgressSet(done)
        End If
    Next i
    WmsView.ProgressEnd()
    SerialSplit = "OK:" & done & "|" & eis
    gUiLastMsg = SerialSplit
    Exit Function
EH:
    sErrTxt = Error$
    nErrLine = Erl
    On Error Resume Next
    WmsView.ProgressEnd()
    SerialSplit = "ERR-SYS:серийные детали не разделены: " & sErrTxt & " (строка " & nErrLine & "). Готово: " & done & " из " & k & IIf(eis <> "", " (" & eis & ")", "")
End Function
