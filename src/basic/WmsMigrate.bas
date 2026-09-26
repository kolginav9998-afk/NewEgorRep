' WmsMigrate — перенос существующих ЕИ старого склада в книгу WMS (задание «FINAL WMS MARATHON», §4; D-056, D-057, D-087).
' Одна операция MIGRATE на строку таблицы переноса («один физический остаток — одна строка»): ЕИ с его прежним
' номером (без перенумерации) → строка «Наличие» этого номера, индекс деталей _ART, NEXT_EI выше номера. Журнал хранит
' происхождение: файл таблицы, его SHA-256, строку таблицы, старую маркировку. Вызывается инструментом
' tools/migrate.py (режим MIGRATE COPY) на копии книги, только после чистого dry-run. У переноса нет исправления и
' сторно: ошибка переноса исправляется повтором на свежей копии до PROMOTE, после PROMOTE — обычными операциями
' (корректировки, перемещение).
Option Explicit

' «Тип источника» of a migrated EI («Наличие».J): the types of Phase 5 and the ordinary supplier
Private Function MigrateStype(ByVal s As String) As String
    Dim names As Variant, i As Integer
    names = Split(STYPE_SUPPLIER & "|" & SRC_STYPES, "|")
    For i = 0 To UBound(names)
        If LCase(Trim(s)) = LCase(names(i)) Then
            MigrateStype = names(i)
            Exit Function
        End If
    Next i
    MigrateStype = ""
End Function

' Posts one migrated EI. OK:<seq>, ERR:<why> (nothing written), BLOCKED:… / ERR-RB / ERR-CRITICAL from ApplyOperation.
' sQty: "0" allowed (an EI kept with a zero balance), otherwise the locale-independent quantity; sOrigin: «<file>|<sha256>|<row>».
Function MigrateEI(sEI As String, sName As String, sArt As String, sUnit As String, sQty As String, sPlace As String, sCat As String, _
    sType As String, sMark As String, sOrigin As String) As String
    Dim n As Long, canon As String, msg As String, q As Double, why As String, srcType As String, key As String, st As Integer
    Dim en As Long, ecanon As String, aRow As Long, sh As Object, sState As String, sSrc As String
    WmsInit()
    On Error GoTo EH
    why = PostingBlockReason()
    If why <> "" Then
        MigrateEI = why
        Exit Function
    End If
    If Not WmsIssue.NormalizeEI(sEI, n, canon, msg) Then
        MigrateEI = "ERR:" & msg
        Exit Function
    End If
    If n < 1 Or n >= MAX_SHEET_ROW Then
        MigrateEI = "ERR:номер ЕИ " & n & " вне допустимого диапазона"
        Exit Function
    End If
    sh = WmsIssue.StockSheet()
    If sh.getCellByPosition(SC_EI, n).getType() <> com.sun.star.table.CellContentType.EMPTY Then
        MigrateEI = "ERR:" & canon & " уже есть в «" & SH_STOCK & "» — перенос не перенумеровывает и не сливает ЕИ"
        Exit Function
    End If
    If Trim(sName) = "" Or Trim(sUnit) = "" Then
        MigrateEI = "ERR:" & canon & ": нет наименования или единицы"
        Exit Function
    End If
    ' a zero balance is allowed (the EI is kept with its number); a negative or malformed quantity is refused
    st = ParseQtyText(sQty, q, msg)
    If st = 5 And q = 0 And InStr(sQty, "-") = 0 Then
        q = 0
    ElseIf st <> 0 Then
        MigrateEI = "ERR:" & canon & ": количество «" & sQty & "» — " & msg
        Exit Function
    End If
    srcType = MigrateStype(sType)
    If srcType = "" Then
        MigrateEI = "ERR:" & canon & ": тип источника «" & sType & "» неизвестен (D-087: пустой тип при переносе не допускается)"
        Exit Function
    End If
    If srcType = SRC_STYPES_PART Then
        key = WmsSpecial.ArticleKey(sArt)
        If key = "" Then
            MigrateEI = "ERR:" & canon & ": у детали нет артикула"
            Exit Function
        End If
        st = WmsSpecial.ArtLookup(key, en, ecanon, aRow, why)
        If st <> 0 Then
            MigrateEI = "ERR:" & canon & ": артикул «" & sArt & "» уже " & IIf(st = 1, "у " & ecanon, "в индексе: " & why) _
                & " — один артикул детали — один ЕИ; конфликт решается до переноса"
            Exit Function
        End If
        aRow = WmsOrders.LastRow(WmsSpecial.ArtSheet()) + 1
    End If
    sState = IIf(srcType = "Иной приход", EI_ST_REVIEW, EI_ST_ACTIVE)
    sSrc = "Перенос" & IIf(Trim(sMark) <> "", " (" & Trim(sMark) & ")", "")
    PlanBegin("MIGRATE", canon)
    PlanField("EI", canon)
    PlanField("NAME", sName)
    PlanField("ART", sArt)
    PlanField("UNIT", sUnit)
    PlanField("QTY", q)
    PlanField("PLACE", sPlace)
    PlanField("CAT", sCat)
    PlanField("STYPE", srcType)
    PlanField("MARK", sMark)
    PlanField("ORIGIN", sOrigin)
    If n >= SysNum(SK_NEXT_EI) Then PlanSetValue(SYS_SHEET, SK_NEXT_EI, 1, n + 1, False)
    WmsSpecial.PlanNewStockRow(n, canon, q, sName, sArt, sUnit, sPlace, sCat, sState, sSrc, srcType)
    If srcType = SRC_STYPES_PART Then
        PlanSetValue(SH_ART, aRow, AR_KEY, key, False)
        PlanSetValue(SH_ART, aRow, AR_EI, canon, False)
        PlanSetValue(SH_ART, aRow, AR_ART, sArt, False)
    End If
    MigrateEI = ApplyOperation(0, 0)
    Exit Function
EH:
    MigrateEI = "ERR:внутренняя ошибка проверки переноса: " & Error$ & " (код " & Err & ", строка " & Erl & ")"
End Function
