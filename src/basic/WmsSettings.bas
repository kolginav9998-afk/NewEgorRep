' WmsSettings — 1.1: the settings of the storekeeper, the sheet «Настройки»: a table «что настроить — значение — пояснение» (the code of a setting is in the hidden column D).
' The sheet is his own (not protected, not journaled): WMS only reads it, and when a value is empty or is not a number it uses what is written here as the default. Library WmsPlus.
Option Explicit

Public Const SH_SET = "Настройки"
Private Const ST_CODE_COL = 3
Private Const ST_VAL_COL = 1
' 1.4.0: 120 rows (it was 40; the same number is SETTINGS_MAX of tools/build_ods.py)
Private Const ST_ROWS = 120

' 1.4.0: the switches of the helpers — the rows «да» / «нет» of the sheet. The cache is read from the sheet once and reset when the sheet changes (OnSettingsChange) or the book changes
Global gFeatOk As Boolean
Global gFeatUrl As String
Global gFeatN As Long
Global gFeatCodes As Variant
Global gFeatVals As Variant
Global gFeatLite As Boolean

' the codes of the settings
Public Const SET_BIG_QTY = "BIG_QTY"
' 1.3.0: who gets the article on the label of an EI: «детали» (parts, an EI with a serial number), «все» (every EI that has an article), «нет» (nobody)
Public Const SET_LABEL_ART = "LABEL_ART"
' 1.4.0: «Лёгкий режим» — every helper off at once (the other codes of the switches are strings: FeatOn("NAME_FIX"), … — a constant costs a lookup on every call, D-118)
Public Const SET_LITE = "LITE"

' the text of a setting; the default when the sheet, the code or the value is not there
Function SettingText(sCode As String, sDefault As String) As String
    Dim sh As Object, d As Variant, i As Long
    SettingText = sDefault
    If Not gDoc.Sheets.hasByName(SH_SET) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_SET)
    d = sh.getCellRangeByPosition(0, 1, ST_CODE_COL, ST_ROWS).getDataArray()
    For i = 0 To UBound(d)
        If CStr(d(i)(ST_CODE_COL)) = sCode Then
            If Trim(CStr(d(i)(ST_VAL_COL))) <> "" Then SettingText = Trim(CStr(d(i)(ST_VAL_COL)))
            Exit Function
        End If
    Next i
End Function

' the number of a setting (a comma or a point as the decimal sign); the default when the value is empty or is not a number
Function SettingNum(sCode As String, dDefault As Double) As Double
    Dim s As String, sh As Object, d As Variant, i As Long
    SettingNum = dDefault
    If Not gDoc.Sheets.hasByName(SH_SET) Then Exit Function
    sh = gDoc.Sheets.getByName(SH_SET)
    d = sh.getCellRangeByPosition(0, 1, ST_CODE_COL, ST_ROWS).getDataArray()
    For i = 0 To UBound(d)
        If CStr(d(i)(ST_CODE_COL)) = sCode Then
            If VarType(d(i)(ST_VAL_COL)) = 5 Then
                SettingNum = d(i)(ST_VAL_COL)
            Else
                s = Replace(Trim(CStr(d(i)(ST_VAL_COL))), ",", ".")
                If s <> "" And IsNumeric(s) Then SettingNum = Val(s)
            End If
            Exit Function
        End If
    Next i
End Function

' ================================================================ 1.4.0 the switches of the helpers
' FeatOn("NAME_FIX") — True when the helper works. A helper is a convenience (a hint, an automatic correction, a list that rebuilds itself when the sheet is opened): switched off, the data stay right —
' the hint is simply not given, the list is rebuilt by its button. There is no switch of the posting, the journal, the checks of integrity, the copies: they do not ask this function.
' «Лёгкий режим» = «да» turns every helper off at once. A code that the sheet does not have (an old book, a deleted row) is a helper that works.
Function FeatOn(sCode As String) As Boolean
    Dim i As Long
    FeatOn = True
    On Error GoTo EH
    If Not gFeatOk Then
        FeatLoad()
    ElseIf gFeatUrl <> gDoc.getURL() Then
        FeatLoad()
    End If
    If gFeatLite Then
        FeatOn = False
        Exit Function
    End If
    For i = 0 To gFeatN - 1
        If gFeatCodes(i) = sCode Then
            FeatOn = (gFeatVals(i) <> "нет")
            Exit Function
        End If
    Next i
    Exit Function
EH:
    FeatOn = True
End Function

' the cache: the codes and the values (lower case) of the sheet, once
Private Sub FeatLoad()
    Dim sh As Object, d As Variant, i As Long, n As Long, sCodes(ST_ROWS) As String, sVals(ST_ROWS) As String, s As String
    gFeatLite = False
    gFeatN = 0
    gFeatOk = True
    gFeatUrl = gDoc.getURL()
    If Not gDoc.Sheets.hasByName(SH_SET) Then Exit Sub
    sh = gDoc.Sheets.getByName(SH_SET)
    d = sh.getCellRangeByPosition(0, 1, ST_CODE_COL, ST_ROWS).getDataArray()
    For i = 0 To UBound(d)
        s = Trim(CStr(d(i)(ST_CODE_COL)))
        If s <> "" And n < ST_ROWS Then
            sCodes(n) = s
            sVals(n) = LCase(Trim(CStr(d(i)(ST_VAL_COL))))
            If s = SET_LITE Then gFeatLite = (sVals(n) = "да")
            n = n + 1
        End If
    Next i
    gFeatCodes = sCodes
    gFeatVals = sVals
    gFeatN = n
End Sub

' the event of the sheet «Настройки» (a value was changed): the cache is read again at the next question
Sub OnSettingsChange(Optional oTarget As Variant)
    gFeatOk = False
End Sub
