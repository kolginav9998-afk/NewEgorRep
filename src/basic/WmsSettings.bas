' WmsSettings — 1.1: the settings of the storekeeper, the sheet «Настройки»: a table «что настроить — значение — пояснение» (the code of a setting is in the hidden column D).
' The sheet is his own (not protected, not journaled): WMS only reads it, and when a value is empty or is not a number it uses what is written here as the default. Library WmsPlus.
Option Explicit

Public Const SH_SET = "Настройки"
Public Const ST_CODE_COL = 3
Public Const ST_VAL_COL = 1
Public Const ST_ROWS = 40

' the codes of the settings
Public Const SET_BIG_QTY = "BIG_QTY"

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
