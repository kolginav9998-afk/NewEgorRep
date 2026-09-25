' WmsConfig — constants, _SYS layout, settings, LibreOffice settings management (AutoInput).
' MASTER SPEC v0.3: §3 sources as text, §10/§29 AutoInput, §16 protection, §20 journal, §22 lock, §24 chunks, §26 backups.
Option Explicit

Public Const WMS_CORE_VERSION = "0.1.0-phase1"
Public Const WMS_SYS_SCHEMA = "WMS-SYS-1"

Public Const SYS_SHEET = "_SYS"
Public Const JOURNAL_DIR = "WMS_Journal"
Public Const BACKUP_DIR = "WMS_Backups"
Public Const LOCK_FILE = "wms.lock"
Public Const JOURNAL_PREFIX = "WMS_journal_"
Public Const JOURNAL_EXT = ".csv"
Public Const AUTOINPUT_STATE_FILE = "wms_autoinput.state"

' Sheet protection guards against accidents, not against a determined user (spec §16: the password is not used in daily work).
Public Const PROTECT_PWD = "wms"

' _SYS: column A = key, column B = value; row index (0-based) = SK_* constant. Layout is checked at every start.
Public Const SK_SCHEMA = 0
Public Const SK_INSTANCE = 1
Public Const SK_MODE = 2
Public Const SK_CORE_VERSION = 3
Public Const SK_LAST_SEQ = 4
Public Const SK_NEXT_EI = 5
Public Const SK_NEXT_NO = 6
Public Const SK_NEXT_RET = 7
Public Const SK_JPOS = 8
Public Const SK_REG_URL = 9
Public Const SK_TX_STATE = 10
Public Const SK_TX_SEQ = 11
Public Const SK_TX_TYPE = 12
Public Const SK_TX_TIME = 13
Public Const SK_TX_BI = 14
Public Const SK_SAVE_STAMP = 15
Public Const SK_SAVE_SEQ = 16
Public Const SK_MAX_QTY = 17
Public Const SK_KEY_SHEETS = 18
Public Const SYS_ROWS = 19

' transaction marker states
Public Const TX_NONE = "NONE"
Public Const TX_STARTED = "STARTED"
Public Const TX_COMMITTED = "COMMITTED"

' settings
Public Const BACKUP_KEEP_DAILY = 14
Public Const BACKUP_KEEP_MANUAL = 10
Public Const BACKUP_KEEP_UPGRADE = 5
Public Const BACKUP_KEEP_MIGRATION = 5
Public Const CHECK_CHUNK_ROWS = 20000
Public Const JPOS_VERIFY_WINDOW = 4096
Public Const QTY_MAX_DECIMALS = 3
Public Const QTY_MAX_INT_DIGITS = 9

Function SysKeyNames() As Variant
    SysKeyNames = Array("SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET", _
        "JOURNAL_POS", "REGISTERED_URL", "TX_STATE", "TX_SEQ", "TX_TYPE", "TX_TIME", "TX_BEFORE_IMAGE", _
        "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS")
End Function

' keys of _SYS that must hold numbers
Function SysNumericKeys() As Variant
    SysNumericKeys = Array(SK_LAST_SEQ, SK_NEXT_EI, SK_NEXT_NO, SK_NEXT_RET, SK_TX_SEQ, SK_SAVE_STAMP, SK_SAVE_SEQ, SK_MAX_QTY)
End Function

' ---------------------------------------------------------------- AutoInput (spec §10, §29; decision D-031)
' The setting belongs to the LibreOffice user profile and is global: it affects every Calc document of this profile
' while WMS runs. The original value is remembered in a small file inside the profile (not in the book), so that it
' survives a crash or an unsaved session and is restored at the next clean close of WMS.

Private Function CalcInputAccess(bUpdate As Boolean) As Object
    Dim cp As Object, a(0) As New com.sun.star.beans.PropertyValue
    cp = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
    a(0).Name = "nodepath"
    a(0).Value = "/org.openoffice.Office.Calc/Input"
    If bUpdate Then
        CalcInputAccess = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", a())
    Else
        CalcInputAccess = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", a())
    End If
End Function

Function AutoInputGet() As Boolean
    AutoInputGet = CalcInputAccess(False).getByName("AutoInput")
End Function

Sub AutoInputSet(bOn As Boolean)
    Dim ua As Object
    ua = CalcInputAccess(True)
    ua.replaceByName("AutoInput", bOn)
    ua.commitChanges()
End Sub

Private Function AutoInputStateUrl() As String
    AutoInputStateUrl = CreateUnoService("com.sun.star.util.PathSettings").UserConfig & "/" & AUTOINPUT_STATE_FILE
End Function

' Called at WMS start. Returns a short note for the startup report.
Function AutoInputDisable() As String
    Dim sfa As Object, url As String, wasOn As Boolean
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    url = AutoInputStateUrl()
    wasOn = AutoInputGet()
    If Not sfa.exists(url) Then
        ' remember the user's own value only once: after a crash the file still holds the original
        WmsCore.WriteTextFile(url, "original=" & IIf(wasOn, "1", "0") & Chr(10))
    End If
    If wasOn Then AutoInputSet(False)
    AutoInputDisable = "автоввод Calc выключен на время работы WMS (настройка общая для LibreOffice)"
    Exit Function
EH:
    AutoInputDisable = "ВНИМАНИЕ: не удалось выключить автоввод Calc: " & Error$
End Function

' Called at clean close. Restores the remembered value and forgets it.
Function AutoInputRestore() As String
    Dim sfa As Object, url As String, s As String
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    url = AutoInputStateUrl()
    If Not sfa.exists(url) Then
        AutoInputRestore = ""
        Exit Function
    End If
    s = WmsCore.ReadTextFile(url)
    If InStr(s, "original=1") > 0 Then AutoInputSet(True)
    If InStr(s, "original=0") > 0 Then AutoInputSet(False)
    sfa.kill(url)
    AutoInputRestore = "автоввод Calc восстановлен"
    Exit Function
EH:
    AutoInputRestore = "ВНИМАНИЕ: не удалось восстановить автоввод Calc: " & Error$
End Function
