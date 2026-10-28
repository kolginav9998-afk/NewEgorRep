' WmsAttach — 1.1: a scan or a photo of the document of a receipt (УПД, накладная) kept beside the book, if the storekeeper wants it. Optional: a receipt is posted without it.
' The file is copied (the original stays where it is) into the folder «WMS_Документы/<заказ>» next to the book; nothing is written in the registry or the journal (no schema change),
' the link is the name of the folder. «Документ…» on «Заказы»: add a file / open the folder of the order. Library WmsPlus.
Option Explicit

Public Const ATTACH_DIR = "WMS_Документы"
' test seam: the file "chosen" in the window of files while a test answers the windows (gUiAuto)
Global gAtTestSrc As String

Sub TestSetAttachSrc(sUrl As String)
    gAtTestSrc = sUrl
End Sub

' the key of the folder for the row of «Заказы» row r: «Заказ <№>», else «Документ <№ документа>»; "" — the row has neither
Function AttachKey(r As Long) As String
    Dim sh As Object, no As String, dc As String
    sh = gDoc.Sheets.getByName(SH_ORDERS)
    no = Trim(sh.getCellByPosition(OC_ORDER, r).getString())
    dc = Trim(sh.getCellByPosition(OC_DOC, r).getString())
    If no <> "" Then
        AttachKey = "Заказ " & WmsPrint.FileSafe(no)
    ElseIf dc <> "" Then
        AttachKey = "Документ " & WmsPrint.FileSafe(dc)
    End If
End Function

' the folder of the key (the URL with the closing slash); created when bMake
Function AttachFolder(sKey As String, bMake As Boolean) As String
    Dim f As String
    f = gBaseDir & ATTACH_DIR & "/"
    If bMake And Not gSfa.exists(f) Then gSfa.createFolder(f)
    f = f & sKey & "/"
    If bMake And Not gSfa.exists(f) Then gSfa.createFolder(f)
    AttachFolder = f
End Function

' the name of a file in the folder that is not taken yet: «scan.jpg» → «scan (2).jpg»
Private Function FreeName(sFolder As String, sName As String) As String
    Dim p As Long, base0 As String, ext As String, n As Long, c As String
    c = sName
    p = Len(sName)
    Do While p > 0
        If Mid(sName, p, 1) = "." Then Exit Do
        p = p - 1
    Loop
    If p > 1 Then
        base0 = Left(sName, p - 1)
        ext = Mid(sName, p)
    Else
        base0 = sName
    End If
    n = 1
    Do While gSfa.exists(sFolder & c)
        n = n + 1
        c = base0 & " (" & n & ")" & ext
    Loop
    FreeName = c
End Function

' the file at sSrc (a URL) copied into the folder of the key: OK:<name>|<path> / ERR:…
Function AttachAdd(sKey As String, sSrc As String) As String
    Dim f As String, nm As String, p As Long, dst As String, nSrc As Double, nDst As Double
    On Error GoTo EH
    If Not gDoc.hasLocation() Then
        AttachAdd = "ERR:книга не сохранена в файл — папку для документов не найти"
        Exit Function
    End If
    If sKey = "" Then
        AttachAdd = "ERR:у строки нет ни номера заказа (A), ни номера документа (C) — документу не к чему привязаться"
        Exit Function
    End If
    If Not gSfa.exists(sSrc) Then
        AttachAdd = "ERR:файл не найден: " & ConvertFromURL(sSrc)
        Exit Function
    End If
    nm = sSrc
    p = Len(nm)
    Do While p > 0
        If Mid(nm, p, 1) = "/" Then Exit Do
        p = p - 1
    Loop
    nm = ConvertFromURL(Mid(nm, p + 1))
    p = Len(nm)
    Do While p > 0
        If Mid(nm, p, 1) = "\" Or Mid(nm, p, 1) = "/" Then Exit Do
        p = p - 1
    Loop
    nm = Mid(nm, p + 1)
    f = AttachFolder(sKey, True)
    nm = FreeName(f, nm)
    dst = f & nm
    gSfa.copy(sSrc, dst)
    nSrc = gSfa.getSize(sSrc)
    nDst = gSfa.getSize(dst)
    If nSrc <> nDst Then
        gSfa.kill(dst)
        AttachAdd = "ERR:копия получилась неполной (" & nDst & " из " & nSrc & " байт) и удалена"
        Exit Function
    End If
    AttachAdd = "OK:" & nm & "|" & ConvertFromURL(dst)
    gUiLastMsg = AttachAdd
    Exit Function
EH:
    AttachAdd = "ERR:документ не добавлен: " & Error$
End Function

' the files of the folder of the key, one per line ("" — none)
Function AttachList(sKey As String) As String
    Dim f As String, a As Variant, i As Long, nm As String, p As Long
    If sKey = "" Or Not gDoc.hasLocation() Then Exit Function
    f = AttachFolder(sKey, False)
    If Not gSfa.exists(f) Then Exit Function
    a = gSfa.getFolderContents(f, False)
    For i = 0 To UBound(a)
        nm = a(i)
        p = Len(nm)
        Do While p > 0
            If Mid(nm, p, 1) = "/" Then Exit Do
            p = p - 1
        Loop
        AttachList = AttachList & IIf(AttachList <> "", Chr(10), "") & ConvertFromURL(Mid(nm, p + 1))
    Next i
End Function

' the button «Документ…» of «Заказы»: for the row under the cursor — add a scan or a photo, or open the folder of the order
Sub BtnAttach(Optional oEvent As Variant)
    Dim sh As Object, r As Long, sKey As String, files As String, k As Integer, fp As Object, a As Variant, res As String, i As Long, sDir As String
    WmsInit()
    On Error GoTo EH
    If gDoc.getCurrentController().getActiveSheet().Name <> SH_ORDERS Then
        WmsUi.UiMessage("«Документ…» работает на листе «" & SH_ORDERS & "»: встаньте на строку прихода и нажмите кнопку.")
        Exit Sub
    End If
    r = gDoc.getCurrentController().getSelection().getRangeAddress().StartRow
    If r < 1 Then
        WmsUi.UiMessage("Встаньте на строку прихода (не на заголовок) и нажмите «Документ…».")
        Exit Sub
    End If
    sKey = AttachKey(r)
    If sKey = "" Then
        WmsUi.UiMessage("У строки нет ни номера заказа (A), ни номера документа (C) — документу не к чему привязаться.")
        Exit Sub
    End If
    files = AttachList(sKey)
    k = WmsPrimeUi.UiButtons("Документ прихода — " & sKey, IIf(files = "", "Документов пока нет.", "Уже лежит:" & Chr(10) & files) & Chr(10) & Chr(10) _
        & "Скан или фото накладной можно положить рядом с книгой — необязательно: приход проводится и без него. Файл копируется, оригинал остаётся на месте.", _
        Array("Добавить файл…", "Открыть папку", "Закрыть"), Array(1, 2, 0))
    If k = 1 Then
        If gUiAuto <> 0 Then
            a = Array(gAtTestSrc)
        Else
            fp = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
            fp.setTitle("Скан или фото документа")
            fp.setMultiSelectionMode(True)
            If fp.execute() <> 1 Then Exit Sub
            a = fp.getFiles()
        End If
        If UBound(a) = 0 Then
            res = AttachAdd(sKey, a(0))
            If Left(res, 3) <> "OK:" Then WmsUi.UiMessage(Mid(res, 5))
        Else
            sDir = a(0)
            If Right(sDir, 1) <> "/" Then sDir = sDir & "/"
            For i = 1 To UBound(a)
                res = AttachAdd(sKey, sDir & a(i))
                If Left(res, 3) <> "OK:" Then
                    WmsUi.UiMessage(Mid(res, 5))
                    Exit For
                End If
            Next i
        End If
        If Left(res, 3) = "OK:" Then WmsUi.UiMessage("Документ добавлен в папку «" & ATTACH_DIR & "/" & sKey & "»:" & Chr(10) & AttachList(sKey))
    ElseIf k = 2 Then
        If files = "" Then
            WmsUi.UiMessage("Папки документов этого прихода ещё нет — сначала «Добавить файл…».")
        ElseIf gUiAuto = 0 Then
            CreateUnoService("com.sun.star.system.SystemShellExecute").execute(AttachFolder(sKey, False), "", 0)
        End If
    End If
    Exit Sub
EH:
    WmsUi.UiMessage("Документ не добавлен: " & Error$)
End Sub
