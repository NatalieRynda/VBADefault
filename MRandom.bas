Attribute VB_Name = "MRandom"
Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ApiShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ApiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Private Const CF_ANSIONLY = &H400&
Private Const CF_APPLY = &H200&
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8
Private Const CF_DIF = 5
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPENHMETAFILE = &H8E
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPTEXT = &H81
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_ENHMETAFILE = 14
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_GDIOBJFIRST = &H300
Private Const CF_GDIOBJLAST = &H3FF
Private Const CF_HDROP = 15
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17
Private Const CF_METAFILEPICT = 3
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_OEMTEXT = 7
Private Const CF_OWNERDISPLAY = &H80
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_PRINTERFONTS = &H2
Private Const CF_PRIVATEFIRST = &H200
Private Const CF_PRIVATELAST = &H2FF
Private Const CF_RIFF = 11
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_SYLK = 4
Private Const CF_TIFF = 6
Private Const CF_TTONLY = &H40000
Private Const CF_UNICODETEXT = 13
Private Const CF_USESTYLE = &H80&
Private Const CF_WAVE = 12
Private Const CF_WYSIWYG = &H8000

#If VBA7 Then
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Private Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#Else
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Function FUserName() As String
    On Error GoTo FUserName_Err

    Dim lngLen As Long, lngX As Long
    Dim strUserName As String

    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)

    If lngX <> 0 Then
        FUserName = Left$(strUserName, lngLen - 1)
    Else
        FUserName = ""
    End If

FUserName_Exit:
    Exit Function

FUserName_Err:
    MsgBox Error$
    Resume FUserName_Exit
End Function

'open a file from a table DocumentAttachments
Public Function FOpenAttachedFile(VDocumentAttachmentID As Long)
    Dim database As DAO.database
    Dim table As DAO.Recordset
    Dim VFileName As String
    Dim Attachments As Object

    Dim CurPath As String, FilePath As String, VCurFileName As String

    CurPath = Application.CurrentProject.Path

    Set database = CurrentDb
    Set table = database.OpenRecordset("SELECT * FROM DocumentAttachments where DocumentAttachmentID=" & VDocumentAttachmentID)
    With table                                   ' For each record in table
        Do Until .EOF                            'exit with loop at end of table
            Set Attachments = table.Fields("DocumentAttachment").Value 'get list of attachments
            VFileName = table.Fields("DocumentAttachmentID").Value ' get record key
            '  Loop through each of the record's attachments'
            While Not Attachments.EOF            'exit while loop at end of record's attachments
                '  Save current attachment to disk in the above-defined folder.
                VFileName = Attachments.FileName
                If FFileExists(CurPath & "\" & VFileName) Then
                    Kill CurPath & "\" & VFileName
                End If
                Attachments.Fields("FileData").SaveToFile CurPath
                Application.FollowHyperlink CurPath & "\" & VFileName
                Attachments.MoveNext             'move to next attachment
            Wend
            .MoveNext                            'move to next record
        Loop
    End With

End Function

'set application icon from an attached file
Public Function FAppIcon()
    Dim database As DAO.database
    Dim table As DAO.Recordset
    Dim VFileName As String
    Dim Attachments As Object

    Dim CurPath As String, FilePath As String, VCurFileName As String

    CurPath = Application.CurrentProject.Path

    Set database = CurrentDb
    Set table = database.OpenRecordset("SELECT * FROM DocumentAttachments where DocumentAttachmentID=5")
    With table                                   ' For each record in table
        Do Until .EOF                            'exit with loop at end of table
            Set Attachments = table.Fields("DocumentAttachment").Value 'get list of attachments
            VFileName = table.Fields("DocumentAttachmentID").Value ' get record key
            '  Loop through each of the record's attachments'
            While Not Attachments.EOF            'exit while loop at end of record's attachments
                '  Save current attachment to disk in the above-defined folder.
                VFileName = Attachments.FileName
                If FFileExists(CurPath & "\" & VFileName) Then
                    Kill CurPath & "\" & VFileName
                End If
                Attachments.Fields("FileData").SaveToFile CurPath
                ChangeIconCurrentDB (CurPath & "\" & VFileName)
                 
                Attachments.MoveNext             'move to next attachment
            Wend
            .MoveNext                            'move to next record
        Loop
    End With
End Function

'Change AppName in current DB
Public Function ChangeAppNameCurrentDB(ByVal NewAppName As String)

    On Error Resume Next

    ChangeAppNameCurrentDB = ""
    ChangeAppNameCurrentDB = CurrentDb.Properties("AppTitle").Value

    If ChangeAppNameCurrentDB = "" Then
        CurrentDb.Properties.Append CurrentDb.CreateProperty("AppTitle", dbText, NewAppName)
    Else
        CurrentDb.Properties("AppTitle").Value = NewAppName
    End If

    Application.RefreshTitleBar

End Function

'Change AppName in other DB
Public Function ChangeAppNameOtherDB(ByVal NewAppName As String, DBPath As String)

    On Error Resume Next

    ChangeAppNameOtherDB = ""
    ChangeAppNameOtherDB = DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties("AppTitle").Value

    If ChangeAppNameOtherDB = "" Then
        DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties.Append DBEngine.Workspaces(0).OpenDatabase(DBPath).CreateProperty("AppTitle", dbText, NewAppName)
    Else
        DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties("AppTitle").Value = NewAppName
    End If

    Application.RefreshTitleBar

End Function

'Change App Icon in current DB
Public Function ChangeIconCurrentDB(ByVal NewIconPath As String)

    On Error Resume Next

    ChangeIconCurrentDB = ""
    ChangeIconCurrentDB = CurrentDb.Properties("AppIcon").Value

    If ChangeIconCurrentDB = "" Then
        CurrentDb.Properties.Append CurrentDb.CreateProperty("AppIcon", dbText, NewIconPath)
    Else
        CurrentDb.Properties("AppIcon").Value = NewIconPath
    End If

    CurrentDb.Properties("UseAppIconForFrmRpt") = 1
    Application.RefreshTitleBar

End Function

'Change App Icon in other DB
Public Function ChangeIconOtherDB(ByVal NewIconPath As String, DBPath As String)

    On Error Resume Next

    ChangeIconOtherDB = ""
    ChangeIconOtherDB = DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties("AppIcon").Value

    If ChangeIconOtherDB = "" Then
        DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties.Append DBEngine.Workspaces(0).OpenDatabase(DBPath).CreateProperty("AppIcon", dbText, NewIconPath)
    Else
        DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties("AppIcon").Value = NewIconPath
    End If

    DBEngine.Workspaces(0).OpenDatabase(DBPath).Properties("UseAppIconForFrmRpt") = 1
    Application.RefreshTitleBar

End Function

'Set AllowDesignChanges To No for all forms in the database
Public Function SetAllowDesignChangesToNo()
    On Error Resume Next
    Dim obj   As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    For Each obj In dbs.AllForms
        DoCmd.OpenForm obj.Name, acDesign
        If Forms(obj.Name).AllowDesignChanges = True Then
            Debug.Print "Updating AllowDesignChanges for " & obj.Name
            Forms(obj.Name).AllowDesignChanges = False
        End If
        DoCmd.Close acForm, obj.Name, acSaveYes
    Next obj
End Function


'pause seconds
Public Function FWait(VNoOfSeconds As Integer)
    Dim varStart As Variant
    varStart = Timer
    Do While Timer < varStart + VNoOfSeconds
    Loop
End Function

Sub CopyToFromClipboard()
'
'    Dim Clipboard As MSForms.DataObject
'    Set Clipboard = New MSForms.DataObject
'    Dim strContents As String
'
'    Clipboard.SetText "A string value"
'    Clipboard.PutInClipboard
'
'    'Or, to copy text from the clipboard into a string variable:
'    Set Clipboard = New MSForms.DataObject
'    Clipboard.GetFromClipboard
'    strContents = Clipboard.GetText

End Sub

Function ClipBoard_SetText(strCopyString As String) As Boolean
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long

    ' Allocate moveable global memory.
    hGlobalMemory = GlobalAlloc(GHND, Len(strCopyString) + 1)

    ' Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, strCopyString)

    ' Unlock the memory and then copy to the clipboard
    If GlobalUnlock(hGlobalMemory) = 0 Then
        If OpenClipboard(0&) <> 0 Then
            Call EmptyClipboard
            hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
            ClipBoard_SetText = CBool(CloseClipboard)
        End If
    End If
End Function

Function ClipBoard_GetText() As String
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
    Dim strCBText As String
    Dim RetVal As Long
    Dim lngSize As Long
    If OpenClipboard(0&) <> 0 Then
        ' Obtain the handle to the global memory block that is referencing the text.
        hClipMemory = GetClipboardData(CF_TEXT)
        If hClipMemory <> 0 Then
            ' Lock Clipboard memory so we can reference the actual data string.
            lpClipMemory = GlobalLock(hClipMemory)
            If lpClipMemory <> 0 Then
                lngSize = GlobalSize(lpClipMemory)
                strCBText = Space$(lngSize)
                RetVal = lstrcpy(strCBText, lpClipMemory)
                RetVal = GlobalUnlock(hClipMemory)
                ' Peel off the null terminating character.
                strCBText = Left(strCBText, InStr(1, strCBText, Chr$(0), 0) - 1)
            Else
                MsgBox "Could not lock memory to copy string from."
            End If
        End If
        Call CloseClipboard
    End If
    ClipBoard_GetText = strCBText
End Function

Function CopyOlePiccy(Piccy As Object)
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
    Dim hClipMemory As Long, x As Long

    ' Allocate moveable global memory.
    hGlobalMemory = GlobalAlloc(GHND, Len(Piccy) + 1)

    ' Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    'Need to copy the object to the memory here
    lpGlobalMemory = lstrcpy(lpGlobalMemory, Piccy)

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    ' Clear the Clipboard.
    x = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If
End Function


Function killHyperlinkWarning()
    Dim oShell As Object
    Dim strReg As String

    strReg = "Software\Microsoft\Office\" & Application.Version & _
             "\Common\Security\DisableHyperlinkWarning"
    '  CreateObject("Wscript.Shell").RegWrite _
    '       "HKCU\Software\Microsoft\Office\" & Application.Version & _
    '       "\Common\Security\DisableHyperlinkWarning", 1, "REG_DWORD"
    Set oShell = CreateObject("Wscript.Shell")
    oShell.RegWrite "HKCU\" & strReg, 1, "REG_DWORD"
End Function

Public Function ShellExecute(ByVal Command As String, Optional ByVal Parameters As String)
    
    If Len(Parameters) = 0 Then
        Parameters = vbNullString
    End If
    ApiShellExecute Application.hWndAccessApp, "Open", Command, Parameters, vbNullString, 1

End Function
