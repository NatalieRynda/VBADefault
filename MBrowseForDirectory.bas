Attribute VB_Name = "MBrowseForDirectory"
Option Compare Database
Option Explicit

'http://www.mvps.org/access/api/api0002.htm
'Code courtesy of Terry Kreft

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Declare PtrSafe Function SHBrowseForFolder Lib "Shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
#Else
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
#End If

Private Const BIF_RETURNONLYFSDIRS = &H1

Public Function BrowseDirectory(szDialogTitle As String) As String
    On Error GoTo Err_BrowseDirectory

    Dim x     As Long, bi As BROWSEINFO, dwIList As Long
    Dim szPath As String, wPos As Integer
 
    With bi
        .hOwner = hWndAccessApp
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
 
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
 
    If x Then
        wPos = InStr(szPath, Chr(0))
        BrowseDirectory = Left$(szPath, wPos - 1)
    Else
        BrowseDirectory = ""
    End If

Exit_BrowseDirectory:
    Exit Function

Err_BrowseDirectory:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_BrowseDirectory

End Function

Public Function TestOpeningDirectory()
    On Error GoTo Err_TestOpeningDirectory
 
    Dim sDirectoryName As String
 
    sDirectoryName = BrowseDirectory("Find and select where to export the Excel report files.")
 
    If sDirectoryName <> "" Then MsgBox "You selected the '" & sDirectoryName & "' directory.", vbInformation
 
Exit_TestOpeningDirectory:
    Exit Function
 
Err_TestOpeningDirectory:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_TestOpeningDirectory

End Function

Sub SimpleBrowse()
  Dim selectedFolder

    With Application.FileDialog(4)
        .Show
        selectedFolder = .SelectedItems(1)
    End With

'                Me.FileSelected = selectedFolder
'                Me.BOK.Visible = True

End Sub

Public Function BrowseToFile()
 
    'f.AllowMultiSelect = False
    'f.Show
    'MsgBox "file choosen = " & f.SelectedItems.AddItem

    Dim FSOFD As Object
    Set FSOFD = Application.FileDialog(3)
    FSOFD.AllowMultiSelect = False
    Dim vrtSelectedItem As Variant
 
    With FSOFD
      .InitialFileName = CurrentProject.Path
      .Title = "Select the backend database"
      .Filters.Clear
      .Filters.Add "Access Databases", "*.accdb"
            
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
           '     Me.FileSelected = vrtSelectedItem
             '   BOK_Click

            Next vrtSelectedItem
        Else
            'The user pressed Cancel.
        End If
    End With
 
    Set FSOFD = Nothing
         
End Function

