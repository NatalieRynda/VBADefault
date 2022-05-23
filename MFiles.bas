Attribute VB_Name = "MFiles"
Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Function ZIP(ZipFile As String, InputFile As String)
    On Error GoTo ErrHandler
    Dim fso   As Object                          'Scripting.FileSystemObject
    Dim oAPP  As Object                          'Shell32.Shell
    Dim oFld  As Object                          'Shell32.Folder
    Dim oShl  As Object                          'WScript.Shell
    Dim i     As Long
    Dim l     As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(ZipFile) Then
        'Create empty ZIP file
        fso.CreateTextFile(ZipFile, True).Write _
        "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
    End If

    Set oAPP = CreateObject("Shell.Application")
    Set oFld = oAPP.NameSpace(CVar(ZipFile))
    i = oFld.OutItems.Count
    oFld.CopyHere (InputFile)

    Set oShl = CreateObject("WScript.Shell")

    'Search for a Compressing dialog
    Do While oShl.AppActivate("Compressing...") = False
        If oFld.OutItems.Count > i Then
            'There's a file in the zip file now, but
            'compressing may not be done just yet
            Exit Do
        End If
        If l > 30 Then
            '3 seconds has elapsed and no Compressing dialog
            'The zip may have completed too quickly so exiting
            Exit Do
        End If
        DoEvents
        Sleep 100
        l = l + 1
    Loop

    ' Wait for compression to complete before exiting
    Do While oShl.AppActivate("Compressing...") = True
        DoEvents
        Sleep 100
    Loop

ExitProc:
    On Error Resume Next
    Set fso = Nothing
    Set oFld = Nothing
    Set oAPP = Nothing
    Set oShl = Nothing
    Exit Function
ErrHandler:
    Select Case Err.Number
    Case Else
        MsgBox "Error " & Err.Number & _
               ": " & Err.Description, _
               vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Function

Public Sub UnZip(ZipFile As String, Optional TargetFolderPath As String = vbNullString, Optional OverwriteFile As Boolean = False)
    On Error GoTo ErrHandler
    Dim oAPP  As Object
    Dim fso   As Object
    Dim fil   As Object
    Dim DefPath As String
    Dim strDate As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(TargetFolderPath) = 0 Then
        DefPath = CurrentProject.Path & ""
    Else
        If fso.FolderExists(TargetFolderPath) Then
            DefPath = TargetFolderPath & ""
        Else
            Err.Raise 53, , "Folder not found"
        End If
    End If

    If fso.FileExists(ZipFile) = False Then
        MsgBox "System could not find " & ZipFile _
             & " upgrade cancelled.", _
               vbInformation, "Error Unziping File"
        Exit Sub
    Else
        'Extract the files into the newly created folder
        Set oAPP = CreateObject("Shell.Application")

        With oAPP.NameSpace(ZipFile & "")
            If OverwriteFile Then
                For Each fil In .OutItems
                    If fso.FileExists(DefPath & fil.Name) Then
                        Kill DefPath & fil.Name
                    End If
                Next
            End If
            oAPP.NameSpace(CVar(DefPath)).CopyHere .OutItems
        End With

        On Error Resume Next
        Kill Environ("Temp") & "Temporary Directory*"

        Kill ZipFile
    End If

ExitProc:
    On Error Resume Next
    Set oAPP = Nothing
    Exit Sub
ErrHandler:
    Select Case Err.Number
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Sub

'see if file exists
Function FFileExists(VFileName As String) As Boolean
    FFileExists = (Dir(VFileName) > "")
End Function

'delete file if exists
Public Function FDeleteFileIfExists(VFileToDelete As String)

    If FFileExists(VFileToDelete) Then             'See above
        SetAttr VFileToDelete, vbNormal
        Kill VFileToDelete
    End If

End Function

'count of files in folder
Function FCountFiles(VDirectory As String) As Double
    Dim fso   As Object, _
    objFiles  As Object
   
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set objFiles = fso.GetFolder(VDirectory).files
    If Err.Number <> 0 Then
        FCountFiles = 0
    Else
        FCountFiles = objFiles.Count
    End If
    On Error GoTo 0
End Function

'list files in folder
Public Function FListFilesInFolder(VFolder As String, Optional VFileType As String = "*.*")

    If Dir(VFolder, vbDirectory) = "" Then
        GoTo ExitProc
    End If

    Dim strFileN As String
    strFileN = Dir(VFolder & VFileType)

    If Len(strFileN) = 0 Then
        GoTo ExitProc
    End If

    Do While Len(strFileN) > 0
        Debug.Print strFileN
        strFileN = Dir
    Loop

ExitProc:
End Function

'loop through files
Sub LoopThroughFiles()
    Dim MyObj As Object, MySource As Object, file As Variant
    file = Dir("\\fffff\ffff\ffff\ffff\ffff\_Load\")
    While (file <> "")
        ' If InStr(file, "test") > 0 Then
        MsgBox "found " & file
        Exit Sub
        '  End If
        file = Dir
    Wend
End Sub

'get files in folder
Function FGetFilesInFolder(VPath As String, Optional VPattern As String = "") As Collection
    Dim rv    As New Collection, f
    If Right(VPath, 1) <> "\" Then VPath = VPath & "\"
    f = Dir(VPath & VPattern)
    Do While Len(f) > 0
        rv.Add VPath & f
        f = Dir()                                'no parameter
    Loop
    Set FGetFilesInFolder = rv
End Function

'test get files in folder
Sub Tester()

    Dim fls, f

    Set fls = FGetFilesInFolder("D:\Analysis\", "*.xls*")
    For Each f In fls
        Debug.Print f
    Next f

End Sub

'loop through folders and subfolders, code for reference
Public Function SLoopThroughFoldersSubFoldersFiles()

    Dim fso, oFolder, oSubfolder, oFile, queue As Collection
    Dim CurProject As String
    Dim VTicketNo As Variant
    Dim VTicketsFolder As String, VTicketResultsFolder As String
    
    VTicketsFolder = "\\xxxx\xxx\xx\Tickets\"
    VTicketResultsFolder = "\\xxxx\xxx\xx\TicketResults\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.Add fso.GetFolder(VTicketsFolder)      'obviously replace

    Do While queue.Count > 0

        Set oFolder = queue(1)
        queue.Remove 1                           'dequeue

        CurProject = Mid(oFolder, 62)

        For Each oSubfolder In oFolder.SubFolders
            queue.Add oSubfolder                 'enqueue
        Next oSubfolder

        For Each oFile In oFolder.files
    
            VTicketNo = Mid(oFile, 44, InStr(Mid(oFile, 44, 99), "\") - 1)

            If DLookup("TicketNo", "tickets", "TicketNo=" & VTicketNo) Then

                If Dir(VTicketsFolder & VTicketNo, vbDirectory) = "" Then
                    MkDir (VTicketsFolder & VTicketNo)
                End If
           
                FileCopy VTicketsFolder & Replace(oFile, VTicketsFolder, ""), VTicketsFolder & Replace(oFile, VTicketsFolder, "")
           
            End If
           
            If DLookup("TicketNo", "tickets", "Results=true and TicketNo=" & VTicketNo) And Not oFile Like "*sql*" Then

                If Dir(VTicketResultsFolder & VTicketNo, vbDirectory) = "" Then
                    MkDir (VTicketResultsFolder & VTicketNo)
                End If
           
                FileCopy VTicketsFolder & Replace(oFile, VTicketsFolder, ""), VTicketResultsFolder & Replace(oFile, VTicketsFolder, "")
           
            End If
    
        Next oFile
    Loop

End Function

Function GetFilesIn(Folder As String) As Collection
    Dim f     As String
    Set GetFilesIn = New Collection
    f = Dir(Folder & "\*")
    Do While f <> ""
        GetFilesIn.Add f
        f = Dir
    Loop
End Function

Function GetFoldersIn(Folder As String) As Collection
    Dim f     As String
    Set GetFoldersIn = New Collection
    f = Dir(Folder & "\*", vbDirectory)
    Do While f <> ""
        If GetAttr(Folder & "\" & f) And vbDirectory Then GetFoldersIn.Add f
        f = Dir
    Loop
End Function

'pick file with latest modified date, code for reference
Public Sub SPickFileWithLatestModifiedDate()

    Dim mainpath As String
    Dim fso, FSOFiles, oSubfolder, oFile, queue As Collection

    mainpath = "\\fff\fff\fff\fffff\fffff\"

    Dim CurFolderC As Collection, CurFolder
    Dim CurFileC As Collection, CurFile
    Dim ChosenFile As String
    Dim ChosenModDate As Date
    Dim CurModDate As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set CurFolderC = GetFoldersIn(mainpath)

    For Each CurFolder In CurFolderC
     
        If CurFolder <> "." And CurFolder <> ".." Then
      
            Set CurFileC = GetFilesIn(mainpath & CurFolder)
            For Each CurFile In CurFileC
        
                If InStr(CurFile, ".xlsx") > 0 Then
              
                    Set FSOFiles = fso.GetFile(mainpath & CurFolder & "\" & CurFile)
              
                    CurModDate = FSOFiles.DateLastModified
              
                    If ChosenModDate = "12:00:00 AM" Then
                        ChosenModDate = CurModDate
                    End If
     
                    If ChosenFile = "" Then
                        ChosenFile = CurFile
                    End If
                                
                    If ChosenModDate < CurModDate Then
                        ChosenModDate = CurModDate
                        ChosenFile = CurFile
                    End If
                  
                End If
              
            Next CurFile
     
            Debug.Print ChosenFile
            ChosenModDate = "12:00:00 AM"
            ChosenFile = ""
     
        End If

    Next CurFolder

    Set fso = Nothing
    Set FSOFiles = Nothing
  
End Sub

'loop through folders and subfolders, code for reference
Public Sub NonRecursiveMethod()
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection

    Dim MainFolder As String

    MainFolder = "\\ddd\ddd\ddd\_tmp\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.Add fso.GetFolder(MainFolder)          'obviously replace

    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1                           'dequeue
        '...insert any folder processing code here...
        For Each oSubfolder In oFolder.SubFolders
            queue.Add oSubfolder                 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.files
        
            Dim TICKETNO As Long
            Dim FileName As String
    
            Debug.Print InStr(Replace(oFile, MainFolder, ""), "\")
            Debug.Print Replace(oFile, MainFolder, "")
            'TicketNo = StripNonNum(Mid(oFile, 71, 99))
            TICKETNO = Left(Replace(oFile, MainFolder, ""), InStr(Replace(oFile, MainFolder, ""), "\") - 1)
            FileName = Mid(Replace(oFile, MainFolder, ""), 8, 99)

            If DLookup("TicketNo", "tickets", "TicketNo=" & TICKETNO) Then
                DoCmd.RunSql "insert into FolderLink select '" & oFile & "' as folderlink, '" & TICKETNO & "' as TicketNo ,'" & FileName & "' as FileName from DUAL"
            End If
        Next oFile
    Loop

End Sub


