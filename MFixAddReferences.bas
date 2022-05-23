Attribute VB_Name = "MFixAddReferences"
Option Compare Database
Option Explicit

Function FixUpRefs()

    Dim loRef As Access.Reference
    Dim intCount As Integer
    Dim intX  As Integer
    Dim blnBroke As Boolean
    Dim strPath As String

    On Error Resume Next

    intCount = Access.References.Count
    'Loop through each reference in the database and determine if the reference is broken. If it is broken, remove the Reference and add it back.
    Debug.Print "----------------- References found -----------------------"
    Debug.Print " reference count = "; intCount

    For intX = intCount To 1 Step -1
        Set loRef = Access.References(intX)
        With loRef
            
            Debug.Print " reference = "; .FullPath
            blnBroke = .IsBroken
        
            If blnBroke = True Or Err <> 0 Then
                strPath = .FullPath
                Debug.Print " ***** Err = "; Err; " and Broke = "; blnBroke
                      
                With Access.References
                    .Remove loRef
                    Debug.Print "path name = "; strPath
                    .AddFromFile strPath
                End With
            End If
           
        End With
    Next
      
    '''Access.References.AddFromFile "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    Set loRef = Nothing
   
    ' Call a hidden SysCmd to automatically compile/save all modules.
    Call SysCmd(504, 16483)
   
End Function

Function AddRefs()

    Dim loRef As Access.Reference
    Dim intCount As Integer
    Dim intX  As Integer
    Dim blnBroke As Boolean
    Dim strPath As String

    On Error Resume Next

    Debug.Print "----------------- Add References -----------------------"

    With Access.References
        .AddFromFile "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
        .AddFromFile "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\vbe6.dll"
        .AddFromFile "C:\Program Files\Microsoft Office\Office\msacc9.olb"
        .AddFromFile "C:\Program Files\Common Files\System\ado\msado15.dll"
        .AddFromFile "C:\Program Files\Common Files\System\ado\msado25.tlb"
        .AddFromFile "C:\Program Files\Common Files\System\ado\msadox.dll"
        .AddFromFile "C:\WINNT\System32\stdole2.tlb"
        .AddFromFile "C:\WINNT\System32\scrrun.dll"

    End With

    ' Call a hidden SysCmd to automatically compile/save all modules.
    Call SysCmd(504, 16483)
End Function

