Attribute VB_Name = "MODBC"
Option Compare Database
Option Explicit

Public Function ODBCConnString()

    Dim CurUser As String
    Dim OracleConnection As String

    'CurUser = fOSUserName
    OracleConnection = DLookup("ConnectionName", "AP_SETTINGS")

    ' If CurUser = "natalie.rynda" Then
    ODBCConnect = "ODBC;DRIVER={" & GetOracleDriver & "};SERVER=" & OracleConnection & ";DBQ=" & OracleConnection & ";Trusted_Connection=Yes"
    '  Else
    '      NewConnect = "ODBC;DRIVER={" & GetOracleDriver & "};SERVER=" & OracleConnection & ";DBQ=" & OracleConnection & ";Trusted_Connection=Yes"
    ' End If

End Function

Public Function ODBCLinkTables()

    ODBCConnString

    'Debug.Print NewConnect
                           
    Dim tdf   As DAO.TableDef
    Dim qdf   As DAO.QueryDef

    For Each tdf In CurrentDb.TableDefs
        If tdf.Connect Like "*ggg.DB*" Then
            tdf.Connect = ODBCConnect
            tdf.RefreshLink
        End If
    Next tdf

    For Each qdf In CurrentDb.QueryDefs
        'Debug.Print qdf.Name
        If (qdf.Type = 144 Or qdf.Type = 112) And qdf.Connect Like "*ggg.DB*" Then
            qdf.Connect = ODBCConnect
        End If
    Next

End Function

Public Function GetOracleDriver()

    Dim strComputer As String
    Dim strValueName As String

    Dim arrValueNames As Variant
    Dim arrValueTypes As Variant
    Dim i     As Long
    Dim R     As Long
    Dim strKeyPath As String
    Dim strValue As String
    Dim objReg As Object
    Dim MyDriverName As String

    Const HKEY_LOCAL_MACHINE = &H80000002

    R = 1
    strComputer = "."

    Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

    strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes

    For i = 0 To UBound(arrValueNames)
        strValueName = arrValueNames(i)
        objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
        If strValue = "Installed" And (arrValueNames(i) Like "*oracle*" And arrValueNames(i) <> "Microsoft ODBC for oracle") Then
            GetOracleDriver = arrValueNames(i)
        End If
        R = R + 1
    Next i

    If IsNull(GetOracleDriver) Then

        R = 1
        strComputer = "."

        Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

        strKeyPath = "SOFTWARE\WOW6432NODE\ODBC\ODBCINST.INI\ODBC Drivers"
        objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes

        For i = 0 To UBound(arrValueNames)
            strValueName = arrValueNames(i)
            objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
            If strValue = "Installed" And (arrValueNames(i) Like "*oracle*" And arrValueNames(i) <> "Microsoft ODBC for oracle") Then
                GetOracleDriver = arrValueNames(i)
            End If
            R = R + 1
        Next i

    End If
    'Debug.Print GetOracleDriver

End Function

Public Function GetOracleDriver64()

    Dim strComputer As String
    Dim strValueName As String

    Dim arrValueNames As Variant
    Dim arrValueTypes As Variant
    Dim i     As Long
    Dim R     As Long
    Dim strKeyPath As String
    Dim strValue As String
    Dim objReg, objCtx, objLocator, objServices, objStdRegProv As Object
    Dim MyDriverName As String

    Const HKEY_LOCAL_MACHINE = &H80000002

    R = 1
    strComputer = "."
        
    '64bit
    
    'The code derives from
    'https://docs.microsoft.com/en-us/windows/win32/wmisdk/requesting-wmi-data-on-a-64-bit-platform
    Const HKLM = &H80000002
    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    objCtx.Add "__ProviderArchitecture", 64
    objCtx.Add "__RequiredArchitecture", True
    Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set objServices = objLocator.ConnectServer(strComputer, "root\default", "", "", , , , objCtx)
    Set objStdRegProv = objServices.Get("StdRegProv")

    'Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    Set objReg = objStdRegProv
    
    strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes

    For i = 0 To UBound(arrValueNames)
        strValueName = arrValueNames(i)
        objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
        Debug.Print strKeyPath, strValueName, strValue
        If strValue = "Installed" And (arrValueNames(i) Like "*oracle*" And arrValueNames(i) <> "Microsoft ODBC for oracle") Then
            GetOracleDriver64 = arrValueNames(i)
        End If
        R = R + 1
    Next i
    
    
    '32bit
    If IsNull(GetOracleDriver64) Then

        R = 1
        strComputer = "."

        Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

        strKeyPath = "SOFTWARE\WOW6432NODE\ODBC\ODBCINST.INI\ODBC Drivers"
        objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes

        For i = 0 To UBound(arrValueNames)
            strValueName = arrValueNames(i)
            objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
            If strValue = "Installed" And (arrValueNames(i) Like "*oracle*" And arrValueNames(i) <> "Microsoft ODBC for oracle") Then
                GetOracleDriver64 = arrValueNames(i)
            End If
            R = R + 1
        Next i


    End If

End Function

Public Function GetTNSName(strFile As String)

    'GetOracleDriver gives my driver name

    Dim hFile As String, strData As String * 60000

    hFile = FreeFile
    Open strFile For Binary Access Read As hFile Len = 4000
    Get hFile, 1, strData
    Close hFile

    Dim LastChar As Long, FirstChar As Long

    LastChar = InStrRev(Mid(strData, 1, InStrRev(Mid(strData, 1, InStr(strData, "SERVICE_NAME=ddd.ddd.com")), "description")), ".ddd.com") - 1
    FirstChar = LastChar

    Do While IsLetterOrNumber(Mid(strData, FirstChar, 1)) = True
        FirstChar = FirstChar - 1
    Loop

    GetTNSName = Mid(strData, FirstChar, LastChar - FirstChar + 1)

End Function

Public Function RecursiveDir(colFiles As Collection, strFolder As String, strFileSpec As String, bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add files in strFolder Match strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString And Not strTemp Like "*diag*"
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each vFolderName In colFolders
            If vFolderName <> "diag" Then
                Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
            End If
        Next vFolderName
    End If

End Function

Public Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function

