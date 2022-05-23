Attribute VB_Name = "MFTPDocs"
Option Compare Database
Option Explicit

Const FTP_TRANSFER_TYPE_ASCII = &H1
Const FTP_TRANSFER_TYPE_BINARY = &H2
Const INTERNET_DEFAULT_FTP_PORT = 21
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_FLAG_PASSIVE = &H8000000
Const GENERIC_WRITE = &H40000000
Const BUFFER_SIZE = 100
Const PassiveConnection As Boolean = True

' Declare wininet.dll API Functions

#If VBA7 Then
    Public Declare PtrSafe Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    Public Declare PtrSafe Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
    Public Declare PtrSafe Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWite As Long, dwNumberOfBytesWritten As Long) As Integer
    Public Declare PtrSafe Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
    Public Declare PtrSafe Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    Public Declare PtrSafe Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
    Public Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
    Public Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Public Declare PtrSafe Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Public Declare PtrSafe Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    Declare PtrSafe Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Long) As Boolean
#Else
    Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
    Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWite As Long, dwNumberOfBytesWritten As Long) As Integer
    Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
    Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
    Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
    Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Long) As Boolean
#End If

Public Sub CallingTheFunction()

    If FTPFile("ftp.domain.com", "myUserName", "myPassword", "Full path and Filename of local file", _
               "Target Filename without path", "Directory on FTP server", "Upload Mode - Binary or ASCII") Then
        MsgBox "Upload - Complete!"
    End If

End Sub

Function FTPFile(ByVal HostName As String, ByVal UserName As String, ByVal Password As String, ByVal LocalFileName As String, ByVal RemoteFileName As String, _
                 ByVal sDir As String, ByVal sMode As String) As Boolean

    On Error GoTo Err_Function
    ' Declare variables

    Dim hConnection, hOpen, hFile As Long        ' Used For Handles
    Dim iSize As Long                            ' Size of file for upload
    Dim RetVal As Variant                        ' Used for progress meter
    Dim iWritten As Long                         ' Used by InternetWriteFile to report bytes uploaded
    Dim iLoop As Long                            ' Loop for uploading chuncks
    Dim iFile As Integer                         ' Used for Local file handle
    Dim FileData(BUFFER_SIZE - 1) As Byte        ' buffer array of BUFFER_SIZE (100) elements 0 to 99

    ' Open Internet Connecion
    hOpen = InternetOpen("FTP", 1, "", vbNullString, 0)

    ' Connect to FTP
    hConnection = InternetConnect(hOpen, HostName, INTERNET_DEFAULT_FTP_PORT, UserName, Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)

    ' Change Directory
    Call FtpSetCurrentDirectory(hConnection, sDir)

    ' Open Remote File
    hFile = FtpOpenFile(hConnection, RemoteFileName, GENERIC_WRITE, IIf(sMode = "Binary", FTP_TRANSFER_TYPE_BINARY, FTP_TRANSFER_TYPE_ASCII), 0)

    ' Check for successfull file handle
    If hFile = 0 Then
        MsgBox "Internet - Failed!"
        ShowError
        FTPFile = False
        GoTo Exit_Function
    End If
        
    ' Set Upload Flag to True
    FTPFile = True

    ' Get next file handle number
    iFile = FreeFile

    ' Open local file
    Open LocalFileName For Binary Access Read As iFile

    ' Set file size
    iSize = LOF(iFile)

    ' Iinitialise progress meter
    RetVal = SysCmd(acSysCmdInitMeter, "Uploading File (" & RemoteFileName & ")", iSize / 1000)

    ' Loop file size
    For iLoop = 1 To iSize \ BUFFER_SIZE
        
        ' Update progress meter
        RetVal = SysCmd(acSysCmdUpdateMeter, (BUFFER_SIZE * iLoop) / 1000)
        
        'Get file data
        Get iFile, , FileData
    
        ' Write chunk to FTP checking for success
        If InternetWriteFile(hFile, FileData(0), BUFFER_SIZE, iWritten) = 0 Then
            MsgBox "Upload - Failed!"
            ShowError
            FTPFile = False
            GoTo Exit_Function
        Else
            ' Check buffer was written
            If iWritten <> BUFFER_SIZE Then
                MsgBox "Upload - Failed!"
                ShowError
                FTPFile = False
                GoTo Exit_Function
            End If
        End If
      
    Next iLoop                                   ' Handle remainder using MOD
  
  
    ' Update progress meter
    RetVal = SysCmd(acSysCmdUpdateMeter, iSize / 1000)

    ' Get file data
    Get iFile, , FileData

    ' Write remainder to FTP checking for success
    If InternetWriteFile(hFile, FileData(0), iSize Mod BUFFER_SIZE, iWritten) = 0 Then
        MsgBox "Upload - Failed!"
        ShowError
        FTPFile = False
        GoTo Exit_Function
    Else
        ' Check buffer was written
        If iWritten <> iSize Mod BUFFER_SIZE Then
            MsgBox "Upload - Failed!"
            ShowError
            FTPFile = False
            GoTo Exit_Function
        End If
    End If
    
Exit_Function:
    ' remove progress meter
    RetVal = SysCmd(acSysCmdRemoveMeter)

    'close remote fileCall
    InternetCloseHandle (hFile)

    'close local file
    Close iFile

    ' Close Internet Connection
    Call InternetCloseHandle(hOpen)
    Call InternetCloseHandle(hConnection)

    Exit Function

Err_Function:
    MsgBox "Error in FTPFile : " & Err.Description
    GoTo Exit_Function
End Function

Public Sub ShowError()

    Dim lErr  As Long, sErr As String, lenBuf As Long

    'get the required buffer size
    InternetGetLastResponseInfo lErr, sErr, lenBuf

    'create a buffer
    sErr = String(lenBuf, 0)

    'retrieve the last respons info
    InternetGetLastResponseInfo lErr, sErr, lenBuf

    'show the last response info
    MsgBox "Last Server Response : " + sErr, vbOKOnly + vbCritical

End Sub


