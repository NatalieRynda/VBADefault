Attribute VB_Name = "MOutlook"
Option Compare Database
Option Explicit

Const OutlFolderInbox As Integer = 6
Const OutlFolderIDeletedItems As Integer = 3

#Const LateBind = True

Const olMinimized As Long = 1
Const olMaximized As Long = 2
Const olFolderInbox As Long = 6

#If LateBind Then

Public Function OutlookApp( _
       Optional WindowState As Long = olMinimized, _
       Optional ReleaseIt As Boolean = False _
       ) As Object
    Static o As Object
#Else
Public Function OutlookApp( _
       Optional WindowState As Outlook.OlWindowState = olMinimized, _
       Optional ReleaseIt As Boolean _
       ) As Outlook.Application
    Static o As Outlook.Application
#End If
On Error GoTo ErrHandler
 
Select Case True
Case o Is Nothing, Len(o.Name) = 0
    Set o = GetObject(, "Outlook.Application")
    If o.Explorers.Count = 0 Then
InitOutlook:
        'Open inbox to prevent errors with security prompts
        o.Session.GetDefaultFolder(olFolderInbox).Display
        o.ActiveExplorer.WindowState = WindowState
    End If
Case ReleaseIt
    Set o = Nothing
End Select
Set OutlookApp = o
 
ExitProc:
Exit Function
ErrHandler:
Select Case Err.Number
Case -2147352567
    'User cancelled setup, silently exit
    Set o = Nothing
Case 429, 462
    Set o = GetOutlookApp()
    If o Is Nothing Then
        Err.Raise 429, "OutlookApp", "Outlook Application does not appear to be installed."
    Else
        Resume InitOutlook
    End If
Case Else
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
End Select
Resume ExitProc
Resume
End Function

#If LateBind Then
Private Function GetOutlookApp() As Object
#Else
Private Function GetOutlookApp() As Outlook.Application
#End If
On Error GoTo ErrHandler
    
Set GetOutlookApp = CreateObject("Outlook.Application")
    
ExitProc:
Exit Function
ErrHandler:
Select Case Err.Number
Case Else
    'Do not raise any errors
    Set GetOutlookApp = Nothing
End Select
Resume ExitProc
Resume
End Function

Public Function OutlookDeletedItems()
  
    CurPath = CurrentProject.Path & "\"

    Dim i, CountOfItems As Long
    Dim EmailContTD, EmailContNew As String, OutlPrefix As String
    
    Set OutlApp = GetObject(, "Outlook.application")
    Set OutlNameSpace = OutlApp.GetNamespace("MAPI")
    Set OutlFolder = OutlNameSpace.GetDefaultFolder(OutlFolderIDeletedItems)
    'Set OutlMail = GetObject(, "Outlook.MailItem")

    OutlMyUTC = 7
    OutlPrefix = "urn:schemas:httpmail:"

    OutlStartDate = Format(DateAdd("h", -OutlMyUTC, Date), "\'m/d/yyyy\") & " 12:00 AM'"
    'OutlEndDate = Format(DateAdd("h", -OutlMyUTC, CDate("1/18/2018 4:00 PM")), "\'m/d/yyyy hh:mm AM/PM\'")
    OutlSentBy = "ddd, ddd"
    OutlSentBy2 = "ddd@ddd.com"
    OutlSubjectCriteria1 = "dddddddddddd *"
    OutlSubjectCriteria2 = "dddddddddddddd"
    OutlSubjectCriteria3 = "ddddddd"

    'OutlFilter = "@SQL= ((urn:schemas:httpmail:sendername = '" & OutlSentBy & "' OR urn:schemas:httpmail:sendername = '" & OutlSentBy2 & "') And urn:schemas:httpmail:datereceived >= " & OutlStartDate & ")"
    OutlFilter = "@SQL= ((" & OutlPrefix & "sendername = '" & OutlSentBy & "' OR " & OutlPrefix & "sendername = '" & OutlSentBy2 & "') AND " & OutlPrefix & "datereceived >= " & OutlStartDate & _
                 " AND (" & OutlPrefix & "subject = '" & OutlSubjectCriteria3 & "' OR " & OutlPrefix & "subject = '" & OutlSubjectCriteria2 & "' OR " & OutlPrefix & "subject Like '" & OutlSubjectCriteria1 & "')) "
    'Debug.Print OutlFilter

    'OutlFilter = "[UnRead] = True"
    'OutlFilter = "@SQL= (urn:schemas:httpmail:sendername = " & OutlSentBy & " And (urn:schemas:httpmail:datereceived >= " & OutlStartDate &  " And urn:schemas:httpmail:datereceived <= " & OutlEndDate & "))"
  
    Set OutlFilteredItem = OutlFolder.Items.Restrict(OutlFilter)
  
    CountOfItems = OutlFilteredItem.Count
    If CountOfItems = 0 Then
        Exit Function
    End If
  
    DoCmd.RunSql "DELETE FROM ddd"
    DoCmd.RunSql "DELETE FROM ddddd"

    With OutlFilteredItem
        For i = CountOfItems To 1 Step -1
            If TypeName(OutlFilteredItem(i)) = "MailItem" Then
                Set OutlMailItem = OutlFilteredItem(i)
        
                'OutlSenderLogin = OutlMailItem.SenderEmailAddress
                'OutlSenderName = OutlMailItem.SenderName
                'OutlSenderEMail = OutlMailItem.Sender.GetExchangeUser.PrimarySmtpAddress
                OutlDateReceived = OutlMailItem.ReceivedTime
                'OutlDateSent = OutlMailItem.CreationTime
                OutlSubject = OutlMailItem.Subject
                OutlMsgBody = OutlMailItem.Body
        
                If OutlSubject Like OutlSubjectCriteria1 Then
                    'Debug.Print oMail.ReceivedTime
                    EmailContTD = Replace(OutlMsgBody, Chr(34), "")
                    ''Debug.Print EmailCont
                    DoCmd.RunSql "INSERT INTO SNTargetDates (Contents, DateReceived) SELECT """ & EmailContTD & """ AS Contents, #" & OutlDateReceived & "# AS DateReceived FROM DUAL;"
                End If
                 
                If OutlSubject = OutlSubjectCriteria2 Then
                    'Debug.Print oMail.ReceivedTime
                    EmailContNew = Replace(OutlMsgBody, Chr(34), "")
                    'Debug.Print "INSERT INTO SNNew (Contents) SELECT """ & EmailContNew & """ AS Expr1 FROM DUAL;"
                    DoCmd.RunSql "INSERT INTO SNNew (Contents) SELECT """ & EmailContNew & """ AS Expr1 FROM DUAL;"
                End If
                 
                If OutlSubject = OutlSubjectCriteria3 Then
                    'Debug.Print oMail.ReceivedTime
                    For Each OutlAttach In OutlMailItem.Attachments
                        OutlAttach.SaveAsFile CurPath & "_Load\MyTickets.xlsx"
                        DoCmd.RunSql "UPDATE MyTicketsDate SET MyTicketsDate.MyTicketsDate = #" & FileDateTime(CurPath & "_Load\MyTickets.xlsx") & "#;"
                    Next OutlAttach
                End If
            End If
          
            EmailContTD = ""
            EmailContNew = ""
        Next
    End With

End Function

Public Function GetOutlookAttachments()

    CurPath = CurrentProject.Path & "\"

    Dim i, CountOfItems As Long
    Dim MyDate As String
    
    '~~> Get Outlook instance
    Set OutlApp = GetObject(, "Outlook.application")
    Set OutlNameSpace = OutlApp.GetNamespace("MAPI")
    Set OutlFolder = OutlNameSpace.GetDefaultFolder(OutlFolderInbox)
    'Set OutlMail = GetObject(, "Outlook.MailItem")

    MyDate = Format(Now, "\'m/d/yyy hh:mm AM/PM\'") '/* will give '1/23/2018 01:36 PM' */


    OutlMyUTC = 0                                '/* this is your UTC, change to suit (in my case 8) */

    OutlStartDate = Format(DateAdd("h", -OutlMyUTC, Date), "\'m/d/yyyy hh:mm AM/PM\'")
    'OutlEndDate = Format(DateAdd("h", -OutlMyUTC, CDate("1/18/2018 4:00 PM")), "\'m/d/yyyy hh:mm AM/PM\'")
    OutlSentBy = "wilson, elizabeth"             '/* can be sendername, "doe, john" */

    '/* filter in one go, where datereceived is
    'expressed in UTC (Universal Coordinated Time) */
    OutlFilter = "@SQL= (urn:schemas:httpmail:sendername = '" & OutlSentBy & "' And urn:schemas:httpmail:datereceived >= " & OutlStartDate & ")"
    'OutlFilter = "[UnRead] = True"
    'Debug.Print OutlFilter
    '
    '  OutlFilter = "@SQL= (urn:schemas:httpmail:sendername = " & OutlSentBy & _
    '       " And (urn:schemas:httpmail:datereceived >= " & OutlStartDate & _
    '       " And urn:schemas:httpmail:datereceived <= " & OutlEndDate & "))"

    '~~> Check if there are any emails that meet the criteria
    CountOfItems = OutlFolder.Items.Restrict(OutlFilter).Count
    If CountOfItems = 0 Then
        Exit Function
    End If
  
    Set OutlItem = OutlFolder.Items.Restrict(OutlFilter)

    With OutlItem
        For i = CountOfItems To 1 Step -1        '/* starting from most recent */
            If TypeName(OutlFolder.Items(i)) = "MailItem" Then
                Set OutlItem = OutlFolder.Items(i)
        
        
                For Each OutlAttach In OutlItem.Attachments
                    If OutlAttach.Type = 1 And InStr(OutlAttach, "_ACT_") > 0 Then
                                    
                        If OutlAttach.FileName Like "*phys*" Then
                            OutlAttach.SaveAsFile CurPath & "_Load\PracsLoad.xlsx"
                        ElseIf OutlAttach.FileName Like "*fac*" Then
                            OutlAttach.SaveAsFile CurPath & "_Load\FacsLoad.xlsx"
                        End If
                    End If
                Next OutlAttach
                '
                '
                '                         OutlSenderLogin = OutlItem.SenderEmailAddress
                '                         OutlSenderName = OutlItem.SenderName
                '                         OutlSenderEMail = OutlItem.Sender.GetExchangeUser.PrimarySmtpAddress
                '                         OutlDateReceived = OutlItem.ReceivedTime
                '                         OutlDateSent = OutlItem.CreationTime
                '                         OutlSubject = OutlItem.Subject
                '                         OutlMsgBody = OutlItem.Body
                '
                '
                '                            Debug.Print OutlSenderLogin
                '                            Debug.Print OutlSenderName
                '                            Debug.Print OutlSenderEMail
                '                            Debug.Print OutlDateReceived
                '                            Debug.Print OutlDateSent
                '                            Debug.Print OutlSubject
                '                            Debug.Print OutlMsgBody
            End If
        Next
    End With
  
  
End Function

'check if outlook is installed, if installed but closed, then open it, then send email
Public Sub SSendEMail()

    Dim xOLApp As Object, V1 As String, VNoOutlook As Boolean
    
    On Error GoTo L1
    
    Set xOLApp = CreateObject("Outlook.Application")
    
    If Not xOLApp Is Nothing Then
         
        Dim OutApp  As Object
        Set OutApp = OutlookApp()
 
        Set OutlNameSpace = OutApp.GetNamespace("MAPI")

        Set OutlMail = OutApp.CreateItem(0)

        With OutlMail

            .To = "gggg@ggggggg.com"
            .Subject = V1 & " cost estimate"
  
            Dim CurFileName As String, CurFileLink As String, NewFileName As String, NewFileLink As String, BodyText As String, BodyTextLinks As String, FinalPath As String
            Dim RsLoopFIlesToEMail As Object
            Dim Attachments As Long, NumberOfLinks As Long, IsZIPFile As Long, VSendTheFile As Long
                
            BodyText1 = "Hello, "
            BodyText3 = "ggggg." & vbCrLf & vbCrLf & _
                        "ALSC Name: " & V1 & vbCrLf & _
                        "Contact: " & V1 & vbCrLf & _
                        "Delivery Address 1: " & V1 & vbCrLf & _
                        "Delivery Address 2: " & V1 & vbCrLf & _
                        "Delivery Address 3: " & V1 & vbCrLf & _
                        "Country: " & DLookup("Country", "CmbCountries", "CountryID=" & V1) & vbCrLf & _
                        "Project Name: " & V1 & vbCrLf
        
            .BodyFormat = 2
            .Body = BodyText1 & vbCrLf & vbCrLf & BodyText3
            ' .Body = BodyText1
         
            Call FSaveReportToFolder("_" & V1 & ".rtf", "ROrder", "RichTextFormat(*.rtf)", False)
            Call FSaveReportToFolder("_" & V1 & ".pdf", "ROrderPdf", "acFormatPDF", False)
        
           ' .Attachments.Add VPDF
          '  .Attachments.Add VRTF
            '  .Display
            .send
            VNoOutlook = True
        End With
        
        Set xOLApp = Nothing
        Exit Sub
    End If
L1:                               VNoOutlook = False
End Sub


Public Function FOutlAttachFiles(VFolder As String, Optional VFileType As String = "*.*")

    ' check if folder exists
    If Dir(VFolder, vbDirectory) = "" Then
        GoTo ExitProc
    End If

    ' check if folder contains specified OutItems
    Dim strFileN As String
    strFileN = Dir(VFolder & VFileType)
    If Len(strFileN) = 0 Then
        GoTo ExitProc
    End If

    ' create mailOutItem
    Set OutlApp = CreateObject("Outlook.Application")
    Set OutlMail = CreateObject("Outlook.MailOutItem")
    Set OutlAttach = CreateObject("Outlook.Attachments")
 
    ' loop through folder and add attachments
    ' if we got this far, strFileN will already contain the
    ' name of the first file in the dir
    Do While Len(strFileN) > 0
        OutlAttach.Add VFolder & strFileN
        strFileN = Dir
    Loop

ExitProc:
    Set OutlApp = Nothing
End Function


