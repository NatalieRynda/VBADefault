Attribute VB_Name = "MRefreshDropDowns"
Option Compare Database
Option Explicit

Public Function FRefreshDropDowns(VFieldName As String, VDropDownFieldID As Long)

    On Error GoTo ErrProc
              
    Dim RsCheckForFields As DAO.Recordset
    Dim VMainForm As String, VSubForm1 As String, VSubForm2 As String

    Set RsCheckForFields = CurrentDb.OpenRecordset("SELECT MainForm, Subform1, Subform2 FROM DropDownForms WHERE DropDownFieldID=" & VDropDownFieldID)
                
    Do Until RsCheckForFields.EOF
         
        VMainForm = vbNullString
        VSubForm1 = vbNullString
        VSubForm2 = vbNullString
         
        If Not IsNull(RsCheckForFields("MainForm")) Then
            VMainForm = RsCheckForFields("MainForm")
        End If
         
        If Not IsNull(RsCheckForFields("Subform1")) Then
            VSubForm1 = RsCheckForFields("Subform1")
        End If
         
        If Not IsNull(RsCheckForFields("Subform2")) Then
            VSubForm2 = RsCheckForFields("Subform2")
        End If
         
        If VMainForm <> vbNullString Then
            If CurrentProject.AllForms(VMainForm).IsLoaded Then
                  
                '                If VRequeryForm = True Then
                '                    Forms(VMainForm).Requery
                '                Else
                If VSubForm1 = vbNullString Then
                    Forms(VMainForm).SetFocus
                    DoCmd.RunCommand acCmdSaveRecord
                    Forms(VMainForm)(VFieldName).Requery
                ElseIf VSubForm1 <> vbNullString And VSubForm2 = vbNullString Then
                    Forms(VMainForm).Form(VSubForm1).SetFocus
                    DoCmd.RunCommand acCmdSaveRecord
                    Forms(VMainForm).Form(VSubForm1)(VFieldName).Requery
                ElseIf VSubForm2 <> vbNullString Then
                    Forms(VMainForm).Form(VSubForm1).Form(VSubForm2).SetFocus
                    DoCmd.RunCommand acCmdSaveRecord
                    Forms(VMainForm).Form(VSubForm1).Form(VSubForm2)(VFieldName).Requery
                End If
                'End If
                           
            End If
        End If
         
        RsCheckForFields.MoveNext
    Loop
         
    RsCheckForFields.Close
    Set RsCheckForFields = Nothing

Exit_ErrProc:                     Exit Function
ErrProc:                     FLogError Err.Number, Err.Description, "MRefreshDropDowns.RefreshDropDowns()"
    Resume Exit_ErrProc

End Function

