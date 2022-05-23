Attribute VB_Name = "MLogError"
Option Compare Database
Option Explicit

Function FLogError(VErrNumber As Long, VErrDescription As String, VCallingProc As String, Optional VParameters, Optional VShowUser As Boolean = True) As Boolean

    On Error GoTo Err_FLogError

    Dim strMsg As String
    Dim rst   As DAO.Recordset
    
    Set rst = CurrentDb.OpenRecordset("LogError", dbOpenDynaset)

    rst.AddNew
    rst![ErrNumber] = VErrNumber
    rst![ErrDescription] = VErrDescription
    rst![ErrDate] = Now()
    rst![CallingProc] = VCallingProc
    rst![UserName] = FUserName()
    If Not IsMissing(VParameters) Then
        rst![Parameters] = VParameters
    End If
    rst.Update
    rst.Close
    FLogError = True

Exit_FLogError:               Set rst = Nothing
    Exit Function

Err_FLogError:
    Resume Exit_FLogError
End Function

