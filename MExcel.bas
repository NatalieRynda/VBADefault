Attribute VB_Name = "MExcel"
Option Compare Database
Option Explicit

Public Function FWorksheetExists(sPath As String, sSheet As String)
    On Error Resume Next
    Dim oExcelApp As Object
    Dim oWB   As Object
    Dim oWS   As Object
    Dim results As Boolean
   
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Workbooks.Open (sPath)
    Set oWS = oExcelApp.Sheets(sSheet)
    If Err Then
        results = False
    Else
        results = True
    End If

    Set oWS = Nothing
    oExcelApp.Quit
    Set oExcelApp = Nothing
    FWorksheetExists = results

End Function

