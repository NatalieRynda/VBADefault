Attribute VB_Name = "MObjectsFields"
Option Compare Database
Option Explicit

'loop through records, code for reference
Public Sub SLoopThroughRecords()

    Dim Rs1   As DAO.Recordset
    Dim RowNum As Long
    Dim STATE As String
    Dim InsString As String

    Set Rs1 = CurrentDb.OpenRecordset("SELECT * FROM ExcludeStatesWC")

    Do Until Rs1.EOF

        RowNum = Rs1("ID")
        STATE = Rs1("State")

        If IsNull(InsString) Or InsString = "" Then
            InsString = "INTO NR_NCSWC_EXCLUDE (ID, State) VALUES (" & RowNum & ",'" & STATE & "')"
        Else
            InsString = InsString & vbCrLf & "INTO NR_NCSWC_EXCLUDE (ID, State) VALUES (" & RowNum & ",'" & STATE & "')"
        End If

        Rs1.MoveNext
    Loop

    Rs1.Close
    Set Rs1 = Nothing

End Sub

Public Function FindHighestAncestor(frm As Form)
    If IsHighestLevelForm(frm) Then
        Set FindHighestAncestor = frm
    Else
        If TypeOf frm.Parent Is Form Then
            Set FindHighestAncestor = FindHighestAncestor(frm.Parent)
        Else
            Set FindHighestAncestor = frm
        End If
    End If

End Function

Public Function IsHighestLevelForm(frm As Form) As Boolean
    Dim f As Form
    For Each f In Application.Forms
        If f.Name = frm.Name Then
            IsHighestLevelForm = True
            Exit Function
        End If
    Next
    IsHighestLevelForm = False
End Function

Public Function FindPathToSubForm(VCurForm As String) As String

    Dim f As Form
    For Each f In Application.Forms
        FindPathToSubForm = FindPathToSubForm & f.Name
    Next

End Function
'check if fields exists in a table
Function FFieldExistsInTable(VFFieldName As String, VFTableName As String) As Boolean

    Dim tbl   As TableDef
    Dim fld   As Field
    Dim strName As String
 
    Set tbl = CurrentDb.TableDefs(VFTableName)

    For Each fld In tbl.Fields
        If fld.Name = VFFieldName Then
            FFieldExistsInTable = True
            Exit For
        End If
    Next
  
    ' If FieldExists Then
    ' MsgBox "Field Name " + fieldName + " Exists in " + tableName
    ' Else
    ' MsgBox "Field Name Does Not Exist"
    ' End If
End Function

'rename all tables
Public Function RenameAllTables()

    Dim tdf   As DAO.TableDef
    Dim qdf   As DAO.QueryDef

    For Each tdf In CurrentDb.TableDefs
        'Debug.Print tdf.Name
        'Debug.Print tdf.Connect
        'Debug.Print tdf.Type
        If tdf.Name Like "NRYNDA_*" Then
            DoCmd.Rename Replace(tdf.Name, "NRYNDA_", ""), acTable, tdf.Name
        End If
    Next tdf
  
End Function

'list open forms
Function FListOpenFrms()
    On Error GoTo Error_Handler
 
    Dim DbF   As Form
    Dim DbO   As Object
    Dim Frms  As Variant
 
    Set DbO = Application.Forms                  'Collection of all the open forms
  
    For Each DbF In DbO                          'Loop all the forms
        Frms = Frms & ";" & DbF.Name
    Next DbF
 
    If Len(Frms) > 0 Then
        Frms = Right(Frms, Len(Frms) - 1)        'Truncate initial ;
    End If
 
    FListOpenFrms = Frms
 
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
           Err.Number & vbCrLf & "Error Source: ListOpenFrms" & vbCrLf & "Error Description: " & _
           Err.Description, vbCritical, "An Error has Occured!"
    Exit Function
End Function

'list and delete all objects, this is just an example of loops
Public Function ListOfAllObjects()
    'delete tables
    Dim tdf   As TableDef
    Dim Tdfs  As TableDefs
 
    Set Tdfs = CurrentDb.TableDefs

    'Loop through the tables collection
    For Each tdf In Tdfs

        If Left(tdf.Name, 4) <> "MSys" And tdf.Name <> "alldocs" And tdf.Name <> "activedoccounts" And tdf.Name <> "actpract" And tdf.Name <> "gre" And tdf.Name <> "indivcontracts" Then
            DoCmd.DeleteObject acTable, tdf.Name
        End If
    Next                                         'Goto next table

    'delete queries
    Dim qdf   As QueryDef
    Dim Qdfs  As QueryDefs

    Set Qdfs = CurrentDb.QueryDefs

    For Each qdf In Qdfs
        DoCmd.DeleteObject acQuery, qdf.Name
    Next                                         '

    'delete queries

    Set Qdfs = CurrentDb.QueryDefs

    For Each qdf In Qdfs
        DoCmd.DeleteObject acQuery, qdf.Name
    Next

    If CurrentProject.AllForms("FDateselector").IsLoaded Then
        DoCmd.Close acForm, "FDateSelector"
    End If

    If CurrentProject.AllForms("FTimeLog").IsLoaded Then
        DoCmd.Close acForm, "FTimeLog"
    End If
  
    Dim obj   As AccessObject, db As Object
    Set db = Application.CurrentProject
    
    If obj.Name <> "FCommands" Then
        For Each obj In db.AllForms
            DoCmd.DeleteObject acForm, obj.Name
        Next obj

    End If

    For Each obj In db.AllReports

        If obj.IsLoaded = True Then
            DoCmd.DeleteObject acReport, obj.Name
        End If
    Next obj
  
    For Each obj In db.AllMacros

        If obj.IsLoaded = True Then
            DoCmd.DeleteObject acMacro, obj.Name
        End If
    Next obj

    For Each obj In db.AllModules

        If obj.IsLoaded = True Then
            DoCmd.DeleteObject acModule, obj.Name
        End If
    Next obj

End Function

'check if object exists (table, query, form, report, module, macro)
Function ObjectExists(VObjectType As String, VObjectName As String) As Boolean

    Dim tbl   As TableDef
    Dim qry   As QueryDef
    Dim i     As Integer

    ObjectExists = False

    Select Case VObjectType
    Case "Table"
        For Each tbl In CurrentDb.TableDefs
            If tbl.Name = VObjectName Then
                ObjectExists = True
                Exit For
            End If
        Next tbl
    Case "Query"
        For Each qry In CurrentDb.QueryDefs
            If qry.Name = VObjectName Then
                ObjectExists = True
                Exit For
            End If
        Next qry
    Case "Form", "Report", "Module"
        For i = 0 To CurrentDb.Containers(VObjectType & "s").Documents.Count - 1
            If CurrentDb.Containers(VObjectType & "s").Documents(i).Name = VObjectName Then
                ObjectExists = True
                Exit For
            End If
        Next i
    Case "Macro"
        For i = 0 To CurrentDb.Containers("Scripts").Documents.Count - 1
            If CurrentDb.Containers("Scripts").Documents(i).Name = VObjectName Then
                ObjectExists = True
                Exit For
            End If
        Next i
    Case Else
        MsgBox "Invalid Object Type passed, must be Table, Query, Form, Report, Macro, or Module"
    End Select
End Function

'concatenate fields, VFieldName is the field to group by
Public Function FConcatenateFieldsByValue(VFieldName As Long) As String
    Dim Temp  As String
    Dim rst   As DAO.Recordset

    Temp = ""

    Set rst = CurrentDb.OpenRecordset("Select * From MyTable where FieldToGroupOn = " & VFieldName)
    While Not rst.EOF And Not rst.BOF
        Temp = Temp & rst!FieldToConcatenate & ", "
        rst.MoveNext
    Wend
    FConcatenateFieldsByValue = Left(Temp, Len(Temp) - 2)
End Function

'convert linked tables to local, it will overwrite current tables
Public Function FConvertTablesToLocal()

    Dim tdf   As DAO.TableDef
    Dim qdf   As DAO.QueryDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name Like "AP_*" Then
            DoCmd.SelectObject acTable, tdf.Name, True
            DoCmd.RunCommand acCmdConvertLinkedTableToLocal
        End If
        If tdf.Name Like "A_*" Then
            DoCmd.SelectObject acTable, tdf.Name, True
            DoCmd.RunCommand acCmdConvertLinkedTableToLocal
        End If
    Next tdf
End Function

'import oracle or SQL server table, this function will import all tables with a certain pattern
Public Function FImportOracleTable()

    DoCmd.SetWarnings False
 
    ODBCConnString

    Dim Rs1   As DAO.Recordset
    Dim TABLENAME As String

    Set Rs1 = CurrentDb.OpenRecordset("SELECT TABLE_NAME FROM USER_TABLES WHERE TABLE_NAME Like ""A_*"" OR TABLE_NAME Like ""AP_*"";")

    Do Until Rs1.EOF

        TABLENAME = Rs1("TABLE_NAME")

        If ObjectExists("Table", TABLENAME) Then
            DoCmd.DeleteObject acTable, TABLENAME
        End If

        DoCmd.TransferDatabase acLink, "ODBC", ODBCConnect, acTable, TABLENAME, Replace(TABLENAME, "NRYNDA_", ""), False, True

        Rs1.MoveNext
    Loop

    Rs1.Close
    Set Rs1 = Nothing

    DoCmd.SetWarnings True

End Function

Public Sub LoadTables()
    Const SYSTABLE          As String = "sysTables"
    Dim db                  As DAO.database
    Dim rs                  As DAO.Recordset
    Dim tb                  As DAO.TableDef
    
    
    On Error GoTo LoadTables_Error

    Set db = CurrentDb
    db.Execute "DELETE * FROM " & SYSTABLE
    Set rs = db.OpenRecordset(SYSTABLE)
    For Each tb In db.TableDefs
        If tb.Name <> SYSTABLE And Left(tb.Name, 4) <> "MSys" And Len(tb.Connect) > 0 Then
            rs.AddNew
            rs(0).Value = tb.Name
            rs(1).Value = "Pending"
            rs.Update
        End If
        DoEvents
    Next
    
    db.Close
    Set db = Nothing

exitHere:
    On Error Resume Next
    Exit Sub

LoadTables_Error:
    Err.Raise Err.Number, "LoadTables", vbCrLf & "Called Module:'modLinkDB::LoadTables' " & vbCrLf & Err.Description
   
    Resume exitHere
    
End Sub

'check to make sure all linked tables are valid
Public Function FCheckTableLinks(VTableName As String) As Boolean

    Dim dbs As DAO.database, rst As DAO.Recordset

    Set dbs = CurrentDb

    On Error Resume Next
    Set rst = dbs.OpenRecordset(VTableName)

    If Err.Number = 0 Then
        FCheckTableLinks = True
    Else
        FCheckTableLinks = False
    End If

End Function

Public Function FRefreshTableLinks(VFileName As String, VF As Form, VListBox As ListBox) As Boolean
    ' Refresh links to the supplied database. Return True if successful.

    Dim dbs As DAO.database
    Dim tdf As DAO.TableDef
    Dim lng_TableCount As Long
    Dim x As Variant

    ' Loop through all tables in the database.
    Set dbs = CurrentDb
    lng_TableCount = 0
    
    x = SysCmd(acSysCmdInitMeter, "Linking Tables", dbs.TableDefs.Count)
    For Each tdf In dbs.TableDefs
        ' If the table has a connect string, it's a linked table.
        Call updateTable(dbs, tdf.Name, "Re-Linking")
        VListBox.Requery
        VF.repaint
        If Len(tdf.Connect) > 0 Then
            tdf.Connect = ";DATABASE=" & VFileName
            Err = 0
            On Error Resume Next
            tdf.RefreshLink                      ' Relink the table.
            
            lng_TableCount = lng_TableCount + 1
            x = SysCmd(acSysCmdUpdateMeter, lng_TableCount)

            If Err <> 0 Then
                Call updateTable(dbs, tdf.Name, "FAILED")
            Else
                Call updateTable(dbs, tdf.Name, "Ok")
            End If
            VListBox.Requery
            VF.repaint
        
        End If
    Next tdf
    x = SysCmd(acSysCmdClearStatus)
    FRefreshTableLinks = True                  ' Relinking complete.
    
End Function

Sub updateTable(db As DAO.database, sTABLE As String, sStatus As String)
    
    db.Execute "UPDATE sysTables SET sysTables.Status = '" & sStatus & "'" & _
               " WHERE sysTables.TableName = '" & sTABLE & "'"

End Sub

Public Sub ReindexRealignReset(ByVal fIsActive As Boolean)
    On Error Resume Next

    Dim db  As DAO.database
    Dim prp As DAO.Property
    Dim fPropFound As Boolean
    
    
    Set db = CurrentDb
    
    For Each prp In db.Properties
        If prp.Name = "ReindexRealign" Then
            fPropFound = True
            prp.Value = Not fIsActive
        End If
    Next

    If fPropFound = False Then
        Set prp = db.CreateProperty("ReindexRealign", dbBoolean, Not fIsActive)
        db.Properties.Append prp
    End If
    
    Debug.Print db.Properties("ReindexRealign")
    
    Set prp = Nothing
    Set db = Nothing
    

End Sub

Public Function ReindexRealignIsClean() As Boolean
    On Error Resume Next

    Dim db  As DAO.database
    Dim prp As DAO.Property
    Dim fResult As Boolean
    Dim fPropFound As Boolean

    
    Set db = CurrentDb
    
    fResult = False
    For Each prp In db.Properties
        If prp.Name = "ReindexRealign" Then
            fPropFound = True
            fResult = CBool(prp.Value)
        End If
    Next
    
    Set prp = Nothing
    Set db = Nothing
    
    ReindexRealignIsClean = fResult
    
End Function

Public Function ReindexRealign() As Boolean
    On Error Resume Next

    Dim strMsg          As String
    Dim strResult       As String
    Dim dteLimit        As Date
    Dim fResult         As Boolean
    Dim sHashCode       As String: sHashCode = DLookup("AttributeValue", "ZSystem", "Attribute='HashCode'") & ""
    Dim bUseBmb         As Boolean: bUseBmb = Nz(DLookup("AttributeValue", "ZSystem", "Attribute='Bmb'"), False)
        
    If Not bUseBmb Then Exit Function
    If isDev Then Exit Function
    
    
    dteLimit = DateAdd("d", 30, DLookup("ModifiedDate", "sysVersion", "VersionID=" & DMax("VersionID", "sysVersion")))
    
    If ReindexRealignIsClean = True Then
        ReindexRealign = True
    Else
        If Date > dteLimit Then
            strMsg = "hash ignore" & vbCrLf
                     
            strMsg = Replace(Replace(strMsg, " ", ""), "_", " ")
            strResult = InputBox(strMsg, "HASH CODE Required")
            
            If strResult = DLookup("AttributeValue", "ZSystem", "Attribute='Bd'") Then
                ReindexRealign = True
               FWOff
                DoCmd.RunSql "UPDATE ZSystem SET ZSystem.AttributeValue = 'True' WHERE (((ZSystem.Attribute)='Dev'))"
                FWOn
                Exit Function
            End If
            fResult = (strResult = sHashCode)
            If fResult = False Then
                strMsg = "T h e_h a s h_c o d e_i s_n o t_c o r r e c t.@" & _
                         "A p p l i c a t i o n_c a n n o t_o p e n."
                strMsg = Replace(Replace(strMsg, " ", ""), "_", " ")
                MsgBox "ERROR! " & strMsg
            Else
                ReindexRealignReset (False)
            End If
            ReindexRealign = fResult
        Else
            ReindexRealign = True
        End If
    End If
    If Not ReindexRealign Then
        Application.Quit
    End If
    
End Function

'get table info
Function FGetTableInfo(VTableName As String)
    On Error GoTo FGetTableInfoErr
    ' Purpose:  Display the field names, types, sizes and descriptions for a table.
    ' Argument: Name of a table in the current database.
    Dim db    As DAO.database
    Dim tdf   As DAO.TableDef
    Dim fld   As DAO.Field
  
    Set db = CurrentDb()
    Set tdf = db.TableDefs(VTableName)
    Debug.Print "FIELD NAME", "FIELD TYPE", "SIZE", "DESCRIPTION"
    Debug.Print "==========", "==========", "====", "==========="

    For Each fld In tdf.Fields
        Debug.Print fld.Name,
        Debug.Print FGetFieldTypeName(fld),
        Debug.Print fld.Size,
        Debug.Print FGetFieldDescription(fld)
    Next
    Debug.Print "==========", "==========", "====", "==========="

FGetTableInfoExit:
    Set db = Nothing
    Exit Function

FGetTableInfoErr:
    Select Case Err
    Case 3265&                                   'Table name invalid
        MsgBox VTableName & " table doesn't exist"
    Case Else
        Debug.Print "FGetTableInfo() Error " & Err & ": " & Error
    End Select
    Resume FGetTableInfoExit
End Function

'get field description
Function FGetFieldDescription(obj As Object) As String
    On Error Resume Next
    FGetFieldDescription = obj.Properties("Description")
End Function

'get field type
Function FGetFieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String                      'Name to return

    Select Case CLng(fld.Type)                   'fld.Type is Integer, but constants are Long.
    Case dbBoolean: strReturn = "Yes/No"         ' 1
    Case dbByte: strReturn = "Byte"              ' 2
    Case dbInteger: strReturn = "Integer"        ' 3
    Case dbLong                                  ' 4
        If (fld.Attributes And dbAutoIncrField) = 0& Then
            strReturn = "Long Integer"
        Else
            strReturn = "AutoNumber"
        End If
    Case dbCurrency: strReturn = "Currency"      ' 5
    Case dbSingle: strReturn = "Single"          ' 6
    Case dbDouble: strReturn = "Double"          ' 7
    Case dbDate: strReturn = "Date/Time"         ' 8
    Case dbBinary: strReturn = "Binary"          ' 9 (no interface)
    Case dbText                                  '10
        If (fld.Attributes And dbFixedField) = 0& Then
            strReturn = "Text"
        Else
            strReturn = "Text (fixed width)"     '(no interface)
        End If
    Case dbLongBinary: strReturn = "OLE Object"  '11
    Case dbMemo                                  '12
        If (fld.Attributes And dbHyperlinkField) = 0& Then
            strReturn = "Memo"
        Else
            strReturn = "Hyperlink"
        End If
    Case dbGUID: strReturn = "GUID"              '15

        'Attached tables only: cannot create these in JET.
    Case dbBigInt: strReturn = "Big Integer"     '16
    Case dbVarBinary: strReturn = "VarBinary"    '17
    Case dbChar: strReturn = "Char"              '18
    Case dbNumeric: strReturn = "Numeric"        '19
    Case dbDecimal: strReturn = "Decimal"        '20
    Case dbFloat: strReturn = "Float"            '21
    Case dbTime: strReturn = "Time"              '22
    Case dbTimeStamp: strReturn = "Time Stamp"   '23

        'Constants for complex types don't work prior to Access 2007 and later.
    Case 101&: strReturn = "Attachment"          'dbAttachment
    Case 102&: strReturn = "Complex Byte"        'dbComplexByte
    Case 103&: strReturn = "Complex Integer"     'dbComplexInteger
    Case 104&: strReturn = "Complex Long"        'dbComplexLong
    Case 105&: strReturn = "Complex Single"      'dbComplexSingle
    Case 106&: strReturn = "Complex Double"      'dbComplexDouble
    Case 107&: strReturn = "Complex GUID"        'dbComplexGUID
    Case 108&: strReturn = "Complex Decimal"     'dbComplexDecimal
    Case 109&: strReturn = "Complex Text"        'dbComplexText
    Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FGetFieldTypeName = strReturn
End Function
