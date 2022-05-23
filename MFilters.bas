Attribute VB_Name = "MFilters"
Option Compare Database
Option Explicit
Function Filters()
    On Error GoTo Filters_Err

    'Beep
    ' Macro can't be run from the navigation pane.
    Exit Function
    DoCmd.SetProperty "cboFilterFavorites", acPropertyVisible, "0"

Filters_Exit:
    Exit Function

Filters_Err:
    MsgBox Error$
    Resume Filters_Exit

End Function
Function Filters_ApplyFilterFavorite(frm As Form)
    On Error GoTo Filters_ApplyFilterFavorite_Err

    With frm
        If (IsNull(Screen.ActiveControl) Or Screen.ActiveControl = 0) Then
            ' Clear Filters.
            Filters_ClearFilter frm
            Exit Function
        End If
        If (Screen.ActiveControl = -1) Then
            ' Manage Filters.
            Filters_Manage
            Exit Function
        End If
        ' Apply Filters
        TempVars.Add "FilterString", DLookup("[Filter String]", "Filters", "ID = " & Screen.ActiveControl)
        TempVars.Add "SortString", DLookup("[Sort String]", "Filters", "ID = " & Screen.ActiveControl)
        If (Not IsNull(TempVars!FilterString)) Then
            DoCmd.ApplyFilter "", TempVars!FilterString, ""
        End If
        If (CurrentProject.IsTrusted And Not IsNull(TempVars!SortString)) Then
            .OrderBy = Nz(TempVars!SortString)
        End If
        If (CurrentProject.IsTrusted) Then
            .OrderByOn = Not IsNull(TempVars!SortString)
        End If
        TempVars.Remove "FilterString"
        TempVars.Remove "SortString"
    End With

Filters_ApplyFilterFavorite_Exit:
    Exit Function

Filters_ApplyFilterFavorite_Err:
    MsgBox Error$
    Resume Filters_ApplyFilterFavorite_Exit

End Function

Function Filters_New(frm As Form)
    On Error GoTo Filters_New_Err

    With frm
        Filters_CheckFilter frm
        Filters_SetupTempVars frm
        DoCmd.OpenForm "fFilterDetails", acNormal, "", "", acAdd, acDialog
        frm!cboFilterFavorites.Requery
        
        !cboFilterFavorites = TempVars!LastFilterCreated
        
        Filters_RemoveTempVars
    End With

Filters_New_Exit:
    Exit Function

Filters_New_Err:
    MsgBox Error$
    Resume Filters_New_Exit

End Function
Function Filters_SetLastFilterID()
    On Error GoTo Filters_SetLastFilterID_Err

    With CodeContextObject
        ' Used in conjunction with NEW to set the value of the combo box when trusted.
        TempVars.Add "LastFilterCreated", .ID
    End With

Filters_SetLastFilterID_Exit:
    Exit Function

Filters_SetLastFilterID_Err:
    MsgBox Error$
    Resume Filters_SetLastFilterID_Exit

End Function
Function Filters_Manage()
    On Error GoTo Filters_Manage_Err

    TempVars.Add "ObjectType", Application.CurrentObjectType
    TempVars.Add "ObjectName", Application.CurrentObjectName
    DoCmd.OpenForm "fFilterDetails", acNormal, "", "[Object Name]=[Application].[CurrentObjectName]", , acDialog
    DoCmd.RunCommand acCmdRefresh
    TempVars.Remove "ObjectType"
    TempVars.Remove "ObjectName"

Filters_Manage_Exit:
    Exit Function

Filters_Manage_Err:
    MsgBox Error$
    Resume Filters_Manage_Exit

End Function

Function Filters_SetupTempVars(frm As Form)
    On Error GoTo Filters_SetupTempVars_Err

    With frm
        TempVars.Add "ObjectType", "Form"
        TempVars.Add "ObjectName", .Name
        TempVars.Add "FilterString", .Filter
        TempVars.Add "SortString", .OrderBy
    End With

Filters_SetupTempVars_Exit:
    Exit Function

Filters_SetupTempVars_Err:
    MsgBox Error$
    Resume Filters_SetupTempVars_Exit

End Function
Function Filters_RemoveTempVars()
    On Error GoTo Filters_RemoveTempVars_Err

    TempVars.Remove "ObjectType"
    TempVars.Remove "ObjectName"
    TempVars.Remove "FilterString"
    TempVars.Remove "SortString"
    TempVars.Remove "Order"
    TempVars.Remove "LastFilterCreated"

Filters_RemoveTempVars_Exit:
    Exit Function

Filters_RemoveTempVars_Err:
    MsgBox Error$
    Resume Filters_RemoveTempVars_Exit

End Function
Function Filters_ClearFilter(frm As Form)
    On Error GoTo Filters_ClearFilter_Err

    With frm
        
        ' Clear Filter
        On Error Resume Next
        DoCmd.ApplyFilter "", """""", ""
        DoCmd.GoToControl "SearchBox"
        DoCmd.SetProperty "SearchClear", acPropertyVisible, "0"
        DoCmd.SetProperty "SearchGo", acPropertyVisible, "-1"
        .SearchBox = ""
        
    End With

Filters_ClearFilter_Exit:
    Exit Function

Filters_ClearFilter_Err:
    MsgBox Error$
    Resume Filters_ClearFilter_Exit

End Function
Function Filters_CheckFilter(frm As Form)
    On Error GoTo Filters_CheckFilter_Err

    With frm
        If (Not (.FilterOn)) Then
            ' If the filter is off, clear the last saved filter
            .ApplyFilter "", """""", ""
        End If
        If (Not (.FilterOn Or .OrderByOn)) Then
            Beep
            MsgBox "You don't have a filter or sort to save.", vbOKOnly, "Save Filter"
            End
        End If
    End With

Filters_CheckFilter_Exit:
    Exit Function

Filters_CheckFilter_Err:
    MsgBox Error$
    Resume Filters_CheckFilter_Exit

End Function
Function Filters_Search()
    On Error GoTo Filters_Search_Err

    DoCmd.RunMacro "Search", , ""

Filters_Search_Exit:
    Exit Function

Filters_Search_Err:
    MsgBox Error$
    Resume Filters_Search_Exit

End Function
