Attribute VB_Name = "MDates"
Option Compare Database
Option Explicit

'get day number from a date
Function FRetrieveDate(VDate As String, VPadWith0 As Boolean) As String
    Dim i     As Integer
    ReDim Months(1 To 12)
    Months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    For i = 0 To 11
        If VDate = Months(i) Then
            If VPadWith0 = True Then
                If i < 10 Then
                    FRetrieveDate = VDate & "0" & i + 1
                Else
                    FRetrieveDate = VDate & i + 1
                End If
            Else
                    FRetrieveDate = VDate & i + 1
            End If
            Exit For
        End If
    Next
End Function

'get month number from month name
Function FRetrieveMonthNumber(VMonthName As String) As Integer

    Select Case VMonthName
    Case "January", "Jan"
        FRetrieveMonthNumber = 1
    Case "February", "Feb"
        FRetrieveMonthNumber = 2
    Case "March", "Mar"
        FRetrieveMonthNumber = 3
    Case "April", "Apr"
        FRetrieveMonthNumber = 4
    Case "May"
        FRetrieveMonthNumber = 5
    Case "June", "Jun"
        FRetrieveMonthNumber = 6
    Case "July", "Jul"
        FRetrieveMonthNumber = 7
    Case "August", "Aug"
        FRetrieveMonthNumber = 8
    Case "September", "Sep"
        FRetrieveMonthNumber = 9
    Case "October", "Oct"
        FRetrieveMonthNumber = 10
    Case "November", "Nov"
        FRetrieveMonthNumber = 11
    Case "December", "Dec"
        FRetrieveMonthNumber = 12
    End Select
End Function

'week number to date
Function FWeekNoToDate(VStartDate As Date, VWeekNo As Double) As String

    Dim VFirstDate As Date, VSecondDate As Date

    If VWeekNo > 52 Or VWeekNo < -52 Then
        MsgBox "Maximum week no is + or - 52", vbCritical
        Exit Function
    End If

    If Int(VWeekNo) <> VWeekNo Then
        MsgBox "Week No must be a whole number", vbCritical
        Exit Function
    End If

    VFirstDate = DateAdd("ww", VWeekNo, VStartDate)
    VSecondDate = DateAdd("d", VFirstDate, 6)
    FWeekNoToDate = Format(VFirstDate, "m/d") & " to " & Format(VSecondDate, "m/d")

End Function

'finds difference between 2 dates

'Print FDiff2Dates("y", #6/1/1998#, #6/26/2002#)
'4 years
'Print FDiff2Dates("ymd", #6/1/1998#, #6/26/2002#)
'4 years 25 days
'Print FDiff2Dates("ymd", #6/1/1998#, #6/26/2002#, True)
'4 years 0 months 25 days
'Print FDiff2Dates("d", #6/1/1998#, #6/26/2002#)
'1486 days
'
'Print FDiff2Dates("h", #1/25/2002 1:23:01 AM#, #1/26/2002 8:10:34 PM#)
'42 hours
'Print FDiff2Dates("hns", #1/25/2002 1:23:01 AM#, #1/26/2002 8:10:34 PM#)
'42 hours 47 minutes 33 seconds
'Print FDiff2Dates("dhns", #1/25/2002 1:23:01 AM#, #1/26/2002 8:10:34 PM#)
'1 day 18 hours 47 minutes 33 seconds
'
'Print FDiff2Dates("ymd", #12/31/1999#, #1/1/2000#)
'1 Day
'Print FDiff2Dates("ymd", #1/1/2000#, #12/31/1999#)
'-1 day
'Print FDiff2Dates("ymd", #1/1/2000#, #1/2/2000#)
'1 Day
'

Public Function FDiff2Dates(VInterval As String, VDate1 As Date, VDate2 As Date, Optional VShowZero As Boolean = False, Optional VYears, Optional VMonths, Optional VDays, Optional VHours, Optional VMinutes, Optional VSeconds) As Variant

    On Error GoTo Err_FDiff2Dates

    Dim booCalcYears As Boolean
    Dim booCalcMonths As Boolean
    Dim booCalcDays As Boolean
    Dim booCalcHours As Boolean
    Dim booCalcMinutes As Boolean
    Dim booCalcSeconds As Boolean
    Dim booSwapped As Boolean
    Dim dtTemp As Date
    Dim intCounter As Integer
    Dim lngDiffYears As Long
    Dim lngDiffMonths As Long
    Dim lngDiffDays As Long
    Dim lngDiffHours As Long
    Dim lngDiffMinutes As Long
    Dim lngDiffSeconds As Long
    Dim varTemp As Variant

    ' Const INTERVALS As String = "dmyhns"
    Const INTERVALs2 As String = "dmyhns"

    'Check that Interval contains only valid characters
    VInterval = LCase$(VInterval)
    For intCounter = 1 To Len(VInterval)
        If InStr(1, INTERVALs2, Mid$(VInterval, intCounter, 1)) = 0 Then
            Exit Function
        End If
    Next intCounter

    'Check that valid dates have been entered
    If Not (IsDate(VDate1)) Then Exit Function
    If Not (IsDate(VDate2)) Then Exit Function

    'If necessary, swap the dates, to ensure that
    'VDate1 is lower than VDate2
    If VDate1 > VDate2 Then
        dtTemp = VDate1
        VDate1 = VDate2
        VDate2 = dtTemp
        booSwapped = True
    End If

    FDiff2Dates = Null
    varTemp = Null

    'What intervals are supplied
    booCalcYears = (InStr(1, VInterval, "y") > 0)
    booCalcMonths = (InStr(1, VInterval, "m") > 0)
    booCalcDays = (InStr(1, VInterval, "d") > 0)
    booCalcHours = (InStr(1, VInterval, "h") > 0)
    booCalcMinutes = (InStr(1, VInterval, "n") > 0)
    booCalcSeconds = (InStr(1, VInterval, "s") > 0)

    'Get the cumulative differences
    If booCalcYears Then
        lngDiffYears = Abs(DateDiff("yyyy", VDate1, VDate2)) - IIf(Format$(VDate1, "mmddhhnnss") <= Format$(VDate2, "mmddhhnnss"), 0, 1)
        VDate1 = DateAdd("yyyy", lngDiffYears, VDate1)
    End If

    If booCalcMonths Then
        lngDiffMonths = Abs(DateDiff("m", VDate1, VDate2)) - IIf(Format$(VDate1, "ddhhnnss") <= Format$(VDate2, "ddhhnnss"), 0, 1)
        VDate1 = DateAdd("m", lngDiffMonths, VDate1)
    End If

    If booCalcDays Then
        lngDiffDays = Abs(DateDiff("d", VDate1, VDate2)) - IIf(Format$(VDate1, "hhnnss") <= Format$(VDate2, "hhnnss"), 0, 1)
        VDate1 = DateAdd("d", lngDiffDays, VDate1)
    End If

    If booCalcHours Then
        lngDiffHours = Abs(DateDiff("h", VDate1, VDate2)) - IIf(Format$(VDate1, "nnss") <= Format$(VDate2, "nnss"), 0, 1)
        VDate1 = DateAdd("h", lngDiffHours, VDate1)
    End If

    If booCalcMinutes Then
        lngDiffMinutes = Abs(DateDiff("n", VDate1, VDate2)) - IIf(Format$(VDate1, "ss") <= Format$(VDate2, "ss"), 0, 1)
        VDate1 = DateAdd("n", lngDiffMinutes, VDate1)
    End If

    If booCalcSeconds Then
        lngDiffSeconds = Abs(DateDiff("s", VDate1, VDate2))
        VDate1 = DateAdd("s", lngDiffSeconds, VDate1)
    End If

    If booCalcYears And (lngDiffYears > 0 Or VShowZero) Then
        varTemp = lngDiffYears & IIf(lngDiffYears <> 1, " years", " year")
    End If

    If booCalcMonths And (lngDiffMonths > 0 Or VShowZero) Then
        If booCalcMonths Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & lngDiffMonths & IIf(lngDiffMonths <> 1, " months", " month")
        End If
    End If

    If booCalcDays And (lngDiffDays > 0 Or VShowZero) Then
        If booCalcDays Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & lngDiffDays & IIf(lngDiffDays <> 1, " days", " day")
        End If
    End If

    If booCalcHours And (lngDiffHours > 0 Or VShowZero) Then
        If booCalcHours Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & lngDiffHours
        End If
    End If

    If booCalcMinutes And (lngDiffMinutes > 0 Or VShowZero) Then
        If booCalcMinutes Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & lngDiffMinutes
        End If
    End If

    If booCalcSeconds And (lngDiffSeconds > 0 Or VShowZero) Then
        If booCalcSeconds Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & lngDiffSeconds
        End If
    End If

    If booSwapped Then
        varTemp = "-" & varTemp
    End If

    FDiff2Dates = Trim$(varTemp)
    VYears = lngDiffYears
    VMonths = lngDiffMonths
    VDays = lngDiffDays
    VHours = lngDiffHours
    VMinutes = lngDiffMinutes
    VSeconds = lngDiffSeconds

End_FDiff2Dates:
    Exit Function

Err_FDiff2Dates:
    Resume End_FDiff2Dates

End Function


Public Function FirstDOW(ByVal dtDate As Date, Optional intWeekBegin As Integer = vbSunday) As Date
    FirstDOW = DateSerial(Year(dtDate), 1, DatePart("y", dtDate, intWeekBegin) - (Weekday(dtDate, intWeekBegin) - 1))
End Function

Public Function LastDOW(ByVal dtDate As Date, Optional intWeekBegin As Integer = vbSunday) As Date
    LastDOW = DateSerial(Year(dtDate), 1, DatePart("y", dtDate, intWeekBegin) + (7 - Weekday(dtDate, intWeekBegin)))
End Function

'Returns the first date of the month for the date passed (dtDate), or the first weekday (DayOfWeek) specified for the month of the date passed.
Public Function FFirstDayOfMonth(VDate As Date, Optional VDayOfWeek As Integer = vbUseSystemDayOfWeek) As Date
    '
    'Example:
    'FirstDayOfMonth(#12/15/2006#) -> Returns the first DATE of the month -> #12/1/2006#
    'FirstDayOfMonth(#12/15/2006#,vbMonday) -> Returns the first MONDAY of the month -> #12/4/2006#
  
    Dim dtTemp As Date
    Dim x     As Byte
  
    'Get the first date of the month for the date passed
    x = 1
    dtTemp = DateSerial(Year(VDate), Month(VDate), x)
  
    If VDayOfWeek <> vbUseSystemDayOfWeek Then
        Do Until Weekday(dtTemp, vbSunday) = VDayOfWeek
            x = x + 1
            dtTemp = DateSerial(Year(VDate), Month(VDate), x)
            If x > 7 Then
                dtTemp = DateSerial(Year(VDate), Month(VDate), 1)
            End If
        Loop
    End If
    
    FFirstDayOfMonth = DateValue(dtTemp)
  
End Function

'finds last day of the month
Function FLastDayOfMonth(VDate As Date) As Date
    FLastDayOfMonth = DateSerial(Year(VDate), Month(VDate) + 1, 0)
End Function

'finds last day of the month
Public Function FLastDayOfMonth2(VDate As Date)
    Dim dFirstDayNextMonth As Date
 
    On Error GoTo lbl_Error
 
    FLastDayOfMonth2 = Empty
    dFirstDayNextMonth = DateSerial(CInt(Format(VDate, "yyyy")), CInt(Format(VDate, "mm")) + 1, 1)
    FLastDayOfMonth2 = DateAdd("d", -1, dFirstDayNextMonth)
 
    Exit Function
lbl_Error:
    MsgBox Err.Description, vbOKOnly + vbExclamation
End Function

