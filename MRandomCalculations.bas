Attribute VB_Name = "MRandomCalculations"
Option Compare Database
Option Explicit

Public intNumNames As Integer                    'largest number in the list (1-70, so 70)
Public intNumPlayers As Integer                  'how many numbers should the combinations consist of (for example, 2 will give you 1-2,1-21,1-34, etc)
Public intPlayerNum As Integer                   'The number of the player whose name is being exported, effectively the Field of the output
Public intRow As Integer                         'A number holding the row number for the next combination
Public i, j, k                                   'counters
Public intDepth As Integer                       'Number of loops that are running
Public intLoopRange As Integer                   'Number of iterations in any one loop
Public arr()  As Integer                         'An array of variables that are used to recursively loop
Public strOutput As String                       'An output string for debugging purposes- easire to read numbers and check them. Note that there are only so many lines in the debug window, so you may not see
'all results for a large number of combinations
Public intexpected As Integer                    'Expected number of combinations

Public Function OutputNames()
  
    intNumNames = 30
    intNumPlayers = 2
    intLoopRange = intNumNames - intNumPlayers + 1
  
    Do
        AnotherLoop (0)
    Loop While intDepth > 0
  
    'intexpected = Combin(intNumNames, intNumPlayers)

End Function

Public Function AnotherLoop(a As Integer)
    intDepth = intDepth + 1
    ReDim Preserve arr(intDepth)
    'For arr(intDepth) = a + 1 To DMin(a + intLoopRange, intNumNames)
    For arr(intDepth) = a + 1 To (a + intLoopRange)
        If intDepth = intNumPlayers Then
            intRow = intRow + 1
            strOutput = Format(intRow, "000") & Space(3)
      
            For j = 1 To intNumPlayers
                'Range("A1").Offset(intRow - 1, j - 1) = rngNames(arr(j))
                strOutput = strOutput & arr(j) & Space(3)
                MsgBox strOutput
            Next j
     
        Else
            AnotherLoop (arr(intDepth))
        End If
    Next arr(intDepth)
    intDepth = intDepth - 1
End Function


Function CalcDist(dblLat1 As Double, dblLon1 As Double, dblLat2 As Double, dblLon2 As Double) As Double

    ' Calculate the distance between two latitudes and longitudes
    Const cnsPI = 3.1415926535

    Dim dblRadLat1 As Double
    Dim dblRadLat2 As Double
    Dim dblRadLon1 As Double
    Dim dblRadLon2 As Double
    Dim dblTheta As Double
    Dim dblRadTheta As Double
    Dim dblDist As Double

    '<cfset radlat1 = Evaluate((pi() * lat1)/180)>
    '<cfset radlat2 = Evaluate((pi() * lat2)/180)>
    '<cfset radlon1 = Evaluate((pi() * lon1)/180)>
    '<cfset radlon2 = Evaluate((pi() * lon2)/180)>
    dblRadLat1 = cnsPI * dblLat1 / 180
    dblRadLat2 = cnsPI * dblLat2 / 180
    dblRadLon1 = cnsPI * dblLon1 / 180
    dblRadLon2 = cnsPI * dblLon2 / 180
    '<cfset theta = lat1-lat2>
    '<cfset radtheta = Evaluate((pi() * theta)/180)>
    dblTheta = dblLat1 - dblLat2
    dblRadTheta = cnsPI * dblTheta / 180
    '<cfset dist = Evaluate((60 * 1.1515) * (180 / pi()) * (ACos((Sin(radlat1) * Sin(radlat2)) + (Cos(radlat1) * Cos(radlat2) * Cos(radtheta)))))>
    CalcDist = (60 * 1.1515) * (180 / cnsPI) * Acos((Sin(dblRadLat1) * Sin(dblRadLat2)) + (Cos(dblRadLat1) * Cos(dblRadLat2) * Cos(dblRadTheta)))
End Function

Function Acos(dblRadian As Double) As Double
    ' Compute the Arc Cosine
    Acos = Atn(-dblRadian / Sqr(-dblRadian * dblRadian + 1)) + 2 * Atn(1)
End Function


