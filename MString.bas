Attribute VB_Name = "MString"
Option Compare Database
Option Explicit

Dim x         As Integer
Dim PadLength As Integer

'pad with 0 from left
Function LPad(MyValue$, MyPadCharacter$, MyPaddedLength%)

    PadLength = MyPaddedLength - Len(MyValue)
    Dim PadString As String
    For x = 1 To PadLength
        PadString = PadString & MyPadCharacter
    Next
    LPad = PadString + MyValue

End Function

'pad with 0 from right
Function RPad(MyValue$, MyPadCharacter$, MyPaddedLength%)

    PadLength = MyPaddedLength - Len(MyValue)
    Dim PadString As String
    For x = 1 To PadLength
        PadString = MyPadCharacter & PadString
    Next
    RPad = MyValue + PadString

End Function

'see if a string has letters or numbers, anything other than a letter or number will give an error
Function IsLetterOrNumber(strValue As String) As Boolean

    Dim intPos As Integer

    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
        Case 65 To 90, 97 To 122, 48 To 57
            IsLetterOrNumber = True
        Case Else
            IsLetterOrNumber = False
            Exit For
        End Select
    Next

End Function

'see if a string has numbers, anything other than a number will give an error
Function IsNumber(strValue As String) As Boolean

    Dim intPos As Integer

    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
        Case 48 To 57
            IsNumber = True
        Case Else
            IsNumber = False
            Exit For
        End Select
    Next

End Function

'remove everyting except letters or numbers
'Option 1 - only leave letters or numbers
'Option 2 - only leave letters or numbers or spaces
Function FStrip(VString As Variant, VOption As Long)

    Dim Clean As String
    Dim Pos, A_Char$

    Pos = 1
    If IsNull(VString) Then Exit Function

    For Pos = 1 To Len(VString)
        A_Char$ = Mid(VString, Pos, 1)
        If VOption = 1 Then
            If (A_Char$ >= "0" And A_Char$ <= "9") Or (A_Char$ >= "a" And A_Char$ <= "z") Then
                Clean$ = Clean$ + A_Char$
            End If
        ElseIf VOption = 2 Then
            If (A_Char$ >= "0" And A_Char$ <= "9") Or (A_Char$ >= "a" And A_Char$ <= "z") Or A_Char$ = " " Then
                Clean$ = Clean$ + A_Char$
            End If
        End If
    Next Pos

    FStrip = Clean$

End Function

'remove everyting except numbers
Public Function FStripNonAlphaNum(VString As String) As String
    Dim n     As Long, b() As Byte
    b = VString
    For n = 0 To UBound(b) Step 2
        Select Case b(n)
        Case 48 To 57, 65 To 90, 97 To 122
        Case Else
            b(n) = 0
        End Select
    Next n
    FStripNonAlphaNum = Replace(b, vbNullChar, vbNullString)
End Function

'removes the rich text characters from a string
Public Function RemoveRichText(VString As Variant) As Variant
    'for the commented part need refernence to Microsoft VBscript regular expressions
    'Dim rgx As New VBScript_RegExp_55.RegExp
 
    Dim rgx   As Object

    Set rgx = CreateObject("vbscript.regexp")
  
    If Not IsNull(VString) Then
        rgx.Pattern = "<[^>]+>"
        'Global: Sets a Boolean value or returns a Boolean value that indicates whether a pattern must match all the occurrences in a whole search string,
        'or whether a pattern must match just the first occurrence
        rgx.Global = True
        'Multiline: enables the regular expression engine to handle an input string that consists of multiple lines. It changes the interpretation of the ^ and $ language elements so that
        'they match the beginning and end of a line, instead of the beginning and end of the input string.
        rgx.MultiLine = True
        VString = rgx.Replace(VString, "")
    End If
    RemoveRichText = VString
  
End Function

'replace characters in a string
Function FReplaceChars(VString As String) As String
    Const REPLACE_WITH As String = ""
    Dim varChars As Variant
    Dim var   As Variant
  
    varChars = Split("., ,")
    For Each var In varChars
        VString = Replace(VString, CStr(var), REPLACE_WITH)
    Next var
    FReplaceChars = VString
End Function

'replace words from a table. A_ReplaceWords is a table that has ReplaceFrom and ReplaceTo fields, this code will loop through all pairs in the table and replace them
Public Function FReplaceWords(VString As String)

    Dim RsLoop As DAO.Recordset
    Dim ReplaceFrom As String, ReplaceTo As String, qry As String, RepName As String

    Set RsLoop = CurrentDb.OpenRecordset("A_ReplaceWords", dbOpenDynaset, dbSeeChanges)

    Do Until RsLoop.EOF
     
        ReplaceFrom = RsLoop("ReplaceFrom")
        ReplaceTo = RsLoop("ReplaceTo")
 
        If InStr(1, VString, ReplaceFrom, vbTextCompare) > 0 Then
            FReplaceWords = FRegExpReplaceWord(VString, ReplaceFrom, ReplaceTo)
        End If
         
        RsLoop.MoveNext
    Loop

    RsLoop.Close
    Set RsLoop = Nothing

End Function

'Finds occurrences of the string inside a string, Instance specifies which occurence you want
Public Function FCharPos(VSearchString As String, VChar As String, VInstance As Long)
    Dim x     As Integer, n As Long
   
    For x = 1 To Len(VSearchString)
        FCharPos = FCharPos + 1
        If Mid(VSearchString, x, Len(VChar)) = VChar Then n = n + 1
        If n = VInstance Then Exit Function
    Next x
   
    'The error below will only be triggered if the function was not already exited due to success
    'CharPos = CVErr(xlErrValue)
   
End Function

    ' Counts occurrences of a particular character or characters.
    ' If lngCompare argument is omitted, procedure performs binary comparison.
Public Function FStringCountOccurrences(VSearchString As String, VFind As String, Optional lngCompare As VbCompareMethod) As Long
    'Testcases:
    '?StringCountOccurrences("","") = 0
    '?StringCountOccurrences("","a") = 0
    '?StringCountOccurrences("aaa","a") = 3
    '?StringCountOccurrences("aaa","b") = 0
    '?StringCountOccurrences("aaa","aa") = 1
    Dim lngPos As Long
    Dim lngTemp As Long
    Dim lngCount As Long
    If Len(VSearchString) = 0 Then Exit Function
    If Len(VFind) = 0 Then Exit Function
    lngPos = 1
    Do
        lngPos = InStr(lngPos, VSearchString, VFind, lngCompare)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            lngPos = lngPos + Len(VFind)
        End If
    Loop Until lngPos = 0
    FStringCountOccurrences = lngCount
End Function

    ' search and replace whole word
    ' [strFind] can be plain text or a regexp pattern;  all occurences of [strFind] are replaced '
Public Function FRegExpReplaceWord(VSource As String, VFind As String, VReplace As String) As String
    'requires reference to Microsoft VBScript Regular Expressions '
    'Dim re As RegExp '
    'Set re = New RegExp '
    'late binding; no reference needed '
    Dim re    As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True                         ' <-- case insensitve
    re.Pattern = "\b" & VFind & "\b"
    FRegExpReplaceWord = re.Replace(VSource, VReplace)
    Set re = Nothing

End Function


' This function splits the sentence in InputText into words and returns a string array of the words. Each element of the array contains one word.
Public Function FSplit(VString As String, Optional VDelimiter As String) As Variant

    ' This constant contains punctuation and characters that should be filtered from the input string.
    Const CHARS = ".!?,;:""'()[]{}"
    Dim strReplacedText As String
    Dim intIndex As Integer
 
    strReplacedText = Trim(Replace(VString, vbTab, " "))
 
    For intIndex = 1 To Len(CHARS)
        strReplacedText = Trim(Replace(strReplacedText, Mid(CHARS, intIndex, 1), " "))
    Next intIndex
 
    Do While InStr(strReplacedText, " ")
        strReplacedText = Replace(strReplacedText, " ", " ")
    Loop
 
    'MsgBox "String:" & strReplacedText
    If Len(VDelimiter) = 0 Then
        FSplit = VBA.Split(strReplacedText)
    Else
        FSplit = VBA.Split(strReplacedText, VDelimiter)
    End If
End Function

Public Function getField(VString As String, FieldNo As Integer, FieldSize As Integer) As String
    Dim intBreakPoint As Integer
    Dim strW  As String
    Dim i     As Integer

    strW = VString
    For i = 1 To FieldNo
        strW = LTrim(strW)
        If Len(strW) > FieldSize Then
            intBreakPoint = InStrRev(Left(strW, FieldSize + 1), " ")
            getField = Left(strW, intBreakPoint - 1)
            strW = Right(strW, Len(strW) - intBreakPoint)
        Else
            getField = strW
            strW = ""
        End If
    Next i

End Function

' This function returns the long integer number of words in InputText.
Public Function FCountWords(VString As String) As Long
    Dim astrWords() As String

    astrWords = Split(VString)
    FCountWords = UBound(astrWords) - LBound(astrWords) + 1
  
End Function

'split string by spaces
Public Function FSplitbySpace(VString As String, intelement As Integer) As String
    On Error Resume Next
    Dim strResult() As String
    strResult = Split([VString], Chr(32))
    FSplitbySpace = strResult(intelement)
End Function

'split string by comma
Public Function FSplitByComma(VString As String, intelement As Integer) As String
    On Error Resume Next
    Dim strResult() As String
    strResult = Split([VString], ",")
    FSplitByComma = strResult(intelement)
End Function

Private Sub SplitTest()
    Dim strTest As String
    strTest = Chr(9) & "  Is this  a  " _
                   & String(2, 9) _
                   & "([long] and  'boring')" & vbTab _
                   & "  sentence? " & String(3, Asc(vbTab))
    MsgBox """" & strTest & """" & vbCr & vbLf & "Words:" & FCountWords(strTest)
    
    strTest = vbTab & " word " & vbTab
    MsgBox """" & strTest & """" & vbCr & vbLf & "Words:" & FCountWords(strTest)
    
    strTest = ""
    MsgBox """" & strTest & """" & vbCr & vbLf & "Words:" & FCountWords(strTest)
    
    strTest = " "
    MsgBox """" & strTest & """" & vbCr & vbLf & "Words:" & FCountWords(strTest)
    
    strTest = String(5, Asc(vbTab))
    MsgBox """" & strTest & """" & vbCr & vbLf & "Words:" & FCountWords(strTest)
End Sub


