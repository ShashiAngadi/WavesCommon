Attribute VB_Name = "wisChest"
 Option Explicit

Public Function GetSereverDate() As String

Dim DBPath As String

'DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "server")

If DBPath = "" Then
    'Give the local path of the MDB FILE
    DBPath = App.Path
Else
    DBPath = "\\" & DBPath
End If
On Error Resume Next
Shell DBPath & "\GetDate.exe"
Dim FIleNo As Integer
Dim DateStr As String
DateStr = String(255, 0)
FIleNo = FreeFile
Open DBPath & "\DateFile.dat" For Input As #FIleNo
Input #FIleNo, DateStr
Close #FIleNo
If Trim(DateStr) = "" Or DateStr = String(255, 0) Then
    DateStr = Format(Now, "MM/DD/YYYY")
End If
GetSereverDate = FormatDate(DateStr)
GetSereverDate = Format(Now, "DD/MM/YYYY")
End Function


'This function returns position of an occurence of a search string
'within another string being searched. And it will search from Right to left
Public Function InstrRev(strString1 As String, strString2 As String, Optional lngStartpos As Integer, Optional Compare As VbCompareMethod) As Long

InstrRev = 0

'Declaring the variables
Dim Pos As Long
Dim i As Integer
Dim StrLen As Long

On Error GoTo ExitLine
'Reversing the string
strString1 = strReverse(strString1)
strString2 = strReverse(strString2)

StrLen = Len(strString1)
If lngStartpos = 0 Then
    lngStartpos = 1
End If
If IsMissing(Compare) Then Compare = vbBinaryCompare

'find the posistion of occurence of string
Pos = InStr(lngStartpos, strString1, strString2, Compare)

If Pos Then InstrRev = StrLen - (Len(strString2) + Pos - 1) + 1

ExitLine:
    Exit Function
    
End Function

'This function will reverses the string being passed to it.
'and returns revesrse of the string.
Private Function strReverse(string1 As String) As String
Dim strRev As String
Dim i As Integer
For i = Len(string1) To 1 Step -1
    strRev = strRev + Mid(string1, i, 1)
Next i
strReverse = strRev
End Function


'*****************************************************************************************************************
'                                   Update Last Accessed Elements
'*****************************************************************************************************************
'This function will be useful if you want to update the last accessed elements
'Eg : Last Accessed Files
'  Suppose you want the last of 4 last accessed files and you have only 2 files.
'  pass the other 2 elements as "" (NULL)
'
'
'   Girish  Desai  May 1st, 1998.
'
Function UpdateLastAccessedElements(Str As String, StrArr() As String, Optional IgnoreCase As Boolean)

Dim CaseVal As Integer
Dim Pos As Integer
Dim Flag As Boolean
Dim Count As Integer
Dim IgnCase As Boolean


    IgnCase = False
    If Not IsMissing(IgnoreCase) Then
        IgnCase = IgnoreCase
    End If


    If IgnCase Then
        CaseVal = vbBinaryCompare
    Else
        CaseVal = vbTextCompare
    End If

'First check out the position
    For Pos = 0 To UBound(StrArr)
        If StrComp(Str, StrArr(Pos), CaseVal) = 0 Then
            Flag = True
            Exit For
        End If
    Next Pos
    
    If Not Flag Then Pos = Pos - 1
    
    For Count = Pos To 1 Step -1
        StrArr(Count) = StrArr(Count - 1)
    Next Count

    StrArr(0) = Str
DoEvents
End Function

Public Function LoadGridSettings(grd As Object, GrdName As String, FileName As String) As Boolean
Dim strIniVal As String
Dim i As Integer
'Prelim Checks
    If FileName = "" Then
        Exit Function
    End If

'strIniVal = ReadFromIniFile(GrdName, "Cols", FileName)
'If Trim$(strIniVal) <> "" Then grd.Cols = Val(strIniVal)

For i = 0 To grd.Cols - 1
    'strIniVal = ReadFromIniFile(GrdName, "ColWidth" & i, FileName)
    If Trim$(strIniVal) <> "" Then grd.ColWidth(i) = Val(strIniVal)
Next i
LoadGridSettings = True
End Function

Function RPad(Str As String, PAdWith As String, LenToPad As Integer) As String
RPad = Str
If LenToPad < Len(Str) Then Exit Function

If Len(PAdWith) > 1 Then Exit Function

RPad = Str & String(LenToPad - Len(Str), PAdWith)


End Function



' Find and remove the next token from this string.
'
' Tokens are stored in the format:
'    name1(value1)name2(value2)...
' Invisible characters (tabs, vbCrLf, spaces, etc.)
'    are allowed before names.
Sub GetToken(Txt As String, token_name As String, _
    token_value As String)
Dim open_pos As Integer
Dim close_pos As Integer
Dim txtlen As Integer
Dim num_open As Integer
Dim i As Integer
Dim ch As String

' Initialize token_name and value.
token_name = ""
token_value = ""

    ' Remove initial invisible characters.
    TrimInvisible Txt

    ' If the string is empty, do nothing.
    If Txt = "" Then Exit Sub

    ' Find the opening parenthesis.
    open_pos = InStr(Txt, "(")
    txtlen = Len(Txt)
    If open_pos = 0 Then open_pos = txtlen

    ' Find the corresponding closing parenthesis.
    num_open = 1
    For i = open_pos + 1 To txtlen
        ch = Mid$(Txt, i, 1)
        If ch = "(" Then
            num_open = num_open + 1
        ElseIf ch = ")" Then
            num_open = num_open - 1
            If num_open = 0 Then Exit For
        End If
    Next i
    If open_pos = 0 Or i > txtlen Then
        ' There is something wrong.
        Err.Raise vbObjectError + 1, _
            "InventoryItem.GetToken", _
            "Error parsing serialization """ & Txt & """"
    End If
    close_pos = i

    ' Get token name and value.
    token_name = Left$(Txt, open_pos - 1)
    token_value = Mid$(Txt, open_pos + 1, _
        close_pos - open_pos - 1)
    'TrimInvisible token_name
    'TrimInvisible token_value

    ' Remove leading spaces.
    token_name = Trim$(token_name)
    token_value = Trim$(token_value)
    
    ' Remove the token name and value
    ' from the serialization string.
    Txt = Right$(Txt, txtlen - close_pos)
End Sub

' Remove leading invisible characters from
' the string (tab, space, CR, etc.)
Public Sub TrimInvisible(Txt As String)
Dim txtlen As Integer
Dim i As Integer
Dim ch As String

    txtlen = Len(Txt)
    For i = 1 To txtlen
        ' See if this character is visible.
        ch = Mid$(Txt, i, 1)
        If ch > " " And ch <= "~" Then Exit For
    Next i
    If i > 1 Then _
        Txt = Right$(Txt, txtlen - i + 1)
End Sub

' Retrieves the value for a specified token
' in a given source string.
' The source should be of type :
'       name1=value1;name2=value2;...;name(n)=value(n)
'   similar to DSN strings maintained by ODBC manager.
Public Function ExtractToken(src As String, TokenName As String) As String

' If the src is empty, exit.
If Len(src) = 0 Or _
    Len(TokenName) = 0 Then Exit Function

' Search for the token name.
Dim token_pos As Integer
Dim strSearch As String
strSearch = TokenName & "="
'token_pos = InStr(src, strSearch)
'If token_pos = 0 Then
'    'Try ignoring the white space
'    strSearch = token_name & " ="
'    token_pos = InStr(src, strSearch)
'    If token_pos = 0 Then Exit Function
'End If

' Search for the token_name in the src string.
 token_pos = InStr(1, src, strSearch, vbTextCompare)
Do
    ' The character before the token_name
    ' should be ";" or, it should be the first word.
    ' Else, search for the next occurance of the token.
    If token_pos = 0 Then
        If token_pos = 0 Then
            'Try ignoring the white space
            strSearch = TokenName & " ="
            token_pos = InStr(src, strSearch)
            If token_pos = 0 Then Exit Function
        End If
    ElseIf token_pos = 1 Then
        Exit Do
    ElseIf Mid$(src, token_pos - 1, 1) = ";" Then
        Exit Do
    Else
        'Get next occurance.
        token_pos = InStr(token_pos + 1, src, TokenName, vbTextCompare)
    End If
Loop

token_pos = token_pos + Len(strSearch)

' Search for the delimiter ";", after the token_pos.
Dim Delim_pos As Integer
Delim_pos = InStr(token_pos, src, ";")
If Delim_pos = 0 Then Delim_pos = Len(src) + 1

' Return the token_value.
ExtractToken = Mid$(src, token_pos, Delim_pos - token_pos)
End Function


Function putToken(src As String, token_name As String, token_value As String) As String
On Error GoTo Err_Line

Dim token_pos As Integer
Dim token_end As Integer
Dim assign_pos As Integer
Dim strTokenVal As String
Dim strBefore As String, strAfter As String

' Search for the token_name in the src string.
token_pos = InStr(1, src, token_name, vbTextCompare)
Do
    ' The character before the token_name
    ' should be ";" or, it should be the first word.
    ' Else, search for the next occurance of the token.
    If token_pos = 0 Then
        token_pos = Len(src) + 1
        Exit Do
    ElseIf token_pos = 1 Then
        Exit Do
    ElseIf Mid$(src, token_pos - 1, 1) = ";" Then
        Exit Do
    Else
        'Get next occurance.
        token_pos = InStr(token_pos + 1, src, token_name, vbTextCompare)
    End If
Loop
strBefore = Left$(src, token_pos - 1)

' Check for assignment symbol (=).
assign_pos = InStr(token_pos + 1, src, "=")
If assign_pos = 0 Then assign_pos = token_pos

' Check for terminating symbol (;).
token_end = InStr(token_pos, src, ";")
If token_end = 0 Then
    token_end = Len(src)
    'strAfter = ""
End If
strAfter = Mid$(src, token_end + 1)

' Ensure a ";" after strBefore
If strBefore <> "" Then
    If Right$(strBefore, 1) <> ";" Then
        strBefore = strBefore & ";"
    End If
End If

' Ensure a ";" before 'strAfter'
If strAfter <> "" Then
    If Left$(strAfter, 1) <> ";" Then
        strAfter = ";" & strAfter
    End If
End If

putToken = strBefore & token_name _
            & "=" & token_value & strAfter


Err_Line:
    If Err Then
        MsgBox "Put_token: " & Err.Description, vbCritical
    End If
End Function


Public Function FormatCurrency(ByVal Curr As Currency) As String
    FormatCurrency = Format(Curr, "##########0.00")
End Function


'***************************************************************************************************************
'                                               DATE VALIDATE FUNCTION
''***************************************************************************************************************
'       Function to Validate a string for date. Supports only the following date formats :
'           1. dd/mm/yyyy       - Indian Format
'           2. mm/dd/yyyy       - American Format
'       A String whose Date Validation has to be checked, The Delimeter should be passed to it.
'
'       Specify the IsIndian Optional parameter as True if you want the validation for format no.1
'
'       Date :  19 May 1998.
'       Last Modified By : Ravindranath M.
'       Dependencies    : GetstringArray()
'                         isLeap()
'
'       Date : 11 Jan 2000
'        Last Modified By : Girish Desai
'        Changes Made :     Fixed problem of 2000  ie (when user specified 00)
'                           Checking Ubound(DateArray) < 2
'                           if len(year) = 2, if < 30 then 19yr elseif > 30 then 20yr !!!
'

Function DateValidate(DateText As String, Delimiter As String, Optional IsIndian As Boolean) As Boolean
DateValidate = False
On Error Resume Next
'Check For The Decimal point in the string.
'If there is any decimal point the cint will

If InStr(1, DateText, ".", vbTextCompare) Then Exit Function

'Breakup the given string into array elements based on the delimiter.
Dim DateArray() As String
GetStringArray DateText, DateArray(), Delimiter

'Quit if ubound is < 3   - GIRISH 11/1/2000
If UBound(DateArray) < 2 Then Exit Function

' Get the date, month and year parts.
Dim DayPart As Integer
Dim MonthPart As Integer
Dim YearPart As Integer
On Error GoTo ErrLine
If IsIndian Then
    DayPart = CInt(DateArray(0))
    MonthPart = CInt(DateArray(1))
Else
    DayPart = CInt(DateArray(1))
    MonthPart = CInt(DateArray(0))
End If

YearPart = CInt(DateArray(2))
On Error GoTo 0
' The day, month and year should not be 0.
If DayPart = 0 Then
    'MsgBox "Inavlid day value.", vbInformation
    Exit Function
End If
If MonthPart = 0 Then
    'MsgBox "Inavlid day value.", vbInformation
    Exit Function
End If
'Changed condition from = to < - Girish 11/1/2000
If YearPart < 0 Then
    'MsgBox "Inavlid year value.", vbInformation
    Exit Function
End If
'The yearpart should not exceed 4 digits.
If Len(CStr(YearPart)) > 4 Then
    'MsgBox "Year is too long.", vbInformation
    Exit Function
End If

' The month part should not exceed 12.
If MonthPart > 12 Then
    'MsgBox "Invalid month.", vbInformation
    Exit Function
End If

' If the year part is only 2 digits long,
' then prefix the century digits.
If Len(CStr(YearPart)) = 2 Then
    'YearPart = Left$(CStr(Year(gStrDate)), 2) & YearPart
    '5 lines added by Girish    11/1/2000
    If Val(YearPart) <= 30 Then
        YearPart = "20" & YearPart
    Else
        YearPart = "19" & YearPart
    End If
End If

' Check if it is a leap year.
Dim bLeapYear As Boolean
bLeapYear = isLeap(YearPart)

' Validations.
Select Case MonthPart
    Case 2  ' Check for February month.
        If bLeapYear Then
            If DayPart > 29 Then
                Exit Function
            End If
        Else
            If DayPart > 28 Then
                
                Exit Function
            End If
        End If
    
    Case 4, 6, 9, 11 ' Months having 30 days...
        If DayPart > 30 Then
            Exit Function
        End If
    Case Else
        If DayPart > 31 Then
            Exit Function
        End If
End Select

DateValidate = True
ErrLine:
    

End Function


Private Function isLeap(Year As Integer) As Boolean

isLeap = ((Year Mod 400) = 0) Or _
    ((Year Mod 4 = 0) And (Year Mod 100 <> 0))

End Function

Function CurrencyValidate(CurStr As String, AcceptZeroes As Boolean) As Boolean
On Error GoTo ErrLine
    Dim MyCur As Currency
    If CurStr = "" Then
        GoTo ErrLine
    End If
    MyCur = CCur(CurStr)
    If Not AcceptZeroes Then
        If MyCur = 0 Then
            GoTo ErrLine
        End If
    End If
        
    
CurrencyValidate = True
Exit Function
ErrLine:

End Function

Public Function FormatAccountNumber(AccNo As Long) As String
    FormatAccountNumber = Format(AccNo, "00000")
End Function


Public Sub ActivateTextBox(Txt As TextBox)
On Error Resume Next
With Txt
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


' Checks for occurance of single quotes in the given string
' and replaces them with additional quotes, so that the
' string can be used in SQL statements for insertion/updation.
'
' INPUT:
'   fldStr - The source string required to be formatted.
'   Enclose (optional) - Indicates that the formatted string
'           be wrapped in quotes. Ex: "'" & string & "'"
'
Public Function AddQuotes(FldStr As String, Optional Enclose As Boolean) As String
Dim QuotePos As Integer
Dim TmpStr As String
Dim TargetStr As String
    
    TmpStr = FldStr
    QuotePos = InStr(TmpStr, "'")
    If QuotePos > 0 Then
            Do While QuotePos > 0
                'Add 2 quotes for one.
                TargetStr = TargetStr & Mid$(TmpStr, 1, QuotePos - 1) & "''"
                TmpStr = Mid$(TmpStr, QuotePos + 1)
                QuotePos = InStr(TmpStr, "'")
            Loop
            TargetStr = TargetStr & TmpStr
    Else
            TargetStr = FldStr
    End If
    AddQuotes = TargetStr
    
    ' If the optional parameter "Enclose" is specified,
    ' enclose the resulting string inside single quotes.
    If Enclose Then AddQuotes = "'" & AddQuotes & "'"

End Function

' Returns the path of a specified file.
Public Function FilePath(strFile As String) As String
On Error GoTo end_line

' Start from the end of the file string,
Dim i As Integer, ch As String
For i = Len(strFile) To 1 Step -1
    ' Check for "\".
    ch = Mid$(strFile, i, 1)
    If ch = "\" Then
        FilePath = Left$(strFile, i - 1)
        Exit For
    End If
Next

end_line:
    Exit Function

End Function

Public Function AppendBackSlash(ByVal strPath As String) As String
If Right$(strPath, 1) <> "\" Then
    strPath = strPath & "\"
End If
AppendBackSlash = strPath
End Function

'This routine creates the directory hierarchy
'specified in the fields information.
'
Function MakeDirectories(DirPath As String) As Boolean
Dim lcount As Integer
Dim DirName As String, OldDir As String
Dim oldDrive As String
Dim PathArray() As String
Dim lRetVal As Integer

MakeDirectories = False 'Initialize the return value.
Screen.MousePointer = vbHourglass
    On Error GoTo ErrorLine

    'Check if the drive is mentioned in the directory path.
    If Mid$(DirPath, 2, 1) <> ":" Then
        If Left$(DirPath, 1) = "\" Then
            'Prefix the drive letter, if the path starts with "\"
            DirPath = Left(CurDir, 2) & DirPath
        Else
            'Prefix the current directory.
            DirPath = CurDir & "\" & DirPath
        End If
    End If

    'Breakup the path into an array
    lRetVal = GetStringArray(DirPath, PathArray(), "\")

    'Save the current drive, and change to the drive of dirpath.
    oldDrive = Left(CurDir, 1)
    OldDir = CurDir
    
    ChDrive Left(DirPath, 1)

    DirName = ""
    For lcount = 0 To UBound(PathArray)
        If PathArray(lcount) <> "" Then
            DirName = DirName & Trim$(PathArray(lcount))
        End If
        If Dir$(DirName, vbDirectory) = "" Then
            MkDir DirName   'create directory
        End If
        DirName = DirName & "\"
        'ChDir DirName   'make it the current directory.
        '
    Next lcount
    MakeDirectories = True

ErrorLine:
    On Error Resume Next
    Screen.MousePointer = vbDefault
    If Left(oldDrive, 1) <> "\" Then
        ChDrive oldDrive
        ChDir OldDir
        If Err > 0 Then
            MsgBox "Error in creating the path '" _
                & DirPath & "'" & vbCrLf & Err.Description, vbCritical
            'MsgBox LoadResString(gLangOffSet + 809) & " " _
                & DirPath & "'" & vbCrLf & Err.Description, vbCritical
        End If
    End If
'Resume
End Function

'*********************************************************************************************************
'                                   GET STRING ARRAY
'*********************************************************************************************************
'
'   To get an array from a string seperated by a delimiter
'   Date : 24th Nov 1997
'   Dependencies : <None>
Function GetStringArray(GivenString As String, strArray() As String, Delim As String)

Dim Pos As Integer
Dim PrevPos As Integer
Dim TmpStr As String


ReDim strArray(0)
TmpStr = GivenString
'check whether the delimeter is there at the end
If Right(TmpStr, 1) = Delim Then
 TmpStr = Left(TmpStr, Len(TmpStr) - 1)
End If

Pos = 0
PrevPos = 1
Do
    Pos = InStr(1, TmpStr, Delim)
    If Pos = 0 Then
        Exit Do
    End If
    
    
    strArray(UBound(strArray)) = Left(TmpStr, Pos - 1)
    'TmpStr = Right(TmpStr, Len(TmpStr) - Pos)
    TmpStr = Mid(TmpStr, Pos + Len(Delim)) 'changed on 27/2/99
    ReDim Preserve strArray(UBound(strArray) + 1)
Loop
    strArray(UBound(strArray)) = TmpStr

End Function


'***********************************************************************
'                           DOES PATH EXIST
'
''***********************************************************************
'Function to check if the path.
'Returns 0 if path does not exist
'Returns 1 if it is a file
'Returns -1 if it is read only file
'Returns 2 if it is a directory
'Returns -2 if it is a read only directory


Function DoesPathExist(ByVal Path As String) As Integer

On Error GoTo ErrLine
Dim Retval As Integer
 
  Retval = GetAttr(Path)
    If Retval >= 32 Then
        Retval = Retval - 32
    End If
    
    If Retval >= 17 Then
        DoesPathExist = -2 'Read Only Directory
        Exit Function
    End If
        
    If Retval >= 16 Then
        DoesPathExist = 2 'Normal Only Directory
        Exit Function
    End If
    
    If Retval = 1 Then
        DoesPathExist = -1  'Read Only File
    Else
        DoesPathExist = 1   'Normal File
    End If
    
Exit Function
ErrLine:
    DoesPathExist = 0
End Function

' Formats the given date string according to DD/MM/YYYY.
' Currently, it assumes that the given date is in MM/DD/YYYY.
Public Function FormatDate(strdate As String) As String
On Error GoTo FormatDateError
' Swap the DD and MM portions of the given date string
Const Delimiter = "/"
Dim YearPart As String
Dim strArray() As String
Dim Particulars As String * 49

'First Check For the Space in the given string
' Because the Date & Time part will be seperated bt a space
strdate = Trim$(strdate)
Dim SpacePos As Integer
SpacePos = InStr(1, strdate, " ")
If SpacePos Then
    strdate = Left(strdate, SpacePos - 1)
End If

' Breakup the date string into array elements.
GetStringArray strdate, strArray(), Delimiter

' Check if the year part contains 2 digits.
ReDim Preserve strArray(2)
YearPart = Left$(strArray(2), 4)
If Len(Trim$(strArray(2))) = 2 Then
    ' Check, if it is greater than 30, in which case,
    ' Add "20", else, add "19".
    If Val(strArray(2)) < 30 Then
        YearPart = "20" & Right$(Trim(YearPart), 2)
    Else
        YearPart = "19" & Right$(Trim(YearPart), 2)
    End If
End If

' Change the month and day portions and concatenate.
FormatDate = strArray(1) & Delimiter & strArray(0) & Delimiter & YearPart

FormatDateError:
'Trap The Settings OF Date Adjusted IN Control Panel
'VbDayOfWeek.vbUseSystemDayOfWeek
 
End Function
Public Function StripExtn(FileName As String) As String
Dim ExtnPos As Integer

' Check for extension
ExtnPos = InStr(FileName, ".")
If ExtnPos = 0 Then ExtnPos = Len(FileName) + 1

' Return the stripped file name.
StripExtn = Mid$(FileName, 1, ExtnPos - 1)

End Function


' -- FormatField:  Formats a given field data
'                  according to its type and returns.
'   Input:  Field object
'   Output: Variant, depends on the data type of the field.
'
Public Function FormatField(fld As Field) As Variant
On Error Resume Next
    If IsNull(fld.value) Then
        ' If the value in the field is NULL,
        ' return it as a Null String rather than NULL.
        ' This will avoid potential run-time errors.
          FormatField = vbNullString
          ' Check if the field is date type.
          If fld.Type = adSingle Or fld.Type = adUnsignedTinyInt Or fld.Type = adInteger Or fld.Type = adDouble Or fld.Type = adNumeric Or fld.Type = 2 Or fld.Type = adCurrency Then
                FormatField = "0"
          End If
    Else
        ' Check if the field is date type.
        If fld.Type = adDate Then
            FormatField = FormatDate(CStr(fld.value))
        ElseIf fld.Type = adCurrency Then
            FormatField = FormatCurrency(fld.value)
        Else
            FormatField = fld.value
        End If
  End If

End Function



Public Function FormatDateField(fld As Field) As String
On Error Resume Next
If fld.Type <> adDate Then Exit Function
If IsNull(fld.value) Then
    ' If the value in the field is NULL,
    ' return it as a Null String rather than NULL.
    ' This will avoid potential run-time errors.
    FormatDateField = "NULL"
Else
    FormatDateField = "#" + CStr(fld.value) + "#"
End If

End Function




Public Function WisDateDiff(IndianDate1 As String, IndianDate2 As String) As Variant
    On Error Resume Next
    WisDateDiff = DateDiff("d", FormatDate(IndianDate1), FormatDate(IndianDate2))
    
    If Err Then MsgBox Err.Number & vbCrLf & Err.Description
    
End Function


