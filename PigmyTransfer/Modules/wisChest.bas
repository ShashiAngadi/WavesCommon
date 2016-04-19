Attribute VB_Name = "wisChest"
Option Explicit
Private m_DateFormat As String
Private m_DayFormat As String
Private m_YearFormat As String
Private m_DateSep As String
Private m_swap As Boolean


''Time
Private Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vtime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FileTime) As Long
Private Declare Function VariantChangeTypeEx Lib "oleaut32.dll" (ByVal pvArgDest As Long, ByVal pvArgSrc As Long, ByVal LCID As Long, ByVal wFlags As Integer, ByVal VarType As Integer) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpFileTime As FileTime, lpLocalFileTime As FileTime) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long

Const ANYSIZE_ARRAY = 1
Public Const TOKEN_ADJUST_PRIVILEGES = 32
Public Const TOKEN_QUERY = 8

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Type LUID_AND_ATTRIBUTES
        pLuid As LARGE_INTEGER
        Attributes As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4

Public Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


Public Property Let DateFormat(NewValue As String)
Dim ChkDate As Date
Dim Chkstr As String

ChkDate = "10/25/2000"
Chkstr = "10/25/2000"
m_swap = IIf(Chkstr = ChkDate, False, True)


    Dim Pos As Integer
    m_DateFormat = NewValue
    If UCase(Mid(NewValue, 2, 1)) = "D" Then
        Pos = 7
        m_DateSep = Mid(NewValue, 3, 1)
        m_DayFormat = "00"
    Else
        Pos = 5
        m_DateSep = Mid(NewValue, 2, 1)
        m_DayFormat = "0"
    End If
    
    If Len(Mid(NewValue, Pos)) = 4 Then
        m_YearFormat = "0000"
    Else
        m_YearFormat = "00"
    End If

End Property


Public Property Get DateFormat() As String
    DateFormat = m_DateFormat
End Property


'This Function Returns the Date with Last day of the given date
'If Input is 16/10/2000 This Function Returns the 31/10/2000
'If Input is 16/09/2000 This Function Returns the 30/09/2000
'This reurns ir Indian Date Format
'Assume the Input is also in the same format
Public Function GetAppLastDate(ByVal IndainDate As String) As String
Dim Pos As Integer
Dim SecPos As Integer
Dim intMonth As Integer


Pos = InStr(1, IndainDate, m_DateSep, vbTextCompare)
If Pos = 0 Then Pos = InStr(1, IndainDate, "/", vbTextCompare)
SecPos = InStr(Pos + 1, IndainDate, m_DateSep, vbTextCompare)
If SecPos = 0 Then SecPos = InStr(Pos + 1, IndainDate, "/", vbTextCompare)

intMonth = CInt(Mid(IndainDate, Pos + 1, SecPos - Pos - 1))
Select Case intMonth
    Case 1, 3, 5, 7, 8, 10, 12
        GetAppLastDate = "31" & m_DateFormat & _
                intMonth & m_DateSep & _
                Mid(IndainDate, SecPos + 1)
    
    Case 2
        If CInt(Mid(IndainDate, SecPos + 1)) Mod 4 Then
            GetAppLastDate = "28" & m_DateFormat & _
                    intMonth & m_DateSep & _
                    Mid(IndainDate, SecPos + 1)
        Else
            GetAppLastDate = "29" & m_DateFormat & _
                    intMonth & m_DateSep & _
                    Mid(IndainDate, SecPos + 1)
        End If
    Case Else
        GetAppLastDate = "30" & m_DateFormat & _
                intMonth & m_DateSep & _
                Mid(IndainDate, SecPos + 1)
    'case end
    
End Select

End Function


'This Function Returns the Date with Last day of the given date
'If Input is 16 october 2000 This Function Returns the 31 october 2000
'If Input is 16 september 2000 This Function Returns the 30 september 2000
'This reurns sstem Date Format
'Assume the Input is also in the same format
Public Function GetSysLastDate(ByVal GivenDate As Date) As Date

On Error GoTo ErrLine

Dim bDay As Byte
Dim bMonth As Byte
Dim IntYear As Integer


bMonth = Month(GivenDate)
IntYear = Year(GivenDate)


Select Case bMonth
    Case 1, 3, 5, 7, 8, 10, 12
        bDay = "31"
    
    Case 2
        bDay = IIf(IntYear Mod 4, 28, 29)
    
    Case Else
        bDay = "30"
    'case end
    
End Select

GetSysLastDate = GetSysFormatDate(bDay & m_DateSep & bMonth & m_DateSep & IntYear)

Exit Function

ErrLine:
    Err.Clear

End Function

'This Function Returns the Date with First of the give date
'If Input is 16/10/2000 This Function Returns the 1/10/2003
'This reurns ir Indian Date Format
'Assume the Input is also in the same format
Public Function GetAppFirstDate(ByVal IndainDate As String) As String
Dim Pos As Integer
Dim SecPos As Integer

Pos = InStr(1, IndainDate, m_DateSep, vbTextCompare)
If Pos = 0 Then Pos = InStr(1, IndainDate, "/", vbTextCompare)
SecPos = InStr(Pos + 1, IndainDate, m_DateSep, vbTextCompare)
If SecPos = 0 Then SecPos = InStr(Pos + 1, IndainDate, "/", vbTextCompare)

GetAppFirstDate = "01" & m_DateFormat & _
                Mid(IndainDate, Pos + 1, SecPos - Pos - 1) & _
                m_DateSep & Mid(IndainDate, SecPos + 1)

End Function


'This Function Returns the Date with First of the given date
'If Input is 16 october 2000 This Function Returns the 1 october 2000
'This returns ir Indian Date Format
'Assume the Input is also in the same format
Public Function GetSysFirstDate(ByVal GivenDate As Date) As Date

On Error GoTo ErrLine

'Dim bDay As Byte
Dim bMonth As Byte
Dim IntYear As Integer

'bDay = Day(GivenDate)
bMonth = Month(GivenDate)
IntYear = Year(GivenDate)

GetSysFirstDate = GetSysFormatDate(1 & m_DateSep & bMonth & m_DateSep & IntYear)

ErrLine:
    Err.Clear

End Function


'This Function Converts the given system date to Appliaction date
'[Inputs} Expression as date
'[Returns] The date format Appliaction is using
'           'Presently it is dd/mm/yyyy
Public Function GetIndianDate(Expression As Date) As String

'GetAppFormatDate = FormatDate1(CStr(Expression))
'GetIndianDate = FormatDate1(CStr(Expression))
'Exit Function


On Error GoTo ErrLine

Dim LDate As Date
Dim SDate As SYSTEMTIME
Dim strYear As String

LDate = Expression
Call VariantTimeToSystemTime(LDate, SDate)
'GetAppFormatDate = SDate.wDay & "/" & SDate.wMonth & "/" & SDate.wYear

strYear = Format(SDate.wYear, "0000")
If Len(m_YearFormat) = 2 Then strYear = Mid(strYear, 3)
GetIndianDate = Format(SDate.wDay, m_DayFormat) & m_DateSep & _
                   Format(SDate.wMonth, m_DayFormat) & m_DateSep & _
                   strYear

Exit Function

ErrLine:

'    GetAppFormatDate = ""
End Function

'This Function Converts the given Indian date format to Application format date
'[Inputs} Expression as date
'[Returns] The date format Application is using
'           'Presently it is dd/mm/yyyy
Private Function GetAppFormatDate(IndianDate As String) As String

'GetAppFormatDate111 = FormatDate1(CStr(Expression))

On Error GoTo FormatDateError
'Swap the DD and MM portions of the given date string
Const Delimiter = "/"

Dim TempDelim As String
Dim YearPart As String
Dim strArray() As String

'First Check For the Space in the given string
'Because the Date & Time part will be saperated by a space
IndianDate = Trim$(IndianDate)
Dim SpacePos As Integer

'check for the deimeter
TempDelim = IIf(InStr(1, IndianDate, m_DateSep), m_DateSep, Delimiter)

SpacePos = InStr(1, IndianDate, " ")
If SpacePos Then IndianDate = Left(IndianDate, SpacePos - 1)

'Breakup the date string into array elements.
'GetStringArray strDate, strArray(), TempDelim
strArray = Split(IndianDate, TempDelim)

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

'Change the month and day portions and concatenate.
TempDelim = IIf(InStr(1, IndianDate, m_DateSep), Delimiter, m_DateSep)

GetAppFormatDate = strArray(1) & TempDelim & strArray(0) & TempDelim & YearPart
'If gIsIndianDate Then FormatDate = strArray(0) & TempDelim & strArray(1) & TempDelim & YearPart

FormatDateError:

Exit Function


On Error GoTo ErrLine

Dim LDate As Date
Dim SDate As SYSTEMTIME
Dim strYear As String

LDate = IndianDate
Call VariantTimeToSystemTime(LDate, SDate)
'GetAppFormatDate = SDate.wDay & "/" & SDate.wMonth & "/" & SDate.wYear

strYear = Format(SDate.wYear, "0000")
If Len(m_YearFormat) = 2 Then strYear = Mid(strYear, 3)
'GetAppFormatDate = Format(SDate.wDay, m_DayFormat) & m_DateSep & _
                   Format(SDate.wMonth, m_DayFormat) & m_DateSep & _
                   strYear

Exit Function

ErrLine:

'    GetAppFormatDate = ""
End Function


'This Function Coverts the given DateString to System date
'[Inputs} Expression as String
'[Returns] The date format Aplliaction is using
''Presently it is dd/mm/yyyy
Public Function GetSysFormatDate(IndainDate As String) As Date

'GetSysFormatDate = FormatDate1(Expression)
If InStr(1, DateFormat, "/") Then
    IndainDate = Replace(IndainDate, "-", "/")
Else
    IndainDate = Replace(IndainDate, "/", "-")
End If

Dim strArr() As String
strArr = Split(CStr(IndainDate), m_DateSep)
GetSysFormatDate = strArr(0) & "-" & Mid("JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC", Val(strArr(1)) * 3 - 2, 3) & "-" & strArr(2)

Exit Function

Dim Var As Variant

    'Var = FormatDate1(Expression)
    Var = IIf(m_swap, FormatDate1(IndainDate), IndainDate)

    If 0 = VariantChangeTypeEx(VarPtr(Var), VarPtr(Var), &H409, 0, vbDate) Then
        GetSysFormatDate = Var
    Else
        'Raise same error as CDate
        Err.Raise 13
        'GetSysFormatDate = "1/1/2000"
    End If

End Function

Function GetFileName(ByVal StrPath As String) As String
    Dim strFileName As String
    Dim iSep As Integer
    
    strFileName = StrPath
    Do
        iSep = InStr(strFileName, "\")
        If iSep = 0 Then iSep = InStr(strFileName, ":")
        If iSep = 0 Then
            GetFileName = strFileName
            Exit Function
        Else
            strFileName = Right(strFileName, Len(strFileName) - iSep)
        End If
    Loop
End Function


'-----------------------------------------------------------
' FUNCTION: GetFileSize
'
' Determine the size (in bytes) of the specified file
'
' IN: [strFileName] - name of file to get size of
'
' Returns: size of file in bytes, or -1 if an error occurs
'-----------------------------------------------------------
'
Function GetFileSize(strFileName As String) As Long
    On Error Resume Next

    GetFileSize = FileLen(strFileName)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function

Private Function GetFileTime(ByVal adate As Date) As FileTime
    Dim lTemp As SYSTEMTIME
    Dim lTime As FileTime
    
    VariantTimeToSystemTime adate, lTemp
    SystemTimeToFileTime lTemp, lTime
    LocalFileTimeToFileTime lTime, GetFileTime
End Function

'Get the path portion of a filename
Function GetPathName(ByVal strFileName As String) As String
    Dim sPath As String
    Dim sFile As String
    
    SeparatePathAndFileName strFileName, sPath, sFile
    
    GetPathName = sPath
End Function

Public Sub LoadParticularsFromFile(Combo As ComboBox, Filename As String)
On Error GoTo ErrLine
'Fill up particulars with default values from SBAcc.INI
Dim Particulars As String
Dim I As Integer
Combo.Clear
Do
    Particulars = ReadFromIniFile("Particulars", "Key" & I, Filename)
    If Trim$(Particulars) <> "" Then Combo.AddItem Particulars
    I = I + 1
Loop Until Trim$(Particulars) = ""
If Combo.ListCount = 0 Then Combo.AddItem " "
ErrLine:
    Err.Clear
End Sub

Public Sub PrintNoRecords(grd As MSFlexGrid)

    With grd
        .Clear
        .rows = 2
        .cols = 2
        .ColWidth(0) = .Width * 0.2
        .ColWidth(1) = .Width * 0.75
        .TextMatrix(1, 1) = "No Records Found"
    End With

End Sub

'Given a fully qualified filename, returns the path portion and the file
'   portion.
Public Sub SeparatePathAndFileName(FullPath As String, ByRef Path As String, _
    ByRef Filename As String)

    Dim nSepPos As Long
    Dim sSEP As String

    nSepPos = Len(FullPath)
    sSEP = Mid$(FullPath, nSepPos, 1)
    Do Until IsSeparator(sSEP)
        nSepPos = nSepPos - 1
        If nSepPos = 0 Then Exit Do
        sSEP = Mid$(FullPath, nSepPos, 1)
    Loop

    Select Case nSepPos
        Case 0
            'Separator was not found.
            Path = CurDir$
            Filename = FullPath
        Case Else
            Path = Left$(FullPath, nSepPos - 1)
            Filename = Mid$(FullPath, nSepPos + 1)
    End Select
End Sub


'Determines if a character is a path separator (\ or /).
Public Function IsSeparator(Character As String) As Boolean
    Select Case Character
        Case "\"
            IsSeparator = True
        Case "/"
            IsSeparator = True
    End Select
End Function
'-----------------------------------------------------------
' SUB: ParseDateTime
'
' Same as CDate with a string argument, except that it
' ignores the current localization settings.  This is
' important because SETUP.LST always uses the same
' format for dates.
'
' IN: [strDate] - string representing the date in
'                 the format mm/dd/yy or mm/dd/yyyy
' OUT: The date which strDate represents
'-----------------------------------------------------------
'
Function ParseDateTime(ByVal strDateTime As String) As Date
Dim Var As Variant
    Var = strDateTime
    If 0 = VariantChangeTypeEx(VarPtr(Var), VarPtr(Var), &H409, 0, vbDate) Then
        ParseDateTime = Var
    Else
        'Raise same error as CDate
        Err.Raise 13
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: CheckDrive
'
' Check to see if the specified drive is ready to be read
' from.  In the case of a drive that holds removable media,
' this would mean that formatted media was in the drive and
' that the drive door was closed.
'
' IN: [strDrive] - drive to check
'     [strCaption] - caption if the drive isn't ready
'
' Returns: True if the drive is ready, False otherwise
'-----------------------------------------------------------
'
Function CheckDrive(ByVal strDrive As String, ByVal strCaption As String) As Integer
    Dim strDir As String
    Dim strMsg As String
    Dim fIsUNC As Boolean

    On Error Resume Next

'    SetMousePtr vbHourglass

    Do
        Err = 0
        fIsUNC = False
        '
        'Attempt to read the current directory of the specified drive.  If
        'an error occurs, we assume that the drive is not ready
        '
        If IsUNCName(strDrive) Then
            fIsUNC = True
            strDir = Dir$(GetUNCShareName(strDrive))
        Else
            strDir = Dir$(Left$(strDrive, 2))
        End If

        If Err > 0 Then
            If fIsUNC Then
                'strMsg = Error$ & vbLf & vbLf & ResolveResString(resCANTREADUNC, "|1", strDrive) & vbLf & vbLf & ResolveResString(resCHECKUNC)
            Else
                'strMsg = Error$ & vbLf & vbLf & ResolveResString(resDRVREAD) & strDrive & vbLf & vbLf & ResolveResString(resDRVCHK)
            End If
'            If MsgError(strMsg, vbExclamation Or vbRetryCancel, strCaption) = vbCancel Then
'                CheckDrive = False
'                Err = 0
'            End If
        Else
            CheckDrive = True
        End If
        Dim gfNoUserInput  As Boolean
        If Err And gfNoUserInput = True Then
            'ExitSetup frmSetup1, gintRET_FATAL
        End If
    Loop While Err

    'SetMousePtr gintMOUSE_DEFAULT
End Function

'-----------------------------------------------------------
' FUNCTION: GetUNCShareName
'
' Given a UNC names, returns the leftmost portion of the
' directory representing the machine name and share name.
' E.g., given "\\SCHWEIZ\PUBLIC\APPS\LISTING.TXT", returns
' the string "\\SCHWEIZ\PUBLIC"
'
' Returns a string representing the machine and share name
'   if the path is a valid pathname, else returns NULL
'-----------------------------------------------------------
'
Function GetUNCShareName(ByVal strFN As String) As Variant
    GetUNCShareName = Null
    If IsUNCName(strFN) Then
        Dim iFirstSeparator As Integer
        iFirstSeparator = InStr(3, strFN, "\")
        If iFirstSeparator > 0 Then
            Dim iSecondSeparator As Integer
            iSecondSeparator = InStr(iFirstSeparator + 1, strFN, "\")
            If iSecondSeparator > 0 Then
                GetUNCShareName = Left$(strFN, iSecondSeparator - 1)
            Else
                GetUNCShareName = strFN
            End If
        End If
    End If
End Function


'-----------------------------------------------------------
' FUNCTION: IsUNCName
'
' Determines whether the pathname specified is a UNC name.
' UNC (Universal Naming Convention) names are typically
' used to specify machine resources, such as remote network
' shares, named pipes, etc.  An example of a UNC name is
' "\\SERVER\SHARE\FILENAME.EXT".
'
' IN: [strPathName] - pathname to check
'
' Returns: True if pathname is a UNC name, False otherwise
'-----------------------------------------------------------
'
Function IsUNCName(ByVal strPathName As String) As Integer
    Const strUNCNAME$ = "\\//\"        'so can check for \\, //, \/, /\

    IsUNCName = ((InStr(strUNCNAME, Left$(strPathName, 2)) > 0) And _
                 (Len(strPathName) > 1))
End Function
'-----------------------------------------------------------
' FUNCTION: LogSilentMsg
'
' If this is a silent install, this routine writes
' a message to the gstrSilentLog file.
'
' IN: [strMsg] - The message
'
' Normally, this routine is called inlieu of displaying
' a MsgBox and strMsg is the same message that would
' have appeared in the MsgBox


Public Function GetWeekDayName(ByVal USdate As String) As String
'Declare the function
Dim Ret
'Dim ChkDate As Date
'ChkDate = GetSysFormatDate(GetAppFormatDate(USdate))
'Setup an error handler...
Ret = Weekday(USdate, vbSunday)

GetWeekDayName = WeekdayName(Ret, False, vbSunday)

GetWeekDayName = GetWeekDayString(CByte(Ret))

Exit Function

ErrLine:
        MsgBox "GetWeekDayName " & Err.Description
        

End Function


'   This function allows only the chars present in the ValidSet passed to it.
'   AllowOtherCase allows the other case also.
'   Eg. If your valid set contains A and you want to allow "a" also,
'   then pass AllowOtherCase as TRUE

Function AllowKeyAscii(txt As Object, ValidSet As String, Key As Integer, Optional AllowOtherCase As Boolean) As Integer
Dim Count As Integer, I As Integer
Dim Flag As Boolean
Dim TempBuf As String

    ReDim InvalidArr(0)
    
    If Not IsMissing(AllowOtherCase) Then
        If AllowOtherCase Then       'We have to consider the case
            ValidSet$ = UCase(ValidSet$) & LCase(ValidSet)
        End If
    End If

    Flag = 0
    For Count = 1 To Len(ValidSet)
        If Key = Asc(Mid(ValidSet, Count, 1)) Then
            Flag = True
        End If
    Next Count
    

    If Key = 22 Then
        TempBuf = Clipboard.GetText
        For Count = 1 To Len(TempBuf)
            Flag = False
            For I = 1 To Len(ValidSet)
                If Asc(Mid(TempBuf, Count, 1)) = Asc(Mid(ValidSet, I, 1)) Then
                    Flag = True
                    Exit For
                End If
            Next I
           If Flag = False Then
                Exit For
           End If

        Next Count
    End If
    
    If Not Flag Then Key = 0
    
End Function


Public Function GetWeekDayString(DayVal As Byte) As String

If gLangOffSet = wis_KannadaOffset Then
        Select Case DayVal
            Case 1
                GetWeekDayString = "§–¡¿—§–"        'sunday
            Case 2
                GetWeekDayString = " Ÿ÷”¿–‘¿—§–" 'monday
            Case 3
                GetWeekDayString = "¿–‘Åñ–Œ–¿—§–"  'Tuesday
            Case 4
                GetWeekDayString = "Ω‘∑–¿—§–"  'Wednsday
            Case 5
                GetWeekDayString = "ñ–‘§–‘¿—§–" 'Thursaday
            Case 6
                GetWeekDayString = "∆–‘ë–ı¿—§–" 'Friday
            Case 7
                GetWeekDayString = "∆–∫¿—§–"  'Sat
        End Select
    Else
        Select Case DayVal
            Case 1
                GetWeekDayString = "Sunday"
            Case 2
                GetWeekDayString = "Monday"
            Case 3
                GetWeekDayString = "Tueseday"
            Case 4
                GetWeekDayString = "WednsDay"
            Case 5
                GetWeekDayString = "ThursDay"
            Case 6
                GetWeekDayString = "Friday"
            Case 7
                GetWeekDayString = "Saturday"
        End Select
    End If
End Function

' this function checks for the financial year which we have
' declared globally
' the function checks for the date does not go beyond the finacaial year's
' duration
' if not then retrurns false
Public Function TextBoxCheckFinYear(ByRef DateTextBox As TextBox, Optional isSetFocus As Boolean) As Boolean

Dim CheckDate As String

CheckDate = DateTextBox.Text

If CDate(GetSysFormatDate(CheckDate)) <= CDate(GetSysFormatDate(FinIndianEndDate)) And CDate(GetSysFormatDate(CheckDate)) >= CDate(GetSysFormatDate(FinIndianFromDate)) Then
    TextBoxCheckFinYear = True
    Exit Function
End If

MsgBox "Please Specify the Date Within the Financial Year ", vbInformation, wis_MESSAGE_TITLE

If isSetFocus Then ActivateTextBox DateTextBox
    
End Function

Function TextBoxCurrencyValidate(CurTextBox As TextBox, ByVal AcceptZeroes As Boolean, Optional isSetFocus As Boolean) As Boolean

On Error GoTo ErrLine:

Dim MyCur As Currency
Dim Curstr As String

If CurTextBox Is Nothing Then Exit Function

Curstr = CurTextBox.Text

If Trim$(Curstr) = "" Then Curstr = "0"

MyCur = CCur(Curstr)

If Not AcceptZeroes Then If MyCur = 0 Then Err.Raise vbObjectError + 513, , "Set the Focus"

If MyCur < 0 Then Err.Raise vbObjectError + 513, , "Set the Focus"

TextBoxCurrencyValidate = True

Exit Function

ErrLine:
            
MsgBox "Invalid Currency!!"

If isSetFocus Then ActivateTextBox CurTextBox

End Function
Function TextBoxDateValidate(DateTextBox As TextBox, ByVal Delimiter As String, Optional ByVal IsIndian As Boolean, Optional ByVal isSetFocus As Boolean, Optional ByVal isCheckFinYear As Boolean = True) As Boolean

' Get the date, month and year parts.
Dim DayPart As Integer
Dim MonthPart As Integer
Dim YearPart As Integer
Dim DateArray() As String
Dim bLeapYear As Boolean
Dim DateText As String

'Check For The Decimal point in the string.
'If there is any decimal point the cint will

On Error GoTo ErrLine

If DateTextBox Is Nothing Then Err.Raise vbObjectError + 513, , "Invalid Date!"

DateText = DateTextBox.Text

If InStr(1, DateText, ".", vbTextCompare) Then Err.Raise vbObjectError + 513, , "Invalid Date!"

GetStringArray DateText, DateArray(), Delimiter

'Quit if ubound is < 3   - GIRISH 11/1/2000
If UBound(DateArray) <> 2 Then Err.Raise vbObjectError + 513, , "Invalid Date!"

If IsIndian Then
    DayPart = CInt(DateArray(0))
    MonthPart = CInt(DateArray(1))
Else
    DayPart = CInt(DateArray(1))
    MonthPart = CInt(DateArray(0))
End If

YearPart = CInt(DateArray(2))

' The day, month and year should not be 0.
If DayPart = 0 Then Err.Raise vbObjectError + 513, , "Invalid Date!"

If MonthPart = 0 Then Err.Raise vbObjectError + 513, , "Invalid Date!"

'Changed condition from = to < - Girish 11/1/2000
If YearPart < 0 Then Err.Raise vbObjectError + 513, , "Invalid Date!"

'The yearpart should not exceed 4 digits.
If Len(CStr(YearPart)) > 4 Then Err.Raise vbObjectError + 513, , "Invalid Date!"


' The month part should not exceed 12.
If MonthPart > 12 Then Err.Raise vbObjectError + 513, , "Invalid Date!"

' If the year part is only 2 digits long,
' then prefix the century digits.
If Len(CStr(YearPart)) = 2 Then
    'YearPart = Left$(CStr(Year(gStrDate)), 2) & YearPart
    '5 lines added by Girish    11/1/2000
    YearPart = "20" & YearPart
    If Val(YearPart) > 40 Then YearPart = "19" & YearPart
    
End If

' Validations.
Select Case MonthPart
    Case 2  ' Check for February month.
    
        bLeapYear = isLeap(YearPart)
        If Not bLeapYear Then If DayPart > 28 Then Err.Raise vbObjectError + 513, , "Invalid Date!"
        If bLeapYear Then If DayPart > 29 Then Err.Raise vbObjectError + 513, , "Invalid Date!"
            
    Case 4, 6, 9, 11 ' Months having 30 days...
    
        If DayPart > 30 Then Err.Raise vbObjectError + 513, , "Invalid Date!"
        
    Case Else
    
        If DayPart > 31 Then Err.Raise vbObjectError + 513, , "Invalid Date!"
        
End Select

If isCheckFinYear Then _
    If Not TextBoxCheckFinYear(DateTextBox, isSetFocus) Then Exit Function

TextBoxDateValidate = True

Exit Function

ErrLine:

    MsgBox Err.Description
    If isSetFocus Then ActivateTextBox DateTextBox
    
End Function

Public Sub SetFontToControls(CurrentForm As Form)
Dim Ctrl As VB.Control
On Error Resume Next
For Each Ctrl In CurrentForm
    If isFontDotNametoContinue(Ctrl) Then
        Ctrl.Font.Name = gFontName
        If Not TypeOf Ctrl Is ComboBox Then Ctrl.Font.Size = gFontSize
        If TypeOf Ctrl Is ComboBox Then Ctrl.Font.Size = 10
    End If
Next

End Sub


'
Private Function isFontDotNametoContinue(ByVal Ctrl) As Boolean

If TypeOf Ctrl Is Frame Then isFontDotNametoContinue = True
If TypeOf Ctrl Is ListView Then isFontDotNametoContinue = True
If TypeOf Ctrl Is MSFlexGrid Then isFontDotNametoContinue = True
If TypeOf Ctrl Is OptionButton Then isFontDotNametoContinue = True
If TypeOf Ctrl Is StatusBar Then isFontDotNametoContinue = True
If TypeOf Ctrl Is TabStrip Then isFontDotNametoContinue = True
If TypeOf Ctrl Is TreeView Then isFontDotNametoContinue = True

If TypeOf Ctrl Is CheckBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is ComboBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is CommandButton Then isFontDotNametoContinue = True
If TypeOf Ctrl Is DirListBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is DriveListBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is FileListBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is Label Then isFontDotNametoContinue = True
If TypeOf Ctrl Is ListBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is PictureBox Then isFontDotNametoContinue = True
If TypeOf Ctrl Is TextBox Then isFontDotNametoContinue = True
End Function













'
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
GetSereverDate = GetSysFormatDate(DateStr)
GetSereverDate = Format(Now, "DD/MM/YYYY")
End Function


'This function returns position of an occurence of a search string
'within another string being searched. And it will search from Right to left
Public Function InstrRev(ByVal strString1 As String, ByVal strString2 As String, Optional lngStartpos As Integer, Optional Compare As VbCompareMethod) As Long

InstrRev = 0

'Declaring the variables
Dim Pos As Long
Dim I As Integer
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
Dim I As Integer
For I = Len(string1) To 1 Step -1
    strRev = strRev + Mid(string1, I, 1)
Next I
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
Function UpdateLastAccessedElements(Str As String, strArr() As String, Optional IgnoreCase As Boolean)

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
    For Pos = 0 To UBound(strArr)
        If StrComp(Str, strArr(Pos), CaseVal) = 0 Then
            Flag = True
            Exit For
        End If
    Next Pos
    
    If Not Flag Then Pos = Pos - 1
    
    For Count = Pos To 1 Step -1
        strArr(Count) = strArr(Count - 1)
    Next Count

    strArr(0) = Str
DoEvents
End Function

'
Public Function GetMonthString(ByVal MonVal As Byte) As String
    Const MonStr = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec"
    If MonVal > 12 Or MonVal < 1 Then Exit Function
    GetMonthString = Mid(MonStr, 4 * (MonVal - 1) + 1, 3)
    
    If gLangOffSet = wis_KannadaOffset Then
        Select Case MonVal
            Case 1
                GetMonthString = "°πŸ”¿–ß"  '"jan
            Case 2
                GetMonthString = "ª∞Ÿ”Ω‘ı¿–ß" 'Feb
            Case 3
                GetMonthString = "¿–¶—õ–˛" 'March
            Case 4
                GetMonthString = "ãºıƒ˝" 'April
            Case 5
                GetMonthString = "¿Ÿ‘”" 'May
            Case 6
                GetMonthString = "°÷π–" ' June
            Case 7
                GetMonthString = "°‘ƒŸ" 'July
            Case 8
                GetMonthString = "Éñ–»–˘"  '"agust
            Case 9
                GetMonthString = " Ÿª–ÂÅΩ§–" 'Sept
            Case 10
                GetMonthString = "Éë–÷ÂΩ§–" 'Oct
            Case 11
                GetMonthString = "π–¿–˚ÅΩ§–" 'Nov
            Case 12
                GetMonthString = "Ø∆–ÅΩ§–" 'dec
        End Select
    Else
        Select Case MonVal
            Case 1
                GetMonthString = "Jan"
            Case 2
                GetMonthString = "Feb"
            Case 3
                GetMonthString = "March"
            Case 4
                GetMonthString = "April"
            Case 5
                GetMonthString = "May"
            Case 6
                GetMonthString = "June"
            Case 7
                GetMonthString = "July"
            Case 8
                GetMonthString = "Agust"
            Case 9
                GetMonthString = "Sept"
            Case 10
                GetMonthString = "Oct"
            Case 11
                GetMonthString = "Nov"
            Case 12
                GetMonthString = "Dec"
        End Select
    End If
End Function


Public Function LoadGridSettings(grd As MSFlexGrid, GrdName As String, Filename As String) As Boolean
Dim strIniVal As String
Dim I As Integer
'Prelim Checks
    If Filename = "" Then
        Exit Function
    End If

'strIniVal = ReadFromIniFile(GrdName, "Cols", FileName)
'If Trim$(strIniVal) <> "" Then grd.Cols = Val(strIniVal)

For I = 0 To grd.cols - 1
    strIniVal = ReadFromIniFile(GrdName, "ColWidth" & I, Filename)
    If Trim$(strIniVal) <> "" Then grd.ColWidth(I) = Val(strIniVal)
Next I
LoadGridSettings = True
End Function

'
Function RPad(Str As String, PAdWith As String, LenToPad As Integer) As String
RPad = Str
If LenToPad < Len(Str) Then
    Exit Function
End If
If Len(PAdWith) > 1 Then
    Exit Function
End If
RPad = Str & String(LenToPad - Len(Str), PAdWith)


End Function



' Find and remove the next token from this string.
'
' Tokens are stored in the format:
'    name1(value1)name2(value2)...
' Invisible characters (tabs, vbCrLf, spaces, etc.)
'    are allowed before names.
Sub GetToken(txt As String, token_name As String, _
    token_value As String)
Dim open_pos As Integer
Dim close_pos As Integer
Dim txtlen As Integer
Dim num_open As Integer
Dim I As Integer
Dim ch As String

' Initialize token_name and value.
token_name = ""
token_value = ""

    ' Remove initial invisible characters.
    TrimInvisible txt

    ' If the string is empty, do nothing.
    If txt = "" Then Exit Sub

    ' Find the opening parenthesis.
    open_pos = InStr(txt, "(")
    txtlen = Len(txt)
    If open_pos = 0 Then open_pos = txtlen

    ' Find the corresponding closing parenthesis.
    num_open = 1
    For I = open_pos + 1 To txtlen
        ch = Mid$(txt, I, 1)
        If ch = "(" Then
            num_open = num_open + 1
        ElseIf ch = ")" Then
            num_open = num_open - 1
            If num_open = 0 Then Exit For
        End If
    Next I
    If open_pos = 0 Or I > txtlen Then
        ' There is something wrong.
        Err.Raise vbObjectError + 1, _
            "InventoryItem.GetToken", _
            "Error parsing serialization """ & txt & """"
    End If
    close_pos = I

    ' Get token name and value.
    token_name = Left$(txt, open_pos - 1)
    token_value = Mid$(txt, open_pos + 1, _
        close_pos - open_pos - 1)
    'TrimInvisible token_name
    'TrimInvisible token_value

    ' Remove leading spaces.
    token_name = Trim$(token_name)
    token_value = Trim$(token_value)
    
    ' Remove the token name and value
    ' from the serialization string.
    txt = Right$(txt, txtlen - close_pos)
End Sub

' Remove leading invisible characters from
' the string (tab, space, CR, etc.)
Public Sub TrimInvisible(txt As String)
Dim txtlen As Integer
Dim I As Integer
Dim ch As String

    txtlen = Len(txt)
    For I = 1 To txtlen
        ' See if this character is visible.
        ch = Mid$(txt, I, 1)
        If ch > " " And ch <= "~" Then Exit For
    Next I
    If I > 1 Then _
        txt = Right$(txt, txtlen - I + 1)
End Sub

'
Public Function SaveGridSettings(grd As MSFlexGrid, GrdName As String, Filename As String) As Boolean
Dim Ret As Integer
Dim I As Integer

'Prelim Checks
    If Filename = "" Then
        Exit Function
    End If

Ret = WritePrivateProfileString(GrdName, vbNullString, vbNullString, Filename)
Ret = WritePrivateProfileString(GrdName, "Cols", CStr(grd.cols), Filename)
For I = 0 To grd.cols - 1
    Ret = WritePrivateProfileString(GrdName, "ColWidth" & I, CStr(grd.ColWidth(I)), Filename)
Next I
SaveGridSettings = True

End Function




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


Sub PauseApplication(Secs As Integer)
Dim PauseTime, Start, Finish, TotalTime
    PauseTime = Secs   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer  ' Set end time.
    TotalTime = Finish - Start  ' Calculate total time.
End Sub

Function PutToken(src As String, token_name As String, token_value As String) As String
On Error GoTo Err_line

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

PutToken = strBefore & token_name _
            & "=" & token_value & strAfter


Err_line:
    If Err Then
        MsgBox "Put_token: " & Err.Description, vbCritical
    End If
End Function


' Fills the listview control with the record set data.
Public Function FillView(view As ListView, rs As ADODB.Recordset, Optional AutoWidth As Boolean) As Boolean
On Error GoTo fillview_error
Const FIELD_MARGIN = 1.5
If rs.EOF And rs.BOF Then Exit Function
' Check if there are any records in the recordset.
rs.MoveLast
rs.MoveFirst
If rs.RecordCount = 0 Then
    FillView = True
    GoTo Exit_line
End If

Dim I As Integer
Dim itmX As ListItem

With view
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    Dim x As Integer
    x = 4
    If rs.Fields.Count <= 4 Then x = rs.Fields.Count - 1
    
    For I = 0 To x  'display only selected fields instead of all the fields
        .ColumnHeaders.Add , rs.Fields(I).Name, rs.Fields(I).Name, _
                     view.Parent.TextWidth(rs.Fields(I).Name) * FIELD_MARGIN
        ' Set the alignment characterstic for the column.
        If I > 0 Then
            If rs.Fields(I).Type = adNumeric Or _
                    rs.Fields(I).Type = adInteger Or _
                    rs.Fields(I).Type = adInteger Or _
                    rs.Fields(I).Type = adDouble Or _
                    rs.Fields(I).Type = adCurrency Then
                .ColumnHeaders(I + 1).Alignment = lvwColumnRight
            End If
        End If
        ' If the autowidth property is set,
        ' check if the width of the column is to be adjusted.
    Next

    ' Begin a loop for processing rows.
    Dim KeyField As String
    Do While Not rs.EOF
         KeyField = FormatField(rs.Fields(0))
         DoEvents
         ' Add the details.
        Set itmX = .ListItems.Add(, "KEY" & KeyField, KeyField)
        'Set itmX = .ListItems.Add(, , FormatField(rs.Fields(0)))
        
        ' If the 'Autowidth' property is enabled,
        ' then check if the width needs to be expanded.
        If AutoWidth Then
            If .ColumnHeaders(1).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(FormatField(rs.Fields(0))) Then
                .ColumnHeaders(1).Width = _
                    .Parent.TextWidth(FormatField(rs.Fields(0))) * FIELD_MARGIN
            End If
        End If
        ' Add sub-items.
       ' For I = 1 To rs.fields.Count - 1
           x = 4
    If rs.Fields.Count <= 4 Then x = rs.Fields.Count - 1

       For I = 1 To x 'display only necessary fields to user
            itmX.SubItems(I) = FormatField(rs.Fields(I))
            ' If the 'Autowidth' property is enabled,
            ' then check if the width needs to be expanded.
            If AutoWidth Then
                If .ColumnHeaders(I + 1).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(FormatField(rs.Fields(I))) Then
                    .ColumnHeaders(I + 1).Width = _
                        .Parent.TextWidth(FormatField(rs.Fields(I))) * FIELD_MARGIN
                End If
            End If
        Next

        rs.MoveNext
    Loop
End With
FillView = True

Exit_line:
view.Visible = True
view.view = lvwReport
view.Tag = ""

Exit Function

fillview_error:
    If Err Then
        MsgBox "FillView: The following error occurred." _
            & vbCrLf & Err.Description, vbCritical
        'Resume
    End If
    GoTo Exit_line
End Function

' Fills the listview control with the record set data.
Public Function FillViewNew(view As ListView, rs As ADODB.Recordset, KeyField As String, Optional AutoWidth As Boolean) As Boolean
'Declare the variables
Dim strKey As String
Dim I As Integer
Dim itmX As ListItem
Dim StartPos As Integer
Dim KeyCount As Long
Dim SlNo As Long
Dim x As Integer
Dim MaxFlds As Integer
Dim fldRatio As Single

Const FIELD_MARGIN = 1.4

'Setup an error handler...
On Error GoTo ErrLine
' Check if there are any records in the recordset.
If Not rs.EOF Then rs.MoveLast
If Not rs.BOF Then rs.MoveFirst

view.ListItems.Clear
If rs.EOF And rs.BOF Then
    FillViewNew = True
    Exit Function
End If

With view
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    x = 5
    
    If rs.Fields.Count <= x Then x = rs.Fields.Count - 1
    MaxFlds = x
    fldRatio = .Width / MaxFlds
    fldRatio = fldRatio - 20
    For I = 0 To x  'display only selected fields instead of all the fields
        If rs.Fields(I).Name = KeyField Then I = I + 1
        
        .ColumnHeaders.Add , rs.Fields(I).Name, rs.Fields(I).Name, _
                     view.Parent.TextWidth(rs.Fields(I).Name) * FIELD_MARGIN
        
        
        ' Set the alignment characterstic for the column.
        If I > 0 Then
            If rs.Fields(I).Type = adNumeric Or _
                        rs.Fields(I).Type = adTinyInt Or _
                        rs.Fields(I).Type = adSmallInt Or _
                        rs.Fields(I).Type = adBigInt Or _
                        rs.Fields(I).Type = adSingle Or _
                        rs.Fields(I).Type = adDouble Or _
                        rs.Fields(I).Type = adCurrency Then
                        
               '.ColumnHeaders(I).Alignment = lvwColumnRight
            End If
        End If
        ' If the autowidth property is set,
        ' check if the width of the column is to be adjusted.
    Next

    ' Begin a loop for processing rows.
    KeyCount = 0
    Do While Not rs.EOF
      KeyCount = KeyCount + 1
      DoEvents
      ' Add the details.
      strKey = "KEY" & FormatField(rs.Fields(KeyField))
      If rs.Fields(0).Name = KeyField Then
        Set itmX = .ListItems.Add(, strKey, FormatField(rs.Fields(1)))
        StartPos = 2
      Else
        Set itmX = .ListItems.Add(, strKey, FormatField(rs.Fields(0)))
        StartPos = 1
      End If
      ' If the 'Autowidth' property is enabled,
      ' then check if the width needs to be expanded.
      If AutoWidth Then .ColumnHeaders(1).Width = fldRatio
                 
      ' Add sub-items.
      ' For I = 1 To rs.fields.Count - 1
        x = 5
      If rs.Fields.Count <= 4 Then x = rs.Fields.Count - 1

       For I = StartPos To x 'display only necessary fields to user
            If rs.Fields(I).Name = KeyField Then I = I + 1
            
            itmX.SubItems(I - 1) = FormatField(rs.Fields(I))
            ' If the 'Autowidth' property is enabled,
            ' then check if the width needs to be expanded.
            If AutoWidth Then .ColumnHeaders(I).Width = fldRatio
            
        Next
        rs.MoveNext
    Loop
End With

FillViewNew = True
view.Visible = True
view.view = lvwReport
view.Tag = "KeyField"

Exit Function

ErrLine:
    MsgBox "FillView: " & Err.Description
    

End Function
'
Public Function FormatCurrency(ByVal Curr As Currency) As String
    FormatCurrency = Format(Curr, "#########0.00")
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
If Len(Trim(DateText)) < 5 Then Exit Function

'Check whether delimeter is computer format or passed argument
If InStr(1, DateText, Delimiter) = 0 Then Delimiter = m_DateSep

'Breakup the given string into array elements based on the delimiter.
Dim DateArray() As String
GetStringArray DateText, DateArray(), Delimiter

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


'
Private Function isLeap(Year As Integer) As Boolean

isLeap = ((Year Mod 400) = 0) Or _
    ((Year Mod 4 = 0) And (Year Mod 100 <> 0))

End Function

'
Function CurrencyValidate(Curstr As String, AcceptZeroes As Boolean) As Boolean
On Error GoTo ErrLine
    Dim MyCur As Currency
    If Curstr = "" And AcceptZeroes Then CurrencyValidate = True: GoTo ErrLine
    
    MyCur = CCur(Curstr)
    If Not AcceptZeroes Then If MyCur = 0 Then GoTo ErrLine
    'End If
    
CurrencyValidate = True
Exit Function
ErrLine:

End Function

'
Public Function FormatAccountNumber(AccNo As Long) As String
    FormatAccountNumber = Format(AccNo, "00000")
End Function


Public Sub ActivateTextBox(txtBox As Object)
On Error Resume Next
With txtBox
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
End With
Err.Clear
End Sub

Public Sub ActivateDateTextBox(txtBox As Object)
On Error Resume Next
Dim Pos As Integer
With txtBox
    Pos = InStr(1, .Text, "/")
    If Pos = 0 Then Exit Sub
    .SetFocus
    .SelStart = 0
    .SelLength = Pos - 1
End With
Err.Clear
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
Public Function AddQuotes(FldStr As String, Optional Enclose As Boolean = True) As String
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
Dim I As Integer, ch As String
For I = Len(strFile) To 1 Step -1
    ' Check for "\".
    ch = Mid$(strFile, I, 1)
    If ch = "\" Then
        FilePath = Left$(strFile, I - 1)
        Exit For
    End If
Next

end_line:
    Exit Function

End Function

'
Public Function AppendBackSlash(ByVal StrPath As String) As String
If Right$(StrPath, 1) <> "\" Then
    StrPath = StrPath & "\"
End If
AppendBackSlash = StrPath
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
Function GetStringArray(GivenString As String, strArray() As String, Delim As String) As Integer
GetStringArray = 0
ReDim strArray(0)
If Trim(GivenString) = "" Then Exit Function

strArray = Split(GivenString, Delim)
GetStringArray = UBound(strArray)

Exit Function

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
    If Pos = 0 Then Exit Do
    
    strArray(UBound(strArray)) = Left(TmpStr, Pos - 1)
    'TmpStr = Right(TmpStr, Len(TmpStr) - Pos)
    TmpStr = Mid(TmpStr, Pos + Len(Delim)) 'changed on 27/2/99
    ReDim Preserve strArray(UBound(strArray) + 1)
Loop
    strArray(UBound(strArray)) = TmpStr
    
    GetStringArray = UBound(strArray)
    
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
Public Function FormatDate1(strDate As String) As String

On Error GoTo FormatDateError
' Swap the DD and MM portions of the given date string
Const Delimiter = "/"

Dim TempDelim As String
Dim YearPart As String
Dim strArray() As String

'First Check For the Space in the given string
'Because the Date & Time part will be seperated bt a space
strDate = Trim$(strDate)
Dim SpacePos As Integer

'check for the deimeter
TempDelim = IIf(InStr(1, strDate, m_DateSep), m_DateSep, Delimiter)

SpacePos = InStr(1, strDate, " ")
If SpacePos Then strDate = Left(strDate, SpacePos - 1)

'Breakup the date string into array elements.
'GetStringArray strDate, strArray(), Delimiter
GetStringArray strDate, strArray(), TempDelim

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

'Change the month and day portions and concatenate.
TempDelim = IIf(InStr(1, strDate, m_DateSep), Delimiter, m_DateSep)

FormatDate1 = strArray(1) & TempDelim & strArray(0) & TempDelim & YearPart
'If gIsIndianDate Then FormatDate = strArray(0) & TempDelim & strArray(1) & TempDelim & YearPart

FormatDateError:

End Function

'
Public Function StripExtn(Filename As String) As String
Dim ExtnPos As Integer

' Check for extension
ExtnPos = InStr(Filename, ".")
If ExtnPos = 0 Then ExtnPos = Len(Filename) + 1

' Return the stripped file name.
StripExtn = Mid$(Filename, 1, ExtnPos - 1)

End Function


' -- FormatField:  Formats a given field data
'                  according to its type and returns.
'   Input:  Field object
'   Output: Variant, depends on the data type of the field.
'
Public Function FormatField(fld As Field) As Variant
On Error Resume Next
    If IsNull(fld.Value) Then
        ' If the value in the field is NULL,
        ' return it as a Null String rather than NULL.
        ' This will avoid potential run-time errors.
          FormatField = vbNullString
          ' Check if the field is date type.
          If fld.Type = 2 Or fld.Type = adSingle Or fld.Type = adUnsignedTinyInt Or fld.Type = adInteger Or fld.Type = adDouble Or fld.Type = adInteger Or fld.Type = adCurrency Then
                FormatField = "0"
          End If
    
    Else
        ' Check if the field is date type.
        If fld.Type = adDate Then
            FormatField = Format(fld.Value, "dd/mm/yyyy")
            If InStr(1, DateFormat, "/") Then
                FormatField = Replace(FormatField, "-", "/")
            Else
                FormatField = Replace(FormatField, "/", "-")
            End If
            
            Dim LDate As Date
            LDate = fld.Value
            'FormatField = GetAppFormatDate(LDate)
        ElseIf fld.Type = adCurrency Then
            FormatField = FormatCurrency(fld.Value)
            If FormatField = "" Then FormatField = 0
        Else
            FormatField = fld.Value
        End If
  End If

End Function



'
Public Function FormatDateField(fld As Field) As String
On Error Resume Next
If fld.Type <> adDate Then Exit Function
If IsNull(fld.Value) Then
    ' If the value in the field is NULL,
    ' return it as a Null String rather than NULL.
    ' This will avoid potential run-time errors.
    FormatDateField = "NULL"
Else
    FormatDateField = "#" + CStr(fld.Value) + "#"
End If

End Function




'
Public Function WisDateDiff(IndianDate1 As String, IndianDate2 As String) As Variant
    On Error Resume Next
    WisDateDiff = DateDiff("d", GetSysFormatDate(IndianDate1), GetSysFormatDate(IndianDate2))
    
    If Err Then MsgBox Err.Number & vbCrLf & Err.Description
    
End Function


Public Function WriteParticularstoFile(strParticular As String, Filename As String)
On Error GoTo ErrLine
Dim ParticularsArr() As String
Dim I As Integer

'Update the Particulars combo   'Read to part array
ReDim ParticularsArr(10)

'Read elements of combo to array
I = 1
Do
    'ParticularsArr(I) = cmbParticulars.List(I)
    ParticularsArr(I) = ReadFromIniFile("Particulars", "Key" & I + 1, Filename)
    If ParticularsArr(I) = "" Then Exit Do
    I = I + 1
    If I = 10 Then Exit Do
Loop

'Update last accessed elements
Call UpdateLastAccessedElements(strParticular, ParticularsArr, True)

'Write to
For I = 0 To UBound(ParticularsArr)
    If Trim$(ParticularsArr(I)) <> "" Then
        Call WriteToIniFile("Particulars", "Key" & I, CStr(ParticularsArr(I)), Filename)
    End If
Next I

ErrLine:
End Function

