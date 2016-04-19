Attribute VB_Name = "Win32API"
Option Explicit

 '---Winapi decl --
'SDA
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'SDA
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'Fonts Sda
Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long

'
'       Retrieves the string from the specified INIFILE
'
'       LAST MODIFICATION ON    :   09.06.1999
'       LAST MODIFICATION BY    :   M. Ravindranath.
'
Function ReadFromIniFile(Section As String, Key As String, IniFileName As String) As String
Dim strRet As String
Dim lRetVal As Long

    strRet = String$(512, 0)
    lRetVal = GetPrivateProfileString(Section, Key, "", _
                strRet, Len(strRet), IniFileName)
    If lRetVal = 0 Then
        ReadFromIniFile = ""
    Else
        ReadFromIniFile = Trim$(Left(strRet, lRetVal))
    End If
End Function

Public Function GetWinDir() As String
Dim strWinDir As String
Dim Lret As Long

strWinDir = String(255, Chr(0))
Lret = GetWindowsDirectory(strWinDir, Len(strWinDir))

If Lret > 0 Then
    strWinDir = Left$(strWinDir, Lret)
End If

GetWinDir = strWinDir
End Function
Function WriteToIniFile(Section As String, Key As String, KeyData As String, IniFileName As String) As Boolean
Dim strRet As String
Dim lRetVal As Long

strRet = String$(255, 0)
lRetVal = WritePrivateProfileString(Section, Key, _
                                    KeyData, IniFileName)
If lRetVal > 0 Then WriteToIniFile = True
End Function
