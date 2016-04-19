Attribute VB_Name = "Win32API"
Option Explicit

' ---Winapi decl --
'SDA
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long

'SDA
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
'Exit Windows
Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'''Not using
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function ShowOwnedPopups Lib "user32" (ByVal hwnd As Long, ByVal fShow As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

'Fonts Sda
Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long

' Declaration for APIs used by to Check isAlreadyFileOpen
Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

      ' Constant declarations:
    Const VK_NUMLOCK = &H90
    Const VK_SCROLL = &H91
    Const VK_CAPITAL = &H14
    Const KEYEVENTF_EXTENDEDKEY = &H1
    Const KEYEVENTF_KEYUP = &H2
    Const VER_PLATFORM_WIN32_NT = 2
    Const VER_PLATFORM_WIN32_WINDOWS = 1
    
Private Type KeyboardInput       '   typedef struct tagINPUT {
   dwType As Long                '     DWORD type;
   wVK As Integer                '     union {MOUSEINPUT mi;
   wScan As Integer              '            KEYBDINPUT ki;
   dwFlags As Long               '            HARDWAREINPUT hi;
   dwTime As Long                '     };
   dwExtraInfo As Long           '   }INPUT, *PINPUT;
   dwPadding As Currency         '
End Type
Public Enum Win_Keys
    winScrlLock = &H91
    winNumLock = &H90
    winCapsLock = &H14
End Enum



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
    lRetVal = WritePrivateProfileString(Section, Key, KeyData, IniFileName)
    If lRetVal > 0 Then WriteToIniFile = True

End Function

Public Sub ToggleWindowsKey(keyName As Win_Keys, turnON As Boolean)
    Dim currentState As Boolean
    currentState = WindowsKeyState(keyName)
    
    If currentState = turnON Then Exit Sub
    
    PressWindowsKey (keyName)
        
End Sub
Private Sub PressWindowsKey(keyName As Win_Keys)
    GenerateKeyboardEvent keyName, 0
    GenerateKeyboardEvent keyName, 2
End Sub
Public Function WindowsKeyState(keyName As Win_Keys) As Boolean
   ' Determine whether CAPSLOCK/SCLLCK/NUMLOC key is toggled on.
   WindowsKeyState = CBool(GetKeyState(keyName) And 1)
End Function
Private Sub GenerateKeyboardEvent(VirtualKey As Long, Flags As Long)
    Dim kevent As KeyboardInput

    With kevent
        .dwType = 1 'INPUT_KEYBOARD
        .wScan = MapVirtualKey(VirtualKey, 0)
        .wVK = VirtualKey
        .dwTime = 0
        .dwFlags = Flags
    End With
    SendInput 1, kevent, Len(kevent)
End Sub


