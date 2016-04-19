Attribute VB_Name = "WisReg"
Option Explicit

'---REGISTRY Constants---
        
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_DYN_DATA = &H80000006
    Public Const HKEY_LOCAL_MACHINE = &H80000002

    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const HKEY_USERS = &H80000003
    Public Const REG_OPTION_NON_VOLATILE = 0
    Public Const REG_BINARY = 3
    Public Const REG_SZ = 1                                 ' Unicode nul terminated string
    Public Const REG_DWORD = 4                        ' 32-bit number
    Public Const REG_EXPAND_SZ = 2
    Public Const REG_LINK = 6
    Public Const REG_MULTI_SZ = 7
    Public Const ERROR_SUCCESS = 0
    Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9
    Public Const READ_CONTROL = &H20000
    Public Const STANDARD_RIGHTS_ALL = &H1F0000
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_CREATE_LINK = &H20
    Public Const SYNCHRONIZE = &H100000
    
    Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                           KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                           KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As Any, ByRef lpcbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'   CreateRegistryKey : Creates a specified key string entry in Registry
'                                   under a specified Root Class.
'
'   [Input] :   1.  Handle of the Root Class key identified by one of the following.
'                            HKEY_CLASSES_ROOT
'                            HKEY_CURRENT_CONFIG
'                            HKEY_CURRENT_USER
'                            HKEY_LOCAL_MACHINE
'
'   [Output] :  Returns True if the specified is created.
'                   Returns False if unsuccessful in creating the key.
'
Function CreateRegistryKey(ByVal lpKeyHandle As Long, ByVal szKeyString As String) As Boolean
    Dim lRetLong As Long
    Dim hKey As Long
    Dim dwDisposition As Long
    Dim lKeyData As String

    CreateRegistryKey = False   'Initialize the return value.
    
    lRetLong = RegCreateKeyEx(lpKeyHandle, szKeyString, _
                        0, vbNull, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                         0, hKey, dwDisposition)
    If lRetLong = ERROR_SUCCESS Then CreateRegistryKey = True
End Function

'       GetRegistryValue :  Gets the value of a specified key from registry.
'
'       [Input] :   1.  Handle of the Root class key.
'                      2.  Key Section name.
'                      3.  Sub key string whose value is to be fetched.
'
'       [Returns]:  The value string if successful.
'                       Null string "" if unsuccessful.
'
Function GetRegistryValue(lpKeyHandle As Long, szKeyName As String, szSubKey As String) As String
    Dim lcount As Long
    Dim lRetLng As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim lKeyValType As Long
    Dim lTmpStr As String
    Dim lKeyValSize As Long
    Dim lKeyVal As String

    'Initialize the return value.
    GetRegistryValue = ""
    
    'Open the specified key.
    lRetLng = RegOpenKeyEx(lpKeyHandle, szKeyName, 0, KEY_ALL_ACCESS, hKey)
    If (lRetLng <> ERROR_SUCCESS) Then GoTo EndLine
        
    'Initialize the string variable to fetch the value.
    lTmpStr = String$(1024, 0)
    lKeyValSize = 1024

    'Query the registry.
    lRetLng = RegQueryValueEx(hKey, szSubKey, 0, _
                         lKeyValType, lTmpStr, lKeyValSize)

    If (lRetLng <> ERROR_SUCCESS) Then GoTo EndLine
    ' Added by Rk,9/2/1998,to be removed
    If lKeyValSize = 0 Then
        lKeyValSize = 1
    End If
    'end of add
    If (Asc(Mid(lTmpStr, lKeyValSize, 1)) = 0) Then
        lTmpStr = Left(lTmpStr, lKeyValSize - 1)
    Else
        lTmpStr = Left(lTmpStr, lKeyValSize)
    End If

    Select Case lKeyValType
    Case REG_SZ
        lKeyVal = lTmpStr
    Case REG_DWORD
        For lcount = Len(lTmpStr) To 1 Step -1
            lKeyVal = lKeyVal + Hex(Asc(Mid(lTmpStr, lcount, 1)))
        Next
        lKeyVal = Format$("&h" + lKeyVal)
    End Select
    
    GetRegistryValue = lKeyVal
    lRetLng = RegCloseKey(hKey)
    Exit Function
    
EndLine:
    lRetLng = RegCloseKey(hKey)
End Function

'   DeleteRegistryKey : Deletes a specified key from the registry.
'
'   [Input]     1.  KeyHandle :  Identifies the handle of the main key. The values can be :
'                                           HKEY_CLASSES_ROOT
'                                           HKEY_CURRENT_CONFIG
'                                           HKEY_CURRENT_USER
'                                           HKEY_LOCAL_MACHINE
'
'                 2.    lpSubKey :  Key string that is to be  deleted.
'
'   [Output]    1.  Returns True when the key is not existing.
'                   2.  Returns True when the key is successfully deleted.
'                   3.  Returns False, if unable to delete key.
'
Function DeleteRegistryKey(KeyHandle As Long, lpSubKey As String) As Boolean
    Dim lRetLng As Long
    Dim hKey As Long

'--Open the specified key.
    lRetLng = RegOpenKeyEx(KeyHandle, lpSubKey, 0, KEY_ALL_ACCESS, hKey)
    If (lRetLng <> ERROR_SUCCESS) Then
        'Since the key is not present, we need not delete it.
        'We will just return true.
        DeleteRegistryKey = True
        Exit Function
    End If

'--Delete the key.
    lRetLng = RegDeleteKey(KeyHandle, lpSubKey)
    If (lRetLng <> ERROR_SUCCESS) Then
        DeleteRegistryKey = False
        Exit Function
    End If

    DeleteRegistryKey = True
End Function

Function OpenRegistryKey(ByVal lpKeyHandle As Long, szKeyString As String) As Boolean
Dim lRetLng As Long
Dim hKey As Long

    OpenRegistryKey = False 'Initialize the return value.
    
    lRetLng = RegOpenKeyEx(lpKeyHandle, szKeyString, 0, KEY_ALL_ACCESS, hKey)
    If (lRetLng = ERROR_SUCCESS) Then OpenRegistryKey = True

End Function




'   SetRegistryValue :  Sets the value of the specified key, subkey to the specified value.
'
'   [Input] :   1.  Handle of the Root Class Key.
'                 2.  Key string.
'                 3.  Sub key.
'                 4.  Value string.
'
'
Function SetRegistryValue(ByVal lpKeyHandle As Long, ByVal szKeyString As String, ByVal szSubKey As String, ByVal szValueKey As String) As Boolean

Dim RetLng As Long
Dim hKey As Long
Dim dwDisposition As Long

'First open the key
SetRegistryValue = False
    'Open the registry key
    RetLng = RegOpenKeyEx(lpKeyHandle, szKeyString, 0, KEY_ALL_ACCESS, hKey)
    If (RetLng <> ERROR_SUCCESS) Then
         RetLng = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szKeyString, 0, vbNull, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hKey, dwDisposition)
                If (RetLng <> ERROR_SUCCESS) Then
                    Exit Function
                End If
    End If
    
    'Set the value
    RetLng = RegSetValueEx(hKey, szSubKey, 0, REG_SZ, ByVal szValueKey, CLng(Len(szValueKey)))
        If RetLng <> ERROR_SUCCESS Then
            Exit Function
        End If
    RetLng = RegCloseKey(hKey)
SetRegistryValue = True

End Function




