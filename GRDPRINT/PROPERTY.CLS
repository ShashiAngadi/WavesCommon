VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim IniFile As String
Dim M_Section As String

'Gets the Properties of any of the Classes
'
Public Function ReadProperty(Category As String, PropertyName As String, DefaultValue As Variant) As Variant
Dim Key As String
Dim RetStr As String
Key = UCase(PropertyName)
    
    RetStr = GetAnyIni("Setup", Key, IniFile)
        If RetStr = "" Then
            RetStr = DefaultValue
        End If
ReadProperty = RetStr
End Function

Public Sub WriteProperty(Category As String, PropertyName As String, PropertyValue As Variant)
Dim RetVal As String

    RetVal = WritePrivateProfileString(M_Section, PropertyName, CStr(PropertyValue), IniFile)
   
End Sub

Private Sub Class_Initialize()
    IniFile = App.Path & "\" & "Urb_Bank.ini"
End Sub


