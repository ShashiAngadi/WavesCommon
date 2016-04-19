VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Member As New Collection

Private mStrId As String
Public Function GetProperty() As String

Dim Propname As String
Dim PropId As String

'Initilise the return Value
GetProperty = ""

Propname = "Common Database Utilities for Testing"

'Return Value
GetProperty = Propname
End Function
Property Get ID() As String
'Initilise the return value
ID = ""
ID = mStrId

End Property
Property Let NewId(newVal As String)

Dim StrName As Boolean

If Not StrName Then
    StrName = True
End If


End Property





