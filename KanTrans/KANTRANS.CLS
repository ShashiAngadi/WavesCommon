VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KanTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function GetIsciName(ByVal strKannada As String, Optional IsBilingual As Boolean = False) As String

If IsBilingual Then GoTo BilingualConversion

Dim ch_array() As Byte
Dim Asc_array() As Byte
Dim Isc_array() As Byte

Dim i As Integer, j As Integer
Dim K As Integer
Dim AscVal As Byte

j = Len(strKannada)
ReDim ch_array(j)
ReDim Asc_array(j)
K = -1
Dim Jump As Boolean
For i = 0 To j - 1
    Jump = True
    If AscVal = 69 Or AscVal = 78 Or AscVal = 79 Then Jump = False
    If AscVal = 85 Or AscVal = 89 Or AscVal = 97 Or AscVal = 79 Then Jump = False
    
    If AscVal > 63 And AscVal < 100 And Jump Then GoTo NEXTCHAR
    
    K = K + 1
    AscVal = Asc(Mid(strKannada, i + 1, 1))
    ch_array(i) = Chr(AscVal)
    Asc_array(i) = AscVal
    'Here get the equalant Name of the Charector
    If AscVal = 37 Then Isc_array(K) = 64
    If AscVal = 78 Then Isc_array(K) = 74
    If AscVal = 79 Then Isc_array(K) = 77
    If AscVal = 110 Then Isc_array(K) = 75
    If AscVal = 111 Then Isc_array(K) = 72
    If AscVal = 85 Then Isc_array(K) = 76
    If AscVal = 85 Then Isc_array(K) = 79
    

NEXTCHAR:

Next









BilgualConversion:





End Function


