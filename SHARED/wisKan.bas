Attribute VB_Name = "basKannada"
Option Explicit
Declare Function WINTOISCII Lib "win2isc.dll" (ByVal inputstr As String) As String
Public gFontName As String
Public gFontSize As Single
Public gLangOffSet As Integer
'public gcuruser
Public Const wis_KannadaOffset = 5000
Public Function ConvertToIscii(AsciStr As String) As String
Dim StrLen As Integer
Dim IsciWord  As String
Dim i As Integer
Dim SingleChar As String * 1
StrLen = Len(AsciStr)
IsciWord = ""
SingleChar = ""
For i = 1 To StrLen
    SingleChar = Hex(Int(Asc(WINTOISCII(Mid(AsciStr, i, 1)))))
    IsciWord = IsciWord & SingleChar
Next i
ConvertToIscii = IsciWord


End Function

Public Sub KannadaInitialize()
gFontName = "MS Sans Serif"
gFontSize = 8
gLangOffSet = 0
'Include  ..\Shared\wisReg.bas File to the project
'First Get The Lanuage Constant From the Registry

'Get the Language information From Database
gDBTrans.SQLStmt = "Select * From Install Where KeyData = 'Language'"
If gDBTrans.SQLFetch > 0 Then
    If UCase(FormatField(gDBTrans.Rst("ValueData"))) = "KANNADA" Then
        gLangOffSet = wis_KannadaOffset
    End If
End If

If gLangOffSet = wis_KannadaOffset Then
    'First Get the  Windows Path &
    'Akruti Installed Path Form The Win.ini
    Dim AkrutiPath As String
    Dim WinPath As String
    WinPath = GetWinDir
    AkrutiPath = ReadFromIniFile("AKRUTI", "DIR", WinPath & "\Win.ini")
    Shell AkrutiPath & "\winaces.exe", vbMinimizedNoFocus
    gFontName = "AkliteKndPadmini"
    gFontSize = 11

End If

End Sub



Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\CustReg.MDB", "") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\CustReg.TAB", "") Then
            MsgBox "unable to create new DataBase", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\CustReg.MDB"
            End
        End If
    End If



End Sub


