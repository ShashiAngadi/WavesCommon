Attribute VB_Name = "basKannada"
Option Explicit
Declare Function WINTOISCII Lib "win2isc.dll" (ByVal inputstr As String) As String
Declare Function NudiStartKeyboardEngine Lib "kannada-nudi.dll" _
    Alias "_NudiStartKeyboardEngineVB@12" (ByVal isGlobal As Boolean, _
    ByVal isMonoLingual As Boolean, _
    ByVal needTrayIcon As Boolean) As Integer
Declare Sub NudiTurnOnScrollLock Lib "kannada-nudi.dll" Alias "_NudiTurnOnScrollLockVB@0" ()
Declare Function NudiStopKeyboardEngine Lib "kannada-nudi.dll" Alias "_NudiStopKeyboardEngineVB@0" () As Integer


Public gFontName As String
Public gFontSize As Single
Public gLangOffSet As Integer
Public Const wis_KannadaOffset = 5000

Public Function ConvertToIscii(AsciStr As String) As String
Dim StrLen As Integer
Dim IsciWord  As String
Dim I As Integer
Dim SingleChar As String * 1
StrLen = Len(AsciStr)
IsciWord = ""
SingleChar = ""
For I = 1 To StrLen
    SingleChar = Hex(Int(Asc(WINTOISCII(Mid(AsciStr, I, 1)))))
    IsciWord = IsciWord & SingleChar
Next I
ConvertToIscii = IsciWord


End Function

Public Sub KannadaInitialize()
gFontName = "MS Sans Serif"
gFontSize = 8
gLangOffSet = 0
Dim rst As ADODB.Recordset
'Include  ..\Shared\wisReg.bas File to the project
'First Get The Lanuage Constant From the Registry

'Get the Language information From Database
gdbTrans.SQLStmt = "Select * From Install Where KeyData = 'Language'"
If gdbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    If UCase(FormatField(rst("ValueData"))) = "KANNADA" Then
        gLangOffSet = wis_KannadaOffset
    End If
End If

If gLangOffSet = wis_KannadaOffset Then
    'First Get the  Windows Path &
    'Akruti Installed Path Form The Win.ini
    Dim AkrutiPath As String
    Dim WinPath As String
'    If NudiStartKeyboardEngine(False, False, True) = 0 Then
'        MsgBox "Cannot Start Nudi", vbInformation
'        Exit Sub
'    End If
    gFontName = "Nudi B-Akshar"
    gFontSize = 11
End If

End Sub
Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gdbTrans Is Nothing Then Set gdbTrans = New clsDBUtils

'Open the data base
    If Not gdbTrans.OpenDB(gAppPath & "\CustReg.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gdbTrans.CreateDB(gAppPath & "\CustReg.TAB", "") Then
            MsgBox "unable to create new DataBase", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\CustReg.MDB"
            End
        End If
    End If



End Sub


