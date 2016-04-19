Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public gStrDate As String
Public Const gAppName = "INDEX2000 - PD Acounts"
'Public gDBTrans As clsTransact
Public gDBTrans As clsDBUtils
Public gcurrUser As clsUsers

Public gCancel As Boolean
Public gWindowHandle As Long
Public gCompanyName As String


'Declare Function HelpFile Lib "hhctrl.ocx" Alias "HelpFileA" _
'(ByVal hWnd As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Sub Initialize()
Dim Rst As ADODB.Recordset

gStrDate = Format(Now, "mm/dd/yy")
'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsDBUtils
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\..\appmain\Index 2000.MDB", "WIS!@#") Then
    'If Not gDBTrans.OpenDB("C:\indx2000\appmain\Index 2000.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            
            End
        End If
'        If Not gDBTrans.CreateDB(gAppPath & "\PDAcc.TAB", "") Then
'            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
'            On Error Resume Next
'            Kill gAppPath & "\PDAcc.MDB"
'            End
'        End If
    End If

If gcurrUser Is Nothing Then
    Set gcurrUser = New clsUsers
End If
    

End Sub

Public Sub Main()
    gLangOffSet = 5000
    Call Initialize
    Call KannadaInitialize
    
If gcurrUser Is Nothing Then Set gcurrUser = New clsUsers

gcurrUser.ShowLoginDialog
If Not gcurrUser.LoginStatus Then
    Set gcurrUser = Nothing
    gDBTrans.CloseDB
    Set gDBTrans = Nothing
    Exit Sub
End If
  wisMain.Show
End Sub


