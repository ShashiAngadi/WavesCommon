Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public Const gAppName = "INDEX2000 - SB Acounts"
Public gDBTrans As clsTransact
Public gStrDate As String
Public gtranstype As wisTransactionTypes
Public gLangOffSet As Long
Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

If Not gDBTrans.OpenDB("c:\indx2000\appmain\index 2000.mdb", "WIS!@#") Then
MsgBox "Unable To Open Db "
End
End If
gStrDate = CStr(FormatDate(Now))

Exit Sub



'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\Index2000.MDB", "") Then
        'If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
        If MsgBox(LoadResString(gLangOffSet + 605) & vbCrLf & vbCrLf & LoadResString(gLangOffSet + 606), vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\Index2000.TAB", "") Then
            'MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
             MsgBox LoadResString(gLangOffSet + 605), vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\Index2000.MDB"
            
        End If
    End If



End Sub

Public Sub Main()
'gLangOffSet = 1000
Call Initialize
'Call KannadaInitialize
'frmMain.Show
frmInt.Show vbModal
End Sub


