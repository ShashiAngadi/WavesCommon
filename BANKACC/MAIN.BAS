Attribute VB_Name = "basMain"
Option Explicit
'Public gtranstype As wisTransactionTypes
Public Sub Initialize()
'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\BankAcc.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            Exit Sub
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\BankAcc.TAB", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\BankAcc.MDB"
        End If
    End If
End Sub
Public Sub Main()
Call Initialize
Call KannadaInitialize
frmBankTest.Show
End Sub


