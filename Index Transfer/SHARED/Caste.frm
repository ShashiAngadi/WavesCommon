VERSION 5.00
Begin VB.Form frmCaste 
   Caption         =   "Enter the Caste"
   ClientHeight    =   1230
   ClientLeft      =   2955
   ClientTop       =   2550
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   4215
   Begin VB.Frame freCaste 
      Height          =   705
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   4035
      Begin VB.TextBox txtCasteName 
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label lblCasteName 
         Caption         =   "Caste Name:"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3450
      TabIndex        =   1
      Top             =   900
      Width           =   675
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   285
      Left            =   2700
      TabIndex        =   0
      Tag             =   "ok"
      Top             =   900
      Width           =   705
   End
End
Attribute VB_Name = "frmCaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CasteId As Long

Private Function UpdatingControls() As Boolean

Dim CasteId As Long
Dim CasteName As String
Dim SCST As Boolean
Dim SqlStr As String

CasteName = txtCasteName.Text
'SCST = chkSCST.value

If cmdOk.Tag = "update" Then
    CasteId = m_CasteId
Else
    SqlStr = " SELECT MAX(CasteId) as MaxCastid FROM CasteTab"
    gDbTrans.SQLStmt = SqlStr
    If gDbTrans.SQLFetch > 0 Then
        CasteId = FormatField(gDbTrans.Rst(0)) + 1
    Else
        CasteId = 1
    End If
End If

If cmdOk.Tag = "ok" Then
    SqlStr = " INSERT INTO CasteTab(CasteID,CasteName)" & _
            " VALUES(" & _
            CasteId & "," & _
            AddQuotes(CasteName, True) & _
            ")"
ElseIf cmdOk.Tag = "update" Then
    SqlStr = " UPDATE caste " & _
            " SET CasteName=" & _
            "'" & CasteName & "'" & "," & _
            "SCST=" & SCST & _
            " WHERE CasteId=" & CasteId
End If
            
gDbTrans.SQLStmt = SqlStr
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
Else
    gDbTrans.CommitTrans
    MsgBox "Data Updated Successfully"
End If
            
            
txtCasteName.Text = ""
'chkSCST.value = 0
cmdOk.Caption = "&OK"
cmdOk.Tag = "ok"

UpdatingControls = True

End Function


Private Function ValidatingControls() As Boolean

ValidatingControls = True

If txtCasteName.Text = "" Then
    MsgBox "Please Enter the Caste"
    txtCasteName.SetFocus
    Exit Function
End If

ValidatingControls = True

End Function


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()

If ValidatingControls Then
    Call UpdatingControls
End If

End Sub


Private Sub Form_Load()

End Sub

Private Sub txtCasteName_LostFocus()
''Dim CasteName As String
''Dim SqlStr As String
''Dim rstCaste As Recordset
''
''If cmdOk.Tag = "ok" Then
''    CasteName = txtCasteName.Text
''    SqlStr = " SELECT * " & _
''             " FROM Caste" & _
''             " WHERE CasteName=" & "'" & CasteName & "'"
''    gDbTrans.SQLStmt = SqlStr
''    If gDbTrans.SQLFetch >= 1 Then
''        Set rstCaste = gDbTrans.rst.Clone
''        m_CasteId = rstCaste("CasteId")
''        If rstCaste("SCST") Then
''            chkSCST.value = True
''        End If
''        cmdOk.Caption = "&Update"
''        cmdOk.Tag = "update"
''        txtCasteName.SetFocus
''    Else
''        cmdOk.Caption = "&Ok"
''        cmdOk.Tag = "ok"
''    End If
''End If

End Sub


