VERSION 5.00
Begin VB.Form frmCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   2295
   ClientTop       =   1800
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Cheque details "
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5175
      Begin VB.TextBox txtSl 
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Text            =   "No."
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtNetAmount 
         Height          =   288
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3705
         Width           =   1116
      End
      Begin VB.TextBox txtExpectedCash 
         Height          =   288
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4050
         Width           =   1116
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Index           =   0
         Left            =   3870
         TabIndex        =   6
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtNo 
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   0
         Left            =   450
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   300
         Left            =   3900
         TabIndex        =   3
         Top             =   4050
         Width           =   1188
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   300
         Left            =   3900
         TabIndex        =   2
         Top             =   4395
         Width           =   1188
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   3900
         TabIndex        =   1
         Top             =   3720
         Width           =   1188
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount: "
         Height          =   210
         Left            =   765
         TabIndex        =   10
         Top             =   3750
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   5340
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Expected Amount: "
         Height          =   210
         Left            =   600
         TabIndex        =   9
         Top             =   4095
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ChequesAccepted(BankName() As String, ChequeNO() As String, ChequeAmount() As Currency)
Public Event ChequesRejected()

Private Sub cmdCancel_Click()
    RaiseEvent ChequesRejected
    Unload Me
End Sub

Private Sub cmdOK_Click()
RaiseEvent ChequesAccepted
Unload Me
End Sub


Private Sub Form_Load()
Dim COunt As Integer
Dim vbGray As Long
vbGray = &HC0C0C0

'Load 5 sets of text boxes...
    For COunt = 1 To 10
        Load txtName(COunt)
        Load txtNo(COunt)
        Load txtAmount(COunt)
        Load txtSl(COunt)
        
        txtSl(COunt).Top = txtSl(COunt - 1).Top + txtSl(COunt - 1).Height
        txtName(COunt).Top = txtName(COunt - 1).Top + txtName(COunt - 1).Height
        txtNo(COunt).Top = txtNo(COunt - 1).Top + txtNo(COunt - 1).Height
        txtAmount(COunt).Top = txtAmount(COunt - 1).Top + txtAmount(COunt - 1).Height
        
        txtSl(COunt).Left = txtSl(COunt - 1).Left
        txtName(COunt).Left = txtName(COunt - 1).Left
        txtNo(COunt).Left = txtNo(COunt - 1).Left
        txtAmount(COunt).Left = txtAmount(COunt - 1).Left
        
        txtSl(COunt).Visible = True
        txtName(COunt).Visible = True
        txtNo(COunt).Visible = True
        txtAmount(COunt).Visible = True
        
        txtSl(COunt).Text = COunt
        txtAmount(COunt).Text = "0.00"
        txtSl(COunt).Enabled = False
    Next COunt
    
    txtSl(0).BackColor = vbGray
    txtName(0).BackColor = vbGray
    txtNo(0).BackColor = vbGray
    txtAmount(0).BackColor = vbGray
    
    txtSl(0).Text = "No."
    txtName(0).Text = "Drawee Bank"
    txtNo(0).Text = "Cheque No."
    txtAmount(0).Text = "Rs."
    
    txtSl(0).Enabled = False
    txtName(0).Enabled = False
    txtNo(0).Enabled = False
    txtAmount(0).Enabled = False
    
End Sub



Private Sub txtAmount_Change(Index As Integer)

Dim NetAmt As Currency
Dim COunt As Integer
On Error GoTo lastline
NetAmt = 0
For COunt = 1 To txtAmount.COunt
    NetAmt = NetAmt + CCur(txtAmount(COunt).Text)
Next COunt
lastline:
txtNetAmount.Text = Format(NetAmt, "#############0.00")

End Sub

Private Sub txtAmount_GotFocus(Index As Integer)
If Index = 0 Then
    Exit Sub
End If

txtAmount(Index).SelStart = 0
txtAmount(Index).SelLength = Len(txtAmount(Index).Text)
End Sub


Private Sub txtAmount_KeyPress(Index As Integer, KeyAscii As Integer)
Dim RetBool As Boolean
Dim Factor As Integer
    RetBool = CashValidateKeyAscii(txtAmount(Index), KeyAscii)

End Sub
Private Sub txtAmount_LostFocus(Index As Integer)
txtAmount(Index).Text = Format(txtAmount(Index).Text, "#############0.00")
End Sub

Private Sub txtName_GotFocus(Index As Integer)
If Index = 0 Then
    Exit Sub
End If
txtName(Index).SelStart = 0
txtName(Index).SelLength = Len(txtName(Index).Text)
End Sub


Private Sub txtNetAmount_Change()
If Val(txtNetAmount.Text) <> Val(txtExpectedCash.Text) Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If

End Sub

Private Sub txtNo_GotFocus(Index As Integer)
If Index = 0 Then
    Exit Sub
End If

txtNo(Index).SelStart = 0
txtNo(Index).SelLength = Len(txtNo(Index).Text)

End Sub


