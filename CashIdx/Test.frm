VERSION 5.00
Object = "*\ACashIdx.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   2025
   ClientTop       =   1800
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   1170
      TabIndex        =   1
      Top             =   2460
      Width           =   2715
   End
   Begin CashIndexControl.CashIndex CashIndex1 
      Left            =   1320
      Top             =   3300
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   1140
      TabIndex        =   0
      Top             =   1710
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents CashDet As CashIndex
Attribute CashDet.VB_VarHelpID = -1
Private Sub Command1_Click()

CashIndex1.PaymentMode = Cash
CashIndex1.ExpectedCash = 200
CashIndex1.DialogTitle = "Enter Cash Details"
CashIndex1.Show

End Sub

Private Sub Command2_Click()

CashIndex1.PaymentMode = Cheque
CashIndex1.ExpectedCash = 200
CashIndex1.DialogTitle = "EnterCashDetails"
CashIndex1.Show

End Sub


