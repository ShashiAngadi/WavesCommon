VERSION 5.00
Object = "*\ACashIdx.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   1725
   ClientTop       =   2985
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6585
   Begin CashIndexControl.CashIndex CashIndex1 
      Left            =   300
      Top             =   330
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ClickHere"
      Height          =   1365
      Left            =   2190
      TabIndex        =   0
      Top             =   2130
      Width           =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
CashIndex1.PaymentMode = Cash
CashIndex1.ExpectedCash = 100
CashIndex1.DialogTitle = "Enter the amount below"
CashIndex1.Show
End Sub
