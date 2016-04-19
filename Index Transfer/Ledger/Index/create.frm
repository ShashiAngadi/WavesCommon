VERSION 5.00
Begin VB.Form frmCreate 
   Caption         =   "Index Test"
   ClientHeight    =   5325
   ClientLeft      =   3810
   ClientTop       =   1425
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.OptionButton Option6 
      Caption         =   "Option3"
      Height          =   465
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   2955
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option2"
      Height          =   495
      Left            =   3330
      TabIndex        =   9
      Top             =   2190
      Width           =   2865
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option1"
      Height          =   525
      Left            =   3330
      TabIndex        =   8
      Top             =   1530
      Width           =   2835
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   465
      Left            =   270
      TabIndex        =   7
      Top             =   2910
      Width           =   2955
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2220
      Width           =   2865
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   525
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2835
   End
   Begin VB.TextBox txtReturnedName 
      Height          =   345
      Left            =   3030
      TabIndex        =   4
      Text            =   "Returned Name"
      Top             =   210
      Width           =   1815
   End
   Begin VB.TextBox txtReturnID 
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Text            =   "Returned ID"
      Top             =   120
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1740
      TabIndex        =   1
      Top             =   3690
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Top             =   3660
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   495
      Left            =   690
      TabIndex        =   2
      Top             =   810
      Width           =   1755
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdCreate_Click()

Dim IndexBankClass As clsBankAcc
Dim HeadID As Long

Set IndexBankClass = New clsBankAcc

gDbTrans.BeginTrans

HeadID = IndexBankClass.GetHeadIDCreatedOnEnum(DepositCA, "Deposit CA", 0)

gDbTrans.CommitTrans

Set IndexBankClass = Nothing

txtReturnID.Text = HeadID


End Sub


Private Sub txtReturnedName_Change()

End Sub


