VERSION 5.00
Begin VB.Form frmBankTest 
   Caption         =   "frmBankAccTest"
   ClientHeight    =   2070
   ClientLeft      =   1830
   ClientTop       =   2295
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   6870
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   585
      Left            =   705
      TabIndex        =   1
      Top             =   1065
      Width           =   5565
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch Bank Accounts Module"
      Height          =   585
      Left            =   675
      TabIndex        =   0
      Top             =   195
      Width           =   5595
   End
End
Attribute VB_Name = "frmBankTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdLaunch_Click()
    Dim BankClass As clsBankAcc
    If BankClass Is Nothing Then
        Set BankClass = New clsBankAcc
    End If
    BankClass.Show
End Sub


Private Sub cmdExit_Click()
gDBTrans.CloseDB
 Unload Me
End Sub


