VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   4455
   ClientTop       =   3765
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   5385
   Begin VB.CommandButton Command1 
      Caption         =   "Dos Print"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   390
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim PrintClass As New clsDosPrint

PrintClass.PrintText ("AS")


End Sub


