VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Index 2000 - Server Installation"
   ClientHeight    =   2805
   ClientLeft      =   1980
   ClientTop       =   1725
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCaption 
      Caption         =   "This is will install Index 2000 server"
      Height          =   1005
      Left            =   390
      TabIndex        =   0
      Top             =   780
      Width           =   5925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
lblCaption.Caption = "This will install INDEX 2000 Server" & vbCrLf & vbCrLf & "Are you sure you want to continue?"
End Sub


