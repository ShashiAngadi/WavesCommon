VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   750
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   300
      Width           =   4275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1485
      Left            =   630
      TabIndex        =   0
      Top             =   1050
      Width           =   4725
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   3090
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Label1.Caption = NumberInFigure(Val(Text1))
End Sub


Private Sub Form_Load()
'Call InitProperties
End Sub


