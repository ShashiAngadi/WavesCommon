VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   0  'None
   Caption         =   "Pop menu"
   ClientHeight    =   2955
   ClientLeft      =   3855
   ClientTop       =   2385
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblMenu 
      Caption         =   "Menu label"
      Height          =   405
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   2025
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event MenuClicked(Index As Integer)

Private Sub SetKannadaCaption()
Dim Ctrl As Control
On Error Resume Next
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next Ctrl

End Sub

Private Sub Form_Load()
Call SetKannadaCaption
End Sub



Private Sub lblMenu_Click(Index As Integer)
RaiseEvent MenuClicked(Index)
End Sub


