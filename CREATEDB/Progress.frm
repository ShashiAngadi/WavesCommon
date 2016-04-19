VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress.... "
   ClientHeight    =   1035
   ClientLeft      =   2235
   ClientTop       =   3630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   -30
      TabIndex        =   0
      Top             =   -120
      Width           =   4785
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   3540
         TabIndex        =   1
         Top             =   780
         Width           =   1035
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   450
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CancelClicked()


Private Sub cmdCancel_Click()
Dim MousePointer As Integer
MousePointer = Screen.MousePointer

MsgBox "Cancel the process"
RaiseEvent CancelClicked
Unload Me

Screen.MousePointer = MousePointer

End Sub
Private Sub Form_Load()

ProgressBar1.Align = vbAlignLeft
ProgressBar1.Visible = True

End Sub

