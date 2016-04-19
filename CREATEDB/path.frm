VERSION 5.00
Begin VB.Form frmpath 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   2130
   ClientTop       =   2025
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   1725
      TabIndex        =   6
      Top             =   2490
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   2475
      Width           =   1005
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   210
      TabIndex        =   4
      Top             =   1200
      Width           =   3540
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   810
      Width           =   3525
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   435
      Width           =   3570
   End
   Begin VB.Label Label1 
      Caption         =   "DataBase Path :"
      Height          =   255
      Left            =   75
      TabIndex        =   1
      Top             =   435
      Width           =   1350
   End
   Begin VB.Label lblDatabase 
      Caption         =   "Select Database Path"
      Height          =   255
      Left            =   -15
      TabIndex        =   0
      Top             =   60
      Width           =   6015
   End
End
Attribute VB_Name = "frmpath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gWorkDir As String
Dim M_MaxCount As Long
Dim StepCount As Integer

Public Event OkClicked()
Public Event CancelClicked()






Private Sub cmdok_Click()
RaiseEvent OkClicked

Me.Hide
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdCancel_Click()
RaiseEvent CancelClicked

txtPath = ""
Me.Hide
End Sub


Private Sub Dir1_Change()
txtPath = Dir1.Path
End Sub

Private Sub Drive1_Change()
RetryLine:
On Error GoTo ErrLine
Me.Dir1.Path = Drive1

Exit Sub

ErrLine:

If Err.Number = 68 Then
    If MsgBox("Drive is Not ready", vbRetryCancel + vbInformation + vbDefaultButton2, "Drive Error") = vbRetry Then
        GoTo RetryLine
    End If
End If
End Sub

