VERSION 5.00
Begin VB.Form frmpath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3225
   ClientTop       =   1905
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   3945
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   210
      TabIndex        =   7
      Top             =   2520
      Width           =   3570
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   1710
      TabIndex        =   6
      Top             =   4830
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   4815
      Width           =   1005
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   210
      TabIndex        =   4
      Top             =   1170
      Width           =   3570
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   840
      Width           =   3570
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



Private Sub cmdok_Click()
Me.Hide
End Sub

Private Sub cmdCancel_Click()
txtPath = ""
Me.Hide
End Sub


Private Sub Dir1_Change()
txtPath = Dir1.Path
File1.Path = Dir1.Path
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

Private Sub File1_Click()
'Dim i As Long
    If Right$(Dir1.Path, 1) <> "\" Then
        txtPath.Text = Dir1.Path & "\" & File1.FileName
    Else
        txtPath.Text = Dir1.Path & File1.FileName
'        For i = 0 To File1.ListCount - 1
'            If File1.Selected(i) Then
'            Debug.Print File1.List(i)
'            End If
'        Next
    End If
End Sub

Private Sub Form_Load()
    File1.Pattern = "*.mdb;*.tab;*.prp;*.rc;*.lyt"
End Sub


