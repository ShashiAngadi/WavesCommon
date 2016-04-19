VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1545
   ClientLeft      =   3180
   ClientTop       =   5130
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3540
      TabIndex        =   2
      Top             =   1110
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   2490
      TabIndex        =   1
      Top             =   1110
      Width           =   975
   End
   Begin VB.Frame fraTrans 
      Caption         =   "Head Types"
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4485
      Begin VB.CheckBox chkOption 
         Caption         =   "Option 4"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   1785
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Option 3"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1785
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Option 2"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   270
         Width           =   1785
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Option 1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkCicked(intAccType As Integer)
Public Event CancelCicked(intAccType As Integer)

Dim m_intSelected As Integer

Private Sub cmdCancel_Click()
    RaiseEvent CancelCicked(0)
    Unload Me
End Sub


Private Sub cmdOK_Click()
Dim intSelect As Integer
Dim Count As Integer
For Count = 0 To chkOption.Count - 1
    If chkOption(Count).value = vbChecked Then _
        intSelect = intSelect + Val(chkOption(Count).Tag)
Next Count
If intSelect = 0 Then Exit Sub
RaiseEvent OkCicked(intSelect)
Unload Me
End Sub


Private Sub Form_Load()
Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
End Sub

