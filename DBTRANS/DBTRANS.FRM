VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   1605
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdMerge 
      Caption         =   "MERGE"
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   1950
      Width           =   2985
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "SPLIT"
      Height          =   435
      Left            =   1920
      TabIndex        =   0
      Top             =   1380
      Width           =   2985
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSplit_Click()
Dim db As New clsTransact

If Not db.BackUPDB("d:\split\new\mail\share3~1.cab") Then
    MsgBox "Unable to backup"
Else
    MsgBox "Backup Successful"
End If

End Sub


