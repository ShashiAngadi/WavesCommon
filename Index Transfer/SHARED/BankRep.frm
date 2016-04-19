VERSION 5.00
Begin VB.Form frmBankReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Bank"
   ClientHeight    =   1215
   ClientLeft      =   3075
   ClientTop       =   3630
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4575
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3630
      TabIndex        =   4
      Top             =   840
      Width           =   825
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   285
      Left            =   2700
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Bank"
      Height          =   645
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   4155
      Begin VB.OptionButton optConsol 
         Caption         =   "Consolidated"
         Height          =   405
         Left            =   2370
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.OptionButton optIndividual 
         Caption         =   "Individual Bank"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmBankReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OKClicked(intSelection As Integer)
Public Event CancelClicked()

Private Sub cmdCancel_Click()
RaiseEvent OKClicked(0)
RaiseEvent CancelClicked
Unload Me
End Sub

Private Sub cmdOk_Click()
RaiseEvent OKClicked(IIf(optIndividual, 1, 2))
Unload Me
End Sub


