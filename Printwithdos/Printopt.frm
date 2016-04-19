VERSION 5.00
Begin VB.Form frmPrintOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WIS   -   Print Options..."
   ClientHeight    =   2640
   ClientLeft      =   2940
   ClientTop       =   2115
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3735
      Begin VB.OptionButton optPrintAllBegin 
         Caption         =   "Print all pages from the beginning"
         Height          =   250
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
      Begin VB.OptionButton optPrintAllCur 
         Caption         =   "Print all pages from the current page"
         Height          =   250
         Left            =   90
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton optPrintCur 
         Caption         =   "Print only the current page"
         Height          =   250
         Left            =   90
         TabIndex        =   5
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chkPause 
         Caption         =   "Pause between pages."
         Height          =   250
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2925
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Print to MSExcel"
         Height          =   250
         Left            =   90
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2970
      TabIndex        =   1
      Top             =   2190
      Width           =   810
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   2190
      Width           =   810
   End
End
Attribute VB_Name = "frmPrintOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As Integer

Private Sub cmdCancel_Click()
Me.Status = wis_CANCEL
Me.Hide
End Sub

Private Sub cmdPrint_Click()
Me.Status = wis_OK
Me.Hide
End Sub
Private Sub Form_Load()
Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)

End Sub


