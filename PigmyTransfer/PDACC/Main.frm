VERSION 5.00
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDEX 2000"
   ClientHeight    =   3435
   ClientLeft      =   2640
   ClientTop       =   1905
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4935
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   330
      TabIndex        =   5
      Top             =   3060
      Width           =   1515
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Add Pigmy agnt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   810
      TabIndex        =   3
      Top             =   1620
      Width           =   3525
   End
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   4770
      Top             =   720
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   750
      TabIndex        =   2
      Top             =   2190
      Width           =   3585
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch Pigmy Deposit Account Module"
      Height          =   495
      Left            =   810
      TabIndex        =   0
      Top             =   930
      Width           =   3585
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pigmy Deposit Account"
      Height          =   2415
      Left            =   270
      TabIndex        =   4
      Top             =   630
      Width           =   4395
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "              Pigmy  Deposit Account Module Ver."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   -735
      TabIndex        =   1
      Top             =   0
      Width           =   5865
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()

Unload Me
End
End Sub

Private Sub cmdLaunch_Click()
'Call gUser.Login("admin", "admin", "")

Dim PDAcc As New clsPDAcc

PDAcc.Show
End Sub

Private Sub cmdUser_Click()
gcurrUser.ShowUserDialog
End Sub

Private Sub Command1_Click()
'Dim hGet As Long
'
'hGet = HelpFile(hWnd, "Ado210.chm", hhHelpLoad, 0)
End Sub
'
Private Sub Form_Load()

lbl.Caption = lbl.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub

