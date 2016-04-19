VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Set Database Options  "
   ClientHeight    =   2295
   ClientLeft      =   2475
   ClientTop       =   1455
   ClientWidth     =   6360
   DrawMode        =   14  'Copy Pen
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   6360
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   60
      TabIndex        =   2
      Top             =   510
      Width           =   6255
      Begin VB.TextBox txtRename 
         Height          =   285
         Left            =   3420
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Options.frx":030A
         Top             =   570
         Width           =   2745
      End
      Begin VB.TextBox txtDualTab 
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   570
         Width           =   2745
      End
      Begin VB.ComboBox cmbDatabase 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Text            =   "Renamed Database Name.."
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblColor 
         Caption         =   "DataBase Name"
         Height          =   255
         Left            =   330
         TabIndex        =   7
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label lblTitle 
         Caption         =   "Rename and Make Copy of Database"
         Height          =   255
         Left            =   3450
         TabIndex        =   6
         Top             =   270
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   285
      Left            =   2130
      TabIndex        =   1
      Top             =   2010
      Width           =   2445
   End
   Begin VB.CommandButton cmdDataBase 
      Caption         =   "Select Database name"
      Height          =   435
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   3045
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   0
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Createdbclass  As clsTransact
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub cmdApply_Click()
Dim dupIniFile As String
Dim TempFile As String
Dim TestOpt As Boolean

If m_Createdbclass Is Nothing Then Set m_Createdbclass = New clsTransact

TempFile = txtDualTab.Text

dupIniFile = TempFile

If m_Createdbclass.CheckDbStructure(dupIniFile, txtDualTab, "WIS!@#") Then
    MsgBox "DataBase Rectified"
  Else
    MsgBox "Error in Database checking Check the database properly", vbInformation, wis_MESSAGE_TITLE
  Exit Sub
End If
'If m_Createdbclass.CreateDB(dupIniFile, "WIS!@#") Then
'    MsgBox "Database Created"
'  Else
'   MsgBox "Not Created"
'End If

MsgBox "Still in progress:"
End Sub
Private Sub cmdDataBase_Click()
Dim dbFile As String

With cmdDataBase
  cdb.Filter = "*.mdb"
  cdb.DefaultExt = "mdb"
  cdb.DialogTitle = "Open the Database"
  cdb.ShowOpen
  cdb.CancelError = False
End With
dbFile = cdb.FileName

txtDualTab = dbFile
End Sub

'
Private Sub Form_Click()

'Dim cx, cy, Msg, XPos, YPos As Integer  ' Declare variables.
'    ScaleMode = 3   ' Set ScaleMode to
'            ' pixels.
'    DrawWidth = 5   ' Set DrawWidth.
'    ForeColor = QBColor(4)  ' Set foreground to red.
'    FontSize = 24   ' Set point size.
'    cx = ScaleWidth / 2 ' Get horizontal center.
'    cy = ScaleHeight / 2    ' Get vertical center.
'    Cls ' Clear form.
'    Msg = "Most Common Database Utilitites!"
'    CurrentX = cx - TextWidth(Msg) / 2  ' Horizontal position.
'    CurrentY = cy - TextHeight(Msg) ' Vertical position.
'
'Print Msg   ' Print message.
'    Do
'        XPos = Rnd * ScaleWidth ' Get horizontal position.
'        YPos = Rnd * ScaleHeight    ' Get vertical position.
'        PSet (XPos, YPos), QBColor(Rnd * 15)    ' Draw confetti.
'        DoEvents    ' Yield to other
'    Loop    ' processing.

'these staments  are Not for the database Use

End Sub
Private Sub Form_Load()

'txtTest.BackColor = vbWhite
'cmbColor.AddItem "Red"
'cmbColor.AddItem "Black"
End Sub

