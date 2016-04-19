VERSION 5.00
Begin VB.Form frmRptDt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIS Report date..."
   ClientHeight    =   1425
   ClientLeft      =   5265
   ClientTop       =   2775
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3360
   Begin VB.CheckBox chkDetails 
      Caption         =   "Details"
      Height          =   405
      Left            =   60
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton cmdEndDate 
      Caption         =   ".."
      Height          =   315
      Left            =   3030
      TabIndex        =   2
      Top             =   540
      Width           =   315
   End
   Begin VB.CommandButton cmdStartDate 
      Caption         =   ".."
      Height          =   285
      Left            =   3030
      TabIndex        =   0
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2550
      TabIndex        =   4
      Top             =   930
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1770
      TabIndex        =   5
      Top             =   930
      Width           =   800
   End
   Begin VB.TextBox txtEndDate 
      Height          =   285
      Left            =   1530
      TabIndex        =   3
      Top             =   540
      Width           =   1455
   End
   Begin VB.TextBox txtStDate 
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      Caption         =   "End Date :"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   570
      Width           =   1455
   End
   Begin VB.Label lblStDate 
      AutoSize        =   -1  'True
      Caption         =   "&Start Date :"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   150
      Width           =   1470
   End
End
Attribute VB_Name = "frmRptDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OKClick(stdate As String, enddate As String)
Public Event CancelClick()

Private Sub SetKannadaCaption()
'declare the variables
Dim Ctrl As Control
' to trap an error
On Error Resume Next
For Each Ctrl In Me
   Ctrl.Font.Name = gFontName
   If Not TypeOf Ctrl Is ComboBox Then
      Ctrl.Font.Size = gFontSize
   End If
Next
cmdOK.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
lblStDate.Caption = LoadResString(gLangOffSet + 87)
lblEndDate.Caption = LoadResString(gLangOffSet + 88)
chkDetails.Caption = LoadResString(gLangOffSet + 147)

End Sub


Private Sub cmdCancel_Click()
RaiseEvent CancelClick
Me.Hide
End Sub

Private Sub cmdEndDate_Click()
With Calendar
    .Left = Screen.Width / 2
    .Top = Screen.Height / 2
    .SelDate = FormatDate(gStrDate)
    .Show vbModal
    Me.txtEndDate.Text = .SelDate
End With

End Sub

Private Sub cmdOK_Click()
'Check For Validate of Dates

If Not DateValidate(txtStDate.Text, "/", True) Then
    'Err.Raise 10012, "Invalid Date"
    MsgBox LoadResString(gLangOffSet + 501), vbCritical, wis_MESSAGE_TITLE
    ActivateTextBox txtStDate
    Exit Sub
End If

If Not DateValidate(txtEndDate.Text, "/", True) Then
    'Err.Raise 10012, "Invalid Date"
    MsgBox LoadResString(gLangOffSet + 501), vbCritical, wis_MESSAGE_TITLE
    ActivateTextBox txtEndDate
    Exit Sub
End If

If WisDateDiff(txtStDate.Text, txtEndDate.Text) < 0 Then
    'Err.Raise 10013, "Invalid date difference"
    MsgBox LoadResString(gLangOffSet + 501), vbCritical, wis_MESSAGE_TITLE
    ActivateTextBox txtEndDate
    Exit Sub
End If
Me.Hide
Screen.MousePointer = vbHourglass
RaiseEvent OKClick(txtStDate.Text, txtEndDate.Text)
Me.Hide
End Sub


Private Sub cmdStartDate_Click()
With Calendar
    .Left = Screen.Width / 2
    .Top = Screen.Height / 2
    .SelDate = "1/4/2000"
    .Show vbModal
    Me.txtStDate.Text = .SelDate
End With
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
' to set kannada fonts
Call SetKannadaCaption

End Sub

