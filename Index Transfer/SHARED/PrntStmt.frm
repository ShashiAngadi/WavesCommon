VERSION 5.00
Begin VB.Form frmPrintTrans 
   Caption         =   "Print Tranascation"
   ClientHeight    =   2085
   ClientLeft      =   2475
   ClientTop       =   2385
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   3885
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   3765
      Begin VB.TextBox txtStDate 
         Height          =   285
         Left            =   1770
         TabIndex        =   4
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtEndDate 
         Height          =   285
         Left            =   1770
         TabIndex        =   7
         Top             =   1140
         Width           =   1455
      End
      Begin VB.CommandButton cmdStartDate 
         Caption         =   ".."
         Height          =   255
         Left            =   3270
         TabIndex        =   5
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmdEndDate 
         Caption         =   ".."
         Height          =   285
         Left            =   3270
         TabIndex        =   8
         Top             =   1140
         Width           =   315
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Print transaction between "
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   450
         Width           =   3075
      End
      Begin VB.OptionButton optLastPrint 
         Caption         =   "Print the transaction from Last Stament to till"
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   3555
      End
      Begin VB.Label lblStDate 
         AutoSize        =   -1  'True
         Caption         =   "&Start Date :"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "End Date :"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1140
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3030
      TabIndex        =   10
      Top             =   1680
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2160
      TabIndex        =   9
      Top             =   1650
      Width           =   800
   End
End
Attribute VB_Name = "frmPrintTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DateClick(StartIndiandate As String, EndIndianDate As String)
Public Event TransClick()
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

If optLastPrint.value Then
    RaiseEvent TransClick
    GoTo Last_Line
End If


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
RaiseEvent DateClick(txtStDate.Text, txtEndDate.Text)
Screen.MousePointer = vbDefault
Last_Line:
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


