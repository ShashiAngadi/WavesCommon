VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCancel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceessing ......."
   ClientHeight    =   1290
   ClientLeft      =   7215
   ClientTop       =   6915
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4335
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   345
      Left            =   2760
      TabIndex        =   1
      Top             =   900
      Width           =   1365
   End
   Begin ComctlLib.ProgressBar prg 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message"
      Height          =   525
      Left            =   270
      TabIndex        =   2
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CancelClicked()

Private Sub SetKannadaCaption()
On Error Resume Next
Dim ctrl As Control
For Each ctrl In Me
    If Not TypeOf ctrl Is ProgressBar And _
        Not TypeOf ctrl Is VScrollBar And _
            Not TypeOf ctrl Is Line And _
                Not TypeOf ctrl Is Image Then
                    ctrl.Font.Name = gFontName
                    If Not TypeOf ctrl Is ComboBox Then
                        ctrl.Font.Size = gFontSize
                    End If
    End If
Next ctrl
cmdCancel.Caption = LoadResString(gLangOffSet + 2) 'Cancel
End Sub

Private Sub cmdCancel_Click()
Dim MousePointer As Integer
MousePointer = Screen.MousePointer
Screen.MousePointer = vbDefault
'If MsgBox("Are you sure U want to cancel this process", vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbYes Then
If MsgBox(LoadResString(gLangOffSet + 613), vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
    RaiseEvent CancelClicked
    gCancel = True
   Unload Me

End If
Screen.MousePointer = MousePointer
End Sub


Private Sub Form_Load()
Call SetKannadaCaption
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Me.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   ""(Me.hwnd, False)
End Sub


Private Sub lblMessage_Change()
Me.Refresh
End Sub

