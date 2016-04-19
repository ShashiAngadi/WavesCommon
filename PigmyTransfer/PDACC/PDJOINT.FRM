VERSION 5.00
Begin VB.Form frmPDJoint 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Joint holder names"
   ClientHeight    =   1410
   ClientLeft      =   2970
   ClientTop       =   3915
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      Left            =   1335
      TabIndex        =   3
      Top             =   1050
      Width           =   850
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   2235
      TabIndex        =   4
      Top             =   1050
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   675
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2715
   End
End
Attribute VB_Name = "frmPDJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Status As String


Private Sub cmdCancel_Click()
Me.Status = "CANCEL"
Me.Hide
End Sub


Private Sub cmdOK_Click()
Me.Status = "OK"
Me.Hide
End Sub

Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).Width = Me.ScaleWidth
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub



Public Property Get JointHolders() As String
Dim i As Integer
For i = 0 To Text1.Count - 1
    If Trim$(Text1(i).Text) <> "" Then
        JointHolders = JointHolders & Text1(i).Text & ";"
    End If
Next
End Property
Public Property Let JointHolders(ByVal vNewValue As String)
On Error GoTo err_line
If vNewValue = "" Then Exit Property

' Breakup the string into array components.
Dim strTmp() As String
Dim i As Integer
GetStringArray vNewValue, strTmp(), ";"
For i = 0 To UBound(strTmp)
    If i > Text1.Count - 1 Then Exit For
    Text1(i).Text = strTmp(i)
Next

err_line:
End Property
Private Sub SetKannadaCaption()
Dim Ctrl As Control
For Each Ctrl In Me
Ctrl.Font.Name = gFontName
If Not TypeOf Ctrl Is ComboBox Then
    Ctrl.Font.Size = gFontSize
End If
Next
Me.cmdOK.Caption = LoadResString(gLangOffSet + 1)  '"«Œâ³ú"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 2)  '"ÇðÁðôà"
End Sub

