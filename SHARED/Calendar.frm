VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Calendar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WIS - Calendar"
   ClientHeight    =   2535
   ClientLeft      =   4170
   ClientTop       =   2475
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   2880
      _Version        =   524288
      _ExtentX        =   5080
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   2
      Day             =   13
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_SelDate As String

Private Sub Calendar1_DblClick()
m_SelDate = Calendar1.value
Me.Hide
End Sub
Private Sub Calendar1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyEscape
        m_SelDate = ""
    Case vbKeyReturn
        m_SelDate = Calendar1.value
End Select
Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    Cancel = True
    m_SelDate = ""
    Me.Hide
End If

End Sub
Private Sub Form_Resize()
With Calendar1
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
End With
End Sub

Public Property Get SelDate() As String
    SelDate = FormatDate(m_SelDate)
End Property
Public Property Let SelDate(ByVal vNewValue As String)
'Validate the Date Format
    If DateValidate(vNewValue, "/", True) Then
        Calendar1.value = FormatDate(vNewValue)
    End If
End Property
