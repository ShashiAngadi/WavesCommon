VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmReport 
   Caption         =   "INDEX-2000   -   Report wizard"
   ClientHeight    =   5715
   ClientLeft      =   825
   ClientTop       =   1320
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2850
      TabIndex        =   3
      Top             =   5295
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   4095
      TabIndex        =   2
      Top             =   5295
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   5295
      Width           =   1200
   End
   Begin ComctlLib.ListView lvwReport 
      Height          =   5250
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   9260
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const m_Caption = "INDEX-2000 - Reports "
' Indicates whether the column should expand
' depending upon the content of the column.
Private m_AutoWidth As Boolean

Public Event SaveClick(strSelection As String)
Public Event PrintClick(strSelection As String)
Public Event SelectClick(strSelection As String)

Public m_SelItem As String
'' Status variable for User action.
'Public Status  As String

Private Sub cmdClose_Click()
On Error Resume Next
'RaiseEvent SelectClick(lvwReport.SelectedItem.Text)
RaiseEvent SelectClick(m_SelItem)
'Me.Status = wis_OK
Me.Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'On Error Resume Next
'RaiseEvent SelectClick(lvwReport.SelectedItem.Text)
'Me.Hide

If UnloadMode = vbFormControlMenu Then
    Cancel = True
    'Me.Status = wis_CANCEL
    Me.Hide
End If
End Sub
Private Sub Form_Resize()
Const MARGIN = 50
Const CTL_MARGIN = 15
On Error Resume Next

' Arrange the command buttons.
With cmdClose
    .Left = Me.ScaleWidth - MARGIN - .Width
    .Top = Me.ScaleHeight - MARGIN - .Height
End With
With cmdPrint
    .Left = cmdClose.Left - CTL_MARGIN - .Width
    .Top = cmdClose.Top
End With
With cmdSave
    .Left = cmdPrint.Left - CTL_MARGIN - .Width
    .Top = cmdClose.Top
End With

' Arrange the list view.
With lvwReport
    .Left = MARGIN
    .Top = MARGIN
    .Width = Me.ScaleWidth - 2 * MARGIN
    .Height = Me.ScaleHeight - 2 * MARGIN - CTL_MARGIN - cmdClose.Height
End With

End Sub


Public Property Let Title(ByVal vNewValue As Variant)
If vNewValue <> "" Then
    Me.Caption = m_Caption & "[" & vNewValue & "]"
Else
    Me.Caption = m_Caption
End If
End Property

' Sets the alignment attribute of a column
Public Property Let Alignment(rvntCol As Variant, ByVal vNewValue As Integer)

With lvwReport.ColumnHeaders(vNewValue)
    Select Case vNewValue
        Case lvwColumnCenter
            .Alignment = lvwColumnCenter
        Case lvwColumnLeft
            .Alignment = lvwColumnLeft
        Case lvwColumnRight
            .Alignment = lvwColumnRight
        Case Else
            MsgBox "Invalid value for column alignment!", vbExclamation
    End Select
End With
End Property



Public Property Get AutoWidth() As Boolean
    AutoWidth = m_AutoWidth
End Property
Public Property Let AutoWidth(ByVal vNewValue As Boolean)
    m_AutoWidth = vNewValue
End Property

Private Sub lvwReport_Click()
m_SelItem = lvwReport.SelectedItem.Text
End Sub


Private Sub lvwReport_DblClick()
Call cmdClose_Click
End Sub


