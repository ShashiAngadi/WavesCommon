VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLookUp 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   2100
   ClientTop       =   2040
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5070
      TabIndex        =   2
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   285
      Left            =   4110
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin ComctlLib.ListView LvwReport 
      Height          =   4965
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8758
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const m_Caption = "Export & Import"

Private m_AutoWidth As Boolean
Public Event SelectClick(strSelection As String)
Public Event CancelClik()

Public m_SelItem As String
Private m_SubItems() As String


Public Property Let AutoWidth(ByVal newVal As Boolean)
m_AutoWidth = newVal
End Property

Public Property Let Title(ByVal vNewValue As Variant)
If vNewValue <> "" Then
    Me.Caption = m_Caption & "[" & vNewValue & "]"
Else
    Me.Caption = m_Caption
End If
End Property


Public Property Get AutoWidth() As Boolean
AutoWidth = m_AutoWidth
End Property

' Sets the alignment attribute of a column
Public Property Let Alignment(rvntCol As Variant, ByVal vNewValue As Integer)

With LvwReport.ColumnHeaders(vNewValue)
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

Private Sub cmdCancel_Click()
RaiseEvent CancelClik

Me.Hide
End Sub

Private Sub cmdOk_Click()
On Error Resume Next

RaiseEvent SelectClick(LvwReport.SelectedItem.Text)

Me.Hide
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Me.Caption = Me.Caption & " - " & gBankName
'Call CenterMe(Me)

ReDim m_SubItems(0)
Screen.MousePointer = vbDefault
End Sub

Private Sub lvwReport_Click()
Dim Count As Integer
m_SelItem = LvwReport.SelectedItem.Text
ReDim m_SubItems(LvwReport.ColumnHeaders.Count)
m_SubItems(0) = m_SelItem
For Count = 1 To LvwReport.ColumnHeaders.Count - 1
   m_SubItems(Count) = LvwReport.SelectedItem.SubItems(Count)
Next
'lvwReport.ColumnHeaders.Item(1).SubItemIndex
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
With cmdOk
    .Left = Me.ScaleWidth - MARGIN - .Width
    .Top = Me.ScaleHeight - MARGIN - .Height
End With
With cmdCancel
    .Left = cmdOk.Left - CTL_MARGIN - .Width
    .Top = cmdOk.Top
End With
' Arrange the list view.
With LvwReport
    .Left = MARGIN
    .Top = MARGIN
    .Width = Me.ScaleWidth - 2 * MARGIN
    .Height = Me.ScaleHeight - 2 * MARGIN - CTL_MARGIN - cmdOk.Height
End With

End Sub


Private Sub LvwReport_DblClick()
Call cmdOk_Click

End Sub


