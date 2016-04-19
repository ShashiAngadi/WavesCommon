VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmLookUp 
   Caption         =   "INDEX-2000   -   Report wizard"
   ClientHeight    =   5715
   ClientLeft      =   1050
   ClientTop       =   1770
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "LookUp.frx":0000
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
      Default         =   -1  'True
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
      _Version        =   327680
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
      MouseIcon       =   "LookUp.frx":0342
      NumItems        =   0
   End
End
Attribute VB_Name = "frmLookUp"
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
Public Event SubItems(strSubItem() As String)

Public m_SelItem As String
Private m_SubItems() As String
'' Status variable for User action.
'Public Status  As String

Private Sub cmdClose_Click()
On Error Resume Next

'RaiseEvent SelectClick(lvwReport.SelectedItem.Text)
'Changed By Shashi on 25/3/00
'Me.lvwReport.ListItems
If CStr(m_SelItem) <> "" Then
    RaiseEvent SelectClick(m_SelItem)
    RaiseEvent SubItems(m_SubItems)
End If
'Me.Status = wis_OK
Me.Hide
'Unload Me
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
'centre the form
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption
ReDim m_SubItems(0)
Screen.MousePointer = vbDefault
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

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
End Sub

Private Sub lvwReport_Click()
Dim Count As Integer
m_SelItem = lvwReport.SelectedItem.Text
ReDim m_SubItems(lvwReport.ColumnHeaders.Count)
m_SubItems(0) = m_SelItem
For Count = 1 To lvwReport.ColumnHeaders.Count - 1
   m_SubItems(Count) = lvwReport.SelectedItem.SubItems(Count)
Next
'lvwReport.ColumnHeaders.Item(1).SubItemIndex
End Sub



Private Sub lvwReport_DblClick()
Call cmdClose_Click
End Sub

Private Sub SetKannadaCaption()
Dim Ctrl As Control
If gLangOffSet = wis_KannadaOffset Then
    For Each Ctrl In Me
        Ctrl.Font.Name = gFontName
        If Not TypeOf Ctrl Is ComboBox Then
            Ctrl.Font.Size = gFontSize
        End If
    Next Ctrl
End If
Me.cmdSave.Caption = LoadResString(gLangOffSet + 7)
Me.cmdPrint.Caption = LoadResString(gLangOffSet + 23)
Me.cmdClose.Caption = LoadResString(gLangOffSet + 11)
End Sub

