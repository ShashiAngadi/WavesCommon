VERSION 5.00
Begin VB.UserControl GridPrint 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "GridPrint.ctx":0000
   PropertyPages   =   "GridPrint.ctx":0792
   ScaleHeight     =   444.444
   ScaleMode       =   0  'User
   ScaleWidth      =   253.659
   ToolboxBitmap   =   "GridPrint.ctx":07B0
   Begin VB.Label lblFont 
      Caption         =   "Label1"
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   540
      Width           =   645
   End
End
Attribute VB_Name = "GridPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents PrintClass As clsPrint
Attribute PrintClass.VB_VarHelpID = -1
Public Event ProcessCount(Count As Long)
Public Event MaxProcessCount(MaxCount As Long)
Public Event Message(strMessage As String)

Private m_grd As Object
Private m_CompanyName As String
Private m_ReportTitle As String

Const wis_MESSAGE_TITLE = "Grid Printing"


Public Sub CancelProcess()
If PrintClass Is Nothing Then Exit Sub
Call PrintClass.CancelProcess
End Sub
Public Property Get CompanyName() As String
CompanyName = m_CompanyName
End Property
Public Property Let CompanyName(strName As String)
m_CompanyName = strName
End Property
Public Property Let GridObject(ByVal Grid As Object)
    Set m_grd = Grid
End Property
Public Property Get Font() As StdFont
    Set Font = gFont
End Property
Public Property Let Font(NewValue As StdFont)
    'lblFont.Font = NewValue
    gFont = NewValue
End Property
Public Property Get ReportTitle() As String
    ReportTitle = m_ReportTitle
End Property
Public Property Let ReportTitle(strTitle As String)
    m_ReportTitle = strTitle
End Property

Public Sub PrintGrid()
If m_grd Is Nothing Then
    MsgBox "Invalid property set", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If PrintClass Is Nothing Then Set PrintClass = New clsPrint
Set PrintClass.DataSource = m_grd
PrintClass.ReportTitle = m_ReportTitle
PrintClass.CompanyName = m_CompanyName
PrintClass.ShowPrint
Set PrintClass = Nothing
End Sub


Public Property Let SetIndividualCellFont(NewValue As Boolean)
    g_NoCellFont = NewValue
End Property

Private Sub PrintClass_MaxProcessCount(MaxCount As Long)
RaiseEvent MaxProcessCount(MaxCount)
End Sub

Private Sub PrintClass_ProcessCount(Count As Long)
RaiseEvent ProcessCount(Count)
End Sub

Private Sub PrintClass_ProcessingMessage(strMessage As String)
RaiseEvent Message(strMessage)
End Sub

Private Sub UserControl_Initialize()
Set gFont = UserControl.Font

End Sub

Private Sub UserControl_InitProperties()
UserControl.Width = 400
UserControl.Height = 400
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'gFont.Name = PropBag.ReadProperty("FontName", "MS Sans Serif")
'gFont.Size = PropBag.ReadProperty("FontSize", "12")
'gFont.Bold = PropBag.ReadProperty("FontBold", gFont.Bold)
'gFont.Italic = PropBag.ReadProperty("FontItalic", gFont.Italic)
'gFont.Strikethrough = PropBag.ReadProperty("FontStrikethrough", gFont.Strikethrough)
'gFont.Underline = PropBag.ReadProperty("FontUnderline", gFont.Underline)

End Sub

Private Sub UserControl_Resize()
UserControl.Width = 400
UserControl.Height = 400

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Call PropBag.WriteProperty("FontName", gFont.Name, "MS Sans Serif")
'Call PropBag.WriteProperty("FontSize", gFont.Size, "12")
'Call PropBag.WriteProperty("FontBold", gFont.Bold, "False")
'Call PropBag.WriteProperty("FontItalic", gFont.Italic, "False")
'Call PropBag.WriteProperty("FontStrikethrough", gFont.Strikethrough, "false")
'Call PropBag.WriteProperty("FontUnderline", gFont.Underline, "False")
End Sub


