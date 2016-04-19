VERSION 5.00
Begin VB.UserControl FilePath 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   LockControls    =   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "FilePath.ctx":0000
   ScaleHeight     =   780
   ScaleWidth      =   480
   ToolboxBitmap   =   "FilePath.ctx":0102
End
Attribute VB_Name = "FilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents m_filePath As frmpath
Attribute m_filePath.VB_VarHelpID = -1

Private m_InitDir As String
Private m_DialogTitle As String
Private m_Path As String
Private m_Caption As String
Private m_CancelError As Boolean
Private m_CancelClicked As Boolean
Public Property Let DialogTitle(strTitle As String)
m_DialogTitle = strTitle
End Property


Public Property Get DialogTitle() As String
DialogTitle = m_DialogTitle
End Property


Public Property Let InitDir(strPath As String)
m_InitDir = strPath
End Property


Public Property Let Path(strPath As String)
'Befor assigning it to the variable
'Check for the existance of path
'If path does not exists '
'then take path of init dir
m_Path = strPath
End Property

Public Property Get Path() As String
'Befor returning the value
'Check for the existance of path
'If path does not exists then create the path
Path = m_Path

End Property

Public Sub Show()
'
'set the Object Variable

Set m_filePath = New frmpath
Load m_filePath
m_filePath.Caption = m_DialogTitle

m_filePath.Show vbModal

m_Path = m_filePath.txtPath

Set m_filePath = Nothing
End Sub

Private Sub m_filePath_cancelClicked()
m_CancelClicked = True
End Sub

Private Sub UserControl_Initialize()

 Set m_filePath = New frmpath
End Sub
Private Sub UserControl_Resize()
UserControl.Width = 400
UserControl.Height = 400
End Sub
