VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWeb 
   Caption         =   "Printing Reports"
   ClientHeight    =   5325
   ClientLeft      =   1410
   ClientTop       =   2100
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6570
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPageSet 
      Caption         =   "Page Setup"
      Height          =   405
      Left            =   2430
      TabIndex        =   3
      Top             =   4770
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   5250
      TabIndex        =   1
      Top             =   4770
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   405
      Left            =   3870
      TabIndex        =   2
      Top             =   4770
      Width           =   1305
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4065
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6315
      ExtentX         =   11139
      ExtentY         =   7170
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents m_HtmlToNavigate As HTMLDocument
Attribute m_HtmlToNavigate.VB_VarHelpID = -1

Public IsDocumentLoaded As Boolean




Private Sub SetKannadaCaption()

'Call SetFontToControls(Me)

'Set Kannada caption all Controls
'lblUserDate.Caption = LoadResString(gLangOffSet + 1)
'lblUserName.Caption = LoadResString(gLangOffSet + 1)
'lblUserPassword.Caption = LoadResString(gLangOffSet + 1)
'cmdLogin.Caption = LoadResString(gLangOffSet + 1)
'cmdCancel.Caption = LoadResString(gLangOffSet + 1)

End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub


Private Sub cmdOK_Click()
'Print the Web Page

'wbp.SetWebBrowser web
'wbp.ReadDlgSettings
'wbp.Orientation = 1
'wbp.[Print]


'Setup an error handler...
On Error GoTo ErrLine

Screen.MousePointer = vbHourglass

'Call web.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT)
Call web.ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT)
Screen.MousePointer = vbDefault

Exit Sub

ErrLine:
    Screen.MousePointer = vbDefault
    MsgBox "Your Computer will not support", vbInformation
    
End Sub

Private Sub cmdPageSet_Click()
'Setup an error handler...
On Error GoTo ErrLine

Call web.ExecWB(OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT)

Exit Sub

ErrLine:
    MsgBox "Your computer will not support this." & vbCrLf & "Upgrade your Internet Explorer.", vbInformation
    

End Sub


Private Sub Form_Load()

web.navigate App.Path & "\material.htm"

'set icon for the form caption
'Set the Icon for the form
'Me.Icon = LoadResPicture(147, vbResIcon)

'If gLangOffSet <> 0 Then SetKannadaCaption
End Sub

Private Sub Form_Resize()

Const MARGIN = 50
Const CTL_MARGIN = 15
Const BOTTOM_MARGIN = 600

On Error Resume Next

web.Left = 0
web.Top = Me.ScaleTop
web.Width = Me.ScaleWidth
web.Height = Me.ScaleHeight - BOTTOM_MARGIN

With cmdCancel
    .Left = Me.ScaleWidth - MARGIN - .Width
    .Top = Me.ScaleHeight - MARGIN - .Height
End With
With cmdOK
    .Left = cmdCancel.Left - CTL_MARGIN - .Width
    .Top = cmdCancel.Top
End With

With cmdPageSet
    .Left = cmdOK.Left - CTL_MARGIN - .Width
    .Top = cmdCancel.Top
End With


End Sub


Private Sub m_HtmlToNavigate_onkeyup()

'Dim myEvent As IHTMLEventObj
'
'Set m_HtmlToNavigate = web.Document
'
'Set myEvent = m_HtmlToNavigate.CreateEventObject()
'If myEvent.KeyCode = vbKeyF5 Then
'    myEvent.KeyCode = 0
'End If

End Sub


Private Sub web_GotFocus()
'Me.SetFocus
End Sub

Private Sub web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
IsDocumentLoaded = True
End Sub


