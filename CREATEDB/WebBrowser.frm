VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWebBrowser 
   ClientHeight    =   5325
   ClientLeft      =   2085
   ClientTop       =   2370
   ClientWidth     =   6585
   Icon            =   "WebBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4905
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6285
      ExtentX         =   11086
      ExtentY         =   8652
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_doc As Document
Private Function webBrowser() As Boolean
'initilise the Variables
Dim webCon As webBrowser

'initilise return value
webBrowser = False

 'webCon.Navigate = "c:\windos\desktop\Pri.doc"


webBrowser = True
End Function

Private Sub Form_Load()
'Set the Caption for web window
frmWebBrowser.Caption = " WebBrowser "

web.AddressBar = True

'Navigate
web.Navigate "c:\windows\Desktop\transtype.txt"

web.Refresh

If Not webBrowser Then Exit Sub
End Sub
Private Sub Form_Resize()
'Resizes the form in screen mode

Call GridResize


End Sub


Private Sub GridResize()

End Sub

