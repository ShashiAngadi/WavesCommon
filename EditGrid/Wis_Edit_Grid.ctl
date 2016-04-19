VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl Wis_Edit_Grid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   5794
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Wis_Edit_Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum UserResize
    flexResizeNone = 0
    flexResizeColumns = 1
    flexResizeRows = 2
    flexResizeBoth = 3
    
End Enum

Private Sub UserControl_Initialize()

End Sub


Private Sub UserControl_Resize()
With grd
    .Left = 0
    .Top = 0
    .Width = UserControl.Width
    .Height = UserControl.Height
End With
End Sub




Public Property Get AllowBigSelection() As Boolean
    
    grd.AllowBigSelection = AllowBigSelection
    
End Property

Public Property Let AllowBigSelection(ByVal vNewValue As Boolean)
grd.AllowBigSelection = vNewValue
End Property

Public Property Get AllowUserResizing() As UserResize
    AllowUserResizing = grd.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal vNewValue As UserResize)
grd.AllowUserResizing = vNewValue
End Property
