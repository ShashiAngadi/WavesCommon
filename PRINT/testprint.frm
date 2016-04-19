VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   2655
   ClientTop       =   2175
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   6585
   Begin VB.PictureBox GridPrint1 
      Height          =   405
      Left            =   480
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   5370
      Width           =   405
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4890
      TabIndex        =   2
      Top             =   5160
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   5160
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4755
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8387
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Me.GridPrint1.GridObject = grd
GridPrint1.PrintGrid

End Sub



Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
grd.Clear
With grd
    .Clear
    .Rows = 15
    .Cols = 4
    .AllowUserResizing = flexResizeBoth
    .FixedCols = 1: .FixedRows = 1
    .Row = 0:
    .FormatString = "SL No|Name|Address|City|Balance"
End With
Dim Count As Integer
Dim MaxCount As Integer

MaxCount = grd.Cols - 1
For Count = 1 To MaxCount
    With grd
        .Row = Count
        .Col = 0: .Text = Count
        .Col = 1: .Text = "Name " & Count
        .Col = 2: .Text = "Address " & Count
        .Col = 3: .Text = "City " & Count
        .Col = 4: .Text = 10000 + 100 * Count
    End With
Next
End Sub


