VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#4.0#0"; "GRDPRINT.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   2490
   ClientTop       =   2100
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6585
   Begin WIS_GRID_Print.GridPrint GridPrint1 
      Left            =   450
      Top             =   4680
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   2100
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4620
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4305
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   7594
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   5610
      TabIndex        =   1
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   4530
      TabIndex        =   0
      Top             =   4770
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   4620
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
With GridPrint1
    .CompanyName = "Waves Information Systems"
    '.Font = "Nudi B-Akshar"
    .GridObject = grd
    .ReportTitle = "Test"
    .PrintGrid
    
End With
End Sub

Private Sub Form_Load()
Dim X As Integer

Dim Fnt As StdFont

grd.Clear
grd.Rows = 2
grd.Cols = 3
grd.FixedRows = 1
grd.FixedCols = 1
grd.AllowUserResizing = flexResizeBoth
Dim i As Integer
On Error Resume Next
grd.Row = 0
grd.FormatString = "Sl no|font Style | Font NAme |Font Size "
For i = 0 To Printer.FontCount - 1
    With grd
        .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = .Row
        .Col = 2: .Text = Printer.Fonts(i) '.Name
        .Col = 1: .Text = "Font Name in": .CellFontName = Printer.Fonts(i)
        .Col = 3: .Text = "Font size " & (.Row + 6) Mod 20: .CellFontSize = (.Row + 6) Mod 20
        If Err Then
            Err.Clear
            .CellFontName = "MS Sans Serif"
        End If
    End With
Next



End Sub

Private Sub MSFlexGrid1_Click()
End Sub


Private Sub Label1_Click()
Label1.Caption = "Åñå®°ý ÍñÈ ²ñÀú³ðÊðô"
Label1.Font.Name = "AkliteKndPadmini"

Debug.Print GridPrint1.Font.Name
Debug.Print GridPrint1.Font.Size
Debug.Print GridPrint1.Font.Bold

End Sub


