VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{B44203C3-6FD7-4C5A-B02A-E52525F0ECEA}#1.0#0"; "WISPrint.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   1875
   ClientTop       =   2190
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6585
   Begin WIS_GRID_PrintNew.WISPrint GridPrint1 
      Left            =   360
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

Private Sub cmdClose_Click()
 Unload Me
 
End Sub


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

'
With grd
            .Clear
            .Rows = 5
            .Cols = 16
            .FixedCols = 2
            .FixedRows = 3
            .Row = 0
            .Col = 0: .Text = "Sl No"
            .ColAlignment(0) = flexAlignLeftCenter
            .Col = 1: .Text = "Name"
            .ColAlignment(1) = flexAlignLeftCenter
            .Col = 2: .Text = "Place"
            .ColAlignment(2) = flexAlignLeftCenter
            
            .Col = 3: .Text = "Loan No"
            .ColAlignment(3) = flexAlignLeftCenter
            .Col = 4: .Text = "Loan Date"
            .ColAlignment(4) = flexAlignLeftCenter
            .Col = 5: .Text = "Loan No"
            .Col = 6: .Text = "Loan Amount"
            .ColAlignment(6) = flexAlignRightCenter
            .Col = 7: .Text = "Deposit"
            .Col = 8: .Text = "Subsidy Amount"
            .ColAlignment(8) = flexAlignRightCenter
            .Col = 9: .Text = "Loan Due Date"
            .ColAlignment(9) = flexAlignRightCenter
            
            '----
            .Col = 10: .Text = "Repayment Details"
            .ColAlignment(10) = flexAlignLeftCenter
            .Col = 11: .Text = "Repayment Details"
            .ColAlignment(11) = flexAlignLeftCenter
            .Col = 12: .Text = "Repayment Details"
            .ColAlignment(12) = flexAlignRightCenter
            .Col = 13: .Text = "Repayment Details"
            .ColAlignment(13) = flexAlignRightCenter
            '----
            .Col = 14: .Text = "Subsidy Amount"
            .ColAlignment(14) = flexAlignRightCenter
            .Col = 15: .Text = "Rebate Amount"
            .ColAlignment(15) = flexAlignRightCenter
            '.Col = 16: .Text = "Remark"
            
            .Row = 1
            .Col = 0: .Text = "Sl No"
            .Col = 1: .Text = "Name"
            .Col = 2: .Text = "Place"
            .Col = 3: .Text = "Loan No"
            .Col = 4: .Text = "Loan Date"
            .Col = 5: .Text = "Loan No"
            .Col = 6: .Text = "Loan Amount"
            .Col = 7: .Text = "Deposit"
            .Col = 8: .Text = "Subsidy Amount"
            .Col = 9: .Text = "Loan Due Date"
            '----
            .Col = 10: .Text = "Voucher No"
            .Col = 11: .Text = "Repayment Date"
            .Col = 12: .Text = "Repayment Amount"
            .Col = 13: .Text = "Days"
            '----
            .Col = 14: .Text = "Subsidy Amount"
            .Col = 15: .Text = "Rebate Amount"
            '.Col = 16: .Text = "Remark"
            
            .Row = 2
            For i = 0 To .Cols - 1
                .Col = i
                .Text = i + 1
                .MergeCol(i) = True
            Next
            .TextMatrix(2, 0) = "  1"
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .MergeRow(2) = True
            .MergeRow(1) = True
        End With
'


Dim Str As String
Dim Slno As Integer
Dim total As Long
Slno = 0
Str = ""
grd.Row = grd.FixedRows
For i = 0 To Printer.FontCount - 1
    If InStr(1, Printer.Fonts(i), "SHR", vbTextCompare) Or InStr(1, Printer.Fonts(i), "NUD", vbTextCompare) Then
    
    Str = IIf(Str = "", " ", "")
    With grd
        Slno = Slno + 1
        .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = Slno
        .Col = 2: .Text = Printer.Fonts(i) + Str '.Name
        .Col = 1: .Text = "Font Name in" + Str: .CellFontName = Printer.Fonts(i)
        .Col = 3: .Text = "Font size " & (.Row + 6) Mod 20: .CellFontSize = (.Row + 6) Mod 20
        .Col = 4: .Text = Val(Slno Mod 10)
        total = total + Val(.Text)
        If Err Then
            Err.Clear
            .CellFontName = "MS Sans Serif"
        End If
    End With
    End If
Next

grd.Rows = grd.Rows + 1
grd.Row = grd.Row + 2
grd.Col = 1: grd.Text = "TOTAL"
grd.Col = 4: grd.Text = total

End Sub

Private Sub MSFlexGrid1_Click()
End Sub


Private Sub Label1_Click()
Label1.Caption = "Åñå®°ý ÍñÈ ²ñÀú³ðÊðô"
Label1.Font.Name = "SUCHI-KAN-0850"

Debug.Print GridPrint1.Font.Name
Debug.Print GridPrint1.Font.Size
Debug.Print GridPrint1.Font.Bold

End Sub


Private Sub WISPrint1_ProcessCount(Count As Long)

End Sub
