VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   2370
   ClientTop       =   1440
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   7740
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   345
      Left            =   570
      TabIndex        =   5
      Top             =   5820
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2430
      TabIndex        =   4
      Text            =   "Text1 Éœ –‘"
      Top             =   5790
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   5505
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9710
      _Version        =   393216
      MergeCells      =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   5940
      TabIndex        =   1
      Top             =   5970
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   4860
      TabIndex        =   0
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1950
      TabIndex        =   3
      Top             =   5790
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function TestFile()

Dim TxtStream As TextStream
Dim flObject As New FileSystemObject

Set TxtStream = flObject.CreateTextFile("C:\TestFile", True)
Dim str As String
Dim I As Integer
Dim X As Byte

Dim Symbol As String
Text1 = ""

With TxtStream

    str = Chr(255) ' & Chr(64)
    .WriteLine str
    str = Chr(27) & Chr(64)
    .write str
    str = Chr(27) & "x0" 'DRaft quality
    .write str
    str = Chr(27) & Chr(65) & Chr(6)   'set 6(n)/72 line spacing
    .write str
    'str = Chr(27) & Chr(67) & Chr(10) '20 lines height
    str = Chr(27) & Chr(67) & Chr(0) & Chr(8) '6 inch
    .write str
    str = Chr(27) & Chr(108) & Chr(5)  'Left Mrgin
    .write str
    str = Chr(27) & Chr(81) & Chr(125)  'WIDTH right Margin
    .write str
    str = Chr(27) & Chr(78) & Chr(2)  'bottom margin '2 LINES
    .write str
    
    'str = "10 line height 2 line bottom margin"
    str = "8 inch height 2 line bottom margin"
    .WriteLine str
    '.write Chr$(27) & Chr$(33) & Chr$(24)
    'str = Chr$(27) & Chr$(97) & Chr$(1) & "Waves Information Systems"
    '.WriteLine str
    
    '6 inch = 39 lines
    '6 inch = 39 lines
    
    Dim RowNo As Integer
    RowNo = 1
    For I = 65 To 90
        RowNo = RowNo + 1
        X = IIf(I Mod 2, I, I + 32)
        'str = String(50, " ") & CStr(RowNo) & String(80, Chr(X))
        str = CStr(RowNo) & String(80, Chr(X))
        .WriteLine str
    Next
    For I = 65 To 90
        RowNo = RowNo + 1
        X = IIf(I Mod 2, I + 32, I)
        'str = String(50, " ") & CStr(RowNo) & String(80, Chr(X))
        str = CStr(RowNo) & String(80, Chr(X))
        .WriteLine str
    Next
    For I = 65 To 90
        RowNo = RowNo + 1
        X = IIf(I Mod 2, I, I + 32)
        'str = String(50, " ") & CStr(RowNo) & String(80, Chr(X))
        str = CStr(RowNo) & String(80, Chr(X))
        .WriteLine str
    Next
    For I = 65 To 90
        RowNo = RowNo + 1
        X = IIf(I Mod 2, I, I + 32)
        str = String(50, " ") & CStr(RowNo) & String(80, Chr(X))
        .WriteLine str
    Next
    For I = 65 To 90
        RowNo = RowNo + 1
        X = IIf(I Mod 2, I + 32, I)
        str = String(50, " ") & CStr(RowNo) & String(80, Chr(X))
        'str = String(50, " ") & CStr(RowNo) & String(80, Chr(X))
        .WriteLine str
    Next
    .write Chr$(12)
    .Close
End With

End Function

Private Sub cmdPrint_Click()

Dim clswebGrid As New clsgrdWeb
With clswebGrid
    Set .GridObject = grd
    .CompanyAddress = ""
    .CompanyName = "waves Information systems"
    .ReportTitle = "Test of Font"
    Call clswebGrid.ShowWebView '(grd)

End With

End Sub

Private Sub Command1_Click()
Call TestFile
Exit Sub
Dim PrintClass As New clsPrint
Set PrintClass.DataSource = grd

With PrintClass
    .MarginBottom = 0.5
    .MarginLeft = 1
    .MarginRight = 1
    .CompanyName = "Waves"
    .PageHeight = 12
    .PageWidth = 8
    
    .ShowPrint
End With

End Sub


Private Sub Form_Load()
Dim X As Integer

Dim Fnt As StdFont
With grd
    .Clear
    .rows = 3
    .cols = 7
    .FixedRows = 2
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
    Dim I As Integer
    On Error Resume Next
    .Row = 0
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Sl No"
    .Col = 2: .Text = "Font"
    .Col = 3: .Text = "Font"
    .Col = 4: .Text = "Font"
    .Col = 5: .Text = "Font"
    .Col = 6: .Text = "Font Size"
    .MergeRow(0) = True
    .Row = 1
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Sl No"
    .Col = 2: .Text = "Style"
    .Col = 3: .Text = "Style"
    .Col = 4: .Text = "Name"
    .Col = 5: .Text = "Name"
    .Col = 6: .Text = "Font Size"
    .MergeRow(1) = True
    '.Col = 0: .Text = "Sl No"
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    
End With

For I = 0 To Printer.FontCount - 1
    With grd
        .rows = .rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = .Row - 2
        .Col = 2: .Text = Printer.Fonts(I) '.Name
        .Col = 1: .Text = "Font Name in": .CellFontName = Printer.Fonts(I)
        .Col = 3: .Text = "Font size " & (.Row + 6) Mod 20: .CellFontSize = (.Row + 6) Mod 20
        .Col = 4: .Text = Printer.Fonts(I) '.Name
        .Col = 5: .Text = "Font Name in": .CellFontName = Printer.Fonts(I)
        .Col = 6: .Text = "Font size " & (.Row + 6) Mod 20: .CellFontSize = (.Row + 6) Mod 20

        If Err Then
            Err.Clear
            .CellFontName = "MS Sans Serif"
        End If
    End With
Next



End Sub


