VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#2.0#0"; "GRDPRINT.OCX"
Begin VB.Form frmQueryRes 
   Caption         =   "Recordset Query Result"
   ClientHeight    =   5535
   ClientLeft      =   2085
   ClientTop       =   1845
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   6585
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   30
      Top             =   5340
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   285
      Left            =   4710
      TabIndex        =   3
      Top             =   5220
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   5220
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   5220
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4980
      Left            =   30
      TabIndex        =   4
      Top             =   240
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8784
      _Version        =   393216
   End
   Begin VB.Label lblqueryTitle 
      Caption         =   "    Title"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2070
      TabIndex        =   0
      Top             =   -60
      Width           =   2235
   End
End
Attribute VB_Name = "frmQueryRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Initailise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)

Private m_frmCancel As frmCancel


Private Function QueryRes() As Boolean
'Here Show the result set in a grid
Dim Count As Integer
Dim Rst As Recordset
Dim fetch As Boolean
Dim Sqlstr As String

QueryRes = False

'Set m_frmCancel = New frmCancel
'm_frmCancel.Show vbModal
'RaiseEvent Processing("Varifying the Records", 0)
'frmCancel.Refresh
Sqlstr = frmCreatDb.txtQuery.Text
If InStr(1, Sqlstr, "SELECT", vbTextCompare) Then fetch = True
If fetch Then
    gDBTrans.SQLStmt = Sqlstr  'Before this remove vbcrlf
    If gDBTrans.SQLFetch < 1 Then
        MsgBox "No Records to Display"
        Exit Function
    End If
    Set Rst = gDBTrans.Rst.Clone
Else
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = Sqlstr
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.CommitTrans
End If


With grd
    .Clear
    .Rows = Rst.RecordCount + 1
    .FixedRows = 1
    .Row = 0
     
    .Cols = Rst.fields.Count
     
     For Count = 0 To Rst.fields.Count - 1
        .Col = Count
        .Text = Rst.fields(Count).Name
        .CellAlignment = 6
        .CellFontBold = True
    Next
   
   RaiseEvent Initailise(0, Rst.RecordCount)
   RaiseEvent Processing("Putting into the Grid", 0)
   
       While Not Rst.EOF
      .Row = .Row + 1
       For Count = 0 To Rst.fields.Count - 1
            .Col = Count
            .Text = FormatField(Rst(Count))
        Next
        Rst.MoveNext
'        DoEvents
'        Me.Refresh
'        RaiseEvent Processing("Reading the Records", Rst.AbsolutePosition / Rst.RecordCount)
    Wend
End With
QueryRes = True
End Function

Private Function Test1() As Boolean
'Declare the variables
grd.ScrollBars = flexScrollBarBoth
grd.CellBackColor = vbGrayed
'frmQueryRes.lblQueryTitle=db.tabledef.recordcount+1 ---for Each rows title
lblqueryTitle.Caption = "Query Results"
'Initilize the grid
With grd
    .Clear
    .WordWrap = True
    .AllowBigSelection = True
    .AllowUserResizing = flexResizeBoth
    .Rows = 30
    .Cols = 5
    .FixedCols = 1
    .FixedRows = 3
End With

With grd
    .MergeCells = flexMergeRestrictAll
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True

    .Cols = 4
    .Row = 0
    .Col = 1: .Text = " Query Result "
    .CellFontBold = True: .CellAlignment = 4: .CellFontSize = 10
    .Col = 2: .Text = " Query Result "
    .CellFontBold = True: .CellAlignment = 4: .CellFontSize = 10
    .Col = 3: .Text = " Query Result "
    .CellFontBold = True: .CellAlignment = 4: .CellFontSize = 10

    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True

'    .Row = 1
'    .Col = 1: .Text = " Result of your Query from the Table"
'    .CellFontBold = True: .CellAlignment = 4: .CellFontSize = 10
'    .Col = 2: .Text = " Result of your Query from the Table"
'    .CellFontBold = True: .CellAlignment = 4: .CellFontSize = 10
'    .Col = 3: .Text = " Result of your Query from the Table"
'    .CellFontBold = True: .CellAlignment = 4: .CellFontSize = 10
'
'    .MergeCol(1) = True
'    .MergeCol(2) = True
'    .MergeCol(3) = True
'
'    .MergeCol(1) = True
'    .MergeCol(2) = True
'    .MergeCol(3) = True

    .Row = 2
    .Col = 0: .Text = "Sl No": .CellFontBold = True
    .Col = 1: .Text = "Test1": .CellFontBold = True
    .Col = 2: .Text = "Test2": .CellFontBold = True
    .Col = 3: .Text = "Test3": .CellFontBold = True
End With
Test1 = True
End Function
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdPrint_Click()

'Set the print options here
grdPrint.CompanyName = "Waves Information Syatems"
grdPrint.ReportTitle = "Query Results"
grdPrint.GridObject = grd
grdPrint.PrintGrid

End Sub
Private Sub Form_Load()
'Set the Caption
lblqueryTitle.Caption = "Query Results"

grd.ScrollBars = flexScrollBarBoth

grd.CellBackColor = vbGrayed
'Call Test1

If Not QueryRes Then Exit Sub

End Sub
Private Sub grd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If grd.Row = 0 Then
'    'Now Write the Coloumn Size
'    Dim ColCount As Integer
'    For ColCount = 0 To grd.Cols - 1
'        Call SaveSetting(App.EXEName, "GridSize=" & "ColWidthRatio" & ColCount, grd.ColWidth(ColCount) / grd.Width)
'    Next
'End If
End Sub


Private Sub Form_Resize()
Const MARGIN = 50
Const CTL_MARGIN = 15
Const BOTTOM_MARGIN = 600

On Error Resume Next

lblqueryTitle.Top = 0
lblqueryTitle.Left = (Me.Width - lblqueryTitle.Width) / 2
grd.Left = 0
grd.Top = lblqueryTitle.Top + lblqueryTitle.Height
grd.Width = Me.Width - 150
grd.Height = Me.ScaleHeight - lblqueryTitle.Height - BOTTOM_MARGIN

With cmdCancel
    .Left = Me.ScaleWidth - MARGIN - .Width
    .Top = Me.ScaleHeight - MARGIN - .Height
End With

With cmdOk
    .Left = cmdCancel.Left - CTL_MARGIN - .Width
    .Top = Me.ScaleHeight - MARGIN - .Height
End With

With cmdPrint
    .Left = cmdOk.Left - CTL_MARGIN - .Width
    .Top = Me.ScaleHeight - MARGIN - .Height
End With

Call GridResize
End Sub
Private Sub GridResize()

Dim ColWidth As Double
Dim ColCount As Integer

For ColCount = 0 To grd.Cols - 1
    grd.ColWidth(ColCount) = grd.Width * GetSetting(App.EXEName, "GridSize=" & "ColWidthRatio" & ColCount, 1 / grd.Cols)
Next ColCount

Exit Sub

End Sub



