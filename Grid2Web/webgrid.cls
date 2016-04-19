VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgrdWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_WebDoc As HTMLDocument
Private m_frmWeb As frmWeb
Private m_WebTable As HTMLTable
Private m_WebTBody As HTMLTableSection
Private m_WebRow As HTMLTableRow
Private m_WebCell As HTMLTableCell
Private m_DOmElement

Private m_DataSource As MSFlexGrid
Private m_CompanyName As String
Private m_companyAddress As String
Private m_ReportTitle As String

Private m_Font As StdFont
''Declaration s for Dos Print

Dim m_txtStream As TextStream

Public Property Set GridObject(ByVal Grid As Object)
    Set m_DataSource = Grid
End Property


Private Sub PrintRows()

m_WebRow.Style.fontWeight = "Bold"

Dim ColCount  As Integer
Dim RowCount As Long
Dim MaxRow As Long
Dim MaxCol As Integer
Dim StrText As String
Dim Wid As Integer
Dim Align As String

MaxCol = m_DataSource.cols - 1
MaxRow = m_DataSource.rows - 1
If MaxRow > 10 Then MaxRow = 10
For RowCount = m_DataSource.FixedRows To MaxRow
    m_DataSource.Row = RowCount
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, m_DataSource.cols)
    With m_WebRow
        For ColCount = 0 To MaxCol
            With .cells(ColCount)
                With m_DataSource
                    .Col = ColCount
                    StrText = .Text
                    Align = "Left"
                End With
                .innerText = StrText
                .Align = Align '"Centre"
                .Style.fontFamily = m_DataSource.CellFontName
                '.Font.Name = grd.CellFontName
            End With
        Next
    End With
Next


End Sub

Public Property Let ReportTitle(NewValue As String)
m_ReportTitle = NewValue
End Property
Public Property Let CompanyName(NewValue As String)
m_CompanyName = NewValue
End Property

Public Property Let CompanyAddress(NewValue As String)
m_companyAddress = NewValue
End Property

Private Sub SetColoumnTitle()


With m_WebTable
    .Width = "700"
End With
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, m_DataSource.cols)

m_WebRow.Style.fontWeight = "Bold"

Dim ColCount  As Integer
Dim MaxCount As Integer
Dim Title As String
Dim Wid As Integer

MaxCount = m_DataSource.cols - 1
m_DataSource.Row = 0
With m_WebRow
    For ColCount = 0 To MaxCount
        With .cells(ColCount)
            With m_DataSource
                .Col = ColCount
                Wid = (.ColWidth(ColCount) / .Width) * 100
                Title = .Text
            End With
            .innerText = Title
            .Width = Wid & "%" '"7%"
            .Align = "Center"
            .Style.fontWeight = "Bold"
        End With
    Next
End With

End Sub

Private Sub SetColoumnTitle1()

With m_WebTable
    .Width = "700"
End With

If m_DataSource.FixedRows = 0 Then Exit Sub

Dim WebRow() As HTMLTableRow
Dim MaxCol As Integer
Dim MaxRow As Integer
Dim RowNo As Integer
Dim ColNo As Integer
Dim Ht() As Single
Dim Wid() As Single
Dim ColWid As Single
Dim RowHt As Single

Dim SpanRow() As Single
Dim SpanCol() As Single

Dim grdVal() As Byte
'Dim Wid() As Byte

MaxCol = m_DataSource.cols - 1
MaxRow = IIf(m_DataSource.FixedRows - 1 > MaxCol, m_DataSource.FixedRows - 1, MaxCol)


ReDim WebRow(MaxRow)
ReDim grdVal(MaxRow, MaxCol)

ReDim Wid(MaxCol)
ReDim Ht(MaxRow)
ReDim SpanRow(MaxRow)
ReDim SpanCol(MaxCol)

'Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, m_DataSource.cols)
With m_DataSource
    For RowNo = 0 To MaxRow
        Ht(RowNo) = .RowHeight(ColNo)
        For ColNo = 0 To MaxCol
            If ColNo = MaxCol Then Wid(ColNo) = .ColWidth(ColNo)
            grdVal(RowNo, ColNo) = 0
        Next
    Next
End With

Dim Span As Byte
Dim Span2 As Byte

Dim NoOfCol As Integer
Dim NoOfRow As Integer

Dim CurrentRow As Byte

PrintLine:

Span = 0: Span2 = 0
NoOfCol = 0: ColWid = 0
With m_DataSource
    RowNo = CurrentRow
    
    For ColNo = 0 To MaxCol
        .Row = CurrentRow
        .Col = ColNo
        SpanCol(ColNo) = 0
        If .CellWidth > .ColWidth(ColNo) Then
            ColWid = ColWid + .ColWidth(ColNo)
            SpanCol(ColNo) = -1
            If ColWid >= .CellWidth + 15 Then
                SpanCol(ColNo - Span) = Span + 1
                ColWid = 0
                Span = 0
            Else
                Span = Span + 1
                NoOfCol = NoOfCol - 1
            End If
        Else
            ColWid = 0
            Span = 0
        End If
        
        RowHt = 0: .Col = ColNo
        'If SpanRow(ColNo) < 0 Then NoOfCol = NoOfCol - 1: GoTo NextCol
        'If CurrentRow = m_DataSource.FixedRows - 1 Then GoTo NextCol
        'If ColNo = 0 Then
          For RowNo = CurrentRow To m_DataSource.FixedRows - 1
            .Row = RowNo
            'If SpanRow(colNo) < 0 Then GoTo NextRow
            If .CellHeight > .RowHeight(RowNo) Then
                RowHt = RowHt + .RowHeight(RowNo)
                SpanRow(ColNo) = -1
                If RowHt > .RowHeight(RowNo) + 15 Then
                    SpanRow(ColNo) = Span2 + 1
                Else
                    Span2 = Span2 + 1
                    NoOfRow = NoOfRow - 1
                    If CurrentRow Then NoOfCol = NoOfCol - 1
                End If
                
            Else
                RowHt = 0
                Span2 = 0
            End If
NextRow:
            NoOfRow = NoOfRow + 1
          Next
          Span2 = 0
        'End If
    NoOfCol = NoOfCol + 1
NextCol:
    
    Next
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, NoOfCol)
m_WebRow.Style.fontWeight = "Bold"

NoOfCol = 0

m_DataSource.Row = CurrentRow

For ColNo = 0 To MaxCol
    m_DataSource.Col = ColNo
    If SpanCol(ColNo) < 0 Then GoTo NextLoop
    If SpanRow(ColNo) < 0 Then GoTo NextLoop
    With m_WebRow.cells(NoOfCol)
        If SpanCol(ColNo) Then .colSpan = SpanCol(ColNo)
        If SpanRow(ColNo) > 0 Then .rowSpan = SpanRow(ColNo)
        .innerText = m_DataSource.Text
        .Width = CStr(m_DataSource.CellWidth * 100 / m_DataSource.Width) & "%"
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    NoOfCol = NoOfCol + 1
NextLoop:

Next

CurrentRow = CurrentRow + 1
If CurrentRow < m_DataSource.FixedRows Then GoTo PrintLine

End Sub


Private Sub SetReportTitle(ReportTitle As String, GapRows As Integer)

If Len(Trim(ReportTitle)) = 0 Then Exit Sub

Dim DOMElement

If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb: Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document
Dim LoopCount  As Integer

With m_WebDoc
        
    'Set m_WebTable = .body.All("HeaderTbl", 0)
    Do
        LoopCount = LoopCount + 1
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
        If LoopCount >= GapRows Then Exit Do
    Loop
    
    With m_WebRow.cells(0)
        .innerText = ReportTitle
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    
    
End With

End Sub

Public Sub ShowWebView()
'If grd Is Nothing Then Exit Sub
'Set m_DataSource = GridObjet 'Form1.grd

'Call StartHtmlTable("TheTable",
Call SetCompanyTitle
Call StartReport
Call SetColoumnTitle1
Call PrintRows
m_frmWeb.Show 1
End Sub

Public Property Let Font(NewValue As StdFont)
    'lblFont.Font = NewValue
    m_Font = NewValue
End Property

Private Sub StartReportMultiRow()

'Dim DOMElement
If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb: Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

            
'Call StartHtmlTable("CashBook", .body, 7)

With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, m_DataSource.cols)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .rowSpan = 2
        '.innerText = LoadResString(gLangOffset + 33) '"Sl No"
        .Width = "5%"
        .Align = "Center"
    End With
    
    With .cells(1)
        .rowSpan = 2
        '.innerText = LoadResString(gLangOffset + 35) '"Name"
        .Width = "35%"
        .Align = "Center"
    End With
    
    With .cells(2)
        .rowSpan = 2
        '.innerText = LoadResString(gLangOffset + 36) & " " & _
                            LoadResString(gLangOffset + 60) '"AccNo"
        .Width = "9%"
        .Align = "Center"
    End With
    
    With .cells(3)
        .colSpan = 2
        '.innerText = LoadResString(gLangOffset + 40) '"Cash"
        '.Width = "13%"
        .Align = "Center"
        
    End With
    
    With .cells(4)
        .colSpan = 2
        '.innerText = LoadResString(gLangOffset + 270) '"contra"
        '.Width = "13%"
        'If m_RpParentID Then .Width = "15%"
        .Align = "Center"
    End With
End With
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
'    With .cells(0)
'        .innerText = LoadResString(gLangOffSet + 33) '"Sl No"
'        .Width = "7%"
'        .Align = "Center"
'    End With
'
'    With .cells(1)
'        .innerText = LoadResString(gLangOffSet + 35) '"Name"
'        .Width = "38%"
'        .Align = "Center"
'    End With
'
'    With .cells(2)
'        .innerText = LoadResString(gLangOffSet + 36) & " " & _
'                            LoadResString(gLangOffSet + 60) '"AccNo"
'        .Width = "13%"
'        .Align = "Center"
'    End With
    
    With .cells(0)
        '.innerText = LoadResString(gLangOffset + 196) '"Receipts"
        .Width = "15%"
        .Align = "Center"
    End With
    
    With .cells(1)
        '.innerText = LoadResString(gLangOffset + 197) '"Payments"
        .Width = "15%"
        'If m_RpParentID Then .Width = "15%"
        .Align = "Center"
    End With
    
    With .cells(2)
        '.innerText = LoadResString(gLangOffset + 196) '"Receipts"
        .Width = "15%"
        .Align = "Center"
    End With
    
    With .cells(3)
        '.innerText = LoadResString(gLangOffset + 197) '"Payments"
        .Width = "15%"
        .Align = "Center"
    End With
    
End With

End Sub

Private Sub StartReport()
If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb: Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document
'If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document
With m_WebDoc
    Call StartHtmlTable("TheTable", .body, 7)
'    Set m_WebTable = .body.All("HeaderTbl", 0)
    
End With

End Sub

Private Function InsertRowAndCellSToTable(Table As HTMLTable, _
                            Optional CellsCount As Integer) As HTMLTableRow

Dim Cell As Integer
Dim Row As HTMLTableRow
If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb: Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

If m_WebDoc Is Nothing Then Set m_WebDoc = New HTMLDocument

With m_WebDoc

    'Set Row = Table.tBodies(0).insertRow(-1)
    Set Row = .createElement("TR")
    Row.Height = 19
    Table.tBodies(0).appendChild Row
    
    'if user has not specified the cols then get the columns from the table
    If CellsCount = 0 Then CellsCount = Table.cols
    
    For Cell = 0 To CellsCount - 1
        Set m_WebCell = .createElement("TD")
        With Row
            .appendChild m_WebCell
            .Style.fontFamily = m_DataSource.CellFontName
            .Style.FontSize = 13
            'If gLangOffset <> 0 Then .Style.FontSize = 15
        End With
    Next Cell
    
    m_WebCell.innerText = " "

End With

Set InsertRowAndCellSToTable = Row

End Function

Private Sub SetCompanyTitle()

Dim DOMElement
If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb: Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc
    PauseApplication (1)
    
    Call StartHtmlTable("HeaderTbl", m_WebDoc.body, 1)
    If Len(Trim(m_CompanyName)) = 0 Then Exit Sub
    
    m_WebTable.border = 0
        
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 0)
    
    With m_WebRow.cells(0)
        .innerText = m_CompanyName
        .Style.FontSize = "17"
        'If gLangOffSet <> 0 Then .Style.FontSize = "19"
        .Align = "Center"
        .Style.fontWeight = "Bold"
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    End With
    
    With m_WebRow.cells(0)
        .innerText = m_companyAddress
        .Style.FontSize = "11"
        .Align = "Center"
        If Len(.innerText) Then _
            Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    End With
    
End With

End Sub

Sub PauseApplication(Secs As Integer)
Dim PauseTime, Start, Finish, TotalTime
    PauseTime = Secs   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer  ' Set end time.
    TotalTime = Finish - Start  ' Calculate total time.
End Sub


Private Sub StartHtmlTable(ByVal TableID As String, ObjectToAppend, _
            ByVal ColsCount As Integer)

If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb: Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc

    Set m_WebTable = .createElement("Table")
    
    With m_WebTable
        .id = TableID
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FontSize = 10
        .Width = "650"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .cols = ColsCount
    End With
    
     ObjectToAppend.appendChild m_WebTable
    
    Set m_WebTBody = .createElement("TBody")
    m_WebTable.appendChild m_WebTBody
    
End With

End Sub


Private Sub Class_Terminate()
Set m_frmWeb = Nothing
Set m_WebDoc = Nothing
Set m_WebTable = Nothing
End Sub


