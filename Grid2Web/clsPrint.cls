VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Define member variables of this print class...
'
Private m_Font As StdFont
Dim m_txtStream As TextStream

Public Event ProcessingMessage(strMessage As String)
Public Event ProcessCount(Count As Long)
Public Event MaxProcessCount(MaxCount As Long)
Private m_CancelProcess As Boolean
Private m_ProcCount  As Long

Private m_RowHeight() As Single
Private m_ColWidth() As Single

' Constants...
Private Const FIELD_MARGIN = 200
Private Const COL_MARGIN = 50
Private Const RECT_MARGIN = 15

Private m_ColLine As Boolean
Private m_RowLine As Boolean

'Declaration of the page margin
Private m_LeftMargin  As Single
Private m_TopMargin  As Single
Private m_RightMargin  As Single
Private m_BottomMargin  As Single

Private m_TitleMargin As Single

Private m_PageWidth As Double
Private m_PageHeight As Double

' Title Object.
Private m_Title As clsField
Private m_clsFooter As clsField
Private m_clsHeader As clsField
Private m_FooterLIne As Boolean
Private m_HeaderLine As Boolean


Private m_WrapHead As Boolean
Private m_WrapCell As Boolean

Private my_printer As Object                  ' Can be a printer or a picturebox object.
'Private m_DataSource As MSFlexGrid            ' Grid to read rows from.
Private m_DataSource As Object
Private m_View As String
'Page Num Details
Public PageNums As Boolean                          ' Prints the pagenos, if set to true
Public ReportDate As Boolean                        ' Prints the date on report, if true.
Public HeaderRectangle As Boolean               ' Prints a rectangle around the header, if true
Private Print_EOF As Boolean                        ' Flag to identify the end of rows.
Private m_PageRow() As Integer

'Font details
Private m_FontSize As Single                ' To Set The Fotn Size And Font NAme for Printing
Private m_FontName As String
Private m_NextPage As Boolean
Private m_MergeText As Boolean


Private title_printed As Boolean            ' Suppresses the title for subsequent pages.

Private rows_per_page As Integer     'Noof rows to be print per page
Private rows_in_first_page As Integer 'Noof rows in the first page
Private rows_in_lastpage As Integer 'Noof rows in the last page
Private num_pages As Integer        'total Num of page excluding split pages
Private page_start_row As Long
Private pages_toPrint() As Long
Private page_printing As Long
Public PageNumber As Integer
Public CompanyName As String             ' This name will appear in the title of reports...

Private pause_between_pages As Boolean      ' Applies while printing to printer.
Private LastPage_Reference As Integer      ' Applies While Moving to next & Previrues pages
Private m_heading_top As Single
Private curPage As Integer
Private curSplitPage As Integer
Private curSplitcolno As Integer
Private saved_view As Integer
Private saved_row As Long
Private saved_page_start_row As Long

'If All coloumns are not fits in
'a page then the remaianing cols will fits into next coloumn
'In such case use below varible- shashi
Private m_PageSplitted As Boolean
Private m_SplitPageNo As Integer
Private m_NoOfSplitPage As Integer
Private m_SplitCol() As Integer
Private Function GetHeadingHeight(RowHt() As Single, ColWid() As Single) As Single
Dim StrText As String
Dim StrData As String
Dim I As Integer, j As Integer
Dim PrintX As Single, PrintY As Single
Dim FromPos As Single, ToPos As Single
Dim Wid As Single, NextPos As Single

On Error GoTo Exit_Line
ReDim ColWid(m_DataSource.cols - 1)
ReDim RowHt(m_DataSource.FixedRows - 1)
ReDim Preserve m_ColWidth(m_DataSource.cols - 1)
ReDim Preserve m_RowHeight(m_DataSource.FixedRows - 1)

Dim PrevCol As Integer

With m_DataSource
    ReDim ColWid(.cols - 1)
    'store the coll width of all coloumns
    .Row = .FixedRows
    For j = 0 To .cols - 1
        .Col = j
        ColWid(j) = .CellWidth
        m_ColWidth(j) = .CellWidth
    Next
End With
Dim MergedCols As Integer

For I = 0 To m_DataSource.FixedRows - 1
    'while printing this we have to Print the box
    'so Before Print this heading
    'we have to calculate the height of this row
    RowHt(I) = my_printer.TextHeight("A") + 5
    MergedCols = 0: PrevCol = 1
    m_DataSource.Row = I
    SetFont
    my_printer.Font.Bold = True
    If m_WrapHead Then
        RowHt(I) = GetRowHeight
    End If
    PrintX = m_LeftMargin - FIELD_MARGIN / 2
    m_DataSource.Col = m_SplitCol(m_SplitPageNo - 1)
    StrText = m_DataSource.Text
    FromPos = PrintX
    ToPos = PrintX
    m_RowHeight(I) = RowHt(I)
Next

Exit_Line:
    If Err Then
        'Resume
        Debug.Print Err.Description
    End If
End Function


Private Function GetPrintHeight(StrData As String, Wid As Single) As Single

Dim Ht As Single
Dim Count As Integer
Dim MaxCount As Integer
Dim RetStrArray() As String

Ht = my_printer.TextHeight("A")

On Error GoTo ErrLine
'Now Search for new line charactors
'If It has THEN bifurcate them with string array
'Call GetStringArray(StrData, RetStrArray, vbCrLf)
RetStrArray = Split(StrData, vbCrLf)
Count = 0:
MaxCount = UBound(RetStrArray) + 1
Do
    RetStrArray(Count) = Trim$(RetStrArray(Count))
    If my_printer.TextWidth(RetStrArray(Count)) > Wid Then
        MaxCount = MaxCount + my_printer.TextWidth(RetStrArray(Count)) / Wid + 0.5
        'if we parsed the string in that case one part of the string
        'already included in the array so reduce the max count by one
        MaxCount = MaxCount - 1
    End If
    Count = Count + 1
    If Count > UBound(RetStrArray) Then Exit Do
Loop

Ht = (my_printer.TextHeight("A") + 5) * MaxCount

ErrLine:
    If Err Then
        MsgBox "error in calulationg print height"
        Err.Clear
    End If
    
GetPrintHeight = Ht

End Function
Private Function ZZZ_GetRowHeightColWidth() As Boolean
Dim I As Integer, j As Integer
Dim Ht As Single, NextPos As Single
Dim CurrRow As Long, CurrCol As Long

CurrRow = m_DataSource.Row
CurrCol = m_DataSource.Col

Debug.Print "Calaaaaaaaaaaaaaaaaaa"
On Error GoTo Exit_Line
ReDim Preserve m_ColWidth(m_DataSource.cols - 1)
ReDim Preserve m_RowHeight(m_DataSource.rows - 1)

With m_DataSource
    ReDim ColWid(.cols - 1)
    'store the coll width of all coloumns
    .Row = .FixedRows
    For j = 0 To .cols - 1
        .Col = j
        m_ColWidth(j) = .CellWidth
    Next
End With
SetFont
Ht = my_printer.TextHeight("A") + 5
    

With m_DataSource
    For I = m_DataSource.FixedRows To m_DataSource.rows - 1
        m_RowHeight(I) = Ht
        If m_WrapHead Then
            .Row = I
            m_RowHeight(I) = GetRowHeight
        End If
    Next
End With

Exit_Line:
    If Err Then
        'Resume
        Debug.Print Err.Description
    End If
    m_DataSource.Row = CurrRow
    m_DataSource.Col = CurrCol
End Function


Private Function GetStringsToPrint(strSource As String, RetStrArray() As String, _
        Wid As Single, Optional Alignment As AlignmentConstants = -1) As Long

Dim StrData As String
Dim Count As Integer
Dim MaxCount As Integer
Dim strTemp As String
Dim LoopCount As Integer

StrData = strSource

'Now Search for new line charactors
'If It has THEN bifurcate them with string array
'Call GetStringArray(StrData, RetStrArray, vbCrLf)
RetStrArray = Split(StrData, vbCrLf)
Count = 0:
MaxCount = UBound(RetStrArray)
Do
    RetStrArray(Count) = Trim$(RetStrArray(Count))
    If my_printer.TextWidth(RetStrArray(Count)) > Wid Then
        strTemp = "": LoopCount = 1
        Do
            If RetStrArray(Count) = "" Then Exit Do
            strTemp = Left(RetStrArray(Count), LoopCount)
            If my_printer.TextWidth(strTemp & "A") >= Wid Then
            ReDim Preserve RetStrArray(UBound(RetStrArray) + 1)
            RetStrArray(UBound(RetStrArray)) = strTemp
            RetStrArray(Count) = Mid(RetStrArray(Count), LoopCount + 1)
            'Arrange the string array
            LoopCount = Count
            Do
                If LoopCount = UBound(RetStrArray) Then Exit Do
                RetStrArray(UBound(RetStrArray)) = RetStrArray(LoopCount + 1)
                RetStrArray(LoopCount + 1) = RetStrArray(LoopCount)
                LoopCount = LoopCount + 1
            Loop
            RetStrArray(Count) = strTemp
            MaxCount = UBound(RetStrArray)
            Exit Do
            RetStrArray(Count) = Mid(RetStrArray(Count), LoopCount + 1)
            End If
            LoopCount = LoopCount + 1
        Loop
    End If
    Count = Count + 1
    If Count > UBound(RetStrArray) Then Exit Do
Loop

GetStringsToPrint = UBound(RetStrArray) + 1

'If IsMissing(Alignment) Then Exit Function
If Alignment < 0 Then _
    Alignment = IIf(IsNumeric(strSource), vbRightJustify, vbLeftJustify)

If Alignment = vbLeftJustify Then Exit Function
If Alignment < 0 Then Exit Function
'Now Allign the grids accoring to the specified alligment
MaxCount = UBound(RetStrArray)
Dim blLeft As Boolean
For Count = 0 To MaxCount
  Do
    If my_printer.TextWidth(RetStrArray(Count)) + my_printer.TextWidth("A") >= Wid Then Exit Do
      If Alignment = vbCenter Then
        If blLeft Then RetStrArray(Count) = " " & RetStrArray(Count)
        If Not blLeft Then RetStrArray(Count) = RetStrArray(Count) & " "
        blLeft = Not blLeft
      ElseIf Alignment = vbRightJustify Then
        RetStrArray(Count) = " " & RetStrArray(Count)
      Else
        RetStrArray(Count) = RetStrArray(Count) & " "
      End If
  Loop
    
Next Count

End Function
Private Function AlignData(strSource As String, _
        Wid As Single, Optional dataFormat As Integer) As String
If my_printer.TextWidth(strSource) < Wid Then
    While my_printer.TextWidth(strSource) < Wid
        strSource = " " & strSource
    Wend
End If
AlignData = strSource
End Function
Public Property Get NextPage() As Boolean
    NextPage = m_NextPage
End Property
Public Property Let NextPage(NewBool As Boolean)
    m_NextPage = NewBool
End Property


Public Property Let PageHeight(NewValue As Double)
    m_PageHeight = NewValue
End Property

Public Property Let PageWidth(NewValue As Double)
    m_PageWidth = NewValue
End Property


Private Sub PrintFooter()

If m_clsFooter.Name = "" Then Exit Sub
Dim j As Integer
Dim curRow As Single
Dim StrData As String
Dim X1 As Single
Dim X2 As Single
Dim Y1 As Single
Dim Y2 As Single

RaiseEvent ProcessingMessage("Printing page footer")
X1 = m_LeftMargin * 0.75
X2 = m_PageWidth - m_RightMargin * 0.75
Y1 = m_PageHeight - 2 * my_printer.TextHeight(m_clsFooter.Name)
Y2 = Y1
If m_FooterLIne Then
    'CurY = m_PageHeight - 2 * my_printer.TextHeight(m_clsFooter.Name)
    my_printer.Line (X1, Y1)-(X2, Y2)
End If

m_clsFooter.Font.Name = m_Font.Name
m_clsFooter.Font.Size = 8
m_clsFooter.Font.Bold = False
m_clsFooter.SetAttrib my_printer



'Left position
If m_clsFooter.Align = vbLeftJustify Then X1 = m_LeftMargin
If m_clsFooter.Align = vbRightJustify Then _
        X1 = m_PageWidth - my_printer.TextWidth(m_clsFooter.Name) - m_RightMargin
If m_clsFooter.Align = vbCenter Then
        X1 = (m_PageWidth - my_printer.TextWidth(m_clsFooter.Name)) / 2
End If
'Top Position
Y1 = my_printer.ScaleHeight - 2 * my_printer.TextHeight(m_clsFooter.Name)

'Set the position of cursor
my_printer.CurrentX = X1
my_printer.CurrentY = Y1

my_printer.Print m_clsFooter.Name


End Sub
Public Property Let Font(NewValue As StdFont)
    'lblFont.Font = NewValue
    m_Font = NewValue
End Property


Private Sub PrintHeader()
If Trim$(m_clsHeader.Name) = "" Then Exit Sub

Dim CurX As Single
Dim CurY As Single
Dim X1 As Single
Dim X2 As Single
Dim Y1 As Single
Dim Y2 As Single

RaiseEvent ProcessingMessage("Printing page header")
'Set the cursor to print header
'Left Position
If m_clsHeader.Align = vbLeftJustify Then CurX = m_LeftMargin
If m_clsHeader.Align = vbRightJustify Then _
        CurX = m_PageWidth - my_printer.TextWidth(m_clsHeader.Name) - m_RightMargin
If m_clsHeader.Align = vbCenter Then _
        CurX = (m_PageWidth - my_printer.TextWidth(m_clsHeader.Name)) / 2
'Top position of cursor
CurY = my_printer.TextHeight(m_clsHeader.Name)

m_clsHeader.Font.Name = m_Font.Name
m_clsHeader.Font.Size = 8
m_clsHeader.Font.Bold = False

m_clsHeader.SetAttrib my_printer

my_printer.CurrentX = CurX '- m_LeftMargin * 0.75
my_printer.CurrentY = CurY

my_printer.Print m_clsHeader.Name

X1 = m_LeftMargin * 0.75
X2 = m_PageWidth - m_RightMargin * 0.75
Y1 = my_printer.CurrentY: Y2 = Y1
If m_HeaderLine Then
    my_printer.Line (X1, Y1)-(X2, Y2)
End If

If Y1 > m_TopMargin Then m_TopMargin = CurY

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------
'                                                  Called by 'PrintReport'
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PrintHeading()

Static PageNumPosition As Single

Dim PrintX As Single, PrintY As Single
Dim StrData As String, pagestr As String
Dim SaveRow As Long, saveCol As Long
Dim SameCol As Boolean
Dim NoBOx As Boolean
Static ColWid() As Single
Static RowHt() As Single
Dim StrArr() As String

Dim I As Integer  'To count the rows
Dim j As Integer 'To Counnt the colomns
Dim K As Integer

On Error Resume Next
SaveRow = m_DataSource.Row
saveCol = m_DataSource.Col

'Inititalise the Col width & row height
If UBound(ColWid) < 0 Then
    Call GetHeadingHeight(RowHt, ColWid)
    'The below function Calulates
    'Printing height of all row and stores that value
    'in a array As it is moving to every cell
    'of the grid it take lot of time
    
    'Call GetRowHeightColWidth
End If

On Error GoTo PrintHeading_Error


RaiseEvent ProcessingMessage("Printing page heading")

' Set the top position (Y-coordinate) for printing.)
PrintY = m_Title.Height

PrintX = m_LeftMargin - FIELD_MARGIN / 2
PrintY = my_printer.CurrentY

SetFont

' Save the current grid row and move to top row.
m_DataSource.Row = 0
m_DataSource.Col = 0

' Loop through collection of fields...
Dim StrText As String
Dim FromPos As Single
Dim ToPos As Single
Dim NextPos As Single
Dim Wid As Single

Dim MergedCols As Integer
Dim NoBottom As Boolean
Dim NoTop As Boolean
Dim PrevCol As Integer

'Print the  Heading Rows
For I = 0 To m_DataSource.FixedRows - 1
    m_DataSource.Row = I
    MergedCols = 0: PrevCol = 1
    PrintX = m_LeftMargin - FIELD_MARGIN / 2
    m_DataSource.Col = m_SplitCol(m_SplitPageNo - 1)
    StrText = m_DataSource.Text
    FromPos = PrintX
    ToPos = PrintX
    SameCol = False
    For j = m_SplitCol(m_SplitPageNo - 1) To m_DataSource.cols - 1
        m_DataSource.Col = j
        'Check if the printx exceeds the width of the printing region.
        If m_PageSplitted And j = m_SplitCol(m_SplitPageNo) Then Exit For
        ' set the attributes of print object for printing.
        SetFont
        'SET THE PRINTING CO-ORDINATES
        my_printer.CurrentX = PrintX
        my_printer.CurrentY = PrintY
        
        'If Grid has merge cells true then the cell width may change
        'So take the cell width after Fixed rows
        With m_DataSource
            If ColWid(j) <= my_printer.TextWidth("A") Then PrevCol = PrevCol + 1: GoTo NextCol
            StrData = .Text
            If StrText <> StrData Then 'Print in given pisition
                If MergedCols = 1 And I > 0 Then
                    .Row = I - 1: .Col = .Col - PrevCol
                    If StrText = .Text Then NoTop = True
                    .Row = I: .Col = .Col + PrevCol
                End If
                If MergedCols = 1 And I <= .FixedRows - 2 Then
                    .Row = I + 1: .Col = .Col - PrevCol
                    If StrText = .Text Then StrText = "": NoBottom = True
                    .Row = I: .Col = .Col + PrevCol
                End If
                PrevCol = 1
                my_printer.CurrentX = FromPos
                my_printer.Font.Bold = True
                Wid = my_printer.TextWidth(StrText)
                Call GetStringsToPrint(StrText, StrArr(), (ToPos - FromPos), vbCenter)
                If Not m_WrapHead Then ReDim Preserve StrArr(0)
                
                For K = 0 To UBound(StrArr)
                    my_printer.Font.Bold = True
                    my_printer.CurrentX = FromPos + (ToPos - FromPos - my_printer.TextWidth(StrArr(K))) / 2
                    my_printer.Print StrArr(K)
                Next
                Call PrintRectangle(FromPos, PrintY, ToPos - FromPos, RowHt(I), NoTop, NoBottom)
                'my_printer.Line (FromPos, PrintY)-(ToPos, PrintY + RowHt(i)), , B
                NoTop = False: NoBottom = False
                FromPos = ToPos
                StrText = StrData
                MergedCols = 0
            End If
        End With
        MergedCols = MergedCols + 1
        ToPos = ToPos + ColWid(j) + FIELD_MARGIN
NextCol:
    Next
    
    'Here we have to print last coloumn
    my_printer.CurrentX = FromPos
    Wid = my_printer.TextWidth(StrText)
    my_printer.CurrentY = PrintY
    Call GetStringsToPrint(StrText, StrArr(), (ToPos - FromPos), vbCenter)
    If Not m_WrapHead Then ReDim Preserve StrArr(0)
    For K = 0 To UBound(StrArr)
        StrArr(K) = TruncateData(StrArr(K), ToPos - FromPos)
        my_printer.CurrentX = FromPos + (ToPos - FromPos - my_printer.TextWidth(StrArr(K))) / 2
        'my_printer.CurrentX = FromPos
        my_printer.Font.Bold = True
        my_printer.Print StrArr(K)
    Next
    'Now Print the rectangle
    Call PrintRectangle(FromPos, PrintY, ToPos - FromPos, RowHt(I))
    'my_printer.Line (FromPos, PrintY)-(ToPos, PrintY + RowHt(i)), , B
    'Before Move to Next row Increase the Y Position
    'PrintY = my_printer.CurrentY
    PrintY = PrintY + RowHt(I)
Next

' Set the y co-ordinate for printing record details...
my_printer.CurrentY = my_printer.CurrentY + RowHt(I - 1) - my_printer.TextHeight("A") * 2
my_printer.CurrentY = PrintY + RowHt(I - 1) - my_printer.TextHeight("A")

PrintHeading_Error:
    If Err Then
        MsgBox "PrintHeading: " & vbCrLf & Err.Description, vbCritical
        'Resume
        Err.Clear
    End If

' Restore the grid row.
m_DataSource.Row = SaveRow
m_DataSource.Col = saveCol

DoEvents

End Sub


'Called by 'PrintReport'
'This Will Check Whether Coloumns will fit on the
'The Existing Printer Or Not
'If All coloumns are not fit into the Then Remaining colomns will fir into the Next Page
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CheckColoumnsSize() As Boolean

Dim PrintX As Single, PrintY As Single
Dim j As Integer
Dim rectLeft As Single, rectTop As Single
Dim rectRight As Single, rectBottom As Single
Dim StrData As String, pagestr As String
Dim SaveRow As Long, saveCol As Long

On Error GoTo ColWidth_Error
' Print the header rectangle......
'SetFont

' Save the current grid row and move to top row.
SaveRow = m_DataSource.Row
saveCol = m_DataSource.Col
m_DataSource.Row = m_DataSource.FixedRows

' Loop through collection of fields...
m_NoOfSplitPage = 1

ReDim m_SplitCol(0)
m_SplitCol(0) = 0
m_PageSplitted = False

For j = 0 To m_DataSource.cols - 1
RepeatCheck:
    m_DataSource.Col = j
    PrintX = PrintX + m_DataSource.ColWidth(j) / m_DataSource.Width * m_PageWidth '+ FIELD_MARGIN
    'Check if the printx exceeds the width of the printing region.
    If PrintX > m_PageWidth - m_RightMargin Then
        'If this problem occurs at 0th coloumn
        'Then it will not print a single line
        'In such case Come out of this function
        If j = 0 Then Exit Function
        If m_SplitCol(UBound(m_SplitCol)) = j Then Exit Function
        
        'if this coloumn does not fit into this page then
        'restrict the no of coloumns to this page
        'and shift all other coloumns into next page
        m_PageSplitted = True
        ReDim Preserve m_SplitCol(UBound(m_SplitCol) + 1)
        'now assign the coloumn no to a varible
        m_SplitCol(UBound(m_SplitCol)) = j
        m_NoOfSplitPage = m_NoOfSplitPage + 1
        'Check the condition for next page
        PrintX = 0 'm_LeftMargin
        GoTo RepeatCheck
    End If
Next
'Now store the Value of last coloumn
ReDim Preserve m_SplitCol(UBound(m_SplitCol) + 1)
m_SplitCol(UBound(m_SplitCol)) = j

'If m_BottomMargin < my_printer.TextHeight(m_clsFooter.Name) * 1.5 Then m_BottomMargin = my_printer.TextHeight(m_clsFooter.Name) * 1.5
'' Set the y-coordinate for printing record details...
'my_printer.CurrentY = my_printer.CurrentY + 200

CheckColoumnsSize = True
ColWidth_Error:
    If Err Then MsgBox "Col width setting " & vbCrLf & Err.Description, vbCritical
    
    'Resume
' Restore the grid row.
m_DataSource.Row = SaveRow
m_DataSource.Col = saveCol
DoEvents

End Function

Private Function PrintRectangle(XPos As Single, YPos As Single, _
        Width As Single, Height As Single, Optional NoTop As Boolean, Optional NoBottom As Boolean) As Boolean
On Error GoTo Exit_Line
Dim X As Single, Y As Single
Dim xLeft As Single, yTop As Single
Dim xRight As Single, yBottom As Single
X = my_printer.CurrentX
Y = my_printer.CurrentY
xLeft = XPos: yTop = YPos
xRight = XPos + Width: yBottom = YPos + Height
'Now Print the Rectange
'By printing 4 lines individually

'Print Top Horizontal line
If Not NoTop Then my_printer.Line (xLeft, yTop)-(xRight, yTop)
'Print Left verticle line
my_printer.Line (xLeft, yTop)-(xLeft, yBottom)
'Print rigth verticle line
my_printer.Line (xRight, yTop)-(xRight, yBottom)
'Print bottom Horizontal line1
If Not NoBottom Then my_printer.Line (xLeft, yBottom)-(xRight, yBottom)

PrintRectangle = True

Exit_Line:
my_printer.CurrentX = X
my_printer.CurrentY = Y

End Function

Private Function PrintReport() As Boolean
Debug.Print "START  " & Now
Dim SaveRow As Long
Dim j As Integer
Dim StrArr() As String
Dim strRange As String

Dim FileObj As New FileSystemObject
Set m_txtStream = FileObj.CreateTextFile("c:\TestFile", True)


' Setup error handler...
On Error GoTo Err_Line

''''This Is Temp code for testing
'frmPrintDailog.Show vbModal
m_CancelProcess = False

On Error GoTo Err_Line


'Round the Page Height & Page width
'm_PageHeight = m_PageHeight \ 1
'm_PageWidth = m_PageWidth \ 1

With m_txtStream
    .write Chr(27) & Chr(64)
    .write Chr(255)
    .write Chr(27) & "l" & Chr(m_LeftMargin * 10)
    .write Chr(27) & "Q" & Chr((m_PageWidth - m_RightMargin) * 10)
    .write Chr(27) & "N" & Chr(m_BottomMargin * 4) 'Bottom Margin
    
End With

'Check Whetehr All Colouns fits in a page or not
If Not CheckColoumnsSize Then
    MsgBox "unable to set the printer values"
'    m_DataSource.Visible = True
    Exit Function
End If

m_SplitPageNo = 1
' Set the initial row of the grid to 1,
' because we have to skip the 0th row.
If m_DataSource.rows > m_DataSource.FixedRows Then
    If PageNumber = 1 Then m_DataSource.Row = m_DataSource.FixedRows
Else
    MsgBox "Internal error: No data in grid!!!", vbCritical
    GoTo Exit_Line
End If

'Check how many no Of pages will be
Call SetPageNumber
If rows_in_first_page = 0 Then GoTo Exit_Line

ReDim pages_toPrint(0)
ReDim Preserve pages_toPrint(num_pages * m_NoOfSplitPage)
If strRange = "" Then strRange = "0"
If strRange <> "0" Then
    'Call GetStringArray(strRange, StrArr, ",")
    StrArr = Split(strRange, ",")
    PageNumber = Val(StrArr(0))
    If PageNumber > 1 Then title_printed = True
    For j = 0 To UBound(StrArr)
        If j > num_pages * m_NoOfSplitPage Then Exit For
        If Val(StrArr(j)) > num_pages Then Exit For
        ReDim Preserve pages_toPrint(j)
        pages_toPrint(j) = Val(StrArr(j))
    Next
Else
    ReDim Preserve pages_toPrint(0)
    pages_toPrint(0) = PageNumber
End If

m_ProcCount = 1

' Loop through available pages.
page_printing = 1: j = 0
Do
    page_printing = j + 1
    If Print_EOF Then Exit Do
    ' Store the current row, in case of printing the same page.
    SaveRow = m_DataSource.Row
    
    'IF YOU ARE PRINTING TO THE PRINTER THEN
    'CHECK FOR THE NEXT PAGE TO BE PRINT
    'If TypeOf my_printer Is Printer Then
    If m_View = "PRINTER" Then
        If j > UBound(pages_toPrint) Then Exit Do
        m_DataSource.Row = m_PageRow(pages_toPrint(j) - 1) ' - 1
    End If
    ' Print a page
    If Not PrintPage Then Exit Do

    ' Increment the page counter
    If m_PageSplitted Then
        If m_SplitPageNo = m_NoOfSplitPage Then
            m_SplitPageNo = 1
            PageNumber = PageNumber + 1: j = j + 1
        Else
            m_SplitPageNo = m_SplitPageNo + 1
        End If
    Else
        PageNumber = PageNumber + 1
        j = j + 1
    End If
'    Debug.Assert PageNumber <> 2
    If m_CancelProcess Then GoTo Exit_Line
Loop
If m_View = "PREVIEW" Then
End If

Exit_Line:
    Screen.MousePointer = vbDefault
'    Exit Property
Err_Line:
    If Err.Number = 482 Then
        j = MsgBox("Printer Error: " & vbCrLf & "Check your printer settings.", vbAbortRetryIgnore)
        If j = vbRetry Then
            
            Resume
        ElseIf j = vbIgnore Then
            Resume Next
        End If
    
    ElseIf Err Then
        MsgBox "PrintReport: " & vbCrLf & Err.Description, vbCritical
        'Resume
    End If
    GoTo Exit_Line

End Function

Private Function GetRowHeight(Optional RowNo As Long = -1) As Single
Dim PrevRow As Long
Dim PrevCol As Long
Dim RowHt As Single, NextPos As Single
Dim ColWid As Single
Dim StrPrintArr() As String

PrevRow = m_DataSource.Row
PrevCol = m_DataSource.Col

On Error GoTo Exit_Line

RowHt = my_printer.TextHeight("A") + 5
If Not m_WrapCell Then GoTo Exit_Line

If RowNo >= 0 Then m_DataSource.Row = RowNo
Dim I As Integer, j As Integer, K As Integer
If Not m_WrapCell Then GoTo Exit_Line
With m_DataSource
    ' Loop through the collection of fields...
    SetFont
    For j = 0 To .cols - 1
        If j = m_SplitCol(m_SplitPageNo) Then Exit For
        .Col = j        ' Set the current cell.
        SetFont     ' Set the font for this field.
        ColWid = .CellWidth
        'If Cell width is smaller than width of letter
        'then do not print that coloumn
        If ColWid <= my_printer.TextWidth("A") Then GoTo NextCol
        
        'cut the printing string length in to multiple part
        'according to the printing length
        'Call GetStringsToPrint(.Text, StrPrintArr, ColWid)
        If my_printer.TextWidth(.Text) > ColWid Then
            NextPos = GetPrintHeight(.Text, ColWid)
            RowHt = IIf(NextPos > RowHt, NextPos, RowHt)
        End If
NextCol:
    Next   'End of for loop
End With

Exit_Line:
    If Err Then
        'MsgBox "Error in RowHeigt", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
    End If
    GetRowHeight = RowHt
m_DataSource.Row = PrevRow
m_DataSource.Col = PrevCol



End Function

' This sub routine will be called, while processing the
' print request to the printer.  Basically, it saves the curpagenumber,
' and the current grid row, to restore it back, later.
Private Sub SaveSettings()
    
    ' Save the current view page.
    curPage = PageNumber
    curSplitPage = m_SplitPageNo
'    saved_view = PrintDlg.picPrint(0).Tag
    
    ' Set the title printed flag to false,
    ' to force the printing of title on the first page.
    'title_printed = False

    ' Set the page_start_row
    saved_row = m_DataSource.Row
    m_DataSource.Row = page_start_row
    
    ' Save the page_start_row
    saved_page_start_row = page_start_row
    
    ' Restore the current printer object.
    Set my_printer = Printer
    
End Sub
' This sub routine will be called, while processing the
' print request to the printer.  Basically, it saves the curpagenumber,
' and the current grid row, to restore it back, later.
' NOTE: This routine should be followed by a call to
'             RestoreSettings, after print processing.
Private Sub RestoreSettings()
    
    ' Restore the current view page.
    PageNumber = curPage
    m_SplitPageNo = curSplitPage
    ' Restore the data source current row
     m_DataSource.Row = saved_row

    ' Reset the title_printed flag
    title_printed = True

  ' Restore the page_start_row
     page_start_row = saved_page_start_row

End Sub


Private Sub SetFont(Optional srcObj As clsField)
On Error GoTo Err_Line
Dim Obj As Object

If Not srcObj Is Nothing Then
    Set Obj = srcObj
Else
    Set Obj = m_DataSource
End If
'
Set Obj = m_DataSource
With my_printer
    .FontName = Obj.CellFontName
    .FontSize = Obj.CellFontSize
    .FontBold = Obj.CellFontBold
    .FontItalic = Obj.CellFontItalic
    .FontUnderline = Obj.CellFontUnderline
    .FontStrikethru = Obj.CellFontStrikeThrough
End With

Exit_Line:
    Exit Sub

Err_Line:
    
    If Err.Number = 380 Then Obj.CellFontName = m_Font.Name: Resume
    If Err Then
        MsgBox "SetFont: " & vbCrLf _
            & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Sub
Private Sub PrintRectangle_old(fldObj As clsField, CurX As Single, CurY As Single)
Dim rectLeft As Single, rectTop As Single
Dim rectRight As Single, rectBottom As Single

With fldObj
    rectLeft = FIELD_MARGIN - .RectMargin
    rectTop = CurY - .RectMargin
    rectRight = rectLeft + .RectWidth(my_printer) + .RectMargin
    rectBottom = rectTop + .RectHeight(my_printer) + .RectMargin
End With
my_printer.Line (rectLeft, rectTop)-(rectRight, rectBottom), , B

' Restore the curx and cury.
my_printer.CurrentX = CurX
my_printer.CurrentY = CurY

End Sub
Public Function PrintPage() As Boolean

' Setup error handler.
On Error GoTo Err_Line
Screen.MousePointer = vbHourglass

' Declare variables...
Dim j As Integer
Dim lcount As Long

Dim SaveRow As Long

'Temp variables ''shashi
Dim BeginRow As Long
Dim EndRow As Long

Dim X1 As Single
Dim X2 As Single
Dim Y1 As Single
Dim Y2 As Single
Dim LinesPrinted As Byte

'Print the Header
PrintHeader

'Now set the Printer X,Y co-ordinates
X1 = m_LeftMargin
Y1 = m_TopMargin
' Print the title.
If Not title_printed Or (PageNumber = 1 And m_SplitPageNo = 1) Then
    PrintTitle
    title_printed = True
ElseIf PageNumber = 1 Then
    'my_printer.CurrentY = m_TitleMargin
    Y1 = m_TitleMargin
End If

' Print the heading.
PrintHeading

' IF the row is zero, force it to 1,
' because we do not want to print the 0th row
If m_DataSource.Row < m_DataSource.FixedRows Then _
                m_DataSource.Row = m_DataSource.FixedRows
m_DataSource.Visible = False
' Starting row for this page.
page_start_row = m_DataSource.Row + 1
' Begin a loop for printing the records...
SaveRow = m_DataSource.Row
BeginRow = SaveRow

RaiseEvent ProcessingMessage("Printing " & PageNumber & "th page")

For lcount = m_DataSource.Row To m_DataSource.rows - 1
    ' Check if page end has reached.
    If LinesPrinted >= m_PageHeight - m_BottomMargin Then
        'cHECK WHETHER AT LEAST ONE ROW HAS PRINTED
        ' IF NOT THEN EXIT PAGEPRINT= FALSE
        If LinesPrinted = 0 Then GoTo Exit_Line
        Exit For
    End If
    ' Print the record details.
    LinesPrinted = LinesPrinted + PrintRow
    If m_CancelProcess Then GoTo Exit_Line
    If m_DataSource.Row <= m_DataSource.rows - 2 Then m_DataSource.Row = m_DataSource.Row + 1
Next
EndRow = m_DataSource.Row

'Now print the first Horizontal line
'and print the verticle lines

X2 = X1
'Print the first line
Dim ColWid As Single
If m_ColLine Then
    'Now print lines after each coloumn
    For lcount = m_SplitCol(m_SplitPageNo - 1) To m_DataSource.cols - 1
        If lcount = m_SplitCol(m_SplitPageNo) Then Exit For
        m_DataSource.Col = lcount
        ColWid = m_DataSource.CellWidth
        If ColWid > my_printer.TextWidth("A") Then
            X1 = X1 + ColWid + FIELD_MARGIN
            X2 = X1
            my_printer.Line (X1, Y1)-(X2, Y2)
        End If
        If m_CancelProcess Then Exit Function
    Next
End If

If InStr(1, m_clsFooter.Name, "PAGE ", vbTextCompare) Then _
        m_clsFooter.Name = "Page " & PageNumber & " Of " & num_pages

'Print the footer
'PrintFooter

If m_PageSplitted And m_SplitPageNo <> m_NoOfSplitPage Then m_DataSource.Row = SaveRow

'
' Check if the last row has been reached.
' IF so, set the eof property to TRUE.
Print_EOF = False
If m_PageSplitted Then
    If m_DataSource.Row >= m_DataSource.rows + 1 And m_SplitPageNo = m_NoOfSplitPage Then Print_EOF = True
    If PageNumber = num_pages And m_SplitPageNo = m_NoOfSplitPage Then Print_EOF = True
Else
    If m_DataSource.Row >= m_DataSource.rows - 1 Then Print_EOF = True
    If PageNumber = num_pages Then Print_EOF = True
End If

Screen.MousePointer = vbDefault
PrintPage = True

Exit_Line:
    Screen.MousePointer = vbDefault
    m_DataSource.Visible = True
    Exit Function

Err_Line:
    
    If Err Then
        MsgBox "PreviewReport: " & Err.Description
        Err.Clear
    End If
    'Resume
    GoTo Exit_Line

End Function

Private Function PrintRow() As Byte

Dim j As Integer
Dim StrData As String
Dim Pos As Integer
Dim ColWid As Single
Dim K As Integer
Dim StrPrintArr() As String
Dim RowHt As Single

Dim TextRow() As String

On Error GoTo printrow_error

If m_CancelProcess Then Exit Function

RaiseEvent ProcessCount(m_ProcCount)
m_ProcCount = m_ProcCount + 1
'RowHt = GetRowHeight
'RowHt = M_RowHeight(m_DataSource.Row)

ReDim TextRow(0)

With m_DataSource
    ' Loop through the collection of fields...
    'SetFont
    For j = m_SplitCol(m_SplitPageNo - 1) To .cols - 1
        If j = m_SplitCol(m_SplitPageNo) Then Exit For
        .Col = j        ' Set the current cell.
        'SetFont     ' Set the font for this field.
        ColWid = .CellWidth
        'If Cell width is smaller than width of letter
        'then do not print that coloumn
        If ColWid <= 30 Then GoTo NextCol
        
        'cut the printing string length in to multple part according to the printing lenth
        Call GetStringsToPrint(.Text, StrPrintArr, ColWid, CInt(.CellAlignment / 3) - 1)
        If Not m_WrapCell Then ReDim Preserve StrPrintArr(0)
        For K = 0 To UBound(StrPrintArr)
            If UBound(TextRow) < K Then
                ReDim Preserve TextRow(K)
                TextRow(K) = String(j - m_SplitCol(m_SplitPageNo - 1) - 1, Chr(9))
            End If
            TextRow(K) = TextRow(K) & StrPrintArr(K) & Chr(9)  'Move the printer by one tab
        Next
        
NextCol:
    Next   'End of for loop
    
    For K = 0 To UBound(TextRow)
        m_txtStream.WriteLine TextRow(K) 'iif(k=ubound(textrow)
    Next
    'm_txtStream.write Chr(13) & Chr(10)
    
    'If m_RowLine Then m_txtStream.WriteLine String(100, "-")

End With
    
PrintRow = K

Exit Function

printrow_error:
    If Err Then
        MsgBox "Printrow: " & vbCrLf & Err.Description, vbCritical '
        'Resume
        Err.Clear
    End If
    
End Function
Private Sub PrintTitle()
RaiseEvent ProcessingMessage("Printing report title")

If m_Title Is Nothing Then
    Set m_Title = New clsField
End If

m_Title.SetAttrib my_printer        ' Set the font, color
' If a rectangle specified, print it.
If m_Title.Rectangle And Trim$(CompanyName) <> "" Then
End If

With m_txtStream
    m_txtStream.write Chr$(27) & Chr$(33) & Chr$(24) ' Print the title
    m_txtStream.write Chr$(27) & Chr$(97) & Chr$(1)
    If Trim$(CompanyName) <> "" Then .WriteLine CompanyName
End With

If m_Title.Rectangle And Trim$(m_Title.Name) <> "" Then
End If

If Trim$(m_Title.Name) <> "" Then m_txtStream.WriteLine m_Title.Name
m_txtStream.write Chr$(27) & Chr$(33) & Chr$(0) ' Print the title

End Sub
Public Property Get ReportTitle() As String
    ReportTitle = m_Title.Name
End Property



Public Property Let ReportTitle(ByVal vNewValue As String)
' Initialize the title object, if not already done.
If m_Title Is Nothing Then Set m_Title = New clsField

m_Title.Name = vNewValue
End Property

Public Sub CancelProcess()
    m_CancelProcess = True
    Screen.MousePointer = vbDefault
End Sub

Public Property Let FontSize(ByVal NewValue As Single)
m_FontSize = NewValue
End Property

Public Property Get FontSize() As Single
FontSize = m_FontSize
End Property

Public Property Get FontName() As String
FontName = m_FontName
End Property

Public Property Let FontName(ByVal NewValue As String)
'Befor Assignin the value Check For the Valid FontName
Dim TmpFontName As StdFont
Dim Retval As Integer, Count As Integer
Retval = Screen.FontCount
For Count = 1 To Retval
    If NewValue = Screen.Fonts(Count) Then GoTo ExitLine
Next
Err.Raise 50003, "Print Class", "Invalid FontName"

Exit Property

ExitLine:
m_FontName = NewValue
End Property

Public Property Let MarginBottom(NewValue As Single)
    m_BottomMargin = NewValue
End Property

Public Property Let MarginLeft(NewValue As Single)
    m_LeftMargin = NewValue
End Property

Public Property Let MarginRight(NewValue As Single)
    m_RightMargin = NewValue
End Property
Public Property Let MarginTop(NewValue As Single)
    m_TopMargin = NewValue
End Property

Private Sub SetPageNumber()

Dim RowCount As Long
Dim ColCount As Integer
Dim RowHt As Single
Dim L As Long
Dim X As Single
Dim Y As Double
'Now set the width & height * Picture box
With my_printer
    Y = m_TopMargin
    m_DataSource.Row = 0: m_DataSource.Col = 0
    RowCount = (m_PageHeight - m_BottomMargin) * 4
    L = RowCount
    If L = m_DataSource.FixedRows Then Exit Sub
    rows_in_first_page = L - m_DataSource.FixedRows
    
    'No Of rows Per page
    num_pages = CInt(m_DataSource.rows / rows_in_first_page)
    ReDim m_PageRow(num_pages)
    m_PageRow(0) = m_DataSource.FixedRows
    m_PageRow(1) = RowCount '+ 1
    If num_pages < 2 Then Exit Sub
    
    'If Number of pages is more than one then
    Y = m_TopMargin
    
    For L = RowCount To m_DataSource.rows - 1
        If Y >= m_PageHeight - m_BottomMargin Then
            Y = Y - 1
            Exit For
        End If
    Next
    
    'Now calulate number of rows per page
    rows_per_page = L '- RowCount + 1
'Now Calulate the no of pages
    L = CInt((m_DataSource.rows - m_DataSource.FixedRows - rows_in_first_page))
    num_pages = (L / rows_per_page + 0.5) \ 1
    'Now add the first page to the no of pages
    num_pages = num_pages + 1
    
    'Now calculate the no of pages in the last row
    rows_in_lastpage = L Mod rows_per_page
    
    If rows_in_lastpage = 0 Then rows_in_lastpage = rows_per_page
    ReDim Preserve m_PageRow(num_pages)
    Debug.Print "FIRST PAGE " & rows_in_first_page
    For L = 2 To num_pages
        m_PageRow(L) = rows_in_first_page + (rows_per_page * (L - 1)) + 1
    Next
    m_PageRow(L - 1) = m_DataSource.rows
End With

End Sub

Public Sub ShowPrint()
' If no recordset available, exit.
If m_DataSource Is Nothing Then
    MsgBox "No records!  Assign a recordset for printing.", vbOKOnly + vbExclamation
    Exit Sub
End If


PrintReport

End Sub

Private Function TruncateData(srcString As String, fldwidth As Single) As String
On Error Resume Next
If my_printer.TextWidth(srcString) > fldwidth Then
    While my_printer.TextWidth(srcString) > fldwidth And srcString <> ""
        srcString = Left$(srcString, Len(srcString) - 1)
    Wend
    srcString = Left$(srcString, Len(srcString) - 3) & "..."
End If
TruncateData = srcString
End Function

Private Sub Class_Initialize()

' Set the destination print object.

' Create a title object.
Set m_Title = New clsField
Set m_clsFooter = New clsField
Set m_clsHeader = New clsField

m_Title.Font.Name = "MS Sans Serif"
Me.PageNumber = 1

m_BottomMargin = 300
m_TopMargin = 800
m_LeftMargin = 600
m_RightMargin = 400


End Sub
Public Property Set DataSource(ByVal GridObject As Object)
 ' This old method assigning a recordset has been
 ' deprecated.
' Set m_DataSource = rs
Set m_DataSource = GridObject '.Object

End Property

