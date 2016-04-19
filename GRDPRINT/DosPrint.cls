VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DosPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Define member variables of this print class...
'
Private m_InitString As String
Priv



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

Private WithEvents PrintDlg As frmPrint     ' Preview display object.
Attribute PrintDlg.VB_VarHelpID = -1

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
ReDim ColWid(m_DataSource.Cols - 1)
ReDim RowHt(m_DataSource.FixedRows - 1)
ReDim Preserve m_ColWidth(m_DataSource.Cols - 1)
ReDim Preserve m_RowHeight(m_DataSource.FixedRows - 1)

Dim PrevCol As Integer

With m_DataSource
    ReDim ColWid(.Cols - 1)
    'store the coll width of all coloumns
    .Row = .FixedRows
    For j = 0 To .Cols - 1
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
    'SetFont
    my_printer.Font.Bold = True
    If m_WrapHead Then RowHt(I) = GetRowHeight
    
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
Call GetStringArray(StrData, RetStrArray, vbCrLf)
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
ReDim Preserve m_ColWidth(m_DataSource.Cols - 1)
ReDim Preserve m_RowHeight(m_DataSource.Rows - 1)

With m_DataSource
    ReDim ColWid(.Cols - 1)
    'store the coll width of all coloumns
    .Row = .FixedRows
    For j = 0 To .Cols - 1
        .Col = j
        m_ColWidth(j) = .CellWidth
    Next
End With
SetFont
Ht = my_printer.TextHeight("A") + 5
    

With m_DataSource
    For I = m_DataSource.FixedRows To m_DataSource.Rows - 1
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


Private Function GetStringsToPrint(strSource As String, RetStrArray() As String, Wid As Single, Optional Alignment As AlignmentConstants = -1) As Long
Dim StrData As String
Dim Count As Integer
Dim MaxCount As Integer
Dim strTemp As String
Dim LoopCount As Integer

StrData = strSource

'Now Search for new line charactors
'If It has THEN bifurcate them with string array
Call GetStringArray(StrData, RetStrArray, vbCrLf)
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
Private Function AlignData(strSource As String, Wid As Single, Optional dataFormat As Integer) As String

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
    'Set the Page Heith trough code
    m_InitString =writeline.
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

m_clsFooter.Font.Name = gFont.Name
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

m_clsHeader.Font.Name = gFont.Name
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

If TypeOf my_printer Is PictureBox Then
    PrintDlg.Init
    PrintDlg.picPrint(0).Width = m_PageWidth
    PrintDlg.picPrint(0).Height = m_PageHeight
End If

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
    For j = m_SplitCol(m_SplitPageNo - 1) To m_DataSource.Cols - 1
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


'       Called by 'PrintReport'
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

'printX = FIELD_MARGIN
PrintX = m_LeftMargin

' Print the header rectangle......
SetFont

' Save the current grid row and move to top row.
SaveRow = m_DataSource.Row
saveCol = m_DataSource.Col
m_DataSource.Row = m_DataSource.FixedRows

' Loop through collection of fields...
m_NoOfSplitPage = 1

ReDim m_SplitCol(0)
m_SplitCol(0) = 0
m_PageSplitted = False

For j = 0 To m_DataSource.Cols - 1
RepeatCheck:
    m_DataSource.Col = j
    ' set the attributes of print object for printing.
'    SetFont
    With my_printer
        .CurrentX = PrintX
        PrintX = PrintX + m_DataSource.CellWidth + FIELD_MARGIN
        ' Check if the printx exceeds the width of the printing region.
        'If printX > my_printer.ScaleWidth - FIELD_MARGIN Then
        If PrintX > m_PageWidth - m_RightMargin Then
            'If this problem occurs at 0th coloumn
            'Then it will not print a single line
            'In such case Come out of this
            If j = 0 Then
                Exit Function
            End If
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
            PrintX = m_LeftMargin
            GoTo RepeatCheck
        End If
    End With
Next
'Now store the Value of last coloumn
ReDim Preserve m_SplitCol(UBound(m_SplitCol) + 1)
m_SplitCol(UBound(m_SplitCol)) = j

If m_BottomMargin < my_printer.TextHeight(m_clsFooter.Name) * 1.5 Then m_BottomMargin = my_printer.TextHeight(m_clsFooter.Name) * 1.5
' Set the y-coordinate for printing record details...
my_printer.CurrentY = my_printer.CurrentY + 200

CheckColoumnsSize = True
ColWidth_Error:
    If Err Then
        MsgBox "Col width setting " & vbCrLf & Err.Description, vbCritical
    End If
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

' Setup error handler...
On Error GoTo Err_Line

''''This Is Temp code for testing
'frmPrintDailog.Show vbModal
With frmPrintDailog
    If UCase(.view) = "CANCEL" Then Exit Function
    ReportDestination = UCase(.view)
    If .ChkExcel.value = vbChecked And ReportDestination = "PRINTER" Then
        Call PrintToExcel(m_DataSource, m_Title.Name)
        Exit Function
    End If
    m_RightMargin = .MarginRight * 1440
    m_LeftMargin = .MarginLeft * 1440
    m_TopMargin = .MarginTop * 1440
    m_BottomMargin = .MarginBottom * 1440
    m_RowLine = .HorizontalLine
    m_ColLine = .VerticleLine
'Word wrapping details
    m_WrapCell = .chkWrapcell
    m_WrapHead = .chkWrapHead
    
''Header Details
    m_clsHeader.Name = .txtHeader
    m_HeaderLine = .HeaderLine
    m_clsHeader.Align = .cmbHeaderAlign.ItemData(.cmbHeaderAlign.ListIndex)

''Footer details
    m_clsFooter.Name = .txtFooter
    m_clsFooter.Align = .cmbFooterAlign.ItemData(.cmbFooterAlign.ListIndex)
    m_FooterLIne = .FooterLine

'Set Page width & Height
    PageWidth = .PaperWidth * 1440
    PageHeight = .PaperHeight * 1440
    On Error Resume Next
    If .Orientation = vbPRORLandscape Then  'THen change the width & Height
        m_PageHeight = .PaperWidth * 1440
        m_PageWidth = .PaperHeight * 1440
    End If
    If TypeOf my_printer Is Printer Then
        my_printer.EndDoc
        my_printer.Orientation = .Orientation
        my_printer.PaperSize = .cmbPaper.ItemData(.cmbPaper.ListIndex)
        If my_printer.PaperSize = vbPRPSUser Then
            my_printer.Height = m_PageHeight
            my_printer.Width = m_PageWidth
        End If
    End If

    'Now check the pause option
    pause_between_pages = .chkPause.value
End With
m_CancelProcess = False

On Error GoTo Err_Line

'Get the pages to print
strRange = frmPrintDailog.PageRange

'Round the Page Height & Page width
m_PageHeight = m_PageHeight \ 1
m_PageWidth = m_PageWidth \ 1
''''''''''''*****
If PrintDlg.picPrint.Count = 0 Then Load PrintDlg.picPrint(0)

With PrintDlg.picPrint(0)
    .Left = (PrintDlg.picViewport.Width - .Width) / 2
    .Top = (PrintDlg.picViewport.Height - .Height) / 2
    If .Left < 0 Then .Left = 30
    If .Top < 0 Then .Top = 30
End With
' Hide the grid, if visible.
'm_DataSource.Visible = False

' Initialize the print form, if type of print is "PREVIEW"
If TypeOf my_printer Is PictureBox Then
'    PrintDlg.Init
    PrintDlg.picPrint(0).Width = m_PageWidth
    PrintDlg.picPrint(0).Height = m_PageHeight
End If

'Check Whetehr All Colouns fits in a page or not
If Not CheckColoumnsSize Then
    MsgBox "unable to set the printer values"
'    m_DataSource.Visible = True
    Exit Function
End If

'If Page has ben splitted then
'To Print as page to the PRINTER
' First print the page to the Picture box
'then  send it to printer as image
If TypeOf my_printer Is Printer And m_PageSplitted Then
    Set my_printer = Nothing
    'Now set To Picture image
    If PrintDlg Is Nothing Then Set PrintDlg = New frmPrint
    Set my_printer = PrintDlg.picPrint(0)
    With PrintDlg.picPrint(0)
        .Left = (PrintDlg.picViewport.Width - .Width) / 2
        .Top = (PrintDlg.picViewport.Height - .Height) / 2
        If .Left < 0 Then .Left = 30
        If .Top < 0 Then .Top = 30
    End With
End If

m_SplitPageNo = 1
' Set the initial row of the grid to 1,
' because we have to skip the 0th row.
If m_DataSource.Rows > m_DataSource.FixedRows Then
    If PageNumber = 1 Then m_DataSource.Row = m_DataSource.FixedRows
Else
    MsgBox "Internal error: No data in grid!!!", vbCritical
    GoTo Exit_Line
End If

'Check how many no Of pages will be
'If ReportDestination = "PRINTER" Then Call SetPageNumber
Call SetPageNumber
If rows_in_first_page = 0 Then GoTo Exit_Line

ReDim pages_toPrint(0)
ReDim Preserve pages_toPrint(num_pages * m_NoOfSplitPage)
If strRange <> "0" Then
    Call GetStringArray(strRange, StrArr, ",")
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
If m_View = "PREVIEW" Then
    RaiseEvent MaxProcessCount(m_DataSource.Rows)
Else
    RaiseEvent MaxProcessCount(rows_per_page * (UBound(pages_toPrint) + 1))
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
    
    ' if print type is "PREVIEW", display the view window.
    If m_View = "PREVIEW" Then
        If PrintDlg Is Nothing Then Set PrintDlg = New frmPrint
    Else        ' Printing to printer.
        If pause_between_pages Then
            If MsgBox("Print next page ?", vbQuestion + vbYesNo) = vbNo Then Exit Do
        End If
    End If

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
    With PrintDlg
        'Now assign the font name & size
        strRange = m_DataSource.Font.Name
        For j = 0 To .cmbFontName.ListCount - 1
            .cmbFontName.ListIndex = j
            If strRange = .cmbFontName.Text Then Exit For
        Next
        .cmbFontName.Tag = .cmbFontName.Text
        'Now assign the font size
        strRange = m_DataSource.Font.Size
        For j = 0 To .cmbFontSize.ListCount - 1
            .cmbFontSize.ListIndex = j
            If strRange = .cmbFontSize.Text Then Exit For
        Next
        
        If j = .cmbFontSize.ListCount Then .cmbFontSize.Text = strRange
        PageNumber = 1
        .cmbFontSize.Tag = .cmbFontSize.Text
        .picPrint(0).Tag = 1
        .picPrint(1).Left = .picPrint(0).Left
        .picPrint(1).Top = .picPrint(0).Top
        .picPrint(1).Visible = True
        .picPrint(1).ZOrder 0
        If .HScroll1.Visible Then .HScroll1.ZOrder 0
        If .VScroll1.Visible Then .VScroll1.ZOrder 0
        .txtPageCount = curPage & "/" & num_pages
        'PageNumber = 1
        .Show vbModal
        If .Status = wis_CANCEL Then GoTo Exit_Line
    End With
End If
Exit_Line:
    Screen.MousePointer = vbDefault
    Exit Function
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
    For j = 0 To .Cols - 1
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
        MsgBox "Error in RowHeigt", vbInformation, wis_MESSAGE_TITLE
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
    saved_view = PrintDlg.picPrint(0).Tag
    
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
    PrintDlg.picPrint(0).Tag = saved_view
    PrintDlg.picPrint(CInt(saved_view)).ZOrder 0
    
    ' Restore the data source current row
     m_DataSource.Row = saved_row

    ' Reset the title_printed flag
    title_printed = True

  ' Restore the page_start_row
     page_start_row = saved_page_start_row

    ' Save the current printer
    Set my_printer = PrintDlg.picPrint(0)
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
    
    If Err.Number = 380 Then Obj.CellFontName = gFont.Name: Resume
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

PrintDlg.picPrint(0).Visible = False

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


' In case of "PREVIEW", clear the screen.
If TypeOf my_printer Is PictureBox Then
    my_printer.Cls
End If

With my_printer
    'Print the Header
    PrintHeader
    
    'Now set the Printer X,Y co-ordinates
    my_printer.CurrentX = m_LeftMargin
    my_printer.CurrentY = m_TopMargin
    ' Print the title.
    If Not title_printed Or (PageNumber = 1 And m_SplitPageNo = 1) Then
        PrintTitle
        title_printed = True
    ElseIf PageNumber = 1 Then
        my_printer.CurrentY = m_TitleMargin
    End If
    
    ' Print the heading.
     PrintHeading
    X1 = my_printer.CurrentX
    Y1 = my_printer.CurrentY
    
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
    
    For lcount = m_DataSource.Row To m_DataSource.Rows - 1
        ' Check if page end has reached.
        If .CurrentY + COL_MARGIN + my_printer.TextHeight("A") >= m_PageHeight - m_BottomMargin Then
            'cHECK WHETHER AT LEAST ONE ROW HAS PRINTED
            ' IF NOT THEN EXIT PAGEPRINT= FALSE
            If lcount = SaveRow Then GoTo Exit_Line
            'the Current row Already printed Print from Next Line
            If m_DataSource.Row <= m_DataSource.Rows - 2 Then
                'm_DataSource.Row = m_DataSource.Row + 1
            End If
            Exit For
        End If
        ' Print the record details.
        PrintRow
        If m_CancelProcess Then GoTo Exit_Line
        If m_DataSource.Row <= m_DataSource.Rows - 2 Then m_DataSource.Row = m_DataSource.Row + 1
    Next
    EndRow = m_DataSource.Row
    
    'If it has not printed the all coloumns in this page
    ' Them set the Datasource row value to the prevValue
    Y2 = my_printer.CurrentY
    X2 = my_printer.CurrentX
End With

'Now print the first Horizontal line
'and print the verticle lines
X1 = m_LeftMargin - FIELD_MARGIN / 2
If m_RowLine Then my_printer.Line (X1, Y1)-(X2, Y1)

X2 = X1
'Print the first line
Dim ColWid As Single
If m_ColLine Then
    my_printer.Line (X1, Y1)-(X2, Y2)
    'Now print lines after each coloumn
    For lcount = m_SplitCol(m_SplitPageNo - 1) To m_DataSource.Cols - 1
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
PrintFooter

' Update pagecount info...
If TypeOf my_printer Is PictureBox Then
    PrintDlg.txtPageCount.Text = PageNumber & "/" & num_pages
End If

If m_PageSplitted And m_SplitPageNo <> m_NoOfSplitPage Then m_DataSource.Row = SaveRow

'
' Check if the last row has been reached.
' IF so, set the eof property to TRUE.
Print_EOF = False
If m_PageSplitted Then
    If m_DataSource.Row >= m_DataSource.Rows + 1 And m_SplitPageNo = m_NoOfSplitPage Then Print_EOF = True
    If PageNumber = num_pages And m_SplitPageNo = m_NoOfSplitPage Then Print_EOF = True
Else
    If m_DataSource.Row >= m_DataSource.Rows - 1 Then Print_EOF = True
    If PageNumber = num_pages Then Print_EOF = True
End If

Screen.MousePointer = vbDefault
If m_View = "PRINTER" Then
    If TypeOf my_printer Is PictureBox Then
        Printer.PaintPicture my_printer.Image, 0, 0
        my_printer.Cls
    End If
    Printer.NewPage
ElseIf TypeOf my_printer Is PictureBox Then
  With PrintDlg
    Load .picPrint(.picPrint.Count)
    .picPrint(.picPrint.Count - 1).PaintPicture .picPrint(0).Image, 0, 0
    .picPrint(.picPrint.Count - 1).Left = .picPrint(0).Left
    .picPrint(.picPrint.Count - 1).Top = .picPrint(0).Top
    .picPrint(.picPrint.Count - 1).Width = .picPrint(0).Width
    .picPrint(.picPrint.Count - 1).Height = .picPrint(0).Height
    If UBound(m_PageRow) < .picPrint.Count - 1 Then ReDim Preserve m_PageRow(UBound(m_PageRow) + 1)
    m_PageRow(.picPrint.Count - 1) = m_DataSource.Row
  End With
End If
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

Private Sub PrintRow()
Dim j As Integer
Dim curRow As Single
Dim CurX As Single
Dim StrData As String
Dim Pos As Integer
Dim ColWid As Single
Dim K As Integer
Dim StrPrintArr() As String
Dim NextYPos As Single
Dim RowHt As Single

On Error GoTo printrow_error

If m_CancelProcess Then Exit Sub
' Save the current row.
curRow = my_printer.CurrentY
'CurX = FIELD_MARGIN
CurX = m_LeftMargin

RaiseEvent ProcessCount(m_ProcCount)
m_ProcCount = m_ProcCount + 1
RowHt = GetRowHeight
'RowHt = M_RowHeight(m_DataSource.Row)
NextYPos = my_printer.CurrentY + RowHt

With m_DataSource
    ' Loop through the collection of fields...
    SetFont
    For j = m_SplitCol(m_SplitPageNo - 1) To .Cols - 1
        If j = m_SplitCol(m_SplitPageNo) Then Exit For
        .Col = j        ' Set the current cell.
        SetFont     ' Set the font for this field.
        ColWid = .CellWidth
        'If Cell width is smaller than width of letter
        'then do not print that coloumn
        If ColWid <= my_printer.TextWidth("A") Then GoTo NextCol
        'cut the printing string length in to multple part to according to the printing lenth
        Call GetStringsToPrint(.Text, StrPrintArr, ColWid, CInt(.CellAlignment / 3) - 1)
        If Not m_WrapCell Then ReDim Preserve StrPrintArr(0)
        my_printer.CurrentY = curRow + COL_MARGIN
        For K = 0 To UBound(StrPrintArr)
            my_printer.CurrentX = CurX - 10
            my_printer.Print StrPrintArr(K)
            'If my_printer.CurrentY > NextYPos Then NextYPos = my_printer.CurrentY
        Next
        CurX = CurX + .CellWidth + FIELD_MARGIN
NextCol:
    Next   'End of for loop
    my_printer.CurrentX = my_printer.CurrentX
    NextYPos = NextYPos + COL_MARGIN
    If m_RowLine Then my_printer.Line (m_LeftMargin - FIELD_MARGIN / 2, NextYPos)-(CurX - FIELD_MARGIN / 2, NextYPos)
End With
    'Set the Y postions of to print next line
    my_printer.CurrentY = NextYPos
    
Exit Sub

printrow_error:
    If Err Then
        MsgBox "Printrow: " & vbCrLf & Err.Description, vbCritical '
        'Resume
        Err.Clear
    End If
    
End Sub
Private Sub PrintTitle()
RaiseEvent ProcessingMessage("Printing report title")

If m_Title Is Nothing Then
    Set m_Title = New clsField
End If

m_Title.SetAttrib my_printer        ' Set the font, color
With my_printer
        .CurrentY = m_TopMargin
        .CurrentX = (m_PageWidth - .TextWidth(CompanyName)) / 2
        ' If a rectangle specified, print it.
        If m_Title.Rectangle And Trim$(CompanyName) <> "" Then
            PrintRectangle my_printer.CurrentX, my_printer.CurrentY, _
                    my_printer.TextWidth(m_Title.Name), my_printer.TextHeight(m_Title.Name)
        End If
End With
' Print the title
If Trim$(CompanyName) <> "" Then my_printer.Print CompanyName

my_printer.CurrentX = (m_PageWidth - my_printer.TextWidth(m_Title.Name)) / 2
If m_Title.Rectangle And Trim$(m_Title.Name) <> "" Then
    PrintRectangle my_printer.CurrentX, my_printer.CurrentY, _
            my_printer.TextWidth(CompanyName), my_printer.TextHeight(CompanyName)
End If

If Trim$(m_Title.Name) <> "" Then my_printer.Print m_Title.Name

' Set the attributes for printing the company name.
my_printer.FontSize = 12
my_printer.CurrentX = (m_PageWidth - my_printer.TextWidth(CompanyName)) / 2
my_printer.FontUnderline = False

If Trim$(CompanyName) <> "" Or Trim$(m_Title.Name) <> "" Then
    m_heading_top = my_printer.CurrentY + 500
    my_printer.CurrentY = my_printer.CurrentY + 250
End If

m_TitleMargin = my_printer.CurrentY

End Sub
Public Property Get ReportTitle() As String
    ReportTitle = m_Title.Name
End Property



Public Property Let ReportTitle(ByVal vNewValue As String)

' Initialize the title object, if not already done.
If m_Title Is Nothing Then
    Set m_Title = New clsField
End If
m_Title.Name = vNewValue
End Property

Public Sub CancelProcess()
    m_CancelProcess = True
    Screen.MousePointer = vbDefault
End Sub
Public Property Get ReportDestination() As String

ReportDestination = m_View
Exit Property
If TypeOf my_printer Is Printer Then
    ReportDestination = "PRINTER"
ElseIf TypeOf my_printer Is PictureBox Then
    ReportDestination = "PREVIEW"
End If
End Property

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

Public Property Let ReportDestination(ByVal vNewValue As String)
If vNewValue = "PRINTER" Then
    Set my_printer = Printer
ElseIf vNewValue = "PREVIEW" Then
    If PrintDlg Is Nothing Then
        Set PrintDlg = New frmPrint
    End If
    Set my_printer = PrintDlg.picPrint(0)
Else
    MsgBox "Invalid Report destination - " & vNewValue, vbExclamation
    Exit Property
End If
m_View = vNewValue

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
    m_Title.SetAttrib my_printer
    Y = Y + .TextHeight(CompanyName)
    Y = Y + .TextHeight(m_Title.Name)
    m_DataSource.Row = 0: m_DataSource.Col = 0
    If ReportDate Then Y = Y + .TextHeight("A")
    Y = Y + 250 'after printing title give a gap
    my_printer.FontSize = 10
    SetFont
    
    Y = Y + .TextHeight("A")  ''Space to print Heading
    Y = Y + 200 'Space after printing Heading
'First calulate the No Of Rows in the first page
    If PageNums Then Y = Y + .TextHeight("A")
    
    For L = m_DataSource.FixedRows To m_DataSource.Rows - 1
        Y = Y + .TextHeight("A") + COL_MARGIN - 5
        If Y >= m_PageHeight - m_BottomMargin Then
            Exit For
        End If
    Next
    RowCount = L
    If L = m_DataSource.FixedRows Then Exit Sub
    rows_in_first_page = L - m_DataSource.FixedRows
    
    'No Of rows Per page
    num_pages = CInt(m_DataSource.Rows / rows_in_first_page)
    ReDim m_PageRow(num_pages)
    m_PageRow(0) = m_DataSource.FixedRows
    m_PageRow(1) = RowCount '+ 1
    If num_pages < 2 Then Exit Sub
    
    'If Number of pages is more than one then
    Y = m_TopMargin
    
    SetFont
    
    Y = Y + .TextHeight("A")  ''Space to print Heading
    Y = Y + 200 'Space after printing Heading
'First calulate the No Of Rows in the first page
    If PageNums Then Y = Y + .TextHeight("A")
    
    For L = RowCount To m_DataSource.Rows - 1
        Y = Y + .TextHeight("A") + COL_MARGIN - 5
        If Y >= m_PageHeight - m_BottomMargin Then
            Y = Y - 1
            Exit For
        End If
    Next
    'Now calulate number of rows per page
    rows_per_page = L - RowCount + 1
'Now Calulate the no of pages
    L = CInt((m_DataSource.Rows - m_DataSource.FixedRows - rows_in_first_page))
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
        'Debug.Print L & "  PAGE " & m_PageRow(L)
    Next
    m_PageRow(L - 1) = m_DataSource.Rows
End With

End Sub

Public Sub ShowPrint()
' If no recordset available, exit.
If m_DataSource Is Nothing Then
    MsgBox "No records!  Assign a recordset for printing.", vbOKOnly + vbExclamation
    Exit Sub
End If

frmPrintDailog.Show vbModal
With frmPrintDailog
    If UCase(.view) = "CANCEL" Then Exit Sub
    ReportDestination = UCase(.view)
    If .ChkExcel.value = vbChecked And ReportDestination = "PRINTER" Then
        Call PrintToExcel(m_DataSource, m_Title.Name)
        Exit Sub
    End If
    m_RightMargin = .MarginRight * 1440
    m_LeftMargin = .MarginLeft * 1440
    m_TopMargin = .MarginTop * 1440
    m_BottomMargin = .MarginBottom * 1440
    m_RowLine = .HorizontalLine
    m_ColLine = .VerticleLine
'Word wrapping details
    m_WrapCell = .chkWrapcell
    m_WrapHead = .chkWrapHead
    
''Header Details
    m_clsHeader.Name = .txtHeader
    m_HeaderLine = .HeaderLine
    m_clsHeader.Align = .cmbHeaderAlign.ItemData(.cmbHeaderAlign.ListIndex)

''Footer details
    m_clsFooter.Name = .txtFooter
    m_clsFooter.Align = .cmbFooterAlign.ItemData(.cmbFooterAlign.ListIndex)
    m_FooterLIne = .FooterLine

'Set Page width & Height
    PageWidth = .PaperWidth * 1440
    PageHeight = .PaperHeight * 1440
    On Error Resume Next
    If .Orientation = vbPRORLandscape Then  'THen change the width & Height
        m_PageHeight = .PaperWidth * 1440
        m_PageWidth = .PaperHeight * 1440
    End If
    If TypeOf my_printer Is Printer Then
        my_printer.EndDoc
        my_printer.Orientation = .Orientation
        my_printer.PaperSize = .cmbPaper.ItemData(.cmbPaper.ListIndex)
        If my_printer.PaperSize = vbPRPSUser Then
            my_printer.Height = m_PageHeight
            my_printer.Width = m_PageWidth
        End If
    End If
End With

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
Set PrintDlg = New frmPrint
Set my_printer = PrintDlg.picPrint

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

frmPrintDailog.optCurrPage.Enabled = False
frmPrintDailog.cmdPreview.Enabled = True

End Sub
Public Property Set DataSource(ByVal GridObject As Object)
 ' This old method assigning a recordset has been
 ' deprecated.
' Set m_DataSource = rs
Set m_DataSource = GridObject '.Object

End Property

Private Sub Class_Terminate()
If Not PrintDlg Is Nothing Then
    Unload PrintDlg
End If
Unload frmPrintDailog
Set PrintDlg = Nothing
End Sub
Private Sub PrintDlg_ProcessEvent(eventNo As Integer)

' Declare variables...
Dim curPage As Integer
Dim PageNo As Integer
Dim LeftPos As Single
Dim TopPos As Single
Dim ColWid() As Single
Dim j As Integer
LeftPos = PrintDlg.picPrint(PrintDlg.picPrint(0).Tag).Left
TopPos = PrintDlg.picPrint(PrintDlg.picPrint(0).Tag).Top
'PrintDlg.picPrint(PrintDlg.picPrint(0).Tag).Visible = False
Select Case eventNo
    Case wis_SHOW_PREVIOUS
        ' Increment the page counter
        If m_PageSplitted Then
            If PageNumber = 1 And m_SplitPageNo = 1 Then GoTo ExitLine
            If m_SplitPageNo = 1 Then
                m_SplitPageNo = m_NoOfSplitPage
                m_DataSource.Row = IIf((page_start_row - rows_per_page) >= 0, _
                    (page_start_row - rows_per_page), 0)
                PageNumber = PageNumber - 1
            Else
                m_SplitPageNo = m_SplitPageNo - 1
            End If
            title_printed = False
        Else
            If PageNumber = 1 Then GoTo ExitLine
            m_DataSource.Row = IIf((page_start_row - rows_per_page) >= 0, (page_start_row - rows_per_page), 0)
            title_printed = False
            PageNumber = PageNumber - 1
            If PageNumber = 1 Then title_printed = False
        End If
        With PrintDlg
            .picPrint(PrintDlg.picPrint(0).Tag - 1).Visible = True
            .picPrint(PrintDlg.picPrint(0).Tag - 1).ZOrder 0
            .picPrint(0).Tag = Val(PrintDlg.picPrint(0).Tag) - 1
        End With
    Case wis_SHOW_NEXT
        If PageNumber = num_pages And m_SplitPageNo = m_NoOfSplitPage Then GoTo ExitLine
        If m_PageSplitted Then
            If m_SplitPageNo = m_NoOfSplitPage Then
                m_SplitPageNo = 1
                PageNumber = PageNumber + 1
            Else
                m_SplitPageNo = m_SplitPageNo + 1
            End If
        Else
            PageNumber = PageNumber + 1
        End If
        With PrintDlg
            If .picPrint.Count = Val(.picPrint(0).Tag + 1) Then .picPrint(0).Tag = .picPrint(0).Tag - 1
            .picPrint(.picPrint(0).Tag + 1).Visible = True
            .picPrint(.picPrint(0).Tag + 1).Left = LeftPos '.picPrint(0).Left
            .picPrint(.picPrint(0).Tag + 1).Top = TopPos '.picPrint(0).Top
            .picPrint(.picPrint(0).Tag + 1).ZOrder 0
            .picPrint(0).Tag = Val(.picPrint(0).Tag) + 1
        End With
    Case wis_SHOW_FIRST
        If m_PageSplitted Then
            If PageNumber = 1 And m_SplitPageNo = 1 Then GoTo ExitLine
            m_SplitPageNo = 1: PageNumber = 1
            m_DataSource.Row = m_DataSource.FixedRows
            m_DataSource.Col = 0
        Else
            ' If already at page 1, exit.
            If PageNumber = 1 Then GoTo ExitLine
            m_DataSource.Row = m_DataSource.FixedRows
            PageNumber = 1
        End If
        With PrintDlg
            .picPrint(1).Visible = True
            .picPrint(1).ZOrder 0
            .picPrint(0).Tag = 1
        End With

    Case wis_SHOW_LAST
        m_SplitPageNo = m_NoOfSplitPage
        m_DataSource.Row = m_DataSource.Rows - rows_in_lastpage - 1
        PageNumber = num_pages
        With PrintDlg
            .picPrint(.picPrint.Count - 1).Visible = True
            .picPrint(.picPrint.Count - 1).ZOrder 0
            .picPrint(0).Tag = .picPrint.Count - 1
        End With

    Case wis_PRINT_CURRENT
        ' Save the current settings...
        Call SaveSettings
        my_printer.PaintPicture PrintDlg.picPrint(PrintDlg.picPrint(0).Tag).Image, 0, 0
        'PageNumber = PageNumber - 1        ' What is the necessity ???
        my_printer.NewPage
        Print_EOF = False
        PrintTitle
        PrintPage
        ' Restore the saved settings...
        Call RestoreSettings

    Case wis_PRINT_CURRENT_PAUSE
        pause_between_pages = True
        Call SaveSettings
        Print_EOF = False
        PrintReport
        Call RestoreSettings
    
    Case wis_PRINT_ALL, wis_PRINT_ALL_PAUSE
        If eventNo = wis_PRINT_ALL_PAUSE Then
            pause_between_pages = True
        End If
        Call SaveSettings
        Print_EOF = False
        With frmPrintDailog
            .optCurrPage.Enabled = True
            .cmdPreview.Enabled = False
        
            'Now check the pause option
            pause_between_pages = .chkPause.value
            .Show 1
            .optCurrPage.Enabled = False
            .cmdPreview.Enabled = True
        End With
        PrintReport
        
        Call RestoreSettings
        
    Case wis_Print_Excel
        Call PrintToExcel(m_DataSource, ReportTitle)
'    Case wis_SHOW_PAGE
'        PageNo = Val(PrintDlg.txtPageCount)
'        If PageNo > PrintDlg.picPrint.Count - 1 Then _
'            PageNo = PrintDlg.picPrint.Count - 1
'        PrintDlg.picPrint(0).Tag = PageNo
    
    Case wis_SHOW_REFRESH
        ReDim ColWid(m_DataSource.Cols - 1)
        m_DataSource.Row = m_DataSource.Rows - 1
        For j = 0 To m_DataSource.Cols - 1
            ColWid(j) = m_DataSource.ColWidth(j)
        Next
        m_DataSource.Font.Name = PrintDlg.cmbFontName.Text
        m_DataSource.Font.Size = PrintDlg.cmbFontSize.Text
        For j = 0 To m_DataSource.Cols - 1
            m_DataSource.ColWidth(j) = ColWid(j)
        Next
        Print_EOF = False
        Unload PrintDlg
        Set PrintDlg = Nothing
        Call PrintReport
        Exit Sub
        
End Select

ExitLine:
With PrintDlg
    .picPrint(.picPrint(0).Tag).Visible = True
    .picPrint(.picPrint(0).Tag).ZOrder 0
    If .HScroll1.Visible Then .HScroll1.ZOrder 0
    If .VScroll1.Visible Then .VScroll1.ZOrder 0
    .txtPageCount = .picPrint(0).Tag & " / " & .picPrint.Count - 1
End With


End Sub


