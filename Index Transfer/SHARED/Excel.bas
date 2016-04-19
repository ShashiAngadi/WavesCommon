Attribute VB_Name = "basExcel"
Option Explicit
' Declare necessary API routines:
Declare Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
                    ByVal lpWindowName As Long) As Long

Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long

Public Function WordExists() As Boolean
    Dim strXLPAth As String
    strXLPAth = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE", "Path")
    If Right(strXLPAth, 1) <> "\" Then strXLPAth = strXLPAth & "\"
    strXLPAth = strXLPAth & "WINWORD.EXE"
    If Dir(strXLPAth) Then
        WordExists = True
    End If
End Function

Public Function ExcelExists() As Boolean
    Dim strXLPAth As String
    ExcelExists = False
    strXLPAth = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE", "Path")
    If Right(strXLPAth, 1) <> "\" Then strXLPAth = strXLPAth & "\"
    strXLPAth = strXLPAth & "Excel.exe"
    If Dir(strXLPAth) <> "" Then
        ExcelExists = True
    End If
End Function

Sub GetExcel()
    Dim myxl As Object  ' Variable to hold reference
    ' to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   ' Flag for final release.

' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.
' Getobject function called without the first argument returns a
' reference to an instance of the application. If the application isn't
' running, an error occurs.
    Set myxl = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel
    
    'Dim XlApp As Excel.Application

'Set the object variable to reference the file you want to see.
    Set myxl = GetObject(App.Path & "\Print.XLS")

' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
    myxl.Application.Visible = True
    myxl.Parent.Windows(1).Visible = True
    myxl.Application.Caption = "Waves Information Systems"
    'Dim XlApp As Excel.Application
    On Error GoTo 0
    Dim Arr() As String
    ReDim Arr(2)
    Arr(0) = "12"
    Arr(1) = "23"
    Arr(2) = "34"
    
    
    ' Do manipulations of your
    ' file here.
    ' ...
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
    'MyXL.Row = 1
    'MyXL.Application.Row = 1
    If ExcelWasNotRunning = True Then
        myxl.Application.Quit
    End If

    Set myxl = Nothing  ' Release reference to the
    ' application and spreadsheet.
End Sub




Function DetectExcel() As Boolean
Dim ExcelPath As String
' Procedure dectects a running Excel and registers it.
    Const WM_USER = 1024
    Dim hWnd As Long
' If Excel is running this API call returns its handle.
    hWnd = FindWindow("XLMAIN", 0)
    If hWnd = 0 Then   ' 0 means Excel not running.
'''        Exit Function
    Else
    ' Excel is running so use the SendMessage API
    ' function to enter it in the Running Object Table.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If
    ExcelPath = GetRegistryValue(HKEY_LOCAL_MACHINE, _
        "Software\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe", _
        "Path")
    If ExcelPath = "" Then Exit Function
    If Dir(ExcelPath & "\Excel.exe", vbNormal) = "" Then Exit Function
DetectExcel = True
End Function
Public Sub PrintToExcel(grd As MSFlexGrid, ReportTitle As String)
#If CodeCompleted Then
Dim xlWorkSheet As Object 'Worksheet
Dim xlWorkBook As Object 'Workbook
Dim FileName As String
Dim Count  As Integer
Dim RowCount As Integer
Dim MaxCols As Long

MaxCols = grd.Cols

With wisMain.cdb
    .Filter = "Excel Files(*.xls)|*.xls|All Files(*.*)|*.* "
    .DefaultExt = "*.xls"
    .CancelError = True
    .ShowSave
    FileName = .FileName
End With
If FileName = "" Then Exit Sub
If Dir(FileName, vbNormal) <> "" Then
    If MsgBox("The file " & FileName & " already exists," & _
    " do you want to overwrite it?", vbYesNo + vbDefaultButton2, _
    "Saving the file ...") = vbNo Then Exit Sub
    Kill FileName
End If

Screen.MousePointer = vbHourglass
'Set xlWorkBook = Workbooks.Add
'Set xlWorkSheet = xlWorkBook.Sheets(1)
    
With xlWorkSheet
    .Range(Cells(1, 1), Cells(1, MaxCols)).Select
    With Selection
        .MergeCells = True: .WrapText = True: .value = gCompanyName
        .Font.Bold = True: .Font.Size = 14: .Font.Name = gFontName
        .HorizontalAlignment = xlCenter
    End With
    .Range(Cells(2, 1), Cells(2, MaxCols)).Select
    With Selection
        .MergeCells = True: .WrapText = True: .value = ReportTitle
        .Font.Bold = True: .Font.Size = 14: .Font.Name = gFontName
        .HorizontalAlignment = xlCenter
    End With
    .Range(Cells(1, 1), Cells(2, MaxCols)).Select
    With Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    .Range(Cells(3, 1), Cells(3, MaxCols)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    grd.Row = 0
    For Count = 1 To MaxCols
        grd.Col = Count - 1
        .Cells(3, Count).Font.Name = grd.CellFontName
        .Cells(3, Count).Font.Bold = grd.CellFontBold
        .Cells(3, Count) = grd.Text
        .Cells(3, Count).HorizontalAlignment = xlCenter
    Next
    RowCount = 3
    Do
        If grd.Row = grd.Rows - 1 Then Exit Do
        grd.Row = grd.Row + 1
        RowCount = RowCount + 1
        For Count = 1 To MaxCols
            grd.Col = Count - 1
            If IsNumeric(grd.Text) And Count > 2 Then
                .Cells(RowCount, Count).Font.Name = "Times New Roman"
                .Cells(RowCount, Count).NumberFormat = "0.00"
            Else
                .Cells(RowCount, Count).Font.Name = grd.CellFontName
            End If
            If grd.CellFontBold Then
                .Cells(RowCount, Count).Font.Bold = grd.CellFontBold
            End If
            .Cells(RowCount, Count) = grd.Text
            .Cells(RowCount, Count).Font.Size = 12
            .Cells(RowCount, Count).RowHeight = 15
        Next
    Debug.Assert RowCount <> 200
    Loop
    For Count = 1 To MaxCols
        .Cells(1, Count).Select
        Selection.EntireColumn.AutoFit
    Next
End With

With ActiveSheet.PageSetup
    .PrintTitleRows = ""
    .PrintTitleColumns = ""
End With
ActiveSheet.PageSetup.PrintArea = ""
With ActiveSheet.PageSetup
    .LeftHeader = ""
    .CenterHeader = ""
    .RightHeader = ""
    .LeftFooter = ""
    .CenterFooter = ""
    .RightFooter = ""
    .LeftMargin = Application.InchesToPoints(0.5)
    .RightMargin = Application.InchesToPoints(0#)
    .TopMargin = Application.InchesToPoints(0.5)
    .BottomMargin = Application.InchesToPoints(1#)
    .HeaderMargin = Application.InchesToPoints(0)
    .FooterMargin = Application.InchesToPoints(0)
    .PrintHeadings = False
    .PrintGridlines = True
    .PrintComments = xlPrintNoComments
    .PrintQuality = Array(360, 180)
    .CenterHorizontally = False
    .CenterVertically = False
    .Orientation = xlPortrait
    .Draft = False
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = 100
End With

xlWorkBook.SaveAs FileName
xlWorkBook.Close savechanges:=True

Set xlWorkSheet = Nothing
Set xlWorkBook = Nothing
Screen.MousePointer = vbNormal
MsgBox "Saved the file to " & FileName, vbInformation, _
    "Saving the File to Excel"

#End If
End Sub

