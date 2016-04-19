Attribute VB_Name = "basExcel"
Option Explicit
Public g_NoCellFont As Boolean

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
    'strXLPAth = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE", "Path")
    If Right(strXLPAth, 1) <> "\" Then strXLPAth = strXLPAth & "\"
    strXLPAth = strXLPAth & "WINWORD.EXE"
    If Dir(strXLPAth) Then
        WordExists = True
    End If
End Function

Public Function ExcelExists() As Boolean
    Dim strXLPAth As String
    ExcelExists = False
    'strXLPAth = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE", "Path")
    If Right(strXLPAth, 1) <> "\" Then strXLPAth = strXLPAth & "\"
    strXLPAth = strXLPAth & "Excel.exe"
    If Dir(strXLPAth) <> "" Then
        ExcelExists = True
    End If
    ExcelExists = True
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




Sub DetectExcel()
' Procedure dectects a running Excel and registers it.
    Const WM_USER = 1024
    Dim hWnd As Long
' If Excel is running this API call returns its handle.
    hWnd = FindWindow("XLMAIN", 0)
    If hWnd = 0 Then   ' 0 means Excel not running.
        Exit Sub
    Else
    ' Excel is running so use the SendMessage API
    ' function to enter it in the Running Object Table.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If
    
End Sub
    


Public Sub PrintToExcel(ByVal grd As Object, ReportTitle As String)

Dim XlApp As Object
'Dim XlApp As EXCEPINFO

Dim ExcelIsRunning As Boolean

' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.
' Getobject function called without the first argument returns a
' reference to an instance of the application. If the application isn't
' running, an error occurs.
    ExcelIsRunning = True
    Set XlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelIsRunning = False
    Err.Clear   ' Clear Err object in case error occurred.
    'Set XlApp = Nothing
    
' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel
    
'Set the object variable to reference the file you want to see.
    'While Installing the SoftWare We have to Install the .xls
    'File(wisPrintXl) in hidden mode
    'Copy wisPrintXl file as priint to file
    'ask File Name to Print
    Dim prtFileName As String
    
    FileCopy App.Path & "\wisPrint.xls", prtFileName
    With frmPrintDailog.cdb
        .DefaultExt = "xls"
        .CancelError = True
        .DialogTitle = "Select the file to print"
        .Filter = "Excel files|*.xls|All Files|(*.*)"
        .ShowSave
        prtFileName = .FileName
    End With
    If Err.Number = 75 Then Kill prtFileName: Err.Clear
    
    
    'prtFileName = InputBox("Name of the Excel file", , App.Path & "\Test1.xls")
    If Len(prtFileName) = 0 Then Exit Sub
    If UCase(Right(prtFileName, 4)) <> ".XLS" Then prtFileName = prtFileName & ".xls"
    
    If prtFileName = "" Then Exit Sub
    FileCopy App.Path & "\wisPrint.xls", prtFileName
    
    Set XlApp = GetObject(prtFileName)
    If Err.Number <> 0 Then
        'Excel is Not Opened
        MsgBox "Unable to detect Microsoft Excel", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    With XlApp.Application
        .Visible = True
        XlApp.Windows(1).Visible = True
        .Caption = "Waves Information Systems - Printing Document"
        .Cells(1, 3) = gCompanyName
        .Cells(1, 3).Font.Name = gFont.Name
        .Cells(1, 3).Font.Size = gFont.Size
        .Cells(2, 2) = ReportTitle
        .Cells(2, 2).Font.Name = gFont.Name
        .Cells(2, 2).Font.Size = gFont.Size
    End With
    'First Print the Title of coloumnsm

Dim RowCount As Integer
Dim ColCount As Integer
Dim MaxRow As Integer
Dim MaxCol As Integer
    
With grd
    MaxRow = .Rows - 1
    MaxCol = .Cols - 1
    .Row = 0
    For ColCount = 0 To MaxCol
        .Col = ColCount
        XlApp.Application.Cells(3, ColCount + 1).Font.Name = .CellFontName
        XlApp.Application.Cells(3, ColCount + 1).Font.Bold = .CellFontBold
        XlApp.Application.Cells(3, ColCount + 1) = .Text
    Next
End With

'Now Print the Remaining Text of the grid
ColCount = 0
grd.Row = 0: RowCount = grd.FixedCols
Do
    If RowCount >= MaxRow Then Exit Do
    grd.Row = RowCount
    
    For ColCount = 0 To MaxCol
        DoEvents
        grd.Col = ColCount
        
        If Not g_NoCellFont Then
            With XlApp.Application.Cells(RowCount, ColCount + 1)
                .Font.Name = grd.CellFontName
                'If IsNumeric(grd.Text) And ColCount <> 1 Then
                '    .Font.Name = "Times New Roman"
                '    .NumberFormat = "0.00"
                'End If
                .Font.Bold = grd.CellFontBold
                .Font.Size = grd.CellFontSize
            End With
            DoEvents
        End If
        
        XlApp.Application.Cells(RowCount, ColCount + 1) = grd.Text
        
    Next
    
    DoEvents
    RowCount = RowCount + 1
Loop

ExitLine:
Set XlApp = Nothing

End Sub

