VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrintDailog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WIS   -   Print Options..."
   ClientHeight    =   4590
   ClientLeft      =   2040
   ClientTop       =   2085
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdb 
      Left            =   1980
      Top             =   4260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      Height          =   315
      Left            =   3270
      TabIndex        =   47
      Top             =   4200
      Width           =   810
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5070
      TabIndex        =   49
      Top             =   4200
      Width           =   810
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   4200
      TabIndex        =   48
      Top             =   4200
      Width           =   810
   End
   Begin VB.Frame fraPaper 
      BorderStyle     =   0  'None
      Caption         =   "Paper size"
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   5685
      Begin VB.Frame fraWrap 
         Caption         =   "Worrd wrapping"
         Height          =   585
         Left            =   90
         TabIndex        =   55
         Top             =   1860
         Width           =   5505
         Begin VB.CheckBox chkWrapcell 
            Caption         =   "Wraps words of cell"
            Height          =   315
            Left            =   2910
            TabIndex        =   57
            Top             =   180
            Width           =   2445
         End
         Begin VB.CheckBox chkWrapHead 
            Caption         =   "Wrap words of heading "
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraHeader 
         Caption         =   "Header"
         Height          =   915
         Left            =   90
         TabIndex        =   22
         Top             =   2490
         Width           =   5505
         Begin VB.ComboBox cmbFooterAlign 
            Height          =   315
            Left            =   3060
            TabIndex        =   29
            Top             =   540
            Width           =   1065
         End
         Begin VB.CheckBox chkFooter 
            Caption         =   "FooterLine"
            Height          =   195
            Left            =   4245
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtFooter 
            Height          =   285
            Left            =   735
            TabIndex        =   28
            Top             =   540
            Width           =   2235
         End
         Begin VB.ComboBox cmbHeaderAlign 
            Height          =   315
            Left            =   3060
            TabIndex        =   25
            Top             =   180
            Width           =   1065
         End
         Begin VB.TextBox txtHeader 
            Height          =   285
            Left            =   720
            TabIndex        =   24
            Top             =   180
            Width           =   2235
         End
         Begin VB.CheckBox chkHeader 
            Caption         =   "Header Line"
            Height          =   285
            Left            =   4230
            TabIndex        =   26
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblFooter 
            Caption         =   "Footer :"
            Height          =   195
            Left            =   30
            TabIndex        =   27
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lblHeader 
            Caption         =   "Header :"
            Height          =   195
            Left            =   30
            TabIndex        =   23
            Top             =   210
            Width           =   645
         End
      End
      Begin VB.Frame fraPaperSize 
         Caption         =   "Paper"
         Height          =   1005
         Left            =   90
         TabIndex        =   2
         Top             =   0
         Width           =   5505
         Begin VB.ComboBox cmbPaper 
            Height          =   315
            Left            =   1470
            TabIndex        =   4
            Top             =   240
            Width           =   3435
         End
         Begin VB.TextBox txtHeight 
            Height          =   285
            Left            =   4320
            TabIndex        =   9
            Top             =   630
            Width           =   585
         End
         Begin VB.TextBox txtWidth 
            Height          =   285
            Left            =   1470
            TabIndex        =   6
            Top             =   630
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "inches"
            Height          =   195
            Left            =   4980
            TabIndex        =   10
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "inches"
            Height          =   225
            Left            =   2250
            TabIndex        =   7
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblPaperWidth 
            Caption         =   "Page width"
            Height          =   225
            Left            =   210
            TabIndex        =   5
            Top             =   660
            Width           =   1125
         End
         Begin VB.Label lblPaperHeight 
            Caption         =   "Page height :"
            Height          =   225
            Left            =   3120
            TabIndex        =   8
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label lblPapSize 
            Caption         =   "Paper Size :"
            Height          =   165
            Left            =   240
            TabIndex        =   3
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame fraMargin 
         Caption         =   "Margins in inches"
         Height          =   855
         Left            =   90
         TabIndex        =   11
         Top             =   990
         Width           =   5505
         Begin VB.CheckBox chkHorLine 
            Caption         =   "Horizontal Lines"
            Height          =   195
            Left            =   3780
            TabIndex        =   21
            Top             =   540
            Width           =   1545
         End
         Begin VB.CheckBox chkVerLine 
            Caption         =   "Verticle Lines"
            Height          =   195
            Left            =   3780
            TabIndex        =   20
            Top             =   210
            Width           =   1545
         End
         Begin VB.TextBox txtMarginRight 
            Height          =   285
            Left            =   2820
            TabIndex        =   19
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtMarginLeft 
            Height          =   285
            Left            =   1890
            TabIndex        =   17
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtMarginBottom 
            Height          =   285
            Left            =   990
            TabIndex        =   15
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtMarginTop 
            Height          =   285
            Left            =   90
            MaxLength       =   6
            TabIndex        =   13
            Top             =   450
            Width           =   705
         End
         Begin VB.Label Label6 
            Caption         =   "Right  :"
            Height          =   195
            Left            =   2760
            TabIndex        =   18
            Top             =   210
            Width           =   645
         End
         Begin VB.Label lblLeft 
            Caption         =   "Left :"
            Height          =   195
            Left            =   1875
            TabIndex        =   16
            Top             =   210
            Width           =   615
         End
         Begin VB.Label lblBottom 
            Caption         =   "Bottom  :"
            Height          =   195
            Left            =   915
            TabIndex        =   14
            Top             =   210
            Width           =   645
         End
         Begin VB.Label lblTop 
            Caption         =   "Top :"
            Height          =   195
            Left            =   90
            TabIndex        =   12
            Top             =   210
            Width           =   645
         End
      End
      Begin VB.Label lblPaperSize 
         Caption         =   "A4"
         Height          =   225
         Left            =   1410
         TabIndex        =   50
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame fraPrinter 
      BorderStyle     =   0  'None
      Caption         =   "Printer settings"
      Height          =   3645
      Left            =   210
      TabIndex        =   31
      Top             =   390
      Width           =   5685
      Begin VB.Frame fraPrinterName 
         Caption         =   "Printer"
         Height          =   675
         Left            =   120
         TabIndex        =   32
         Top             =   30
         Width           =   5505
         Begin VB.ComboBox cmbPrinter 
            Height          =   315
            Left            =   2070
            TabIndex        =   34
            Top             =   210
            Width           =   3285
         End
         Begin VB.Label lblPrinter 
            Caption         =   "Printer Name :"
            Height          =   285
            Left            =   240
            TabIndex        =   33
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Frame fraCopy 
         Caption         =   "Copies"
         Height          =   2685
         Left            =   3840
         TabIndex        =   43
         Top             =   750
         Width           =   1785
         Begin VB.OptionButton optLandScape 
            Caption         =   "Landscape"
            Height          =   255
            Left            =   60
            TabIndex        =   54
            Top             =   1680
            Width           =   1215
         End
         Begin VB.OptionButton optPortrait 
            Caption         =   "Portrait"
            Height          =   255
            Left            =   60
            TabIndex        =   53
            Top             =   1410
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CheckBox ChkExcel 
            Caption         =   "Print to &Excel"
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1545
         End
         Begin VB.ListBox lstCopy 
            Height          =   255
            ItemData        =   "PrintDlg.frx":0000
            Left            =   990
            List            =   "PrintDlg.frx":000D
            TabIndex        =   45
            Top             =   990
            Width           =   705
         End
         Begin VB.CheckBox chkCollate 
            Caption         =   "Colla&te"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   630
            Width           =   1575
         End
         Begin VB.Image imgOrientation 
            Height          =   525
            Left            =   330
            Top             =   2100
            Width           =   1155
         End
         Begin VB.Label lblCopies 
            Caption         =   "Copies  :"
            Height          =   255
            Left            =   90
            TabIndex        =   44
            Top             =   1020
            Width           =   885
         End
      End
      Begin VB.Frame fraPageRange 
         Caption         =   "Page range"
         Height          =   2685
         Left            =   90
         TabIndex        =   35
         Top             =   750
         Width           =   3675
         Begin VB.CheckBox chkPause 
            Caption         =   "&Pause between pages"
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   2250
            Width           =   2535
         End
         Begin VB.OptionButton optAllPage 
            Caption         =   "&All Pages"
            Height          =   255
            Left            =   150
            TabIndex        =   36
            Top             =   300
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optCurrPage 
            Caption         =   "Curr&ent page"
            Height          =   195
            Left            =   1860
            TabIndex        =   37
            Top             =   330
            Width           =   1305
         End
         Begin VB.OptionButton optPageRange 
            Caption         =   "Pa&ges"
            Height          =   285
            Left            =   150
            TabIndex        =   40
            Top             =   1080
            Width           =   1035
         End
         Begin VB.TextBox txtPageRange 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1230
            TabIndex        =   41
            Top             =   1050
            Width           =   2235
         End
         Begin VB.OptionButton optEven 
            Caption         =   "All e&ven pages"
            Height          =   195
            Left            =   1860
            TabIndex        =   39
            Top             =   630
            Width           =   1425
         End
         Begin VB.OptionButton optOdd 
            Caption         =   "All &Odd Pages"
            Height          =   255
            Left            =   150
            TabIndex        =   38
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblDesc 
            Caption         =   "Enter page numbers and/or page ranges saperated by commas. For example 1,3,5 -12"
            Height          =   435
            Left            =   150
            TabIndex        =   42
            Top             =   1560
            Width           =   3315
         End
         Begin VB.Line Line1 
            X1              =   90
            X2              =   3570
            Y1              =   960
            Y2              =   960
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4035
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7117
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Paper"
            Key             =   "Paper"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set the paper margin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Printer"
            Key             =   "Printer"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Shows the printer setting"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgLst 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PrintDlg.frx":001A
            Key             =   "Portrait"
            Object.Tag             =   "Portrait image"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PrintDlg.frx":0334
            Key             =   "Landscape"
            Object.Tag             =   "Lansd scapeimage"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPrintDailog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_View As String
Private m_PageRange As String
Public NoOfPages As Integer

Private Function CheckControls() As Boolean

If cmbPaper.ListIndex < 0 Then
    MsgBox "Plese specify the paper size", vbInformation, wis_MESSAGE_TITLE
    cmbPaper.SetFocus
    Exit Function
End If

If cmbPaper.ListIndex = cmbPaper.ListCount - 1 Then
    If PaperWidth = 0 Then
        MsgBox "Please specify the page width", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtWidth
    End If
    If PaperHeight = 0 Then
        MsgBox "Please specify the page height", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtHeight
    End If
End If
    
Dim strRange As String

'Now Check for the print rang
If Me.optAllPage Then m_PageRange = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"
If Me.optCurrPage Then m_PageRange = "0"
If Me.optEven Then m_PageRange = "2,4,6,8,10,12,14,16,18,20,22,24,26,28,30"
If Me.optOdd Then m_PageRange = "1,3,5,7,9,11,13,15,17,19,21,21,23,25,27,29,31,33"

If optPageRange Then
    Dim I As Integer
    Dim j As Integer
    Dim Pos As Integer
    Dim PrevPos As Integer
    Dim PageNo As Integer
    Dim PageArr() As String
    Dim FirstPart As String
    Dim SecPart As String
    Dim strValue As Integer
    
    strRange = txtPageRange
    If Trim(strRange) = "" Then
        MsgBox "Please specify the pages to be print", vbInformation, App.EXEName
        ActivateTextBox txtPageRange
        Exit Function
    End If
    I = 1: j = Len(strRange)
    PrevPos = Asc("0")
    Pos = Asc("9")
    
    'check for other charectors
    While I <= j
        FirstPart = Mid(strRange, I, 1)
        strValue = Asc(FirstPart)
        If strValue < PrevPos Or strValue > Pos Then
            If Not (strValue = Asc(" ") Or strValue = Asc("-") Or strValue = Asc(",")) Then
                MsgBox "Invalid page range specified", vbInformation, wis_MESSAGE_TITLE
                Exit Function
            End If
        End If
        I = I + 1
    Wend
    If GetStringArray(strRange, PageArr, ",") > 0 Then
        For I = LBound(PageArr) To UBound(PageArr)
            PrevPos = 0
            'Check the range here specified with the hypken ("-")
            Pos = InStr(PrevPos + 1, PageArr(I), "-", vbBinaryCompare)
            If Pos Then
                FirstPart = Left(PageArr(I), Pos - 1)
                SecPart = Mid(PageArr(I), Pos + 1)
                Pos = InStr(PrevPos + 1, strRange, "-", vbBinaryCompare)
                strRange = Left(strRange, Pos - 1) & "," & Mid(strRange, Pos + 1)
                If Val(SecPart) < Val(FirstPart) Then
                    MsgBox "Invalid page range specified", vbInformation, wis_MESSAGE_TITLE
                    Exit Function
                End If
                PrevPos = Val(FirstPart): Pos = Val(SecPart)
                For j = PrevPos + 1 To Pos - 1
                    strRange = strRange & "," & j
                Next
            End If
        Next
    End If
    
    'Here check page no which has repeated in the page range
    Call GetStringArray(strRange, PageArr(), ",")
    m_PageRange = ""
    For I = 0 To UBound(PageArr)
Recheck:
        Pos = Val(PageArr(I))
        If Pos = 0 Then GoTo NextCount
        For j = I + 1 To UBound(PageArr)
            PrevPos = Val(PageArr(j))
            If PageArr(I) = PageArr(j) Then PageArr(j) = 0
            If Pos > PrevPos Then  'swap the value
                Pos = Pos + PrevPos
                PageArr(j) = Pos - PrevPos
                PageArr(I) = Pos - Val(PageArr(j))
                GoTo Recheck
            End If
        Next
NextCount:
    Next
    
    For I = 0 To UBound(PageArr)
    If Val(PageArr(I)) > 0 Then m_PageRange = m_PageRange & "," & PageArr(I)
    Next
    m_PageRange = Mid(m_PageRange, 2)
    If Trim(m_PageRange) = "" Then
        MsgBox "Please specify the pages to be print", vbInformation, App.EXEName
        ActivateTextBox txtPageRange
        Exit Function
    End If
    
End If

'Finally give the page ranges
CheckControls = True
End Function


Public Property Get HorizontalLine() As Boolean
If chkHorLine.value = vbChecked Then HorizontalLine = True
End Property

Public Property Get Orientation() As Byte
If optPortrait Then
    Orientation = 1
ElseIf optLandScape Then
    Orientation = 2
End If
End Property

Public Property Get PageRange() As String


If optPageRange Then
    PageRange = txtPageRange
Else
    PageRange = ""
End If
PageRange = m_PageRange
End Property

Public Property Get PaperWidth() As Single

Dim StrWid As String
Dim Pos As Long

StrWid = txtWidth
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos - 1)

PaperWidth = Val(StrWid)
End Property

Public Property Get PaperHeight() As Single

Dim StrWid As String
Dim Pos As Long

StrWid = txtHeight
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos)

PaperHeight = Val(StrWid)
End Property


Public Property Get MarginLeft() As Single

Dim StrWid As String
Dim Pos As Long

StrWid = txtMarginLeft
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos - 1)

MarginLeft = Val(StrWid)
End Property
Public Property Get MarginRight() As Single

Dim StrWid As String
Dim Pos As Long

StrWid = txtMarginRight
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos - 1)

MarginRight = Val(StrWid)
End Property


Public Property Get MarginTop() As Single

Dim StrWid As String
Dim Pos As Long

StrWid = txtMarginTop
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos - 1)

MarginTop = Val(StrWid)
End Property


Public Property Get MarginBottom() As Single

Dim StrWid As String
Dim Pos As Long

StrWid = txtMarginBottom
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos - 1)

MarginBottom = Val(StrWid)
End Property

Public Property Get Footer() As String
Footer = txtFooter
End Property
Public Property Get Header() As String
Header = txtHeader
End Property
Public Property Get FooterLine() As Boolean
If chkFooter.value = vbChecked Then FooterLine = True
End Property

Public Property Get HeaderLine() As Boolean
If chkHeader.value = vbChecked Then HeaderLine = True
End Property


Public Property Get ToPage() As Long

End Property
Public Property Get FromPage() As Long

End Property
Public Property Get VerticleLine() As Boolean
If chkVerLine.value = vbChecked Then VerticleLine = True
End Property
Public Property Get view() As String
view = m_View
End Property

Private Sub cmbPaper_Click()
If cmbPaper.ListIndex < 0 Then Exit Sub

Dim Ht As String
Dim Wid As String

If cmbPaper.ListIndex = cmbPaper.ListCount - 1 Then
    txtWidth.Enabled = True
    txtHeight.Enabled = True
    Ht = (Printer.ScaleHeight / 1440) \ 1
    Wid = (Printer.ScaleWidth / 1440) \ 1
    Ht = GetSetting(App.EXEName, "Printer", "PaperHeight", Ht)
    Wid = GetSetting(App.EXEName, "Printer", "PaperWidth", Wid)
    GoTo ExitLine
Else
    txtWidth.Enabled = False
    txtHeight.Enabled = False
End If

With cmbPaper
    If .ListIndex = 0 Then Ht = "16.5 ": Wid = "23.3 "
    If .ListIndex = 1 Then Ht = "23.3 ": Wid = "15.5 "
    If .ListIndex = 2 Then Ht = "15.5 ": Wid = "23.3 "
    If .ListIndex = 3 Then Ht = "8.2 ": Wid = "11.7 "
    If .ListIndex = 4 Then Ht = "11.7 ": Wid = "8.2 "
    If .ListIndex = 5 Then Ht = "10.1 ": Wid = "14.3 "
    If .ListIndex = 6 Then Ht = "7.2 ": Wid = "10.1 "
    If .ListIndex = 7 Then Ht = "7.25 ": Wid = "10.5 "
    If .ListIndex = 8 Then Ht = "15 ": Wid = "12 "
    If .ListIndex = 9 Then Ht = "8.2 ": Wid = "12 "
    If .ListIndex = 10 Then Ht = "8.5 ": Wid = "14 "
    If .ListIndex = 11 Then Ht = "8.5 ": Wid = "11 "
End With

Ht = Ht & """"
Wid = Wid & """"

ExitLine:
txtHeight = Wid  'Ht
Call txtHeight_LostFocus

txtWidth = Ht 'Wid
Call txtWidth_LostFocus

End Sub


'Public event OKClicked
'set printer.DriverName
Private Sub cmdCancel_Click()
m_View = "Cancel"
Me.Hide
End Sub

Private Sub cmdPreview_Click()
If Not CheckControls Then Exit Sub
    m_View = "Preview"
    Me.Hide
End Sub
Private Sub cmdPrint_Click()
If Not CheckControls Then Exit Sub
m_View = "Printer"
Me.Hide
End Sub

Private Sub Form_Activate()
m_View = "Cancel"
'Set the default printer
Dim X As Integer
For X = 0 To cmbPrinter.ListCount - 1
    If cmbPrinter.ItemData(X) Then
        cmbPrinter.ListIndex = X
        Exit For
    End If
Next
'ChkExcel.Enabled = True

End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print "SHashi"
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
    If TabStrip1.SelectedItem.Index = TabStrip1.Tabs.Count Then
          TabStrip1.Tabs(1).Selected = True
    Else
          TabStrip1.Tabs(TabStrip1.SelectedItem.Index + 1).Selected = True
    End If
End If

End Sub


Private Sub Form_Load()

'Call CenterMe(Me)

'Load the Printere
Dim Prt As Printer

cmbPrinter.Clear
For Each Prt In Printers
    cmbPrinter.AddItem Prt.DeviceName
    If Printer.DeviceName = Prt.DeviceName Then
        cmbPrinter.ItemData(cmbPrinter.NewIndex) = 1
    Else
        cmbPrinter.ItemData(cmbPrinter.NewIndex) = 0
    End If
Next

'Now load the Paper Details
With cmbPaper
    .Clear
    .AddItem "A2 420 x 594 mm"
    .ItemData(.NewIndex) = 256  'Paper size
    .AddItem "A3 397 x 420 mm"
    .ItemData(.NewIndex) = 8  'Paper size
    .AddItem "A3Transeverse 420 x 397 mm" '2
    .ItemData(.NewIndex) = 8  'Paper size
    .AddItem "A4 210 x 297 mm"
    .ItemData(.NewIndex) = 9  'Paper size
    .AddItem "A4Transeverse 297 x 210 mm" '4
    .ItemData(.NewIndex) = 10  'Paper size
    .AddItem "B4(JIS) 257 x 364 mm"
    .ItemData(.NewIndex) = 12  'Paper size
    .AddItem "B5(JIS) 182 x 257 mm" '6
    .ItemData(.NewIndex) = 13  'Paper size
    .AddItem "Executive 7 1/4 x 10 1/2 in"
    .ItemData(.NewIndex) = 7  'Paper size
    .AddItem "Fanfold 15 x 12 in"  '8
    .ItemData(.NewIndex) = 39  'Paper size
    .AddItem "Fanfold 210mm x 12in"
    .ItemData(.NewIndex) = 40  'Paper size
    .AddItem "Legal 8 1/2 x 14 in"   '10
    .ItemData(.NewIndex) = 5  'Paper size
    .AddItem "Legal 8 1/2 x 11 in"  '11
    .ItemData(.NewIndex) = 1  'Paper size
    .AddItem "User - defined size"
    .ItemData(.NewIndex) = 256  'Paper size
End With

'Now Load Aligmment
Dim X As AlignmentConstants
X = vbLeftJustify
cmbHeaderAlign.AddItem "Left Align"
cmbHeaderAlign.ItemData(cmbHeaderAlign.NewIndex) = X
cmbFooterAlign.AddItem "Left Align"
cmbFooterAlign.ItemData(cmbFooterAlign.NewIndex) = X
X = vbCenter
cmbHeaderAlign.AddItem "Center Align"
cmbHeaderAlign.ItemData(cmbHeaderAlign.NewIndex) = X
cmbFooterAlign.AddItem "Center Align"
cmbFooterAlign.ItemData(cmbFooterAlign.NewIndex) = X
X = vbRightJustify
cmbHeaderAlign.AddItem "Right Align"
cmbHeaderAlign.ItemData(cmbHeaderAlign.NewIndex) = X
cmbFooterAlign.AddItem "Right Align"
cmbFooterAlign.ItemData(cmbFooterAlign.NewIndex) = X


'Now Load the Previous values
chkCollate.value = GetSetting(App.EXEName, "Printer", "Collate", "1")
cmbPaper.ListIndex = GetSetting(App.EXEName, "Printer", "PaperSize", "4")

txtMarginTop = GetSetting(App.EXEName, "Printer", "TopMargin", "1 """)
txtMarginLeft = GetSetting(App.EXEName, "Printer", "LeftMargin", "1.25 """)
txtMarginBottom = GetSetting(App.EXEName, "Printer", "BottomMargin", "1 """)
txtMarginRight = GetSetting(App.EXEName, "Printer", "RightMargin", "0.5 """)

chkVerLine.value = GetSetting(App.EXEName, "Printer", "VetricleLine", "0")
chkHorLine.value = GetSetting(App.EXEName, "Printer", "HorizontalLine", "0")
optPortrait.value = GetSetting(App.EXEName, "Printer", "Portrait", "-1")
optLandScape.value = Not optPortrait.value
Call optPortrait_Click

''Now load the details of Wrapping
chkWrapcell.value = GetSetting(App.EXEName, "Printer", "CellWrap", "0")
chkWrapHead.value = GetSetting(App.EXEName, "Printer", "HeadWrap", "0")

txtHeader = GetSetting(App.EXEName, "Printer", "Header", "Waves Information Systems")
cmbHeaderAlign.ListIndex = GetSetting(App.EXEName, "Printer", "HeaderAlign", "0")
chkHeader.value = GetSetting(App.EXEName, "Printer", "HeaderLine", "1")

txtFooter = GetSetting(App.EXEName, "Printer", "Footer", "Page X of Y")
cmbFooterAlign.ListIndex = GetSetting(App.EXEName, "Printer", "FooterAlign", "2")
chkFooter.value = GetSetting(App.EXEName, "Printer", "FooterLine", "1")

'optExcel.value = GetSetting(App.EXEName, "Printer", "PrintToExcel", -1)

X = GetSetting(App.EXEName, "Printer", "TabStrip", 1)
TabStrip1.Tabs(X).Selected = True

'Now chck the whther ms exceis exist in
'this coompouter or not If not disable
'the option to print in excel
ChkExcel.Enabled = ExcelExists

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Save the Values as default
Call SaveSetting(App.EXEName, "Printer", "Collate", chkCollate.value)
Call SaveSetting(App.EXEName, "Printer", "PaperSize", cmbPaper.ListIndex)
If cmbPaper.ListIndex = cmbPaper.ListCount - 1 Then
    Call SaveSetting(App.EXEName, "Printer", "PaperWidth", txtWidth)
    Call SaveSetting(App.EXEName, "Printer", "PaperHeight", txtHeight)
End If
Call SaveSetting(App.EXEName, "Printer", "TopMargin", txtMarginTop)
Call SaveSetting(App.EXEName, "Printer", "LeftMargin", txtMarginLeft)
Call SaveSetting(App.EXEName, "Printer", "BottomMargin", txtMarginBottom)
Call SaveSetting(App.EXEName, "Printer", "RightMargin", txtMarginRight)

Call SaveSetting(App.EXEName, "Printer", "CellWrap", chkWrapcell.value)
Call SaveSetting(App.EXEName, "Printer", "HeadWrap", chkWrapHead.value)

Call SaveSetting(App.EXEName, "Printer", "VetricleLine", chkVerLine.value)
Call SaveSetting(App.EXEName, "Printer", "HorizontalLine", chkHorLine.value)
Call SaveSetting(App.EXEName, "Printer", "Portrait", optPortrait.value)

Call SaveSetting(App.EXEName, "Printer", "Header", txtHeader)
Call SaveSetting(App.EXEName, "Printer", "HeaderAlign", cmbHeaderAlign.ListIndex)
Call SaveSetting(App.EXEName, "Printer", "HeaderLine", chkHeader.value)

Call SaveSetting(App.EXEName, "Printer", "Footer", txtFooter)
Call SaveSetting(App.EXEName, "Printer", "FooterAlign", cmbFooterAlign.ListIndex)
Call SaveSetting(App.EXEName, "Printer", "FooterLine", chkFooter.value)

Call SaveSetting(App.EXEName, "Printer", "TabStrip", TabStrip1.SelectedItem.Index)
End Sub


Private Sub optAllPage_Click()
txtPageRange.Enabled = False
End Sub

Private Sub optCurrPage_Click()
txtPageRange.Enabled = False
End Sub


Private Sub optEven_Click()
txtPageRange.Enabled = False
End Sub

Private Sub optLandScape_Click()
imgOrientation.Stretch = False
imgOrientation.Picture = imgLst.ListImages("Landscape").Picture
imgOrientation.Stretch = True

End Sub

Private Sub optOdd_Click()
txtPageRange.Enabled = False
End Sub


Private Sub optPageRange_Click()
txtPageRange.Enabled = True
End Sub

Private Sub optPortrait_Click()
If optPortrait Then
    imgOrientation.Stretch = False
    imgOrientation.Picture = imgLst.ListImages("Portrait").Picture
    imgOrientation.Stretch = True
Else
    imgOrientation.Picture = imgLst.ListImages("Landscape").Picture
End If
End Sub

Private Sub TabStrip1_Click()
On Error Resume Next
If TabStrip1.SelectedItem.Index = 1 Then
    fraPaper.ZOrder 0
    cmbPaper.SetFocus
Else
    fraPrinter.ZOrder 0
    cmbPrinter.SetFocus
End If
Err.Clear
'TabStrip1.Object = Me.ActiveControl
End Sub


Private Sub txtHeight_LostFocus()
Dim Pos As Integer
Dim StrWid As String

StrWid = txtHeight
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos)

txtHeight = Val(StrWid) & " """

End Sub

Private Sub txtMarginBottom_LostFocus()
Dim strMargin As String
strMargin = txtMarginBottom
If Trim(strMargin) = "" Then Exit Sub
'Now serach for the double quote(")
Dim Pos As Long
Dim SecPos As Long
Dim PtPos As Long

'Now remove the double quotes
Pos = InStr(1, strMargin, """", vbTextCompare)
If Pos Then strMargin = Left(strMargin, Pos - 1)

strMargin = Trim(strMargin) & " """

txtMarginBottom = strMargin

End Sub


Private Sub txtMarginLeft_LostFocus()
Dim strMargin As String
strMargin = txtMarginLeft
If Trim(strMargin) = "" Then Exit Sub
'Now serach for the double quote(")
Dim Pos As Long
Dim SecPos As Long
Dim PtPos As Long

'Now remove the double quotes
Pos = InStr(1, strMargin, """", vbTextCompare)
If Pos Then strMargin = Left(strMargin, Pos - 1)

strMargin = Trim(strMargin) & " """

txtMarginLeft = strMargin


End Sub

Private Sub txtMarginRight_LostFocus()
Dim strMargin As String

strMargin = txtMarginRight
If Trim(strMargin) = "" Then Exit Sub
'Now serach for the double quote(")
Dim Pos As Long

'Now remove the double quotes
Pos = InStr(1, strMargin, """", vbTextCompare)
If Pos Then strMargin = Left(strMargin, Pos - 1)

strMargin = Trim(strMargin) & " """

txtMarginRight = strMargin


End Sub


Private Sub txtMarginTop_LostFocus()
Dim strMargin As String
strMargin = txtMarginTop
If Trim(strMargin) = "" Then Exit Sub
'Now serach for the double quote(")
Dim Pos As Long
Dim SecPos As Long
Dim PtPos As Long

'Now remove the double quotes
Pos = InStr(1, strMargin, """", vbTextCompare)
If Pos Then strMargin = Left(strMargin, Pos - 1)

strMargin = Trim(strMargin) & " """

txtMarginTop = strMargin
End Sub

Private Sub txtWidth_LostFocus()
Dim Pos As Integer
Dim StrWid As String

StrWid = txtWidth
Pos = InStr(1, StrWid, """", vbTextCompare)
If Pos Then StrWid = Left(StrWid, Pos)

txtWidth = Val(StrWid) & " """
End Sub

