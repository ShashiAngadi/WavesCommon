VERSION 5.00
Begin VB.Form frmAccReports 
   Caption         =   "Account Reports ..."
   ClientHeight    =   3600
   ClientLeft      =   1995
   ClientTop       =   2025
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6390
   Begin VB.TextBox txtDate1 
      Enabled         =   0   'False
      Height          =   395
      Left            =   1680
      TabIndex        =   7
      Top             =   2310
      Width           =   1275
   End
   Begin VB.TextBox txtDate2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   5010
      TabIndex        =   9
      Text            =   "12/12/2222"
      Top             =   2280
      Width           =   1185
   End
   Begin VB.ComboBox cmbReportList 
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   1
      Text            =   "Report List"
      Top             =   180
      Width           =   3885
   End
   Begin VB.ComboBox cmbRepParentHead 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   990
      Width           =   3915
   End
   Begin VB.ComboBox cmbRepHeadID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1500
      Width           =   3915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5100
      TabIndex        =   11
      Top             =   3030
      Width           =   1185
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3810
      TabIndex        =   10
      Top             =   3030
      Width           =   1185
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   6330
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label lblDate1 
      Caption         =   "From Date"
      Height          =   390
      Left            =   90
      TabIndex        =   6
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Label lblDate2 
      Caption         =   "To Date"
      Height          =   390
      Left            =   3480
      TabIndex        =   8
      Top             =   2295
      Width           =   1485
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   6210
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6225
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Select Report Type"
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   2205
   End
   Begin VB.Label lblRepAccHead 
      Caption         =   " Account Head  :"
      Height          =   360
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   2175
   End
   Begin VB.Label lblRepAccName 
      Caption         =   "Account Name :"
      Height          =   360
      Left            =   60
      TabIndex        =   4
      Top             =   1530
      Width           =   2175
   End
End
Attribute VB_Name = "frmAccReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event OKClick(StDate As String, EndDate As String, ParentID As Long, HeadID As Long, ReportSelected As Wis_AccountReportList)
Public Event CancelClick()

Private Sub SetKannadaCaption()

Dim ctrl As VB.Control

On Error Resume Next
For Each ctrl In Me
 ctrl.FontName = gFontName
 If Not TypeOf ctrl Is ComboBox Then
    ctrl.FontSize = gFontSize
 End If
Next

'fraReport.Caption = LoadResString(gLangOffSetNew + 505) & " " & LoadResString(gLangOffSet + 27)
lblRepAccHead.Caption = LoadResString(gLangOffSet + 232)
lblRepAccName.Caption = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 35)
lblDate1.Caption = LoadResString(gLangOffSet + 109)
lblDate2.Caption = LoadResString(gLangOffSet + 110)

End Sub

Private Sub cmbReportList_Click()

Dim ReportList As Wis_AccountReportList

With cmbReportList
    If .ListIndex = -1 Then Exit Sub
    ReportList = .ItemData(.ListIndex)
End With

cmbRepParentHead.Enabled = False
cmbRepHeadID.Enabled = False
txtDate1.Enabled = False
txtDate2.Enabled = False
cmdView.Enabled = False

Select Case ReportList

    Case AccountLedger
        
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        'txtDate1.Enabled = True
        'txtDate2.Enabled = True
        cmdView.Enabled = True
    
    Case AccountLedgerOnDate
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        txtDate1.Enabled = True
        txtDate2.Enabled = True
        cmdView.Enabled = True
                
    Case DayBook
        txtDate2.Enabled = True
        cmdView.Enabled = True
    
    Case AccountsClosed
    
    Case SubDayBook
        
        cmbRepParentHead.Enabled = True
        txtDate2.Enabled = True
        cmdView.Enabled = True
    
    Case BalancesAsON
    
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        txtDate2.Enabled = True
        cmdView.Enabled = True

    Case GeneralLedger
        
        'cmbRepParentHead.Enabled = True
        'cmbRepHeadID.Enabled = True
        
        txtDate2.Enabled = True
        cmdView.Enabled = True

    Case ProfitandLossTrans
    Case ReportNothing
    Case TotalTransActionsMade
    
End Select


End Sub


Private Sub cmbRepParentHead_Click()

With cmbRepParentHead
    If .ListIndex = -1 Then Exit Sub
    Call LoadLedgersToCombo(cmbRepHeadID, .ItemData(.ListIndex))
End With

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdView_Click()

Dim StartDate As String
Dim EndDate As String
Dim ParentID As Long
Dim HeadID As Long

Dim ReportList As Wis_AccountReportList

With cmbReportList
    If .ListIndex = -1 Then Exit Sub
    ReportList = .ItemData(.ListIndex)
End With
   
If cmbRepParentHead.Enabled Then
    With cmbRepParentHead
        If Not .ListIndex = -1 Then ParentID = .ItemData(.ListIndex)
    End With
End If

If cmbRepHeadID.Enabled Then
    With cmbRepHeadID
        If Not .ListIndex = -1 Then HeadID = .ItemData(.ListIndex)
    End With
End If

'If txtDate1.Enabled Then
    StartDate = txtDate1.Text
    'Check For Validate of Dates
    If Not DateValidate(StartDate, "/", True) Then
        txtDate1.SetFocus
        Exit Sub
    End If
'End If

'If txtDate2.Enabled Then
    EndDate = txtDate2.Text
    If Not DateValidate(EndDate, "/", True) Then
        txtDate2.SetFocus
        Exit Sub
    End If
'End If

If DateDiff("d", CDate(StartDate), CDate(EndDate)) < 0 Then
    MsgBox "Start date should be earlier than the end date ", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


Me.Hide

Screen.MousePointer = vbHourglass
RaiseEvent OKClick(StartDate, EndDate, ParentID, HeadID, ReportList)

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

CenterMe Me

Me.Icon = LoadResPicture(147, vbResIcon)

If gLangOffSet <> 0 Then SetKannadaCaption

LoadReportList

Call LoadParentHeads(cmbRepParentHead)

End Sub

Private Sub LoadReportList()

Dim ReportList As Wis_AccountReportList


With cmbReportList
    ReportList = BalancesAsON
    .AddItem "Balances As On"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = AccountsClosed
    .AddItem "Account Closed"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = GeneralLedger
    .AddItem "General Ledger"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = ProfitandLossTrans
    .AddItem "Profit and Loss Transactions"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = TotalTransActionsMade
    .AddItem "Total TransActionsMade"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = AccountLedger
    .AddItem "Ledger of Account"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = AccountLedgerOnDate
    .AddItem "Ledger of Account As On Date"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = DayBook
    .AddItem "Day Book"
    .ItemData(.NewIndex) = ReportList
    
    ReportList = SubDayBook
    .AddItem "Sub Day Book"
    .ItemData(.NewIndex) = ReportList
    
End With

End Sub








Private Sub Form_Resize()
On Error Resume Next

cmbReportList.SetFocus

End Sub


