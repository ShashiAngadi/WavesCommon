VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAccTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger Entry"
   ClientHeight    =   5880
   ClientLeft      =   420
   ClientTop       =   1785
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clear"
      Height          =   345
      Index           =   0
      Left            =   9240
      TabIndex        =   2
      Top             =   5400
      Width           =   1035
   End
   Begin VB.TextBox txtParticulars 
      Height          =   825
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4350
      Width           =   4155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Index           =   1
      Left            =   10380
      TabIndex        =   3
      Top             =   5400
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid grdLedger 
      Height          =   4755
      Index           =   0
      Left            =   4380
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8387
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdLedger 
      Height          =   4755
      Index           =   1
      Left            =   4380
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8387
      _Version        =   393216
   End
   Begin VB.Frame fraTab 
      Height          =   3705
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   510
      Width           =   4215
      Begin VB.ComboBox cmbTab0Ledger 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1530
         TabIndex        =   18
         Text            =   "Ledger Type"
         Top             =   1230
         Width           =   2595
      End
      Begin VB.TextBox txtTab0Amount 
         Height          =   360
         Left            =   2280
         TabIndex        =   17
         Top             =   2700
         Width           =   1515
      End
      Begin VB.ComboBox cmbTab0Ledger 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1530
         TabIndex        =   16
         Text            =   "Ledger Name"
         Top             =   1740
         Width           =   2565
      End
      Begin VB.TextBox txtTab0CurrentDate 
         Height          =   330
         Left            =   1530
         TabIndex        =   15
         Top             =   690
         Width           =   1815
      End
      Begin VB.ComboBox cmbTab0Ledger 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1530
         TabIndex        =   14
         Text            =   "Voucher Type"
         Top             =   240
         Width           =   2625
      End
      Begin VB.OptionButton optTab0From 
         Caption         =   "From"
         Enabled         =   0   'False
         Height          =   435
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   2250
         Width           =   855
      End
      Begin VB.OptionButton optTab0From 
         Caption         =   "To"
         Enabled         =   0   'False
         Height          =   465
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   2250
         Width           =   1065
      End
      Begin VB.CommandButton cmdTab0AddtoGrid 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   2070
         TabIndex        =   11
         Top             =   3300
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdTab0AddtoGrid 
         Caption         =   "Add"
         Height          =   345
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   3300
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   345
         Left            =   3180
         TabIndex        =   9
         Top             =   3300
         Width           =   975
      End
      Begin VB.CommandButton cmdTab0Date 
         Caption         =   ".."
         Height          =   315
         Left            =   3420
         TabIndex        =   8
         Top             =   690
         Width           =   345
      End
      Begin VB.Label lbltab0Amount 
         Caption         =   "Amount"
         Height          =   270
         Index           =   4
         Left            =   930
         TabIndex        =   24
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label lbltab0Date 
         Caption         =   "Date"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblTab0AccountType 
         Caption         =   "Ledger Name"
         Height          =   435
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label lblTab0HeadType 
         Caption         =   "Ledger Type"
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1290
         Width           =   1365
      End
      Begin VB.Label lblTab0VoucherType 
         Caption         =   "Voucher Type"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblTab0Balance 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   360
         Index           =   0
         Left            =   2280
         TabIndex        =   19
         Top             =   2250
         Width           =   1515
      End
   End
   Begin VB.Frame fraTab 
      Height          =   3705
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   510
      Width           =   4215
      Begin VB.CommandButton cmdTab1Ledger 
         Caption         =   ".."
         Height          =   315
         Index           =   2
         Left            =   3870
         TabIndex        =   37
         Top             =   870
         Width           =   285
      End
      Begin VB.CommandButton cmdTab1Ledger 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   3900
         TabIndex        =   36
         Top             =   300
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cmbTab1Ledger 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1530
         TabIndex        =   30
         Text            =   "Ledgers"
         Top             =   270
         Width           =   2325
      End
      Begin VB.ComboBox cmbTab1Ledger 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1500
         TabIndex        =   29
         Top             =   840
         Width           =   2325
      End
      Begin VB.CommandButton cmdTab1Show 
         Caption         =   "&Show"
         Height          =   315
         Left            =   3030
         TabIndex        =   28
         Top             =   3120
         Width           =   1035
      End
      Begin VB.TextBox txtTab1Dates 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1620
         TabIndex        =   27
         Top             =   2040
         Width           =   2325
      End
      Begin VB.TextBox txtTab1Dates 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1620
         TabIndex        =   26
         Top             =   2520
         Width           =   2325
      End
      Begin VB.CheckBox chkTab1EnterDates 
         Caption         =   "EnterDates  "
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   25
         Top             =   1620
         Width           =   3705
      End
      Begin VB.Label lblTab1Ledgers 
         Caption         =   "Ledger Type"
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   34
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label lblTab1Ledgers 
         Caption         =   "Ledger Name"
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   33
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label lblTab1Dates 
         Caption         =   "From Date"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   32
         Top             =   2130
         Width           =   1425
      End
      Begin VB.Label lblTab1Dates 
         Caption         =   "To Date"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   31
         Top             =   2580
         Width           =   1575
      End
   End
   Begin ComctlLib.TabStrip tabTrans 
      Height          =   4275
      Left            =   30
      TabIndex        =   35
      Top             =   60
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7541
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transaction"
            Key             =   "Trans"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Show Ledger"
            Key             =   "Ledger"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Shows the ledger transctions"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   11340
      Y1              =   5310
      Y2              =   5310
   End
   Begin VB.Label lblLedgerName 
      AutoSize        =   -1  'True
      Caption         =   "Ledger Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4470
      TabIndex        =   1
      Top             =   60
      Width           =   6765
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAccTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' For the Combos
Private Const VoucherType = 0
Private Const LedgerType = 1
Private Const LedgerName = 2

'For the Options
Private Const FromOpt = 0
Private Const ToOpt = 1

' For the Add Commands
Private Const AddToGrid = 0
Private Const DeleteGrid = 1

' To Handle Grid Functions
Private m_GrdFunctions As clsGrdFunctions
Private m_AccTransClass As clsAccTrans

' To show ledgers
Private m_LedgerClass As clsLedger

Private m_DBOperation As wis_DBOperation
Private m_IsAllowTransDate As Boolean
Private m_IsStartParticulars As Boolean
Private m_ActiveTab As Byte

Private Function CheckTransDate() As Boolean

On Error GoTo Hell:

Dim LastTransDate As String
Dim CurrentDate As String

CheckTransDate = True
If m_IsAllowTransDate Then Exit Function


CheckTransDate = False

' Get the Last TransDate for the HeadID
LastTransDate = LoadLastTransDate

If LastTransDate = "" Then Exit Function

CurrentDate = txtTab0CurrentDate.Text

If Not TextBoxDateValidate(txtTab0CurrentDate, "/", True, True) Then Exit Function

If CDate(FormatDate(CurrentDate)) < CDate(FormatDate(LastTransDate)) Then
    
    If MsgBox("Current Date is Smaller than Last Entered Date!" & _
        vbCrLf & "Do You Want To Continue ?", vbQuestion + vbYesNo) = vbNo Then
        Exit Function
    Else
        m_IsAllowTransDate = True
        CheckTransDate = True
        Exit Function
    End If
Else
    CheckTransDate = True
End If

Exit Function

Hell:
    MsgBox "Check TransDate :" & vbCrLf & Err.Description
    
End Function

Private Sub ClearControls()

If MsgBox("Do You Want To Clear Controls ? ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
grdLedger(m_ActiveTab).Clear
If m_ActiveTab = 0 Then
    RefreshForTab0
    m_AccTransClass.frmClearClicked
Else
    RefreshForTab1
End If

End Sub

'This Procedure  will get the last transacted date from acctrans and load it to
' m_LastTransDate
Private Function LoadLastTransDate() As String

Dim rstTransDate As ADODB.Recordset
Dim HeadID As Long

LoadLastTransDate = ""

With cmbTab0Ledger(LedgerName)
    If .ListIndex = -1 Then Exit Function
    HeadID = .ItemData(.ListIndex)
End With
    
gDbTrans.SQLStmt = " SELECT MAX(TransDate) as MaxTransDate " & _
                   " FROM AccTrans WHERE HeadID = " & HeadID
                 
Call gDbTrans.Fetch(rstTransDate, adOpenForwardOnly)

LoadLastTransDate = FormatField(rstTransDate.Fields("MaxTransDate"))

' If no transaction is made then last trans date will be first day
If LoadLastTransDate = "" Then LoadLastTransDate = FinIndianFromDate

Set rstTransDate = Nothing

End Function

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the kannada for the Tabs
tabTrans.Tabs(1).Caption = LoadResString(gLangOffSet + 28)
tabTrans.Tabs(2).Caption = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 13)

tabTrans.Font.Name = gFontName
tabTrans.Font.Size = gFontSize

'set the Kannada for all controls
lblTab0VoucherType(0).Caption = LoadResString(gLangOffSet + 41)    'Voucher
lbltab0Date(1).Caption = LoadResString(gLangOffSet + 37)  'NAme
lblTab0HeadType(2).Caption = LoadResString(gLangOffSet + 160) & " " & _
            LoadResString(gLangOffSet + 36)  'Main Account
lblTab0AccountType(3).Caption = LoadResString(gLangOffSet + 36) & " " & _
                LoadResString(gLangOffSet + 35)  'Account NAme
optTab0From(0).Caption = LoadResString(gLangOffSet + 107)  'From
optTab0From(1).Caption = LoadResString(gLangOffSet + 108)   'TO
lbltab0Amount(4).Caption = LoadResString(gLangOffSet + 40)  'Amount
lblLedgerName.Caption = LoadResString(gLangOffSet + 36) & " " & _
            LoadResString(gLangOffSet + 35)  'Account NAme

cmdTab0AddtoGrid(0).Caption = LoadResString(gLangOffSet + 10)   'Add
cmdTab0AddtoGrid(1).Caption = LoadResString(gLangOffSet + 14)   ''Delete
cmdOk.Caption = LoadResString(gLangOffSet + 7)   'Save
cmdCancel(0).Caption = LoadResString(gLangOffSet + 8)   'Clear
cmdCancel(1).Caption = LoadResString(gLangOffSet + 11)  'Close

lblTab1Ledgers(0).Caption = LoadResString(gLangOffSet + 160) & " " & _
            LoadResString(gLangOffSet + 36)  'Main Account
lblTab1Ledgers(1).Caption = LoadResString(gLangOffSet + 36) & " " & _
            LoadResString(gLangOffSet + 35)  'Account NAme
lblTab1Dates(0).Caption = LoadResString(gLangOffSet + 109) 'After Date
lblTab1Dates(1).Caption = LoadResString(gLangOffSet + 110)  'Before Date
cmdTab1Show.Caption = LoadResString(gLangOffSet + 13)   'Show
chkTab1EnterDates.Caption = LoadResString(gLangOffSet + 106)    'Specify dat range

End Sub


' Handles total entries made
'Private m_Entries As Integer
Public Sub InitTab0Grid()

With grdLedger(0)
    .Clear
    .Enabled = True
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 5
    .FixedCols = 1
    .FixedRows = 1
    
    .Row = 0
    
    .Col = 0: .CellFontBold = True: .Text = LoadResString(gLangOffSet + 33)  '"SlNo"
    .Col = 1: .CellFontBold = True: .Text = LoadResString(gLangOffSet + 36)  '"Ledger Name"
    .Col = 2: .CellFontBold = True: .Text = LoadResString(gLangOffSet + 277)  '"Dr "
    .Col = 3: .CellFontBold = True: .Text = LoadResString(gLangOffSet + 276)  '"Cr "
    .Col = 4: .CellFontBold = True: .Text = LoadResString(gLangOffSet + 42)  '"Total "
    
    .ColWidth(0) = 435
    .ColWidth(1) = 1900
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1500
    
    .Row = 1
End With


End Sub



'If will when tab0 is clicked
Public Sub RefreshForTab0()
    
    Call InitTab0Grid
    cmbTab0Ledger(LedgerName).Text = ""
    cmbTab0Ledger(0).Locked = False
    'm_AccTransClass.frmDeleteClicked
    
End Sub

'If will when tab1 is clicked

Public Sub RefreshForTab1()
    
'    grdLedger(m_ActiveTab).Visible = True
'    grdLedger(m_ActiveTab).ZOrder 0
    
    Call LoadParentHeads(cmbTab1Ledger(0))
    
    cmbTab1Ledger(1).Clear
    txtTab1Dates(0).Text = ""
    txtTab1Dates(1).Text = ""
    txtParticulars.Text = ""
    txtParticulars.Enabled = False
    
End Sub

'
Private Sub RefreshOptionButtons()
    
Dim VoucherTypes As Wis_VoucherTypes

' This Will Fetch The Current Voucher Type Selected by the User

With cmbTab0Ledger(VoucherType)
    VoucherTypes = .ItemData(.ListIndex)
End With

'Please dont change the case creiteria(Why?)
Select Case VoucherTypes

    Case Payment
        optTab0From(ToOpt).value = True
        optTab0From(ToOpt).Enabled = True
    Case Receipt
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case Sales
        optTab0From(ToOpt).value = True
        optTab0From(ToOpt).Enabled = True
    Case Purchase
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case FreePurchase
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case FreeSales
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case Contra
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case Journal
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case RejectionsIn
        optTab0From(FromOpt).value = True
        optTab0From(FromOpt).Enabled = True
    Case RejectionsOut
        optTab0From(ToOpt).value = True
        optTab0From(ToOpt).Enabled = True
End Select

End Sub

'
Private Sub SwapCheckedOptions()

If chkTab1EnterDates.value = vbChecked Then
    txtTab1Dates(0).Text = FinIndianFromDate
    txtTab1Dates(1).Text = FinIndianEndDate
    txtTab1Dates(0).Enabled = True
    txtTab1Dates(1).Enabled = True
Else
    txtTab1Dates(0).Text = ""
    txtTab1Dates(1).Text = ""
    txtTab1Dates(0).Enabled = False
    txtTab1Dates(1).Enabled = False
End If

End Sub


'
Private Sub UnLoadME()
    Unload Me
End Sub

'
Private Sub chkTab1EnterDates_Click()
    SwapCheckedOptions
End Sub

Private Sub cmbTab0Ledger_Click(Index As Integer)

Dim VoucherType As Wis_VoucherTypes

If cmbTab0Ledger(Index).ListIndex = -1 Then Exit Sub

'Don't Change this Case creteria
Select Case Index
        
        Case VoucherType
             ' RefreshOptionButtons
             m_AccTransClass.frmVoucherClicked
             
        Case LedgerType
        
            ' This load the Ledgers to combo
            Call LoadLedgersToCombo(cmbTab0Ledger(LedgerName), _
            cmbTab0Ledger(LedgerType).ItemData(cmbTab0Ledger(LedgerType).ListIndex))
            
            'lblTab0Balance(0).Caption = "Balance"
                        
            With cmbTab0Ledger(LedgerName)
                If .ListCount > 0 Then .ListIndex = 0
            End With
            cmbTab0Ledger_Click (LedgerName)
        Case LedgerName
        
            If Not DateValidate(txtTab0CurrentDate.Text, "/", True) Then Exit Sub
            
            ' This will Fetch the Balance for the HeadId
            With cmbTab0Ledger(LedgerName)
                m_AccTransClass.frmHeadClicked (.ItemData(.ListIndex))
            End With
End Select

End Sub


'
Private Sub cmbTab1Ledger_Click(Index As Integer)

If Index = 0 Then

    If cmbTab1Ledger(0).ListIndex = -1 Then Exit Sub
    
    ' This load the Ledgers to combo
    
    Call LoadLedgersToCombo(cmbTab1Ledger(1), _
    cmbTab1Ledger(0).ItemData(cmbTab1Ledger(0).ListIndex))
    
Else
    
End If

End Sub


'
Private Sub cmdTab0AddtoGrid_Click(Index As Integer)

If Index = DeleteGrid Then m_AccTransClass.frmDeleteClicked

If Index = AddToGrid Then
    If m_AccTransClass.UpdatingGrid Then
        m_AccTransClass.frmUpdateClicked
    Else
        m_AccTransClass.frmAddClicked
    End If
End If

End Sub
'
Private Sub cmdCancel_Click(Index As Integer)

If Index = 0 Then ClearControls
If Index = 1 Then UnLoadME

End Sub

Private Sub cmdOk_Click()

If m_AccTransClass.frmOKClicked = Success Then m_IsAllowTransDate = False
    
End Sub

Private Sub cmdTab0Date_Click()
With Calendar
    .Top = Me.Top + tabTrans.Top + cmdTab0Date.Top
    .Left = Me.Left + tabTrans.Left + cmdTab0Date.Left
    .SelDate = IIf(DateValidate(txtTab0CurrentDate, "/", True), txtTab0CurrentDate, FormatDate(gStrDate))
    .Show vbModal
    txtTab0CurrentDate = .SelDate
End With

End Sub

Private Sub cmdTab1Ledger_Click(Index As Integer)
If Index = LedgerName Then
    If m_LedgerClass Is Nothing Then Set m_LedgerClass = New clsLedger
    If cmbTab1Ledger(0).ListIndex >= 0 Then _
        m_LedgerClass.ParentID = cmbTab1Ledger(0).ItemData(cmbTab1Ledger(0).ListIndex)
    
    m_LedgerClass.ShowLedger

ElseIf Index = LedgerType Then
    
    frmSubParent.Show vbModal
    Call LoadParentHeads(cmbTab0Ledger(LedgerType))

End If

'Load the parent heads
Call cmbTab0Ledger_Click(LedgerType)
'LOad the sub heads
Call cmbTab1Ledger_Click(0)

End Sub

Private Sub cmdTab1Show_Click()

On Error GoTo Hell:

' Declarations
Dim lngHeadID As Long
Dim strFromDate As String
Dim strToDate As String
Dim lngCheckedValue As Long

If cmbTab1Ledger(0).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Parent Ledger Not Selected"
If cmbTab1Ledger(1).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Ledger Not Selected"

lngCheckedValue = chkTab1EnterDates.value
     
Set m_LedgerClass = New clsLedger

If lngCheckedValue = vbChecked Then
    
    strFromDate = txtTab1Dates(0).Text
    If Not TextBoxDateValidate(txtTab1Dates(0), "/", True, True) Then Exit Sub
    
    strToDate = txtTab1Dates(1).Text
    If Not TextBoxDateValidate(txtTab1Dates(1), "/", True, True) Then Exit Sub
    
End If

With cmbTab1Ledger(1)

    lngHeadID = .ItemData(.ListIndex)

End With

Me.MousePointer = vbHourglass

m_IsStartParticulars = False

If lngCheckedValue = vbChecked Then
    Call m_AccTransClass.ShowNewLedgerToGrid(lngHeadID, grdLedger(m_ActiveTab), True, _
                            strFromDate, strToDate)
Else
    Call m_AccTransClass.ShowNewLedgerToGrid(lngHeadID, grdLedger(m_ActiveTab), False, "", "")
End If

m_IsStartParticulars = True

Me.MousePointer = vbDefault
    
If lngCheckedValue = vbChecked Then _
    lblLedgerName.Caption = cmbTab1Ledger(1).Text & "           " _
        & strFromDate & " To " & strToDate
        
If lngCheckedValue <> vbChecked Then _
    lblLedgerName.Caption = cmbTab1Ledger(1).Text & "           " _
        & FinIndianFromDate & " To " & FinIndianEndDate
   
   
Exit Sub

Hell:
    
    MsgBox "Voucher Entry : " & vbCrLf & Err.Description
    
End Sub


Private Sub Form_Initialize()
Debug.Print "FOrm Init"
If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions
If m_AccTransClass Is Nothing Then Set m_AccTransClass = New clsAccTrans

Set m_GrdFunctions.fGrd = grdLedger(m_ActiveTab)

End Sub

'
Private Sub Form_Load()
'
CenterMe Me

Me.Icon = LoadResPicture(147, vbResIcon)

If gLangOffSet <> 0 Then SetKannadaCaption

'LOad The Current Date
txtTab0CurrentDate = FormatDate(gStrDate)

'Load the parent heads
Call LoadVouchersToCombo(cmbTab0Ledger(VoucherType))
Call LoadParentHeads(cmbTab0Ledger(LedgerType))

Call LoadParentHeads(cmbTab1Ledger(0))


'Call LoadLedgersToCombo(cmbTab1Ledger(1), _
    cmbTab1Ledger(0).ItemData(cmbTab1Ledger(0).ListIndex))

'RefreshForTab0
Call tabTrans_Click
'Set The Grid
grdLedger(1).Top = grdLedger(0).Top
grdLedger(1).Left = grdLedger(0).Left
grdLedger(1).Width = grdLedger(0).Width
grdLedger(1).Height = grdLedger(0).Height

Call InitTab0Grid
Call m_AccTransClass.InitTab1Grid(grdLedger(1))

grdLedger(0).Visible = True
grdLedger(0).ZOrder 0

Call Form_Resize

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set m_AccTransClass = Nothing
Set m_LedgerClass = Nothing
Set m_GrdFunctions = Nothing

End Sub

Private Sub Form_Resize()
Call m_AccTransClass.SetAccTrans(Me)
Set m_GrdFunctions.fGrd = grdLedger(0)
End Sub

Private Sub Form_Terminate()
Debug.Print "Form Term"
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmAccTrans = Nothing

End Sub

Public Property Get DbOperation() As wis_DBOperation

    DbOperation = m_DBOperation
    
End Property

Public Property Let DbOperation(ByVal NewValue As wis_DBOperation)
    m_DBOperation = NewValue
End Property

Private Sub grdLedger_DblClick(Index As Integer)
Dim RowNum As Integer

grdLedger(Index).Visible = False

    RowNum = grdLedger(Index).Row
    Call m_AccTransClass.frmGridClicked(RowNum, Index)

grdLedger(Index).Visible = True

End Sub


Private Sub grdLedger_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 And KeyAscii = 4 Then
    DbOperation = DeleteRec
    m_AccTransClass.frmDeleteIDClicked
    m_IsAllowTransDate = False
End If
End Sub

Private Sub grdLedger_RowColChange(Index As Integer)

If m_IsStartParticulars Then
    txtParticulars.Text = m_LedgerClass.GetTransIDParticulars(grdLedger(Index).RowData(grdLedger(Index).Row))
End If

End Sub

Private Sub tabTrans_Click()

m_ActiveTab = tabTrans.SelectedItem.Index - 1
fraTab(m_ActiveTab).ZOrder 0

grdLedger(1 - m_ActiveTab).Visible = False
grdLedger(m_ActiveTab).Visible = True
grdLedger(m_ActiveTab).ZOrder 0
    
Set m_GrdFunctions.fGrd = grdLedger(m_ActiveTab)

'Shashi On 16/12/2002

End Sub

