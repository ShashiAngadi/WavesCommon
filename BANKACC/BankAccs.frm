VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBankAcc 
   Caption         =   "Bank Accounts..."
   ClientHeight    =   7125
   ClientLeft      =   1410
   ClientTop       =   1305
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6540
      TabIndex        =   30
      Top             =   6660
      Width           =   1185
   End
   Begin VB.Frame fra 
      Height          =   5835
      Index           =   1
      Left            =   420
      TabIndex        =   31
      Top             =   570
      Width           =   7245
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "..."
         Height          =   300
         Left            =   2730
         TabIndex        =   43
         Top             =   1200
         Width           =   345
      End
      Begin VB.CommandButton cmdAccNames 
         Caption         =   "..."
         Height          =   315
         Left            =   6645
         TabIndex        =   2
         Top             =   675
         Width           =   315
      End
      Begin VB.ComboBox cmbAccNames 
         Height          =   315
         Left            =   2235
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   4260
      End
      Begin VB.ComboBox cmbAccHeads 
         Height          =   315
         ItemData        =   "BankAccs.frx":0000
         Left            =   2235
         List            =   "BankAccs.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   4260
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Index           =   0
         Left            =   150
         TabIndex        =   33
         Top             =   1080
         Width           =   6795
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1545
         Width           =   1965
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   4830
         TabIndex        =   5
         Top             =   1185
         Width           =   2085
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   180
         TabIndex        =   32
         Top             =   2340
         Width           =   6825
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5820
         TabIndex        =   8
         Top             =   5400
         Width           =   1305
      End
      Begin VB.CheckBox chkBackLog 
         Caption         =   "Clear backlog"
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   5430
         Width           =   1995
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   1215
         Width           =   1485
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4170
         TabIndex        =   9
         Top             =   5400
         Width           =   1425
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   1905
         Width           =   5805
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Left            =   4845
         TabIndex        =   6
         Top             =   1500
         Width           =   2100
      End
      Begin VB.Frame fraPassBook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2175
         Left            =   300
         TabIndex        =   34
         Top             =   2970
         Width           =   6675
         Begin VB.CommandButton cmdNextTrans 
            Caption         =   ">"
            Height          =   315
            Left            =   6270
            TabIndex        =   13
            Top             =   690
            Width           =   375
         End
         Begin VB.CommandButton cmdPrevTrans 
            Caption         =   "<"
            Height          =   315
            Left            =   6270
            TabIndex        =   12
            Top             =   135
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2025
            Left            =   150
            TabIndex        =   35
            Top             =   120
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3572
            _Version        =   393216
            Rows            =   5
            AllowUserResizing=   1
         End
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2835
         Left            =   165
         TabIndex        =   11
         Top             =   2505
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   5001
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Pass book"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAccName 
         AutoSize        =   -1  'True
         Caption         =   "Select the account name :"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label lblAccTitle 
         AutoSize        =   -1  'True
         Caption         =   "Select Account head type :"
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblTrans 
         Caption         =   "Transaction : "
         Height          =   195
         Left            =   150
         TabIndex        =   40
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   255
         Left            =   3810
         TabIndex        =   39
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label lblInstrNo 
         Caption         =   "Instument no:"
         Height          =   195
         Left            =   3750
         TabIndex        =   38
         Top             =   1605
         Width           =   1005
      End
      Begin VB.Label lblParticular 
         Caption         =   "Particulars : "
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   1995
         Width           =   945
      End
      Begin VB.Label lblDate 
         Caption         =   "Date : "
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   1245
         Width           =   735
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   6390
      Left            =   240
      TabIndex        =   14
      Top             =   150
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   11271
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transactions"
            Key             =   "Transactions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "REPORT"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   5835
      Index           =   2
      Left            =   420
      TabIndex        =   44
      Top             =   570
      Width           =   7230
      Begin VB.ComboBox cmbRepAccHead 
         Height          =   315
         Left            =   2250
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   270
         Width           =   4650
      End
      Begin VB.Frame fraReport 
         Caption         =   "Choose a report"
         Height          =   1785
         Left            =   195
         TabIndex        =   54
         Top             =   1305
         Width           =   6735
         Begin VB.OptionButton optReports 
            Caption         =   "Account Balances where"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   17
            Top             =   270
            Value           =   -1  'True
            Width           =   2385
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Balances as on"
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   18
            Top             =   630
            Width           =   1725
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Total transactions made"
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   19
            Top             =   1020
            Width           =   2325
         End
         Begin VB.OptionButton optReports 
            Caption         =   "General ledger"
            Height          =   285
            Index           =   5
            Left            =   3390
            TabIndex        =   22
            Top             =   1050
            Width           =   1755
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Profit && loss Transaction"
            Height          =   285
            Index           =   3
            Left            =   3375
            TabIndex        =   20
            Top             =   225
            Width           =   3135
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Accounts Closed"
            Height          =   285
            Index           =   4
            Left            =   3390
            TabIndex        =   21
            Top             =   630
            Width           =   1755
         End
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   195
         TabIndex        =   49
         Top             =   4020
         Width           =   6795
         Begin VB.TextBox txtAmt2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5145
            TabIndex        =   28
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txtAmt1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1755
            TabIndex        =   27
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lblAmt1 
            Caption         =   "amount exceeds Rs."
            Enabled         =   0   'False
            Height          =   255
            Left            =   195
            TabIndex        =   51
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label lblAmt2 
            Caption         =   "and lies within Rs."
            Enabled         =   0   'False
            Height          =   255
            Left            =   3495
            TabIndex        =   50
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Index           =   1
         Left            =   165
         TabIndex        =   48
         Top             =   1200
         Width           =   6915
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5775
         TabIndex        =   29
         Top             =   5310
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Height          =   765
         Left            =   180
         TabIndex        =   45
         Top             =   3150
         Width           =   6765
         Begin VB.CommandButton cmdDate2 
            Height          =   300
            Left            =   6240
            TabIndex        =   25
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdDate1 
            Height          =   300
            Left            =   2835
            TabIndex        =   23
            Top             =   315
            Width           =   390
         End
         Begin VB.TextBox txtDate2 
            Height          =   285
            Left            =   5130
            TabIndex        =   26
            Top             =   330
            Width           =   1065
         End
         Begin VB.TextBox txtDate1 
            Height          =   285
            Left            =   1770
            TabIndex        =   24
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label lblDate2 
            Caption         =   "and before (dd/mm/yyyy)"
            Height          =   255
            Left            =   3300
            TabIndex        =   47
            Top             =   375
            Width           =   1875
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Height          =   225
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.ComboBox cmbRepAccName 
         Height          =   315
         Left            =   2235
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   705
         Width           =   4650
      End
      Begin VB.Label lblRepAccHead 
         Caption         =   " Account Head  :"
         Height          =   270
         Left            =   300
         TabIndex        =   53
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label lblRepAccName 
         Caption         =   "Account Name :"
         Height          =   240
         Left            =   315
         TabIndex        =   52
         Top             =   780
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmBankAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_AccHead  As String
Dim m_AccID As Long
Dim m_CustReg As clsCustReg
Dim m_rstPassBook As Recordset

Private WithEvents m_frmBankReport As frmBankReport
Attribute m_frmBankReport.VB_VarHelpID = -1

Private Function ComputeInterest(TillDate As String, AccountID As Long) As Currency
Dim Days As Integer
Dim IntAmount As Currency
Dim lastDate As String
Dim Amount As Currency
Dim Balance As Currency
    gDBTrans.SQLStmt = "select transdate, transtype, balance, amount from" & _
        " acctrans where transdate <= #" & TillDate & "# " & " And accid=" & AccountID
    If gDBTrans.SQLFetch <= 0 Then
        GoTo ExitLine
    End If
    gDBTrans.Rst.MoveFirst
    lastDate = FormatField(gDBTrans.Rst("Transdate"))
    While Not gDBTrans.Rst.EOF
        Amount = FormatField(gDBTrans.Rst("Amount"))
        Balance = FormatField(gDBTrans.Rst("balance"))
        If gDBTrans.Rst("Transtype") = wDeposit Then
            Balance = Balance - Amount
        ElseIf gDBTrans.Rst("transtype") = wWithDraw Then
            gDBTrans.Rst.MovePrevious
            Balance = FormatField(gDBTrans.Rst("balance"))
            gDBTrans.Rst.MoveNext
        End If
        Days = WisDateDiff(FormatField(gDBTrans.Rst("transdate")), TillDate)
        If Not Days < 0 Then
            IntAmount = IntAmount + (Balance * Days * 18 / 36500)
        End If
        gDBTrans.Rst.MoveNext
    Wend
    ComputeInterest = IntAmount

ExitLine:

End Function
Private Sub LoadAccountHeads()
    
    'Load the AccountHeads  in to Combo Box
    Dim BankHeadId As wisBankHeads
 
 'Load BankHeads
BankHeadId = wis_BankHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 451) '"Bank Accounts"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId   '1 * wis_BankHeadOffSet

'Load Bank Loan Accounts
BankHeadId = wis_BankLoanHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 452) '"Bank Loan Accounts"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId   '2 * wis_BankHeadOffSet

'Load Advances
BankHeadId = wis_AdvanceHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 453) '"Advances"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId  '3 * wis_BankHeadOffSet
'Invests
BankHeadId = wis_InvestmentHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 454) '"Investments"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId

'Load Income Heads
BankHeadId = wis_IncomeHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 455) '"Income Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId   '5 * wis_BankHeadOffSet

'Load Expense heads
BankHeadId = wis_ExpenditureHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 456) '"Expense Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '6 * wis_BankHeadOffSet

'if Firm is bank then need not to load these Expense
If gBank Then
    'Load Trading Income heads
     BankHeadId = wis_TradingIncomeHead
    cmbAccHeads.AddItem LoadResString(gLangOffSet + 457) '"Trading Income Heads"
    cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '7 * wis_BankHeadOffSet
    
    'Load Trading Expense heads
    BankHeadId = wis_TradingExpenditureHead
    cmbAccHeads.AddItem LoadResString(gLangOffSet + 458) '"Trading Expense Heads"
    cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '8 * wis_BankHeadOffSet
End If
  
'Load Balance Sheet Heads
BankHeadId = wis_ReserveFundHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 459) '"Fund Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '9 * wis_BankHeadOffSet

'Load Share Capital Heads
BankHeadId = wis_ShareCapitalHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 264) '"Share Capitol Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '10 * wis_BankHeadOffSet

'Load Subsidy Heads
BankHeadId = wis_GovtLoanSubsidyHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 263) '"Govt Loan Subsidy" '"Subsidy Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '11 * wis_BankHeadOffSet

'PaymnetHeads
BankHeadId = wis_PaymentHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 426) '"Payment Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '12 * wis_BankHeadOffSet

'Repaymnet Heads
BankHeadId = wis_RepaymentHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 427) '"Repayment Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '14 * wis_BankHeadOffSet

'Load Balance Sheet Asset headsHeads
BankHeadId = wis_AssetHead
cmbAccHeads.AddItem LoadResString(gLangOffSet + 428) '"Asset Heads"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '9 * wis_BankHeadOffSet

'These two Head are particularly for Hirepadasalagi
'the below code should me removed while shifting to other banks

BankHeadId = wis_MemberDeposits
cmbAccHeads.AddItem LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 43)  '"Member Deposits"
cmbAccHeads.ItemData(cmbAccHeads.NewIndex) = BankHeadId     '20 * wis_BankHeadOffSet

End Sub

Private Sub SetKannadaCaption()

Dim Ctrl As Control
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
         Ctrl.Font.Size = gFontSize
     End If
Next

'Assign Kannada caption to all the controls

'Set kannadacaption for the tabs
Me.TabStrip.Tabs(1).Caption = LoadResString(gLangOffSet + 210)
Me.TabStrip.Tabs(2).Caption = LoadResString(gLangOffSet + 212)

'Set kannadacaption for the controls in general fom
Me.cmdOK.Caption = LoadResString(gLangOffSet + 1)

'Set kannada caption for the transaction frame
Me.lblAccTitle.Caption = LoadResString(gLangOffSet + 162)
Me.lblAccName.Caption = LoadResString(gLangOffSet + 163)
Me.lblDate.Caption = LoadResString(gLangOffSet + 37)
Me.lblTrans.Caption = LoadResString(gLangOffSet + 38)
Me.lblParticular.Caption = LoadResString(gLangOffSet + 39)
Me.lblAmount.Caption = LoadResString(gLangOffSet + 40)
Me.lblInstrNo.Caption = LoadResString(gLangOffSet + 41)
Me.chkBackLog.Caption = LoadResString(gLangOffSet + 164)
Me.cmdUndo.Caption = LoadResString(gLangOffSet + 5)
Me.cmdAccept.Caption = LoadResString(gLangOffSet + 4)

' set kannada caption for tabstrip 2
Me.TabStrip2.Tabs(1).Caption = LoadResString(gLangOffSet + 218)

' Set kannada captions for the reports frame
Me.lblRepAccHead.Caption = LoadResString(gLangOffSet + 159)
Me.lblRepAccName.Caption = LoadResString(gLangOffSet + 160)
Me.fraReport.Caption = LoadResString(gLangOffSet + 288)
Me.optReports(0).Caption = LoadResString(gLangOffSet + 61)
Me.optReports(1).Caption = LoadResString(gLangOffSet + 67)
Me.optReports(2).Caption = LoadResString(gLangOffSet + 62)
Me.optReports(3).Caption = LoadResString(gLangOffSet + 403) & " && " & LoadResString(gLangOffSet + 404) + " " + LoadResString(gLangOffSet + 28)
Me.optReports(4).Caption = LoadResString(gLangOffSet + 65)
Me.optReports(5).Caption = LoadResString(gLangOffSet + 63)
Me.lblDate1.Caption = LoadResString(gLangOffSet + 109)
Me.lblDate2.Caption = LoadResString(gLangOffSet + 110)
Me.lblAmt1.Caption = LoadResString(gLangOffSet + 107)
Me.lblAmt2.Caption = LoadResString(gLangOffSet + 108)
Me.cmdView.Caption = LoadResString(gLangOffSet + 13)
End Sub




Private Sub UpdateRF(AccId As Long, Transdate As String)
Dim TempBalance As Currency
Dim TransType As wisTransactionTypes
Dim TransID As Long
Dim Balance As Currency

    gDBTrans.SQLStmt = "SELECT TransType, Amount, Balance, TransID FROM AccTrans" & _
                " WHERE AccID = " & AccId & _
                " ORDER BY TransID"
    If gDBTrans.SQLFetch > 0 Then
        With gDBTrans
            .Rst.MoveFirst
            TransID = FormatField(.Rst("TransID"))
            If TransID = 100 Then
               TempBalance = Val(FormatField(.Rst("Amount")))
               If Val(FormatField(.Rst("Balance"))) <> TempBalance Then
                    gDBTrans.BeginTrans
                    gDBTrans.SQLStmt = "UPDATE AccTrans Set Balance = " & TempBalance & " WHERE AccID = " & _
                        AccId & " And TransID = " & TransID
                    If Not gDBTrans.SQLExecute Then
                        MsgBox "The TransID " & TransID & " Could Not Update", vbCritical, "Error in Updating"
                    End If
                    gDBTrans.CommitTrans
               End If
            Else
               TempBalance = Val(FormatField(.Rst("Balance")))
            End If
            .Rst.MoveNext
            While Not .Rst.EOF
                TransID = FormatField(.Rst("transid"))
                TransType = FormatField(.Rst("TransType"))
                If TransType = wDeposit Or TransType = wContraDeposit Then
                    Balance = TempBalance + FormatField(.Rst("Amount"))
                Else
                    Balance = TempBalance - FormatField(.Rst("Amount"))
                End If
                If Balance <> FormatField(.Rst("Balance")) Then
                    gDBTrans.BeginTrans
                    gDBTrans.SQLStmt = "UPDATE AccTrans Set Balance = " & Balance & " WHERE AccID = " & _
                        AccId & " And TransID = " & TransID
                    If Not gDBTrans.SQLExecute Then
                        MsgBox "The TransID " & TransID & " Could Not Update", vbCritical, "Error in Updating"
                    End If
                    gDBTrans.CommitTrans
                End If
                TempBalance = FormatField(.Rst("Balance"))
                .Rst.MoveNext
            Wend
        End With
    End If

End Sub

Private Sub chkBackLog_Click()
If chkBackLog.value = vbChecked Then
    cmdUndo.Caption = LoadResString(gLangOffSet + 31) 'undo first
End If
End Sub

Private Sub cmbAccHeads_Click()
    
    '**************************************
    ' WARNING
    ' DO not delete any commented lines in ths module
    ' *************
    
    If m_AccHead = cmbAccHeads.Text Then Exit Sub
    m_AccHead = cmbAccHeads.Text
    Dim cmbListIndex As wisBankHeads
   'Clear the Combo boxes
    cmbTrans.Clear
    cmbAccNames.Clear
    Dim TransType As wisTransactionTypes
    '
    cmbListIndex = cmbAccHeads.ItemData(cmbAccHeads.ListIndex)
    
   'Load Apprapriate Transaction Type in Trans Combo
    
    'If cmbAccHeads.ListIndex = 0 Then
    If cmbListIndex = wis_BankHead Then                      ' Bank Accounts
        cmbTrans.AddItem LoadResString(gLangOffSet + 481)
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '  "WithDrawn":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 483) '  "Interest Recieved"
        TransType = wCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 484) '  "Charges Paid":
        TransType = wInterest
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
        cmbTrans.AddItem LoadResString(gLangOffSet + 481) & " " & LoadResString(gLangOffSet + 270)
        TransType = wContraWithdraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) & " " & LoadResString(gLangOffSet + 270) '  "WithDrawn":
        TransType = wContraDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_BankLoanHead Then               'Bank Loans Account
        cmbTrans.AddItem LoadResString(gLangOffSet + 485) '  "Loan Recieved":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 486) '  "Repayment":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
            
        cmbTrans.AddItem LoadResString(gLangOffSet + 487) '  "Interest Paid":
        TransType = wInterest
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 488) '  "Amount Recieved":
        TransType = wCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_AdvanceHead Then                  'Advances OR Advance Accounts
        cmbTrans.AddItem LoadResString(gLangOffSet + 488) '  "Amount Recieved":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 489)     '"Amount Returned":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 488) '  "Amount Recieved":
        TransType = wContraDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 489)     '"Amount Returned":
        TransType = wContraWithdraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

    ElseIf cmbListIndex = wis_InvestmentHead Then 'Investments
        cmbTrans.AddItem LoadResString(gLangOffSet + 481) '  "Deposited":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '  "WithDrawn":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 483)   '"Interest Recieved":
        TransType = wCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 483) '  "Interest Recieved"
        TransType = wContraCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 481) '  "Deposited":
        TransType = wContraWithdraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 482)  '  "WithDrawn":
        TransType = wContraDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

    ElseIf cmbListIndex = wis_IncomeHead Then                                   ''Income Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 488) '"Amount Recieved":
        TransType = wCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_ExpenditureHead Then  'Expense Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 489) '  "Amount Paid":
        TransType = wInterest
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_TradingIncomeHead Then          'Trading 'Income Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 488) '  "Amount Recieved":
        TransType = wCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_TradingExpenditureHead Then     'Trading Expense Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 489) '"Amount Paid":
        TransType = wInterest
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
    ElseIf cmbListIndex = wis_ReserveFundHead Then                   ' Balance Sheet Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 481) '"Deposited":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 270) + LoadResString(gLangOffSet + 481) '"Deposited":
        TransType = wContraDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '"WithDrawn":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
        cmbTrans.AddItem LoadResString(gLangOffSet + 388)  'From Last Years Profit
        TransType = wStock
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_ShareCapitalHead Then           '"Share Capital Head
        cmbTrans.AddItem LoadResString(gLangOffSet + 481) '"Deposited":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '"WithDrawn":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 481) '"Deposited":
        TransType = wContraDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 482)  '"WithDrawn":
        TransType = wContraWithdraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
    ElseIf cmbListIndex = wis_GovtLoanSubsidyHead Then         'Subsidy Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 481) '"Deposited":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '"WithDrawn":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 487) '  "Interest Paid":
        TransType = wInterest
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
        cmbTrans.AddItem LoadResString(gLangOffSet + 483)   '"Interest Recieved":
        TransType = wCharges
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

    ElseIf cmbListIndex = wis_PaymentHead Then               ' Payments
        cmbTrans.AddItem LoadResString(gLangOffSet + 271)
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 271) & " " & LoadResString(gLangOffSet + 403) & " " & LoadResString(gLangOffSet + 207)  'From Last Years Profit
        TransType = wStock
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 271)
        TransType = wContraDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 272)
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 270) & " " & LoadResString(gLangOffSet + 272)  '"WithDrawn":
        TransType = wContraWithdraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

    ElseIf cmbListIndex = wis_RepaymentHead Then                ' Receivables
        cmbTrans.AddItem LoadResString(gLangOffSet + 481) ' "Deposited":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '"WithDrawn":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
    
   ElseIf cmbListIndex = wis_AssetHead Then                     'Balance Sheet Heads
        cmbTrans.AddItem LoadResString(gLangOffSet + 271) '"Deposited":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 272) '"WithDrawn":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

        cmbTrans.AddItem LoadResString(gLangOffSet + 175) '"ContraCharges in case of depreciation":
        TransType = wContraInterest
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType

   ElseIf cmbListIndex = wis_MemberDeposits Then
         cmbTrans.AddItem LoadResString(gLangOffSet + 481) '  "Deposited":
        TransType = wDeposit
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
        
        cmbTrans.AddItem LoadResString(gLangOffSet + 482) '  "WithDrawn":
        TransType = wWithDraw
        cmbTrans.ItemData(cmbTrans.NewIndex) = TransType
     End If

cmdAccept.Enabled = False
cmdUndo.Enabled = False

'Load approriate Heads In Sub Head Combo
Dim Lret As Long
Dim Rst As Recordset

gDBTrans.SQLStmt = "Select AccName, AccId From AccMaster Where accId > " & _
                    cmbAccHeads.ItemData(cmbAccHeads.ListIndex) & _
                    " And AccId < " & cmbAccHeads.ItemData(cmbAccHeads.ListIndex) + wis_BankHeadOffSet
                    
If gDBTrans.SQLFetch < 1 Then
    Exit Sub
Else
Set Rst = gDBTrans.Rst.Clone
    While Not Rst.EOF
        cmbAccNames.AddItem FormatField(Rst("AccName"))
        cmbAccNames.ItemData(cmbAccNames.NewIndex) = FormatField(Rst("AccId"))
        Rst.MoveNext
    Wend
End If

End Sub


Private Sub cmbAccNames_Click()
        
    If cmbAccNames.ListIndex <> -1 Then
        cmdAccept.Enabled = True
        cmdUndo.Enabled = True
    Else
        Exit Sub
    End If
    
If m_AccID <> cmbAccNames.ItemData(cmbAccNames.ListIndex) Then
    m_AccID = cmbAccNames.ItemData(cmbAccNames.ListIndex)
    Call AccountLoad(m_AccID)
End If

'''If cmbAccNames.ItemData(cmbAccNames.ListIndex) = 13001 Or _
'''                cmbAccNames.ItemData(cmbAccNames.ListIndex) = 13002 Or _
'''                        cmbAccNames.ItemData(cmbAccNames.ListIndex) = 13003 Then
'''    cmbTrans.ListIndex = -1
'''    cmbTrans.Enabled = False
'''End If

End Sub
Private Sub cmbRepAccHead_Click()
'Load approriate Heads In Sub Head Combo
Dim Lret As Long
Dim Rst As Recordset
cmbRepAccName.Clear
If cmbRepAccHead.ListIndex >= 0 Then
gDBTrans.SQLStmt = "Select AccName, AccId From AccMaster Where accId > " & _
                    cmbRepAccHead.ItemData(cmbRepAccHead.ListIndex) & _
                    " And AccId < " & cmbRepAccHead.ItemData(cmbRepAccHead.ListIndex) + wis_BankHeadOffSet
Else
    Exit Sub
End If
If gDBTrans.SQLFetch < 1 Then
    Exit Sub
Else
Set Rst = gDBTrans.Rst.Clone
    While Not Rst.EOF
        cmbRepAccName.AddItem FormatField(Rst("AccName"))
        cmbRepAccName.ItemData(cmbRepAccName.NewIndex) = FormatField(Rst("AccId"))
        Rst.MoveNext
    Wend
End If

cmdView.Enabled = True
    

'Set the Option Buttons Of Reports
'Initilally Set all optiopn Buttons enabled
Dim Count As Integer
For Count = 0 To optReports.Count - 1
    optReports(Count).Enabled = True
Next
gDBTrans.SQLStmt = "Select Sum(Amount)  ,TransType From Acctrans Where accId > " & _
                    cmbRepAccHead.ItemData(cmbRepAccHead.ListIndex) & _
                    " And AccId < " & cmbRepAccHead.ItemData(cmbRepAccHead.ListIndex) + wis_BankHeadOffSet & _
                    "  Group By TransType"
If gDBTrans.SQLFetch < 1 Then
    Exit Sub
End If

Dim TransType As wisTransactionTypes
Set Rst = gDBTrans.Rst.Clone
    
    For Count = 1 To Rst.RecordCount
        TransType = Val(FormatField(Rst("TransType")))
        If TransType = wCharges Or TransType = wInterest Then
            optReports(3).Enabled = True
            optReports(5).Enabled = True
        ElseIf TransType = wDeposit Or TransType = wWithDraw Then
            optReports(0).Enabled = True
            optReports(1).Enabled = True
            optReports(2).Enabled = True
            optReports(5).Enabled = True
        End If
    Next

End Sub

Private Sub cmbRepAccName_Click()
If cmbRepAccName.ListIndex >= 0 Then
    cmdView.Enabled = True
End If
End Sub


Private Sub cmbTrans_GotFocus()
cmbTrans.ListIndex = 0
End Sub

Private Sub cmdAccept_Click()
'Check and perform appropriate transaction
    If chkBackLog.value = vbChecked Then
        If Not AccountBackLogTransaction() Then
            Exit Sub
        End If
    Else
        If Not AccountTransaction() Then
            Exit Sub
        End If
    End If

'Reload the account
    If Not AccountLoad(m_AccID) Then
        Exit Sub
    End If
    
    
    If txtDate.Enabled Then
        txtDate.SetFocus
        txtDate.SelStart = 0: txtDate.SelLength = 2 '''Len(txtDate.Text)
    End If
End Sub
Private Function AccountTransaction() As Boolean
Dim AccountCloseFlag As Boolean
Dim L_AccID As wisBankHeads
'Prelim check
    If m_AccID <= 0 Then
        'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 523), vbCritical, gAppName & " - Error"
        cmdUndo.Enabled = False
        Exit Function
    End If

L_AccID = m_AccID

'Check if account exists
    Dim ClosedON As String
    If Not AccountExists(m_AccID, ClosedON) Then
        'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    If ClosedON <> "" Then
        'MsgBox "This account has been closed !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 524), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

'Validate the date and assign to variable
    Dim Transdate As String
    If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
        'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    Else
        Transdate = txtDate.Text
    End If

'Check if the date of transaction is earlier than account opening date itself
    Dim Ret As Integer
    gDBTrans.SQLStmt = "Select * from AccMaster where AccID = " & m_AccID
    Ret = gDBTrans.SQLFetch
    If Ret <> 1 Then
        'MsgBox "DB error !", vbCritical, gAppName & " - ERROR"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - ERROR"
        Exit Function
    Else
        If WisDateDiff(Trim$(txtDate.Text), FormatField(gDBTrans.Rst("CreateDate"))) > 0 Then
            'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
        '    MsgBox LoadResString(gLangOffSet + 568), vbExclamation, gAppName & " - Error"
         '   ActivateTextBox txtDate
          '  Exit Function
        End If
    End If

'Get the Transaction Type
    Dim Trans As wisTransactionTypes
    If cmbTrans.ListIndex = -1 Then
        'MsgBox "Transaction type not specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 588), vbExclamation, gAppName & " - Error"
        If cmbTrans.Enabled Then
            cmbTrans.SetFocus
        End If
        Exit Function
    Else
        Trans = cmbTrans.ItemData(cmbTrans.ListIndex)
    End If

'Validate the Amount
    Dim Amount As Currency
    If Not CurrencyValidate(txtAmount.Text, True) Then
        'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtAmount
        Exit Function
    Else
        Amount = CCur(Trim$(txtAmount.Text))
    End If

'Validate the Cheque No
    Dim ChequeNo As Long
    Dim ChequeStr As String

If Trans = wDeposit Then
    If txtCheque.Text = "" Then
        'MsgBox "Cheque / Scroll number not specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 511), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtCheque
    '    Exit Function
    Else
        ChequeNo = Val(txtCheque.Text)
        If Val(ChequeNo) <= 0 Then
            'MsgBox "Invalid Cheque / Scroll number specified !", vbExclamation, gAppName & " - Error"
            MsgBox LoadResString(gLangOffSet + 511), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtCheque
            Exit Function
        End If
    End If
End If
        ChequeNo = Val(txtCheque.Text)

'Get the Particulars
    Dim Particulars As String
    If Trim$(cmbParticulars.Text) = "" Then
        'MsgBox "Transaction particulars not specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 621), vbExclamation, gAppName & " - Error"
        cmbParticulars.SetFocus
        Exit Function
    Else
        Particulars = Trim$(cmbParticulars.Text)
    End If

'Get the Balance and new transid
    Dim Balance As Currency
    Dim TransID As Long
    gDBTrans.SQLStmt = "Select TOP 1 * from AccTrans where AccID = " & m_AccID & " order by TransID desc"
    If gDBTrans.SQLFetch = 0 Then
        TransID = 100
        Balance = CCur(Val((InputBox("Please enter a balance to start with as this account has not transaction performed", "Initial Balance", "0.00"))))
        If Balance = 0 Then
            'MsgBox "Invalid initial balance specified !", vbExclamation, gAppName & " - Error"
            If MsgBox(LoadResString(gLangOffSet + 517) & vbCrLf & LoadResString(gLangOffSet + 541), vbExclamation + vbYesNo, gAppName & " - Error") = vbNo Then Exit Function
         '   Exit Function
        End If
    Else
        Balance = CCur(FormatField(gDBTrans.Rst("Balance")))
        TransID = FormatField(gDBTrans.Rst("TransID")) + 1

Dim InsertReserveFund As Boolean
        'See if the date is earlier than last date of transaction
        If WisDateDiff(FormatField(gDBTrans.Rst("TransDate")), txtDate.Text) < 0 _
         Or InStr(1, txtDate.Text, "31/3/", vbTextCompare) <> 0 Then
            If Not (m_AccID > wis_ReserveFundHead And _
                m_AccID < wis_ReserveFundHead + wis_BankHead) And _
                    Not InStr(1, txtDate.Text, "31/3/", vbTextCompare) Then
                        'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
                        MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
                        ActivateTextBox txtDate
                        Exit Function
                ElseIf Not Trans = wContraDeposit Then
                        MsgBox "Only Contra Operations are allowed before date!", vbExclamation, "Reserve Fund Operation"
                        cmbTrans.SetFocus
                        Exit Function
                Else
                    InsertReserveFund = True
            End If
        End If
    End If

'Calculate new balance
Select Case L_AccID
    Case wis_BankHead To wis_BankHead + wis_BankHeadOffSet, _
            wis_InvestmentHead To wis_InvestmentHead + wis_BankHeadOffSet, _
            wis_AssetHead To wis_AssetHead + wis_BankHeadOffSet, _
            wis_RepaymentHead To wis_RepaymentHead + wis_BankHead
        If Trans = wDeposit Or Trans = wContraInterest Then
            Balance = CCur(Balance - Amount)
        ElseIf Trans = wWithDraw Or Trans = wCharges Then
            Balance = CCur(Balance + Amount)
        End If
        If Not (L_AccID > wis_RepaymentHead And L_AccID < wis_RepaymentHead + wis_BankHead) Then _
            If Balance < 0 Then Exit Function
    Case wis_BankLoanHead To wis_BankLoanHead + wis_BankHeadOffSet, _
            wis_AdvanceHead To wis_AdvanceHead + wis_BankHeadOffSet, _
            wis_ReserveFundHead To wis_ReserveFundHead + wis_BankHeadOffSet, _
            wis_ShareCapitalHead To wis_ShareCapitalHead + wis_BankHeadOffSet, _
            wis_GovtLoanSubsidyHead To wis_GovtLoanSubsidyHead + wis_BankHeadOffSet, _
            wis_PaymentHead To wis_PaymentHead + wis_BankHeadOffSet
        If Trans = wDeposit Or Trans = wContraDeposit Or Trans = wStock Then
            Balance = CCur(Balance + Amount)
        ElseIf Trans = wWithDraw Then
            Balance = CCur(Balance - Amount)
        End If
        If Not (L_AccID > wis_AdvanceHead And L_AccID < wis_AdvanceHead + wis_BankHead) Then
            If Balance < 0 Then Exit Function
        End If
    Case wis_IncomeHead To wis_IncomeHead + wis_BankHeadOffSet * 4
        Balance = Balance + Amount
End Select

If InsertReserveFund Then
    TransID = GetTransID(m_AccID, Transdate)
End If

'Perform the Transaction to the Database

    gDBTrans.BeginTrans
    
    gDBTrans.SQLStmt = "Insert into AccTrans (AccID, TransID, TransDate, Amount, " & _
                        " Balance, Particulars, TransType, ChequeNo ) values ( " & _
                        m_AccID & "," & _
                        TransID & "," & _
                        "#" & FormatDate(Transdate) & "#," & _
                        Amount & "," & _
                        Balance & "," & "'" & Particulars & "'," & _
                        Trans & "," & _
                        ChequeNo & ")"
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.CommitTrans
    
    If InsertReserveFund Then
        Call UpdateRF(m_AccID, Transdate)
        Balance = OBOfAccount(wis_ProfitOrLoss, Transdate)
        Call StoreBalances(wis_ProfitOrLoss, FormatDate(Transdate), Balance - Amount)
    End If

    If Trans = wStock Then
        Balance = OBOfAccount(wis_FromProfit, Transdate)
        If Balance <= 0 Then
            Balance = OBOfAccount(wis_ProfitOrLoss, Transdate)
        End If
        Call StoreBalances(wis_FromProfit, FormatDate(Transdate), Balance - Amount)
    End If
    Dim i As Long
    If Trans = wWithDraw Then
        'Update the Cheque book
        Dim ChequeArr() As Long
        'Dim i As Integer
        ReDim ChequeArr(0)
    End If
'Update the Particulars combo
    'Read to part array
    Dim ParticularsArr() As String
    ReDim ParticularsArr(20)
    
    'Read elements of combo to array
    For i = 0 To cmbParticulars.ListCount - 1
        ParticularsArr(i) = cmbParticulars.List(i)
    Next i
    
    'Update last accessed elements
    Call UpdateLastAccessedElements(Trim$(cmbParticulars.Text), ParticularsArr, True)
    
    'Write to
    cmbParticulars.Clear
    For i = 0 To UBound(ParticularsArr)
        If Trim$(ParticularsArr(i)) <> "" Then
            Call WriteToIniFile("Particulars", "Key" & i, ParticularsArr(i), App.Path & "\BankAcc.ini")
            cmbParticulars.AddItem ParticularsArr(i)
        End If
    Next i

If AccountCloseFlag = True Then
    If Not AccountClose() Then
        'MsgBox "Unable to close the account !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 534), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
End If

AccountTransaction = True
End Function


Private Function GetTransID(AccId As Long, Transdate As String) As Long
Dim TransID As Long
Dim TransRecord As Recordset
Dim TempTransID As Long

gDBTrans.SQLStmt = "SELECT TransID, TransDate FROM AccTrans WHERE AccID = " & AccId & _
        " AND TransDate >= #" & Transdate & "# ORDER BY TransDate, TransID"
If gDBTrans.SQLFetch > 0 Then
    Set TransRecord = gDBTrans.Rst.Clone
    TransRecord.MoveFirst
    TempTransID = TransRecord.fields("TransID")
    TransRecord.MoveLast
    While TransID <> TempTransID
        TransID = TransRecord.fields("TransID")
        gDBTrans.BeginTrans
        gDBTrans.SQLStmt = "UPDATE AccTrans SET TransID = " & TransID + 1 & _
                " WHERE AccID = " & AccId & _
                " And TransID = " & TransID
        If Not gDBTrans.SQLExecute Then
            MsgBox "The Account No. " & AccId & " Could Not Update", vbCritical, "Error in Updating"
            gDBTrans.RollBack
        End If
        gDBTrans.CommitTrans
        TransRecord.MovePrevious
    Wend
End If
Set TransRecord = Nothing
If TempTransID = 0 Then TempTransID = 101
GetTransID = TempTransID

End Function

Private Function AccountClose() As Boolean
'exit
Dim Ret As Integer
Dim AccNo As Long

'Prelim checks
    AccNo = m_AccID
    If AccNo <= 0 Then Exit Function

'Check if account exists
    If Not AccountExists(AccNo) Then
        Exit Function
    End If
    
'Check date format
    If Not DateValidate(txtDate.Text, "/", True) Then
        'MsgBox "Invalid date specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 501), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Close the account
    gDBTrans.BeginTrans
        gDBTrans.SQLStmt = "Update AccMaster set ClosedDate = #" & FormatDate(Trim$(txtDate.Text)) & "# where AccID = " & AccNo
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Function
        End If
    gDBTrans.CommitTrans
    
AccountClose = True

End Function


Public Function AccountExists(AccId As Long, Optional ClosedON As String) As Boolean
Dim Ret As Integer

'Query Database
    gDBTrans.SQLStmt = "Select * from AccMaster where " & _
                        " AccID = " & AccId
    Ret = gDBTrans.SQLFetch
    If Ret <= 0 Then Exit Function
    
    If Ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

'Check the closed status
    If Not IsMissing(ClosedON) Then
        'ClosedON = FormatField(gDBTrans.Rst("ClosedDate"))
    End If
AccountExists = True
End Function

Private Function AccountName(AccId As Long) As String

Dim Lret As Long

'Prelim checks
    If AccId <= 0 Then Exit Function

'Query DB
        gDBTrans.SQLStmt = "SELECT AccID, Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM CAmaster, " _
                & "NameTab WHERE CAmaster.AccID = " & AccId _
                & " AND CAmaster.CustomerID = NameTab.CustomerID"
        'Lret = gDBTrans.SQLFetch
        If Lret = 1 Then
            AccountName = FormatField(gDBTrans.Rst("Name"))
        ElseIf Lret > 1 Then
            'MsgBox "Data base error !", vbCritical, gAppName & " - Error"
            MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - Error"
            Exit Function
        End If

End Function




Private Function AccountReopen(AccNo As Long) As Boolean

'Check if account number exists in data base
    gDBTrans.SQLStmt = "Select * from AccMaster where AccID = " & AccNo
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If


    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "Update AccMaster set ClosedDate = NULL where AccID = " & AccNo
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.CommitTrans
    
AccountReopen = True
End Function


'****************************************************************************************
'Returns a new account number
'Author: Girish
'Date : 29th Dec, 1999
'Modified by Ravindra on 25th Jan, 2000
'****************************************************************************************
Private Function GetNewAccountNumber() As Long
    Dim NewAccNo As Long
    'gDBTrans.SQLStmt = "Select TOP 1 AccID from CAmaster order by AccID desc"
    gDBTrans.SQLStmt = "SELECT MAX(AccID) FROM AccMaster"
    If gDBTrans.SQLFetch = 0 Then
        NewAccNo = 1
    Else
        NewAccNo = Val(FormatField(gDBTrans.Rst(0))) + 1
    End If
    GetNewAccountNumber = NewAccNo
End Function

Public Function AccountLoad(AccId As Long) As Boolean
Dim rstMaster As Recordset
Dim ClosedDate As String
Dim Ret As Integer
Dim JointHolders() As String
Dim i As Integer

'Check if account number is valid
    If AccId <= 0 Then GoTo DisableUserInterface
'Check if account number exists
    If Not AccountExists(AccId) Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        GoTo DisableUserInterface
    End If

'Query data base
    gDBTrans.SQLStmt = "Select * from AccMaster where AccID = " & AccId
    If gDBTrans.SQLFetch <= 0 Then
        GoTo DisableUserInterface
    Else  'Set record set to local rec set
        Set rstMaster = gDBTrans.Rst.Clone
    End If

    
'Get the transaction details of this account holder
    gDBTrans.SQLStmt = "Select * from AccTrans where AccID = " & AccId & " order by TransID"
    Ret = gDBTrans.SQLFetch
    If Ret < 0 Then
        GoTo DisableUserInterface
    ElseIf Ret > 0 Then
        Dim BalanceAmount As Currency
        Set m_rstPassBook = gDBTrans.Rst.Clone
        m_rstPassBook.MoveLast
        BalanceAmount = CCur(m_rstPassBook("Balance"))
        
        'Position to first record of last page
        With m_rstPassBook
            .Move -1 * (.AbsolutePosition Mod 10)
        End With
        cmdUndo.Enabled = True
        'cmdDelete.Enabled = False       'There are transactions, Do not Allow delete
    Else
        Set m_rstPassBook = Nothing
        PassBookPageInitialize
        cmdUndo.Enabled = False
        'cmdDelete.Enabled = True        'No transactions, Allow delete
    End If
    
'Assign to some module level variables
    m_AccID = AccId
    'm_accUpdatemode = wis_UPDATE
    'm_AccClosed = IIf(FormatField(rstMaster("ClosedDate")) <> "", True, False)
    
'Load account to the User Interface
    'TAB 1
    'ClosedDate = FormatField(rstMaster("ClosedDate"))
    With Me
        With .txtDate
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            If .Text = "" Then
                .Text = FormatDate(gStrDate)
            End If
        End With
        
        With .cmbTrans
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .ListIndex = -1
        End With

        With .cmbParticulars
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .ListIndex = -1
        End With
        
        With .txtAmount
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Text = ""
        End With
        
        With .txtCheque
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Text = ""
        End With
        
'        cmdAddNote.Enabled = IIf(ClosedDate = "", True, False)
        cmdPrevTrans.Enabled = IIf(ClosedDate = "", True, False)
        cmdNextTrans.Enabled = IIf(ClosedDate = "", True, False)
        
        .cmdAccept.Enabled = IIf(ClosedDate = "", True, False)
        '.cmdUndo.Enabled = IIf(ClosedDate = "", True, False)
        .chkBackLog.Enabled = IIf(ClosedDate = "", True, False)
        
        Call PassBookPageShow
    End With
        
AccountLoad = True

Exit Function

DisableUserInterface:
    Call ResetUserInterface
    

Exit Function
    
ErrLine:
'MsgBox "Account Load:" & vbCrLf & "     Error Loading account", vbCritical, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 521) & vbCrLf & "     Error Loading account", vbCritical, gAppName & " - Error"
End Function


Private Function AccountBackLogTransaction() As Boolean
Dim i As Integer

'Prelim check
    If m_AccID <= 0 Then
        'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 523), vbCritical, gAppName & " - Error"
        cmdUndo.Enabled = False
        Exit Function
    End If

'Check if account exists
    Dim ClosedON As String
    If Not AccountExists(m_AccID, ClosedON) Then
        'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    If ClosedON <> "" Then
        'MsgBox "This account has been closed !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 524), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

'Validate the date and assign to variable
    Dim Transdate As String
    If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
        'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    Else
        Transdate = FormatDate(txtDate.Text)
    End If

'Check if the date of transaction is earlier than account opening date itself
    Dim Ret As Integer
    gDBTrans.SQLStmt = "Select * from AccMaster where AccID = " & m_AccID
    Ret = gDBTrans.SQLFetch
    If Ret <> 1 Then
        'MsgBox "DB error !", vbCritical, gAppName & " - ERROR"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - ERROR"
        Exit Function
    Else
        If WisDateDiff(Trim$(txtDate.Text), FormatField(gDBTrans.Rst("CreateDate"))) > 0 Then
            'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
            MsgBox LoadResString(gLangOffSet + 568), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtDate
            Exit Function
        End If
    End If

'Validate the Amount
    Dim Amount As Currency
    If Not CurrencyValidate(txtAmount.Text, True) Then
        'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtAmount
        Exit Function
    Else
        Amount = CCur(Trim$(txtAmount.Text))
    End If

'Get the Transaction Type
    Dim Trans As wisTransactionTypes
    If cmbTrans.ListIndex = -1 Then
        'MsgBox "Transaction type not specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 588), vbExclamation, gAppName & " - Error"
        cmbTrans.SetFocus
        Exit Function
    Else
       Trans = cmbTrans.ItemData(cmbTrans.ListIndex)
    End If

'Validate the Cheque No
    Dim ChequeNo As Long
    Dim ChequeStr As String
    
'Check out the cheque book. He has to enter a number explicitly and not from the existing book
    If txtCheque.Text = "" Then
        'MsgBox "Cheque / Scroll number not specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 511), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtCheque
        Exit Function
    Else
        ChequeNo = Val(txtCheque.Text)
        If Val(ChequeNo) <= 0 Then
            'MsgBox "Invalid Cheque / Scroll number specified !", vbExclamation, gAppName & " - Error"
            MsgBox LoadResString(gLangOffSet + 511), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtCheque
            Exit Function
        End If
    End If
'End If
    
'Get the Particulars
    Dim Particulars As String
    If Trim$(cmbParticulars.Text) = "" Then
        'MsgBox "Transaction particulars not specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 621), vbExclamation, gAppName & " - Error"
        cmbParticulars.SetFocus
        Exit Function
    Else
        Particulars = Trim$(cmbParticulars.Text)
    End If

'Get the Balance
    Dim Balance As Currency
    Dim TransID As Long
    Dim TransType As wisTransactionTypes
    
    gDBTrans.SQLStmt = "Select TOP 1 * from AccTrans where AccID = " & m_AccID & " order by TransID"
    If gDBTrans.SQLFetch = 0 Then
        'MsgBox "You can clear the back log only if you have some transaction on the account !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 554), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
        
'Get Balance and TransID for back log
    Dim OLDTransType As wisTransactionTypes
    Balance = CCur(FormatField(gDBTrans.Rst("Balance")))
    TransID = FormatField(gDBTrans.Rst("TransID")) - 1  'Back Log clearance
    OLDTransType = FormatField(gDBTrans.Rst("TransType"))

'Calculate the OLD Balance
    If OLDTransType = wWithDraw Or OLDTransType = wCharges Then
        Balance = CCur(Balance + Val(FormatField(gDBTrans.Rst("Amount"))))
    ElseIf OLDTransType = wDeposit Or OLDTransType = wInterest Then
        Balance = CCur(Balance - Val(FormatField(gDBTrans.Rst("Amount"))))
    End If
        
'See if the later is earlier than last date of transaction
    If WisDateDiff(FormatField(gDBTrans.Rst("TransDate")), txtDate.Text) < 0 Then
        'MsgBox "You have specified a transaction date that is later than the first date of transaction !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    End If

'Perform the Transaction to the Database
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "Insert into AccTrans (AccID, TransID, TransDate, Amount, " & _
                        " Balance, Particulars, TransType, ChequeNo ) values ( " & _
                        m_AccID & "," & _
                        TransID & "," & _
                        "#" & Transdate & "#," & _
                        Amount & "," & _
                        Balance & "," & "'" & Particulars & "'," & _
                        Trans & "," & _
                        ChequeNo & ")"
                    
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
    
    gDBTrans.CommitTrans

'Update the Particulars combo

    For i = 0 To cmbParticulars.ListCount - 1
        If UCase(cmbParticulars.List(i)) = UCase(cmbParticulars.Text) Then
            Exit For
        End If
    Next i

    If i = cmbParticulars.ListCount Then
        Call WriteToIniFile("Particulars", "Key" & i, cmbParticulars.Text, App.Path & "\BankAcc.ini")
    End If
        
'Update the Particulars combo
    i = 0
    cmbParticulars.Clear
    Do
        Particulars = ReadFromIniFile("Particulars", "Key" & i, App.Path & "\BankAcc.INI")
        If Particulars <> "" Then
            cmbParticulars.AddItem Particulars
        End If
        i = i + 1
    Loop Until Particulars = ""

AccountBackLogTransaction = True

End Function


Private Sub cmdAccNames_Click()

If cmbAccHeads.ListIndex < 0 Then Exit Sub

Dim HeadId As wisBankHeads
'First Check whether this name is in the DataBase or not
'If itis not there insert it into this Database
gDBTrans.SQLStmt = "Select * From AccMaster Where AccName = '" & m_AccHead & "' "
If gDBTrans.SQLFetch < 1 Then 'Insert This Into the DataBase
    HeadId = cmbAccHeads.ItemData(cmbAccHeads.ListIndex)
    
    gDBTrans.SQLStmt = "Insert Into AccMaster (AccId,AccName,CreateDate) Values " & _
                        " ( " & HeadId & ", '" & cmbAccHeads.Text & "', #1/1/1975# )"
    
    gDBTrans.BeginTrans
    If Not gDBTrans.SQLExecute Then
        'MsgBox "DataBase Corroption"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " -ERROR"
        gDBTrans.RollBack
        Exit Sub
    End If
    'If These heads are income & Expense Heads Then Insert Sub Head called 'Miscelleneous'
    If HeadId >= 5000 And HeadId <= 8000 Then
    gDBTrans.SQLStmt = "Insert Into accMaster (AccId,AccName,CreateDate) Values " & _
                        " ( " & HeadId + 1 & ", '" & LoadResString(gLangOffSet + 327) & "', #1/1/1975#  )"
        If Not gDBTrans.SQLExecute Then
            'MsgBox "DataBase Corroption"
            MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " -ERROR"
            gDBTrans.RollBack
            Exit Sub
        End If
            
    End If
    
    'If this Head is Expense then Add On more Sub head as 'Pigmy Commission'
    If HeadId = 6000 Then  'ExpenseHead
        gDBTrans.SQLStmt = "Insert Into accMaster (AccId,AccName,CreateDate) Values " & _
                        " ( " & HeadId + 2 & ", '" & LoadResString(gLangOffSet + 328) & "', #1/1/1975#  )"
        If Not gDBTrans.SQLExecute Then
            'MsgBox "DataBase Corroption"
            MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - ERROR"
            gDBTrans.RollBack
            Exit Sub
        End If
    End If
    
    'If this Head is Other Payables then Add One more Sub head as 'Interest Payable'
    If HeadId = 13000 Then  'Other Payables
'Fixed Deposit Payable
        gDBTrans.SQLStmt = "Insert Into AccMaster (AccId,AccName,CreateDate) Values " & _
                        " ( " & HeadId + 1 & ", '" & LoadResString(gLangOffSet + 423) & _
                        " " & LoadResString(gLangOffSet + 450) & "', #1/1/1975#  )"
        If Not gDBTrans.SQLExecute Then
            'MsgBox "DataBase Corroption"
            MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " -ERROR"
            gDBTrans.RollBack
            Exit Sub
        End If
' Pigmy Deposit Payable
        gDBTrans.SQLStmt = "Insert Into AccMaster (AccId,AccName,CreateDate) Values " & _
                        " ( " & HeadId + 2 & ", '" & LoadResString(gLangOffSet + 425) & _
                        " " & LoadResString(gLangOffSet + 450) & "', #1/1/1975#  )"
        If Not gDBTrans.SQLExecute Then
            'MsgBox "DataBase Corroption"
            MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " -ERROR"
            gDBTrans.RollBack
            Exit Sub
        End If
        
' Recurring Deposit Payable
        gDBTrans.SQLStmt = "Insert Into AccMaster (AccId,AccName,CreateDate) Values " & _
                        " ( " & HeadId + 3 & ", '" & LoadResString(gLangOffSet + 424) & _
                        " " & LoadResString(gLangOffSet + 450) & "', #1/1/1975#  )"
        If Not gDBTrans.SQLExecute Then
            'MsgBox "DataBase Corroption"
            MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " -ERROR"
            gDBTrans.RollBack
            Exit Sub
        End If

    End If
    gDBTrans.CommitTrans

End If
'Check For the Date of Craetation
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Invalid Date Specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Sub
End If

'First assign the headid to the calling form
frmAcDef.p_HeadId = cmbAccHeads.ItemData(cmbAccHeads.ListIndex)
frmAcDef.txtAccHead.Text = m_AccHead
frmAcDef.txtAccName.Text = cmbAccNames.Text
frmAcDef.Show vbModal, Me

'Reflect the Changes in Combo Box
m_AccHead = ""
Call cmbAccHeads_Click
Call LoadReportAccountHeads
End Sub


Private Sub ResetUserInterface()

'First the TAB 1
    'Disable the UI if you are unable to load the specified account number
    With cmbAccNames
        .BackColor = wisGray: .Enabled = False: .Clear
    End With
    With txtDate
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With txtAmount
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With cmbTrans
        .BackColor = wisGray: .Enabled = False
    End With
    With txtCheque
        .BackColor = wisGray: .Enabled = False
    End With
    With cmbParticulars
        .BackColor = wisGray: .Enabled = False
    End With
    With cmdAccept
        .Enabled = False
    End With
    With cmdUndo
        .Enabled = False
    End With
    With chkBackLog
        .Enabled = False
    End With

    Call PassBookPageInitialize
    
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    
    
'Now the Tab 2
    Dim i As Integer
    Dim strField As String
    Dim TxtIndex As Integer
    
    
'The form level variables
    m_CustReg.NewCustomer
    m_AccID = 0
    Set m_rstPassBook = Nothing
End Sub

Private Function PassBookPageInitialize()
    With grd
        .Clear: .Rows = 11: .FixedRows = 1: .FixedCols = 0
        .Row = 0
        If cmbAccHeads.ListIndex = 1 Then
            .Cols = 7
            .Col = 0: .Text = LoadResString(gLangOffSet + 37): .CellFontBold = True: .ColWidth(0) = 900
            .Col = 1: .Text = LoadResString(gLangOffSet + 39): .CellFontBold = True: .ColWidth(1) = 500
            .Col = 2: .Text = LoadResString(gLangOffSet + 41): .CellFontBold = True: .ColWidth(2) = 500
            .Col = 3: .Text = LoadResString(gLangOffSet + 276): .CellFontBold = True: .ColWidth(3) = 1000
            .Col = 4: .Text = LoadResString(gLangOffSet + 277): .CellFontBold = True: .ColWidth(4) = 900
            .Col = 5: .Text = LoadResString(gLangOffSet + 47): .CellFontBold = True: .ColWidth(5) = 900
            .Col = 6: .Text = LoadResString(gLangOffSet + 42): .CellFontBold = True: .ColWidth(6) = 1000
        Else
            .Cols = 6
            .Col = 0: .Text = LoadResString(gLangOffSet + 37): .CellFontBold = True: .ColWidth(0) = 1000
            .Col = 1: .Text = LoadResString(gLangOffSet + 39): .CellFontBold = True: .ColWidth(1) = 900
            .Col = 2: .Text = LoadResString(gLangOffSet + 41): .CellFontBold = True: .ColWidth(2) = 800
            .Col = 3: .Text = LoadResString(gLangOffSet + 276): .CellFontBold = True: .ColWidth(3) = 1000
            .Col = 4: .Text = LoadResString(gLangOffSet + 277): .CellFontBold = True: .ColWidth(4) = 1000
            .Col = 5: .Text = LoadResString(gLangOffSet + 42): .CellFontBold = True: .ColWidth(5) = 1000
        End If
    End With
End Function

Private Sub PassBookPageShow()
Dim i As Integer
Dim CmbIndex As Long
Dim TransType As wisTransactionTypes
'Check if Rec Set has been set
    If m_rstPassBook Is Nothing Then
        Exit Sub
    End If

   CmbIndex = Me.cmbAccHeads.ItemData(cmbAccHeads.ListIndex)
'Show 10 records or till eof of the page being pointed to
    With grd
        Call PassBookPageInitialize
        .Visible = False
        i = 0
        Dim BankTrans As wisBankHeads
        BankTrans = wis_AdvanceHead
        Do
            i = i + 1
            .Row = i
            .Col = 0: .Text = FormatField(m_rstPassBook("TransDate"))
            .Col = 1: .Text = FormatField(m_rstPassBook("Particulars"))
            .Col = 2: .Text = " " & FormatField(m_rstPassBook("ChequeNo"))
            TransType = FormatField(m_rstPassBook("TransType"))
            If cmbAccHeads.ListIndex = 1 Then
                If TransType = wWithDraw Then
                    .Col = 4
                ElseIf TransType = wDeposit Then
                    .Col = 3
                ElseIf TransType = wInterest Then
                    .Col = 5
                End If
            ElseIf CmbIndex = wis_ReserveFundHead Or CmbIndex = wis_AdvanceHead _
               Or CmbIndex = wis_ShareCapitalHead Then
                If TransType = wWithDraw Then
                    .Col = 4
                Else
                    .Col = 3
                End If
            Else
                If TransType = wWithDraw Or TransType = wCharges Or _
                    TransType = wContraWithdraw Or TransType = wContraCharges Or TransType = wStock Then
                    .Col = 3
                ElseIf TransType = wDeposit Or TransType = wInterest Or _
                    TransType = wContraDeposit Or TransType = wContraInterest Then
                    .Col = 4
                End If
            End If
                .Text = FormatField(m_rstPassBook("Amount"))
            If cmbAccHeads.ListIndex = 1 Then
                .Col = 6
            Else
                .Col = 5
            End If
            If Me.cmbAccHeads.ItemData(cmbAccHeads.ListIndex) = BankTrans Then
                .Text = FormatCurrency(Abs(FormatField(m_rstPassBook("Balance"))))
            Else
                .Text = FormatCurrency(Abs(FormatField(m_rstPassBook("Balance"))))
            End If
            If i < 10 Then
                m_rstPassBook.MoveNext
                If m_rstPassBook.EOF Then Exit Do
            Else
                Exit Do
            End If
        Loop
        .Visible = True
        .Row = 1
    End With

End Sub
Private Sub cmdCalendar_Click()
With Calendar
    .Left = Me.Left + Me.Fra(1).Left + Me.txtDate.Left
    .Top = Me.Top + Me.Fra(1).Top + Me.txtDate.Top
    If DateValidate(txtDate.Text, "/", True) Then
        .SelDate = txtDate.Text
    Else
        .SelDate = FormatDate(gStrDate)
    End If
    .Show vbModal
    Me.txtDate.Text = .SelDate
End With
End Sub

Private Sub cmdDate1_Click()
With Calendar
    .SelDate = FormatDate(gStrDate)
    If DateValidate(txtDate1.Text, "/", True) Then .SelDate = txtDate1.Text
    .Left = Me.Left + Me.Fra(1).Left + cmdDate1.Left - .Width / 2
    .Top = Me.Top + Me.Fra(2).Top + cmdDate2.Top + 2800
    .Show vbModal
    If .SelDate <> "" Then txtDate1.Text = .SelDate
End With
End Sub

Private Sub cmdDate2_Click()
With Calendar
    .SelDate = FormatDate(gStrDate)
    If DateValidate(txtDate2.Text, "/", True) Then .SelDate = txtDate2.Text
    .Left = Me.Left + Me.Fra(1).Left + cmdDate2.Left - .Width / 2
    .Top = Me.Top + Me.Fra(2).Top + cmdDate2.Top + 2800
    .Show vbModal
    If .SelDate <> "" Then txtDate2.Text = .SelDate
End With
End Sub


Private Sub cmdNextTrans_Click()
If m_rstPassBook Is Nothing Then
    Exit Sub
End If

Dim CurPos As Integer

'Position cursor to start of next page
    If m_rstPassBook.EOF Then
        m_rstPassBook.MoveLast
    End If
    CurPos = m_rstPassBook.AbsolutePosition
    CurPos = 10 - (CurPos Mod 10)
    If m_rstPassBook.AbsolutePosition + CurPos >= m_rstPassBook.RecordCount Then
        Beep
        Exit Sub
    Else
        m_rstPassBook.Move CurPos
    End If

Call PassBookPageShow

#If Junk Then
If m_rstPassBook.AbsolutePosition < m_rstPassBook.RecordCount - 10 Then
    If m_rstPassBook.AbsolutePosition Mod 10 <> 0 Then
        m_rstPassBook.Move 10 - m_rstPassBook.AbsolutePosition Mod 10
        If m_rstPassBook.AbsolutePosition >= m_rstPassBook.RecordCount - 10 Then
            cmdNextTrans.Enabled = False
        End If
    End If
Else
    cmdNextTrans.Enabled = False
End If

Call PassBookPageShow
If m_rstPassBook.AbsolutePosition >= m_rstPassBook.RecordCount Then
    cmdPrevTrans.Enabled = False
Else
    cmdPrevTrans.Enabled = True
End If
#End If


End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPrevTrans_Click()
If m_rstPassBook Is Nothing Then
    Exit Sub
End If

Dim CurPos As Integer

'Position cursor to previous page
    If m_rstPassBook.EOF Then
        'm_rstPassBook.MoveFirst
        m_rstPassBook.MoveLast
        'm_rstPassBook.MovePrevious
    End If
    
    CurPos = m_rstPassBook.AbsolutePosition
    
    CurPos = CurPos - CurPos Mod 10 - 10
    If CurPos < 0 Then
        Beep
        Exit Sub
    Else
        m_rstPassBook.MoveFirst
        m_rstPassBook.Move (CurPos)
    End If
    Call PassBookPageShow
    
#If Junk Then
If m_rstPassBook.AbsolutePosition > 10 Then
    If m_rstPassBook.AbsolutePosition Mod 10 = 0 Then
        'm_rstpassbook.MovePrevious
        m_rstPassBook.Move -1 * (m_rstPassBook.AbsolutePosition Mod 10 + 20)
    Else
        m_rstPassBook.Move -1 * (m_rstPassBook.AbsolutePosition Mod 10 + 10)
    End If
    
    If m_rstPassBook.AbsolutePosition < 10 Then
        cmdPrevTrans.Enabled = False
    End If
End If
Call PassBookPageShow
If m_rstPassBook.AbsolutePosition < 10 Then
    cmdNextTrans.Enabled = False
Else
    cmdNextTrans.Enabled = True
End If
#End If
End Sub




Private Sub cmdUndo_Click()
If chkBackLog.value = vbUnchecked Then
    If Not AccountUndoLastTransaction() Then
        Exit Sub
    End If
End If
If chkBackLog.value = vbChecked Then
    If Not AccountUndoFirstTransaction() Then
        Exit Sub
    End If
End If
If Not AccountLoad(m_AccID) Then
   ' MsgBox "Unable to undo transaction !", vbCritical, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 609), vbCritical, gAppName & " - Error"
    Exit Sub
End If
End Sub

Private Sub cmdView_Click()

'First validate the Data
If txtDate1.Enabled Then
    If Not DateValidate(txtDate1.Text, "/", True) Then
        'MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate1
        Exit Sub
    End If
End If
If txtDate2.Enabled Then
    If Not DateValidate(txtDate2.Text, "/", True) Then
        'MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate2
        Exit Sub
    End If
End If

If txtAmt1.Enabled And Trim(txtAmt1.Text) <> "" Then
    If Not CurrencyValidate(txtAmt1.Text, True) Then
        MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtAmt1
        Exit Sub
    End If
End If

If txtAmt2.Enabled And Trim(txtAmt2.Text) <> "" Then
    If Not CurrencyValidate(txtAmt2.Text, True) Then
        MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtAmt2
        Exit Sub
    End If
End If
gCancel = False
frmCancel.Show
frmCancel.Refresh
Set m_frmBankReport = New frmBankReport
If gCancel Then GoTo ExitLine
Load m_frmBankReport
Unload frmCancel
frmBankReport.Show vbModal

ExitLine:
   On Error Resume Next
   Set frmCancel = Nothing
   Set m_frmBankReport = Nothing



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
      If TabStrip.SelectedItem.Index = TabStrip.Tabs.Count Then
            TabStrip.Tabs(1).Selected = True
      Else
            TabStrip.Tabs(TabStrip.SelectedItem.Index + 1).Selected = True
      End If
End If

End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption
    
'Load The Main Heads in CmbAccHeads
 LoadAccountHeads
    
'load the Account names in combobox
 Call cmbAccNames_Click
'Load the AccountHeads  in to reports Combo Box
 Call LoadReportAccountHeads
        
 txtDate.Text = FormatDate(gStrDate)
    cmbRepAccName.Clear
    optReports(2).value = True
    optReports(0).value = True
    Call TabStrip_Click
Screen.MousePointer = vbDefault
grd.Cols = 3
grd.Rows = 6
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
gWindowHandle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gWindowHandle = 0
    Set frmBankAcc = Nothing
End Sub

Private Sub m_frmBankReport_Initialise(Min As Long, Max As Long)
If Max <> 0 Then
    With frmCancel
        .prg.Min = Min
        .prg.Visible = True
        If Max > 32500 Then Max = 32500
        .prg.Max = Max
    End With
End If
End Sub

Private Sub m_frmBankReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS " & vbCrLf & strMessage
    If Ratio > 0 Then
        .prg.value = Ratio
    End If
End With
End Sub


Private Sub optReports_Click(Index As Integer)
    Select Case Index
        Case 0:
            txtDate1.Enabled = False: txtDate2.Enabled = False
            txtAmt1.Enabled = True: txtAmt2.Enabled = True
            lblDate1.Enabled = False: lblDate2.Enabled = False
            lblAmt1.Enabled = True: lblAmt2.Enabled = True
            cmdDate1.Enabled = False: cmdDate2.Enabled = False
        Case 1:
            txtDate1.Enabled = True: txtDate2.Enabled = False
            txtAmt1.Enabled = False: txtAmt2.Enabled = False
            lblDate1.Enabled = True: lblDate2.Enabled = False
            lblAmt1.Enabled = False: lblAmt2.Enabled = False
            cmdDate1.Enabled = True: cmdDate2.Enabled = False
        Case 2:
            txtDate1.Enabled = True: txtDate2.Enabled = True
            txtAmt1.Enabled = False: txtAmt2.Enabled = False
            lblDate1.Enabled = True: lblDate2.Enabled = True
            lblAmt1.Enabled = False: lblAmt2.Enabled = False
            cmdDate1.Enabled = True: cmdDate2.Enabled = True
        Case 3:
            txtDate1.Enabled = True: txtDate2.Enabled = True
            txtAmt1.Enabled = False: txtAmt2.Enabled = False
            lblDate1.Enabled = True: lblDate2.Enabled = True
            lblAmt1.Enabled = False: lblAmt2.Enabled = False
            cmdDate1.Enabled = True: cmdDate2.Enabled = True
            
        Case 4:
            txtDate1.Enabled = True: txtDate2.Enabled = False
            txtAmt1.Enabled = True: txtAmt2.Enabled = True
            lblDate1.Enabled = True: lblDate2.Enabled = False
            lblAmt1.Enabled = True: lblAmt2.Enabled = True
            cmdDate1.Enabled = True: cmdDate2.Enabled = False
        Case 5:
            txtDate1.Enabled = True: txtDate2.Enabled = True
            txtAmt1.Enabled = False: txtAmt2.Enabled = False
            lblDate1.Enabled = True: lblDate2.Enabled = True
            lblAmt1.Enabled = False: lblAmt2.Enabled = False
            cmdDate1.Enabled = True: cmdDate2.Enabled = True
    End Select
        
End Sub


Private Sub TabStrip_Click()
Dim Count As Byte
For Count = 1 To Fra.Count
    Fra(Count).Visible = False
Next
With TabStrip
    Fra(.SelectedItem.Index).Visible = True
    Fra(.SelectedItem.Index).ZOrder 0
End With


End Sub

Private Sub LoadReportAccountHeads()
'Load The Above Heads Into Report Combo Box
gDBTrans.SQLStmt = "Select * From AccMaster Where AccId mod " & _
                        wis_BankHeadOffSet & "  = 0 ORDER BY AccID"
cmbRepAccHead.Clear
If gDBTrans.SQLFetch > 0 Then
    Dim Rst As Recordset
    Set Rst = gDBTrans.Rst.Clone
    While Not Rst.EOF
        cmbRepAccHead.AddItem FormatField(Rst("AccName"))
        cmbRepAccHead.ItemData(cmbRepAccHead.NewIndex) = Val(FormatField(Rst("AccId")))
        Rst.MoveNext
    Wend
End If

End Sub
Public Function AccountUndoFirstTransaction() As Boolean

'Prelim check
    If m_AccID <= 0 Then
        'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 523), vbCritical, gAppName & " - Error"
        cmdUndo.Enabled = False
        Exit Function
    End If

'Check if account exists
    Dim ClosedON As String
    If Not AccountExists(m_AccID, ClosedON) Then
        'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    
    If ClosedON <> "" Then
        'MsgBox "Account has been closed previously. Ths ", vbExclamation, gAppName & " - Error"
        'If MsgBox("Account has been closed previously." & vbCrLf & _
                "This action will reopen the account." & vbCrLf & _
                "Do you want to continue ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
        If MsgBox(LoadResString(gLangOffSet + 524) & vbCrLf & _
                LoadResString(gLangOffSet + 548) & vbCrLf & _
                LoadResString(gLangOffSet + 541), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
                Exit Function
        Else  'Reopen the account first
            If Not AccountReopen(m_AccID) Then
                'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
                MsgBox LoadResString(gLangOffSet + 536), vbExclamation, gAppName & " - Error"
                Exit Function
            End If
        End If
    End If
    
    'Get first transaction record
    Dim Amount As Currency
    Dim Ret As Integer
    Dim TransID As Long
    gDBTrans.SQLStmt = "Select TOP 1 * from AccTrans where AccID = " & m_AccID & " order by TransID Asc"
    Ret = gDBTrans.SQLFetch
    If Ret >= 1 Then
        Amount = CCur(FormatField(gDBTrans.Rst("Amount")))
        TransID = FormatField(gDBTrans.Rst("TransID"))
    ElseIf Ret = 0 Then
        'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 551), vbInformation, gAppName & " - Error"
        Exit Function
    End If
    
    'Confirm UNDO
    'If MsgBox("Are you sure you want to undo the previous transaction of Rs." & Amount & "?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
    If MsgBox(LoadResString(gLangOffSet + 627) & Amount & "?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
        Exit Function
    End If
    
'Delete record from Data base
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "Delete from AccTrans where AccID = " & m_AccID & " and TransID = " & TransID
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.CommitTrans

AccountUndoFirstTransaction = True
End Function

Public Function AccountUndoLastTransaction() As Boolean

'Prelim check
    If m_AccID <= 0 Then
        'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 523), vbCritical, gAppName & " - Error"
        cmdUndo.Enabled = False
        Exit Function
    End If

'Check if account exists
    Dim ClosedON As String
    If Not AccountExists(m_AccID, ClosedON) Then
        'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    If ClosedON <> "" Then
        'MsgBox "Account has been closed previously. Ths ", vbExclamation, gAppName & " - Error"
        'If MsgBox("Account has been closed previously." & vbCrLf & _
                "This action will reopen the account." & vbCrLf & _
                "Do you want to continue ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
        If MsgBox(LoadResString(gLangOffSet + 524) & vbCrLf & _
                LoadResString(gLangOffSet + 548) & vbCrLf & _
                LoadResString(gLangOffSet + 541), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
                Exit Function
        Else  'Reopen the account first
            If Not AccountReopen(m_AccID) Then
                'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
                MsgBox LoadResString(gLangOffSet + 536), vbExclamation, gAppName & " - Error"
                Exit Function
            End If
        End If
    End If
    
    'Get last transaction record
    Dim Amount As Currency
    Dim Ret As Integer
    Dim TransID As Long
    Dim Transdate As String
    gDBTrans.SQLStmt = "Select TOP 1 * from AccTrans where AccID = " & m_AccID & " order by TransID desc"
    Ret = gDBTrans.SQLFetch
    If Ret >= 1 Then
        Amount = FormatField(gDBTrans.Rst("Amount"))
        TransID = FormatField(gDBTrans.Rst("TransID"))
        Transdate = FormatField(gDBTrans.Rst("TransDate"))
    ElseIf Ret = 0 Then
        'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 551), vbInformation, gAppName & " - Error"
        Exit Function
    End If
    
    'Confirm UNDO
    'If MsgBox("Are you sure you want to undo the last transaction of Rs." & Amount & "?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
    If MsgBox(LoadResString(gLangOffSet + 583) & Amount & "?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
        Exit Function
    End If
    
Dim InsertReserveFund As Boolean
        'See if the date is earlier than last date of transaction
If m_AccID > wis_ReserveFundHead And _
   m_AccID < wis_ReserveFundHead + wis_BankHead And _
   FormatField(gDBTrans.Rst.fields("TransType")) = wContraDeposit Then
   InsertReserveFund = True
End If
   
   
'Delete record from Data base
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "Delete from AccTrans where AccID = " & m_AccID & " and TransID = " & TransID
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.CommitTrans

    If InsertReserveFund Then
      Dim Balance As Currency
      Balance = OBOfAccount(wis_ProfitOrLoss, Transdate)
      Call StoreBalances(wis_ProfitOrLoss, Transdate, Balance + Amount)
    End If

AccountUndoLastTransaction = True
End Function



