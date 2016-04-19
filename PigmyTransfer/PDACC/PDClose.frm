VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmPDClose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PD Close"
   ClientHeight    =   5100
   ClientLeft      =   405
   ClientTop       =   1575
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   3810
      Width           =   8295
      Begin VB.OptionButton optTransfer 
         Caption         =   "&Transfer  (Contra)"
         Height          =   315
         Left            =   4410
         TabIndex        =   29
         Top             =   210
         Width           =   3165
      End
      Begin VB.OptionButton optClose 
         Caption         =   "&Close (Cash)"
         Height          =   285
         Left            =   150
         TabIndex        =   28
         Top             =   240
         Width           =   3225
      End
      Begin VB.CommandButton cmdTfr 
         Caption         =   "..."
         Height          =   315
         Left            =   7830
         TabIndex        =   27
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.Frame fraDepDetail 
      Caption         =   "Deposit Details"
      Height          =   3495
      Left            =   60
      TabIndex        =   15
      Top             =   450
      Width           =   4305
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   315
         Left            =   3960
         TabIndex        =   31
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInterest 
         Height          =   345
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   1038
         Width           =   1335
      End
      Begin VB.TextBox txtDate 
         Height          =   345
         Left            =   2610
         TabIndex        =   16
         Top             =   210
         Width           =   1335
      End
      Begin WIS_Currency_Text_Box.CurrText txtIntPayable 
         Height          =   345
         Left            =   2610
         TabIndex        =   22
         Top             =   1452
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtDepInterest 
         Height          =   345
         Left            =   2610
         TabIndex        =   24
         Top             =   1866
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotalInt 
         Height          =   345
         Left            =   2610
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtDepositAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2610
         TabIndex        =   32
         Top             =   624
         Width           =   1335
      End
      Begin VB.Label txtPayableAmount 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2610
         TabIndex        =   35
         Top             =   2820
         Width           =   1425
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   3930
         Y1              =   2730
         Y2              =   2730
      End
      Begin VB.Label lblNetAmount 
         Caption         =   "Net payable amount:"
         Height          =   255
         Left            =   180
         TabIndex        =   34
         Top             =   2850
         Width           =   2145
      End
      Begin VB.Label lblTotalIntAmount 
         Caption         =   "Toatal Interest on deposit:"
         Height          =   225
         Left            =   150
         TabIndex        =   25
         Top             =   2340
         Width           =   2145
      End
      Begin VB.Label lblIntAmount 
         Caption         =   "Interest on deposit:"
         Height          =   225
         Left            =   150
         TabIndex        =   23
         Top             =   1920
         Width           =   2145
      End
      Begin VB.Label lblIntPayable 
         Caption         =   "Payable Interest"
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   1500
         Width           =   2145
      End
      Begin VB.Label lblDepInt 
         Caption         =   "Rate of Interest:"
         Height          =   225
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Width           =   2115
      End
      Begin VB.Label lblDepAmount 
         Caption         =   "Deposited amount:"
         Height          =   225
         Left            =   180
         TabIndex        =   18
         Top             =   690
         Width           =   2115
      End
   End
   Begin VB.Frame fraCharges 
      Caption         =   "Charges"
      Height          =   1305
      Left            =   4350
      TabIndex        =   10
      Top             =   2640
      Width           =   4005
      Begin WIS_Currency_Text_Box.CurrText txtCharges 
         Height          =   345
         Left            =   2670
         TabIndex        =   14
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTax 
         Height          =   345
         Left            =   2670
         TabIndex        =   12
         Top             =   720
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblPreClose 
         Caption         =   "Premature closure charges: "
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Width           =   2145
      End
      Begin VB.Label lblOthers 
         Caption         =   "Other charges (Tax, etc.)"
         Height          =   225
         Left            =   150
         TabIndex        =   11
         Top             =   840
         Width           =   2145
      End
   End
   Begin VB.Frame fraLoanDet 
      Caption         =   "Loan Details"
      Height          =   2265
      Left            =   4350
      TabIndex        =   5
      Top             =   450
      Width           =   4005
      Begin VB.CheckBox ChkDeductLoan 
         Alignment       =   1  'Right Justify
         Caption         =   "Deduct Loan Amount :"
         Height          =   405
         Left            =   210
         TabIndex        =   30
         Top             =   1620
         Width           =   3615
      End
      Begin VB.TextBox txtLoanAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox txtLoanRate 
         Height          =   345
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   615
         Width           =   1275
      End
      Begin WIS_Currency_Text_Box.CurrText txtLoanInterest 
         Height          =   345
         Left            =   2610
         TabIndex        =   9
         Top             =   1020
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblIntOnLoan 
         Caption         =   "Interest on loans: "
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   1110
         Width           =   2145
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Total loan amount:"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   2115
      End
      Begin VB.Label lblLoanInt 
         Caption         =   "Rate of interest:"
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   690
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7170
      TabIndex        =   4
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   5790
      TabIndex        =   3
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "Customer Name"
      Height          =   345
      Left            =   60
      TabIndex        =   33
      Top             =   60
      Width           =   8295
   End
End
Attribute VB_Name = "frmPDClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_AccID As Long

Private m_AccHeadId As Long

Private m_AgentID As Long
Private m_retVar
Private m_LoanID As Long
Private m_RstMaster As Recordset
Private m_DeposiName As String

Private M_setUp As New clsSetup
Private m_ContraClass As clsContra
Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1


Public Property Let AccountId(NewValue As Long)
    m_AccID = NewValue
End Property


Private Function PDClose() As Boolean

Dim TransDate As Date
Dim MatDate As Date
Dim tmpTransID As Long
Dim TransID As Long
Dim Amount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim Balance As Currency
'Dim IntBalance As Currency
Dim PayableBalance As Currency
Dim Rst As Recordset

Dim PDHeadID As Long
Dim IntHeadID As Long
Dim PayableHeadID As Long
Dim HeadName As String
Dim BankClass As clsBankAcc


'Date specified must be latest
gDbTrans.SQLStmt = "Select Balance from PDTrans where" & _
            " AccID = " & m_AccID & " order by transid desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then _
            Balance = FormatField(Rst.Fields("Balance"))

TransDate = GetSysFormatDate(txtDate)


'Do not allow if this deposit has loans
If Val(txtLoanAmount) > 0 And ChkDeductLoan.Value = vbUnchecked Then
    'MsgBox "The deposit you are trying to close has loans against it" & vbCrLf & _
            "You must first get the loan repayment and then close this deposit", vbInformation, gAppName & " - Message"
   If MsgBox(LoadResString(gLangOffSet + 574) & vbCrLf & vbCrLf & _
        LoadResString(gLangOffSet + 541), vbQuestion + vbYesNo + vbDefaultButton2, _
         gAppName & " - Confirmation") = vbNo Then Exit Function
End If

gDbTrans.SQLStmt = "Select Balance from PDIntPayable where" & _
            " AccID = " & m_AccID & " order by transid desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) Then PayableBalance = FormatField(Rst.Fields("Balance"))

'Do not allow if this deposit has loans
If Val(txtLoanAmount.Text) > Val(txtDepositAmount) Then
    'MsgBox "The deposit you are trying to close has loans against it" & vbCrLf & _
            "You must first get the loan repayment and then close this deposit", vbInformation, gAppName & " - Message"
    MsgBox LoadResString(gLangOffSet + 574) & vbCrLf & _
            LoadResString(gLangOffSet + 541), vbInformation, gAppName & " - Message"
       'Check for the LoanDeduction IF He has not checked the LoanDedcuction
    'Check for the LoanDeduction IF He has not checked the LoanDedcuction
    If MsgBox(LoadResString(gLangOffSet + 541) _
            , vbQuestion + vbYesNo, gAppName & " - Message") = vbNo Then Exit Function
    
End If


Dim VoucherNo As String

'Get free TransID
Dim Trans As wisTransactionTypes
Dim IntTransType As wisTransactionTypes

Dim boolContra As Boolean

PayableAmount = txtIntPayable
IntAmount = txtDepInterest
Amount = Val(txtDepositAmount)

PayableBalance = PayableBalance - PayableAmount
Balance = Balance - Amount

If PayableBalance < 0 Then
    If MsgBox("You are withdrawing more amount than that of deposited from Payble account" & _
        vbCrLf & "Do you wnat to continue?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    PayableBalance = 0
End If
'If He closes the account before three months so Taking some Charges
'for Pigmy agent's commission
IntAmount = IntAmount - Val(txtCharges.Text)
' Now Get the INterest Amount deposited in InterestPayble Account
'Get the Next TransId
If optTransfer Then boolContra = True


Dim InTrans As Boolean

'Get the MAx Transction ID
TransID = GetPigmyMaxTransID(m_AccID) + 1

gDbTrans.BeginTrans

Set BankClass = New clsBankAcc

Dim UserID As Long
UserID = gCurrUser.UserID
'Close the account by giving deposit
Trans = wWithdraw
If optTransfer Then Trans = wContraWithdraw
'Get pigmy HeadId
HeadName = LoadResString(gLangOffSet + 425)
PDHeadID = BankClass.GetHeadIDCreated(HeadName, parMemberDeposit, 0, wis_PDAcc)

gDbTrans.SQLStmt = "Insert into PDTrans " & _
                " (AccID,TransID, TransType, " & _
                " TransDate, Amount, Balance,VoucherNo, " & _
                " Particulars,UserID ) values ( " & _
                m_AccID & "," & _
                TransID & "," & _
                Trans & "," & _
                "#" & TransDate & "#," & _
                Amount & "," & Balance & "," & _
                AddQuotes(VoucherNo, True) & "," & _
                " 'By Deposit Repayment' ," & _
                UserID & " )"
If Not gDbTrans.SQLExecute Then GoTo ErrLine
If Trans = wWithdraw Then _
    If Not BankClass.UpdateCashWithDrawls(PDHeadID, Amount, _
            TransDate) Then GoTo ErrLine

If IntAmount Then
    IntTransType = IIf(IntAmount > 0, wWithdraw, wDeposit)
    If ChkDeductLoan = vbChecked Then _
        IntTransType = IIf(IntAmount > 0, wContraWithdraw, wContraDeposit)
    
    If boolContra Then IntTransType = IIf(IntTransType = wDeposit, wContraDeposit, wContraWithdraw)
    
    gDbTrans.SQLStmt = "Insert into PDIntTrans " & _
                " (AccID, TransID, TransType, " & _
                " TransDate, Amount,Balance, VoucherNo," & _
                " Particulars,UserID ) values ( " & _
                m_AccID & "," & _
                TransID & "," & _
                IntTransType & "," & _
                "#" & TransDate & "#," & _
                Abs(IntAmount) & ", 0," & _
                AddQuotes(VoucherNo, True) & "," & _
                "'By interest' ," & _
                UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    'Get pigmy interest HeadId
    HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 487)
    IntHeadID = BankClass.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_PDAcc)

    If IntTransType = wWithdraw Then _
        If Not BankClass.UpdateCashWithDrawls(IntHeadID, IntAmount, _
                TransDate) Then GoTo ErrLine
    If IntTransType = wDeposit Then _
        If Not BankClass.UpdateCashDeposits(IntHeadID, Abs(IntAmount), _
                TransDate) Then GoTo ErrLine
End If

If PayableAmount <> 0 Then
    Trans = wWithdraw
    If ChkDeductLoan = vbChecked Then Trans = wContraWithdraw
    
    'Get pigmy payable HeadId
    'HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 450)
    HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 375) & " " & LoadResString(gLangOffSet + 47)
    PayableHeadID = BankClass.GetHeadIDCreated(HeadName, parDepositIntProv, 0, wis_PDAcc)
    
    gDbTrans.SQLStmt = "Insert into PDIntPayable (AccID, TransID, TransType, " & _
            " TransDate, Amount,Balance, VoucherNo,Particulars,UserID ) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            Trans & "," & _
            "#" & TransDate & "#," & _
            PayableAmount & "," & _
            PayableBalance & "," & _
            AddQuotes(VoucherNo, True) & "," & _
            "'By interest' ," & _
            UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    If Trans = wWithdraw Then _
        If Not BankClass.UpdateCashWithDrawls(PayableHeadID, PayableAmount, _
                TransDate) Then GoTo ErrLine
    
End If
'Update the first transaction with close = 1
gDbTrans.SQLStmt = "UPdate PDMaster Set ClosedDate = #" & TransDate & "# " & _
        " where AccID = " & m_AccID

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'/////Contra And Suspense account goes here
If boolContra Then
    Dim ContraID As Long
    'Get the Contra ID
     ContraID = GetMaxContraTransID + 1
    'put withdrawal transction details int to contra Table
     '''/////
     'First insert the Deposit Amount
     gDbTrans.SQLStmt = "Insert into ContraTrans " & _
            "(ContraId,AccId,AccHeadID," _
             & " TransType,TransId,Amount,VoucherNo)" & _
            "Values(" _
             & ContraID & "," & _
             m_AccID & "," & _
             m_AccHeadId & "," & _
             Trans & ", " & TransID & "," & _
             Amount & "," & _
             AddQuotes(VoucherNo, True) & _
             " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    'Now insert the Intesrt amount
     gDbTrans.SQLStmt = "Insert into ContraTrans " & _
            "(ContraId,AccId,AccHeadID," _
             & " TransType,TransId,Amount,VoucherNo)" & _
            "Values(" _
             & ContraID & "," & _
             m_AccID & "," & _
             IntHeadID & "," & _
             IntTransType & ", " & TransID & "," & _
             IntAmount & "," & _
             AddQuotes(VoucherNo, True) & _
             " )"
    If IntAmount Then If Not gDbTrans.SQLExecute Then GoTo ErrLine
    'Now insert the payable amount
     gDbTrans.SQLStmt = "Insert into ContraTrans " & _
            "(ContraId,AccId,AccHeadID," _
             & " TransType,TransId,Amount,VoucherNo)" & _
            "Values(" _
             & ContraID & "," & _
             m_AccID & "," & _
             PayableHeadID & "," & _
             Trans & ", " & TransID & "," & _
             PayableAmount & "," & _
             AddQuotes(VoucherNo, True) & _
             " )"
    If IntAmount Then If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
End If

If boolContra And m_ContraClass Is Nothing Then Call cmdTfr_Click

If Not m_ContraClass Is Nothing Then
    m_ContraClass.TransDate = TransDate
    If Not m_ContraClass.SaveDetails Then GoTo ErrLine
    Set m_ContraClass = Nothing
End If

'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If Trans = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(m_AccHeadId, _
                m_AccID, TransDate, TransID, Val(txtPayableAmount)) < 1 Then GoTo ErrLine
    Set Cashclass = Nothing
End If

gDbTrans.CommitTrans
InTrans = False

Set BankClass = Nothing
Unload Me

Exit Function

ErrLine:
    
    If InTrans Then gDbTrans.RollBack
    'MsgBox "Unable to perform transaction !", vbCritical, gAppName & " - Critical Error"
    MsgBox LoadResString(gLangOffSet + 535), vbCritical, gAppName & " - Critical Error"
    Set BankClass = Nothing
    Exit Function


End Function

Private Function TransferToLoanAccount() As Boolean

Dim VoucherNo As String
Dim TransDate As Date
Dim TransID As Long

Dim DepAmount As Currency
Dim MatAmount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim Balance As Currency
'Dim IntBalance As Currency
Dim PayableBalance As Currency
Dim Rst As Recordset

Dim PDHeadID As Long
Dim IntHeadID As Long
Dim PayableHeadID As Long
Dim HeadName As String
Dim CustId As Long

'Get pigmy HeadId
HeadName = LoadResString(gLangOffSet + 425)
PDHeadID = GetIndexHeadID(HeadName)
'Get pigmy interest HeadId
HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 487)
IntHeadID = GetIndexHeadID(HeadName)
'Get pigmy payable HeadId
HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 450)
PayableHeadID = GetIndexHeadID(HeadName)

'Get the Customer ID of this account
gDbTrans.SQLStmt = "Select CustomerID from PDMaster where" & _
                    " AccID = " & m_AccID
If gDbTrans.Fetch(Rst, adOpenForwardOnly) Then _
        CustId = FormatField(Rst.Fields("CustomerID"))

'get the Balance of the account
gDbTrans.SQLStmt = "Select Balance from PDTrans " & _
                " where AccID = " & m_AccID & " order by transid desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then _
                    Balance = FormatField(Rst.Fields("Balance"))

gDbTrans.SQLStmt = "Select TransDate,Transid,Balance " & _
                    " From PDIntPayable where" & _
                    " AccID = " & m_AccID & " order by Transid desc"
            
If gDbTrans.Fetch(Rst, adOpenForwardOnly) Then _
            PayableBalance = FormatField(Rst.Fields("Balance"))

Dim Trans As wisTransactionTypes

PayableAmount = Val(txtIntPayable)
IntAmount = txtDepInterest
DepAmount = Val(txtDepositAmount)
MatAmount = DepAmount + IntAmount + PayableAmount

PayableBalance = PayableBalance - PayableAmount
Balance = Balance - DepAmount

If PayableBalance < 0 Then
    If MsgBox("You are withdrawing more amount than that of deposited from Payble account" & _
        vbCrLf & LoadResString(gLangOffSet + 541), vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    PayableBalance = 0
End If
'If He closes the account before three months so Taking some Charges
'for Pigmy agent's commission
IntAmount = IntAmount - Val(txtCharges.Text)
' Now Get the INterest Amount deposited in InterestPayble Account

'Get the Next TransId
TransID = GetPigmyMaxTransID(m_AccID) + 1

If TransID = 1 Then Exit Function

Dim BankClass As clsBankAcc
Dim InTrans As Boolean

gDbTrans.BeginTrans
InTrans = True
Set BankClass = New clsBankAcc

Dim UserID As Long
Dim SuspAmount As Currency

SuspAmount = Val(txtPayableAmount)
UserID = gCurrUser.UserID

'Close the account by giving deposit
Trans = wContraWithdraw
gDbTrans.SQLStmt = "Insert into PDTrans (AccID,TransID, TransType, " & _
                " TransDate, Amount, Balance," & _
                " VoucherNo ,Particulars,UserID ) values ( " & _
                m_AccID & "," & _
                TransID & "," & _
                Trans & "," & _
                "#" & TransDate & "#," & _
                DepAmount & "," & Balance & "," & _
                AddQuotes(VoucherNo, True) & "," & _
                " 'By Deposit Repayment' ," & _
                UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

If IntAmount Then
    Trans = IIf(IntAmount > 0, wContraWithdraw, wContraDeposit)
    gDbTrans.SQLStmt = "Insert into PDIntTrans (AccID, TransID, TransType, " & _
            " TransDate, Amount,Balance, VoucherNo,Particulars,UserID ) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            Trans & "," & _
            "#" & TransDate & "#," & _
            Abs(IntAmount) & ", 0," & _
            AddQuotes(VoucherNo, True) & "," & _
            "'By interest' ," & _
            UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    If Trans = wContraWithdraw Then _
        If Not BankClass.UpdateContraTrans(IntHeadID, PDHeadID, IntAmount, _
                TransDate) Then GoTo ErrLine
    If Trans = wContraDeposit Then
        If Not BankClass.UpdateContraTrans(IIf(PayableAmount, PayableHeadID, PDHeadID), _
                    PDHeadID, Abs(IntAmount), TransDate) Then GoTo ErrLine
        PayableAmount = IIf(PayableAmount, PayableAmount + IntAmount, 0)
     End If
End If

If PayableAmount Then
    Trans = wContraWithdraw
    
    gDbTrans.SQLStmt = "Insert into PDIntPayable (AccID, TransID, TransType, " & _
            " TransDate, Amount,Balance, VoucherNo,Particulars,UserID ) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            Trans & "," & _
            "#" & TransDate & "#," & _
            PayableAmount & "," & _
            PayableBalance & "," & _
            AddQuotes(VoucherNo, True) & "," & _
            "'By interest' ," & _
            UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine

    If Not BankClass.UpdateContraTrans(PayableHeadID, _
                        PDHeadID, PayableAmount, TransDate) Then GoTo ErrLine
    
End If

'Update the first transaction with close = 1
gDbTrans.SQLStmt = "UPdate PDMaster " & _
        " Set ClosedDate = #" & TransDate & "# " & _
        " Where AccID = " & m_AccID

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'/////Contra And Suspense account goes here
'if he is repaying the loan amount or trnsferring to other account
'then the transction will be contra
Dim ContraID As Long
'Get the Contra ID
ContraID = GetMaxContraTransID + 1

'put withdrawal transction details int to contra Table
 '''/////
 gDbTrans.SQLStmt = "Insert into ContraTrans " & _
        "(ContraId,AccId,AccHeadID," _
         & " TransType,TransId,Amount,VoucherNo)" & _
        "Values(" _
         & ContraID & "," & _
         m_AccID & "," & _
         m_AccHeadId & "," & _
         Trans & ", " & TransID & "," & _
         DepAmount & "," & _
         AddQuotes(VoucherNo, True) & _
         " )"

If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

'if he is transfering the Amount to the loan account,
'transfer it to loan account
Dim LoanAmount As Currency
Dim LoanIntAmount As Currency

MatAmount = txtDepositAmount + txtTotalInt
LoanAmount = Val(txtLoanAmount)
LoanIntAmount = txtLoanInterest

Dim DepLOanClass As clsDepLoan
Set DepLOanClass = New clsDepLoan

'if loan amount is more than the matured amount
'then first take the interest then remaining amount as principal
'loanclass
MatAmount = MatAmount - LoanIntAmount
LoanAmount = IIf(LoanAmount < MatAmount, LoanAmount, MatAmount)
MatAmount = MatAmount - LoanAmount
If DepLOanClass.DepositAmount(CInt(Rst(0)), LoanAmount, _
        LoanIntAmount, "PD Close", TransDate, VoucherNo) = 0 Then GoTo ErrLine

Dim LoanIntHeadID As Long
Dim LoanHeadID As Long
HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 58)
LoanHeadID = GetIndexHeadID(HeadName)
HeadName = HeadName & " " & LoadResString(gLangOffSet + 483)
LoanIntHeadID = GetIndexHeadID(HeadName)
If LoanIntAmount Then _
    If Not BankClass.UpdateContraTrans(PDHeadID, LoanIntHeadID, _
        LoanIntAmount, TransDate) Then GoTo ErrLine
If Not BankClass.UpdateContraTrans(PDHeadID, LoanHeadID, _
    LoanAmount, TransDate) Then GoTo ErrLine

'if matured amount is more than the loan amount
'then transfer tht remianing amount to the suspence account
If MatAmount > 0 Then
    Dim SuspHeadID As Long
    SuspHeadID = GetIndexHeadID(LoadResString(gLangOffSet + 365))
    Debug.Assert MatAmount = 0
    SuspHeadID = BankClass.GetHeadIDCreated(LoadResString(gLangOffSet + 365), _
                                    parSuspAcc, 0, wis_SuspAcc)
    
    Dim SuspClass As New clsSuspAcc
    If SuspClass.DepositAmount(PDHeadID, m_AccID, CustId, "", _
                        TransDate, MatAmount, TransID, VoucherNo) < 1 Then GoTo ErrLine
    
    
    If Not BankClass.UpdateContraTrans(PDHeadID, SuspHeadID, _
        MatAmount, TransDate) Then GoTo ErrLine
    
End If

gDbTrans.CommitTrans
InTrans = False
Set BankClass = Nothing

TransferToLoanAccount = True

ErrLine:

    If InTrans Then gDbTrans.RollBack
    'MsgBox "Unable to perform transaction !", vbCritical, gAppName & " - Critical Error"
    MsgBox LoadResString(gLangOffSet + 535), vbCritical, gAppName & " - Critical Error"
        
    Set BankClass = Nothing
    Exit Function


End Function

'
Private Sub UpdateDetails()
Dim Days As Long
Dim DepDate As Date
Dim DepAmt As Currency
Dim MatDate As Date
Dim TransType As wisTransactionTypes
Dim AsOnDate As Date
Dim LoanID As Long
Dim Rst As Recordset

Dim ClsBank As clsBankAcc
If m_AccHeadId = 0 Then
    Set ClsBank = New clsBankAcc
    gDbTrans.BeginTrans
    m_AccHeadId = ClsBank.GetHeadIDCreated(LoadResString(gLangOffSet + 425), _
                    parMemberDeposit, 0, wis_PDAcc)
    gDbTrans.CommitTrans
    Set ClsBank = Nothing
End If

'Check for valid date
    If Trim$(txtDate.Text) = "" Then
        txtDate.Text = gStrDate
        AsOnDate = gStrDate
    Else
        If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub
        AsOnDate = GetSysFormatDate(txtDate)
    End If

'Get the deposited amount
'    TransType = wDeposit
    gDbTrans.SQLStmt = "Select TOP 1 Balance from PDTrans where " & _
                    " AccID = " & m_AccID & _
                    " ORDER BY TransID DESC"
    
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 570), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        DepAmt = FormatField(Rst("Balance"))
    End If
    gDbTrans.SQLStmt = "Select * from PDMaster where AccID = " & m_AccID
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 570), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    MatDate = Rst.Fields("MaturityDate")
    DepDate = Rst.Fields("CreateDate")
    LoanID = FormatField(Rst("LoanID"))
    
  '''now get the depositIntreat
Dim IntAmount As Currency

IntAmount = ComputePDInterestAmount(m_AccID, DepDate, True)
IntAmount = IntAmount \ 1
txtDepInterest = IntAmount
'Get the Total
txtTotalInt = txtIntPayable + txtDepInterest

'Calculate the number of days
    Days = DateDiff("d", DepDate, AsOnDate)
    If Days > 0 Then  'Account being closed prematurely
        Days = DateDiff("D", DepDate, AsOnDate)
    Else
        Days = DateDiff("D", DepDate, MatDate)
    End If
'Extract the rate of interest from Setup values
    Dim IntRate As Single
    
    IntRate = GetPDDepositInterest(Days, GetIndianDate(DepDate))
    txtInterest.Text = Format(Val(IntRate), "#0.00")
    
'Now Read the Int Rates from SetUp calss
'Dim SetupClass As New clsSetup
'If Val(txtLoanRate.Text) = 0 Then
    'txtloanrate.Text=format(setupclass.ReadSetupValue ("PDAcc",
'End If
'Set SetupClass = Nothing

txtDepositAmount = FormatCurrency(DepAmt)

'IntAmount = PDInterest(m_Accid)

If IntAmount < 0 Then
    txtCharges = (IntAmount * -1) \ 1
    txtDepInterest = 0
Else
    txtDepInterest.Text = IntAmount \ 1
End If


gDbTrans.SQLStmt = " SELECT * From PDIntPayable WHERE ACCID = " & m_AccID & _
            " AND TransID = (SELECT MAx(TransID) FROM PDIntPayable " & _
                " WHERE ACCID = " & m_AccID & ")"
txtIntPayable = "0.00"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) Then _
        txtIntPayable = FormatField(Rst.Fields("Balance"))

'Get total loan amount
Dim LnAmt As Currency
Dim LnTransDate As Date

'Get sum of withdrawals as loans drawn
gDbTrans.SQLStmt = "Select Top 1 Balance as TotalLoan,TransDate " & _
            "from DepositLoanTrans where LoanID = " & LoanID

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
    
LnAmt = Val(FormatField(Rst.Fields("TotalLoan")))
LnTransDate = Rst.Fields("Transdate")
gDbTrans.SQLStmt = "Select *  from DepositLoanMaster " & _
        " where LoanID = " & LoanID
If gDbTrans.Fetch(Rst, adOpenForwardOnly) Then _
    txtLoanRate = Val(FormatField(Rst.Fields("InterestRate")))

Days = DateDiff("D", LnTransDate, AsOnDate)
If Days < 0 Then Days = 0
'Get date of last transaction
txtLoanAmount.Text = FormatCurrency(LnAmt)
'Doubt
'computetefdinterest
txtLoanInterest.Text = FormatCurrency(LnAmt * Days / 365 * Val(txtLoanRate) / 100)

optClose = True
    

End Sub

Private Sub ChkDeductLoan_Click()
If ChkDeductLoan.Value = vbChecked Then
    optClose = True
    optTransfer.Enabled = False
Else
    optTransfer.Enabled = True
End If
Call txtDepositAmount_Change
End Sub

Private Sub cmdAccept_Click()
'Perform the transaction with closed flag and send the guy home
'Check date
Dim TransDate As Date
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date not in dd/mm/yyyy format ", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
TransDate = GetSysFormatDate(txtDate)
'Check For Last Date of Transaction
If DateDiff("d", TransDate, GetPigmyLastTransDate(m_AccID)) > 0 Then
    'MsgBox "Early Date transaction", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If

'Warn for premature closure
Dim Rst As Recordset
gDbTrans.SQLStmt = "Select MaturityDate from PDMaster where " & _
                    " AccID = " & m_AccID
If gDbTrans.Fetch(Rst, adOpenForwardOnly) <= 0 Then
    'MsgBox "Deposit not found !", vbCritical, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 658), vbCritical, gAppName & " - Error"
    Exit Sub
End If

If DateDiff("d", Rst.Fields("MaturityDate"), TransDate) < 0 Then
    'If MsgBox("You are attempting to close this deposit prematurely !" & vbCrLf & vbCrLf & "Are you sure you want to continue this operation ?", vbQuestion + vbYesNo + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then
    If MsgBox(LoadResString(gLangOffSet + 576) & vbCrLf & _
        LoadResString(gLangOffSet + 541), vbQuestion + vbYesNo + vbDefaultButton2, _
        gAppName & " - Confirmation") = vbNo Then
        Unload Me
        Exit Sub
    End If
End If

If ChkDeductLoan.Value = vbChecked Then
    If Not TransferToLoanAccount Then Exit Sub
Else
    If Not PDClose Then Exit Sub
End If

Unload Me

End Sub



Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDate_Click()
Dim strDate As String
With Calendar
    .Left = Me.Left + cmdDate.Left - .Width / 2
    .Top = Me.Top + cmdDate.Top
    .selDate = txtDate.Text
    strDate = .selDate
    .Show vbModal
    txtDate.Text = .selDate
    If .selDate = strDate Then Exit Sub
End With

Call txtDate_LostFocus

End Sub

'
Private Sub cmdTfr_Click()
Dim AccNum As String
Dim Rst As Recordset

Dim PayableAmount As Currency
Dim IntAmount As Currency
Dim Amount As Currency

Dim PDHeadID As Long
Dim IntHeadID As Long
Dim PayableHeadID As Long
Dim HeadName As String
Dim BankCls As clsBankAcc

gDbTrans.SQLStmt = "SELECT AccNum From PdMaster Where AccId = " & m_AccID
If gDbTrans.Fetch(Rst, adOpenDynamic) Then AccNum = FormatField(Rst(0))

''Get pigmy HeadId
HeadName = LoadResString(gLangOffSet + 425)
PDHeadID = GetIndexHeadID(HeadName)

PayableAmount = Val(txtIntPayable)
IntAmount = Val(txtDepInterest)
Amount = Val(txtDepositAmount)

Set BankCls = New clsBankAcc
Set m_ContraClass = New clsContra
With m_ContraClass
    
    If PayableAmount Then
        HeadName = LoadResString(gLangOffSet + 425) & " " & _
            " " & LoadResString(gLangOffSet + 375) & " " & LoadResString(gLangOffSet + 47)
        PayableHeadID = BankCls.GetHeadIDCreated(HeadName, parDepositIntProv, 0, wis_PDAcc)
        
        Call .Transfer(GetSysFormatDate(txtDate), "12", PayableHeadID, AccNum, PayableAmount)
    End If
    
    If IntAmount Then
        HeadName = LoadResString(gLangOffSet + 425) & " " & LoadResString(gLangOffSet + 487)
        IntHeadID = BankCls.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_PDAcc)
                        
        Call .Transfer(GetSysFormatDate(txtDate), "12", IntHeadID, AccNum, IntAmount)
    End If
    
    Call .Transfer(GetSysFormatDate(txtDate), "12", m_AccHeadId, AccNum, Amount)
    .Show
End With


End Sub

'
Private Sub Form_Load()
'Center the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'Set kannada fonts
Call SetKannadaCaption
         
'Todays date
    txtDate.Text = gStrDate

Call UpdateDetails

If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub
'
Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
Set frmPDClose = Nothing
End Sub


'
Private Sub optClose_Click()
cmdTfr.Enabled = optTransfer.Value
End Sub


'
Private Sub optTransfer_Click()
cmdTfr.Enabled = optTransfer.Value
End Sub

Private Sub txtCharges_Change()
Call txtDepositAmount_Change
End Sub

Private Sub txtDate_LostFocus()
Call UpdateDetails
End Sub


Private Sub txtDepInterest_Change()
Call txtIntPayable_Change
End Sub

Private Sub txtDepositAmount_Change()
If ChkDeductLoan.Value = vbChecked Then
    txtPayableAmount = Val(txtDepositAmount) + txtTotalInt - _
            Val(txtLoanAmount) - txtLoanInterest - txtCharges - txtTax
Else
    txtPayableAmount = Val(txtDepositAmount) + txtTotalInt - _
             txtCharges - txtTax
End If
End Sub

Private Sub txtInterest_Change()
Call txtIntPayable_Change
End Sub

Private Sub txtIntPayable_Change()
    txtTotalInt = txtIntPayable + txtDepInterest
    
End Sub

Private Sub txtLoanInterest_Click()
    Call txtDepositAmount_Change
End Sub

Private Sub txtLoanRate_LostFocus()
Call UpdateDetails
End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

fraDepDetail = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 295)
fraLoanDet = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 295)
fraCharges = LoadResString(gLangOffSet + 237) & " " & LoadResString(gLangOffSet + 273)

'Set the Kannada caption to the Command buttons
cmdAccept.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)

lblDate = LoadResString(gLangOffSet + 37)
lblDepAmount = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 40)
lblDepInt = LoadResString(gLangOffSet + 186)
lblIntPayable = LoadResString(gLangOffSet + 450)
lblIntAmount = LoadResString(gLangOffSet + 233)
lblTotalIntAmount = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 233)

lblLoanAmount = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 91)
lblLoanInt = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 186)
lblIntOnLoan = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 47)

lblPreClose = LoadResString(gLangOffSet + 238)
lblOthers = LoadResString(gLangOffSet + 237) & " " & LoadResString(gLangOffSet + 273)   '
lblNetAmount = LoadResString(gLangOffSet + 240)

End Sub


Private Sub txtTax_Click()
Call txtDepositAmount_Change
End Sub

Private Sub txtTotalInt_Change()
'Call txtIntPayable_Change
txtPayableAmount = Val(txtDepositAmount) + txtTotalInt.Value
End Sub

