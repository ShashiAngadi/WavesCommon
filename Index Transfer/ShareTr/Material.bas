Attribute VB_Name = "basMaterial"
Option Explicit

Public Sub CreateDefaultView()

'Call NewIndexTrans.DeleteAllViews

NewIndexTrans.SQLStmt = "SELECT A.TransID,A.HeadID,A.Debit,A.Credit,B.HeadID," & _
                " B.Debit,B.Credit,B.TransDate,B.VoucherType" & _
                " FROM AccTrans AS A" & _
                " INNER JOIN AccTrans AS B ON A.TransID=B.TransID "
    
NewIndexTrans.CreateView ("QryAccTransMerge")
NewIndexTrans.SQLStmt = "SELECT A.TransID, A.TransDate, A.HeadID, B.HeadID," & _
            " A.Credit, A.Debit, A.VoucherType " & _
            " FROM AccTrans AS A, AccTrans AS B, BankHeadIds AS C " & _
            " WHERE A.TransID=B.TransID And A.HeadID=C.HeadID;"
NewIndexTrans.CreateView ("QryAccBankTrans")

End Sub

'This function Adds the Project Defined Vouchers to Given Combo
'
' Inputs :
'           Combobox
'
' Pradeep
'
Public Sub LoadVouchersToCombo(cmbVouch As ComboBox)

' Handle the Error
On Error GoTo ComboFailed:

' do not delete this table
''   Receipt = 1
''   Payment = 2
''   Purchase = 3
''   Sales = 4
''   Free = 5
''   Journal = 6
''   Contra = 7
''   RejectionsIn = 8
''   RejectionsOut = 9

' Declare Variables

' Check and Clear the Combo
cmbVouch.Clear

' Start wtih Statement
With cmbVouch
    ' Start Adding Vocuhers
    .AddItem LoadResString(gLangOffSet + 196) '"Receipt"
    .ItemData(.NewIndex) = 1
    .AddItem LoadResString(gLangOffSet + 197) '"Payment"
    .ItemData(.NewIndex) = 2
'    .AddItem LoadResString(gLangOffSet + 176) '"Purchase"
'    .ItemData(.NewIndex) = 3
'    .AddItem LoadResString(gLangOffSet + 180) '"Sales"
'    .ItemData(.NewIndex) = 4
'    .AddItem LoadResString(gLangOffSet + 105) '"Free"
'    .ItemData(.NewIndex) = 5
    
''''''''The Following  Block is changed by shashi on 17/12/2002
    'As the old Index 2000 is haveing only two type transction
    ' i.e. 1) Deposit( receipt), 2) Withdraw(payment)
    'But in this new Index 2000 we introduced two More transction
    'of the same concept called contraDeposit, contraWithdraw
    'where Physical Cash is not hndels or the internal transfer
    ' is called Contra Transction
    'So to faclitate the Index 2000 concept Here we changed the
    'Tranction Type n the Cobo Box
    'And we are keeping the Same Voucher Type as in theis Project
    'We may change it later
    'And the one major change may be The
    'Contra Transaction Of this Project will become Cash Transction
    
    'Previous Code
''    .AddItem LoadResString(gLangOffSet + 198) '"Journal"
''    .ItemData(.NewIndex) = 6
''    .AddItem LoadResString(gLangOffSet + 199) '"Contra"
''    .ItemData(.NewIndex) = 7
    'NEW CODE
    .AddItem LoadResString(gLangOffSet + 199) '"Contra"
    .ItemData(.NewIndex) = 6

End With


Exit Sub

ComboFailed:
   
        
End Sub

' This sub adds all the heads to the Combobox for the given combobox
'
' Inputs :
'        Combobox
'        ParentID for which Heads to fetch
' On error It will clear the Combobox
'
' Pradeep--sir
'
Public Sub LoadLedgersToCombo(cmbLedger As ComboBox, ByVal ParentID As Long)

' Handle Errors
On Error GoTo NoLoadLedgers:

' Declarations
Dim rstHeads As ADODB.Recordset

' Check if Variable is empty
If ParentID = 0 Then Exit Sub

' clear the combobox
cmbLedger.Clear

' Set the Sql Statement
NewIndexTrans.SQLStmt = " SELECT HeadID,HeadName" & _
                   " FROM Heads WHERE ParentID =  " & ParentID & _
                   " ORDER BY HeadName"

' Fetch the Records
If NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly) < 0 Then Exit Sub
' Start the Loop
Do While Not rstHeads.EOF
    'Add the item to combo
    cmbLedger.AddItem FormatField(rstHeads("HeadName"))
    ' Set the itemdata
    cmbLedger.ItemData(cmbLedger.NewIndex) = FormatField(rstHeads("HeadID"))
    'Move to the next record
    rstHeads.MoveNext
Loop

' Exit
Exit Sub

NoLoadLedgers:
    ' on error
    cmbLedger.Clear
    
End Sub

' This sub adds all the heads to the Combobox for the given combobox
'
' Inputs :
'        Combobox
'        ParentID for which Heads to fetch
' On error It will clear the Combobox
'
' Pradeep
'
Public Sub LoadHeadsToCombo(cmbLedger As ComboBox, ByVal AccountType As wis_AccountType)

' Handle Errors
On Error GoTo NoLoadLedgers:

' Declarations
Dim rstHeads As ADODB.Recordset


' clear the combobox
cmbLedger.Clear

' Set the Sql Statement
NewIndexTrans.SQLStmt = " SELECT HeadID,HeadName" & _
                   " FROM Heads A,ParentHeads B" & _
                   " WHERE A.ParentID = B.ParentID" & _
                   " AND B.AccountType=" & AccountType & _
                   " AND B.UserCreated <= 2 " & _
                   " ORDER BY HeadName"

' Fetch the Records

If NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly) < 0 Then Exit Sub


' Start the Loop
Do While Not rstHeads.EOF
    ' Add the item to combo
    cmbLedger.AddItem rstHeads.Fields("HeadName")
    ' Set the itemdata
    cmbLedger.ItemData(cmbLedger.NewIndex) = rstHeads.Fields("HeadID")
    
    'Move to the next record
    rstHeads.MoveNext
    
Loop

' Exit
Exit Sub

NoLoadLedgers:
    ' on error
    cmbLedger.Clear
    
End Sub

'
Public Sub LoadParentHeads(ctrlComboBox As ComboBox)

On Error GoTo NoLoadParents:

Dim rstParent As ADODB.Recordset

ctrlComboBox.Clear

NewIndexTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads WHERE UserCreated <= 2" & _
                   " ORDER BY ParentName "

Call NewIndexTrans.Fetch(rstParent, adOpenForwardOnly)

If rstParent Is Nothing Or (rstParent.EOF And rstParent.BOF) Then
    InsertParentHeads
NewIndexTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads WHERE UserCreated <= 2" & _
                   " ORDER BY ParentName "
  
  Call NewIndexTrans.Fetch(rstParent, adOpenForwardOnly)
End If
Do While Not rstParent.EOF
    ctrlComboBox.AddItem FormatField(rstParent("ParentName"))
    ctrlComboBox.ItemData(ctrlComboBox.NewIndex) = FormatField(rstParent("ParentID"))
    
    'Move to the next record
    rstParent.MoveNext
Loop

Exit Sub

NoLoadParents:
    ctrlComboBox.Clear

End Sub
Public Sub InsertParentHeads()

Dim Rst As Recordset
Dim ParentName() As String
Dim AccountType As wis_AccountType
'Trap an error
On Error GoTo ErrLine

NewIndexTrans.SQLStmt = " SELECT * FROM ParentHeads"

If NewIndexTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then Exit Sub

ReDim Preserve ParentName(3, 20)
Dim i As Integer

i = 0
AccountType = Liability
'Share Capital
ParentName(0, i) = parShareCapital
ParentName(1, i) = LoadResString(gLangOffSet + 351)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Memebr Share
ParentName(0, i) = parMemberShare
ParentName(1, i) = LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 53)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Reserve and Suplus funds
ParentName(0, i) = parReserveFunds
ParentName(1, i) = LoadResString(gLangOffSet + 352)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1

'''Deposit (Liability)
''ParentName(0, i) = parDepositLiab
''ParentName(1, i) = LoadResString(gLangOffSet + 45)
''ParentName(2, i) = AccountType
''ParentName(3, i) = 1
''i = i + 1

'Memebr Deposit
ParentName(0, i) = parMemberDeposit
ParentName(1, i) = LoadResString(gLangOffSet + 49) & _
        " " & LoadResString(gLangOffSet + 45)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1   '5

'Bank LOan Accounts
ParentName(0, i) = parBankLoanAccount
ParentName(1, i) = LoadResString(gLangOffSet + 356)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Govt Loan subsidy
ParentName(0, i) = parGovtLoanSubsidy
ParentName(1, i) = LoadResString(gLangOffSet + 263)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1

'Other Loans
ParentName(0, i) = parOtherLoans
ParentName(1, i) = LoadResString(gLangOffSet + 237) & " " & LoadResString(gLangOffSet + 18)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Other Payables
ParentName(0, i) = parPayAble
ParentName(1, i) = LoadResString(gLangOffSet + 357)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Deposit Interest Provision
ParentName(0, i) = parDepositIntProv
ParentName(1, i) = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 450)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1   '10
'Loan Interest Provision
ParentName(0, i) = parLoanIntProv
ParentName(1, i) = LoadResString(gLangOffSet + 80) & " " & _
        LoadResString(gLangOffSet + 450)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Suspence Account
ParentName(0, i) = parSuspAcc
ParentName(1, i) = LoadResString(gLangOffSet + 365)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Account to vcarry forward the current years profit or loss
'to the next Finaancial year
'so in the next financial yaer it will be distributed to the funds
' And till then head is called previousyearProfit( Or Loss)
'Profit & Loss Account
ParentName(0, i) = parProfitORLoss
ParentName(1, i) = LoadResString(gLangOffSet + 443) & " " & LoadResString(gLangOffSet + 36)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'''ASSETS
AccountType = Asset
'Cash in Hand
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parCash
ParentName(1, i) = LoadResString(gLangOffSet + 350)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Bank Accounts
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parBankAccount
ParentName(1, i) = LoadResString(gLangOffSet + 359)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Investments (Assets)
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parInvestment
ParentName(1, i) = LoadResString(gLangOffSet + 361)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1   '15

'Loan And Advances (Assets)
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parLoanAdvanceAsset
ParentName(1, i) = LoadResString(gLangOffSet + 360)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1

'Member Loans
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parMemberLoan
ParentName(1, i) = LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 18)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'Member Deposit Loans
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parMemDepLoan
ParentName(1, i) = LoadResString(gLangOffSet + 49) & " " & _
    LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 18)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'Salary Advance
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parSalaryAdvance
ParentName(1, i) = LoadResString(gLangOffSet + 90) & " " & _
    LoadResString(gLangOffSet + 355)
ParentName(2, i) = AccountType
ParentName(3, i) = 2
i = i + 1

'Fixed Assets
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parFixedAsset
ParentName(1, i) = LoadResString(gLangOffSet + 363)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1

'ReceivAbles
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parReceivable
ParentName(1, i) = LoadResString(gLangOffSet + 364)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1

''INCOME HEADS
AccountType = Profit
'Income
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parIncome
ParentName(1, i) = LoadResString(gLangOffSet + 366)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1   '20
'Trading Income
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parTradingIncome
ParentName(1, i) = LoadResString(gLangOffSet + 367)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Regular Interest received on Member LOans
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parMemLoanIntReceived
ParentName(1, i) = LoadResString(gLangOffSet + 80) & " " & _
        LoadResString(gLangOffSet + 102) & " " & LoadResString(gLangOffSet + 344)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Penal Interest received on Member LOans
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parMemLoanPenalInt
ParentName(1, i) = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 345)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'Interest received on deposit Loans
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parMemDepLoanIntReceived
ParentName(1, i) = LoadResString(gLangOffSet + 43) & " " & _
        LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 483)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Interest received on Deposits
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parDepIntReceived
ParentName(1, i) = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 483)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'Income other Bank incomes
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parBankIncome
ParentName(1, i) = LoadResString(gLangOffSet + 418) & " " & _
        LoadResString(gLangOffSet + 366)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'''EXPENSE
AccountType = Loss
'Expense
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parExpense
ParentName(1, i) = LoadResString(gLangOffSet + 368) '& " " & LoadResString(gLangOffSet + 36)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1
'Trading expense
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parTradingExpense
ParentName(1, i) = LoadResString(gLangOffSet + 369)
ParentName(2, i) = AccountType
ParentName(3, i) = 1
i = i + 1   '25
'Interest paid on Member Deposit
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parMemDepIntPaid 'interest paid on Deposits parent head id
ParentName(1, i) = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 487)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Interest paid on Loans
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parLoanIntPaid
ParentName(1, i) = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 487)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Other expenditure in Bank accounts
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parBankExpense
ParentName(1, i) = LoadResString(gLangOffSet + 418) & " " & LoadResString(gLangOffSet + 368)
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

'Salary Expense
ReDim Preserve ParentName(3, i + 1)
ParentName(0, i) = parSalaryExpense
ParentName(1, i) = LoadResString(gLangOffSet + 90) & " " & _
    LoadResString(gLangOffSet + 36)
ParentName(2, i) = AccountType
ParentName(3, i) = 2
i = i + 1


'Purchase and Sales Account
AccountType = ItemPurchase
'Purchase
ReDim Preserve ParentName(3, i)
ParentName(0, i) = parPurchase
ParentName(1, i) = LoadResString(gLangOffSet + 176) & " " & LoadResString(gLangOffSet + 36)  ''"Purchase Account"
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1
'Sales account
ReDim Preserve ParentName(3, i)
AccountType = ItemSales
ParentName(0, i) = parSales
ParentName(1, i) = LoadResString(gLangOffSet + 180) & " " & LoadResString(gLangOffSet + 36) '"Sales Account"
ParentName(2, i) = AccountType
ParentName(3, i) = 3
i = i + 1

NewIndexTrans.BeginTrans

Dim MaxCount As Integer
Dim lpCount As Integer
Dim PrevType As wis_AccountType
Dim PrintOrder As Integer

PrintOrder = 1
MaxCount = i - 1

For lpCount = 0 To MaxCount
    'Change the print order as the acount type changes
    If PrevType = Val(ParentName(0, lpCount)) Then PrintOrder = 1
    
    NewIndexTrans.SQLStmt = " INSERT INTO ParentHeads " & _
        "(ParentID,ParentName,AccountType," & _
        " PrintDetailed,PrintOrder,UserCreated )" & _
        " VALUES ( " & _
        CLng(ParentName(0, lpCount)) & "," & _
        AddQuotes(ParentName(1, lpCount), True) & "," & _
        CLng(ParentName(2, lpCount)) & "," & _
        1 & "," & _
        PrintOrder & "," & _
        CLng(ParentName(3, lpCount)) & _
        " )"

    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack
    PrintOrder = PrintOrder + 1
Next lpCount

Dim OPDate As Date
OPDate = "4/1/" & (Year(Now) - IIf(Month(Now) > 3, 0, 1))

''NOw Insert the necessary Sub Heads
'CASH HEAD
'Insert Cash Head
    NewIndexTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        wis_CashHeadID & "," & _
                        AddQuotes(LoadResString(gLangOffSet + 350), True) & "," & _
                        wis_CashParentID & _
                        " )"
    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack
  'Insert Opening Balance
    NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_CashHeadID & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack

'PREVIOUS YEAR PROFIT OR LOSS
    Dim HeadName As String
    HeadName = LoadResString(gLangOffSet + 250) & " " & LoadResString(gLangOffSet + 251) & _
        " " & LoadResString(gLangOffSet + 403) 'Previous Year' Profit
    
    NewIndexTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        wis_PrevProfitHeadID & "," & _
                        AddQuotes(HeadName, True) & "," & _
                        parProfitORLoss & _
                        " )"
    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack
  'Insert Opening Balance
    NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_PrevProfitHeadID & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack

'MISCELENEOUS(INCOME)
'Insert Misceleneous
    NewIndexTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        parIncome + 1 & "," & _
                        AddQuotes(LoadResString(gLangOffSet + 327), True) & "," & _
                        parIncome & _
                        " )"
    'If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack
  'Insert Opening Balance
    NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_IncomeParentID + 1 & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    'If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack

'MISCELENEOUS(EXPENSE)
'Insert Misceleneous
    NewIndexTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        parExpense + 1 & "," & _
                        AddQuotes(LoadResString(gLangOffSet + 327), True) & "," & _
                        parExpense & _
                        " )"
    'If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack
  'Insert Opening Balance
    NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     parExpense + 1 & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    'If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack

NewIndexTrans.CommitTrans

MsgBox "Parent Heads inserted ", vbInformation

Exit Sub

ErrLine:

    MsgBox "InsertParentHeads" & vbCrLf & Err.Description, vbCritical
    'Resume
    Err.Clear
        
End Sub


' This function returns the AccountType From the given ParentID
' Input is ParentId as long
' Returns AccountType long
'
' Lingappa Sindhanur
'
Public Function GetAccountType(ParentID As Long) As wis_AccountType

' Declare Variables
Dim rstParentID As ADODB.Recordset


' Check the Input Received if Zero then Exit
If ParentID = 0 Then Exit Function

'set the sqlstmt
NewIndexTrans.SQLStmt = " SELECT AccountType " & _
                   " FROM ParentHeads " & _
                   " WHERE ParentID=" & ParentID
                   
' Now fetch the record
If NewIndexTrans.Fetch(rstParentID, adOpenForwardOnly) < 0 Then Exit Function

GetAccountType = rstParentID.Fields("AccountType")

End Function

' This function returns the ParentID from the given Headid
' Input is Headid as long
' Returns ParentID long
'
' Pradeep
'
Public Function GetParentID(HeadID As Long) As Long

' Handle Error
On Error GoTo NoParentID:

' Declare Variables
Dim rstParentID As ADODB.Recordset

' Intialiase the Variable
GetParentID = 0

' Check the Input Received if Zero then Exit
If HeadID = 0 Then Exit Function

' set the sqlstmt
NewIndexTrans.SQLStmt = " SELECT ParentID " & _
                   " FROM Heads " & _
                   " WHERE HeadID = " & HeadID
                   
' Now fetch the record
If NewIndexTrans.Fetch(rstParentID, adOpenForwardOnly) < 0 Then Exit Function

' Here is the ParentID!
GetParentID = FormatField(rstParentID("ParentID"))

Exit Function

NoParentID:
    
End Function

