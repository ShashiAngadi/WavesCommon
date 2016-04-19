Attribute VB_Name = "basMaterial"
Option Explicit


' This function will checks the HeadID and retuns the Head NAme
' If the headID is not avaialble in the heads Table it retuns ""
' Inputs :
'           HeadID as Long
' OutPut : HeadName as string
Public Function GetHeadName(ByVal HeadID As Long) As String

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset

'initialise the function
GetHeadName = ""

'Validate the inputs
If HeadID = 0 Then Exit Function

'Check the given Heads in the database
gDbTrans.SQLStmt = " SELECT HeadName FROM BankHeadIds " & _
                   " WHERE Headid = " & HeadID
                
'if exists then exit function
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then _
    GetHeadName = FormatField(rstHeads.Fields(0))

Set rstHeads = Nothing

Exit Function

ErrLine:
    MsgBox "GetHeadID: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Set rstHeads = Nothing
    
    Exit Function
End Function


' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it retuns 0
' Inputs :
'           HeadName as String
' OutPut : Headid
Public Function GetIndexHeadID(ByVal HeadName As String) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset

'initialise the function
GetIndexHeadID = 0

'Validate the inputs
If HeadName = "" Then Exit Function

'Check the given Heads in the database
'gDbTrans.SQLStmt = " SELECT HeadID FROM Heads " & _
            " WHERE HeadName = " & AddQuotes(HeadName, True)
gDbTrans.SQLStmt = " SELECT HeadID FROM BankHeadIds " & _
                   " WHERE HeadName = " & AddQuotes(HeadName, True)
                
'if exists then exit function
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then _
    GetIndexHeadID = FormatField(rstHeads.Fields(0))

Set rstHeads = Nothing

Exit Function

ErrLine:
    MsgBox "GetHeadID: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Set rstHeads = Nothing
    
    Exit Function
End Function

' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it retuns 0
' Inputs :
'           HeadName as String
' OutPut : Headid
Public Function GetHeadID(ByVal HeadName As String, ParentID As Long) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset

'initialise the function
GetHeadID = 0

'Validate the inputs
If HeadName = "" Then Exit Function

'Check the given Heads in the database
gDbTrans.SQLStmt = " SELECT HeadID FROM Heads " & _
            " WHERE HeadName = " & AddQuotes(HeadName, True) & _
            " And ParentID = " & ParentID

'If ParentId Then _
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & " And ParentID = " & ParentId
            
'gDbTrans.SQLStmt = " SELECT HeadID FROM BankHeadIds " & _
                   " WHERE HeadName = " & AddQuotes(HeadName, True)
                
'if exists then exit function
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then _
    GetHeadID = FormatField(rstHeads.Fields(0))

If rstHeads.RecordCount > 1 Then GetHeadID = 0
Set rstHeads = Nothing

Exit Function

ErrLine:
    MsgBox "GetHeadID: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Set rstHeads = Nothing
    
    Exit Function
End Function


'
Public Function IsBankHead(HeadID As Long) As Boolean

Dim rstTemp As Recordset

gDbTrans.SQLStmt = "SELECT * FRom ParentHeads " & _
        " WHERE ParentID = (SELECT ParentID " & _
            " From Heads Where HeadID = " & HeadID & " )"

If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Function

If FormatField(rstTemp("UserCreated")) > 2 Then IsBankHead = True
Set rstTemp = Nothing
  
End Function

'This Function Checks whetehr the given head id is created by user
'or created by system(i.e. head is predifined
Public Function IsUserCreatedHead(HeadID As Long) As Boolean

Dim rstTemp As Recordset

gDbTrans.SQLStmt = "SELECT * FRom ParentHeads " & _
            " WHERE ParentID = (SELECT ParentID " & _
                " From Heads Where HeadID = " & HeadID & " )"

If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Function
If FormatField(rstTemp("UserCreated")) Mod 2 = 0 Then IsUserCreatedHead = True

Set rstTemp = Nothing

End Function


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
Dim PurAndSale As Boolean
Dim Rst As Recordset
gDbTrans.SQLStmt = "Select UserCreated From ParentHeads " & _
        "Where parentId = " & parPurchase
'If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
    If FormatField(Rst("UserCreated")) < 3 Then PurAndSale = True

' Start wtih Statement
With cmbVouch
    ' Start Adding Vocuhers
    .AddItem LoadResString(gLangOffSet + 196) '"Receipt"
    .ItemData(.NewIndex) = 1
    .AddItem LoadResString(gLangOffSet + 197) '"Payment"
    .ItemData(.NewIndex) = 2
    
    If PurAndSale Then
        .AddItem LoadResString(gLangOffSet + 176) '"Purchase"
        .ItemData(.NewIndex) = 3
        .AddItem LoadResString(gLangOffSet + 180) '"Sales"
        .ItemData(.NewIndex) = 4
        .AddItem LoadResString(gLangOffSet + 105) '"Free"
        .ItemData(.NewIndex) = 5
    End If
    
''''''''The Following  Block is changed by shashi on 17/12/2002
    'As the old Index 2000 is having only two type transction
    ' i.e. 1) Deposit( receipt), 2) Withdraw(payment)
    'But in this new Index 2000 we introduced two More transction
    'of the same concept called contraDeposit, contraWithdraw
    'where Physical Cash is not handled or the internal transfer
    'is called Contra Transction
    'So to faclitate the Index 2000 concept Here we changed the
    'Tranction Type n the Cobo Box
    'And we are keeping the Same Voucher Type as in theis Project
    'We may change it later
    'And the one major change may be The
    'Contra Transaction Of this Project will become Cash Transction
    
    'Previous Code
''    .AddItem LoadResString(gLangOffSet + 198) '"Journal"
''    .ItemData(.NewIndex) = 6
''    .AddItem LoadResString(gLangOffSet + 270) '"Contra"
''    .ItemData(.NewIndex) = 7
    'NEW CODE
    .AddItem LoadResString(gLangOffSet + 270) '"Contra"
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
Public Sub LoadLedgersToCombo(cmbLedger As ComboBox, _
        ByVal ParentID As Long, Optional ClearCombo As Boolean = True)

' Handle Errors
On Error GoTo NoLoadLedgers:

' Declarations
Dim rstHeads As ADODB.Recordset

' Check if Variable is empty
If ParentID = 0 Then Exit Sub

' clear the combobox
If ClearCombo Then cmbLedger.Clear

' Set the Sql Statement
gDbTrans.SQLStmt = " SELECT HeadID,HeadName" & _
                   " FROM Heads WHERE ParentID =  " & ParentID & _
                   " ORDER BY HeadName"

' Fetch the Records
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) < 1 Then Exit Sub
' Start the Loop
Do While Not rstHeads.EOF
    With cmbLedger
        'Add the item to combo
        .AddItem FormatField(rstHeads("HeadName"))
        ' Set the itemdata
        .ItemData(.NewIndex) = FormatField(rstHeads("HeadID"))
    End With
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
gDbTrans.SQLStmt = " SELECT HeadID,HeadName" & _
                   " FROM Heads A,ParentHeads B" & _
                   " WHERE A.ParentID = B.ParentID" & _
                   " AND B.AccountType=" & AccountType & _
                   " AND B.UserCreated <= 2 " & _
                   " ORDER BY HeadName"

' Fetch the Records

If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) < 1 Then Exit Sub


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
Public Sub LoadParentHeads(ctrlComboBox As ComboBox, _
                        Optional LoadIndexHeads As Boolean = False)

On Error GoTo NoLoadParents:

Dim rstParent As ADODB.Recordset

ctrlComboBox.Clear

gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
                   " WHERE UserCreated <= 2" & _
                   " ORDER BY AccountType,ParentName "
If (gCurrUser.UserPermissions And perOnlyWaves) Or LoadIndexHeads Then
    gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
                   " ORDER BY AccountType,ParentName "
End If

Call gDbTrans.Fetch(rstParent, adOpenForwardOnly)

If rstParent Is Nothing Or (rstParent.EOF And rstParent.BOF) Then
    InsertParentHeads

gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
                   " ORDER BY ParentName "
'WHERE UserCreated <= 2
  
  Call gDbTrans.Fetch(rstParent, adOpenForwardOnly)
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

'
Private Sub InsertParentHeads()

Dim Rst As Recordset
Dim ParentName() As String
Dim AccountType As wis_AccountType
'Trap an error
On Error GoTo ErrLine

gDbTrans.SQLStmt = " SELECT * FROM ParentHeads"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then Exit Sub

ReDim Preserve ParentName(3, 20)
Dim I As Integer

I = 0
AccountType = Liability
'Share Capital
ParentName(0, I) = parShareCapital
ParentName(1, I) = LoadResString(gLangOffSet + 351)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Memebr Share
ParentName(0, I) = parMemberShare
ParentName(1, I) = LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 53)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Reserve and Suplus funds
ParentName(0, I) = parReserveFunds
ParentName(1, I) = LoadResString(gLangOffSet + 352)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1

'''Deposit (Liability)
''ParentName(0, i) = parDepositLiab
''ParentName(1, i) = LoadResString(gLangOffSet + 45)
''ParentName(2, i) = AccountType
''ParentName(3, i) = 1
''i = i + 1

'Memebr Deposit
ParentName(0, I) = parMemberDeposit
ParentName(1, I) = LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 45)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1   '5

'Bank LOan Accounts
ParentName(0, I) = parBankLoanAccount
ParentName(1, I) = LoadResString(gLangOffSet + 356)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Govt Loan subsidy
ParentName(0, I) = parGovtLoanSubsidy
ParentName(1, I) = LoadResString(gLangOffSet + 263)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1

'Other Loans
ParentName(0, I) = parOtherLoans
ParentName(1, I) = LoadResString(gLangOffSet + 237) & " " & LoadResString(gLangOffSet + 18)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Other Payables
ParentName(0, I) = parPayAble
ParentName(1, I) = LoadResString(gLangOffSet + 357)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Deposit Interest Provision
ParentName(0, I) = parDepositIntProv
ParentName(1, I) = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 450)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1   '10
'Loan Interest Provision
ParentName(0, I) = parLoanIntProv
ParentName(1, I) = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 450)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Suspence Account
ParentName(0, I) = parSuspAcc
ParentName(1, I) = LoadResString(gLangOffSet + 365)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Account to vcarry forward the current years profit or loss
'to the next Finaancial year
'so in the next financial yaer it will be distributed to the funds
' And till then head is called previousyearProfit( Or Loss)
'Profit & Loss Account
ParentName(0, I) = parProfitORLoss
ParentName(1, I) = LoadResString(gLangOffSet + 443) & " " & LoadResString(gLangOffSet + 36)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1


'''ASSETS
AccountType = Asset
'Cash in Hand
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parCash
ParentName(1, I) = LoadResString(gLangOffSet + 350)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Bank Accounts
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parBankAccount
ParentName(1, I) = LoadResString(gLangOffSet + 359)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Investments (Assets)
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parInvestment
ParentName(1, I) = LoadResString(gLangOffSet + 361)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1   '15

'LOan And Advances (Assets)
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parLoanAdvanceAsset
ParentName(1, I) = LoadResString(gLangOffSet + 360)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Member Loans
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parMemberLoan
ParentName(1, I) = LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 18)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

'Member Deposit Loans
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parMemDepLoan
ParentName(1, I) = LoadResString(gLangOffSet + 49) & " " & _
    LoadResString(gLangOffSet + 53) & " " & LoadResString(gLangOffSet + 18)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

'Salary Advance
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parSalaryAdvance
ParentName(1, I) = LoadResString(gLangOffSet + 90) & " " & _
    LoadResString(gLangOffSet + 355)
ParentName(2, I) = AccountType
ParentName(3, I) = 2
I = I + 1


'Fixed Assets
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parFixedAsset
ParentName(1, I) = LoadResString(gLangOffSet + 363)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'ReceivAbles
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parReceivable
ParentName(1, I) = LoadResString(gLangOffSet + 364)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1

''INCOME HEADS
AccountType = Profit
'Income
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parIncome
ParentName(1, I) = LoadResString(gLangOffSet + 366)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1   '20
'Trading Income
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parTradingIncome
ParentName(1, I) = LoadResString(gLangOffSet + 367)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Regual Interest received on Member LOans
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parMemLoanIntReceived
ParentName(1, I) = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 344)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Penal Interest received on Member LOans
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parMemLoanPenalInt
ParentName(1, I) = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 345)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
  
'Interest received on deposit Loans
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parMemDepLoanIntReceived
ParentName(1, I) = LoadResString(gLangOffSet + 43) & " " & _
        LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 483)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

'Interest received on Deposits
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parDepIntReceived
ParentName(1, I) = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 483)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

'Income other Bank incomes
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parBankIncome
ParentName(1, I) = LoadResString(gLangOffSet + 418) & " " & _
        LoadResString(gLangOffSet + 366)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

'''EXPENSE
AccountType = Loss
'Expense
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parExpense
ParentName(1, I) = LoadResString(gLangOffSet + 368) '& " " & LoadResString(gLangOffSet + 36)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1
'Trading expense
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parTradingExpense
ParentName(1, I) = LoadResString(gLangOffSet + 369)
ParentName(2, I) = AccountType
ParentName(3, I) = 1
I = I + 1   '25
'Interest paid on Member Deposit
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parMemDepIntPaid
ParentName(1, I) = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 487)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Interest paid on Loans
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parLoanIntPaid
ParentName(1, I) = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 487)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Other expenditure in Bank accounts
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parBankExpense
ParentName(1, I) = LoadResString(gLangOffSet + 418) & " " & LoadResString(gLangOffSet + 368)
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

'Salary expense
ReDim Preserve ParentName(3, I + 1)
ParentName(0, I) = parSalaryExpense
ParentName(1, I) = LoadResString(gLangOffSet + 90) & " " & _
    LoadResString(gLangOffSet + 36)
ParentName(2, I) = AccountType
ParentName(3, I) = 2
I = I + 1


'Purchase and Sales Account
AccountType = ItemPurchase
'Purchase
ReDim Preserve ParentName(3, I)
ParentName(0, I) = parPurchase
ParentName(1, I) = LoadResString(gLangOffSet + 176) & " " & LoadResString(gLangOffSet + 36)  ''"Purchase Account"
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1
'Sales account
ReDim Preserve ParentName(3, I)
AccountType = ItemSales
ParentName(0, I) = parSales
ParentName(1, I) = LoadResString(gLangOffSet + 180) & " " & LoadResString(gLangOffSet + 36) '"Sales Account"
ParentName(2, I) = AccountType
ParentName(3, I) = 3
I = I + 1

gDbTrans.BeginTrans

Dim MaxCount As Integer
Dim lpCount As Integer
Dim PrevType As wis_AccountType
Dim PrintOrder As Integer

PrintOrder = 1
MaxCount = I - 1

For lpCount = 0 To MaxCount
    'Change the print order as the acount type changes
    If PrevType = Val(ParentName(0, lpCount)) Then PrintOrder = 1
    
    gDbTrans.SQLStmt = " INSERT INTO ParentHeads " & _
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

    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    PrintOrder = PrintOrder + 1
Next lpCount

Dim OPDate As Date
OPDate = GetSysFormatDate("1/4/" & (Year(gStrDate) - IIf(Month(gStrDate) > 3, 0, 1)))

''NOw Insert the necessary Sub Heads
'CASH HEAD
Dim SubHeadName As String
'Insert Cash Head
    SubHeadName = LoadResString(gLangOffSet + 350)
    gDbTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        wis_CashHeadID & "," & _
                        AddQuotes(SubHeadName, True) & "," & _
                        wis_CashParentID & _
                        " )"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
  'Insert Opening Balance
    gDbTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_CashHeadID & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

'PREVIOUS YEAR PROFIT OR LOSS
    SubHeadName = LoadResString(gLangOffSet + 250) & " " & LoadResString(gLangOffSet + 251) & _
        " " & LoadResString(gLangOffSet + 403) 'Previous Year' Profit
    gDbTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        wis_PrevProfitHeadID & "," & _
                        AddQuotes(SubHeadName, True) & "," & _
                        parProfitORLoss & _
                        " )"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
  'Insert Opening Balance
    gDbTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_PrevProfitHeadID & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

'MISCELENEOUS(INCOME)
'Insert Misceleneous
    gDbTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        wis_IncomeParentID + 1 & "," & _
                        AddQuotes(LoadResString(gLangOffSet + 327), True) & "," & _
                        wis_IncomeParentID & _
                        " )"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
  'Insert Opening Balance
    gDbTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_IncomeParentID + 1 & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

'MISCELENEOUS(EXPENSE)
'Insert Misceleneous
    gDbTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID )" & _
                      " VALUES ( " & _
                        wis_ExpenseParentID + 1 & "," & _
                        AddQuotes(LoadResString(gLangOffSet + 327), True) & "," & _
                        wis_ExpenseParentID & _
                        " )"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
  'Insert Opening Balance
    gDbTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     wis_ExpenseParentID + 1 & "," & _
                     "#" & OPDate & "#," & _
                     0 & ")"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

MsgBox "Parent Heads inserted ", vbInformation

Exit Sub

ErrLine:
        MsgBox "InsertParentHeads" & vbCrLf & Err.Description, vbCritical
        'Resume
        Err.Clear
End Sub

Public Sub CreateDefaultView()

Call gDbTrans.DeleteAllViews

gDbTrans.SQLStmt = "SELECT A.TransID,A.HeadID,A.Debit,A.Credit,B.HeadID," & _
                " B.Debit,B.Credit,B.TransDate,B.VoucherType" & _
                " FROM AccTrans AS A" & _
                " INNER JOIN AccTrans AS B ON A.TransID=B.TransID "
    
gDbTrans.CreateView ("QryAccTransMerge1")
gDbTrans.SQLStmt = "SELECT A.TransID, A.TransDate, A.HeadID, B.HeadID," & _
            " A.Credit, A.Debit, A.VoucherType " & _
            " FROM AccTrans AS A, AccTrans AS B, BankHeadIds AS C " & _
            " WHERE A.TransID=B.TransID And A.HeadID=C.HeadID;"
gDbTrans.CreateView ("QryAccBankTrans")

gDbTrans.SQLStmt = "SELECT A.TransID, A.HeadID, A.Debit, A.Credit, " & _
            " B.HeadID, B.Debit, B.Credit, B.TransDate, B.VoucherType " & _
            " FROM AccTrans AS A INNER JOIN AccTrans AS B ON A.TransID = B.TransID;"
gDbTrans.CreateView ("QryAccTransMerge")

gDbTrans.SQLStmt = "SELECT CustomerID,FirstNAme + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,Place,Caste,Gender,IsciName From NameTab"
gDbTrans.CreateView ("QryName")

gDbTrans.SQLStmt = "SELECT CustomerID,FirstNAme + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,IsciName From NameTab"
gDbTrans.CreateView ("QryOnlyName")



End Sub

' This function gives the purchase price of the item from the database
' Inputs : RelationID as long
' Output : Gives the Purchase Price of the Item
Public Sub CreateItemPriceQuery(ByVal FromDate As String, _
                                 ByVal ToDate As String, _
                                 qryName As String)
'Trap an error
On Error GoTo ErrLine
'declare Variables

Dim VoucherType As Wis_VoucherTypes
Dim eVoucherType As Wis_VoucherTypes

VoucherType = Purchase
eVoucherType = StockIn

gDbTrans.SQLStmt = " SELECT Max(TransID) AS maxTransID, RelationID" & _
                   " FROM Stock " & _
                   " WHERE (VoucherType = " & VoucherType & _
                   " OR VoucherType = " & eVoucherType & " ) " & _
                   " AND TransDate BETWEEN #" & GetSysFormatDate(FromDate) & "#" & _
                   " AND #" & GetSysFormatDate(ToDate) & "#" & _
                   " GROUP BY RelationID"
                   
Call gDbTrans.CreateView("QryMaxTransID")

'Fire the Query
gDbTrans.SQLStmt = " SELECT UnitPrice,a.RelationID " & _
                   " FROM Stock a,qryMaxTransID b " & _
                   " WHERE a.TransID=b.maxTransID " & _
                   " AND a.RelationID=b.RelationID"

gDbTrans.CreateView ("qryPrice")

gDbTrans.SQLStmt = " SELECT UnitPrice as UnitPrice1,A.RelationID " & _
                   " FROM qryPrice A " & _
                   " WHERE UnitPrice > 0"
gDbTrans.CreateView ("qryPrice1")

'There are Some Items which will have purchase price will be 0
'Like Container whih will come free with some materails
'and User Sales and Gets some amount From It
'So For Such Items Get the Item Price from Sales details
'Some thng Like Max Price

'So get the the Free items from getting whose purchase price is 0
gDbTrans.SQLStmt = " SELECT Max(UnitPrice) As UnitPrice1,RelationID " & _
                   " FROM Stock A " & _
                   " WHERE RelationID IN (SELECT RelationID " & _
                   " FROM qryPrice WHERE UnitPrice = 0) " & _
                   " Group BY RelationID"
gDbTrans.CreateView ("qryPrice2")

gDbTrans.SQLStmt = " Select RelationID,UnitPrice1 as UnitPrice" & _
                   " From qryPrice1 " & _
                   " UNION " & _
                   " Select RelationID,UnitPrice1 as UnitPrice" & _
                   " From qryPrice2 "

gDbTrans.CreateView (qryName)
'If gDbTrans.Fetch(rstPrice, adOpenDynamic) < 1 Then Exit Sub

Exit Sub

ErrLine:
    
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
gDbTrans.SQLStmt = " SELECT AccountType " & _
                   " FROM ParentHeads " & _
                   " WHERE ParentID=" & ParentID
                   
' Now fetch the record
If gDbTrans.Fetch(rstParentID, adOpenForwardOnly) < 1 Then Exit Function

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
gDbTrans.SQLStmt = " SELECT ParentID " & _
                   " FROM Heads " & _
                   " WHERE HeadID=" & HeadID
                   
' Now fetch the record
If gDbTrans.Fetch(rstParentID, adOpenForwardOnly) < 1 Then Exit Function

' Here is the ParentID!
GetParentID = FormatField(rstParentID("ParentID"))

Exit Function

NoParentID:
    
End Function


