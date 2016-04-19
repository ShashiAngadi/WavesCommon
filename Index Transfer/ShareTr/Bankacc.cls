VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBankAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IsOnLine As Boolean

Public Function Balance(ByVal AsOnIndiandate As String, ByVal HeadID As Long) As Currency
#If AccTrans Then
'Trap an Error
On Error GoTo ErrLine

Dim AccTransClass As clsAccTrans

Set AccTransClass = New clsAccTrans

Balance = AccTransClass.GetOpBalance(HeadID, AsOnIndiandate)

Set AccTransClass = Nothing

Exit Function

ErrLine:
    
    MsgBox "Balance: " & vbCrLf & Err.Description, vbCritical
#End If
End Function

Public Function CreateCallDeposits(HeadID() As Long, ByVal HeadName As String) As Long

Dim ParentID As Long

ReDim HeadID(0)

''    Select Case IndexIds
''
''        Case IndexIds = DepositBKCC, DepositCA, DepositPigmy, DepositRD, DepositSB
''
''            ParentID = wis_Deposits
''
''        Case IndexIds = LoansBKCC, LoansDeposit, LoansNonAgri, LoansPigmy, LoansRD
''
''            ParentID = wis_Loans
''
''        Case IndexIds = LossDepositBKCC, LossDepositCA, LossDepositPigmy, LossDepositRD, LossDepositSB, _
''                        LossLoansBKCC, LossLoansDeposit, LossLoansNonAgri, LossLoansPigmy, LossLoansRD
''
''            ParentID = wis_ExpenseID
''
''        Case IndexIds = PayAbleDepositPigmy, PayAbleDepositRD
''
''            ParentID = wis_Payables
''
''        Case IndexIds = ProfitDepositBKCC, ProfitDepositCA, ProfitDepositPigmy, ProfitDepositRD, ProfitDepositSB, _
''                        ProfitLoansBKCC, ProfitLoansDeposit, ProfitLoansNonAgri, ProfitLoansPigmy, ProfitLoansRD
''
''            ParentID = wis_IncomeID
''
''    End Select

'''COMMENTED BY SHASHI ''12/12/ like "'''"
'''' Create Deposit
'''ParentID = wis_Deposits
'''' Create DepositLoans
'''ParentID = wis_Loans
'''' Create Profit
'''ParentID = wis_IncomeID
'''' Create Loss
'''ParentID = wis_ExpenseID
'''' Create Payable
'''ParentID = wis_Payables

End Function
Public Function CreateDemandDeposits() As Boolean

End Function


Public Function CreateLoans() As Boolean

End Function


Private Function InsertIntoIndexIds(ByVal IndexID As Long, ByVal HeadID As Long) As Boolean

InsertIntoIndexIds = False

If IndexID = 0 Then Exit Function
If HeadID = 0 Then Exit Function

NewIndexTrans.SQLStmt = " INSERT INTO IndexIDs (IndexID,MaterialID) " & _
                 " VALUES ( " & _
                 IndexID & "," & _
                 HeadID & ")"

If Not NewIndexTrans.SQLExecute Then Exit Function

InsertIntoIndexIds = True

End Function

Public Function Loss(ByVal FromIndianDate As String, ByVal ToIndianDate As String, ByVal HeadID As Long) As Currency

' To Trap an Error
On Error GoTo ErrLine

'Declare variables
Dim AccountType As wis_AccountType
Dim rstLoss As ADODB.Recordset

Loss = 0

' Validations
If HeadID = 0 Then Exit Function
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

' set the accounttype
AccountType = Loss

' This is the sql to fetch the profits

NewIndexTrans.SQLStmt = " SELECT b.Credit as LossReceived, b.Debit as LossGiven, e.AccountType " & _
                   " FROM qryAccTransmerge AS d, Heads AS c, ParentHeads AS e " & _
                   " WHERE a.HeadId = " & HeadID & _
                   " AND b.HeadID<>" & HeadID & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromIndianDate) & "#" & _
                   " AND #" & FormatDate(ToIndianDate) & "#" & _
                   " AND b.HeadID=c.HeadID " & _
                   " AND c.ParentID=e.ParentID " & _
                   " AND e.AccountType=" & AccountType & _
                   " ORDER BY a.TransDate, a.TransID "
                       
' this will create a view
Call NewIndexTrans.CreateView("QryHeadIDLoss")

' this will fetch the details

NewIndexTrans.SQLStmt = " SELECT (SUM(LossGiven)-SUM(LossReceived)) as Loss" & _
                   " FROM qryHeadIDLoss"
                   
Call NewIndexTrans.Fetch(rstLoss, adOpenForwardOnly)


Loss = rstLoss.Fields("Loss")

Exit Function

ErrLine:
    MsgBox "Loss: " & vbCrLf & Err.Description, vbCritical
        
End Function



'Private Function TransactionDetails(FromIndianDate As String, ToIndianDate As String, HeadId As Long, TransType As Wis_VoucherTypes) As Currency
'
''Trap an error
'On Error GoTo ErrLine
'
'    gDbTrans.SqlStmt = " Select Sum(Amount) as Deposits from ACCTrans where " & _
'        " AccId = " & HeadId & " And TransType = " & TransType & _
'        " AND TransDate Between #" & FormatDate(FromIndianDate) & "#" & _
'        " And #" & FormatDate(ToIndianDate) & "#"
Public Function UpdateContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As String) As Boolean

UpdateContraTrans = False

If m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function

If Not DateValidate(TransDate, "/", True) Then Exit Function

VoucherType = Journal

Set AccTransClass = New clsAccTrans
                
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

UpdateContraTrans = True

Exit Function

Hell:
    
    UpdateContraTrans = False
    
End Function
Public Function AddContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As String) As Boolean

AddContraTrans = False

If Not m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function

If Not DateValidate(TransDate, "/", True) Then Exit Function

VoucherType = Journal

Set AccTransClass = New clsAccTrans
                
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.SaveVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

AddContraTrans = True

Exit Function

Hell:
    
    AddContraTrans = False
    
End Function

Public Function UpdateCashDeposits(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean


UpdateCashDeposits = False

If m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If HeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function

If Not DateValidate(CStr(TransDate), "/", True) Then Exit Function

CrHeadID = HeadID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

Set AccTransClass = New clsAccTrans
                
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

UpdateCashDeposits = True

Exit Function

Hell:
    
    UpdateCashDeposits = False
    
End Function
Public Function AddCashDeposits(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddCashDeposits = False

If Not m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function


CrHeadID = HeadID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

Set AccTransClass = New clsAccTrans

If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.SaveVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

AddCashDeposits = True

Exit Function

Hell:
    
    AddCashDeposits = False
    
End Function

Public Function UndoCashDeposits(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal IndianTransDate As String) As Boolean

UndoCashDeposits = False

If m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If HeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function

If Not DateValidate(IndianTransDate, "/", True) Then Exit Function

CrHeadID = HeadID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

Set AccTransClass = New clsAccTrans
                
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UndoVouchers(VoucherType, IndianTransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

UndoCashDeposits = True

Exit Function

Hell:
    
    UndoCashDeposits = False
    
End Function

Public Function UpdateCashWithDrawls(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As String) As Boolean

UpdateCashWithDrawls = False

If m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function

If Not DateValidate(TransDate, "/", True) Then Exit Function

DrHeadID = HeadID
CrHeadID = wis_CashHeadID

VoucherType = Payment

Set AccTransClass = New clsAccTrans
                
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

UpdateCashWithDrawls = True

Exit Function

Hell:
    
    UpdateCashWithDrawls = False
    
End Function
Public Function AddCashWithDrawls(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddCashWithDrawls = False

'If Not m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function


CrHeadID = wis_CashHeadID
DrHeadID = HeadID

VoucherType = Payment

Set AccTransClass = New clsAccTrans

If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.SaveVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

AddCashWithDrawls = True

Exit Function

Hell:
    
    AddCashWithDrawls = False
    
End Function


Public Function UndoBalances(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal IndianTransDate As String, _
                                ByVal VoucherType As Wis_VoucherTypes) As Boolean
UndoBalances = False

If m_IsOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then Exit Function
If Not DateValidate(IndianTransDate, "/", True) Then Exit Function

Set AccTransClass = New clsAccTrans

If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UndoVouchers(VoucherType, IndianTransDate) <> Success Then Exit Function

UndoBalances = True

Exit Function

Hell:
    UndoBalances = False
    
End Function
''
''If gDbTrans.SQLFetch < 1 Then Exit Function
''
''TransactionDetails = Val(FormatField(gDbTrans.Rst(0)))
''
''ExitLine:
''    Exit Function
''
''
''ErrLine:
''    If Err Then
''        MsgBox "TransactionDetails :" & vbCrLf & Err.Description
''        GoTo ExitLine
''    End If
''
''End Function
''
Public Function WithDrawls(ByVal FromIndianDate As String, ByVal ToIndianDate As String, ByVal HeadID As Long) As Currency

On Error GoTo Err_Line:

Dim rstWithdrawal As ADODB.Recordset

WithDrawls = 0

If HeadID = 0 Then Exit Function
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

NewIndexTrans.SQLStmt = " SELECT Sum(Credit) as WithDrawal " & _
                   " FROM AccTrans " & _
                   " WHERE HeadID=" & HeadID & _
                   " AND TransDate BETWEEN " & "#" & FormatDate(FromIndianDate) & "#" & _
                   " AND #" & FormatDate(ToIndianDate) & "#"
                  
Call NewIndexTrans.Fetch(rstWithdrawal, adOpenForwardOnly)
  
WithDrawls = rstWithdrawal.Fields("WithDrawal")
  
Exit Function

Err_Line:

    MsgBox "WithDrawls : " & Err.Description
        
End Function
Public Function Profit(ByVal FromIndianDate As String, ByVal ToIndianDate As String, ByVal HeadID As Long) As Currency

' To trap an error
On Error GoTo ErrLine

'Declare variables
Dim AccountType As wis_AccountType
Dim rstProfit As ADODB.Recordset

Profit = 0

' Validations
If HeadID = 0 Then Exit Function
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

' set the accounttype
AccountType = Profit

' This is the sql to fetch the profits

NewIndexTrans.SQLStmt = " SELECT b.Credit as ProfitReceived, b.Debit as ProfitGiven, e.AccountType " & _
                   " FROM qryAccTransmerge AS d, Heads AS c, ParentHeads AS e " & _
                   " WHERE a.HeadId = " & HeadID & _
                   " AND b.HeadID<>" & HeadID & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromIndianDate) & "#" & _
                   " AND #" & FormatDate(ToIndianDate) & "#" & _
                   " AND b.HeadID=c.HeadID " & _
                   " AND c.ParentID=e.ParentID " & _
                   " AND e.AccountType=" & AccountType & _
                   " ORDER BY a.TransDate, a.TransID "
                       
Call NewIndexTrans.CreateView("QryHeadIDProfit")

NewIndexTrans.SQLStmt = " SELECT (SUM(ProfitReceived)-SUM(ProfitGiven)) as Profits" & _
                   " FROM qryHeadIDProfit"
                   
Call NewIndexTrans.Fetch(rstProfit, adOpenForwardOnly)

Profit = rstProfit.Fields("Profits")

ErrLine:
    
    MsgBox "Profit: " & vbCrLf & Err.Description, vbCritical
    
End Function


Public Function Deposits(FromIndianDate As String, ToIndianDate As String, HeadID As Long) As Currency

Dim rstDeposits As ADODB.Recordset

Deposits = 0

If HeadID = 0 Then Exit Function
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

NewIndexTrans.SQLStmt = " SELECT Sum(Debit) as Deposits " & _
                   " FROM AccTrans " & _
                   " WHERE HeadID=" & HeadID & _
                   " AND TransDate BETWEEN " & "#" & FormatDate(FromIndianDate) & "#" & _
                   " AND #" & FormatDate(ToIndianDate) & "#"
                 
Call NewIndexTrans.Fetch(rstDeposits, adOpenForwardOnly)
  
Deposits = rstDeposits.Fields("Deposits")

End Function
Public Function MainHeads(HeadsName() As String, HeadID() As Long) As Integer

On Error GoTo Hell:

Dim RetLng As Long
Dim Rst As ADODB.Recordset

ReDim HeadNames(0)
ReDim HeadIds(0)
    
NewIndexTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
                   " WHERE ParentID mod " & HEAD_OFFSET & "=0" & _
                   " ORDER BY ParentName "

RetLng = NewIndexTrans.Fetch(Rst, adOpenStatic)

If RetLng < 1 Then Exit Function
    
ReDim HeadNames(RetLng - 1)
ReDim HeadIds(RetLng - 1)

RetLng = 0

Do While Rst.EOF
    HeadNames(RetLng) = FormatField(Rst("ParentName"))
    HeadIds(RetLng) = FormatField(Rst("ParentID"))
    RetLng = RetLng + 1
    Rst.MoveNext
Loop

MainHeads = RetLng
Exit Function

Hell:
    
    MsgBox "MainHeads :" & vbCrLf & Err.Description
    
End Function
Public Function GetHeadNames(HeadNames() As String, HeadIds() As Long, ByVal ParentID As Long) As Integer

Dim RetLng As Long
Dim Rst As ADODB.Recordset

If ParentID = 0 Then Exit Function

ReDim HeadNames(0)
ReDim HeadIds(0)
    
NewIndexTrans.SQLStmt = " SELECT HeadName,HeadID " & _
                   " FROM Heads " & _
                   " WHERE ParentID=" & ParentID
                   
RetLng = NewIndexTrans.Fetch(Rst, adOpenDynamic)
If RetLng < 1 Then Exit Function
    
ReDim HeadNames(RetLng - 1)
ReDim HeadIds(RetLng - 1)

RetLng = 0
Do While Not Rst.EOF
    HeadNames(RetLng) = FormatField(Rst("HeadName"))
    HeadIds(RetLng) = FormatField(Rst("HeadId"))
    RetLng = RetLng + 1
    Rst.MoveNext
Loop

GetHeadNames = RetLng
    
End Function

' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it will create the
' Head with given HeadID & parentID
' Inputs :
'           HeadName as String
'           ParentID as long
' OutPut :
Public Function GetHeadIDCreatedONParentID(ByVal ParentID As Long, ByVal HeadName As String, ByVal OpBalance As Currency) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset
Dim USOpDate As String
Dim HeadID As Long



'initialise the function
GetHeadIDCreatedONParentID = -1

'Validate the inputs
If ParentID = 0 Then Exit Function
If HeadName = "" Then Exit Function

'Check the given Heads in the database
NewIndexTrans.SQLStmt = " SELECT HeadID" & _
                   " FROM Heads " & _
                   " WHERE HeadName = " & AddQuotes(HeadName, True) & _
                   " AND ParentID = " & ParentID
                
'if exists then exit function
If NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then
    GetHeadIDCreatedONParentID = FormatField(rstHeads.Fields(0))
    Exit Function
End If

'Get the Maximum Head From the database
NewIndexTrans.SQLStmt = " SELECT MAX(HeadID)" & _
                   " FROM Heads " & _
                   " WHERE ParentID = " & ParentID

Call NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly)

HeadID = FormatField(rstHeads.Fields(0)) + 1
    
If HeadID < ParentID Then HeadID = HeadID + ParentID

' Else create the head with given headID & parentiD
'Insert the heads into the database
NewIndexTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID) " & _
                  " VALUES ( " & _
                  HeadID & "," & _
                  AddQuotes(HeadName, True) & "," & _
                  ParentID & ")"

If Not NewIndexTrans.SQLExecute Then Exit Function


USOpDate = FormatDate(FinIndianFromDate)

If Not DateValidate(USOpDate, "/", False) Then USOpDate = "4/1/" & IIf(Month(Now) < 3, Year(Now) - 1, Year(Now))
'Insert the opening balance
NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 "#" & USOpDate & "#," & _
                 OpBalance & ")"
If Not NewIndexTrans.SQLExecute Then Exit Function

GetHeadIDCreatedONParentID = HeadID

Set rstHeads = Nothing

Exit Function

ErrLine:
    MsgBox "CreateHead: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    GetHeadIDCreatedONParentID = -1
    
    Set rstHeads = Nothing
    
    Exit Function
End Function
' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it will create the
' Head with given HeadID & parentID
' Inputs :
'           HeadID as long
'           WIs_IndexIDs
' OutPut :
Public Function GetHeadIDCreatedOnEnum(ByVal IndexIds As WIS_IndexIDs, ByVal HeadName As String, ByVal OpBalance As Currency) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset
Dim USOpDate As String
Dim HeadID As Long
Dim ParentID As Long

' Initialise the function
GetHeadIDCreatedOnEnum = -1

' Validate the inputs

If IndexIds = 0 Then Exit Function

Select Case IndexIds

    Case DepositBKCC, DepositCA, DepositPigmy, DepositRD, DepositSB
        
        ParentID = parMemberDeposit
    
    Case LoansBKCC, LoansNonAgri
        
        ParentID = parMemberLoan
    
    Case LoansDeposit, LoansPigmy, LoansRD
    
            ParentID = parMemDepLoan
    
    Case LossDepositBKCC, LossDepositCA, LossLoansBKCC, LossLoansNonAgri
        
        ParentID = parLoanIntPaid
    
    Case LossDepositPigmy, LossDepositRD, LossDepositSB, LossLoansDeposit, LossLoansPigmy, LossLoansRD
    
    Case PayAbleDepositPigmy, PayAbleDepositRD
        
        ParentID = parDepositIntProv
        
    Case ProfitLoansBKCC, ProfitLoansDeposit, ProfitLoansNonAgri
        
        ParentID = parMemLoanIntReceived
    
    Case ProfitDepositBKCC, ProfitDepositCA, ProfitDepositPigmy, ProfitDepositRD, ProfitDepositSB, ProfitLoansPigmy, ProfitLoansRD
        
        ParentID = parDepIntReceived
    
End Select

If ParentID = 0 Then Exit Function
If HeadName = "" Then Exit Function

' Check the given Heads in the database
NewIndexTrans.SQLStmt = " SELECT HeadID" & _
                   " FROM Heads " & _
                   " WHERE HeadName = " & AddQuotes(HeadName, True) & _
                   " AND ParentID = " & ParentID
                
' if exists then exit function
If NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then
    GetHeadIDCreatedOnEnum = FormatField(rstHeads.Fields(0))
    Exit Function
End If

' Get the Maximum Head From the database
NewIndexTrans.SQLStmt = " SELECT MAX(HeadID)" & _
                   " FROM Heads " & _
                   " WHERE ParentID = " & ParentID

Call NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly)

HeadID = FormatField(rstHeads.Fields(0)) + 1
    
If HeadID < ParentID Then HeadID = HeadID + ParentID

' Else create the Head with given headID & parentiD

' Insert the heads into the database
NewIndexTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID) " & _
                  " VALUES ( " & _
                  HeadID & "," & _
                  AddQuotes(HeadName, True) & "," & _
                  ParentID & ")"

If Not NewIndexTrans.SQLExecute Then Exit Function

USOpDate = FinUSFromDate

' Insert the opening balance
NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 "#" & USOpDate & "#," & _
                 OpBalance & ")"
                 
If Not NewIndexTrans.SQLExecute Then Exit Function


' It will insert into the IndexIDs Table
If Not InsertIntoIndexIds(IndexIds, HeadID) Then Exit Function

GetHeadIDCreatedOnEnum = HeadID

Set rstHeads = Nothing

Exit Function

ErrLine:
    
    MsgBox "CreateHead: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    GetHeadIDCreatedOnEnum = -1
    
    Set rstHeads = Nothing
    
    Exit Function
        
End Function

