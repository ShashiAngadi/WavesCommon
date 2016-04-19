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
Public gOnLine As Boolean
Public Function Balance(ByVal AsOnIndiandate As String, ByVal HeadID As Long) As Currency

'Trap an Error
On Error GoTo ErrLine

Dim AccTransClass As clsAccTrans

Set AccTransClass = New clsAccTrans

Balance = AccTransClass.GetOpBalance(HeadID, AsOnIndiandate)

Set AccTransClass = Nothing

Exit Function

ErrLine:
    
    MsgBox "Balance: " & vbCrLf & Err.Description, vbCritical
        
End Function



Private Function InsertIntoIndexIds(ByVal HeadID As Long, _
                ByVal HeadName As String, AccType As wisModules) As Boolean

InsertIntoIndexIds = False

If HeadID = 0 Then Exit Function

NewIndexTrans.SQLStmt = " INSERT INTO BankHeadIds " & _
                "(HeadID,HeadName,AccType,AliasName) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 AddQuotes(HeadName, True) & ", " & _
                 AccType & "," & _
                 AddQuotes(HeadName, True) & ")"

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

Public Function UndoContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

UndoContraTrans = False

If gOnLine Then UndoContraTrans = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then UndoContraTrans = True: Exit Function

VoucherType = Journal

If UndoTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
UndoContraTrans = True

Exit Function

Hell:
    
    UndoContraTrans = False
    
End Function
Public Function UpdateContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean



UpdateContraTrans = False

If gOnLine Then UpdateContraTrans = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function

If Amount = 0 Then UpdateContraTrans = True: Exit Function

VoucherType = Journal

If UpdateTrasnction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
UpdateContraTrans = True

Exit Function

Hell:
    
    UpdateContraTrans = False
    
End Function

Public Function AddContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddContraTrans = False

If Not gOnLine Then AddContraTrans = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then AddContraTrans = True: Exit Function

'If Not DateValidate(TransDate, "/", False) Then Exit Function
Dim USTransDate As Date
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

If gOnLine Then UpdateCashDeposits = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If HeadID = 0 Then Exit Function
If Amount = 0 Then UpdateCashDeposits = True: Exit Function

CrHeadID = HeadID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

If UpdateTrasnction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    UpdateCashDeposits = True

Exit Function

Hell:
    
    UpdateCashDeposits = False
    
End Function

Public Function AddCashDeposits(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddCashDeposits = False

If Not gOnLine Then AddCashDeposits = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then AddCashDeposits = True: Exit Function

CrHeadID = HeadID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddCashDeposits = True

Exit Function

Hell:
    
    AddCashDeposits = False
    
End Function

Public Function AddContraDeposits(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddContraDeposits = False

If Not gOnLine Then AddContraDeposits = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then AddContraDeposits = True: Exit Function

VoucherType = Receipt

CrHeadID = HeadID
DrHeadID = 1

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddContraDeposits = True

Exit Function

Hell:
    
    AddContraDeposits = False
    
End Function


Public Function AddCashWithDrawls(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddCashWithDrawls = False

If Not gOnLine Then AddCashWithDrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then AddCashWithDrawls = True: Exit Function


CrHeadID = wis_CashHeadID
DrHeadID = HeadID

VoucherType = Payment

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddCashWithDrawls = True

Exit Function

Hell:
    
    AddCashWithDrawls = False
    
End Function

Public Function AddContraWithDrawls(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddContraWithDrawls = False

If Not gOnLine Then AddContraWithDrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then AddContraWithDrawls = True: Exit Function


CrHeadID = 1
DrHeadID = HeadID

VoucherType = Payment

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddContraWithDrawls = True

Exit Function

Hell:
    
    AddContraWithDrawls = False
    
End Function

Public Function UndoCashWithdrawls(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal USTransDate As Date) As Boolean

UndoCashWithdrawls = False

If gOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If HeadID = 0 Then Exit Function
If Amount = 0 Then UndoCashWithdrawls = True: Exit Function

DrHeadID = HeadID
CrHeadID = wis_CashHeadID

VoucherType = Payment

If UndoTransaction(CrHeadID, DrHeadID, Amount, USTransDate, VoucherType) = Success Then _
UndoCashWithdrawls = True

Exit Function

Hell:
    
    UndoCashWithdrawls = False
    
End Function


Public Function UndoCashDeposits(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal USTransDate As Date) As Boolean

UndoCashDeposits = False

If gOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If HeadID = 0 Then Exit Function
If Amount = 0 Then UndoCashDeposits = True: Exit Function


CrHeadID = HeadID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

If UndoTransaction(CrHeadID, DrHeadID, Amount, USTransDate, VoucherType) = Success Then _
UndoCashDeposits = True

Exit Function

Hell:
    
    UndoCashDeposits = False
    
End Function

Public Function UpdateCashWithDrawls(ByVal HeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

UpdateCashWithDrawls = False

If gOnLine Then UpdateCashWithDrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If HeadID = 0 Then Exit Function
If Amount = 0 Then UpdateCashWithDrawls = True: Exit Function

DrHeadID = HeadID
CrHeadID = wis_CashHeadID

VoucherType = Payment

If UpdateTrasnction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    UpdateCashWithDrawls = True

Exit Function

Hell:
    
    UpdateCashWithDrawls = False
    
End Function

Public Function UndoBalances(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal USTransDate As Date, _
                                ByVal VoucherType As Wis_VoucherTypes) As Boolean
UndoBalances = False

If gOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then UndoBalances = True: Exit Function

Set AccTransClass = New clsAccTrans

If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UndoVouchers(VoucherType, USTransDate) <> Success Then Exit Function

UndoBalances = True

Exit Function

Hell:
    UndoBalances = False
    
End Function
'This Function Update the transction to the Acctrans Table
'Here All the HeadId's tobe transferred
Private Function AddTransaction(ByVal CrHeadID As Long, ByVal DrHeadID As Long, _
        ByVal Amount As Currency, ByVal TransDate As Date, _
        ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

On Error GoTo Hell

If Not gOnLine Then AddTransaction = Success: Exit Function

Dim AccTransClass As clsAccTrans
Dim rstTrans As Recordset
Dim TransID As Long
Dim HeadID As Long
Dim CrHeadExists As Boolean
Dim DrHeadExists As Boolean
Dim Debit As Currency
Dim Credit As Currency

HeadID = IIf(VoucherType = Payment, CrHeadID, DrHeadID)

Set AccTransClass = New clsAccTrans
NewIndexTrans.SQLStmt = "Select * From AccTrans Where TransID = " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID = " & HeadID & _
            " AND TransDate = #" & TransDate & "#" & _
            " ANd A.VoucherType = " & VoucherType & ")"

If NewIndexTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then Set rstTrans = Nothing
Set AccTransClass = New clsAccTrans

If Not rstTrans Is Nothing Then
    ' Start the Loop
    TransID = rstTrans.Fields("TransID")
    AccTransClass.TransID = TransID
    Do While Not rstTrans.EOF
        If TransID <> rstTrans.Fields("TransID") Then Exit Do
        ' Feed data to Controls
        HeadID = FormatField(rstTrans.Fields("HeadID"))
        If HeadID = CrHeadID Then CrHeadExists = True
        If HeadID = DrHeadID Then DrHeadExists = True
        Debit = FormatField(rstTrans.Fields("Debit"))
        Credit = FormatField(rstTrans.Fields("Credit"))
        
        'Update the Collections
        If AccTransClass.AllTransHeadsAdd(HeadID, Debit, Credit) <> Success Then Exit Function
        rstTrans.MoveNext
    Loop
End If

If DrHeadExists Then AccTransClass.AllTransHeadIDRemove (DrHeadID)
If CrHeadExists Then AccTransClass.AllTransHeadIDRemove (CrHeadID)

If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

AddTransaction = Success
Hell:

Exit Function

End Function

'This Function Update the transction to the Acctrans Table
'Here All the HeadId's tobe transferred
Private Function UpdateTrasnction(ByVal CrHeadID As Long, ByVal DrHeadID As Long, _
        ByVal Amount As Currency, ByVal TransDate As Date, _
        ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

On Error GoTo Hell

If gOnLine Then UpdateTrasnction = Success: Exit Function

Dim AccTransClass As clsAccTrans
Dim rstTrans As Recordset
Dim TransID As Long
Dim HeadID As Long
Dim CrHeadExists As Boolean
Dim DrHeadExists As Boolean
Dim CRDebit As Currency
Dim DRDebit As Currency

Dim CRCredit As Currency
Dim DRCredit As Currency

Dim SqlSupport As String

HeadID = IIf(VoucherType = Payment, CrHeadID, DrHeadID)

If HeadID = wis_CashHeadID Then
    SqlSupport = " = " & wis_CashHeadID & " "
Else
    SqlSupport = " IN (Select Distinct HeadId From BankHeadIds) "
End If


Set AccTransClass = New clsAccTrans
NewIndexTrans.SQLStmt = "Select * From AccTrans Where TransID IN " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID " & SqlSupport & _
            " AND TransDate = #" & TransDate & "#" & _
            " ANd A.VoucherType = " & VoucherType & ")"

If NewIndexTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then Set rstTrans = Nothing
Set AccTransClass = New clsAccTrans

If Not rstTrans Is Nothing Then
    Dim Debit As Currency
    Dim Credit As Currency
    
    ' Start the Loop
    TransID = rstTrans.Fields("TransID")
    AccTransClass.TransID = TransID
    Do While Not rstTrans.EOF
        If TransID <> rstTrans.Fields("TransID") Then Exit Do
        ' Feed data to Controls
        HeadID = FormatField(rstTrans.Fields("HeadID"))
        Debit = FormatField(rstTrans.Fields("Debit"))
        Credit = FormatField(rstTrans.Fields("Credit"))
    
        If HeadID = CrHeadID Then
            CrHeadExists = True
            CRDebit = Debit
            CRCredit = Credit
        ElseIf HeadID = DrHeadID Then
            DrHeadExists = True
            DRDebit = Debit
            DRCredit = Credit
        End If
        'Update the Collections
        If AccTransClass.AllTransHeadsAdd(HeadID, Debit, Credit) _
                    <> Success Then Exit Function
        rstTrans.MoveNext
    Loop
End If

If DrHeadExists Then AccTransClass.AllTransHeadIDRemove (DrHeadID)
If CrHeadExists Then AccTransClass.AllTransHeadIDRemove (CrHeadID)

If AccTransClass.AllTransHeadsAdd(DrHeadID, DRDebit + Amount, 0) _
        <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, CRCredit + Amount) _
        <> Success Then Exit Function
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

Dim Rst As Recordset

UpdateTrasnction = Success

Hell:

Exit Function

End Function


'This Function Update the transction to the Acctrans Table
'Here All the HeadId's tobe transferred
Private Function UndoTransaction(ByVal CrHeadID As Long, ByVal DrHeadID As Long, _
        ByVal Amount As Currency, ByVal TransDate As Date, _
        ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

On Error GoTo Hell

Dim AccTransClass As clsAccTrans
Dim rstTrans As Recordset
Dim TransID As Long
Dim HeadID As Long
Dim ObjectCount As Integer
Dim CrHeadExists As Boolean
Dim DrHeadExists As Boolean

Dim CRDebit As Currency
Dim DRDebit As Currency
Dim CRCredit As Currency
Dim DRCredit As Currency
Dim Debit As Currency
Dim Credit As Currency

Dim SqlSupport As String

Set AccTransClass = New clsAccTrans

HeadID = IIf(VoucherType = Payment Or VoucherType = Journal, CrHeadID, DrHeadID)
'If (CrHeadID = wis_CashHeadID) Or (DrHeadID = wis_CashHeadID) Then HeadID = wis_CashHeadID

If HeadID = wis_CashHeadID Then
    SqlSupport = " = " & wis_CashHeadID & " "
Else
    SqlSupport = " IN (Select Distinct HeadId From BankHeadIds) "
End If

NewIndexTrans.SQLStmt = "Select * From AccTrans Where TransID IN " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID " & SqlSupport & _
            " AND TransDate = #" & TransDate & "#" & _
            " ANd A.VoucherType = " & VoucherType & ")"

If NewIndexTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then Set rstTrans = Nothing
Set AccTransClass = New clsAccTrans

If Not rstTrans Is Nothing Then
    ' Start the Loop
    TransID = rstTrans.Fields("TransID")
    AccTransClass.TransID = TransID
    Do While Not rstTrans.EOF
        If TransID <> rstTrans.Fields("TransID") Then Exit Do
        ' Feed data to Controls
        HeadID = FormatField(rstTrans.Fields("HeadID"))
        Debit = FormatField(rstTrans.Fields("Debit"))
        Credit = FormatField(rstTrans.Fields("Credit"))
            
        If HeadID = CrHeadID Then
            CrHeadExists = True
            CRDebit = Debit
            CRCredit = Credit
        ElseIf HeadID = DrHeadID Then
            DrHeadExists = True
            DRDebit = Debit
            DRCredit = Credit
        End If
            'Update the Collections
        If AccTransClass.AllTransHeadsAdd(HeadID, Debit, Credit) _
                <> Success Then Exit Function
        ObjectCount = ObjectCount + 1
        rstTrans.MoveNext
    Loop
End If

''Information for the Developer
Debug.Assert DrHeadExists And CrHeadExists

If DrHeadExists Then AccTransClass.AllTransHeadIDRemove (DrHeadID)
If CrHeadExists Then AccTransClass.AllTransHeadIDRemove (CrHeadID)

Debit = DRDebit - Amount: Credit = DRCredit
If Debit Or Credit Then
    If AccTransClass.AllTransHeadsAdd(DrHeadID, Debit, Credit) <> Success Then Exit Function
Else
    ObjectCount = ObjectCount - 1
End If

Debit = CRDebit: Credit = CRCredit - Amount
If Debit Or Credit Then
    If AccTransClass.AllTransHeadsAdd(CrHeadID, Debit, Credit) <> Success Then Exit Function
Else
    ObjectCount = ObjectCount - 1
End If

If ObjectCount Then
    If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function
Else
    NewIndexTrans.SQLStmt = "DELETE * FROM AccTrans Where TransID = " & TransID
    If Not NewIndexTrans.SQLExecute Then GoTo Hell
End If
UndoTransaction = Success

Hell:

Exit Function

End Function

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
Public Function Profit(ByVal FromDate As Date, ByVal ToDate As Date, ByVal HeadID As Long) As Currency

' To trap an error
On Error GoTo ErrLine

'Declare variables
Dim AccountType As wis_AccountType
Dim rstProfit As ADODB.Recordset

Profit = 0

' Validations
If HeadID = 0 Then Exit Function

' set the accounttype
AccountType = Profit

' This is the sql to fetch the profits

NewIndexTrans.SQLStmt = " SELECT b.Credit as ProfitReceived, b.Debit as ProfitGiven, e.AccountType " & _
                   " FROM qryAccTransmerge AS d, Heads AS c, ParentHeads AS e " & _
                   " WHERE a.HeadId = " & HeadID & _
                   " AND b.HeadID<>" & HeadID & _
                   " AND a.TransDate BETWEEN #" & FromDate & "#" & _
                   " AND #" & ToDate & "#" & _
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


Public Function Deposits(FromDate As Date, ToDate As Date, HeadID As Long) As Currency

Dim rstDeposits As ADODB.Recordset

Deposits = 0

If HeadID = 0 Then Exit Function

NewIndexTrans.SQLStmt = " SELECT Sum(Debit) as Deposits " & _
                   " FROM AccTrans " & _
                   " WHERE HeadID=" & HeadID & _
                   " AND TransDate BETWEEN " & "#" & FromDate & "#" & _
                   " AND #" & ToDate & "#"
                 
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
Public Function GetHeadIDCreated(ByVal HeadName As String, _
            Optional ParentID As Long, Optional OpBalance As Currency, _
            Optional AccType As wisModules) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset
Dim USOpDate As String
Dim HeadID As Long

'initialise the function
GetHeadIDCreated = 0

'Validate the inputs
If HeadName = "" Then Exit Function

'Check the given Heads in the database
NewIndexTrans.SQLStmt = " SELECT HeadID FROM BankHeadIds " & _
                   " WHERE HeadName = " & AddQuotes(HeadName, True)
                
'if exists then exit function
If NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then
    GetHeadIDCreated = FormatField(rstHeads.Fields(0))
    Exit Function
End If

If ParentID = 0 Then Exit Function
'Get the Maximum Head From the database
NewIndexTrans.SQLStmt = "SELECT MAX(HeadID)" & _
                   " FROM Heads " & _
                   " WHERE ParentID = " & ParentID

Call NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly)
HeadID = FormatField(rstHeads.Fields(0)) + 1

Set rstHeads = Nothing
    
If HeadID < ParentID Then HeadID = HeadID + ParentID

' Else create the head with given headID & parentiD
'Insert the heads into the database
NewIndexTrans.SQLStmt = "INSERT INTO Heads (HeadID,HeadName,ParentID) " & _
                  " VALUES ( " & _
                  HeadID & "," & _
                  AddQuotes(HeadName, True) & "," & _
                  ParentID & ")"

If Not NewIndexTrans.SQLExecute Then Exit Function


'USOpDate = FormatDate(FinIndianFromDate)
USOpDate = frmMain.txtDate
USOpDate = "1 Apr " & Right(USOpDate, 2)



USOpDate = DateAdd("D", 1, gFromDate) ', "4/1/2003"
'Insert the opening balance
NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 "#" & USOpDate & "#," & _
                 OpBalance & ")"
If Not NewIndexTrans.SQLExecute Then Exit Function

If Year(gFromDate) = 2002 Then
    USOpDate = DateAdd("YYYY", 1, gFromDate) ', "4/1/2004"
    NewIndexTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 "#" & USOpDate & "#," & _
                 OpBalance & ")"
If Not NewIndexTrans.SQLExecute Then Exit Function
End If
If Not InsertIntoIndexIds(HeadID, HeadName, AccType) Then Exit Function

GetHeadIDCreated = HeadID

Exit Function

ErrLine:
    MsgBox "CreateHead: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    GetHeadIDCreated = -1
    
    Set rstHeads = Nothing
    
    Exit Function
End Function

