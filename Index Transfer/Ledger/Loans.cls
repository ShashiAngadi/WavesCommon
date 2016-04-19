VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function LoanRepay() As Boolean

On Error GoTo Err_Line

' Variables used in this procedure...

Dim lLoanID As Long
Dim newTransID As Long
Dim inTransaction As Boolean
Dim NewBalance As Currency
Dim Balance As Currency
Dim IntAmount  As Currency
Dim PenalIntAmount As Currency
Dim PrincAmount  As Currency
Dim IntBalance As Currency
Dim PayAmount As Currency
Dim RegInt As Currency
Dim PenalInt As Currency
Dim IntPaidDate As String
Dim TransType As wisTransactionTypes
 

Dim RepayDate As String

RepayDate = txtRepayDate.Text

'Calculate the RegInt & Penal Int on the Specified date
RegInt = ComputeRegularInterest(RepayDate, lLoanID)
RegInt = RegInt \ 1
PenalInt = ComputePenalInterest(RepayDate, lLoanID)
PenalInt = PenalInt \ 1

IntBalance = 0

If IntAmount < 0 Then IntBalance = IntAmount

If (Val(txtRegInterest.Text) + Val(txtPenalInterest.Text)) - 1 > (RegInt + PenalInt) Then

    If MsgBox(LoadResString(gLangOffSet + 772) & _
        vbCrLf & LoadResString(gLangOffSet + 786), vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            IntBalance = 0
    Else
        IntBalance = (RegInt + PenalInt) - IntAmount
    End If
End If

If PrincAmount < 0 Then
    IntBalance = IntAmount - PayAmount
    'Upadate the Varaible as paying interest
    IntAmount = PayAmount
    PrincAmount = 0
End If

' Get a new transactionID.
gDbTrans.SqlStmt = " SELECT MAX(TransID) " & _
                   " FROM  " & m_TransTable & "  " & _
                   " WHERE loanid = " & lLoanID
        
If gDbTrans.SQLFetch <= 0 Then GoTo Exit_Line

newTransID = Val(FormatField(gDbTrans.Rst(0)))

gDbTrans.SqlStmt = " SELECT Balance,TRANSDATE " & _
                   " FROM  " & m_TransTable & _
                   " WHERE loanid = " & lLoanID & _
                   " AND TransID =  " & newTransID
        
If gDbTrans.SQLFetch <= 0 Then GoTo Exit_Line

Balance = Val(FormatField(gDbTrans.Rst("Balance")))

NewBalance = Balance - PrincAmount

If WisDateDiff(FormatField(gDbTrans.Rst("TransDate")), RepayDate) < 0 Then
   
    MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"

    ActivateTextBox txtRepayDate
    Exit Function
End If

'if the amount paying is equal to the loan balance and
'interest what he is paying is nill in such case warn him
If IntAmount = 0 And NewBalance = 0 And RegInt > 0 Then
    If MsgBox("amount specified is less than interest amount" & vbCrLf & _
        "Do you want to continue ?", vbYesNo + vbDefaultButton2, _
        wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If


'Begin the transaction
If Not gDbTrans.BeginTrans Then GoTo Exit_Line

inTransaction = True
    
If Balance > 0 Then
    ''Very First update the Penal Interest amount
    TransType = wCharges
    ' Get the balance.
    'He is paying Only interest need not to change the BALANCE
    NewBalance = CCur(txtBalance.Caption)

    ' Update the Penal interest amount.
    newTransID = newTransID + 1
    gDbTrans.SqlStmt = "INSERT INTO  " & m_TransTable & "  (LoanID, TransID, " _
            & "TransType, Amount, transDate, Balance, Particulars ) " _
            & "VALUES (" & lLoanID & ", " & newTransID & ", " _
            & TransType & ", " & PenalIntAmount & ", #" _
            & FormatDate(txtRepayDate.Text) & "#, " _
            & NewBalance & ",'" & LoadResString(gLangOffSet + 345) & "')"
            
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    m_TotalInterest = m_TotalInterest - PenalIntAmount

    ' Update the interest amount.
    If IntAmount < 0 And NewBalance - PrincAmount > 0 Then
    
        IntAmount = 0
        
    ElseIf NewBalance - PrincAmount <= 0 Then
    
    'Repay the interest if he  has paid extra interst earlier
        If IntAmount < 0 Then
            TransType = wInterest
            IntAmount = Abs(IntAmount)
           MsgBox "He has paid extra interest in the prevoius Payment " & vbCrLf & _
                "Return the extra interest Rs." & IntAmount
       End If
    End If
        
    newTransID = newTransID + 1
    gDbTrans.SqlStmt = "INSERT INTO  " & m_TransTable & "  (LoanID, TransID, " _
            & "TransType, Amount, transDate, Balance, Particulars ) " _
            & "VALUES (" & lLoanID & ", " & newTransID & ", " _
            & TransType & ", " & IntAmount & ", #" _
            & FormatDate(txtRepayDate.Text) & "#, " _
            & NewBalance & ", '" & LoadResString(gLangOffSet + 344) & "')"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    m_TotalInterest = m_TotalInterest - IntAmount
    
    'If he has not paid any interest and loan is BKCC
    'then keep the deatil of last interest paid date
    If IntAmount <> 0 And m_BKCC Then
        gDbTrans.SqlStmt = "UPDATE LoanMaster Set LastIntDate = " & _
            "#" & FormatDate(txtRepayDate.Text) & "# " & _
            " WHERE LoanId = " & lLoanID
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
    ''Next Pay the Principle amount
    TransType = wDeposit
    ' Increment the transaction ID.
    newTransID = newTransID + 1
    
    'Update the principal amount.
    NewBalance = NewBalance - PrincAmount
    gDbTrans.SqlStmt = "INSERT INTO  " & m_TransTable & " (LoanID, TransID, " _
        & "TransType, Amount, transDate, Balance, Particulars) " _
        & "VALUES (" & lLoanID & ", " & newTransID & ", " _
        & TransType & ", " & _
        IIf(NewBalance <= 0, PrincAmount + NewBalance, PrincAmount) & ", #" _
        & FormatDate(txtRepayDate.Text) & "#, " _
        & IIf(NewBalance <= 0, 0, NewBalance) & ", '" & LoadResString(gLangOffSet + 343) & "')"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

If NewBalance < 0 Then
    If Not m_BKCC Then
        MsgBox "Invalid amount specified!", vbCritical, "BKCC DEPOSIT"
        ActivateTextBox txtRepayAmt
        GoTo Exit_Line
    End If
    newTransID = newTransID + 1
    TransType = wBKCCDeposit
    gDbTrans.SqlStmt = "INSERT INTO  " & m_TransTable & "  (LoanID, TransID, " _
        & "TransType, Amount, transDate, Balance, Particulars) " _
        & "VALUES (" & lLoanID & ", " & newTransID & ", " _
        & TransType & ", " & IIf(Balance > 0, Abs(NewBalance), PrincAmount) & ", #" _
        & FormatDate(txtRepayDate.Text) & "#, " _
        & NewBalance & ", 'Access Amount Deposited')"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

    ' Update the loan master with interest balance.
    ' In this case, the total repaid amt is less than the interest payable.
    ' Therefore, put this difference amt, to loan master table.
    If m_TotalInterest >= 0 And NewBalance > 0 Then
        gDbTrans.SqlStmt = "UPDATE loanmaster SET [InterestBalance] = " _
                & AddQuotes(CStr(m_TotalInterest), True) & " WHERE loanid = " & lLoanID
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If

    ' If the balance amount is fully paidup, then set the flag "LoanClosed" to True.
    If NewBalance = 0 Then
        gDbTrans.SqlStmt = "UPDATE loanmaster SET loanclosed = TRUE, " & _
        "InterestBalance = 0 Where LoanID =  " & lLoanID
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
   
    ' Commit the transaction.
    If Not gDbTrans.CommitTrans Then GoTo Exit_Line
    inTransaction = False
    
'Now UpDate the  Miscalleneous Amount
Dim BankClass As New clsBankAcc
If Val(txtMIsc.Text) > 0 Then
    ' while undoing transaction we have to undo this profit also
    ' so for the identification I'm sendig Loan Id & TransID
    
    Call BankClass.UPDateMiscProfit(Val(txtMIsc.Text), txtRepayDate.Text, "LoanRepay " & lLoanID & "-" & newTransID)
End If
Set BankClass = Nothing

    LoanRepay = True
    'MsgBox "Loan repayment accepted.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 706), vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    If inTransaction Then gDbTrans.RollBack
    Exit Function

Err_Line:
    If Err Then
        MsgBox "AcceptPayment: " & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
        'MsgBox LoadResString(gLangOffSet + 707) & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
    End If
    GoTo Exit_Line
End Function


Private Function LoanValidations() As Boolean

' Check if a valid amount is entered.
If Not CurrencyValidate(txtRepayAmt.Text, False) Then
    'MsgBox "Enter valid amount.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayAmt
    GoTo Exit_Line
End If

'If Package is Considering The Interest Paid date Then
 If Not m_InterestBalance Then  'Check the Validate
     If Not DateValidate(txtIntBalance.Text, "/", True) Then
       MsgBox LoadResString(gLangOffSet + 501), vbExclamation, wis_MESSAGE_TITLE
       ActivateTextBox txtIntBalance
       Exit Function
     End If

' Check if the repayment date is later than today's date.
     If WisDateDiff(FormatDate(gStrDate), txtIntBalance.Text) > 0 Then
         'MsgBox "Repayment date cannot be greater than today's date", _
                 vbExclamation, wis_MESSAGE_TITLE
         MsgBox "Repayment date cannot be greater than today's date", _
                 vbExclamation, wis_MESSAGE_TITLE
         GoTo Exit_Line
     End If
End If

' Get the loanID.
lLoanID = Val(Mid(tabLoans.SelectedItem.Key, 4))
IntAmount = CCur(Val(txtRegInterest.Text))
PenalIntAmount = CCur(Val(txtPenalInterest.Text))
PayAmount = CCur(txtRepayAmt.Text)
PrincAmount = PayAmount - IntAmount - PenalIntAmount - CCur(Val(txtMIsc.Text))
                        
If PayAmount < CCur(Val(txtIntBalance.Text)) And m_InterestBalance Then
    If MsgBox("The amount he is paying is less than his Previous Interest Balance " & _
        vbCrLf & " Do you want to continue with transaction", vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If


End Function


