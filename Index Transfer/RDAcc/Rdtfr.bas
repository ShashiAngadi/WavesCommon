Attribute VB_Name = "RDTransfer"
'This BAs file is used to Transfer
'RDMaster & Sb TranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit

Dim m_AccOfSet As Long
Dim m_LoanOffSet As Long
Private m_DepositType As wis_DepositType



Private Function CreateRDHeads(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean

Dim HeadID As Long

Dim ClsBank As clsBankAcc

On Error GoTo ErrLine
'First Creat the Dl Deposit Heads

Set ClsBank = New clsBankAcc
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans

'First Create the fixed Deposit Heads
Dim HeadName As String
Dim Prefix As String
Dim HeadBalance As Currency
Dim LoanBalance As Currency
Dim rstTemp As Recordset

'Get the Head Balance
OldTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#"
If OldTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    Do
        If rstTemp.EOF Then Exit Do
        If rstTemp("Module") = 55 Then _
            HeadBalance = FormatField(rstTemp("ObAmount"))
        If rstTemp("Module") = 56 Then _
            LoanBalance = FormatField(rstTemp("ObAmount"))
        If HeadBalance And LoanBalance Then Exit Do
        rstTemp.MoveNext
    Loop
End If

'Prefix = IIf(gLangOffSet, "„ÀÐ´þ³Ð ¬ÙÓÀÐ±", "Recurring Deposit")
Prefix = LoadResString(gLangOffSet + 424)

Dim PrgVal As Integer
frmMain.lblProgress = "Transferring the Rd Ledger Heads"
frmMain.Refresh
With frmMain.prg
    .Max = 365
    .Min = 0
    .Value = 0
End With

Dim FromDate As Date
Dim rstTrans As ADODB.Recordset
Dim TransDate As Date
Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency
Dim TransType As wisTransactionTypes

FromDate = gFromDate '"3/31/03"

'First Insert the Deposit Transaction
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM RDTrans " & _
                "Where TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
    NewIndexTrans.BeginTrans
    
    'Recurring Deposit Heads
    HeadName = Prefix
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemberDeposit, 0, wis_RDAcc)

    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        
        PrgVal = PrgVal + 1
        With frmMain.prg
            If PrgVal >= .Max Then .Max = PrgVal * 1.5
            .Value = PrgVal
        End With

        rstTrans.MoveNext
    Wend

    Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert the interest transaction of rd deposit
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM RDIntTrans " & _
                "Where TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit Interest Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 487)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_RDAcc)

    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
            'Call ClsBank.UpdateCashWithDrawls(HeadID, TotalDeposit, TransDate)
            'Call ClsBank.UpdateCashDeposits(HeadID, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        
        
        PrgVal = PrgVal + 1
        With frmMain.prg
            If PrgVal >= .Max Then .Max = PrgVal * 1.5
            .Value = PrgVal
        End With

        rstTrans.MoveNext
    Wend
    Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
    'Call ClsBank.UpdateCashWithDrawls(HeadID, TotalDeposit, TransDate)
    'Call ClsBank.UpdateCashDeposits(HeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert Interest payable
NewTrans.SQLStmt = "Select Sum(Amount) As TotalAmount ," & _
                "TransType,TransDate FROM RDIntPayable " & _
                "Where TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType "
                
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    OldTrans.SQLStmt = "Select Top 1 bALANCE From AccTrans " & _
                " WHERE TransDate <=#" & gFromDate & "#" & _
                " AND AccId = 13003 ORder By TransID Desc"
    If OldTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                    HeadBalance = FormatField(rstTemp("Balance"))

    NewIndexTrans.BeginTrans
    'Deposit Interest Provision Head
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 375) & " " & LoadResString(gLangOffSet + 47)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parDepositIntProv, HeadBalance, wis_RDAcc)

    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        rstTrans.MoveNext
    Wend
    Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
End If

'NOW INSERT THE TRANSACTION RDEPOSIT LOAN
'Insert the Deposit loan Transaction
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM DepositLoanTrans " & _
                "Where LoanId in (Select LoanID From DepositLoanMaster " & _
                    "Where DepositType = " & m_DepositType & ")" & _
                " And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit LOan Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 58)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemDepLoan, LoanBalance, wis_DepositLoans + wis_RDAcc)

    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        
        PrgVal = PrgVal + 1
        With frmMain.prg
            If PrgVal >= .Max Then .Max = PrgVal * 1.5
            .Value = PrgVal
        End With

        rstTrans.MoveNext
    Wend

    Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert the interest transaction of deposit Loans
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM DepositLoanIntTrans " & _
                "Where loanId in (Select LOanID From DepositLoanMaster " & _
                    "Where DepositType = " & m_DepositType & ") " & _
                "And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit Loan Interest Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 483)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemDepLoanIntReceived, 0, wis_DepositLoans + wis_RDAcc)

    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        
        
        PrgVal = PrgVal + 1
        With frmMain.prg
            If PrgVal >= .Max Then .Max = PrgVal * 1.5
            .Value = PrgVal
        End With

        rstTrans.MoveNext
    Wend
    Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

CreateRDHeads = True

ExitLine:

    Set NewIndexTrans = Nothing
    Set rstTrans = Nothing
    Set ClsBank = Nothing
    
    
ErrLine:
    
If Err.Number = 380 Then
    frmMain.prg.Max = PrgVal * 1.5
    Resume Next
ElseIf Err.Number Then
    MsgBox "Error In Chcking data Base"
    GoTo ExitLine
    'Resume
End If
    
End Function

'this function is used to transfer the
'RD MAster details form OLdb to new one
'and NewRDTrans has assigned to new database
Private Function TransferRDLoanMaster(OldRDTrans As clsOldUtils, NewRDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim Sql_2 As String
Dim Rst As Recordset

Dim rstTemp As Recordset
On Error GoTo Err_Line

'Now TransFer the Loan Details
'Get the account Having Details

With frmMain
    .lblProgress = "Transferring recurring deposit Loan details"
    .prg.Value = 0
    .Refresh
End With

SqlStr = "SELECT A.AccID,CustomerID,MaturityDate,Loan, " & _
    " TransDate,RateOfInterest,Amount,TransID,LedgerNo,FolioNo " & _
    " FROM RDTrans A, RDMaster B WHERE A.AccID = B.AccID " & _
    " AND TransId = (SELECT Min(TransID) " & _
        " From RDTrans C Where " & _
        " C.AccID = B.accID AND C.Loan = A.Loan ) "


OldRDTrans.SQLStmt = SqlStr
If OldRDTrans.Fetch(Rst, adOpenDynamic) <= 0 Then GoTo ExitLine
'Set Rst = OldRDTrans.Rst.Clone

NewRDTrans.SQLStmt = "SELECT MAX(LoanID) FROM DepositLoanMaster "
If NewRDTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then m_LoanOffSet = FormatField(rstTemp(0))

Dim AccID As Long
Dim LoanId As Long
Dim LoanNum As String


'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM RDMASTER "
OldRDTrans.SQLStmt = SqlStr
Call OldRDTrans.Fetch(rstTemp, adOpenDynamic)

'm_AccOffSet = FormatField(Rsttemp(0))
'm_AccOffSet = AccOffSet + 100 - (AccOffSet Mod 100)

With frmMain
    .lblProgress = "Transferring recuring Loan accounts"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With

LoanId = m_LoanOffSet: AccID = 0
While Not Rst.EOF
    'If AccId = FormatField(Rst("AccID")) Then GoTo NextAccount
    LoanNum = Rst("AccId")
    
    If Not Rst("Loan") Then GoTo NextAccount
    LoanId = LoanId + 1
    
    NewRDTrans.BeginTrans
    
'Now insert the Loans master details
    
'''' Here  we are inserting into the new table called DepositLoanMAster
''the above said table will be common for all type deposit(eg. FD,Rd,Pd,DL)

    m_DepositType = wisDeposit_RD
    SqlStr = "Insert INTO PledgeDeposit (" & _
        "LoanID,AccID,DepositType,PledgeNum)" & _
        " VALUES (" & _
        LoanId & "," & _
        AccID & "," & _
        m_DepositType & "," & _
        " 1 )"
            
    NewRDTrans.SQLStmt = SqlStr
    If Not NewRDTrans.SQLExecute Then
        NewRDTrans.RollBack
        MsgBox "Unable to transafer the RD Loan Master data"
        Exit Function
    End If
    
    SqlStr = "Insert INTO DepositLoanMaster (" & _
        "CustomerID,LoanID,AccNum,DepositType,LoanIssuedate," & _
        "LoanDueDate,PledgeDescription, " & _
        "InterestRate,LoanAmount,LedgerNo,FolioNo ," & _
        " LastPrintId )"
    SqlStr = SqlStr & " VALUES (" & _
        Rst("CustomerID") & "," & LoanId & "," & _
        AddQuotes(LoanNum, True) & "," & _
        m_DepositType & "," & _
        FormatDateField(Rst("TransDate")) & ", " & FormatDateField(Rst("MaturityDate")) & "," & _
        AccID & " ," & _
        Rst("RateOfInterest") & "," & _
        Rst("Amount") & "," & _
        "'" & Val(Rst("LedgerNo")) & "'," & _
        "'" & Val(Rst("FolioNo")) & "'," & _
        "1 )"
        
    NewRDTrans.SQLStmt = SqlStr
    If Not NewRDTrans.SQLExecute Then
        NewRDTrans.RollBack
        MsgBox "Unable to transafer the pigmy loan data"
        Exit Function
    End If
    
    SqlStr = "UPdate RDMaster Set LOanid = " & LoanId & _
            " Where AccNum = '" & Rst("AccId") & "'"
    NewRDTrans.SQLStmt = SqlStr
    If Not NewRDTrans.SQLExecute Then
        NewRDTrans.RollBack
        MsgBox "Unable to transafer the pigmy loan data"
        Exit Function
    End If
    
    NewRDTrans.CommitTrans
    
NextAccount:
    
    With frmMain
        .lblProgress = "Transferring recurring deposit Loan accounts"
        .prg.Value = Rst.AbsolutePosition
    End With
    Rst.MoveNext

Wend

ExitLine:
TransferRDLoanMaster = True

    Debug.Print Now
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If

    If Err Then MsgBox "eror In RD LoanMaster " & Err.Description
    
End Function


'this function is used to transfer the
'RD Loan transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferRDLoanTrans(OldRDTrans As clsOldUtils, NewRDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim IsIntTrans As Boolean

On Error GoTo Err_Line

Dim OldTransType As Integer, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim NewTransId As Long

Dim Rst As Recordset
Dim rstTemp As Recordset

Dim AccID As Long
Dim oldAccid As Long
Dim LoanId As Long
Dim Amount As Currency
Dim IntBalance As Currency
Dim TransDate As Date
Dim DepositType As wis_DepositType

DepositType = wisDeposit_RD
    
'Fetch the detials of pigmy Account

SqlStr = "SELECT * FROM RDTrans Where Loan = True " & _
    " ORDER BY AccID,TransId"

OldRDTrans.SQLStmt = SqlStr
If OldRDTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo ExitLine
'Set Rst = OldRDTrans.Rst.Clone


'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM RDMASTER "
OldRDTrans.SQLStmt = SqlStr
Call OldRDTrans.Fetch(rstTemp, adOpenDynamic)

'M_AccOffSet = FormatField(Rsttemp(0))
'M_AccOffSet = AccOffSet + 100 - (AccOffSet Mod 100)

With frmMain
    .lblProgress = "Transferring recurring deposit Loan transaction"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With

TransID = 10000

LoanId = m_LoanOffSet
Balance = 0
While Not Rst.EOF
    IsIntTrans = False
    OldTransType = FormatField(Rst("TransType"))
    If OldTransType = 4 Or OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Or OldTransType = -4 Then IsIntTrans = True
    
    If TransID >= FormatField(Rst("TransID")) Then
        'LoanId = LoanId + 1
        NewRDTrans.SQLStmt = "Select LoanID From DepositLoanMaster" & _
            " Where AccNum = '" & Rst("AccID") & "' And DepositType = " & wisDeposit_RD
        If NewRDTrans.Fetch(rstTemp, adOpenDynamic) < 0 Then GoTo NextAccount
        If LoanId = rstTemp("LoanID") Then
            LoanId = LoanId
            If rstTemp.RecordCount > 1 Then rstTemp.MoveNext: LoanId = rstTemp("LoanID"): NewTransId = 0
        Else
            LoanId = rstTemp("LoanID")
            NewTransId = 0
        End If
        IntBalance = 0
        Balance = -1
    End If
    NewTransId = NewTransId + 1
    TransID = FormatField(Rst("TransID"))
    oldAccid = Rst("accId")
    Amount = Rst("Amount")
                
    'NewTransType = OldTransType / Abs(OldTransType)
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    If IsIntTrans Then
        NewTransType = IIf(OldTransType > 0, wWithDraw, wDeposit)
        If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithDraw  'interest Paid to the customer
        If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit  'Interest collected form the customer
        TransDate = Rst("TransDate")
        IntBalance = IntBalance + Rst("Amount")
            
        'Insert INto table Called Depsoit LoanIntTrans
        SqlInt = "Insert INTO DepositLoanIntTrans ( " & _
            "LoanId,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            LoanId & "," & _
            NewTransId & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            IntBalance & " ," & _
            NewTransType & " )"
        'If Balance <> rst("Balance") Then GoTo NextAccount
        'If Balance <> rst("Balance") Then rst.MovePrevious
        Rst.MoveNext
        Amount = Rst("Amount")
    End If
    
    OldTransType = FormatField(Rst("TransType"))
    'NewTransType = (OldTransType / Abs(OldTransType))
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    
    NewRDTrans.BeginTrans
    
    ''INSERt INTO DEPOSITLOANTRANS
    SqlStr = "Insert INTO DepositLoanTrans ( " & _
        "LoanID,TransID,TransDate," & _
        "Amount,Balance,Particulars," & _
        "TransType )"
    SqlStr = SqlStr & "VALUES (" & _
        LoanId & "," & _
        NewTransId & "," & _
        "#" & Rst("TransDate") & "#," & _
        Rst("Amount") & "," & _
        Rst("Balance") & " ," & _
        AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
        NewTransType & " )"
    
    If Balance = Rst("Balance") Then SqlStr = ""
    If oldAccid <> Rst("accId") Then SqlStr = ""
    If Rst("Amount") = 0 Then SqlStr = ""
    
    If SqlStr <> "" Then
        NewRDTrans.SQLStmt = SqlStr
        If Not NewRDTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy loan Trans "
            NewRDTrans.RollBack
            Exit Function
        End If
    End If
    If SqlInt <> "" Then
        NewRDTrans.SQLStmt = SqlInt
        If Not NewRDTrans.SQLExecute Then
            MsgBox "Unable to transafer the RD loan Trans "
            NewRDTrans.RollBack
            Exit Function
        End If
        SqlInt = ""
    End If

    '''END DEPOSITLOANTRANS
    NewRDTrans.CommitTrans
    
    If Rst.AbsolutePosition Mod 5000 = 0 Then _
            Debug.Print Now & "  " & Rst.RecordCount
    Balance = FormatField(Rst("Balance"))

NextAccount:
    With frmMain
        .lblProgress = "Transferring recurring deposit loans"
        .prg.Value = Rst.AbsolutePosition
    End With
    Rst.MoveNext
Wend

ExitLine:
TransferRDLoanTrans = True

    With frmMain
        .lblProgress = "Transferred the RD details"
        .prg.Value = 0
        .Refresh
    End With

Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err Then MsgBox "Error in RDLoanTrans" & Err.Description
    
End Function


'just calling this function we can transafer the sbmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferRD(OldDBName As String, NewDBName As String) As Boolean
Debug.Print Now
Dim OldTrans As New clsOldUtils
Dim NewTrans As New clsDBUtils
If Not OldTrans.OpenDB(OldDBName, OldPwd) Then Exit Function
If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    Exit Function
End If

If Not gOnlyLedgerHeads Then
    If Not TransferRDMaster(OldTrans, NewTrans) Then Exit Function
    If Not TransferRDTrans(OldTrans, NewTrans) Then Exit Function
    If Not TransferRDLoanMaster(OldTrans, NewTrans) Then Exit Function
    If Not TransferRDLoanTrans(OldTrans, NewTrans) Then Exit Function
End If
    
If Not CreateRDHeads(OldTrans, NewTrans) Then Exit Function
Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
    
TransferRD = True
If Not PutVoucherNumber(NewTrans) Then
    MsgBox "Unable to set the voucher No"
    Exit Function
End If
    
End Function


'this function is used to transfer the
'RD MAster details form OLdb to new one
'and NewRDTrans has assigned to new database
Private Function TransferRDMaster(OldRDTrans As clsOldUtils, NewRDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim rstTemp As Recordset


On Error GoTo Err_Line

'Before Fetching Update the Values
'where It can be Null with default value
'Then Fetch the records

'Update Modify date
With frmMain
    .lblProgress = "Transferring recurring accounts"
    .prg.Min = 0
    .prg.Max = 100
    .prg.Value = 0
    .Refresh
End With

'Fetch the detials of RD Account
SqlStr = "SELECT A.* FROM RDMASTER A, RDTrans B " & _
    " WHERE A.AccID = B.AccID AND TransID = (SELECT MIn(TransID) " & _
        " FROm RDTrans C WHERE C.AccID = B.AccID AND Loan = FALSE )" & _
    " ORDER BY A.AccID"

Dim AccID As Long, IntroId As Long
Dim Rst As Recordset
OldRDTrans.SQLStmt = SqlStr

If OldRDTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo Exit_line
'Set Rst = OldRDTrans.Rst.Clone

With frmMain
    .lblProgress = "Transferring recurring accounts"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With


While Not Rst.EOF
    AccID = AccID + 1
    IntroId = FormatField(Rst("Introduced"))
    'Get the Introducer ID
    If IntroId > 0 Then
        OldRDTrans.SQLStmt = "SELECT CustomerID FROM RDMASTeR " & _
            " WHERE AccID = " & IntroId
        If OldRDTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then IntroId = FormatField(rstTemp("CustomerID"))
    End If
    
'    AccId = Rst("AccID")
    'First insert into Rd joint table
    SqlStr = "Insert INTO RDJOINT (" & _
        "AccID,CustomerID,CustomerNum)" & _
        "VALUES (" & _
        Rst("AccID") & "," & Rst("CustomerID") & "," & _
        "1 )"
        
    NewRDTrans.BeginTrans
    NewRDTrans.SQLStmt = SqlStr
'    If Not NewRDTrans.SQLExecute Then
'        NewRDTrans.RollBack
'        MsgBox "Unable to transafer the CA MAster data"
'        NewRDTrans.RollBack
'        Exit Function
'    End If
    
'NOW insert into
    SqlStr = "Insert INTO RDMASTER (" & _
        "AccID,CustomerID,AccNum,CreateDate,ModifiedDate," & _
        "ClosedDate,MaturityDate, " & _
        "InstallMentAmount,NoOfInstallMents,RateOfINterest," & _
        "IntroducerID,LedgerNo,FolioNo ," & _
        "NomineeID ,InOperative,LastPrintId,AccGroupID )"
    
    SqlStr = SqlStr & " VALUES (" & _
        Rst("AccID") & "," & Rst("CustomerID") & "," & _
        AddQuotes(Rst("AccID"), True) & "," & _
        FormatDateField(Rst("CreateDate")) & "," & _
        FormatDateField(Rst("Modifieddate")) & "," & _
        FormatDateField(Rst("ClosedDate")) & " ," & _
        FormatDateField(Rst("MaturityDate")) & " ," & _
        Rst("InstallmentAmount") & "," & _
        Rst("NoOfinstallments") & "," & _
        Rst("RateOfINterest") & "," & _
        Rst("Introduced") & "," & _
        Val(Rst("LedgerNo")) & "," & _
        Val(Rst("FolioNo")) & " ," & _
        "0 ," & _
        False & "," & _
        "1, 1 )"
        
    NewRDTrans.SQLStmt = SqlStr
    If Not NewRDTrans.SQLExecute Then
        NewRDTrans.RollBack
        MsgBox "Unable to transafer the RD MAster data"
        Exit Function
    End If
    NewRDTrans.CommitTrans
    
NextAccount:
    
    With frmMain
        .lblProgress = "Transferring recurring retails"
        .prg.Value = Rst.AbsolutePosition
    End With
    Rst.MoveNext
    
Wend

'Now reverse the change made before transfer
'Update Modify date
SqlStr = "UPDATE RDMASTER Set ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
OldRDTrans.SQLStmt = SqlStr
OldRDTrans.BeginTrans
If Not OldRDTrans.SQLExecute Then
    OldRDTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldRDTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE RDMASTER Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
OldRDTrans.SQLStmt = SqlStr
OldRDTrans.BeginTrans
If Not OldRDTrans.SQLExecute Then
    OldRDTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldRDTrans.CommitTrans



'Now Update the smae with new database
'Update Modify date
SqlStr = "UPDATE RDMAster Set ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
NewRDTrans.SQLStmt = SqlStr
NewRDTrans.BeginTrans
If Not NewRDTrans.SQLExecute Then
    NewRDTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewRDTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE RDMaster Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
NewRDTrans.SQLStmt = SqlStr
NewRDTrans.BeginTrans
If Not NewRDTrans.SQLExecute Then
    NewRDTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewRDTrans.CommitTrans
    
Exit_line:
TransferRDMaster = True
    Debug.Print Now
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "Error In RDMaster " & vbCrLf & Err.Description
        Err.Clear
    End If
    
End Function

'this function is used to transfer the
'SB transaction details form OLd Db to new one
'and NewRDTrans has assigned to new database
Private Function TransferRDTrans(OldRDTrans As clsOldUtils, NewRDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim SqlPayable As String

Dim IsIntTrans As Boolean
Dim IsPaybleTrans As Boolean

On Error GoTo Err_Line

Dim OldTransType As Integer, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim Rst As Recordset
Dim AccID As Long
Dim oldAccid As Long
Dim OldTransID As Long
Dim Amount As Currency
Dim PayableBalance As Currency
Dim IntBalance As Currency

Dim TransDate As Date
    'Fetch the detials of Sb Account

SqlStr = "SELECT * FROM RDTrans WHERE Loan = False " & _
    " ORDER BY AccID,TransId"
OldRDTrans.SQLStmt = SqlStr
If OldRDTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo Exit_line
'Set Rst = OldRDTrans.Rst.Clone

OldTransID = 100000
With frmMain
    .lblProgress = "Transferring recuring transaction"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With

'Balance = FormatField(Rst("Balance"))
Balance = 0
While Not Rst.EOF
    
    If OldTransType = 1 Then NewTransType = wDeposit
    If OldTransType = -1 Then NewTransType = wWithDraw
    If OldTransType = 3 Then NewTransType = wContraDeposit
    If OldTransType = -3 Then NewTransType = wContraWithDraw
    
    IsIntTrans = False: IsPaybleTrans = False
    OldTransType = FormatField(Rst("TransType"))
    oldAccid = Rst("AccId")
    
    If OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Then IsIntTrans = True
    If OldTransType = -4 Or OldTransType = 4 Then IsPaybleTrans = True
    If OldTransType = -5 Or OldTransType = 5 Then IsPaybleTrans = True
    
    'If the last record's Transaction id is greater or equal to present transid then
    'It means that the account no has been changed
    'If OldTransID >= FormatField(Rst("TransID")) Then
    If OldTransID >= FormatField(Rst("TransID")) Or AccID <> Rst("AccID") Then
        'AccID = AccID + 1
        AccID = Rst("AccID")
        PayableBalance = 0
        IntBalance = 0
        TransID = Rst("TransID") - 1
        Balance = -1
    End If
    
    TransID = TransID + 1
    Amount = Rst("Amount")
    If OldTransType = 0 Then GoTo NextAccount
    'NewTransType = OldTransType / Abs(OldTransType)
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    If IsPaybleTrans Then
        If OldTransType = 4 Then NewTransType = wContraDeposit
        If OldTransType = 5 Then NewTransType = wContraWithDraw
         'The above transactions also effect the profit & loss
        If OldTransType = -5 Then NewTransType = wContraDeposit
        If OldTransType = -4 Then NewTransType = wDeposit
        TransDate = Rst("TransDate")
        PayableBalance = PayableBalance + Rst("Amount") * IIf(OldTransType > 0, 1, -1)
        If PayableBalance < 0 Then PayableBalance = 0
        
        SqlPayable = "Insert INTO RDIntPayable ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlPayable = SqlPayable & " VALUES (" & _
            AccID & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            PayableBalance & " ," & _
            NewTransType & " )"
        
        IntBalance = IntBalance + Rst("Amount")
        If OldTransType = 4 Then
            NewTransType = wContraWithDraw
            SqlInt = "Insert INTO RDIntTrans ( " & _
                "AccID,TransID,TransDate," & _
                "Amount,Balance," & _
                "TransType )"
            SqlInt = SqlInt & "VALUES (" & _
                AccID & "," & _
                TransID & "," & _
                "#" & Rst("TransDate") & "#," & _
                Rst("Amount") & "," & _
                IntBalance & "," & _
                NewTransType & " )"
        End If
        If Rst("Amount") = 0 Then SqlPayable = "": SqlInt = ""
        If OldTransType = 5 Then
            Rst.MoveNext
            'Check the Next TransCtion Date If Both are not same
            'then do Not Count the Next record for this transaction
            'if Next Transactiojn is not of withdraw
            If TransDate <> Rst("TransDate") Or Abs(Rst("Transtype")) <> 1 Then Rst.MovePrevious: GoTo NextAccount
            OldTransType = Rst("Transtype")
        End If
    End If
    
    If IsIntTrans Then
        If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithDraw
        If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit
        IntBalance = IntBalance + Rst("Amount")
        TransDate = Rst("TransDate")
        SqlInt = "Insert INTO RDIntTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            AccID & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            IntBalance & " ," & _
            NewTransType & " )"
        
        If Rst("Amount") = 0 Then SqlInt = ""
        
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        If Balance = FormatField(Rst("balance")) Then Rst.MoveNext Else GoTo NextAccount
        OldTransType = Rst("TransType")
    End If
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    If OldTransType = 1 Then NewTransType = wDeposit
    If OldTransType = -1 Then NewTransType = wWithDraw
    If OldTransType = 3 Then NewTransType = wContraDeposit
    If OldTransType = -3 Then NewTransType = wContraWithDraw
    'If OldTransType = -3 Then NewTransType = wContraWithDraw
    
    SqlStr = "Insert INTO RDTrans ( " & _
        "AccID,TransID,TransDate," & _
        "Amount,Balance,Particulars," & _
        "TransType,VoucherNo)"
    
    SqlStr = SqlStr & "VALUES (" & _
        AccID & "," & _
        TransID & "," & _
        "#" & Rst("TransDate") & "#," & _
        Rst("Amount") & "," & _
        Rst("Balance") & "," & _
        AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
        NewTransType & "," & FormatField(Rst("ChequeNo")) & " )"
    
    If Balance = Rst("Balance") And OldTransID < 10000 Then SqlStr = ""
    If oldAccid <> Rst("accid") Then SqlStr = ""
    
    NewRDTrans.BeginTrans
        
    If SqlStr <> "" Then
        NewRDTrans.SQLStmt = SqlStr
        If Not NewRDTrans.SQLExecute Then
            MsgBox "Unable to transafer the Rd Trans data"
            NewRDTrans.RollBack
            Exit Function
        End If
    End If
    If SqlInt <> "" Then
        NewRDTrans.SQLStmt = SqlInt
        If Not NewRDTrans.SQLExecute Then
            MsgBox "Unable to transafer the RD Trans data"
            NewRDTrans.RollBack
            Exit Function
        End If
        SqlInt = ""
    End If
    If SqlPayable <> "" Then
        NewRDTrans.SQLStmt = SqlPayable
        If Not NewRDTrans.SQLExecute Then
            MsgBox "Unable to transafer the RD Trans data"
            NewRDTrans.RollBack
            Exit Function
        End If
        SqlPayable = ""
    End If
    'If Rst.AbsolutePosition Mod 5000 = 0 Then Debug.Print Now
    NewRDTrans.CommitTrans
    Balance = FormatField(Rst("Balance"))
    OldTransID = Rst("TransID")
    
NextAccount:
    With frmMain
        .lblProgress = "Transferring recurring transaction"
        .prg.Value = Rst.AbsolutePosition
    End With
    Rst.MoveNext

Wend

Debug.Print Now & "  " & Rst.RecordCount


Exit_line:
TransferRDTrans = True
Exit Function


Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "Error in RD Trans" & Err.Description
        'Resume
    End If
    
End Function
'this function is used to transfer the
'set the voucher no fot the transaferred data
Private Function PutVoucherNumber(NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String

PutVoucherNumber = True
End Function
