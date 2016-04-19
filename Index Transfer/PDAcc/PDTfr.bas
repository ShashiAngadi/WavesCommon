Attribute VB_Name = "PDTransfer"
'This BAs file is used to Transfer
'Pigmy Master & pigmyTranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit
Private m_AccOffSet As Long
Private m_LoanOffSet As Long
Private m_DepositType As wis_DepositType



Private Function CreatePigmyHeads(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean

Dim HeadID As Long

Dim ClsBank As clsBankAcc


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
        If rstTemp("Module") = 57 Then _
            HeadBalance = FormatField(rstTemp("ObAmount"))
        If rstTemp("Module") = 58 Then _
            LoanBalance = FormatField(rstTemp("ObAmount"))
        If HeadBalance And LoanBalance Then Exit Do
        rstTemp.MoveNext
    Loop
End If

'Prefix = IIf(gLangOffSet, "¼—ó ¬ÙÓÀÐ±Ò", "Pigmy Deposit")
Prefix = LoadResString(gLangOffSet + 425)

CreatePigmyHeads = True

On Error GoTo ErrLine
Dim PrgVal As Integer

frmMain.lblProgress = "Transferring the Pigmy ledger"
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

FromDate = gFromDate ' "3/31/03"

'First Insert the Deposit Transaction
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM PDTrans " & _
                "Where TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
    
    NewIndexTrans.BeginTrans
    'Pigmy  Deposit Heads
    HeadName = Prefix
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemberDeposit, HeadBalance, wis_PDAcc)

    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            'Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
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

    'Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert the interest transaction of pigmy deposit
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM PDIntTrans " & _
                "Where TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit Interest Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 487)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_PDAcc)

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

'Now insert Interest payable
NewTrans.SQLStmt = "Select Sum(Amount) As TotalAmount ," & _
                "TransType,TransDate FROM PDIntPayable " & _
                "Where TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType "
                
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    HeadBalance = 0
    OldTrans.SQLStmt = "Select Top 1 Balance From AccTrans " & _
                " WHERE TransDate <= #" & gFromDate & "#" & _
                " AND AccId = 13002 ORder By TransID Desc"
    If OldTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                HeadBalance = FormatField(rstTemp("Balance"))


    NewIndexTrans.BeginTrans
    'Deposit payable heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 375) & " " & LoadResString(gLangOffSet + 47)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parDepositIntProv, HeadBalance, wis_PDAcc)

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

'NOW INSERT THE TRANSACTION PIGMY DEPOSIT LOAN
'Insert the Deposit loan Transaction
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM DepositLoanTrans " & _
                "Where LoanId in (Select LoanID From DepositLoanMaster " & _
                    "Where DepositType = " & m_DepositType & ")" & _
                " And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    

    NewIndexTrans.BeginTrans
    'Deposit Interest Provision Head
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 58)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemDepLoan, LoanBalance, wis_DepositLoans + wis_PDAcc)

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
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 58) & " " & _
            LoadResString(gLangOffSet + 483)
    HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemDepLoanIntReceived, 0, wis_DepositLoans + wis_PDAcc)

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

'just calling this function we can transafer the PDmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferPD(OldDBName As String, NewDBName As String) As Boolean
Debug.Print Now
Dim OldTrans As New clsOldUtils
Dim NewTrans As New clsDBUtils
If Not OldTrans.OpenDB(OldDBName, OldPwd) Then Exit Function
If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    Exit Function
End If

Screen.MousePointer = vbHourglass

If Not gOnlyLedgerHeads Then
    If Not TransferPDMaster(OldTrans, NewTrans) Then GoTo ExitLine
    If Not TransferPDTrans(OldTrans, NewTrans) Then GoTo ExitLine
    If Not TransferPDLoanMaster(OldTrans, NewTrans) Then GoTo ExitLine
    If Not TransferPDLoanTrans(OldTrans, NewTrans) Then GoTo ExitLine
End If

If Not CreatePigmyHeads(OldTrans, NewTrans) Then GoTo ExitLine
    
    TransferPD = True
    If Not PutVoucherNumber(NewTrans) Then
         MsgBox "Unable to set the voucher No"
        GoTo ExitLine
    End If
    
ExitLine:
Screen.MousePointer = vbDefault
End Function
'this function is used to transfer the
'Pigmy  MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferPDMaster(OldPdTrans As clsOldUtils, NewPDTrans As clsDBUtils) As Boolean
Dim SqlStr As String

Dim Rst As Recordset
Dim rstTemp As Recordset

Dim ProcCount As Long

On Error GoTo Err_Line


'Fetch the detials of Pifmy  Account
SqlStr = "SELECT * FROM PDMASTER ORDER BY UserId,AccID"
Dim AccID As Long, IntroId As Long
Dim UserID As Long, AccNum As String
OldPdTrans.SQLStmt = SqlStr

If OldPdTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo Exit_line
'Set Rst = OldPdTrans.Rst.Clone

'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldPdTrans.SQLStmt = SqlStr
'Call OldPdTrans.SQLFetch

'Before Insering the Record delete if any records are there in the
'PDMaster due to invonvince of the last transfr
NewPDTrans.BeginTrans
SqlStr = "DELETE * FROM PDMASTER"
NewPDTrans.SQLStmt = SqlStr
If Not NewPDTrans.SQLExecute Then
    NewPDTrans.RollBack
    MsgBox "Unable to transafer the pigmy Master data"
    Exit Function
End If
NewPDTrans.CommitTrans


With frmMain
    .lblProgress = "Transferring pigmy account details"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With


'm_AccOffSet= FormatField(OldPdTrans.Rst(0))
'm_AccOffSet= AccOffSet + 100 - (AccOffSet Mod 100)

AccID = m_AccOffSet
While Not Rst.EOF
    'If AccID = FormatField(Rst("AccID")) Then GoTo NextAccount
    IntroId = FormatField(Rst("Introduced"))
    'Get the Introducer ID
    
    If IntroId > 0 Then
        SqlStr = "SELECT CustomerID FROM PDMASTER " & _
            " WHERE AccID = " & IntroId
        OldPdTrans.SQLStmt = SqlStr
        If OldPdTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then IntroId = FormatField(rstTemp("CustomerID"))
    End If
        
    'AccNum = Format(Rst("AccId"), "000")
    'AccNum = Rst("UserID") & "_" & Rst("AccId")
    AccNum = Rst("AccId")
    NewPDTrans.BeginTrans
    
'NOW insert into
    'AccID = Val(Rst("UserID")) * AccOffSet + Val(Rst("AccId"))
    'AccID = AccID - AccOffSet
    AccID = AccID + 1
    
    SqlStr = "Insert INTO PDMASTER (" & _
        "AccID,AgentID,CustomerID,AccNum," & _
        "CreateDate,ModifiedDate,ClosedDate," & _
        "MaturityDate,PigmyAmount,PigmyType," & _
        "RateOfInterest,Nominee,Introduced," & _
        " LedgerNo,FolioNo,NomineeID,LastPrintId,AccGroupID )"
    
    SqlStr = SqlStr & " VALUES (" & _
        AccID & "," & _
        Rst("UserID") & "," & _
        Rst("CustomerID") & "," & _
        AddQuotes(AccNum, True) & "," & _
        FormatDateField(Rst("CreateDate")) & "," & _
        FormatDateField(Rst("Modifieddate")) & "," & _
        FormatDateField(Rst("ClosedDate")) & " ," & _
        FormatDateField(Rst("MaturityDate")) & " ," & _
        Rst("PigmyAmount") & "," & _
        AddQuotes(FormatField(Rst("PigmyType")), True) & "," & _
        Rst("RateOfInterest") & "," & _
        AddQuotes(FormatField(Rst("Nominee")), True) & "," & _
        Rst("Introduced") & "," & _
        AddQuotes(FormatField(Rst("LedgerNo")), True) & "," & _
        AddQuotes(FormatField(Rst("FolioNo")), True) & " ," & _
        "0 ," & _
        "1, 1 )"
        
    NewPDTrans.SQLStmt = SqlStr
    If Not NewPDTrans.SQLExecute Then
        NewPDTrans.RollBack
        'Now Check related Customer Is Missed
        NewPDTrans.SQLStmt = "SELECT * From NameTab Where CustomerID = " & Rst("CustomerID")
        If NewPDTrans.Fetch(rstTemp, adOpenDynamic) = 0 Then GoTo NextAccount
        NewPDTrans.SQLStmt = "SELECT * From UserTab Where UserID = " & Rst("UserID")
        If NewPDTrans.Fetch(rstTemp, adOpenDynamic) = 0 Then GoTo NextAccount
        
        MsgBox "Unable to transfer the pigmy MAster data"
        Exit Function
    End If
    NewPDTrans.CommitTrans
    
NextAccount:
    ProcCount = ProcCount + 1
    With frmMain
        .lblProgress = "Transferring pigmy accounts"
        .prg.Value = ProcCount
        If ProcCount Mod 50 = 0 Then .Refresh
    End With
    Rst.MoveNext
Wend

TransferPDMaster = True

'Now reverse the change made before transfer
'Update Modify date
SqlStr = "UPDATE PDMASTER Set ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
OldPdTrans.SQLStmt = SqlStr
OldPdTrans.BeginTrans
If Not OldPdTrans.SQLExecute Then
    OldPdTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldPdTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE PDMASTER Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
OldPdTrans.SQLStmt = SqlStr
OldPdTrans.BeginTrans
If Not OldPdTrans.SQLExecute Then
    OldPdTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldPdTrans.CommitTrans

'Now Update the smae with new database
'Update Modify date
SqlStr = "UPDATE PDMAster Set ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
NewPDTrans.SQLStmt = SqlStr
NewPDTrans.BeginTrans
If Not NewPDTrans.SQLExecute Then
    NewPDTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewPDTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE PDMAster Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
NewPDTrans.SQLStmt = SqlStr
NewPDTrans.BeginTrans
If Not NewPDTrans.SQLExecute Then
    NewPDTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewPDTrans.CommitTrans
    

Exit_line:
TransferPDMaster = True
    Debug.Print Now
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next

    If Err Then
        MsgBox "eror In SBMaster " & Err.Description
        'Resume
    End If
    
End Function



'this function is used to transfer the
'SB MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferPDLoanMaster(OldPdTrans As clsOldUtils, NewPDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim Sql_2 As String

Dim Rst As Recordset
Dim rstTemp As Recordset

Dim ProcCount As Long

On Error GoTo Err_Line

'Now TransFer the Loan Details
'Get the account Having Details
With frmMain
    .lblProgress = "Transferring pigmy Loan details"
    .prg.Value = 0
    .Refresh
End With

SqlStr = "SELECT A.UserId,A.AccID,CustomerID,MaturityDate," & _
    " TransDate,RateOfInterest,Amount,TransID,LedgerNo,FolioNo " & _
    " FROM PDTrans A, PDMAster B WHERE A.UserId = B.UserID " & _
    " AND A.AccID = B.AccID And TransId = (SELECT Min(TransID) " & _
        " From PDTrans C Where C.USerID = B.UserID " & _
        " AND C.AccID = B.accID AND Loan = True ) AND Loan = TRue "

OldPdTrans.SQLStmt = SqlStr
If OldPdTrans.Fetch(Rst, adOpenDynamic) <= 0 Then GoTo ExitLine
'Set Rst = OldPdTrans.Rst.Clone

Dim AccID As Long
Dim LoanId As Long
Dim LoanNum As String
Dim AccNum As String

'Get the Account Offset From the oLddataBase

NewPDTrans.SQLStmt = "DELETE * FROM DepositLoanMaster Where DepositType = " & wisDeposit_PD
NewPDTrans.BeginTrans
Call NewPDTrans.SQLExecute
NewPDTrans.CommitTrans

NewPDTrans.SQLStmt = "SELECT MAX(LOanId) FROM DepositLoanMaster"
If NewPDTrans.Fetch(rstTemp, adOpenDynamic) Then m_LoanOffSet = FormatField(rstTemp(0))

SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldPdTrans.SQLStmt = SqlStr
'Call OldPdTrans.SQLFetch
'm_AccOffSet= FormatField(OldPdTrans.Rst(0))
'm_AccOffSet= AccOffSet + 100 - (AccOffSet Mod 100)

With frmMain
    .lblProgress = "Transferring pigmy Loan accounts"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With


LoanId = m_LoanOffSet
AccID = m_AccOffSet

While Not Rst.EOF
    AccNum = Rst("UserID") & "_" & Rst("AccId")
    LoanNum = Rst("UserID") & "_" & Rst("AccId")
    'LoanNum = Rst("Accid")
    LoanId = LoanId + 1
    'Get the Account ID
    NewPDTrans.BeginTrans
    
    
'''' Here  we are inserting into the new table called DepositLoanMAster
''the above said table will be common for all type deposit(eg. FD,Rd,Pd,DL)
    m_DepositType = wisDeposit_PD
    SqlStr = "Insert INTO PledgeDeposit (" & _
        "LoanID,AccID,DepositType,PledgeNum)" & _
        " VALUES (" & _
        LoanId & "," & _
        AccID & "," & _
        m_DepositType & "," & _
        " 1 )"
            
    NewPDTrans.SQLStmt = SqlStr
    If Not NewPDTrans.SQLExecute Then
        NewPDTrans.RollBack
        MsgBox "Unable to transafer the pigmy MAster data"
        Exit Function
    End If
    
    SqlStr = "Insert INTO DepositLoanMASTER (" & _
        " CustomerID,LoanID,AccNum,DepositType," & _
        " LoanIssuedate,LoanDueDate,PledgeDescription, " & _
        " InterestRate,LoanAmount,LedgerNo,FolioNo ," & _
        " LastPrintId )"
    SqlStr = SqlStr & " VALUES (" & _
        Rst("CustomerID") & "," & LoanId & "," & _
        AddQuotes(LoanNum, True) & "," & _
        m_DepositType & "," & _
        "#" & Rst("TransDate") & "#, #" & Rst("MaturityDate") & "#," & _
        AccID & " ," & _
        Rst("RateOfInterest") & "," & _
        Rst("AMount") & "," & _
        AddQuotes(Rst("LedgerNo"), True) & "," & _
        "'" & Val(Rst("FolioNo")) & "'," & _
        "1 )"
        
    NewPDTrans.SQLStmt = SqlStr
    If Not NewPDTrans.SQLExecute Then
        NewPDTrans.RollBack
        MsgBox "Unable to transafer the pigmy loan data"
        Exit Function
    End If
    
    SqlStr = "UPdate PDMaster Set LoanId = " & LoanId & " Where AccNum = '" & AccNum & "'"
    NewPDTrans.SQLStmt = SqlStr
    If Not NewPDTrans.SQLExecute Then
        NewPDTrans.RollBack
        MsgBox "Unable to transafer the pigmy loan data"
        Exit Function
    End If
    
    NewPDTrans.CommitTrans
    
NextAccount:
    ProcCount = ProcCount + 1
    With frmMain
        .lblProgress = "Transferring pigmy Loan accounts"
        .prg.Value = ProcCount
        If ProcCount Mod 50 = 0 Then .Refresh
    End With
    Rst.MoveNext
    
Wend

ExitLine:
TransferPDLoanMaster = True

    Debug.Print Now
Exit Function

Err_Line:


If Err.Number = 3021 Then Err.Clear: Resume Next

    If Err Then MsgBox "eror In Pigmy LoanMaster " & Err.Description
    'Resume
End Function


'this function is used to transfer the
'pigmy transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferPDTrans(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim SqlPayable As String

Dim IsIntTrans As Boolean
Dim IsPaybleTrans As Boolean
Dim PayableBalance As Currency
Dim IntBalance As Currency
On Error GoTo Err_Line

Dim OldTransType As Integer
Dim NewTransType As wisTransactionTypes
Dim Balance As Currency
Dim TransID As Long
Dim NewTransId As Long

Dim Rst As Recordset
Dim AccID As Long

Dim oldAccid As Long
Dim OldUserID As Long

Dim Amount As Currency
Dim TransDate As Date

Dim ProcCount As Long

    'Fetch the detials of Sb Account
With frmMain
    .lblProgress = "Transferring pigmy transaction"
    .prg.Value = 0
    .Refresh
End With

SqlStr = "SELECT * FROM PDTrans Where Loan = False " & _
    "ORDER BY UserID,AccID,TransId"

OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo ExitLine
'Set Rst = OldTrans.Rst.Clone

'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldTrans.SQLStmt = SqlStr
'Call OldTrans.SQLFetch
'm_AccOffSet= FormatField(OldTrans.Rst(0))
'm_AccOffSet = AccOffSet + 100 - (AccOffSet Mod 100)

With frmMain
    .lblProgress = "Transferring pigmy transaction"
    .prg.Value = 0
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With

TransID = 10000000: AccID = m_AccOffSet
Dim AccNum As String
Dim RstAcc As Recordset

While Not Rst.EOF
    IsIntTrans = False: IsPaybleTrans = False
    OldTransType = Rst("TransType")
    If OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Then IsIntTrans = True
    If OldTransType = -4 Or OldTransType = 4 Then IsPaybleTrans = True
    If OldTransType = -5 Or OldTransType = 5 Then IsPaybleTrans = True
    
    'If the last record's Transaction id is greater or equal to present transid then
    'It means that the account no has been changed
    'If TransID <= rst("TransID") Then
    If AccNum <> CStr(Rst("UserID") & "_" & Rst("AccId")) Then
        AccID = AccID + 1
        'get the AccID
        AccNum = Rst("UserID") & "_" & Rst("AccId")
        NewTrans.SQLStmt = "SELECT accid FROM PDMAster Where AccNum = '" & AccNum & "';"
        If NewTrans.Fetch(RstAcc, adOpenForwardOnly) Then AccID = FormatField(RstAcc("AccId"))
        'End If
        PayableBalance = 0
        IntBalance = 0
        TransID = Rst("TransID")
        NewTransId = 0
        Balance = -1
    End If
    TransID = TransID + 1: NewTransId = NewTransId + 1
    
    Amount = Rst("Amount")
    oldAccid = Rst("accId")
    OldUserID = Rst("UserID")
                
    'NewTransType = OldTransType / Abs(OldTransType)
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    If IsPaybleTrans Then
        If OldTransType = 4 Then NewTransType = wContraDeposit
        If OldTransType = 5 Then NewTransType = wContraWithDraw
        'The above transactions also effect the profit & loss
        If OldTransType = -5 Then NewTransType = wContraDeposit
        If OldTransType = -4 Then NewTransType = wDeposit
        TransDate = Rst("TransDate")
        PayableBalance = PayableBalance + Rst("Amount") * IIf(NewTransType > 0, 1, -1)
        If PayableBalance < 0 Then PayableBalance = 0
        
        If OldTransType = 4 Then NewTransType = wContraDeposit
        If OldTransType = 5 Then NewTransType = wWithDraw
        SqlPayable = "Insert INTO PDIntPayable ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlPayable = SqlPayable & "VALUES (" & _
            AccID & "," & _
            NewTransId & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            PayableBalance & " ," & _
            NewTransType & " )"
        
        If OldTransType = 4 Then
            NewTransType = wContraWithDraw
            IntBalance = IntBalance + Rst("Amount")
            SqlInt = "Insert INTO PDIntTrans ( " & _
                "AccID,TransID,TransDate," & _
                "Amount,Balance," & _
                "TransType )"
            SqlInt = SqlInt & "VALUES (" & _
                AccID & "," & _
                NewTransId & "," & _
                "#" & Rst("TransDate") & "#," & _
                Rst("Amount") & "," & _
                IntBalance & " ," & _
                NewTransType & " )"
        End If
        If Rst("Amount") = 0 Then SqlPayable = "": SqlInt = ""
        
         'If TransType is  5 then he has withdraw the amount from the Interest Payble
        'and in the next transaction he might have closed the deposit
        'So move the record and check the next transaction
        If OldTransType = 5 Then
            Rst.MoveNext
            'Check the TransCtion Date If Both are not same
            'then do Not Count the the record for this transaction
            'if Next Transactiojn is not of withdraw
            If TransDate <> Rst("TransDate") Or Abs(Rst("Transtype")) <> 1 Then Rst.MovePrevious
        End If
    End If
    If IsIntTrans Then
        'If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithDraw
        'If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit
        NewTransType = IIf(OldTransType > 0, wWithDraw, wDeposit)
        TransDate = Rst("TransDate")
        IntBalance = IntBalance + Rst("Amount")
        SqlInt = "Insert INTO PDIntTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            AccID & "," & _
            NewTransId & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0," & _
            NewTransType & " )"
        Amount = Rst("Amount")
        
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        If Balance = FormatField(Rst("balance")) Then Rst.MoveNext
        OldTransType = Rst("TransType")
        'After this transaction the transaction in the PDTable is contra
        'Therefore
        NewTransType = (OldTransType / Abs(OldTransType)) * 3
    End If
    'If the transaction is payble then Need not do the '
    'transaction ir depsoit accounts
   'NewTransType = (OldTransType / Abs(OldTransType))
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    SqlStr = "Insert INTO PDTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
    SqlStr = SqlStr & "VALUES (" & _
            AccID & "," & _
            NewTransId & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            Rst("Balance") & "," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
            NewTransType & " )"
    
    If Balance = Rst("Balance") Then SqlStr = ""
    If oldAccid <> Rst("AccId") Or OldUserID <> Rst("UserId") Then SqlStr = ""
    If Rst("Amount") = 0 Then SqlStr = ""
    
    NewTrans.BeginTrans
    If SqlStr <> "" Then
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy trans data"
            NewTrans.RollBack
            Exit Function
        End If
        SqlStr = ""
    End If
    If SqlInt <> "" Then
        NewTrans.SQLStmt = SqlInt
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy Trans data"
            NewTrans.RollBack
            Exit Function
        End If
        SqlInt = ""
    End If
    If SqlPayable <> "" Then
        NewTrans.SQLStmt = SqlPayable
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy Trans data"
            NewTrans.RollBack
            Exit Function
        End If
        SqlPayable = ""
    End If
    
    NewTrans.CommitTrans
    Balance = Rst("Balance")

NextAccount:
    ProcCount = ProcCount + 1
    With frmMain
        .lblProgress = "Transferring pigmy transaction"
        .prg.Value = ProcCount
    End With
    Rst.MoveNext
    
Wend


ExitLine:
TransferPDTrans = True
Exit Function


Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next

    If Err Then
        MsgBox "Error in PDTrans" & Err.Description
        'Resume
        Err.Clear
    End If
    
End Function


'this function is used to transfer the
'pigmy Loan transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferPDLoanTrans(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim IsIntTrans As Boolean

Dim OldUserID As Long
Dim oldAccid As Long

On Error GoTo Err_Line

Dim OldTransType As Integer, NewTransType As wisTransactionTypes
Dim Balance As Currency
Dim TransID As Long
Dim NewTransId As Long
Dim Rst As Recordset
Dim rstTemp As Recordset
Dim LoanId As Long
Dim Amount As Currency
Dim TransDate As Date
Dim DepositType As wis_DepositType
Dim ProcCount As Long

DepositType = wisDeposit_PD
    
'Fetch the detials of pigmy Account
With frmMain
    .lblProgress = "Transferring pigmy Loan transaction"
    .prg.Value = 0
    .Refresh
End With

SqlStr = "SELECT * FROM PDTrans Where Loan = True " & _
    " ORDER BY UserID,AccID,TransId"

OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo ExitLine
'Set Rst = OldTrans.Rst.Clone

'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldTrans.SQLStmt = SqlStr
Call OldTrans.Fetch(rstTemp, adOpenDynamic)
m_AccOffSet = FormatField(rstTemp(0))
m_AccOffSet = m_AccOffSet + 100 - (m_AccOffSet Mod 100)


With frmMain
    .lblProgress = "Transferring pigmy Loan transaction"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With


TransID = 100000
LoanId = m_LoanOffSet
Balance = 0
While Not Rst.EOF
    IsIntTrans = False: SqlInt = ""
    OldTransType = FormatField(Rst("TransType"))
    
    If OldTransType = 4 Or OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Or OldTransType = -4 Then IsIntTrans = True
    
    If TransID >= Rst("TransID") Then
'        LoanId = LoanId + 1
        NewTrans.SQLStmt = "Select LoanID From DepositLoanMaster" & _
            " Where AccNum = '" & Rst("UserID") & "_" & Rst("AccId") & "'" & _
            " And DepositType = " & wisDeposit_PD
        If NewTrans.Fetch(rstTemp, adOpenDynamic) < 0 Then GoTo NextAccount
        If LoanId = rstTemp("LoanID") Then
            LoanId = LoanId
            If rstTemp.RecordCount > 1 Then rstTemp.MoveNext: LoanId = rstTemp("LoanID"): NewTransId = 0
        Else
            LoanId = rstTemp("LoanID")
            NewTransId = 0
        End If
        TransID = Rst("TransID") - 1
        Balance = -1
    End If
    
    TransID = TransID + 1
    NewTransId = NewTransId + 1
    Amount = Rst("Amount")
    OldUserID = Rst("UserId")
    oldAccid = Rst("AccID")
    
    NewTransType = OldTransType / Abs(OldTransType)
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    If IsIntTrans Then
        'If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithDraw  'interest Paid to the customer
        'If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit  'Interest collected form the customer
        NewTransType = IIf(OldTransType > 0, wWithDraw, wDeposit)
        TransDate = Rst("TransDate")
        SqlInt = "Insert INTO DepositLoanIntTrans ( " & _
            "LoanId,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            LoanId & "," & _
            NewTransId & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0," & _
            NewTransType & " )"
        If Amount = 0 Then SqlInt = ""
        Rst.MoveNext
        Amount = Rst("Amount")
    End If
    
    OldTransType = FormatField(Rst("TransType"))
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    
    ''INSERT INTO DEPOSITLOANTRANS
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
    
    If oldAccid = Rst("AccID") Or OldUserID <> Rst("UserID") Then SqlStr = ""
    If Balance <> Rst("Balance") Then SqlStr = ""
    If Rst("Amount") = 0 Then SqlStr = ""
    
    NewTrans.BeginTrans
    If SqlStr <> "" Then
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy loan Trans "
            NewTrans.RollBack
            Exit Function
        End If
    End If
    If SqlInt <> "" Then
        NewTrans.SQLStmt = SqlInt
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy loan Trans "
            NewTrans.RollBack
            Exit Function
        End If
    End If
    NewTrans.CommitTrans
    
    Balance = FormatField(Rst("Balance"))

NextAccount:
    ProcCount = ProcCount + 1
    With frmMain
        .lblProgress = "Transferring pigmy loan transaction"
        .prg.Value = ProcCount + 1
        If ProcCount Mod 50 = 0 Then .Refresh
    End With
    Rst.MoveNext

Wend

ExitLine:

TransferPDLoanTrans = True

    With frmMain
        .lblProgress = "Transferred the pigmy detials"
        .prg.Value = 0
        .Refresh
    End With


Exit Function


Err_Line:


If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "Error in PD LOAn Trans" & Err.Description
        'Resume
    End If
    
End Function

'this function is used to transfer the
'set the voucher no fot the transaferred data
Private Function PutVoucherNumber(NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String
PutVoucherNumber = True
End Function
