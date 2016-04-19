Attribute VB_Name = "FDTransfer"
'This BAs file is used to Transfer
'FD Master & FD Transaction dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit
Dim m_AccOffSet As Long
Dim m_LoanOffSet As Long
Dim m_DepositType As Integer





Private Function CreateFDHeads(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean

Dim AccHeadId As Long

Dim ClsBank As clsBankAcc

On Error GoTo ExitLine
'First Creat the Dl Deposit Heads

Set ClsBank = New clsBankAcc
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans


'First Create the fixed Deposit Heads
Dim HeadName As String
Dim Prefix As String
frmMain.lblProgress = "Transferring the Fixed deposit"
frmMain.Refresh
Dim PrgVal As Integer

With frmMain.prg
    .Max = 365
    .Min = 0
    .Value = 0
End With


Prefix = IIf(gLangOffSet, "¿–‘‘∑–Ï≥–‘Í ¨Ÿ”¿–±“", "Fixed Deposit")


Dim FromDate As Date
Dim rstTrans As ADODB.Recordset
Dim TransDate As Date
Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency
Dim TransType As wisTransactionTypes


'Get the Head Balance
Dim HeadBalance As Currency
Dim LoanBalance As Currency

Dim rstTemp As Recordset

OldTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#"
If OldTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Set rstTemp = Nothing
            
FromDate = gFromDate '"3/31/03"

'First Insert the Deposit Transaction
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM FDTrans " & _
                "Where AccId in (Select AccID From FdMaster " & _
                    "Where DepositType = " & m_DepositType & ") " & _
                "And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    'Get the Head Balance
    If Not rstTemp Is Nothing Then
        Do
            If rstTemp.EOF Then Exit Do
            If rstTemp("Module") = 53 Then _
                HeadBalance = FormatField(rstTemp("ObAmount"))
            If rstTemp("Module") = 54 Then _
                LoanBalance = FormatField(rstTemp("ObAmount"))
            If HeadBalance And LoanBalance Then Exit Do
            rstTemp.MoveNext
        Loop
    End If
    NewIndexTrans.BeginTrans
    'Fixed Deposit Heads
    HeadName = Prefix
    AccHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemberDeposit, HeadBalance, wis_Deposits + m_DepositType)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
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

    Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert the interest transaction of dl deposit
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM FDIntTrans " & _
                "Where AccId in (Select AccID From FdMaster " & _
                    "Where DepositType = " & m_DepositType & ") " & _
                "And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"
                
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit Interest Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 487)
    AccHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_Deposits + m_DepositType)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
            'Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalDeposit, TransDate)
            'Call ClsBank.UpdateCashDeposits(AccHeadId, TotalWithdraw, TransDate)
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
    Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
    'Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalDeposit, TransDate)
    'Call ClsBank.UpdateCashDeposits(AccHeadId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert Interest payable
NewTrans.SQLStmt = "Select Sum(Amount)as TotalAmount, " & _
                "TransType,TransDate FROM FDIntPayable " & _
                "Where AccId in (Select AccID From FdMaster " & _
                    "Where DepositType = " & m_DepositType & ") " & _
                "And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    HeadBalance = 0
    OldTrans.SQLStmt = "Select Top 1 Balance From AccTrans " & _
                " WHERE TransDate <= #" & gFromDate & "#" & _
                " AND AccID = 13001 ORder By TransID Desc"
    If OldTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                HeadBalance = FormatField(rstTemp("Balance"))

    NewIndexTrans.BeginTrans
    'Deposit Interest Provision Head
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 375) & " " & LoadResString(gLangOffSet + 47)
    AccHeadId = ClsBank.GetHeadIDCreated(HeadName, parDepositIntProv, HeadBalance, wis_Deposits + m_DepositType)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
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
    Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
End If

'NOW INSERT THE TRANSACTION FIXED DEPOSIT LOAN
'Insert the Deposit loan Transaction
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM DepositLoanTrans " & _
                "Where LoanId in (Select LoanID From DepositLoanMaster " & _
                    "Where DepositType = " & m_DepositType & ") " & _
                "And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit LOan Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 58)
    AccHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemDepLoan, LoanBalance, wis_DepositLoans + m_DepositType)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
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

    Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now insert the interest transaction of deposit Loans
NewTrans.SQLStmt = "Select Sum(Amount) as TotalAmount, " & _
                "TransType,TransDate FROM DepositLoanIntTrans " & _
                "Where loanId in (Select LOanID From DepositLoanMaster " & _
                    "Where DepositType = " & m_DepositType & ") " & _
                "And TransDate >= #" & FromDate & "# " & _
                "Group By TransDate,TransType " & _
                "Order By TransDate,TransType"
                
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Deposit Interest Heads
    HeadName = Prefix & " " & LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 483)
    AccHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemDepLoanIntReceived, 0, wis_DepositLoans + m_DepositType)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
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
    Call ClsBank.UpdateCashDeposits(AccHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(AccHeadId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If


CreateFDHeads = True

ExitLine:
    
If Err.Number = 380 Then
    frmMain.prg.Max = PrgVal * 1.5
    Resume Next
ElseIf Err.Number Then
    MsgBox "Error In Chcking data Base"
    'GoTo Exit_line
    'Resume
End If
    
    Debug.Assert Err.Number = 0
    Set NewIndexTrans = Nothing
    Set rstTrans = Nothing
    Set ClsBank = Nothing
    

End Function


'just calling this function we can transafer the PDmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferFD(OldDBName As String, NewDBName As String) As Boolean

Dim OldTrans As clsOldUtils
Dim NewTrans As clsDBUtils

Set OldTrans = New clsOldUtils
Set NewTrans = New clsDBUtils

If Not OldTrans.OpenDB(OldDBName, OldPwd) Then Exit Function
If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    Exit Function
End If
    
If Not gOnlyLedgerHeads Then
    If Not TransferFDMaster(OldTrans, NewTrans) Then Exit Function
    If Not TransferFDTrans(OldTrans, NewTrans) Then Exit Function
    If Not TransferFDLoanMaster(OldTrans, NewTrans) Then Exit Function
    If Not TransferFDLoanTrans(OldTrans, NewTrans) Then Exit Function
End If

If Not CreateFDHeads(OldTrans, NewTrans) Then Exit Function
Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
    
    TransferFD = True
    
    If Not PutVoucherNumber(NewTrans) Then
        MsgBox "Unable to set the voucher No"
        Exit Function
    End If

End Function
'
'this function is used to transfer the
'FD MAster details form OLdb to new one
'and NewFDTrans has assigned to new database
Private Function TransferFDMaster(OldFDTrans As clsOldUtils, NewFDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim Rst As ADODB.Recordset
Dim DepAmount As Currency

On Error GoTo Err_Line

'Before Fetching Update the Values
'where It can be Null with default value
'Then Fetch the records
'Delete the any details in the PDMAster Whose transactions has not done
SqlStr = "SELECT Distinct DepositID FROM FDMASTER"
OldFDTrans.SQLStmt = SqlStr

If OldFDTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo Exit_line

While Not Rst.EOF
    SqlStr = "DELETE * FROM FDMAster WHERE DepositID = " & Rst(0) & _
        " AND AccID NOT IN (SELECT Distinct AccId From FDTrans " & _
        " WHERE DepositID = " & Rst(0) & ")"
    OldFDTrans.SQLStmt = SqlStr
    OldFDTrans.BeginTrans
    If Not OldFDTrans.SQLExecute Then
        OldFDTrans.RollBack
        MsgBox "UPdate the Data Base Instead Of transfer"
        Exit Function
    End If
    OldFDTrans.CommitTrans
    Rst.MoveNext
Wend

'Fetch the detials of Pifmy  Account
SqlStr = "SELECT * FROM FDMASTER ORDER BY AccID,DepositID"
OldFDTrans.SQLStmt = SqlStr

If OldFDTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Function

Dim AccID As Long, IntroId As Long
Dim DepositID As Integer, AccNum As String
Dim rstTemp As ADODB.Recordset

'Get The Deposit Type
m_DepositType = wisDeposit_FD
NewFDTrans.SQLStmt = "SELECT Max(DepositID) FROM DepositName"
If NewFDTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    m_DepositType = FormatField(rstTemp(0))
    m_DepositType = IIf(m_DepositType < wisDeposit_FD, wisDeposit_FD, m_DepositType) + 1
End If

Dim DepName  As String


NewFDTrans.SQLStmt = "SELECT Max(DepositID) FROM DepositName"
m_DepositType = wisDeposit_FD
If NewFDTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    m_DepositType = FormatField(rstTemp(0))
    m_DepositType = IIf(m_DepositType <= wisDeposit_FD, wisDeposit_FD, m_DepositType) + 1
End If

'DepName = LoadResString(gLangOffSet + 423)
DepName = IIf(gLangOffSet, "¿–‘‘∑–Ï≥–‘Í ¨Ÿ”¿–±“", "Fixed Deposit")

NewFDTrans.SQLStmt = "INSERT INTO DepositName " & _
    "(DepositID, DepositName,Cumulative) VALUES " & _
    "(" & m_DepositType & "," & AddQuotes(DepName, True) & ",False)"

NewFDTrans.BeginTrans
If Not NewFDTrans.SQLExecute Then
    NewFDTrans.RollBack
    MsgBox "Cannot Create The Deposit", vbCritical
Else
    NewFDTrans.CommitTrans
End If

'\\\\\\\\\\\\\\\\\\\\\\\
'Get the Account Offset From the new dataBase
SqlStr = "SELECT Max(AccID) FROM FDMASTER"
NewFDTrans.SQLStmt = SqlStr
If NewFDTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
        m_AccOffSet = FormatField(rstTemp(0))

AccID = m_AccOffSet
Dim CertNo As String

With frmMain
    .lblProgress = "Inserting the Fd Account Details"
    .prg.Min = 0
    .prg.Max = Rst.RecordCount + 1
    .prg.Value = 0
End With

Dim oldAccid As Long
Dim oldDepositID As Long

While Not Rst.EOF
        
        If oldAccid = Rst("AccID") And oldDepositID = Rst("DepositID") Then GoTo NextAccount
        Debug.Assert Rst("AccID") < 10000
        IntroId = FormatField(Rst("IntroducedID"))
        'Get the Introducer ID
        If IntroId > 0 Then
            SqlStr = "SELECT CustomerID FROM FDMASTeR " & _
                " WHERE AccID = " & IntroId
            OldFDTrans.SQLStmt = SqlStr
            If OldFDTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then IntroId = FormatField(rstTemp("CustomerID"))
        End If
        
        'Now get the DepositAmount
        OldFDTrans.SQLStmt = "SELECT Amount,Particulars FROM FDTrans WHERE AccID = " & Rst("AccID") & _
            " AND DepositID = " & Rst("DepositID") & _
            " AND Loan = False ORDER By TransID"
        
        If OldFDTrans.Fetch(rstTemp, adOpenForwardOnly) < 1 Then GoTo NextAccount
        DepAmount = rstTemp("Amount")
        CertNo = rstTemp("Particulars")
        If Val(CertNo) = 0 Then CertNo = Rst("Accid") * 1000 + Rst("DepositID")
    'NOW insert into FD MAster
        'Increase the AccID By 1
        AccID = AccID + 1
        AccNum = Rst("AccId")
        
        SqlStr = "Insert INTO FDMASTER (" & _
            "AccID,CustomerID,AccNum,CertificateNo," & _
            "EffectiveDate,CreateDate,ClosedDate," & _
            "MaturityDate,MaturedOn,DepositAmount," & _
            "RateOfInterest,Introduced," & _
            "LedgerNo,FolioNo,NomineeID," & _
            "LastPrintId,LastIntDate,DepositType,AccGroupId)"
        
        SqlStr = SqlStr & " VALUES (" & _
            AccID & "," & _
            Rst("CustomerID") & "," & _
            AddQuotes(AccNum, True) & ",'" & CertNo & "'," & _
            IIf(IsNull(Rst("EffectiveDate")), FormatDateField(Rst("CreateDate")), FormatDateField(Rst("EffectiveDate"))) & "," & _
            FormatDateField(Rst("CreateDate")) & "," & _
            FormatDateField(Rst("ClosedDate")) & " ," & _
            FormatDateField(Rst("MaturityDate")) & " ," & _
            FormatDateField(Rst("MaturedOn")) & " ," & DepAmount & "," & _
            Rst("RateOfInterest") & "," & _
            IntroId & "," & _
            AddQuotes(FormatField(Rst("LedgerNo")), True) & "," & _
            AddQuotes(FormatField(Rst("FolioNo")), True) & " ," & _
            "0 ," & _
            "1 ," & _
            FormatDateField(Rst("EffectiveDate")) & _
            ", " & m_DepositType & ", " & _
            "1 )"
            
        NewFDTrans.BeginTrans
        NewFDTrans.SQLStmt = SqlStr
        
        If Not NewFDTrans.SQLExecute Then
            NewFDTrans.RollBack
            AccID = AccID - 1
            'Now Check related Customer Is Missed
            NewFDTrans.SQLStmt = "SELECT CustomerID From NameTab Where CustomerID = " & Rst("CustomerID")
            If NewFDTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then GoTo NextAccount
            MsgBox "Unable to transafer the Fd MAster data"
            Exit Function
        End If
        NewFDTrans.CommitTrans

NextAccount:
    With frmMain
        .lblProgress = "Inserting the Fd Account Details"
        .prg.Value = Rst.AbsolutePosition
        .Refresh
    End With
    
    oldAccid = Rst("AccID"): oldDepositID = Rst("DepositID")
    
    Rst.MoveNext
Wend
Exit_line:
TransferFDMaster = True

Set Rst = Nothing
Set rstTemp = Nothing
Exit Function

Err_Line:
If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "Error In FDMaster " & Err.Description
        Err.Clear
        'Resume
    End If
    
End Function

'this function is used to transfer the
'FD MAster details form OLdb to new one
'and NewFDTrans has assigned to new database
Private Function TransferFDLoanMaster(OldFDTrans As clsOldUtils, NewFDTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim Sql_2 As String
Dim Rst As ADODB.Recordset
Dim rstTemp As ADODB.Recordset

'Dim DepositType As wis_DepositType

On Error GoTo Err_Line

'Now TransFer the Loan Details
'Get the account Having Details
SqlStr = "SELECT A.DepositID,A.AccID,LOan,CustomerID,MaturityDate," & _
    " TransDate,RateOfInterest,Amount,TransID,LedgerNo,FolioNo " & _
    " FROM FDTrans A, FDMAster B WHERE A.DepositId = B.DepositID " & _
    " AND A.AccID = B.AccID And TransId = (SELECT Min(TransID) " & _
        " From FDTrans C Where C.DepositID = B.DepositID " & _
        " AND C.AccID = B.AccID AND C.Loan = A.Loan AND Balance > 0 ) " & _
    " ORDER BY A.AccID, A.DepositID,TransID Asc,Loan Desc"

OldFDTrans.SQLStmt = SqlStr
If OldFDTrans.Fetch(Rst, adOpenForwardOnly) <= 0 Then GoTo ExitLine

Dim AccID As Long
Dim DepositID As Integer
Dim LoanId As Long
Dim LoanNum As String

'Get the Account Offset From the new dataBase
SqlStr = "SELECT Max(LoanID) FROM DepositLoanMASTER "
NewFDTrans.SQLStmt = SqlStr
If NewFDTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
    m_LoanOffSet = FormatField(rstTemp(0))

LoanId = m_LoanOffSet
AccID = m_AccOffSet


With frmMain
    .lblProgress = "Inserting the Fd LOan account Details"
    .prg.Min = 0
    .prg.Max = Rst.RecordCount + 1
    .prg.Value = 0
End With

While Not Rst.EOF
    AccID = AccID + 1
    If Rst("Loan") = False Then GoTo NextAccount
    AccID = AccID - 1
    LoanId = LoanId + 1
    LoanNum = Rst("ACCID")
    Debug.Assert LoanId <> 0
    
    NewFDTrans.BeginTrans
    
    '''' Here  we are inserting into the new table called DepositLoanMAster
    ''the above said table will be common for all type deposit(eg. FD,Rd,Pd,DL)

    'First insert the pledge details
    SqlStr = "Insert INTO PledgeDeposit (" & _
        "LoanID,AccID,DepositType,PledgeNum)" & _
        " VALUES (" & _
        LoanId & "," & _
        AccID & "," & _
        m_DepositType & "," & _
        " 1 )"
            
    NewFDTrans.SQLStmt = SqlStr
    If Not NewFDTrans.SQLExecute Then
        NewFDTrans.RollBack
        MsgBox "Unable to transafer the pigmy MAster data"
        Exit Function
    End If
    
    'Now insert the Loans master details
    SqlStr = "Insert INTO DepositLoanMASTER (" & _
        "CustomerID,LoanID, AccNum,DepositType," & _
        " LoanIssuedate,LoanDueDate,PledgeDescription, " & _
        "InterestRate,LoanAmount,LedgerNo,FolioNo ," & _
        " LastPrintId )"
    
    SqlStr = SqlStr & " VALUES (" & _
        Rst("CustomerID") & "," & LoanId & "," & _
        AddQuotes(LoanNum, True) & "," & _
        m_DepositType & "," & _
        FormatDateField(Rst("TransDate")) & "," & _
        FormatDateField(Rst("MaturityDate")) & "," & _
        "'" & Format(Rst("AccID"), "000") & "' ," & _
        Rst("RateOfInterest") & "," & _
        Rst("AMount") & "," & _
        "'" & Val(FormatField(Rst("LedgerNo"))) & "'," & _
        "'" & Val(FormatField(Rst("FolioNo"))) & "'," & _
        "1 )"
        
    NewFDTrans.SQLStmt = SqlStr
    If Not NewFDTrans.SQLExecute Then
        NewFDTrans.RollBack
        MsgBox "Unable to transafer the FD loan data"
        Exit Function
    End If
    
    SqlStr = "Update FDMaster Set Loanid = " & LoanId & " Where AccID = " & AccID
    NewFDTrans.SQLStmt = SqlStr
    If Not NewFDTrans.SQLExecute Then
        NewFDTrans.RollBack
        MsgBox "Unable to transafer the pigmy loan data"
        Exit Function
    End If
    
    NewFDTrans.CommitTrans

NextAccount:
    LoanNum = Rst("AccID")
    DepositID = Rst("DepositID")
    With frmMain
        .lblProgress = "Inserting the Fd LOan account Details"
        .prg.Max = Rst.AbsolutePosition
        .Refresh
    End With
    Rst.MoveNext
    
Wend

Set Rst = Nothing

ExitLine:
TransferFDLoanMaster = True

Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next

If Err Then MsgBox "Error In FD LoanMaster " & Err.Description
    
End Function

'this function is used to transfer the
'FD transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferFDTrans(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim SqlPayble As String

Dim IsIntTrans As Boolean
Dim IsPaybleTrans As Boolean

On Error GoTo Err_Line

Dim OldTransType As Integer
Dim NewTransType As Integer
Dim Balance As Currency
Dim PayBalance As Currency
Dim TransID As Long
Dim Rst As ADODB.Recordset
Dim AccID As Long
Dim Amount As Currency
Dim TransDate As Date
Dim oldDepositID As Long
Dim oldAccid As Long
Dim LastIntDate As Date

Dim RstAcc As Recordset

OldTrans.SQLStmt = SqlStr
'If OldTrans.Fetch(Rst, adOpenStatic) < 1 Then GoTo ExitLine



'Fetch the detials of FD Account
SqlStr = "SELECT * FROM FDTrans Where Loan = False " & _
    " ORDER BY AccID,DepositId,TransId"

OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(Rst, adOpenStatic) < 1 Then GoTo ExitLine

With frmMain
    .lblProgress = "Inserting the Fd Account Trans Details"
    .prg.Min = 0
    .prg.Max = Rst.RecordCount + 1
    .prg.Value = 0
    .Refresh
End With


TransID = 10000000
AccID = m_AccOffSet
SqlInt = "": SqlPayble = ""

While Not Rst.EOF
    IsIntTrans = False: IsPaybleTrans = False
    OldTransType = FormatField(Rst("TransType"))
    If OldTransType = 0 Then GoTo NextAccount
    If OldTransType = 2 Then IsIntTrans = True   'Interest paid
    If OldTransType = -2 Then IsIntTrans = True  'Access Interest received
    If OldTransType = -4 Or OldTransType = 4 Then IsPaybleTrans = True 'deposited to Interes PAyable
    If OldTransType = -5 Or OldTransType = 5 Then IsPaybleTrans = True 'Withdrawn from Interes PAyable
    
    'If the last record's Transaction id is greater or equal to present transid then
    'It means that the account no has been changed
    If oldDepositID <> Rst("DepositId") Or oldAccid <> Rst("AccId") Then
        oldAccid = Rst("ACCID")
        oldDepositID = FormatField(Rst("DepositId"))
        AccID = AccID + 1
        TransID = 0
        Balance = 0
        PayBalance = 0
    End If
    TransID = TransID + 1
    Amount = Rst("Amount")
    If Amount = 0 Then GoTo NextAccount
    'NewTransType = OldTransType / Abs(OldTransType)
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    TransDate = Rst("TransDate")

    If IsPaybleTrans Then
        If OldTransType = 4 Then NewTransType = wContraDeposit: PayBalance = PayBalance + Amount
        If OldTransType = 5 Then NewTransType = wWithDraw: PayBalance = PayBalance - Amount
        'The above transactions also effect the profit & loss
        If OldTransType = -5 Then NewTransType = wDeposit: PayBalance = PayBalance + Amount
        If OldTransType = -4 Then NewTransType = wContraWithDraw: PayBalance = PayBalance - Amount
        If PayBalance < 0 Then PayBalance = 0
        SqlPayble = "Insert INTO FDIntPayable ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlPayble = SqlPayble & "VALUES (" & _
            AccID & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            PayBalance & " ," & _
            NewTransType & ")"
        If OldTransType = 4 Then
            If NewTransType = wContraDeposit Then NewTransType = wContraWithDraw
            
            If NewTransType = wDeposit Then NewTransType = wWithDraw
            If NewTransType = wWithDraw Then NewTransType = wDeposit
            If NewTransType = wContraWithDraw Then NewTransType = wContraDeposit
                        
            SqlInt = "Insert INTO FDIntTrans ( " & _
                "AccID,TransID,TransDate," & _
                "Amount,Balance," & _
                "TransType )"
            SqlInt = SqlInt & "VALUES (" & _
                AccID & "," & _
                TransID & "," & _
                "#" & Rst("TransDate") & "#," & _
                Rst("Amount") & "," & _
                "0 ," & _
                NewTransType & " )"
            If Rst("Amount") = 0 Then SqlInt = ""
        End If
        
        'if oldTranstype = 5 then he has withdraw the amount from the
        'interest payable 'So consider the next transaction
        'If OldTransType = 5 Then Rst.MoveNext
        'changed code on 13/11/2003
        If OldTransType = 5 Then
            Rst.MoveNext
            OldTransType = FormatField(Rst("TransType"))
            If Abs(OldTransType) = 2 Then IsIntTrans = True
            'If Abs(OldTransType) <> 1 Then Rst.MovePrevious: GoTo NextAccount
        End If
    End If
    If IsIntTrans Then
        If OldTransType = 2 Then NewTransType = wWithDraw
        If OldTransType = -2 Then NewTransType = wDeposit
        NewTransType = IIf(OldTransType < 0, wDeposit, wWithDraw)
        SqlInt = "Insert INTO FDIntTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            AccID & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0  ," & _
            NewTransType & " )"
        Amount = Rst("Amount")
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        LastIntDate = Rst("TransDate")
        If Balance = FormatField(Rst("balance")) And _
            oldAccid = Rst("AccID") And oldDepositID = Rst("DepositID") Then Rst.MoveNext
        OldTransType = Rst("TransType")
        'After this transaction the transaction in the Table is contra
        'Therefore
        NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
        'If TransDate <> rst("TransDate") Then SqlInt = ""
    End If
    
    'If the transaction is payble then Need not do the '
    'transaction if depsoit accounts
    Amount = Rst("Amount")
    'NewTransType = (OldTransType / Abs(OldTransType))
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    SqlStr = "Insert INTO FDTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
    SqlStr = SqlStr & "VALUES (" & _
            AccID & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Amount & "," & _
            Rst("Balance") & "," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
            NewTransType & " )"
    
    
    If oldDepositID <> Rst("DepositId") Or oldAccid <> Rst("AccId") Then SqlStr = ""
    If Amount = 0 Then SqlStr = ""
    If TransDate <> Rst("TransDate") Then SqlStr = ""
    If Abs(OldTransType) <> 1 Then SqlStr = ""
    
    If Balance = FormatField(Rst("balance")) Then SqlStr = ""
    'This Transaction is the INterest Paid Where The Next Transaction
    If Rst("TransType") = 5 And Not IsPaybleTrans Then Rst.MovePrevious
        'If rst("TransType") = 5 Then
            'rst.MovePrevious
            'If It is interest transaction and
            'Balance has not changed in the Nexr Record
            'It means next to this (Current Position)record is
            'amount withdrawn from Interest payable  'Move back to the old position
            'rst.MovePrevious
            'TransID = TransID - 1
        'End If
    'End If
    
    NewTrans.BeginTrans
    If SqlStr <> "" Then
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the FD Trans data"
            NewTrans.RollBack
            Exit Function
        End If
    End If
    If SqlInt <> "" Then
        NewTrans.SQLStmt = SqlInt
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the FD Trans data"
            NewTrans.RollBack
            Exit Function
        End If
        SqlInt = ""
    End If
    If SqlPayble <> "" Then
        NewTrans.SQLStmt = SqlPayble
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the FD Trans data"
            NewTrans.RollBack
            Exit Function
        End If
        SqlPayble = ""
    End If

    NewTrans.CommitTrans
    
    'If the Record moved and next record is not on the same date
    'The go back to the previus position
    If TransDate <> Rst("TransDate") Then Rst.MovePrevious
    Balance = FormatField(Rst("Balance"))
    
NextAccount:
    With frmMain
        .lblProgress = "Inserting the Fd LOan Transaction Details"
        .prg.Value = Rst.AbsolutePosition
    End With
    Rst.MoveNext
Wend

Set Rst = Nothing


ExitLine:
TransferFDTrans = True
'Resume
Exit Function


Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "Error in FDTrans" & Err.Description
        Err.Clear
        'Resume
    End If
    
End Function

'this function is used to transfer the
'pigmy Loan transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferFDLoanTrans(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim IsIntTrans As Boolean
Dim rstTemp As Recordset
Dim SqlDep As String
Dim DepCount As Byte

On Error GoTo Err_Line

Dim OldTransType As Integer, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim Rst As ADODB.Recordset
Dim AccID As Long
Dim LoanId As Long
Dim Amount As Currency
Dim TransDate As Date
'Dim DepositType As wis_DepositType

Dim oldAccid As Long
Dim oldDepositID As Long

'DepositType = wisDeposit_FD
    
'Fetch the detials of FD Account
SqlStr = "SELECT * FROM FDTrans Where Loan = True " & _
    " ORDER BY AccID,DepositID,TransId"

OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo ExitLine

TransID = 10000
Balance = 0
LoanId = m_LoanOffSet
AccID = m_AccOffSet

With frmMain
    .lblProgress = "Inserting the Fd LOan Transaction Details"
    .prg.Min = 0
    .prg.Max = Rst.RecordCount + 1
    .prg.Value = 0
End With

While Not Rst.EOF
    SqlDep = ""
    IsIntTrans = False
    OldTransType = FormatField(Rst("TransType"))
    If OldTransType = 4 Or OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Or OldTransType = -4 Then IsIntTrans = True
    
    If TransID >= Rst("TransID") Then
        NewTrans.SQLStmt = "Select LoanID From DepositLoanMaster" & _
            " Where AccNum = '" & Rst("AccId") & "'" & _
            " And DepositType = " & m_DepositType
        If NewTrans.Fetch(rstTemp, adOpenDynamic) < 0 Then GoTo NextAccount
        If LoanId = rstTemp("LoanID") Then
            LoanId = LoanId
            rstTemp.MoveNext
            If DepCount Then rstTemp.Move DepCount
            If rstTemp.EOF Then rstTemp.MoveLast
            LoanId = rstTemp("LoanID")
            TransID = 0
            DepCount = DepCount + 1
            SqlDep = "Update DepositLoanMaster Set " & _
                " AccNum = '" & Rst("Accid") & "_" & DepCount & "' " & _
                " Where LoanID = " & LoanId & " And DepositType = " & m_DepositType
        Else
            LoanId = rstTemp("LoanID")
            TransID = 0
            DepCount = 0
        End If
        Balance = -1
    End If
    TransID = TransID + 1
    
    oldAccid = Rst("AccID"): oldDepositID = Rst("DepositID")
    TransDate = Rst("TransDate")
    
    NewTransType = IIf(OldTransType < 0, wWithDraw, wDeposit)
    If IsIntTrans Then
        If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithDraw  'interest Paid to the customer
        If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit  'Interest collected form the customer
        SqlInt = "Insert INTO DepositLoanIntTrans ( " & _
            "LoanId,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            LoanId & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0," & _
            NewTransType & " )"
        
        Rst.MoveNext
        Amount = Rst("Amount")
    End If
    
    OldTransType = FormatField(Rst("TransType"))
    NewTransType = (OldTransType / Abs(OldTransType))
    NewTransType = IIf(OldTransType > 0, wDeposit, wWithDraw)
    Debug.Assert LoanId <> 0
    
    
    NewTrans.BeginTrans
    If SqlDep <> "" Then
        NewTrans.SQLStmt = SqlDep
        SqlStr = ""
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the FD loan Trans "
            NewTrans.RollBack
            Exit Function
        End If
        SqlDep = ""
    End If
    
    ''INSERT INTO DEPOSITLOANTRANS
    
    SqlStr = "Insert INTO DepositLoanTrans ( " & _
            "LoanID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
    SqlStr = SqlStr & "VALUES (" & _
        LoanId & "," & _
        TransID & "," & _
        "#" & Rst("TransDate") & "#," & _
        Rst("Amount") & "," & _
        Rst("Balance") & " ," & _
        AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
        NewTransType & ")"
    If Rst("Amount") = 0 Then SqlStr = ""
    
    If SqlInt <> "" Then
        'Now Check whether the Deposit or accuont has changed after inserting the
        'Interest detalas Check whether the Date has changed after this insertion
        If oldAccid <> Rst("AccID") Or oldDepositID <> Rst("DepositID") Then
            SqlStr = ""
            Rst.MovePrevious
        ElseIf DateDiff("d", TransDate, Rst("TransDate")) Then
            SqlStr = ""
            Rst.MovePrevious
        End If
        NewTrans.SQLStmt = SqlInt
        SqlInt = ""
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the FD loan Trans "
            NewTrans.RollBack
            Exit Function
        End If
    End If
    If SqlStr <> "" Then
        NewTrans.SQLStmt = SqlStr
        SqlStr = ""
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the FD loan Trans "
            NewTrans.RollBack
            Exit Function
        End If
    End If
    
    NewTrans.CommitTrans
    Balance = FormatField(Rst("Balance"))

NextAccount:
    With frmMain
        .lblProgress = "Inserting the Fd Loan Transaction Details"
        .prg.Value = Rst.AbsolutePosition
        .Refresh
    End With
    Rst.MoveNext
    
Wend

ExitLine:
Set Rst = Nothing
TransferFDLoanTrans = True
    With frmMain
        .lblProgress = "Transferred the Fixed deposit transaction"
        .prg.Value = 0
        .Refresh
    End With

Exit Function


Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "Error in SBTrans" & Err.Description
        Err.Clear
        'Resume
    End If
    
End Function

'this function is used to transfer the
'set the voucher no fot the transaferred data
Private Function PutVoucherNumber(NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String



PutVoucherNumber = True
End Function
