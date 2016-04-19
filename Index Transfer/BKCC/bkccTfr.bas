Attribute VB_Name = "BKCCTransfer"
Option Explicit

Dim m_HeadBalance As Currency
Private Function CreateBKCCHeads(NewTrans As clsDBUtils) As Boolean
Dim ClsBank As clsBankAcc
Set ClsBank = New clsBankAcc


Dim BkccLoanId As Long
Dim BkccDepositId As Long
Dim BKCCRegIntId As Long
Dim BKCCPenalIntId As Long
Dim BKccDepIntId As Long

Dim HeadName As String
Dim Suffix As String

Dim FromDate As Date
FromDate = gFromDate '"3/31/03"

On Error GoTo ErrLine
'"BKCC"= LoadResString(gLangOffSet + 229)
HeadName = LoadResString(gLangOffSet + 229)

'aSIGN THE vARIABLE
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans

frmMain.lblProgress = "Transferring the BKCc Ledger"
frmMain.Refresh

Dim PrgVal As Integer
With frmMain.prg
    .Max = 365
    .Min = 0
    .Value = PrgVal
End With


Dim rstTrans As ADODB.Recordset
'''Variables required for the Bank accounts
Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency


Dim TotalRegInt As Currency
Dim TotalPenalInt As Currency
'Dim TotalDepInt As Currency
Dim TotalMiscAmount  As Currency

Dim TransDate As Date
Dim TransType As wisTransactionTypes

'Now Insert the loan Transction to the Acctrans  table
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    " TransType,TransDate From BKCCTrans " & _
    " Where Deposit = 0 ANd TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
    NewIndexTrans.BeginTrans
    'Create the Head Id In the Asset Side
    Suffix = " " & LoadResString(gLangOffSet + 58) 'Loan
    BkccLoanId = ClsBank.GetHeadIDCreated(HeadName & Suffix, parMemberLoan, m_HeadBalance, wis_BKCCLoan)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(BkccLoanId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(BkccLoanId, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
            frmMain.prg.Value = frmMain.prg.Value + 1
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        rstTrans.MoveNext
    Wend

    Call ClsBank.UpdateCashDeposits(BkccLoanId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(BkccLoanId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now Insert the transaction of interest fee deatils
NewTrans.SQLStmt = "Select sum(IntAmount) as TotalReg, " & _
    "Sum(PenalIntAmount) as TotalPenal, sum(MiscAmount) as TotalMisc," & _
    "TransType,TransDate From BKCCIntTrans " & _
    " Where Deposit = 0 ANd TransDate >= #" & FromDate & "# " & _
    "Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
    NewIndexTrans.BeginTrans
    'Regular Interest Head for Bkcc Loan
    Suffix = " " & LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 344) 'Loan Regular Interest
    BKCCRegIntId = ClsBank.GetHeadIDCreated(HeadName & Suffix, parMemLoanIntReceived, 0, wis_BKCCLoan)
    'Penal Interest Head for Bkcc Loan
    Suffix = " " & LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 345) 'Loan Regular Interest
    BKCCPenalIntId = ClsBank.GetHeadIDCreated(HeadName & Suffix, parMemLoanPenalInt, 0, wis_BKCCLoan)

    TransDate = rstTrans("Transdate")
    frmMain.prg.Value = 0
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(BKCCRegIntId, TotalRegInt, TransDate)
            Call ClsBank.UpdateCashDeposits(BKCCPenalIntId, TotalPenalInt, TransDate)
            Call ClsBank.UpdateCashDeposits(parBankIncome + 1, TotalMiscAmount, TransDate)
            
            Call ClsBank.UpdateCashWithDrawls(BKCCRegIntId, TotalWithdraw, TransDate)
            TotalRegInt = 0: TotalPenalInt = 0
            TotalMiscAmount = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
            frmMain.prg.Value = frmMain.prg.Value + 1
        End If
        
        TransType = FormatField(rstTrans("TransType"))
        
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalRegInt = TotalRegInt + FormatField(rstTrans("TotalReg"))
            TotalPenalInt = TotalPenalInt + FormatField(rstTrans("TotalPenal"))
            TotalMiscAmount = TotalMiscAmount + FormatField(rstTrans("TotalMisc"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalReg"))
        End If
        
        rstTrans.MoveNext
    Wend
    
    Call ClsBank.UpdateCashDeposits(BKCCRegIntId, TotalRegInt, TransDate)
    Call ClsBank.UpdateCashDeposits(BKCCPenalIntId, TotalPenalInt, TransDate)
    Call ClsBank.UpdateCashDeposits(parBankIncome + 1, TotalMiscAmount, TransDate)
    
    Call ClsBank.UpdateCashWithDrawls(BKCCRegIntId, TotalWithdraw, TransDate)
    TotalRegInt = 0: TotalPenalInt = 0
    TotalMiscAmount = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If


''''Now Insert the depoist Transction details of tot acctrans table
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    " TransType,TransDate From BKCCTrans " & _
    " Where Deposit <> 0 ANd TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Create the Head in the Laibility Side i.e. BKCC Deposit
    Suffix = " " & LoadResString(gLangOffSet + 43) 'Deposit
    BkccDepositId = ClsBank.GetHeadIDCreated(HeadName & Suffix, parMemberDeposit, 0, wis_BKCC)
    frmMain.prg.Value = 0
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(BkccDepositId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(BkccDepositId, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
            frmMain.prg.Value = frmMain.prg.Value + 1
        End If
        TransType = FormatField(rstTrans("TransType"))
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalAmount"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalAmount"))
        End If
        rstTrans.MoveNext
    Wend

    Call ClsBank.UpdateCashDeposits(BkccDepositId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(BkccDepositId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now Insert the Interest paid for the Deposit
NewTrans.SQLStmt = "Select sum(IntAmount) as TotalReg, " & _
    "Sum(PenalIntAmount) as TotalPenal, sum(MiscAmount) as TotalMisc," & _
    "TransType,TransDate From BKCCIntTrans " & _
    " Where Deposit <> 0 ANd TransDate >= #" & FromDate & "# " & _
    "Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
    NewIndexTrans.BeginTrans
    'Expense Id For the BKCC Deposit Interest paid
    Suffix = " " & LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 487) 'Deposit Interst Paid
    BKccDepIntId = ClsBank.GetHeadIDCreated(HeadName & Suffix, parMemDepIntPaid, wis_BKCC)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(BKccDepIntId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(BKccDepIntId, TotalWithdraw, TransDate)
            TotalDeposit = 0: TotalWithdraw = 0
            TransDate = rstTrans("Transdate")
            frmMain.prg.Value = frmMain.prg.Value + 1
        End If
        
        TransType = FormatField(rstTrans("TransType"))
        
        If TransType = wDeposit Or TransType = wContraDeposit Then
            TotalDeposit = TotalDeposit + FormatField(rstTrans("TotalReg"))
        Else
            TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalReg"))
        End If
        
        rstTrans.MoveNext
    Wend
    
    Call ClsBank.UpdateCashDeposits(BKccDepIntId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(BKccDepIntId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    NewIndexTrans.SQLStmt = "UPDATe opBalance Set opdate = #4/1/2003#"
    NewIndexTrans.SQLExecute
    NewIndexTrans.CommitTrans
End If

CreateBKCCHeads = True

ExitLine:
    Set ClsBank = Nothing
    Set NewIndexTrans = Nothing
    Set rstTrans = Nothing

ErrLine:
    
If Err.Number = 380 Then
    frmMain.prg.Max = PrgVal * 1.5
    Resume Next
ElseIf Err.Number Then
    MsgBox "Error In Chcking data Base"
    'GoTo Exit_line
    'Resume
End If

End Function

'This Bas file is made to transfer the Index 2000
'data base to the this existing loansdatabase

Function TransferBKCC(OldDBName As String, NewDBName As String) As Boolean

Dim OldTrans As New clsOldUtils
Dim NewTrans As New clsDBUtils


If Not OldTrans.OpenDB(OldDBName, OldPwd) Then
    MsgBox " No Index Db"
    Exit Function
End If

If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    MsgBox " No Index Db"
    Exit Function
End If
Set NewIndexTrans = Nothing
Set NewIndexTrans = NewTrans

'Dim dbIndex As clsDBUtils


If Not gOnlyLedgerHeads Then
    If Not BKCCMasterTransfer(OldTrans, NewTrans) Then Exit Function
    If Not BKCCTransTransfer(OldTrans, NewTrans) Then Exit Function
    '''Now Create transaction Heads In New Database
End If
If Not CreateBKCCHeads(NewTrans) Then Exit Function
Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
    
'If Not BKCCDetailsTransfer(NewTrans) Then Exit Function

TransferBKCC = True

Set NewIndexTrans = Nothing


End Function


Private Function BKCCMasterTransfer(oldLoanTrans As clsOldUtils, NewLoanTrans As clsDBUtils) As Boolean

Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim AccNum As String
Dim BankID As Long
Dim RecNo As Integer

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Database of loans

With frmMain
    .lblProgress = "Fetching the data Of BKCC Loan Details"
    '.prg.Max
    .prg.Value = 1
    .Refresh
End With

SqlStr = "SELECT A.*,B.CustomerID FROM LoanMaster A, MMMaster B " & _
    " WHERE B.AccID=A.MemberID AND LoanId IN (SELECT LoanID " & _
        " From LOanMaster Where SchemeID In " & _
            "(SELECT SchemeID FRom LoanTypes Where BKCC = TRUE ))" & _
    " ORDER BY LoanID"

oldLoanTrans.SQLStmt = SqlStr
If oldLoanTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then
    MsgBox "There are no BKCC loans to transfer"
    Exit Function
End If

'In the Loan Dtaabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id

With frmMain
    .lblProgress = "Transferring the data of BKCC Loan Master Details"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
End With

On Error GoTo Err_Line

Dim SchemeID As Integer
Dim FarmerType As Integer

Dim rstTemp As Recordset
oldLoanTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#" & _
            " Order By obDate Desc"
If oldLoanTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Set rstTemp = Nothing

SchemeID = 0 'FormatField(RstIndex("SchemeID"))
FarmerType = 0 '1

NewLoanTrans.BeginTrans
While Not RstIndex.EOF
    
    If SchemeID <> FormatField(RstIndex("SchemeID")) Then
        'Get the Head Balance
        If Not rstTemp Is Nothing Then
            rstTemp.MoveFirst
            Do
                If rstTemp.EOF Then Exit Do
                If rstTemp("Module") = SchemeID Then _
                    m_HeadBalance = m_HeadBalance + FormatField(rstTemp("ObAmount")): Exit Do
                rstTemp.MoveNext
            Loop
        End If
        'm_HeadBalance = m_HeadBalance + FormatField(rstTemp("ObAmount"))
        Set rstTemp = Nothing
        SchemeID = FormatField(RstIndex("SchemeID"))
        FarmerType = FarmerType + 1
        If FarmerType > 3 Then FarmerType = 1
    End If
    
    'first get the mem Id of the old db
     MemID = FormatField(RstIndex("MemberId"))
    'now get the customer Of this Member
    CustomerId = 0
    CustomerId = FormatField(RstIndex("CustomerId"))
    AccNum = FormatField(RstIndex("LoanAccNo"))
    
    If Trim(AccNum) = "" Then _
        AccNum = FormatField(RstIndex("SchemeID")) & "_" & FormatField(RstIndex("LoanId"))
    
    SqlStr = "INSERT INTO BKCCMaster (" & _
            " LoanId,MemID,CustomerID," & _
            " AccNum,FarmerType,Issuedate," & _
            " SanctionAmount," & _
            " Intrate,PenalIntRate,DepIntRate, " & _
            " Guarantor1,Guarantor2,LoanClosed, Remarks) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LoanId") & "," & _
            MemID & ", " & CustomerId & "," & _
            AddQuotes(AccNum, True) & "," & FarmerType & "," & _
            FormatDateField(RstIndex("IssueDate")) & "," & _
            FormatField(RstIndex("LoanAmt")) & "," & _
            FormatField(RstIndex("InterestRate")) & "," & _
            FormatField(RstIndex("PenalInterestrate")) & ", 10 ," & _
            FormatField(RstIndex("GuarantorId1")) & "," & _
            FormatField(RstIndex("GuarantorId2")) & "," & _
            FormatField(RstIndex("LoanClosed")) & "," & _
            AddQuotes(FormatField(RstIndex("Remarks")), True) & ")"

    NewLoanTrans.SQLStmt = SqlStr
    If Not NewLoanTrans.SQLExecute Then
        NewLoanTrans.RollBack
        Exit Function
    End If
    
NextAccount:
    RstIndex.MoveNext

    With frmMain
        .lblProgress = "Transferring the data of BKCC Loan Master Details"
        .prg.Value = RecNo
    End With
    RecNo = RecNo + 1
    
Wend

NewLoanTrans.CommitTrans


Exit_line:

BKCCMasterTransfer = True
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
    
End If

End Function
Private Function BKCCTransTransfer(oldLoanTrans As clsOldUtils, NewLoanTrans As clsDBUtils) As Boolean
Dim RstIndex As Recordset
Dim rstTemp As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim BankID As Long
Dim TransID As Long
Dim TransType As Integer
Dim Particualrs As String
Dim ItIsIntTrans As Boolean
Dim LoanId As Long
Dim RegInt As Double
Dim PenalInt As Double
Dim InstType As Integer
Dim Amount As Currency
Dim InstAmount As Currency
Dim InstNo As Integer
Dim InstBalance As Currency
Dim DepTrans As Boolean




'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM BKCCTrans ORDER By LoanID, TransID"
oldLoanTrans.SQLStmt = SqlStr
If oldLoanTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then GoTo Exit_line

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id
    
Particualrs = "Penal Interest"  ' Extarcted from Data BAse Differnet for Kannada
Particualrs = IIf(gLangOffSet, "·Ð®Ð·Ð ½¯ç", "Penal interest") ' LoadResString(345)

Dim InTrans As Boolean

DepTrans = False
With frmMain
    .lblProgress = "Transferring the data of BKCC Loan Transaction"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
End With

On Error GoTo Err_Line

While Not RstIndex.EOF
    
    ItIsIntTrans = False
    TransType = RstIndex("TransType")
    If LoanId <> RstIndex("LoanID") Then
'        Debug.Assert RstIndex("loanID") <> 566
        TransID = 0: DepTrans = False
        LoanId = RstIndex("LoanID")
        NewLoanTrans.SQLStmt = "SELECT CustomerId From BKCCMaster " & _
            " WHere LoanID = " & LoanId
        If NewLoanTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then
            RstIndex.Find "LoanId > " & LoanId
            GoTo NextAccount
        End If
    End If
    'Begin the transaction
    NewLoanTrans.BeginTrans
    InTrans = True
    TransID = TransID + 1
    If TransType = -2 Or TransType = 2 Then
        ItIsIntTrans = True
        RegInt = FormatField(RstIndex("Amount"))
        If InStr(1, Trim$(FormatField(RstIndex("Particulars"))), Particualrs, vbTextCompare) Then
            PenalInt = FormatField(RstIndex("Amount"))
            RegInt = 0
            
            RstIndex.MoveNext
            If LoanId <> RstIndex("LoanID") Then GoTo NextAccount
            RegInt = FormatField(RstIndex("Amount"))
            TransType = RstIndex("TransType")
            'if Next transaction is Payment or Receipt
            If TransType = 1 Or TransType = -1 Then
                RegInt = 0
                RstIndex.MovePrevious
                TransType = RstIndex("TransType")
            End If
        End If
        
        If DepTrans Then
            'If TransType > 0 Then TransType = wWithDraw
            TransType = IIf(TransType > 0, wWithDraw, wDeposit)
            Debug.Assert TransType = wWithDraw
        Else
            TransType = IIf(TransType < 0, wDeposit, wWithDraw)
        End If
        
        SqlStr = "INSERT INTO BKccIntTrans (" & _
                " LoanId,TransDate," & _
                " TransID,TransType," & _
                " IntAmount,PenalIntAmount, " & _
                " IntBalance,Deposit,Particulars)"
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("LoanId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & TransType & "," & _
                RegInt & "," & _
                PenalInt & "," & _
                " 0, " & _
                DepTrans & "," & _
                AddQuotes(FormatField(RstIndex("Particulars")), True) & ")"
        If RegInt Or PenalInt Then
            NewLoanTrans.SQLStmt = SqlStr
            If Not NewLoanTrans.SQLExecute Then
                NewLoanTrans.RollBack
                InTrans = False
                Exit Function
            End If
            NewLoanTrans.SQLStmt = "Update BKCCMaster Set LastIntDate = " & _
                "#" & RstIndex("TransDate") & "# WHERE LoanID = " & RstIndex("LoanId")
            If Not NewLoanTrans.SQLExecute Then
                NewLoanTrans.RollBack
                InTrans = False
                Exit Function
            End If
            RegInt = 0: PenalInt = 0
        End If
        RstIndex.MoveNext
        
        If RstIndex.EOF Then GoTo NextAccount
        If LoanId <> RstIndex("LoanID") Then GoTo NextAccount
        
    End If
    
    TransType = RstIndex("TransType")
    If RstIndex("Balance") < 0 Then DepTrans = True
    If RstIndex("Balance") > 0 Then DepTrans = False
    
    
    If Not (TransType = -1 Or TransType = 1) Then
        If TransType = 7 Then
            DepTrans = True
            TransType = wDeposit
        ElseIf TransType = -7 Then
            DepTrans = True
            TransType = wWithDraw
        Else
            GoTo NextRecord
        End If
    Else
        TransType = IIf(TransType > 0, wDeposit, wWithDraw)
    End If
    
    Amount = FormatField(RstIndex("Amount"))
    
    SqlStr = "INSERT INTO BKCCTrans (" & _
            " LoanId,TransDate," & _
            " TransID,TransType," & _
            " Amount, Balance, Deposit,Particulars)"
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            "#" & RstIndex("TransDate") & "#," & _
            TransID & "," & TransType & "," & _
            Amount & "," & _
            FormatField(RstIndex("Balance")) & "," & _
            DepTrans & ", " & _
            AddQuotes(FormatField(RstIndex("Particulars")), True) & ")"
            If TransID = 1 Then TransID = 2
    
    NewLoanTrans.SQLStmt = SqlStr
    If Not NewLoanTrans.SQLExecute Then
        NewLoanTrans.RollBack
        InTrans = False
        Exit Function
    End If
    
NextRecord:
    RstIndex.MoveNext

NextAccount:
    If InTrans Then NewLoanTrans.CommitTrans
    InTrans = False
    If RstIndex.EOF Then GoTo Exit_line
    With frmMain
        .lblProgress = "Transferring the data of BKCC Loan Master Details"
        .prg.Value = RstIndex.AbsolutePosition
    End With
    
Wend


Exit_line:
Debug.Print Now
BKCCTransTransfer = True
With frmMain
    .lblProgress = "Transferred the BKCC transaction"
    .prg.Value = 0
    .Refresh
End With

Exit Function

Err_Line:
If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Trans"
    Resume
    Err.Clear
End If

End Function


