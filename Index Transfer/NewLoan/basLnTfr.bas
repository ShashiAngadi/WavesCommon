Attribute VB_Name = "LoanTransfer"
Option Explicit



Private Function CreateLoanAccountHeads(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean

Dim RstLoan As ADODB.Recordset
Dim FromDate As Date
Dim PrgVal As Long


FromDate = gFromDate '"3/31/2003"


'Get the Head Balance
Dim HeadBalance As Currency
Dim rstTemp As Recordset

OldTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#"
If OldTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Set rstTemp = Nothing

NewTrans.SQLStmt = "SELECT SchemeID, SchemeName " & _
            " From LoanScheme Order By SchemeID "

If NewTrans.Fetch(RstLoan, adOpenDynamic) < 1 Then Exit Function 'GoTo ExitLine
 
Dim HeadID As Long
Dim RegIntHeadID As Long
Dim PenalIntHeadId As Long

Dim LoanPayableId As Long

Dim rstTrans As ADODB.Recordset
Dim ClsBank  As clsBankAcc

Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency
Dim TotalRegInt As Currency
Dim TotalPenalInt As Currency
Dim TotalMiscAmount As Currency


Dim TransDate As Date
Dim TransType As wisTransactionTypes

Set ClsBank = New clsBankAcc
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans

Dim HeadName As String
Dim SchemeName As String
Dim SchemeID As String

frmMain.lblProgress = "Tranferring Loan ledger"
With frmMain.prg
    .Min = 0
    .Max = 100
    .Value = 0
End With

frmMain.lblProgress = "Creating the Loan Account heads"

'Dim AccType As wisModules
Dim AccType As Long
AccType = wis_Loans

While Not RstLoan.EOF
    
    SchemeName = FormatField(RstLoan("SchemeName"))
    SchemeID = FormatField(RstLoan("SchemeID"))
    AccType = wis_Loans + Val(SchemeID)
        
    'Insert the transction details to the acctrans table
     NewTrans.SQLStmt = "SELECT Sum(Amount) As TotalAmount, " & _
                        "TransDate,TransType From LoanTrans " & _
                        "Where LoanID In (Select LoanID  From LoanMaster " & _
                            "Where SchemeID = " & SchemeID & ") " & _
                        "And Transdate >= #" & gFromDate & "# " & _
                        "Group BY TransDate,TransType " & _
                        "Order BY TransDate,TransType"
    
    If NewTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        'Get the Head Balance
        'Dim HeadBalance As Currency
        'Dim rstTemp As Recordset
        
        If Not rstTemp Is Nothing Then
            rstTemp.MoveFirst
            Do
                If rstTemp.EOF Then Exit Do
                If rstTemp("Module") = SchemeID Then _
                    HeadBalance = FormatField(rstTemp("ObAmount")): Exit Do
                rstTemp.MoveNext
            Loop
        End If
        NewIndexTrans.BeginTrans
        'Create Loan Head
        HeadName = SchemeName
        HeadID = ClsBank.GetHeadIDCreated(HeadName, parMemberLoan, HeadBalance, AccType)
        TotalDeposit = 0: TotalWithdraw = 0

        TransDate = rstTrans("Transdate")
        While Not rstTrans.EOF
            If rstTrans("Transdate") <> TransDate Then
                Call ClsBank.UpdateCashDeposits(HeadID, TotalDeposit, TransDate)
                Call ClsBank.UpdateCashWithDrawls(HeadID, TotalWithdraw, TransDate)
                TotalDeposit = 0
                TotalWithdraw = 0
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
    
    'Insert the Interest details to the AccTrans table
    NewTrans.SQLStmt = "SELECT Sum(IntAmount) As TotalReg, " & _
                "Sum(PenalIntAmount) As TotalPenal, Sum(MiscAmount) As TotalMisc, " & _
                "TransDate,TransType From LoanIntTrans " & _
                "Where LoanID In (Select LoanID  From LoanMaster " & _
                    "Where SchemeID = " & SchemeID & ") " & _
                " And TransDate > #" & gFromDate & "#" & _
                "Group BY TransDate,TransType " & _
                "Order BY TransDate,TransType"

    If NewTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        
        NewIndexTrans.BeginTrans
        'Create Loan interest received head in income
        HeadName = SchemeName & " " & LoadResString(gLangOffSet + 344)   'Regular Interest
        RegIntHeadID = ClsBank.GetHeadIDCreated(HeadName, parMemLoanIntReceived, 0, AccType)
        HeadName = SchemeName & " " & LoadResString(gLangOffSet + 345)   'Penal Interest
        PenalIntHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemLoanPenalInt, 0, AccType)
        TransDate = rstTrans("Transdate")
        
        While Not rstTrans.EOF
            If rstTrans("Transdate") <> TransDate Then
                Call ClsBank.UpdateCashDeposits(RegIntHeadID, TotalRegInt, TransDate)
                Call ClsBank.UpdateCashDeposits(PenalIntHeadId, TotalPenalInt, TransDate)
                Call ClsBank.UpdateCashDeposits(parIncome + 1, TotalMiscAmount, TransDate)
                
                Call ClsBank.UpdateCashWithDrawls(RegIntHeadID, TotalWithdraw, TransDate)
                TotalRegInt = 0: TotalPenalInt = 0
                TotalMiscAmount = 0: TotalWithdraw = 0
                TransDate = rstTrans("Transdate")
            End If
            
            TransType = FormatField(rstTrans("TransType"))
            
            If TransType = wDeposit Or TransType = wContraDeposit Then
                TotalRegInt = TotalRegInt + FormatField(rstTrans("TotalReg"))
                TotalPenalInt = TotalPenalInt + FormatField(rstTrans("TotalPenal"))
                TotalMiscAmount = TotalMiscAmount + FormatField(rstTrans("TotalMisc"))
            Else
                TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalReg"))
                TotalWithdraw = TotalWithdraw + FormatField(rstTrans("TotalPenal"))
            End If
            rstTrans.MoveNext
        Wend
        Call ClsBank.UpdateCashDeposits(RegIntHeadID, TotalRegInt, TransDate)
        Call ClsBank.UpdateCashDeposits(PenalIntHeadId, TotalPenalInt, TransDate)
        Call ClsBank.UpdateCashDeposits(parIncome + 1, TotalMiscAmount, TransDate)
        
        Call ClsBank.UpdateCashWithDrawls(RegIntHeadID, TotalWithdraw, TransDate)
        TotalRegInt = 0: TotalPenalInt = 0
        TotalMiscAmount = 0: TotalWithdraw = 0
        NewIndexTrans.CommitTrans
    End If
NextLoan:
    
    RstLoan.MoveNext
Wend


CreateLoanAccountHeads = True
ReDim m_AccHeadid(0)
ReDim m_Amount(0)
ReDim m_AmountType(0)


ExitLine:
    Set rstTrans = Nothing
    Set RstLoan = Nothing
    Set NewIndexTrans = Nothing
    Set ClsBank = Nothing
    Exit Function
    
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

Function TransferLoan(OldDBName As String, NewDBName As String) As Boolean

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

If Not gOnlyLedgerHeads Then
    If Not SchemeTransfer(OldTrans, NewTrans) Then Exit Function
    If Not LoanMasterTransfer(OldTrans, NewTrans) Then Exit Function
    Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
        
    If Not LoanTransTransfer(OldTrans, NewTrans) Then Exit Function
End If
If Not CreateLoanAccountHeads(OldTrans, NewTrans) Then Exit Function
Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
        
TransferLoan = True

End Function



Private Function LoanMasterTransfer(oldLoanTrans As clsOldUtils, NewLoanTrans As clsDBUtils) As Boolean
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim AccNum As String
Dim BankID As Long
Dim InstMode As Integer
Dim InstAmount As Currency
Dim NoOfINstall As Integer
Dim RecNo As Integer

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Database of loans

With frmMain
    .lblProgress = "Transferring Loan details"
    .prg.Value = 0
    .Refresh
End With




SqlStr = "SELECT A.*,B.CustomerID FROM LoanMaster A, MMMaster B " & _
    " WHERE B.AccID=A.MemberID AND " & _
        " A.SchemeID In (SELECT SchemeID FRom LoanTypes " & _
                " Where BKCC = False OR BKCC is NULL)" & _
    " ORDER BY SchemeID,LoanID"
oldLoanTrans.SQLStmt = SqlStr
If oldLoanTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then
    MsgBox "There are no loan Types and Loan accounts"
    Exit Function
End If
    
'Set RstIndex = oldLoanTrans.rst.Clone
With frmMain
    .lblProgress = "Transferring Loan details"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
End With

'RstIndex.Find "MemberId = 1500"

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id

NewLoanTrans.BeginTrans
While Not RstIndex.EOF
    
    'first get the mem Id of the old db
     MemID = FormatField(RstIndex("MemberId"))
     
    'now get the customer Of this Member
    CustomerId = 0
    InstMode = 0: NoOfINstall = 0: InstAmount = 0
    InstMode = FormatField(RstIndex("InstalmentMode"))
    'Set oldLoanTrans.Rst = Nothing
    'SqlStr = "SELECT CustomerID from MMMaster where AccID = " & MemId
    'oldLoanTrans.SqlStmt = SqlStr
    'If oldLoanTrans.SQLFetch > 0 Then
    CustomerId = FormatField(RstIndex("CustomerId")) 'FormatField(oldLoanTrans.Rst(0))
    'Else
        'GoTo NextAccount
    'End If
    'RstIndex.AbsolutePosition = RecNo
    AccNum = FormatField(RstIndex("LoanAccNo"))
    If Trim(AccNum) = "" Then _
        AccNum = FormatField(RstIndex("SchemeID")) & "_" & FormatField(RstIndex("LoanId"))
    On Error Resume Next
    
    If InstMode > 0 Then
        If FormatField(RstIndex("InstalmentAmt")) < 10 Then
            
        Else
            InstAmount = FormatField(RstIndex("InstalmentAmt"))
            If InstAmount = 0 Then
                NoOfINstall = 0
            Else
                NoOfINstall = FormatField(RstIndex("LoanAmt")) / InstAmount
            End If
            If NoOfINstall > 2000 Then NoOfINstall = 0
            If NoOfINstall = 1 Then NoOfINstall = 0: InstAmount = 0: InstMode = 0
        End If
    End If
    On Error GoTo Err_Line
    SqlStr = "INSERT INTO LoanMaster (" & _
            " LoanId,SchemeID, MemID,CustomerID," & _
            " AccNUm,Issuedate,LoanDueDate," & _
            " PledgeItem,Pledgevalue,LoanAmount," & _
            " InstMode,InstAmount, NooFInstall, " & _
            " EMI, Intrate,PenalIntRate, " & _
            " Guarantor1,Guarantor2,LoanClosed, Remarks,AccGroupId) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            RstIndex("SchemeId") & "," & _
            MemID & ", " & CustomerId & "," & _
            AddQuotes(AccNum, True) & "," & _
            "#" & RstIndex("IssueDate") & "#," & _
            "#" & RstIndex("LoanDueDate") & "#," & _
            AddQuotes(FormatField(RstIndex("PledgeDescription")), True) & "," & _
            FormatField(RstIndex("PledgeValue")) & "," & _
            FormatField(RstIndex("LoanAmt")) & "," & _
            InstMode & ", " & InstAmount & "," & _
            NoOfINstall & ", False, " & _
            FormatField(RstIndex("InterestRate")) & "," & _
            FormatField(RstIndex("PenalInterestrate")) & "," & _
            FormatField(RstIndex("GuarantorId1")) & "," & _
            FormatField(RstIndex("GuarantorId2")) & "," & _
            FormatField(RstIndex("LoanClosed")) & "," & _
            AddQuotes(FormatField(RstIndex("Remarks")), True) & _
            " , 1 )"

    NewLoanTrans.SQLStmt = SqlStr
    If Not NewLoanTrans.SQLExecute Then    'THere ARE MORE FIELDS
        NewLoanTrans.RollBack              'IN lOANS THAN iNDEX
        Exit Function
    End If
    'If loan has got the installment then insert the installment details
    If InstMode > 0 Then
        'so we have consder daily loan installment in the new loans
        'and that is not the case with index 2000 so increase insttype by one
        NewLoanTrans.CommitTrans
        If NoOfINstall > 255 Then
            MsgBox "Improper information of the instalament Of memberid " & MemID, vbInformation
            NoOfINstall = 200
        End If
        InstMode = InstMode + 1
        If Not SaveInstallmentDetails(NewLoanTrans, RstIndex("LoanID"), InstMode, NoOfINstall, FormatField(RstIndex("LoanAmt")), _
                     InstAmount, FormatField(RstIndex("IssueDate"))) Then
            MsgBox "Unable to save the installment details of the " & RstIndex("LoanID")
            Exit Function
        End If
        NewLoanTrans.BeginTrans
    End If
NextAccount:
    With frmMain
        .lblProgress = "Transferring Loan Master details"
        .prg.Value = RecNo
    End With
    RecNo = RecNo + 1
    RstIndex.MoveNext
    
Wend

NewLoanTrans.CommitTrans


LoanMasterTransfer = True

Exit_line:

LoanMasterTransfer = True
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If
'Resume
End Function



Private Function SaveInstallmentDetails(NewLoanTrans As clsDBUtils, LoanId As Long, InstMode As Integer, NoOfInst As Integer, LoanAmount As Currency, InstAmount As Currency, IssueIndianDate As String) As Boolean

Dim InstNo As Integer
Dim lpCount As Integer
Dim SqlStr As String
Dim Rst As Recordset
Dim NextDate As Date
Dim FortNight As Boolean
Dim TotalInstAmount As Currency

NextDate = FormatDate(IssueIndianDate)
NewLoanTrans.BeginTrans
InstNo = 1
Do
     If InstNo > NoOfInst Then Exit Do
     If TotalInstAmount >= LoanAmount Then Exit Do
     
     'Get The Next INstallment date
     If InstMode = Inst_Daily Then NextDate = DateAdd("d", 1, NextDate)
     If InstMode = Inst_Weekly Then NextDate = DateAdd("WW", 1, NextDate)
     If InstMode = Inst_FortNightly Then
         If FortNight Then
             FortNight = False
             NextDate = DateAdd("d", 15, NextDate)
         Else
             FortNight = True
             NextDate = DateAdd("d", -15, NextDate)
             NextDate = DateAdd("m", 1, NextDate)
         End If
     End If
     If InstMode = Inst_Monthly Then NextDate = DateAdd("M", 1, NextDate)
     If InstMode = Inst_BiMonthly Then NextDate = DateAdd("m", 2, NextDate)
     If InstMode = Inst_Quartery Then NextDate = DateAdd("q", 1, NextDate)
     If InstMode = Inst_HalfYearly Then
         If FortNight Then
             FortNight = False
             NextDate = DateAdd("M", 6, NextDate)
         Else
             FortNight = True
             NextDate = DateAdd("M", -6, NextDate)
             NextDate = DateAdd("YYYY", 1, NextDate)
         End If
     End If
     If InstMode = Inst_Yearly Then NextDate = DateAdd("YYYY", 1, NextDate)
     
     'WRITE Into the databsae
     SqlStr = "INSERT INTO LoanInst (LoanID,InstNo," & _
            " InstDate,InstAmount,InstBalance )" & _
         " Values ( " & _
         LoanId & "," & _
         InstNo & "," & _
         " #" & NextDate & "#," & _
         InstAmount & "," & _
         InstAmount & " ) "
    NewLoanTrans.SQLStmt = SqlStr
    If Not NewLoanTrans.SQLExecute Then
       NewLoanTrans.RollBack
       Exit Function
    End If
     TotalInstAmount = TotalInstAmount + InstAmount
     InstNo = InstNo + 1
     
Loop
NewLoanTrans.CommitTrans

SaveInstallmentDetails = True


End Function


Private Function LoanTransTransfer(oldLoanTrans As clsOldUtils, NewLoanTrans As clsDBUtils) As Boolean
'Dim
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
Dim RstInst As Recordset
Dim Amount As Currency
Dim InstAmount As Currency
Dim InstNo As Integer
Dim InstBalance As Currency

Dim BKCC As Boolean

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM LoanTrans ORDER By LoanID, TransID"
oldLoanTrans.SQLStmt = SqlStr

If oldLoanTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then GoTo ExitLine
    
'Set RstIndex = oldLoanTrans.Rst.Clone

'In the Loan Database of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id
    
Particualrs = "Penal interest"  ' Extarcted from Data BAse Differnet for Kannada
Particualrs = IIf(gLangOffSet, "·Ð®Ð·Ð ½¯ç", "Penal interest") ' LoadResString(345)

Dim InTrans As Boolean
With frmMain
    .lblProgress = "Transferring Loan Transaction"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
End With

TransactionTransfer:

While Not RstIndex.EOF
    ItIsIntTrans = False
    TransType = RstIndex("TransType")
    If LoanId <> RstIndex("LoanID") Then
        Set RstInst = Nothing
        TransID = 0
        LoanId = RstIndex("LoanID")
   
        'Get the installment type
        NewLoanTrans.SQLStmt = "SELECT InstMode From LoanMaster Where LoanID = " & LoanId
        If NewLoanTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then
            'There is no information about this loan in the
            'Loan Master 'we will Drop this loan's transaction
            'and move to the next loanid
            
            MsgBox "Insufficient information of Instalment LoanID = " & LoanId
            
            'Get the Infroamtion Of the LAst Line
            oldLoanTrans.SQLStmt = "Select * from LoanMaster Where LoanID = " & LoanId
            If oldLoanTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
                MsgBox "Loan Details of Member No = " & rstTemp("MemberID") & " not transferred", vbInformation, wis_MESSAGE_TITLE
            End If
            Set rstTemp = Nothing
            'Search For the Next Loan
            
            RstIndex.Find "LoanID > " & LoanId
            'mOve to the last record of skipping loan
            'Check whether this recordset has moved to the last record
            If RstIndex.EOF Then GoTo Exit_line
            RstIndex.MovePrevious
            GoTo NextRecord
            'Exit Function
            InstType = 0
        Else
            InstType = FormatField(rstTemp("InstMode"))
        End If
    End If
    
    If InstType > 0 Then
        Set RstInst = Nothing
        SqlStr = "SELECT * FROM LoanInst Where LoanID = " & LoanId & _
            " AND InstBalance > 0 ORDER BY InstDate"
        NewLoanTrans.SQLStmt = SqlStr
        If NewLoanTrans.Fetch(RstInst, adOpenDynamic) < 0 Then
            MsgBox "Error in loan installment of " & LoanId
            Exit Function
        End If
        'Set RstInst = NewLOanTrans.Rst.Clone
    End If
    
    'Begin the transaction
    NewLoanTrans.BeginTrans
    InTrans = True
    TransID = TransID + 1
    
    If TransType = -2 Or TransType = 2 Then
        ItIsIntTrans = True
        RegInt = FormatField(RstIndex("Amount"))
        If InStr(1, Trim$(FormatField(RstIndex("Particulars"))), Particualrs, vbTextCompare) Then
            RegInt = 0
            PenalInt = FormatField(RstIndex("Amount"))
            RstIndex.MoveNext
            If RstIndex.EOF Then GoTo NextAccount
            If LoanId <> RstIndex("LoanID") Then GoTo NextAccount
            RegInt = FormatField(RstIndex("Amount"))
        End If
        'If TransType = -2 Then TransType = wDeposit
        TransType = IIf(TransType < 0, wDeposit, wWithDraw)
        SqlStr = "INSERT INTO LoanIntTrans (" & _
                " LoanId,TransDate," & _
                " TransID,TransType," & _
                " IntAmount,PenalIntAmount, IntBalance)"
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("LoanId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & TransType & "," & _
                RegInt & "," & _
                PenalInt & "," & _
                FormatField(RstIndex("Balance")) & _
                ")"
        NewLoanTrans.SQLStmt = SqlStr
        If RegInt Or PenalInt Then
            If Not NewLoanTrans.SQLExecute Then
                NewLoanTrans.RollBack
                InTrans = False
                Exit Function
            End If
        End If
        RegInt = 0: PenalInt = 0
        
        RstIndex.MoveNext
        If RstIndex.EOF Then GoTo NextAccount
        If LoanId <> RstIndex("LoanID") Then GoTo NextAccount
    End If
    
    TransType = FormatField(RstIndex("TransType"))
    
    If Abs(TransType) <> 1 Then
        If TransType = 7 Or TransType = -7 Then
            Amount = Amount * -1
        Else
            GoTo NextRecord
        End If
    End If
    Amount = FormatField(RstIndex("Amount"))
    TransType = IIf(TransType > 0, wDeposit, wWithDraw)
     
    SqlStr = "INSERT INTO LoanTrans (" & _
            " LoanId,TransDate," & _
            " TransID,TransType," & _
            " Amount, Balance, Particulars)"
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            "#" & RstIndex("TransDate") & "#," & _
            TransID & "," & TransType & "," & _
            Amount & "," & _
            FormatField(RstIndex("Balance")) & "," & _
            AddQuotes(FormatField(RstIndex("Particulars")), True) & ")"
    
    If TransID = 1 Then TransID = 2
    
    NewLoanTrans.SQLStmt = SqlStr
    If Not NewLoanTrans.SQLExecute Then
        NewLoanTrans.RollBack
        Exit Function
    End If
    
    If Not RstInst Is Nothing And (TransType = wContraDeposit Or TransType = wDeposit) Then
        Do
            If RstInst.EOF Then Exit Do
            If Amount <= 0 Then Exit Do
            InstAmount = FormatField(RstInst("InstBalance"))
            InstNo = FormatField(RstInst("InstNo"))
            If InstAmount >= Amount Then
                InstBalance = InstAmount - Amount
                Amount = 0 'Amount - Instp
            Else
                InstBalance = 0
                Amount = Amount - InstAmount
            End If
            SqlStr = "UPDATE LoanInst  Set InstBalance = " & InstBalance & _
                ", PaidDate = #" & RstIndex("TransDate") & "#" & _
                " WHERE LoanID = " & LoanId & _
                " AND InstNo = " & InstNo
            NewLoanTrans.SQLStmt = SqlStr
            If Not NewLoanTrans.SQLExecute Then
                NewLoanTrans.RollBack
                InTrans = False
                Exit Function
            End If
            RstInst.MoveNext
        Loop
    End If

NextRecord:
    RstIndex.MoveNext
    
NextAccount:
    If InTrans Then NewLoanTrans.CommitTrans
    InTrans = False
    If RstIndex.EOF Then GoTo ExitLine
    With frmMain
        .lblProgress = "Transferring Loan Transaction"
        .prg.Value = RstIndex.AbsolutePosition
    End With
    
Wend

ExitLine:
Debug.Print Now
    With frmMain
        .lblProgress = "Transferred the Loan  transaction"
        .prg.Value = 0
        .Refresh
    End With

Exit_line:
LoanTransTransfer = True

Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If
'Resume
End Function


Private Function SchemeTransfer(oldLoanTrans As clsOldUtils, NewLoanTrans As clsDBUtils) As Boolean
'Now transfer the Loans Schemes
Dim RstIndex As Recordset
oldLoanTrans.SQLStmt = "SELECT  * From LoanTypes"
If oldLoanTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then
    MsgBox "No Loan Schemes"
    GoTo Exit_line
End If
'Set RstIndex = oldLoanTrans.Rst.Clone

Dim SqlStr As String
Dim SchemeID As Integer
Dim SchemeName As String
Dim LoanCategary As wisLoanCategories
Dim TermType As Integer
Dim LoanType As Integer
LoanType = 1
Dim Monthduration As Integer
Dim DayDuration As Integer
Dim MaxRepayments As Integer
Dim InterestRate As Single
Dim PenalInterestRaste As Single
Dim InsurenceFee As Currency
Dim LegalFee As Currency
Dim Description  As String
Dim CreateDate As Date

CreateDate = Now

NewLoanTrans.BeginTrans
While Not RstIndex.EOF
    On Error Resume Next
    SchemeID = RstIndex("SchemeID")
    SchemeName = RstIndex("SchemeName")
    LoanType = 1
    LoanCategary = RstIndex("Category")
    TermType = FormatField(RstIndex("TermType"))
    CreateDate = RstIndex("Createdate")
    Monthduration = FormatField(RstIndex("MaxRepaymentTime")) * 12
    DayDuration = 0
    
    If Not Monthduration Then Monthduration = 4
    If CreateDate = Null Then CreateDate = Now
    
    On Error GoTo Err_Line
    
    SqlStr = "INSERT INTO LoanScheme (SchemeID, SchemeName," & _
            " Category,TermType,LoanType, MonthDuration," & _
            " DayDuration,Intrate, PenalIntrate," & _
            " LOanPurpose, InsuranceFee,LegalFee," & _
            " Description,Createdate ) "
    SqlStr = SqlStr & " Values (" & _
            SchemeID & "," & AddQuotes(SchemeName, True) & "," & _
            LoanCategary & ", " & TermType & "," & _
            LoanType & "," & Monthduration & "," & _
            DayDuration & "," & FormatField(RstIndex("InterestRate")) & _
            "," & FormatField(RstIndex("PenalInterestRate")) & "," & _
            " 'Individual'," & FormatField(RstIndex("InsuranceFee")) & "," & _
            FormatField(RstIndex("LegalFee")) & ",'Loan Name'," & _
            "#" & CreateDate & "#)"
    NewLoanTrans.SQLStmt = SqlStr
    If Not NewLoanTrans.SQLExecute Then
        NewLoanTrans.RollBack
        Exit Function
    End If
    
    RstIndex.MoveNext
Wend

NewLoanTrans.CommitTrans

Exit_line:
SchemeTransfer = True

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If
End Function


