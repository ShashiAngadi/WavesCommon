Attribute VB_Name = "MMTransfer"
Option Explicit

Private m_HeadBalance As Currency

'
'this function creates the Member share head in Memeber share head
'under the parnt head of shre captial-memebr share
'in the new created head we can do the transaction
Private Function CreateMemberHeads(NewTrans As clsDBUtils) As Boolean

Dim ShareHeadID As Long
Dim MemberFeeId As Long
Dim ShareFeeID As Long

Dim ClsBank As clsBankAcc
Dim HeadName As String

Set ClsBank = New clsBankAcc

Dim FromDate As Date

On Error GoTo ErrLine
Dim PrgVal As Integer
frmMain.lblProgress = "Transferring the Member Ledger Head"
frmMain.Refresh
With frmMain.prg
    .Max = 365
    .Min = 0
    .Value = 0
End With


FromDate = gFromDate '"3/31/03"

'First Create the share Heads
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans

    
''Begin the Transaction
NewIndexTrans.BeginTrans

HeadName = LoadResString(gLangOffSet + 53) & " " & LoadResString(gLangOffSet + 36) 'Share account
ShareHeadID = ClsBank.GetHeadIDCreated(HeadName, parMemberShare, m_HeadBalance, wis_Members)
'Create the Memeberhip Fee
HeadName = LoadResString(gLangOffSet + 79) & " " & LoadResString(gLangOffSet + 191) 'Memebrship Fee
MemberFeeId = ClsBank.GetHeadIDCreated(HeadName, parBankIncome, 0, wis_Members)
'Create the share Fee
HeadName = LoadResString(gLangOffSet + 53) & " " & LoadResString(gLangOffSet + 191)  'Share Fee
ShareFeeID = ClsBank.GetHeadIDCreated(HeadName, parBankIncome, 0, wis_Members)

'commit the tranaction
NewIndexTrans.CommitTrans
CreateMemberHeads = True

Dim rstTrans As ADODB.Recordset
Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency
Dim TransDate As Date
Dim TransType As wisTransactionTypes

'Now INsert the Transcted Details to the acctrans table
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    " TransType,TransDate From MemTrans " & _
    " Where TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
    NewIndexTrans.BeginTrans
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(ShareHeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(ShareHeadID, TotalWithdraw, TransDate)
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

    Call ClsBank.UpdateCashDeposits(ShareHeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(ShareHeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now Insert the Memeber fee deatils
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    "TransType,TransDate From MemIntTrans " & _
    " Where TransID = 1 And TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(MemberFeeId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(MemberFeeId, TotalWithdraw, TransDate)
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
    Call ClsBank.UpdateCashDeposits(MemberFeeId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(MemberFeeId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now Insert the Share fee deatils
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    "TransType,TransDate From MemIntTrans " & _
    "Where TransID <> 1 AND TransDate >= #" & FromDate & "# " & _
    "Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(ShareFeeID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(ShareFeeID, TotalWithdraw, TransDate)
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
    
    Call ClsBank.UpdateCashDeposits(ShareFeeID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(ShareFeeID, TotalWithdraw, TransDate)
    
    TotalDeposit = 0: TotalWithdraw = 0
    NewIndexTrans.SQLStmt = "UPDATe opBalance Set opdate = #4/1/2003#"
    NewIndexTrans.SQLExecute
    
    NewIndexTrans.CommitTrans
    
End If

NewIndexTrans.BeginTrans
    NewIndexTrans.SQLStmt = "UPDATe opBalance Set opdate = #4/1/2003#"
    NewIndexTrans.SQLExecute
NewIndexTrans.CommitTrans

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
    'GoTo Exit_line
    'Resume
End If

'NewIndexTrans.RollBack

End Function

Private Function MemMasterTransfer(oldMMTrans As clsOldUtils, NewMMTrans As clsDBUtils) As Boolean
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

On Error GoTo Err_Line

    
oldMMTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#"
If oldMMTrans.Fetch(RstIndex, adOpenDynamic) > 0 Then
    Do
        If RstIndex.EOF Then Exit Do
        If RstIndex("Module") = 61 Then _
            m_HeadBalance = FormatField(RstIndex("ObAmount")): Exit Do
        RstIndex.MoveNext
    Loop
End If

'This function Tranfers all the data from MMMaster of Index 200
'to MMMaster Of New data Base
'Fetch the data from Old Database

SqlStr = "SELECT * FROM MMMaster ORDER BY AccId"
oldMMTrans.SQLStmt = SqlStr
If oldMMTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then Exit Function

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id

With frmMain
    .lblProgress = "Transferring member details"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
End With

Dim Count As Integer

NewMMTrans.BeginTrans
While Not RstIndex.EOF
 
'Count = Count + 1
'Debug.Assert Count Mod 20 <> 0
    'first get the mem Id of the old db
     MemID = FormatField(RstIndex("AccId"))
    'now get the customer Of this Member
    CustomerId = 0
    CustomerId = FormatField(RstIndex("CustomerId")) 'FormatField(oldLoanTrans.Rst(0))
    'If CustomerId = 0 Then GoTo NextAccount
    On Error GoTo Err_Line
    SqlStr = "INSERT INTO MemMaster (" & _
            " AccID,AccNum,CustomerID," & _
            " CreateDate,ModifiedDate,ClosedDate," & _
            " NomineeID,IntroducerID,LedgerNo," & _
            " FolioNo,MemberType,AccGroupID) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("AccID") & "," & _
            "'" & RstIndex("AccID") & "' ," & _
            RstIndex("CustomerID") & "," & _
            FormatDateField(RstIndex("CreateDate")) & "," & _
            FormatDateField(RstIndex("ModifiedDate")) & "," & _
            FormatDateField(RstIndex("ClosedDate")) & "," & _
            "0 ," & RstIndex("Introduced") & ", " & _
            AddQuotes(Trim(FormatField(RstIndex("LedgerNo"))), True) & "," & _
            AddQuotes(Trim(FormatField(RstIndex("FolioNo"))), True) & "," & _
            RstIndex("MemberType") + 1 & ",1 ) "
    
    NewMMTrans.SQLStmt = SqlStr
    If Not NewMMTrans.SQLExecute Then    'THere aRE MORE FIELDS
        'gDBTrans.CommitTrans
        NewMMTrans.RollBack      'IN lOANS THAN iNDEX
        Exit Function
    End If
    'If loan has got the installment then insert the installment details
    
NextAccount:
    With frmMain
        .lblProgress = "Transferring member details"
        .prg.Value = RecNo
         If RecNo Mod 50 = 0 Then .Refresh
    End With
    RecNo = RecNo + 1
    RstIndex.MoveNext
Wend

NewMMTrans.CommitTrans


'Update Modify date
SqlStr = "UPDATE MMMASTER SET ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
oldMMTrans.SQLStmt = SqlStr
oldMMTrans.BeginTrans
If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If

'Update Closed date
SqlStr = "UPDATE MMMASTER Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
oldMMTrans.SQLStmt = SqlStr

If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
oldMMTrans.CommitTrans

SqlStr = "UPDATE MemMASTER SET ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
NewMMTrans.SQLStmt = SqlStr
NewMMTrans.BeginTrans
If Not NewMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If

'Update Closed date
SqlStr = "UPDATE MemMASTER Set ClosedDate = NULL Where ClosedDate = #1/1/100#"
NewMMTrans.SQLStmt = SqlStr
If Not NewMMTrans.SQLExecute Then
    NewMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewMMTrans.CommitTrans

MemMasterTransfer = True
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If
'Resume

End Function



Private Function MemTransTransfer(oldMemTrans As clsOldUtils, NewMemTrans As clsDBUtils) As Boolean
'Dim
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim TransID As Long
Dim TransType As Integer
Dim ItIsIntTrans As Boolean
Dim Amount As Currency
Dim InstAmount As Currency
Dim InstNo As Integer
Dim InstBalance As Currency
Dim ProcCount As Long


'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans
With frmMain
    .lblProgress = "Transferring shar transaction "
    .prg.Value = 0
    .Refresh
End With

SqlStr = "SELECT * FROM MMTrans ORDER By AccID, TransID"
oldMemTrans.SQLStmt = SqlStr

If oldMemTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then GoTo ExitLine
    
'Set RstIndex = oldMemTrans.rst.Clone

With frmMain
    .lblProgress = "Transferring share transaction"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
    TransID = 0
End With
    
NewMemTrans.BeginTrans
While Not RstIndex.EOF

    ItIsIntTrans = False
    TransType = FormatField(RstIndex("TransType"))
    If MemID <> RstIndex("accID") Then TransID = 1
    MemID = RstIndex("accID")
'    Debug.Assert MemID <> 44
    'Begin the transaction
    If TransType = -2 Or TransType = 2 Then
        TransID = TransID + 1
        ItIsIntTrans = True
        TransType = IIf(TransType > 0, wWithDraw, wDeposit)
        Amount = FormatField(RstIndex("Amount"))
        If RstIndex("transid") <> 1 And Amount = 0 Then GoTo NextAccount
        If RstIndex("Transid") = 1 Then TransID = 1
        SqlStr = "INSERT INTO MemIntTrans (" & _
                " AccId,TransDate," & _
                " TransID,TransType," & _
                " Amount,Balance " & _
                " )"
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("AccId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & TransType & "," & _
                Amount & "," & _
                " 0 )"
    Else
        TransID = TransID + 1
        TransType = FormatField(RstIndex("TransType"))
        Amount = FormatField(RstIndex("Amount"))
        
        TransType = IIf(TransType < 0, wWithDraw, wDeposit)
        
        SqlStr = "INSERT INTO MemTrans (" & _
                " AccId,TransDate," & _
                " TransID,Leaves,TransType," & _
                " Amount, Balance)"
        
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("AccId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & _
                RstIndex("Leaves") & "," & _
                TransType & "," & _
                Amount & "," & _
                RstIndex("Balance") & "  )"
    End If
    
    NewMemTrans.SQLStmt = SqlStr
    If Not NewMemTrans.SQLExecute Then
        NewMemTrans.RollBack
        Exit Function
    End If

NextAccount:
    ProcCount = ProcCount + 1
    With frmMain
        .lblProgress = "Transferring share transaction"
        .prg.Value = ProcCount
        If ProcCount Mod 50 = 0 Then .Refresh
    End With
 RstIndex.MoveNext

Wend
        
    

    NewMemTrans.CommitTrans

ExitLine:

Debug.Print Now
MemTransTransfer = True
    With frmMain
        .lblProgress = "Transferred the member details"
        .prg.Value = 0
        .Refresh
    End With

Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If
'Resume
End Function


Private Function ShareTransfer(oldMemTrans As clsOldUtils, NewMemTrans As clsDBUtils) As Boolean
'Now transfer the Loans Schemes

Dim RstIndex As Recordset
Dim ProcCount As Long

oldMemTrans.SQLStmt = "SELECT  * From ShareLeaves"
If oldMemTrans.Fetch(RstIndex, adOpenDynamic) < 1 Then
    MsgBox "No Loan Schemes"
    Exit Function
End If
'Set RstIndex = oldMemTrans.rst.Clone

Dim SqlStr As String
Dim AccID As Integer
Dim SchemeName As String
Dim Description  As String
Dim CreateDate As Date

CreateDate = Now
With frmMain
    .lblProgress = "Transferring share certificate"
    .prg.Max = RstIndex.RecordCount + 1
    .Refresh
End With

NewMemTrans.BeginTrans
While Not RstIndex.EOF
    
    SqlStr = "INSERT INTO ShareTrans (AccID, SaleTransID," & _
            " ReturnTransID, CertNo,CertID," & _
            " FaceValue ) "
    SqlStr = SqlStr & " Values (" & _
            RstIndex("AccID") & "," & _
            RstIndex("SaleTransID") & "," & _
            FormatField(RstIndex("ReturnTransID")) & "," & _
            AddQuotes(RstIndex("CertNo"), True) & "," & _
            RstIndex("CertNo") & "," & _
            RstIndex("FaceValue") & ")"
    NewMemTrans.SQLStmt = SqlStr
'    AccId   SaleTransID ReturnTransID   CertNo  CertId  FaceValue
'    675     7            0              7452    7452    $100.00
    If Not NewMemTrans.SQLExecute Then
        NewMemTrans.RollBack
        Exit Function
    End If
    With frmMain
        .lblProgress = "Transferring share Ledger"
        ProcCount = ProcCount + 1
        .prg.Value = ProcCount
        If ProcCount Mod 50 = 0 Then .Refresh
    End With
    RstIndex.MoveNext
    
Wend
NewMemTrans.CommitTrans
ShareTransfer = True

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In BKCC Master"
    Err.Clear
End If

End Function


Public Function MemberTransfer(OldDBName As String, NewDBName As String) As Boolean
Dim OldTrans As New clsOldUtils
Dim NewTrans As New clsDBUtils

If Not OldTrans.OpenDB(OldDBName, OldPwd) Then
    MsgBox "No old Index Db"
    Exit Function
End If

If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    MsgBox " No new Index Db"
    Exit Function
End If

Screen.MousePointer = vbHourglass
If Not gOnlyLedgerHeads Then
    If Not MemMasterTransfer(OldTrans, NewTrans) Then GoTo ExitLine
    If Not MemTransTransfer(OldTrans, NewTrans) Then GoTo ExitLine
    
    Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
        
    If Not ShareTransfer(OldTrans, NewTrans) Then GoTo ExitLine
End If

If Not CreateMemberHeads(NewTrans) Then GoTo ExitLine

Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
    
MemberTransfer = True

ExitLine:

Screen.MousePointer = vbDefault

End Function

