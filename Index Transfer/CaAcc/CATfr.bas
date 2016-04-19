Attribute VB_Name = "CaTransfer"
'This BAs file is used to Transfer
'CAMaster & CA TranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit

Private m_HeadBalance As Currency



'this function creates the CA Account head
'under the parnt head of Member deposits
'the new created head we can do the transaction
Private Function CreateCurrnetAccountHead(NewTrans As clsDBUtils) As Boolean

Dim CAHeadID As Long
Dim CAIntHeadId As Long

Dim ClsBank As clsBankAcc
Dim HeadName As String

Set ClsBank = New clsBankAcc
Dim FromDate As Date

FromDate = gFromDate

On Error GoTo ErrLine
'First Create the share Heads
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans
Dim PrgVal As Integer
frmMain.lblProgress = "Trasferring the Current Account Ledger"
frmMain.Refresh

With frmMain.prg
    .Max = 365
    .Min = 0
    .Value = PrgVal
End With


Dim rstTrans As ADODB.Recordset
Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency
Dim TransDate As Date
Dim TransType As wisTransactionTypes

'Now INsert the Transcted Details to the acctrans table
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    " TransType,TransDate From CATrans " & _
    " Where TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    ''Begin the Transaction
    NewIndexTrans.BeginTrans
    HeadName = LoadResString(gLangOffSet + 422) 'Current Account
    CAHeadID = ClsBank.GetHeadIDCreated(HeadName, parMemberDeposit, m_HeadBalance, wis_CAAcc)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(CAHeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(CAHeadID, TotalWithdraw, TransDate)
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

    Call ClsBank.UpdateCashDeposits(CAHeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(CAHeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now Insert the CA Interest details
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    "TransType,TransDate From CAPLTrans " & _
    " Where TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    
    'Create the Interest Head
    HeadName = LoadResString(gLangOffSet + 422) & " " & LoadResString(gLangOffSet + 487) 'CA Interes Paid
    CAIntHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_CAAcc)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(CAIntHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(CAIntHeadId, TotalWithdraw, TransDate)
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
    Call ClsBank.UpdateCashDeposits(CAIntHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(CAIntHeadId, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

CreateCurrnetAccountHead = True


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

'just calling this function we can transafer the CAmaster Old to new
'Arguments for this function are OldcaTrans & new caTrans
'Old ca Trans is assigned to Old database
'and NewcaTrans has assigned to new database
Public Function TransferCA(OldDBName As String, NewDBName As String) As Boolean
Debug.Print Now
Dim OldTrans As New clsOldUtils
Dim NewTrans As New clsDBUtils
If Not OldTrans.OpenDB(OldDBName, OldPwd) Then Exit Function
If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    Exit Function
End If

If Not gOnlyLedgerHeads Then
    If Not TransferCAMaster(OldTrans, NewTrans) Then Exit Function
    If Not TransferCATrans(OldTrans, NewTrans) Then Exit Function
End If

If Not CreateCurrnetAccountHead(NewTrans) Then Exit Function
    
    TransferCA = True
    If Not PutVoucherNumber(NewTrans) Then
        MsgBox "Unable to set the voucher No"
        Exit Function
    End If
    
End Function


'this function is used to transfer the
'ca MAster details form OLdb to new one
'and NewcaTrans has assigned to new database
Private Function TransferCAMaster(OldSBTrans As clsOldUtils, NewSBTrans As clsDBUtils) As Boolean
Dim SqlStr As String
On Error GoTo Err_Line

'Update Modify date
With frmMain
    .lblProgress = "Getting current account Details"
    .prg.Value = 0
    .Refresh
End With

Dim AccID As Long, IntroId As Long
Dim Rst As Recordset
Dim rstTemp As Recordset
 'Get the Head Balance

OldSBTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#" & _
            " ORder By obDate Desc"
If OldSBTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    Do
        If rstTemp.EOF Then Exit Do
        If rstTemp("Module") = 52 Then _
            m_HeadBalance = FormatField(rstTemp("ObAmount")): Exit Do
        rstTemp.MoveNext
    Loop
End If

'Fetch the detials of ca Account
SqlStr = "SELECT * FROM CAMASTER ORDER BY AccID"
OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(Rst, adOpenDynamic) < 1 Then TransferCAMaster = True: Exit Function


With frmMain
    .lblProgress = "Transferring the data of current account transactions "
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With

On Error GoTo Err_Line
Dim NomineeInfo() As String


While Not Rst.EOF
    If AccID = FormatField(Rst("AccID")) Then GoTo NextAccount
    
    IntroId = FormatField(Rst("Introduced"))
    'Get the Introducer ID
    If IntroId > 0 Then
        SqlStr = "SELECT CustomerID FROM CAMASTER " & _
            " WHERE AccID = " & IntroId
        NewSBTrans.SQLStmt = SqlStr
        If NewSBTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then IntroId = FormatField(rstTemp("CustomerID"))
    End If
    
    AccID = Rst("AccID")
    'First insert into Cb joint table
    SqlStr = "Insert INTO CAJOINT (" & _
        "AccID,CustomerID,CustomerNum)" & _
        "VALUES (" & _
        Rst("AccID") & "," & Rst("CustomerID") & "," & _
        "1 )"
        
    Call GetStringArray(Rst("Nominee"), NomineeInfo, ";")
    ReDim Preserve NomineeInfo(2)
    
    NewSBTrans.BeginTrans
    NewSBTrans.SQLStmt = SqlStr
    
'NOW insert into
    SqlStr = "Insert INTO CAMASTER (" & _
        "AccID,CustomerID,AccNUM,CreateDate,ModifiedDate,ClosedDate," & _
        "JointHolder,NomineeName,NomineeAge,NomineeRelation," & _
        "IntroducerId,LedgerNo,FolioNo," & _
        "AccGroupID, InOperative,LastPrintId )"
    
    SqlStr = SqlStr & " VALUES (" & _
        Rst("AccID") & "," & Rst("CustomerID") & "," & _
        AddQuotes(Rst("AccID"), True) & "," & _
        FormatDateField(Rst("CreateDate")) & "," & FormatDateField(Rst("Modifieddate")) & "," & _
        FormatDateField(Rst("ClosedDate")) & " ," & _
        AddQuotes(FormatField(Rst("JointHolder")), True) & ", " & _
        AddQuotes(NomineeInfo(0), True) & "," & _
        Val(NomineeInfo(1)) & "," & _
        AddQuotes(NomineeInfo(2), True) & "," & _
        Rst("Introduced") & ",'" & Val(Rst("LedgerNo")) & "'," & _
        "'" & Val(Rst("FolioNo")) & "' , 1, " & _
        False & "," & _
        "1 )"
        
    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        NewSBTrans.RollBack
        MsgBox "Unable to transafer the CA MAster data"
        Exit Function
    End If
    NewSBTrans.CommitTrans
    
NextAccount:
    With frmMain
        .lblProgress = "Transferring the data of BKCC Loan Master Details"
        .prg.Value = Rst.AbsolutePosition
    End With
    Rst.MoveNext
    
Wend

'Now reverse the change made before transfer
'Update Modify date
SqlStr = "UPDATE CAMASTER Set ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
OldSBTrans.SQLStmt = SqlStr
OldSBTrans.BeginTrans
If Not OldSBTrans.SQLExecute Then
    OldSBTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldSBTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE CAMASTER Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
OldSBTrans.SQLStmt = SqlStr
OldSBTrans.BeginTrans
If Not OldSBTrans.SQLExecute Then
    OldSBTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldSBTrans.CommitTrans

'Now Update the smae with new database
'Update Modify date
SqlStr = "UPDATE CAMAster Set ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
NewSBTrans.SQLStmt = SqlStr
NewSBTrans.BeginTrans
If Not NewSBTrans.SQLExecute Then
    NewSBTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewSBTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE CAMAster Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
NewSBTrans.SQLStmt = SqlStr
NewSBTrans.BeginTrans
If Not NewSBTrans.SQLExecute Then
    NewSBTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewSBTrans.CommitTrans
    
SqlStr = "UPDATE CAMASTER Set NomineeAge = NULL WHERE NomineeAge = 0 "
NewSBTrans.SQLStmt = SqlStr
NewSBTrans.BeginTrans
If Not NewSBTrans.SQLExecute Then
    NewSBTrans.RollBack
Else
    NewSBTrans.CommitTrans
End If
    
TransferCAMaster = True
    Debug.Print Now
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then
        MsgBox "eror In CAMaster " & vbCrLf & Err.Description
        Err.Clear
    End If
    
End Function


'this function is used to transfer the
'ca transaction details form OLd Db to new one
'and NewCATrans has assigned to new database
Private Function TransferCATrans(OldSBTrans As clsOldUtils, NewSBTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim IsIntTrans As Boolean

On Error GoTo Err_Line

Dim OldTrans As Integer, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim Rst As Recordset
Dim AccID As Long
Dim Amount As Currency
Dim TransDate As Date
    'Fetch the detials of Sb Account
With frmMain
    .lblProgress = "Transferring current account transaction"
    .prg.Value = 0
    .Refresh
End With

Dim FirstRound As Boolean

FirstRound = True

FirstLine:

If FirstRound Then
  
    SqlStr = "SELECT A.* FROM CATrans A " & _
            " Where A.Transid = (SELECT Max(TransID) FROM CATrans C " & _
                " Where C.AccID = A.AccID " & _
                " And TransDate < #" & gFromDate & "#) " & _
            " ORDER BY AccID,TransId"
            
SqlStr = "SELECT Max(TransID) as MaxID, AccID as Acc FROM CATrans " & _
    " Where TransDate < #" & gFromDate & "# group by AccID"
    OldSBTrans.SQLStmt = SqlStr
    
    OldSBTrans.CreateView ("qryCAMaxTransID")
    
    SqlStr = "SELECT A.* FROM CATrans A , qryCAMaxTransID B " & _
            " Where A.Transid = B.MaxID And A.AccID = B.Acc " & _
            " ORDER BY A.AccID,A.TransId"
 
            
Else
    SqlStr = "SELECT * FROM CATrans Where TransDate >= #" & gFromDate & "# " & _
        " ORDER BY AccID,TransId"
End If

OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(Rst, adOpenDynamic) < 1 Then TransferCATrans = True: Exit Function

'Set rst = OldSBTrans.rst.Clone
With frmMain
    .lblProgress = "Transferring current account transaction"
    .prg.Max = Rst.RecordCount + 1
    .Refresh
End With

TransID = IIf(FirstRound, 1, 2)
While Not Rst.EOF
    IsIntTrans = False
    OldTrans = Rst("TransType")
'    If OldTrans = wContraInterest Or OldTrans = wInterest Then IsIntTrans = True
'    If OldTrans = wContraCharges Or OldTrans = wCharges Then IsIntTrans = True
    If OldTrans = 4 Or OldTrans = 2 Then IsIntTrans = True
    If OldTrans = -2 Or OldTrans = -4 Then IsIntTrans = True
    
    If OldTrans = 1 Then NewTransType = wDeposit
    If OldTrans = -1 Then NewTransType = wWithDraw
    If OldTrans = 3 Then NewTransType = wContraDeposit
    If OldTrans = -3 Then NewTransType = wContraWithDraw
    
    If AccID <> Rst("accID") Then TransID = IIf(FirstRound, 0, 2)
    TransID = TransID + 1
    
    If IsIntTrans Then
'        If OldTrans = wInterest Then NewTransType = -1
'        If OldTrans = wContraInterest Then NewTransType = -3
'        If OldTrans = wCharges Then NewTransType = 3
'        If OldTrans = wContraCharges Then NewTransType = 1
        If OldTrans = 2 Or OldTrans = 4 Then NewTransType = wContraWithDraw
        If OldTrans = -2 Or OldTrans = -4 Then NewTransType = wContraDeposit
        
        TransDate = Rst("TransDate")
        SqlInt = "Insert INTO CAPLTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            Rst("AccID") & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0  ," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
            NewTransType & " )"
        
        Amount = Rst("Amount")
        AccID = Rst("AccID")
        
        OldTrans = Rst("TransType")
        'After this transaction the transaction in the sb Table is contra
        'Therefore
        If OldTrans = 1 Then NewTransType = wDeposit
        If OldTrans = -1 Then NewTransType = wWithDraw
        If OldTrans = 3 Then NewTransType = wContraDeposit
        If OldTrans = -3 Then NewTransType = wContraWithDraw
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        If Balance = FormatField(Rst("balance")) Then
            Rst.MoveNext
            NewTransType = IIf(Rst("transType") < 0, wContraDeposit, wContraWithDraw)
        End If
    End If
    
    SqlStr = "Insert INTO CATrans ( " & _
        "AccID,TransID,TransDate," & _
        "Amount,Balance,Particulars," & _
        "TransType,ChequeNo)"
    
    SqlStr = SqlStr & "VALUES (" & _
        Rst("AccID") & "," & _
        TransID & "," & _
        "#" & Rst("TransDate") & "#," & _
        Rst("Amount") & "," & _
        Rst("Balance") & "," & _
        AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
        NewTransType & "," & FormatField(Rst("ChequeNo")) & " )"
    
    NewSBTrans.BeginTrans
    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        MsgBox "Unable to transafer the Current A/c Trans data"
        NewSBTrans.RollBack
        Exit Function
    End If
    If SqlInt <> "" Then
        NewSBTrans.SQLStmt = SqlInt
        If Not NewSBTrans.SQLExecute Then
            MsgBox "Unable to transafer the Current A/c Trans data"
            NewSBTrans.RollBack
            Exit Function
        End If
        SqlInt = ""
    End If
    'If Rst.AbsolutePosition Mod 5000 = 0 Then Debug.Print Now
    NewSBTrans.CommitTrans
    Balance = FormatField(Rst("Balance"))
NextAccount:
    
    With frmMain
        .lblProgress = "Transferring current account transaction"
        .prg.Value = Rst.AbsolutePosition
    End With
    AccID = Rst("AccID")
    Rst.MoveNext
    
Wend

If FirstRound Then
    FirstRound = False
    GoTo FirstLine
End If



TransferCATrans = True
    With frmMain
        .lblProgress = "Transferred the CA transaction"
        .prg.Value = 0
        .Refresh
    End With

Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err Then
    MsgBox "Error in CATrans" & Err.Description
    Err.Clear
End If
    
End Function
'this function is used to transfer the
'set the voucher no fot the transaferred data
Private Function PutVoucherNumber(NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String


PutVoucherNumber = True
End Function
