Attribute VB_Name = "SbTransfer"
'This BAs file is used to Transfer
'SbMaster & Sb TranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit

Private m_HeadBalance As Currency

'this function creates the Sb Account head
'under the parnt head of Member deposits
'the new created head we can do the transaction
Private Function CreateSavingsHead(NewTrans As clsDBUtils) As Boolean

Dim SBHeadID As Long
Dim SBIntHeadId As Long

Dim ClsBank As clsBankAcc
Dim HeadName As String

Set ClsBank = New clsBankAcc
Dim FromDate As Date

FromDate = gFromDate '"3/31/03"

'First Create the share Heads
On Error GoTo ErrLine
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans


Dim rstTrans As ADODB.Recordset
Dim TotalDeposit As Currency
Dim TotalWithdraw As Currency
Dim TransDate As Date
Dim TransType As wisTransactionTypes

Dim PrgVal As Integer
frmMain.lblProgress = "Tranferring the Sb Ledger Accounts"
frmMain.Refresh
With frmMain.prg
    .Min = 0
    .Max = 165
    .Value = 0
End With


'Now INsert the Transcted Details to the acctrans table
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    " TransType,TransDate From SBTrans " & _
    " Where TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"

If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    
   
    ''Begin the Transaction
    NewIndexTrans.BeginTrans
    
    HeadName = LoadResString(gLangOffSet + 421) 'Savings Account
    SBHeadID = ClsBank.GetHeadIDCreated(HeadName, parMemberDeposit, m_HeadBalance, wis_SBAcc)
   
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            Call ClsBank.UpdateCashDeposits(SBHeadID, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashWithDrawls(SBHeadID, TotalWithdraw, TransDate)
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

    Call ClsBank.UpdateCashDeposits(SBHeadID, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashWithDrawls(SBHeadID, TotalWithdraw, TransDate)
    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
    
End If

'Now Insert the sb Interest details
NewTrans.SQLStmt = "Select sum(Amount) as TotalAmount," & _
    "TransType,TransDate From SBPLTrans " & _
    " Where TransDate >= #" & FromDate & "# " & _
    " Group By TransDate,TransType"
If NewTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then

    NewIndexTrans.BeginTrans
    'Create the Interest Head
    HeadName = LoadResString(gLangOffSet + 421) & " " & LoadResString(gLangOffSet + 487) 'Sb Interes Paid
    SBIntHeadId = ClsBank.GetHeadIDCreated(HeadName, parMemDepIntPaid, 0, wis_SBAcc)
    
    TransDate = rstTrans("Transdate")
    While Not rstTrans.EOF
        If TransDate <> rstTrans("Transdate") Then
            'Call ClsBank.UpdateCashDeposits(SBIntHeadId, TotalDeposit, TransDate)
            'Call ClsBank.UpdateCashWithDrawls(SBIntHeadId, TotalWithdraw, TransDate)
            Call ClsBank.UpdateCashWithDrawls(SBIntHeadId, TotalDeposit, TransDate)
            Call ClsBank.UpdateCashDeposits(SBIntHeadId, TotalWithdraw, TransDate)
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
    'Call ClsBank.UpdateCashDeposits(SBIntHeadId, TotalDeposit, TransDate)
    'Call ClsBank.UpdateCashWithDrawls(SBIntHeadId, TotalWithdraw, TransDate)
    Call ClsBank.UpdateCashWithDrawls(SBIntHeadId, TotalDeposit, TransDate)
    Call ClsBank.UpdateCashDeposits(SBIntHeadId, TotalWithdraw, TransDate)

    TotalDeposit = 0: TotalWithdraw = 0
    
    NewIndexTrans.CommitTrans
End If


CreateSavingsHead = True


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



'just calling this function we can transafer the sbmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferSB(OldDBName As String, NewDBName As String) As Boolean
Screen.MousePointer = vbHourglass
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
    If Not TransferSBMaster(OldTrans, NewTrans) Then GoTo ErrLine
    If Not TransferSBTrans(OldTrans, NewTrans) Then GoTo ErrLine
End If

If Not CreateSavingsHead(NewTrans) Then GoTo ErrLine
Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
TransferSB = True

ErrLine:

OldTrans.CloseDB
NewTrans.CloseDB
Set OldTrans = Nothing
Set NewTrans = Nothing

Screen.MousePointer = vbNormal

End Function

'this function is used to transfer the
'SB MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferSBMaster(OldSBTrans As clsOldUtils, NewSBTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SngSpace As String
Dim AccID As Long, IntroId As Long
Dim rstMain As ADODB.Recordset
Dim Rst As ADODB.Recordset

On Error GoTo Err_Line

 'Get the Head Balance
OldSBTrans.SQLStmt = "Select * From ObTab " & _
            " WHERE obDate = #" & DateAdd("D", 1, gFromDate) & "#"
If OldSBTrans.Fetch(rstMain, adOpenForwardOnly) > 0 Then
    Do
        If rstMain.EOF Then Exit Do
        If rstMain("Module") = 51 Then _
            m_HeadBalance = FormatField(rstMain("ObAmount")): Exit Do
        rstMain.MoveNext
    Loop
End If

'Fetch the detials of Sb Account
SqlStr = "SELECT * FROM SBMASTER A Where CustomerID in (Select Distinct CustomerID from NameTab) ORDER BY AccID"

OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(rstMain, adOpenForwardOnly) < 1 Then GoTo Exit_line

With frmMain
    .lblProgress = "Transferring Sb Account details"
    .prg.Max = rstMain.RecordCount + 1
    .prg.Value = 0
    .Refresh
End With
Dim NomineeInfo() As String

While Not rstMain.EOF
    If AccID = rstMain("AccID") Then GoTo NextAccount
    IntroId = FormatField(rstMain("Introduced"))
    'Get the Introducer ID
    If IntroId > 0 Then
        SqlStr = "SELECT CustomerID FROM SBMASTER " & _
            " WHERE AccID = " & IntroId
        OldSBTrans.SQLStmt = SqlStr
        If OldSBTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then IntroId = FormatField(Rst("CustomerID"))
    End If
    
    AccID = rstMain("AccID")
    Call GetStringArray(rstMain("Nominee"), NomineeInfo, ";")
    ReDim Preserve NomineeInfo(2)
    
    NewSBTrans.BeginTrans
'NOW insert into
    SqlStr = "Insert INTO SBMASTER (" & _
        "AccID,CustomerID,AccNUM,CreateDate,ModifiedDate,ClosedDate," & _
        "NomineeName,NomineeAge,NomineeRelation, " & _
        "IntroducerId,LedgerNo,FolioNo ," & _
        " InOperative,LastPrintId ,AccGroupID)"
    
    SqlStr = SqlStr & " VALUES (" & _
        AccID & "," & rstMain("CustomerID") & "," & _
        AddQuotes(rstMain("AccID"), True) & "," & _
        FormatDateField(rstMain("CreateDate")) & "," & _
        FormatDateField(rstMain("ModifiedDate")) & "," & _
        FormatDateField(rstMain("ClosedDate")) & "," & _
        AddQuotes(Left(NomineeInfo(0), 25), True) & "," & _
        Val(NomineeInfo(1)) & "," & _
        AddQuotes(Left(NomineeInfo(2), 15), True) & "," & _
        rstMain("Introduced") & ",'" & Val(rstMain("LedgerNo")) & "','" & Val(rstMain("FolioNo")) & "' ," & _
        "" & _
        False & "," & _
        "1, 1 )"

    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        NewSBTrans.RollBack
        OldSBTrans.BeginTrans
        OldSBTrans.SQLStmt = "DELETE * From SBTRANS " & _
            " Where AccID = " & rstMain("AccId")
        OldSBTrans.SQLExecute
        OldSBTrans.CommitTrans
        MsgBox "Unable to transafer the SB Account No " & rstMain("AccId")
        GoTo NextAccount
    End If
    NewSBTrans.CommitTrans
    
NextAccount:
    With frmMain
        .lblProgress = "Transferring Sb account details"
        .prg.Value = rstMain.AbsolutePosition
    End With
    rstMain.MoveNext
Wend

Exit_line:
TransferSBMaster = True
Exit Function

Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then MsgBox "eror In SBMaster " & Err.Description
    
End Function

'this function is used to transfer the
'SB transaction details form OLd Db to new one
'and NewSBTrans has assigned to new database
Private Function TransferSBTrans(OldSBTrans As clsOldUtils, NewSBTrans As clsDBUtils) As Boolean

On Error GoTo Err_Line

Dim SqlStr As String
Dim SqlInt As String
Dim IsIntTrans As Boolean
Dim OldTrans As Integer, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long

Dim rstMain As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim AccID As Long
Dim Amount As Currency
Dim TransDate As Date
'Fetch the detials of Sb Account
Dim FirstRound As Boolean
FirstRound = True

FirstLine:

If FirstRound Then

SqlStr = "SELECT Max(TransID) as MAxID, AccID as Acc FROM SbTrans " & _
    " Where TransDate < #" & gFromDate & "# group by AccID"
    OldSBTrans.SQLStmt = SqlStr
    
    OldSBTrans.CreateView ("qrySBMaxTransID")
    
    SqlStr = "SELECT A.* FROM SBTrans A , qrySBMaxTransID B " & _
            " Where A.Transid = B.MaxID And A.AccID = B.Acc " & _
            " ORDER BY A.AccID,A.TransId"
    
    'SqlStr = "SELECT A.* FROM SBTrans A " & _
            " Where A.Transid = (SELECT Max(TransID) FROM SbTrans C " & _
                " Where C.AccID = A.AccID " & _
                " And TransDate < #" & gFromDate & "#) " & _
            " ORDER BY AccID,TransId"
Else
    SqlStr = "SELECT * FROM SBTrans WHERE " & _
        " TransDate >= #" & gFromDate & "# " & _
        " ORDER BY AccID,TransId"
End If

OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(rstMain, adOpenStatic) < 1 Then GoTo Exit_line
'1701+ 1664
With frmMain
    .lblProgress = "Transferring Sb transaction details"
    .prg.Max = rstMain.RecordCount + 1
    .prg.Value = 0
    .Refresh
End With

'Open the Transaction
NewSBTrans.BeginTrans

While Not rstMain.EOF
    IsIntTrans = False
    OldTrans = FormatField(rstMain("TransType"))
    'TransID = FormatField(rstMain("TransID"))
    If OldTrans = 4 Or OldTrans = 2 Then IsIntTrans = True
    If OldTrans = -2 Or OldTrans = -4 Then IsIntTrans = True
    
    NewTransType = OldTrans
    NewTransType = IIf(OldTrans < 0, wWithDraw, wDeposit)
    If OldTrans = 1 Then NewTransType = wDeposit
    If OldTrans = -1 Then NewTransType = wWithDraw
    If OldTrans = 3 Then NewTransType = wContraDeposit
    If OldTrans = -3 Then NewTransType = wContraWithDraw
    
    If AccID <> rstMain("AccID") Then TransID = IIf(FirstRound, 0, 2)
    TransID = TransID + 1
    
    If IsIntTrans Then
        If OldTrans = 2 Or OldTrans = 4 Then NewTransType = wContraDeposit
        If OldTrans = -2 Or OldTrans = -4 Then NewTransType = wContraWithDraw
        
        TransDate = rstMain("TransDate")
        
        SqlInt = "Insert INTO SBPLTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            rstMain("AccID") & "," & _
            TransID & "," & _
            "#" & rstMain("TransDate") & "#," & _
            rstMain("Amount") & "," & _
            "0," & _
            NewTransType & " )"
        
        Amount = rstMain("Amount")
        AccID = rstMain("AccID")
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        
        'If Balance = FormatField(Rst("balance")) Then Rst.MoveNext
        
        'Insted of the above code the beleow code is written for
        'shiggaon case onmly
        'Except shiggaon in all banks the aboce code works 100%
        Do
            If TransDate <> rstMain("transdate") Or AccID <> rstMain("AccId") Then
                rstMain.MovePrevious
                Exit Do
            End If
            If OldTrans > 0 And Balance + Amount = rstMain("Balance") Then Exit Do
            If OldTrans < 0 And Balance - Amount = rstMain("Balance") Then Exit Do
            rstMain.MoveNext
        Loop
        OldTrans = rstMain("TransType")
        'After this transaction the transaction in the sb Table is contra
        'Therefore
        'NewTransType = (OldTrans / Abs(OldTrans)) * 3
        NewTransType = IIf(OldTrans < 0, wWithDraw, wDeposit)
        If OldTrans = 1 Then NewTransType = wDeposit
        If OldTrans = -1 Then NewTransType = wWithDraw
        If OldTrans = 3 Then NewTransType = wContraDeposit
        If OldTrans = -3 Then NewTransType = wContraWithDraw
    End If
    
    SqlStr = "Insert INTO SBTrans ( " & _
        "AccID,TransID,TransDate," & _
        "Amount,Balance,Particulars," & _
        "TransType,ChequeNo)"
    
    SqlStr = SqlStr & "VALUES (" & _
        rstMain("AccID") & "," & _
        TransID & "," & _
        "#" & rstMain("TransDate") & "#," & _
        rstMain("Amount") & "," & _
        rstMain("Balance") & "," & _
        AddQuotes(Left(FormatField(rstMain("Particulars")), 25), True) & "," & _
        NewTransType & "," & FormatField(rstMain("ChequeNo")) & " )"
    
    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        MsgBox "Unable to transafer the SB Trans data"
        NewSBTrans.RollBack
        Exit Function
    End If
    'if any transaction s do it
    If SqlInt <> "" Then
        NewSBTrans.SQLStmt = SqlInt
        If Not NewSBTrans.SQLExecute Then
            MsgBox "Unable to transafer the SB Trans data"
            NewSBTrans.RollBack
            Exit Function
        End If
        SqlInt = ""
    End If
    'If Rst.AbsolutePosition Mod 5000 = 0 Then Debug.Print Now
    Balance = FormatField(rstMain("Balance"))
    
NextAccount:
    
    With frmMain
        .lblProgress = "Transferring Sb transactions"
        .prg.Value = rstMain.AbsolutePosition
    End With
    AccID = rstMain("AccID")
    
    rstMain.MoveNext
Wend

NewSBTrans.CommitTrans

If FirstRound Then
    FirstRound = False
    GoTo FirstLine
End If


Exit_line:
TransferSBTrans = True
Exit Function
Err_Line:

If Err.Number = 3021 Then Err.Clear: Resume Next
If Err Then MsgBox "Error in SBTrans" & Err.Description
    
End Function


