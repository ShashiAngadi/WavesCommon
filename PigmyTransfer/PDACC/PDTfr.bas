Attribute VB_Name = "PDTransfer"
'This BAs file is used to Transfer
'Pigmy Master & pigmyTranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit
Private m_AccOffSet As Long
Private m_LoanOffSet As Long
'just calling this function we can transafer the PDmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferPD(OldDBName As String, NewDBName As String) As Boolean
Debug.Print Now
Dim OldTrans As New clsTransact
Dim NewTrans As New clsTransact
If Not OldTrans.OpenDB(OldDBName, "WIS!@#") Then Exit Function
If Not NewTrans.OpenDB(NewDBName, "WIS!@#") Then
    OldTrans.CloseDB
    Exit Function
End If
    Screen.MousePointer = vbHourglass
    If Not TransferPDMaster(OldTrans, NewTrans) Then GoTo ExitLine
    If Not TransferPDTrans(OldTrans, NewTrans) Then GoTo ExitLine
    If Not TransferPDLoanMaster(OldTrans, NewTrans) Then GoTo ExitLine
    If Not TransferPDLoanTrans(OldTrans, NewTrans) Then GoTo ExitLine
    
    TransferPD = True
    If Not PutVoucherNumber(NewTrans) Then
        MsgBox "Unable to set the voucher No"
        GoTo ExitLine
    End If
    
ExitLine:
Screen.MousePointer = vbDefault
End Function
'this function is used to transfer the
'SB MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferPDMaster(OldPdTrans As clsTransact, NewPDTrans As clsTransact) As Boolean
Dim SqlStr As String

Dim Rst As Recordset

On Error GoTo Err_Line

'Before Fetching Update the Values
'where It can be Null with default value
'Then Fetch the records

'Delete the any details in the PDMAster Whose transactions has not done
SqlStr = "SELECT Distinct UserID FROM PDMASTER"
OldPdTrans.SQLStmt = SqlStr

If OldPdTrans.SQLFetch < 1 Then GoTo Exit_Line

Set Rst = OldPdTrans.Rst.Clone
While Not Rst.EOF
    SqlStr = "DELETE * FROM PDMAster WHERE USERID = " & Rst(0) & _
        " AND AccId NOT IN (SELECT Distinct AccID From PDTrans " & _
            " WHERE USERID = " & Rst(0) & ")"
    OldPdTrans.SQLStmt = SqlStr
    OldPdTrans.BeginTrans
'    If Not OldPdTrans.SQLExecute Then
'        OldPdTrans.RollBack
'        MsgBox "UPdate the Data Base Instead Of transfer"
'        Exit Function
'    End If
    OldPdTrans.CommitTrans
    Rst.MoveNext
Wend

'Update Modify date
SqlStr = "UPDATE PDMASTER SET ModifiedDate = #1/1/100# WHERE ModifiedDate = NULL"
OldPdTrans.SQLStmt = SqlStr
OldPdTrans.BeginTrans
If Not OldPdTrans.SQLExecute Then
    OldPdTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldPdTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE PDMASTER Set ClosedDate = #1/1/100# WHEre ClosedDate = NULL"
OldPdTrans.SQLStmt = SqlStr
OldPdTrans.BeginTrans
If Not OldPdTrans.SQLExecute Then
    OldPdTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldPdTrans.CommitTrans

'Update Nominee
Dim SngSpace As String
SngSpace = ""
SqlStr = "UPDATE PDMASTER set Nominee = '" & SngSpace & "'" & _
        " WHERE Nominee = NULL"
OldPdTrans.SQLStmt = SqlStr
OldPdTrans.BeginTrans
If Not OldPdTrans.SQLExecute Then
    OldPdTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldPdTrans.CommitTrans

SqlStr = "UPDATE PDMASTER set JointHolder= '" & SngSpace & "'" & _
        " WHERE JointHolder = NULL"
OldPdTrans.SQLStmt = SqlStr
OldPdTrans.BeginTrans
If Not OldPdTrans.SQLExecute Then
    OldPdTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
OldPdTrans.CommitTrans

'Fetch the detials of Pifmy  Account
SqlStr = "SELECT * FROM PDMASTER ORDER BY UserId,AccID"
Dim AccId As Long, IntroId As Long
Dim UserID As Long, AccNum As String
OldPdTrans.SQLStmt = SqlStr

If OldPdTrans.SQLFetch < 1 Then Exit Function
Set Rst = OldPdTrans.Rst.Clone

'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldPdTrans.SQLStmt = SqlStr
Call OldPdTrans.SQLFetch

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
'm_AccOffSet= FormatField(OldPdTrans.Rst(0))
'm_AccOffSet= AccOffSet + 100 - (AccOffSet Mod 100)

AccId = m_AccOffSet
While Not Rst.EOF
    'If AccID = FormatField(Rst("AccID")) Then GoTo NextAccount
    IntroId = FormatField(Rst("Introduced"))
    'Get the Introducer ID
    
    If IntroId > 0 Then
        SqlStr = "SELECT CustomerID FROM PDMASETR " & _
            " WHERE AccID = " & IntroId
        If OldPdTrans.SQLFetch > 0 Then IntroId = FormatField(OldPdTrans.Rst("CustomerID"))
    End If
        
    'AccNum = Format(Rst("AccId"), "000")
    AccNum = Rst("UserID") & "_" & Rst("AccId")
    NewPDTrans.BeginTrans
    
'NOW insert into
    'AccID = Val(Rst("UserID")) * AccOffSet + Val(Rst("AccId"))
    'AccID = AccID - AccOffSet
    AccId = AccId + 1
    
    SqlStr = "Insert INTO PDMASTER (" & _
        "AccID,AgentID,CustomerID,AccNum," & _
        "CreateDate,ModifiedDate,ClosedDate," & _
        "MaturityDate,PigmyAmount,PigmyType," & _
        "RateOfInterest,Nominee,Introduced," & _
        " LedgerNo,FolioNo,NomineeID,LastPrintId )"
    
    SqlStr = SqlStr & " VALUES (" & _
        AccId & "," & _
        Rst("UserID") & "," & _
        Rst("CustomerID") & "," & _
        AddQuotes(AccNum, True) & "," & _
        "#" & Rst("CreateDate") & "#," & _
        "#" & Rst("Modifieddate") & "#," & _
        "#" & Rst("ClosedDate") & "# ," & _
        "#" & Rst("MaturityDate") & "# ," & _
        Rst("PigmyAmount") & "," & _
        AddQuotes(FormatField(Rst("PigmyType")), True) & "," & _
        Rst("RateOfInterest") & "," & _
        AddQuotes(FormatField(Rst("Nominee")), True) & "," & _
        Rst("Introduced") & "," & _
        AddQuotes(FormatField(Rst("LedgerNo")), True) & "," & _
        AddQuotes(FormatField(Rst("FolioNo")), True) & " ," & _
        "0 ," & _
        "1 )"
        
    NewPDTrans.SQLStmt = SqlStr
    If Not NewPDTrans.SQLExecute Then
        NewPDTrans.RollBack
        'Now Check related Customer Is Missed
        NewPDTrans.SQLStmt = "SELECT * From NameTab Where CustomerID = " & Rst("CustomerID")
        If NewPDTrans.SQLFetch = 0 Then GoTo NextAccount
        NewPDTrans.SQLStmt = "SELECT * From UserTab Where UserID = " & Rst("UserID")
        If NewPDTrans.SQLFetch = 0 Then GoTo NextAccount
        
        MsgBox "Unable to transafer the pigmy MAster data"
        Exit Function
    End If
    NewPDTrans.CommitTrans
    
NextAccount:
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
    

Exit_Line:
TransferPDMaster = True
    Debug.Print Now
Exit Function

Err_Line:
    If Err Then
        MsgBox "eror In SBMaster " & Err.Description
        'Resume
    End If
    
End Function



'this function is used to transfer the
'SB MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferPDLoanMaster(OldPdTrans As clsTransact, NewPDTrans As clsTransact) As Boolean
Dim SqlStr As String
Dim Sql_2 As String
Dim Rst As Recordset
Dim DepositType As wis_DepositType

On Error GoTo Err_Line

'Now TransFer the Loan Details
'Get the account Having Details
SqlStr = "SELECT A.UserId,A.AccID,CustomerID,MaturityDate," & _
    " TransDate,RateOfInterest,Amount,TransID,LedgerNo,FolioNo " & _
    " FROM PDTrans A, PDMAster B WHERE A.UserId = B.UserID " & _
    " AND A.AccID = B.AccID And TransId = (SELECT Min(TransID) " & _
        " From PDTrans C Where C.USerID = B.UserID " & _
        " AND C.AccID = B.accID AND Loan = True ) AND Loan = TRue "

OldPdTrans.SQLStmt = SqlStr
If OldPdTrans.SQLFetch <= 0 Then GoTo ExitLine
Set Rst = OldPdTrans.Rst.Clone

Dim AccId As Long
Dim LoanId As Long
Dim LoanNum As String
Dim AccNum As String

'Get the Account Offset From the oLddataBase

NewPDTrans.SQLStmt = "DELETE * FROM DepositLoanMaster Where DepositType = " & wisDeposit_PD
NewPDTrans.BeginTrans
Call NewPDTrans.SQLExecute
NewPDTrans.CommitTrans

NewPDTrans.SQLStmt = "SELECT MAX(LOanId) FROM DepositLoanMaster"
If NewPDTrans.SQLFetch Then m_LoanOffSet = FormatField(NewPDTrans.Rst(0))

SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldPdTrans.SQLStmt = SqlStr
Call OldPdTrans.SQLFetch
'm_AccOffSet= FormatField(OldPdTrans.Rst(0))
'm_AccOffSet= AccOffSet + 100 - (AccOffSet Mod 100)

LoanId = m_LoanOffSet
AccId = m_AccOffSet

While Not Rst.EOF
    AccNum = Rst("UserID") & "_" & Rst("AccId")
    LoanNum = Rst("UserID") & "_" & Rst("AccId")
    'LoanNum = Format(Rst("Accid"), "000")
    LoanId = LoanId + 1
    'Get the Account ID
    NewPDTrans.BeginTrans
    
    
'''' Here  we are inserting into the new table called DepositLoanMAster
''the above said table will be common for all type deposit(eg. FD,Rd,Pd,DL)
    DepositType = wisDeposit_PD
    SqlStr = "Insert INTO PledgeDeposit (" & _
        "LoanID,AccID,DepositType,PledgeNum)" & _
        " VALUES (" & _
        LoanId & "," & _
        AccId & "," & _
        DepositType & "," & _
        " 1 )"
            
    NewPDTrans.SQLStmt = SqlStr
    If Not NewPDTrans.SQLExecute Then
        NewPDTrans.RollBack
        MsgBox "Unable to transafer the pigmy MAster data"
        Exit Function
    End If
    
    SqlStr = "Insert INTO DepositLoanMASTER (" & _
        " CustomerID,LoanID,LoanAccNo,DepositType," & _
        " LoanIssuedate,LoanDueDate,PledgeDescription, " & _
        " InterestRate,LoanAmount,LedgerNo,FolioNo ," & _
        " LastPrintId )"
    SqlStr = SqlStr & " VALUES (" & _
        Rst("CustomerID") & "," & LoanId & "," & _
        AddQuotes(LoanNum, True) & "," & _
        DepositType & "," & _
        "#" & Rst("TransDate") & "#, #" & Rst("MaturityDate") & "#," & _
        AccId & " ," & _
        Rst("RateOfInterest") & "," & _
        Rst("AMount") & "," & _
        "'" & Rst("LedgerNo") & "'," & _
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
    Rst.MoveNext
Wend

ExitLine:
TransferPDLoanMaster = True

    Debug.Print Now
Exit Function

Err_Line:

    If Err Then MsgBox "eror In Pigmy LoanMaster " & Err.Description
    'Resume
End Function


'this function is used to transfer the
'pigmy transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferPDTrans(OldTrans As clsTransact, NewTrans As clsTransact) As Boolean
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
Dim Rst As Recordset
Dim AccId As Long

Dim OldAccId As Long
Dim OldUserID As Long

Dim Amount As Currency
Dim TransDate As Date
    'Fetch the detials of Sb Account

SqlStr = "SELECT * FROM PDTrans Where Loan = False " & _
    " ORDER BY UserID,AccID,TransId"

OldTrans.SQLStmt = SqlStr
If OldTrans.SQLFetch < 1 Then GoTo ExitLine

Set Rst = OldTrans.Rst.Clone

'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldTrans.SQLStmt = SqlStr
Call OldTrans.SQLFetch
'm_AccOffSet= FormatField(OldTrans.Rst(0))
'm_AccOffSet = AccOffSet + 100 - (AccOffSet Mod 100)


TransID = 10000000: AccId = m_AccOffSet

While Not Rst.EOF
    SqlInt = "": SqlPayable = ""
    IsIntTrans = False: IsPaybleTrans = False
    OldTransType = FormatField(Rst("TransType"))
    If OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Then IsIntTrans = True
    If OldTransType = -4 Or OldTransType = 4 Then IsPaybleTrans = True
    If OldTransType = -5 Or OldTransType = 5 Then IsPaybleTrans = True
    
    'If the last record's Transaction id is greater or equal to present transid then
    'It means that the account no has been changed
    If TransID >= FormatField(Rst("TransID")) Then
        AccId = AccId + 1
        PayableBalance = 0
        IntBalance = 0
        TransID = Rst("TransID") - 1
    End If
    
    TransID = TransID + 1
    Amount = Rst("Amount")
    OldAccId = Rst("accId")
    OldUserID = Rst("UserID")
                
    NewTransType = OldTransType / Abs(OldTransType)
    If IsPaybleTrans Then
        If OldTransType = 4 Then NewTransType = wContraDeposit
        If OldTransType = 5 Then NewTransType = wContraWithdraw
        'The above transactions also effect the profit & loss
        If OldTransType = -5 Then NewTransType = wContraDeposit
        If OldTransType = -4 Then NewTransType = wDeposit
        TransDate = Rst("TransDate")
        PayableBalance = PayableBalance + Rst("Amount") * IIf(NewTransType > 0, 1, -1)
        If PayableBalance < 0 Then PayableBalance = 0
        
        If OldTransType = 4 Then NewTransType = wContraDeposit
        If OldTransType = 5 Then NewTransType = wWithdraw
        SqlPayable = "Insert INTO PDIntPayable ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
        SqlPayable = SqlPayable & "VALUES (" & _
            AccId & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            PayableBalance & " ," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
            NewTransType & " )"
        
        If OldTransType = 4 Then
            NewTransType = wContraWithdraw
            IntBalance = IntBalance + Rst("Amount")
            SqlInt = "Insert INTO PDIntTrans ( " & _
                "AccID,TransID,TransDate," & _
                "Amount,Balance,Particulars," & _
                "TransType )"
            SqlInt = SqlInt & "VALUES (" & _
                AccId & "," & _
                TransID & "," & _
                "#" & Rst("TransDate") & "#," & _
                Rst("Amount") & "," & _
                IntBalance & " ," & _
                AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
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
        If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithdraw
        If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit
        TransDate = Rst("TransDate")
        IntBalance = IntBalance + Rst("Amount")
        SqlInt = "Insert INTO PDIntTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            AccId & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0  ," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
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
    
    NewTransType = (OldTransType / Abs(OldTransType))
    SqlStr = "Insert INTO PDTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
    SqlStr = SqlStr & "VALUES (" & _
            AccId & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            Rst("Balance") & "," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
            NewTransType & " )"
    
    If Balance = FormatField(Rst("Balance")) Then SqlStr = ""
    If OldAccId <> Rst("AccId") Or OldUserID <> Rst("UserId") Then SqlStr = ""
    If Rst("Amount") = 0 Then SqlStr = ""
    
    NewTrans.BeginTrans
    If SqlStr <> "" Then
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy trans data"
            NewTrans.RollBack
            Exit Function
        End If
    End If
    If SqlInt <> "" Then
        NewTrans.SQLStmt = SqlInt
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy Trans data"
            NewTrans.RollBack
            Exit Function
        End If
    End If
    If SqlPayable <> "" Then
        NewTrans.SQLStmt = SqlPayable
        If Not NewTrans.SQLExecute Then
            MsgBox "Unable to transafer the pigmy Trans data"
            NewTrans.RollBack
            Exit Function
        End If
    End If
    
    NewTrans.CommitTrans
    Balance = FormatField(Rst("Balance"))

NextAccount:
    Rst.MoveNext

Wend
Debug.Print Now & "  " & Rst.RecordCount

ExitLine:
TransferPDTrans = True
Exit Function


Err_Line:
    If Err Then
        MsgBox "Error in PDTrans" & Err.Description
        'Resume
    End If
    
End Function


'this function is used to transfer the
'pigmy Loan transaction details form OLd Db to new one
'and newtrans has assigned to new database
Private Function TransferPDLoanTrans(OldTrans As clsTransact, NewTrans As clsTransact) As Boolean
Dim SqlStr As String
Dim SqlInt As String
Dim IsIntTrans As Boolean

Dim OldUserID As Long
Dim OldAccId As Long

On Error GoTo Err_Line

Dim OldTransType As wisTransactionTypes, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim Rst As Recordset
Dim LoanId As Long
Dim Amount As Currency
Dim TransDate As Date
Dim DepositType As wis_DepositType

DepositType = wisDeposit_PD
    
'Fetch the detials of pigmy Account

SqlStr = "SELECT * FROM PDTrans Where Loan = True " & _
    " ORDER BY UserID,AccID,TransId"

OldTrans.SQLStmt = SqlStr
If OldTrans.SQLFetch < 1 Then GoTo ExitLine
Set Rst = OldTrans.Rst.Clone

'Get the Account Offset From the oLddataBase
SqlStr = "SELECT Max(AccID) FROM PDMASTER "
OldTrans.SQLStmt = SqlStr
Call OldTrans.SQLFetch
m_AccOffSet = FormatField(OldTrans.Rst(0))
m_AccOffSet = m_AccOffSet + 100 - (m_AccOffSet Mod 100)

TransID = 100000
LoanId = m_LoanOffSet
Balance = 0
While Not Rst.EOF
    IsIntTrans = False: SqlInt = ""
    OldTransType = FormatField(Rst("TransType"))
    
    If OldTransType = 4 Or OldTransType = 2 Then IsIntTrans = True
    If OldTransType = -2 Or OldTransType = -4 Then IsIntTrans = True
    
    If TransID >= Rst("TransID") Then LoanId = LoanId + 1: TransID = Rst("TransID") - 1
    TransID = TransID + 1
    Amount = Rst("Amount")
    OldUserID = Rst("UserId")
    OldAccId = Rst("AccID")
    
    NewTransType = OldTransType / Abs(OldTransType)
    If IsIntTrans Then
        If OldTransType = 2 Or OldTransType = 4 Then NewTransType = wWithdraw  'interest Paid to the customer
        If OldTransType = -2 Or OldTransType = -4 Then NewTransType = wDeposit  'Interest collected form the customer
        TransDate = Rst("TransDate")
        SqlInt = "Insert INTO DepositLoanIntTrans ( " & _
            "LoanId,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
        SqlInt = SqlInt & "VALUES (" & _
            LoanId & "," & _
            TransID & "," & _
            "#" & Rst("TransDate") & "#," & _
            Rst("Amount") & "," & _
            "0  ," & _
            AddQuotes(FormatField(Rst("Particulars")), True) & "," & _
            NewTransType & " )"
        If Amount = 0 Then SqlInt = ""
        Rst.MoveNext
        Amount = Rst("Amount")
    End If
    
    OldTransType = FormatField(Rst("TransType"))
    NewTransType = (OldTransType / Abs(OldTransType))
    
    ''INSERT INTO DEPOSITLOANTRANS
    SqlStr = "Insert INTO DEPOSITLOanTrans ( " & _
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
        NewTransType & " )"
    
    If OldAccId = Rst("AccID") Or OldUserID <> Rst("UserID") Then SqlStr = ""
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
    Rst.MoveNext
Wend

ExitLine:
TransferPDLoanTrans = True
Exit Function


Err_Line:
    If Err Then
        MsgBox "Error in PD LOAn Trans" & Err.Description
        'Resume
    End If
    
End Function

'this function is used to transfer the
'set the voucher no fot the transaferred data
Private Function PutVoucherNumber(NewTrans As clsTransact) As Boolean
Dim SqlStr As String



End Function
