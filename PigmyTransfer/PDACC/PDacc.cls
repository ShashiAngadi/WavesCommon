VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This Function will withdraw the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class

Private m_rstCustBalance As ADODB.Recordset
Private WithEvents m_frmPDAcc As frmPDAcc
Attribute m_frmPDAcc.VB_VarHelpID = -1
Private m_AccID As Long
Private m_AgentID As Integer
Private WithEvents m_frmPDReport As frmPDReport
Attribute m_frmPDReport.VB_VarHelpID = -1

Private m_ParentHandle As Long

Public Property Get ObjectName()
    ObjectName = "PDClass"
End Property


Public Property Let AccountId(NewValue As Long)
m_AccID = NewValue

End Property


Public Property Let CustomerID(NewValue As Long)
Dim Rst As Recordset

gDbTrans.SQLStmt = "Select AccID From PDMAster " & _
    "Where CustomerID = " & NewValue
If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Property

m_AccID = FormatField(Rst("AccID"))
Set Rst = Nothing


End Property


Public Function InterestAmount(ByVal AccountId As Long, ByVal AsOnDate As Date)

InterestAmount = ComputePDInterestAmount(AccountId, AsOnDate)

End Function

Public Function WithdrawAmount(AccID As Long, _
                Amount As Currency, Description As String, TransDate As Date, _
                Optional VoucherNo As String = " ", _
                Optional CashTransaction As Boolean = False) As Integer
    
Dim Rst As ADODB.Recordset
Dim Loan As Boolean

'Check Whether Account Exist or Closed
Loan = False
gDbTrans.SQLStmt = "Select TransId,TransDate,Balance From PDtrans " & _
                " Where AccId = " & AccID & " And TransId = " & _
                " (SELECT MAX(TransID) From PDTrans" & _
                    " WHERE Accid = " & AccID & ")"

If gDbTrans.Fetch(Rst, adOpenStatic) < 1 Then GoTo ErrLine
' Check fotr the Tranasction date
If DateDiff("d", TransDate, Rst("Transdate")) > 0 Then GoTo ErrLine

Dim TransID As Long
Dim Balance As Currency
Dim Particulars As String

'Assign The Necessary Values
TransID = Val(FormatField(Rst(0))) + 1
Balance = Val(FormatField(Rst("Balance")))
Particulars = Description

If Balance - Amount < 1 Then GoTo ErrLine
Dim TransType As wisTransactionTypes


'gDbTrans.BeginTrans
TransType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

Balance = Balance - Amount
Loan = False

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SQLStmt = "INSERT INTO PDTrans " & _
            " (Accid ,TransId ,Amount ,Balance," & _
            " TransType, Particulars,TransDate)" & _
            " VALUES(" & AccID & "," & TransID & "," & _
            Amount & "," & Balance & "," & TransType & "," & _
            AddQuotes(Particulars, True) & ",#" & TransDate & "# )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

Dim ClsBank As clsBankAcc
Dim AccHeadID As Long

Set ClsBank = New clsBankAcc
AccHeadID = ClsBank.GetHeadIDCreated(LoadResString(gLangOffSet + 425))
Set ClsBank = Nothing
If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID + 1
    
    gDbTrans.SQLStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccID," & _
        "TransType, TransID,Amount,VoucherNo)" & _
        "Values (" & ContraID & "," & AccHeadID & "," & _
        AccID & ", " & TransType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

WithdrawAmount = TransID


Exit Function
ErrLine:

End Function

Public Function CustomerBalance(ByVal CustomerID As Long, _
        retAccNo As String, ByVal AsOnDate As Date) As Currency

Dim AccID As Long
Dim AccNum As Long

'To trap an error
On Error GoTo ErrLine

CustomerBalance = 0
retAccNo = ""

'Sql stmt to get Accid
gDbTrans.SQLStmt = "Select A.Accid,AccNum,Balance from " & _
        " PDMaster A,PDTrans B where Customerid=" & CustomerID & _
        " AND A.AccID = B.AccID And TransID = " & _
            "(SELECT Max(TransID) From PDTrans C Where " & _
            " C.AccID = A.AccID ANd TransDate <= #" & AsOnDate & "# )"
If gDbTrans.Fetch(m_rstCustBalance, adOpenStatic) < 1 Then
    AccID = 0
    Exit Function
End If

AccID = FormatField(m_rstCustBalance("AccId"))
retAccNo = FormatField(m_rstCustBalance("AccNum"))
CustomerBalance = FormatField(m_rstCustBalance("Balance"))

ExitLine:
    Exit Function

ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        Err.Clear
        GoTo ExitLine
    End If
End Function
Public Function CashDeposits(FromDate As Date, ToDate As Date) As Currency

Dim TransType As wisTransactionTypes
CashDeposits = 0

TransType = wDeposit

CashDeposits = DepositTransactionDetails(FromDate, ToDate, TransType)

End Function

Public Function CashWithdrawls(FromDate As Date, ToDate As Date) As Currency

Dim TransType As wisTransactionTypes
CashWithdrawls = 0

TransType = wWithdraw

CashWithdrawls = DepositTransactionDetails(FromDate, ToDate, TransType)

End Function

Public Function ContraDeposits(FromDate As Date, ToDate As Date) As Currency

Dim TransType As wisTransactionTypes
ContraDeposits = 0

TransType = wContraDeposit

ContraDeposits = DepositTransactionDetails(FromDate, ToDate, TransType)

End Function

Public Function ContraWithdrawls(FromDate As Date, ToDate As Date) As Currency

Dim TransType As wisTransactionTypes
ContraWithdrawls = 0

TransType = wContraWithdraw

ContraWithdrawls = DepositTransactionDetails(FromDate, ToDate, TransType)

End Function

Public Function DepositAgents(FromDate As Date, ToDate As Date, Optional AgentID As Long) As Currency

Dim Rst As ADODB.Recordset
Dim SqlStr As String
Dim Trans As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

On Error GoTo ErrLine
DepositAgents = 0
Trans = wDeposit
ContraTrans = wContraDeposit

SqlStr = "SELECT Sum(Amount) as Deposit From AgentTrans " & _
    " WHERE TransDate >= #" & FromDate & "#" & _
    " AND TransDate <= #" & ToDate & "#" & _
    " AND (TransType = " & Trans & " OR TransType = " & ContraTrans & " )"

If AgentID Then SqlStr = SqlStr & " AND AgentID = " & AgentID

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(Rst, adOpenStatic) > 0 Then DepositAgents = FormatField(Rst(0))

Exit Function

ErrLine:
    If Err Then
        MsgBox "Error in AgentTrans " & Err.Number & vbCrLf & Err.Description _
            , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If

End Function
'This Function will deposit the Amount to The Speicfied Agent
'And It put the particulars as deposited notes
'This function Can Be Used in Cash Counter only'
Public Function DepositAgentAmount(ByVal AgentID As Long, ByVal Amount As Currency, _
    ByVal Description As String, ByVal TransDate As Date, _
    Optional VoucherNo As String = " ", _
    Optional CashTransaction As Boolean) As Integer

'Check Whether Account Exist or Closed
Dim RstDep As Recordset

'Now Get the Max TransID,Balance On specified  Date
'Then compare it With Max(TransID)
gDbTrans.SQLStmt = "Select Top 1 * " & _
    " From AgentTrans Where AgentId = " & AgentID & _
    " ORDER By TransID Desc"

If gDbTrans.Fetch(RstDep, adOpenStatic) < 1 Then GoTo ErrLine
'Check for the Last Transcated date

If DateDiff("d", TransDate, RstDep("TransDate")) > 0 Then GoTo ErrLine

Dim TransID As Long
Dim Balance As Currency
Dim TransType As wisTransactionTypes
Dim UserID As Long
Dim Rst As ADODB.Recordset

UserID = gCurrUser.UserID

Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(LoadResString(gLangOffSet + 425))

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID
End If

'Assign The Necessary Values
TransID = Val(FormatField(RstDep("TransID")))
Balance = Val(FormatField(RstDep("Balance")))

TransType = IIf(CashTransaction, wDeposit, wContraDeposit)

Balance = Balance + Amount
TransID = TransID + 1

If IsMissing(VoucherNo) Then VoucherNo = " "

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SQLStmt = "INSERT INTO AgentTrans " & _
    "(AgentId,TransId,Amount,Balance," & _
    "TransType,TransDate,VoucherNo,UserID) " & _
    "VALUES(" & AgentID & "," & TransID & "," & _
    Amount & "," & Balance & "," & TransType & "," & _
    " #" & TransDate & "# ," & _
    AddQuotes(VoucherNo, True) & ", " & UserID & ")"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

If Not CashTransaction Then
    'update it to the Contra Trans table
    gDbTrans.SQLStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccId," & _
        "TransType, TransID,Amount,VoucherNo)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        AgentID & ", " & TransType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

DepositAgentAmount = TransID

Exit Function

ErrLine:

End Function

'This Function will deposit the Amount to The Speicfied Account No
'And It put the particulars as deposited notes
'This function Can Be Used in Standing Instuctions
'Presently this is Using by Clearing Class and Cashier class
Public Function DepositAmount(ByVal AccID As Long, ByVal Amount As Currency, _
                    ByVal Description As String, ByVal TransDate As Date, _
                    Optional VoucherNo As String = " ", _
                    Optional CashTransaction As Boolean = False) As Integer

'Check Whether Account Exist or Closed
Dim RstDep As Recordset

'Now Get the Max TransID,Balance On specified  Date
'Then compare it With Max(TransID)
gDbTrans.SQLStmt = "Select Top 1 TransID,Balance,TransDate " & _
            " From PDTrans Where AccId = " & AccID & _
            " ORDER By TransID Desc"

If gDbTrans.Fetch(RstDep, adOpenStatic) < 1 Then GoTo ErrLine
'Check for the Last Transcated date

If DateDiff("d", TransDate, RstDep("TransDate")) > 0 Then GoTo ErrLine

Dim TransID As Long
Dim Balance As Currency
Dim TransType As wisTransactionTypes
Dim UserID As Long
Dim Rst As ADODB.Recordset

UserID = gCurrUser.UserID

Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(LoadResString(gLangOffSet + 425))

'Assign The Necessary Values
TransID = Val(FormatField(RstDep("TransID")))
Balance = Val(FormatField(RstDep("Balance")))

gDbTrans.SQLStmt = "Select Top 1 TransID,Balance,TransDate " & _
            " From PDIntTrans Where AccId = " & AccID & _
            " ORDER By TransID Desc"
If gDbTrans.Fetch(RstDep, adOpenStatic) > 0 Then
    TransID = IIf(TransID > RstDep("TransID"), TransID, RstDep("TransID"))
    'Check for the Last Transcated date
    If DateDiff("d", TransDate, RstDep("TransDate")) > 0 Then GoTo ErrLine
End If

TransType = IIf(CashTransaction, wDeposit, wContraDeposit)

Balance = Balance + Amount
If IsMissing(VoucherNo) Then VoucherNo = " "

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SQLStmt = "INSERT INTO PDTrans " & _
        "(Accid,TransId,Amount,Balance," & _
        "TransType,Particulars,TransDate,VoucherNo,UserID) " & _
        "VALUES(" & AccID & "," & TransID & "," & _
        Amount & "," & Balance & "," & TransType & "," & _
        AddQuotes(Description, True) & "," & _
        " #" & TransDate & "# ," & _
        AddQuotes(VoucherNo, True) & ", " & UserID & ")"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID
    'update it to te Contra Trans table
    gDbTrans.SQLStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccId," & _
        "TransType, TransID,Amount,VoucherNo)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        AccID & ", " & TransType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

DepositAmount = TransID

Exit Function

ErrLine:

End Function

'
Public Function Deposits(FromIndianDate As String, ToIndianDate As String) As Currency
Dim FromDate As Date
Dim ToDate As Date
FromDate = GetSysFormatDate(FromIndianDate)
ToDate = GetSysFormatDate(ToIndianDate)

' Total Fixed Deposit flown in the bank between 2 given dates

Dim TransType As wisTransactionTypes
Dim Amount As Currency
Deposits = 0


TransType = wDeposit
Amount = DepositTransactionDetails(FromDate, ToDate, TransType)
TransType = wContraDeposit
Amount = Amount + DepositTransactionDetails(FromDate, ToDate, TransType)

Deposits = Amount
End Function

Private Function DepositPLTransactionDetails(FromDate As Date, ToDate As Date, _
    TransType As wisTransactionTypes) As Currency


Dim SqlStr As String
Dim Rst As ADODB.Recordset

DepositPLTransactionDetails = 0

SqlStr = " Select Sum(Amount) as Deposits from PDTrans where " & _
    " TransType = " & TransType & _
    " And TransDate >= #" & FromDate & "# " & _
    " And TransDate <= #" & ToDate & "#"
gDbTrans.SQLStmt = SqlStr

If gDbTrans.Fetch(Rst, adOpenStatic) < 1 Then Exit Function
DepositPLTransactionDetails = Val(FormatField(Rst(0)))

End Function
Private Function DepositTransactionDetails(FromDate As Date, ToDate As Date, _
    TransType As wisTransactionTypes) As Currency
Dim SqlStr As String
Dim Rst As ADODB.Recordset

DepositTransactionDetails = 0

SqlStr = " Select Sum(Amount) as Deposits from PDTrans where " & _
    " TransType = " & TransType & _
    " And TransDate >= #" & FromDate & "# " & _
    " And TransDate <= #" & ToDate & "#"
gDbTrans.SQLStmt = SqlStr

If gDbTrans.Fetch(Rst, adOpenStatic) < 1 Then Exit Function
DepositTransactionDetails = Val(FormatField(Rst(0)))

End Function

Public Function InterestOnDeposits(AsonIndianDate As String) As Currency
   On Error GoTo ErrLine
  'declare the variables necessary
Dim rstMain As ADODB.Recordset
  
Dim Retval As Long
Dim AccID As Long
Dim InterestRate As Currency
Dim LastIntDate As String
Dim CreateDate As String
Dim MatDate As String
Dim Duration As Long
Dim IntAmount As Currency
Dim Balance As Currency
Dim Loan As Boolean

Dim TotalAmount As Currency

'initialize some variables
Loan = False

'Build The Querry
gDbTrans.SQLStmt = " SELECT Balance, A.AccId, A.UserID, " & _
        " CreateDate, MaturityDate, TransDate, RateOfInterest" & _
        " From PDMaster A, PDTrans B Where A.AccId = B.AccId" & _
        " And (ClosedDate is NULL OR ClosedDate = #1/1/100#) " & _
        " And CreateDate < #" & GetSysFormatDate(AsonIndianDate) & "# And TransID =  " & _
            " ( Select max(TransID) From PDTrans C Where A.AccId = C.AccId  " & _
            " and TransDate <= #" & GetSysFormatDate(AsonIndianDate) & "# ) " & _
        " Order By A.UserID, A.AccId "

Retval = gDbTrans.Fetch(rstMain, adOpenStatic)
If Retval < 0 Then GoTo ErrLine
If Retval < 1 Then GoTo ExitLine


While Not rstMain.EOF
    AccID = Val(FormatField(rstMain("AccId")))
    Balance = CCur(FormatField(rstMain("Balance")))
    LastIntDate = FormatField(rstMain("TransDate"))
    CreateDate = FormatField(rstMain("CreateDate"))
    InterestRate = CCur(FormatField(rstMain("RateofInterest")))
    MatDate = FormatField(rstMain("MaturityDate"))
    
    'Now Get The Date Difference
    ' On that Difference Get the InterestRate
    Duration = WisDateDiff(LastIntDate, AsonIndianDate)
    If Duration <= 0 Then GoTo NextDeposit
    
    If InterestRate = 0 Then
        InterestRate = GetPDDepositInterest(CInt(Duration), CreateDate)
    End If
    
    Duration = WisDateDiff(LastIntDate, AsonIndianDate)
    If Duration <= 0 Then GoTo NextDeposit
    IntAmount = ((InterestRate / 100) * Balance * Duration) / 365
    TotalAmount = CCur(TotalAmount + IntAmount)
    
NextDeposit:
    rstMain.MoveNext
Wend
InterestOnDeposits = TotalAmount
GoTo ExitLine


ErrLine:
MsgBox "Error In PDClass--Interest On Deposits", vbCritical, wis_MESSAGE_TITLE
'Resume

ExitLine:
End Function


Public Function Liability(AsonIndianDate As String) As Currency
' Computes the Liability (Deposit plus interest on Deposit)

Dim SetUp As New clsSetup
Dim Loan As Boolean
Dim LiabBeforeInterest As Currency
Dim InterestOnDeposit As Double

Loan = False
Liability = 0

If Not DateValidate(AsonIndianDate, "/", True) Then Exit Function

LiabBeforeInterest = ComputeTotalPDLiability(AsonIndianDate)

InterestOnDeposit = CDbl(SetUp.ReadSetupValue("PDacc", "Interest On PDDeposit", 12#))

Liability = ComputePDInterest(LiabBeforeInterest, InterestOnDeposit)

End Function

Public Sub Show()
Dim Perms As wis_Permissions
Dim StartPD As Boolean

StartPD = False

If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers: Exit Sub
    
    'Now Check the permission
    'whether logged inperson is having permission of sb account or not
    Perms = gCurrUser.UserPermissions
    If Perms And perBankAdmin Then StartPD = True
    If Perms And perCreateAccount Then StartPD = True
    If Perms And perClerk Then StartPD = True
    If Perms And perCashier Then StartPD = True
    If Perms And perReadOnly Then StartPD = True
    If Perms And perPigmyAgent Then StartPD = True
    If StartPD Then
        If m_frmPDAcc Is Nothing Then Set m_frmPDAcc = New frmPDAcc
        Load m_frmPDAcc
        If gWindowHandle Then m_ParentHandle = gWindowHandle
        gWindowHandle = m_frmPDAcc.hwnd
        m_frmPDAcc.Show
        If m_AccID Then m_frmPDAcc.AccountLoad (m_AccID)
    Else
        MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
    End If
    
    
End Sub

'This Function Will undo the transaction of the Given
'Account ID and the transaction ID
'This Help Full In Case Of The Contras Transaction Has to UNdo
Public Function UndoAgentTransaction(AgentID As Long, TransID As Long) As Boolean

Dim TransDate As String
Dim Amount As Currency
Dim LastTransId As Long
Dim SQLStmt As String
Dim LoanBalance As Currency
Dim TransType As wisTransactionTypes

'Get last transaction record
Dim Rst As Recordset
gDbTrans.SQLStmt = "Select TOP 1 * from AgentTrans where " & _
            " AgentID = " & AgentID & " Order by TransID desc"
        
If gDbTrans.Fetch(Rst, adOpenStatic) < 1 Then Exit Function

Amount = FormatField(Rst.Fields("Balance"))
LastTransId = FormatField(Rst.Fields("TransID"))
TransType = FormatField(Rst("TransType"))
TransDate = FormatField(Rst.Fields("TransDate"))
        
If LastTransId <> TransID Then
    MsgBox LoadResString(gLangOffSet + 530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Delete record from Data base
gDbTrans.SQLStmt = "Delete * from AgentTrans where " & _
        " AgentID = " & AgentID & " and TransID = " & LastTransId
If Not gDbTrans.SQLExecute Then Exit Function
 
UndoAgentTransaction = True

End Function

'This Function Will undo the transaction of the Given
'Account ID and the transaction ID
'This Help Full In Case Of The Contras Transaction Has to UNdo
Public Function UndoTransaction(AccID As Long, TransID As Long) As Boolean

Dim TransDate As String
Dim Amount As Currency
Dim LastTransId As Long
Dim SQLStmt As String
Dim LoanBalance As Currency
Dim TransType As wisTransactionTypes

'Get last transaction record
Dim Rst As Recordset
gDbTrans.SQLStmt = "Select TOP 1 * from PDTrans where " & _
        " AccID = " & AccID & " Order by TransID desc"
        
If gDbTrans.Fetch(Rst, adOpenStatic) < 1 Then Exit Function


Amount = FormatField(Rst.Fields("Balance"))
LastTransId = FormatField(Rst.Fields("TransID"))
TransType = FormatField(Rst("TransType"))
TransDate = FormatField(Rst.Fields("TransDate"))
        
If LastTransId <> TransID Then
    MsgBox LoadResString(gLangOffSet + 530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Delete record from Data base
    gDbTrans.SQLStmt = "Delete from PDTrans where " & _
        " AccID = " & AccID & " and TransID = " & LastTransId
    If Not gDbTrans.SQLExecute Then Exit Function
    
    If LastTransId = 100 Then
        If MsgBox(LoadResString(gLangOffSet + 539) & "Do you Want To Delete This " & _
            "Account Permanently ?", vbInformation + vbYesNo + vbDefaultButton2, _
            "Undo Last") = vbYes Then
            gDbTrans.SQLStmt = "Delete from PDMaster where AccID = " & AccID
            If Not gDbTrans.SQLExecute Then Exit Function
            
        End If
    End If


UndoTransaction = True

End Function


Public Function Loss(FromIndianDate As String, ToIndianDate As String) As Currency
Dim FromDate As Date
Dim ToDate As Date
FromDate = GetSysFormatDate(FromIndianDate)
ToDate = GetSysFormatDate(ToIndianDate)

Loss = 0

Dim TransType As wisTransactionTypes
Dim Amount As Currency
TransType = wWithdraw
Amount = DepositPLTransactionDetails(FromDate, ToDate, TransType)
TransType = wContraWithdraw
Amount = Amount + DepositPLTransactionDetails(FromDate, ToDate, TransType)

Loss = Amount

End Function

'THis Function Retuns the Next Account Blance of the customewr
'Before calling this function, we should call the CustomerBlance function
Public Function NextBalance(retAccNum As String) As Currency

'To trap an error
On Error GoTo ErrLine

NextBalance = 0
retAccNum = ""

If m_rstCustBalance Is Nothing Then GoTo ExitLine
If m_rstCustBalance.EOF Then GoTo ExitLine
m_rstCustBalance.MoveNext
If m_rstCustBalance.EOF Then GoTo ExitLine

retAccNum = Val(FormatField(m_rstCustBalance("AccNum")))
NextBalance = FormatField(m_rstCustBalance("Balance"))

ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If

End Function


Public Function Profit(FromIndianDate As String, ToIndianDate As String) As Currency
Dim FromDate As Date
Dim ToDate As Date
FromDate = GetSysFormatDate(FromIndianDate)
ToDate = GetSysFormatDate(ToIndianDate)

' profit to bank thru deposits
Profit = 0
Dim TransType As wisTransactionTypes
Dim Amount As Currency

TransType = wDeposit
Amount = DepositPLTransactionDetails(FromDate, ToDate, TransType)
TransType = wContraDeposit
Amount = Amount + DepositPLTransactionDetails(FromDate, ToDate, TransType)

Profit = Amount

End Function

Public Function WithDrawls(FromIndianDate As String, ToIndianDate As String) As Currency
Dim FromDate As Date
Dim ToDate As Date
FromDate = GetSysFormatDate(FromIndianDate)
ToDate = GetSysFormatDate(ToIndianDate)

'Pigmy Deposit Given back to the customer
WithDrawls = 0

Dim TransType As wisTransactionTypes
Dim Amount As Currency

TransType = wWithdraw
Amount = DepositTransactionDetails(FromDate, ToDate, TransType)

TransType = wContraWithdraw
Amount = Amount + DepositTransactionDetails(FromDate, ToDate, TransType)
WithDrawls = Amount

End Function

Public Function Balance(AsonIndianDate As String) As Currency
' fixed deposits in the bank as on the given date
Balance = 0
Balance = ComputeTotalPDLiability(AsonIndianDate)

End Function

Private Sub Class_Terminate()
On Error Resume Next
If Not m_frmPDAcc Is Nothing Then
    Unload m_frmPDAcc
    Set m_frmPDAcc = Nothing
End If

End Sub



Private Sub m_frmPDAcc_AccountChanged(ByVal AccID As Long)
m_AccID = AccID

End Sub

Private Sub m_frmPDAcc_AgentChanged(ByVal AgentID As Integer)
m_AgentID = AgentID
End Sub


Private Sub m_frmPDAcc_ShowReport(ShowAgent As Boolean, _
            ReportType As wis_PDReports, ReportOrder As wis_ReportOrder, _
            FromDate As String, ToDate As String, _
            RepOption As clsRepOption, AgentID As Integer)

If Not m_frmPDReport Is Nothing Then Set m_frmPDReport = Nothing
Set m_frmPDReport = New frmPDReport
    
frmCancel.Show
With m_frmPDReport
    .Caste = RepOption.Caste
    .Place = RepOption.Place
    .Gender = RepOption.Gender
    .FromIndianDate = FromDate
    .ToIndianDate = ToDate
    .FromAmount = RepOption.FromAmount
    .ToAmount = RepOption.ToAmount
    .AgentID = AgentID
    
    .ReportOrder = ReportOrder
    .ReportType = ReportType
    .AccountGroup = RepOption.AccountGroupID
    
    .DisplayAgentName = ShowAgent
    Load m_frmPDReport
    Unload frmCancel
    Screen.MousePointer = vbDefault
    .Show 'vbModal
    
End With


Screen.MousePointer = vbDefault


End Sub

Private Sub m_frmPDAcc_WindowClosed()
If m_ParentHandle Then _
    gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set m_frmPDAcc = Nothing
End Sub


Public Property Get ObjectHandle() As Long
    If Not m_frmPDAcc Is Nothing Then _
            ObjectHandle = m_frmPDAcc.hwnd: Exit Property
    If Not m_frmPDReport Is Nothing Then _
            Unload m_frmPDReport: Set m_frmPDReport = Nothing
            
End Property

Private Sub m_frmPDReport_Initialise(Min As Long, Max As Long)
On Error Resume Next
With frmCancel
    If Max <> 0 And Max > Min Then
        .PicStatus.Visible = True
    End If
End With

End Sub


Private Sub m_frmPDReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS : " & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
End With

End Sub


