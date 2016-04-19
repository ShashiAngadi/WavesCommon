Attribute VB_Name = "basBankAcc"
Option Explicit

Dim m_BankExpHeadID As Long
Dim m_BankIncomeHeadID As Long

Public Const wis_BankHeadOffSet = 1000
Enum wisBankHeads
  wis_BankHead = 1000
  wis_BankLoanHead = 2000
  wis_AdvanceHead = 3000
  wis_InvestmentHead = 4000
  wis_IncomeHead = 5000
  wis_ExpenditureHead = 6000
  wis_TradingIncomeHead = 7000
  wis_TradingExpenditureHead = 8000
  wis_ReserveFundHead = 9000
  wis_ShareCapitalHead = 10000
  wis_GovtLoanSubsidyHead = 11000
  wis_AssetHead = 12000
  wis_PaymentHead = 13000
  wis_RepaymentHead = 14000
  
  'Particularly for Hiewpadasalagi
  'Particularly for Hirepadasalagi
  wis_MemberDeposits = 20000
  
End Enum


Private Function GetNewParentID(Rst As Recordset) As Long
    Dim ParentID As Long
    Dim AccID As Long
    If Rst Is Nothing Then Exit Function
    If Rst.EOF Or Rst.BOF Then Exit Function
    AccID = (Rst("AccID") - (Rst("AccID") Mod 1000))
    Select Case AccID
        Case 1000
            ParentID = parBankAccount
        Case 2000
            ParentID = parBankLoanAccount
        Case 3000
            ParentID = parLoanAdvanceAsset
        Case 4000
            ParentID = parInvestment
        Case 5000
            ParentID = parIncome
            'Rst.Find "AccID > " & AccID + 1
        Case 6000
            ParentID = parExpense
            'Rst.Find "AccID > " & AccID + 1
        Case 7000
            ParentID = parTradingIncome
        Case 8000
            ParentID = parTradingExpense
        Case 9000
            ParentID = parReserveFunds
        Case 10000
            ParentID = parShareCapital
        Case 11000
            ParentID = parGovtLoanSubsidy
        Case 12000
            ParentID = parFixedAsset
        Case 13000
            ParentID = parPayAble
            Rst.Find "AccID > " & AccID + 3
        Case 14000
            ParentID = parReceivable
        'Case 15000
        '    ParentID = parBankAccount
        'Case 16000
         '   ParentID = parBankAccount
        Case 20000
            ParentID = parMemberShare
        Case 21000
            ParentID = parProfitORLoss
        Case Else
            Rst.Find "AccID > " & AccID + 100
      
      End Select
      
    GetNewParentID = ParentID
End Function

Private Function GetNewVoucherType(OldTrans As Integer, _
        NewHeadID As Long, OldParentID As Long, VoucherType As Wis_VoucherTypes) As Wis_VoucherTypes
    
Dim NewTransType As wisTransactionTypes
'Dim VoucherType As Wis_VoucherTypes
Dim ParentID As Long
Dim AccID As Long

    If NewHeadID < wis_CashHeadID Then
        'Laibility Heads consider Only 1 and -1 transtype
        If OldTrans = 1 Then NewTransType = wDeposit
        If OldTrans = -1 Then NewTransType = wWithDraw
        
    End If
    If NewHeadID > wis_CashHeadID And NewHeadID < parIncome Then
        'Asset Heads
        If OldTrans = 1 Then NewTransType = wDeposit
        If OldTrans = -1 Then NewTransType = wWithDraw
    End If
    If NewHeadID > parIncome And NewHeadID < parExpense Then
        'Income Heads
        
    End If
    If NewHeadID > parExpense And NewHeadID < parProfitORLoss Then
        'expense Heads
        
    End If
    
    Select Case OldParentID
        Case 1000
            If OldTrans = 1 Then VoucherType = Receipt
            If OldTrans = -1 Then NewTransType = wWithDraw

        Case 2000
            ParentID = parBankLoanAccount
        Case 3000
            ParentID = parLoanAdvanceAsset
        Case 4000
            ParentID = parInvestment
        Case 5000
            ParentID = parIncome
'            Rst.Find "AccID > " & AccID + 1
        Case 6000
            ParentID = parExpense
'            Rst.Find "AccID > " & AccID + 1
        Case 7000
            ParentID = parTradingIncome
        Case 8000
            ParentID = parTradingExpense
        Case 9000
            ParentID = parReserveFunds
        Case 10000
            ParentID = parShareCapital
        Case 11000
            ParentID = parGovtLoanSubsidy
        Case 12000
            ParentID = parFixedAsset
        Case 13000
            ParentID = parPayAble
'            rstMain.Find "AccID > " & AccID + 3
        Case 14000
            ParentID = parReceivable
        'Case 15000
        '    ParentID = parBankAccount
        'Case 16000
         '   ParentID = parBankAccount
        Case 20000
            ParentID = parMemberShare
        Case 21000
            ParentID = parProfitORLoss
        Case Else
'            rstMain.Find "AccID > " & AccID + 100
      
    End Select

    
    
End Function

'just calling this function we can transafer the sbmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferBank(OldDBName As String, NewDBName As String) As Boolean
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
    If Not TransferBankMaster(OldTrans, NewTrans) Then GoTo ErrLine
End If
'    Call NewTrans.WISCompactDB(NewDBName, newpwd, newpwd)
    If Not TransferBankTrans(OldTrans, NewTrans) Then GoTo ErrLine
    'If Not CreateSavingsHead(NewTrans) Then GoTo ErrLine
    TransferBank = True
    Call NewTrans.WISCompactDB(NewDBName, NewPwd, NewPwd)
ErrLine:

OldTrans.CloseDB
NewTrans.CloseDB
Set OldTrans = Nothing
Set NewTrans = Nothing

Screen.MousePointer = vbNormal

End Function

'this function is used to transfer the
'Bank MAster details form OLdb to new one
'and NewTrans has assigned to new database
Private Function TransferBankMaster(OldTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim AccID As Long
Dim rstMain As ADODB.Recordset
Dim Rst As ADODB.Recordset
Dim ParentID As Long
Dim i As Integer
Dim FromDate As Date

On Error GoTo Err_Line

FromDate = FormatDate(frmMain.txtDate)

'Fetch the detials of Sb Account
SqlStr = "SELECT * FROM AccMaster ORDER BY AccID"

OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(rstMain, adOpenForwardOnly) < 1 Then GoTo Exit_line

With frmMain
    .lblProgress = "Transferring Bank Account details"
    .prg.Max = rstMain.RecordCount + 1
    .prg.Value = 0
    .Refresh
End With

Dim NewHeadID As Long
Dim Balance As Currency

NewTrans.BeginTrans

i = 1

While Not rstMain.EOF
    
    If AccID = rstMain("AccID") Then GoTo NextAccount
    AccID = rstMain("AccID")
    
    If AccID Mod 1000 = 0 Then
        rstMain.MoveNext
        If rstMain.EOF Then GoTo NextAccount
        ParentID = GetNewParentID(rstMain)
        If ParentID = 0 Then rstMain.Find "AccID = " & AccID + 1000
        
        NewTrans.SQLStmt = "Select Max(HeadID) From Heads " & _
                        " Where Headid > " & ParentID & _
                        " And HeadID < " & ParentID + 100
        If NewTrans.Fetch(Rst, adOpenDynamic) > 0 Then NewHeadID = FormatField(Rst(0))
        NewHeadID = IIf(NewHeadID > ParentID, NewHeadID, ParentID)
        
        If rstMain("AccID") Mod 1000 = 0 Then GoTo NextAccount
        AccID = rstMain("AccID")
        
    End If
    
    NewHeadID = NewHeadID + 1

    'NOW insert into Heads Table
    SqlStr = "Insert INTO Heads (" & _
                " HeadID,ParentID,HeadName)" & _
                " VALUES (" & NewHeadID & "," & ParentID & "," & _
                AddQuotes(FormatField(rstMain("AccName")), True) & " )"

    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        MsgBox "Unable to transafer the Account No " & rstMain("AccId")
        Exit Function
        GoTo NextAccount
    End If
    
    Balance = 0
    OldTrans.SQLStmt = "Select Top 1 Balance From AccTrans " & _
                " Where AccID =" & AccID & " And TransDate <= " & _
                "#" & FromDate & "# Order By TransID Desc"
    If OldTrans.Fetch(Rst, adOpenDynamic) > 0 Then Balance = FormatField(Rst(0))
    
    If NewHeadID > parLoanAdvanceAsset And _
            NewHeadID < parLoanAdvanceAsset + 100 Then Balance = Balance * -1
    
    SqlStr = "Insert INTO OPBalance (" & _
                " HeadID,OpDate,OpAmount)" & _
                " VALUES (" & NewHeadID & ", #" & DateAdd("d", 1, FromDate) & "# ," & _
                 Balance & " )"

    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        MsgBox "Unable to transafer the Account No " & rstMain("AccId")
        Exit Function
        GoTo NextAccount
    End If
    If Year(FromDate) = 2002 Then
        Balance = 0
        OldTrans.SQLStmt = "Select Top 1 Balance From AccTrans " & _
                " Where AccID =" & AccID & " And TransDate <= #3/31/2003#" & _
                " Order By TransID Desc"
        If OldTrans.Fetch(Rst, adOpenDynamic) > 0 Then Balance = FormatField(Rst(0))
        
        SqlStr = "Insert INTO OPBalance (" & _
                    " HeadID,OpDate,OpAmount)" & _
                    " VALUES (" & NewHeadID & ", " & _
                    "#4/1/2003# ," & _
                     Balance & " )"
    
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            NewTrans.RollBack
            MsgBox "Unable to transafer the Account No " & rstMain("AccId")
            Exit Function
            GoTo NextAccount
        End If
    End If
    
    If ParentID = parLoanAdvanceAsset Or _
        ((ParentID = parReceivable Or ParentID = parPayAble) And AccID Mod 1000 >= 100) Then
        'gDbTrans.SQLStmt = " SELECT MAX(HeadID) FROM Heads " & _
                " WHERE HeadID BETWEEN " & ParentID & " AND " & (wis_CreditorsParentID + SUB_HEAD_OFFSET)
        'Insert into the database
        SqlStr = " INSERT INTO CompanyCreation " & _
                " (HeadID,CompanyName,CompanyType, " & _
                " KST,CST,Address,PhoneNo,ContactPerson," & _
                " MobileNo,Email,SameState ) " & _
                " VALUES ( " & _
                NewHeadID & "," & _
                AddQuotes(FormatField(rstMain("AccName")), True) & "," & _
                IIf(ParentID = parReceivable, 2, 3) & "," & _
                "'KST','CST','Addre','PhNo'," & _
                "'CP','MN',''," & _
                "1 ) "
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            NewTrans.RollBack
            MsgBox "Unable to transafer the Account No " & rstMain("AccId")
            Exit Function
            GoTo NextAccount
        End If
      'End If
    End If
    
NextAccount:
    
    i = i + 1
    With frmMain
        .lblProgress = "Transferring bank account details"
        .prg.Value = i
    End With
    If Not rstMain.EOF Then rstMain.MoveNext
Wend

    NewTrans.CommitTrans

Exit_line:

TransferBankMaster = True
Exit Function

Err_Line:

NewTrans.RollBack

If Err.Number = 3021 Then Err.Clear: Resume Next
    If Err Then MsgBox "eror In Account Master " & Err.Description
'Resume

End Function

'this function is used to transfer the
'transaction details form OLd Db to new one
'and NewTrans has assigned to new database
Private Function TransferBankTrans(OldBankTrans As clsOldUtils, NewTrans As clsDBUtils) As Boolean

On Error GoTo Err_Line

Dim BankAccExpId As Long
Dim BankAccIncId As Long

Dim OldTrans As Integer, NewTransType As Integer
Dim TransID As Long
Dim AccNAme As String
Dim DrAmount As Currency
Dim CrAmount As Currency
Dim Amount_2 As Currency

Dim AccID As Long
Dim Amount As Currency
Dim TransDate As Date
Dim OldParentID As Long

Dim rstTemp As Recordset
Dim rstMain As ADODB.Recordset
Dim AccTrans As clsAccTrans
Dim BankClass As clsBankAcc

'Fetch the detials of Account
OldBankTrans.SQLStmt = "SELECT * FROM AccTrans " & _
                "WHERE AccId in (Select AccId From AccMaster) " & _
                "AND TransDate >=#" & gFromDate & "# " & _
                "ORDER BY AccID,TransId"

If OldBankTrans.Fetch(rstMain, adOpenStatic) < 1 Then
'    MsgBox "No transaction to transfer", vbInformation, "Bank Account Trans"
    OldBankTrans.SQLStmt = "SELECT Max(TransID)as  MaxID," & _
                        " AccId as AccID1 FROM AccTrans " & _
                        " Where TransDate <= #" & gFromDate & "# Group by AccID "
    OldBankTrans.CreateView ("QryMaxID")
    
    OldBankTrans.SQLStmt = "SELECT A.* FROM AccTrans A, QryMaxID B WHERE " & _
                " A.AccId in (Select AccId From AccMaster)" & _
                " And A.AccId = B.AccID1 And A.TransID = B.MAxID " & _
                " ORDER BY AccID,TransId"
    If OldBankTrans.Fetch(rstMain, adOpenStatic) < 1 Then GoTo Exit_line
    'Exit Function
End If

NewTrans.SQLStmt = "Select Max(TransID) From AccTrans "
If NewTrans.Fetch(rstTemp, adOpenStatic) > 0 Then TransID = FormatField(rstTemp(0))


With frmMain
    .lblProgress = "Transferring bank transaction details"
    .prg.Max = rstMain.RecordCount + TransID + 5
    .prg.Min = TransID
    .prg.Value = TransID
    .Refresh
End With

Dim VoucherType As Wis_VoucherTypes
Dim strRemarks As String
Dim TransHeadID As Long
Dim NewHeadID As Long

Set AccTrans = New clsAccTrans
Set BankClass = New clsBankAcc
Set NewIndexTrans = NewTrans

NewTrans.BeginTrans

While Not rstMain.EOF
    TransID = TransID + 1
    TransHeadID = 0
    strRemarks = ""
    
    If AccID <> rstMain("AccID") Then
        AccID = rstMain("AccID")
        If (AccID - (AccID Mod 1000)) <> OldParentID Then
            OldParentID = (AccID - (AccID Mod 1000))
            NewHeadID = GetNewParentID(rstMain)
            If NewHeadID = 70000 Then rstMain.MoveLast: GoTo NextAccount
            AccID = rstMain("AccID")
        End If
        NewHeadID = NewHeadID + 1
        If gOnlyLedgerHeads Then
            NewTrans.SQLStmt = "Select * From Heads Where HeadID = " & NewHeadID
            If NewTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then
                Set rstTemp = Nothing
                Do
                    If rstMain.EOF Then Exit Do
                    If rstMain("AccID") <> AccID Then Exit Do
                    rstMain.MoveNext
                Loop
                rstMain.MovePrevious
                GoTo NextAccRecord
            End If
        End If
        OldBankTrans.SQLStmt = "Select AccName From AccMaster" & _
                                " WHERE AccId = " & AccID
        If OldBankTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                                AccNAme = FormatField(rstTemp(0))
        NewTrans.SQLStmt = "Select HeadID From Heads" & _
                " WHERE HeadName = " & AddQuotes(AccNAme, True) & _
                " and HeadID >= " & NewHeadID - (NewHeadID Mod 1000) & _
                " And HeadID < " & NewHeadID + 10000
        If NewTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                                NewHeadID = FormatField(rstTemp(0))
        
        If rstMain("AccID") >= OldParentID + 1000 Then GoTo NextAccount
    End If
        
    AccID = rstMain("AccID")
    TransDate = rstMain("TransDate")
    OldTrans = rstMain("TransType")
    Amount = rstMain("Amount")
    DrAmount = Amount
    CrAmount = Amount
    
    VoucherType = VouNothing
    If OldParentID = 1000 Then
        VoucherType = Contra  'IIf(OldTrans > 0, Receipt, Payment)
        
        If Abs(OldTrans) = 1 Then
            If OldTrans > 0 Then CrAmount = 0 Else DrAmount = 0
        Else
            If OldTrans > 0 Then DrAmount = 0 Else CrAmount = 0
            VoucherType = IIf(OldTrans < 0, Receipt, Payment)
            TransHeadID = BankClass.GetHeadIDCreated(LoadResString(gLangOffSet + 418) _
                 & " " & LoadResString(gLangOffSet + 366), _
                parBankIncome, 0, wis_BankAccounts)
            strRemarks = "interest From " & AccNAme
        End If

    End If
    If OldParentID = 2000 Then
        VoucherType = IIf(OldTrans > 0, Receipt, Payment)
        If Abs(OldTrans) = 1 Then
            If OldTrans > 0 Then CrAmount = 0 Else DrAmount = 0
        Else
            If OldTrans > 0 Then DrAmount = 0 Else CrAmount = 0
            VoucherType = IIf(OldTrans < 0, Receipt, Payment)
            'TransHeadID = parIncome + 1
            TransHeadID = BankClass.GetHeadIDCreated(LoadResString(gLangOffSet + 418) _
                & " " & LoadResString(gLangOffSet + 368), _
                parBankExpense, 0, wis_BankAccounts)
            strRemarks = "interest Paid  to " & AccNAme
        End If
    End If

    If OldParentID = 3000 Or OldParentID = 4000 Or OldParentID = 12000 Or OldParentID = 14000 Then
        VoucherType = IIf(OldTrans > 0, Receipt, Payment)
        If OldTrans > 0 Then CrAmount = 0 Else DrAmount = 0
    End If
    If OldParentID = 9000 Or OldParentID = 10000 Or OldParentID = 11000 Or OldParentID = 13000 Then
        VoucherType = IIf(OldTrans > 0, Receipt, Payment)
        If OldTrans > 0 Then CrAmount = 0 Else DrAmount = 0
    End If
    
    If NewHeadID > parIncome And NewHeadID < parExpense Then
        'Income Heads
        If Abs(OldTrans) <> 2 Then TransID = TransID - 1: GoTo NextAccount
        VoucherType = IIf(OldTrans > 0, Payment, Receipt)
        If OldTrans > 0 Then DrAmount = 0 Else CrAmount = 0
    End If
    If NewHeadID > parExpense And NewHeadID < parExpense + 10000 Then
        'expense Heads
        If Abs(OldTrans) <> 2 Then VoucherType = VouNothing
        VoucherType = IIf(OldTrans > 0, Payment, Receipt)
        If OldTrans > 0 Then DrAmount = 0 Else CrAmount = 0
    End If
    
    If TransHeadID = 0 Then TransHeadID = NewHeadID
    If Amount > 0 Then
        'Amount_2 = 0
        'If VoucherType = Receipt Then Amount_2 = Amount: Amount = 0
    With AccTrans
        .TransID = TransID
        If .AllTransHeadsAdd(wis_CashHeadID, DrAmount, CrAmount) <> Success Then GoTo Exit_line
        If .AllTransHeadsAdd(TransHeadID, CrAmount, DrAmount) <> Success Then GoTo Exit_line
        If .SaveVouchers(VoucherType, TransDate, strRemarks) <> Success Then GoTo Exit_line
    End With
    End If
NextAccount:
        
    With frmMain
        .lblProgress = "Transferring bank transactions"
        .prg.Value = TransID
    End With
    AccID = rstMain("AccID")
    
NextAccRecord:
    rstMain.MoveNext
Wend

NewTrans.CommitTrans

Set AccTrans = Nothing
Set NewIndexTrans = Nothing

TransferBankTrans = True
Exit Function

Exit_line:
'If InTrans Then NewTrans.RollBack
NewTrans.RollBack

Exit Function

Err_Line:
'Resume
If Err.Number = 3021 Then Err.Clear: Resume Next
If Err Then MsgBox "Error in Bank Trans" & Err.Description
    
End Function

