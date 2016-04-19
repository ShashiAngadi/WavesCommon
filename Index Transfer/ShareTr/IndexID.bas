Attribute VB_Name = "basIndexID"
Option Explicit

Public Enum WIS_IndexIDs

    DepositSB = 1
    DepositCA
    DepositPigmy
    DepositRD
    DepositBKCC
    
    LoansDeposit
    LoansRD
    LoansPigmy
    LoansNonAgri
    LoansBKCC
    
    ProfitDepositSB
    ProfitDepositCA
    ProfitDepositPigmy
    ProfitDepositRD
    ProfitDepositBKCC
    ProfitLoansDeposit
    ProfitLoansRD
    ProfitLoansPigmy
    ProfitLoansNonAgri
    ProfitLoansBKCC
    
    LossDepositSB
    LossDepositCA
    LossDepositPigmy
    LossDepositRD
    LossDepositBKCC
    LossLoansDeposit
    LossLoansRD
    LossLoansPigmy
    LossLoansNonAgri
    LossLoansBKCC
    
    PayAbleDepositPigmy
    PayAbleDepositRD
    
End Enum
' This Function Will Read the IndexIDs and Return the Respective Material ID
' ID will be kept the IndexIDs Table in the Database
Public Function GetIDForIndexEnum(ByVal IndexIds As WIS_IndexIDs) As Long

On Error GoTo Hell:

Dim rstID As ADODB.Recordset

GetIDForIndexEnum = 0

NewIndexTrans.SQLStmt = " SELECT MaterialID From IndexIDs WHERE IndexID=" & IndexIds

If NewIndexTrans.Fetch(rstID, adOpenForwardOnly) < 0 Then Exit Function

GetIDForIndexEnum = FormatField(rstID.Fields("MaterialID"))

Set rstID = Nothing

Exit Function

Hell:
        
End Function
