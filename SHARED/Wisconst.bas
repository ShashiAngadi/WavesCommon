Attribute VB_Name = "wisConst"
Option Explicit
Public Const wisGray = &H80000000
'Shashi 4/12/2000
Public Const vbWhite = &H80000005    '&H80000005&
' Status variable constants...
Public Const wis_CANCEL = 0
Public Const wis_FAILURE = 0
Public Const wis_OK = 1
Public Const wis_SUCCESS = 2
Public Const wis_COMPLETE = 3
Public Const wis_EVENT_SUCCESS = 4
Public Const wis_SHOW_FIRST = 5
Public Const wis_SHOW_PREVIOUS = 6
Public Const wis_SHOW_NEXT = 7
Public Const wis_SHOW_LAST = 8
Public Const wis_PRINT_CURRENT = 9
Public Const wis_PRINT_ALL = 10
Public Const wis_PRINT_ALL_PAUSE = 11
Public Const wis_PRINT_CURRENT_PAUSE = 12
Public Const wis_Print_Excel = 13
' Database updation mode.
Public Const wis_INSERT = 1
Public Const wis_UPDATE = 2

' Query mode constants...
Public Const wis_QUERY_BY_CUSTOMERID = 1
Public Const wis_QUERY_BY_CUSTOMERNAME = 2

' The key name for this application in Registry.
Public Const wis_INDEX2000_KEY = "Software\Waves Information Systems\Index2000"

' Password for database.
Public Const wis_PWD = "wis"

' Title for Message boxes.
Public Const wis_MESSAGE_TITLE = " DataBase Design Info..."

' Module name constants...
'Public Const wis_SB = 1
'Public Const wis_CA = 2
'Public Const wis_FD = 3
'Public Const wis_RD = 4

' Report Constants
Enum wisReports
    wisTradingAccount = 1
    wisDebitCreditStatement = 2
    wisProfitLossStatement = 3
    wisBalanceSheet = 4
    wisDailyRegister = 5
    wisBankBalance = 6
    wisDailyDebitCredit = 7
End Enum


Enum wisModules
    wis_None = 0
    wis_CustReg = 1
    wis_sbacc = 2
    wis_FDAcc = 4
    wis_CAAcc = 8
    wis_RDAcc = 16
    wis_PDAcc = 32
    wis_dlAcc = 64
    wis_Members = 128
    wis_Users = 256
    wis_MatAcc = 512
    wis_Loans = 1024
End Enum

' Enumerated error values...
Enum errors
    wis_DATABASE_NOT_OPEN
    wis_INVALID_DATABASE
    wis_DUPLICATE_ACCNO
    wis_INVALID_MODULEID
    wis_INVALID_ACCNO
    wis_ACCNO_NOT_SET
    wis_MODULEID_NOT_SET
    wis_INIT_FAIL
    wis_FILENOTFOUND
End Enum

Public Enum wisTransactionTypes
    '+ values indicate cust / bank money entering account
    '- values indicate cust / bank money drawn from account
    '1 indicates Money from / to customer
    '2 indicates money from / to bank
    ' wDeposit & wWithdraw are w.r.t to any Accounts
    ' wInterest & wCharges Are w.r.t Bank/Society
    wDeposit = 1        'Customers money into account
    'Redefnation Money Comes Into Accountirrespsctive of the account type
    wWithDraw = -1      'Customers money out of account
    'Money Go out of the account irrespsctive of the account type
    wCharges = -2       'Banks money out of account  (Fines, charges, etc)
    ' Money Comes in to the bank as Loss
    wInterest = 2       'Banks money into account   (Interest provided)
    ' Banks money go out of the bank as Profit
    
    
    'SHASHI 16/9/2000
    ' Extended For Contra Transactions
    ' This Not Iplemented till now on any account except Materail
    ' For Which is special case
    ' The Contra is same as Deposit/Withdrawl
    ' Here in this case no physical money has transferred
    ' Money will be transferred through papers
    ' For Examples if customer submits a cheque of the same bank
    '     In such case the money simply transferred from one accoun to other
    wContraDeposit = 3
    wContraWithdraw = -3
    wContraInterest = 4   ' The Loss which effects the PL but not RP
    wContraCharges = -4   ' The Profit which effects the PL but not RP
    wRPInterest = 5       ' The amount which effects the RP (Cash) as receipts but not PL
    wRPCharges = -5       ' The amount which effects the RP (Cash) as payments but not PL
End Enum



Enum wisLoanCategories
    wisAgriculural = 1
    wisNonAgriculural = 2
End Enum

Enum wisLoanTerm
    wisShortTerm = 1
    wisMidTerm = 2
    wisLongTerm = 3
End Enum

Enum wisStatus
   wisPending = 1
   wisCleared = 2
   wisBounced = 3
End Enum
Function ErrMsg(errNum As Integer, Optional ByVal errParam As String) As String
' Returns the user-defined error description.
Select Case errNum
    Case wis_ACCNO_NOT_SET
        ErrMsg = "The member variable AccNo is not set."
    Case wis_MODULEID_NOT_SET
        ErrMsg = "The member variable 'ModuleID' is not set."
    Case wis_INVALID_MODULEID
        'ErrMsg = "No module installed having id number " & errParam & "."
        ErrMsg = "Module license information not found.  Please contact the vendor for licensed version of software modules."
    Case wis_INIT_FAIL
        ErrMsg = "Error in initializing the module."
    Case wis_INVALID_ACCNO
        ErrMsg = "Account number should be a valid number."
    Case wis_FILENOTFOUND
        ErrMsg = "Could not find the file  - " & errParam & "."
    Case wis_DATABASE_NOT_OPEN
        ErrMsg = "A database must be open, for a query to be executed."
    Case wis_DUPLICATE_ACCNO
        ErrMsg = "The account number " & errParam & " is in use."
    Case wis_INVALID_DATABASE
        ErrMsg = "The database is not proper or is corrupt."
End Select
End Function


