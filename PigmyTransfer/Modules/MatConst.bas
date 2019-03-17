Attribute VB_Name = "basMatConst"
Option Explicit
Public gOnLine As Boolean
'Public Const constAPPLICATION_NAME = "Material Management"
'Public Const wis_MESSAGE_TITLE = "Material Management"
Public Const constFINYEARFILE = "FinYear.Fin"
Public Const constBAKFILEPREFIX = "BakI2K"
Public Const constDBName = "Index 2000.mdb"
Public Const constDBPWD = "WIS@$)*"
Public Const constREGKEYNAME = "Software\Waves Information Systems\Index 2000V3\Settings"

Public Const wis_CashParentID = 100000
Public Const wis_BanksParentID = 110000
Public Const wis_BanksODParentID = 40100
Public Const wis_CreditorsParentID = 50000
Public Const wis_DebitorsParentID = 150000
Public Const wis_PurchaseParentID = 210000
Public Const wis_SalesParentID = 220000

Public Const wis_IncomeParentID = 180000 ' all income heads
Public Const wis_ExpenseParentID = 190000 ' all expense heads

Public Const wis_TradingIncParentID = 180100 ' all Trading income heads
Public Const wis_TradingExpParentID = 190100 ' all Trading expense heads

Public Const wis_DummyHeadID = 1
Public Const wis_CashHeadID = 100001
Public Const HEAD_OFFSET = 10000
Public Const SUB_HEAD_OFFSET = 100

Public Const wis_PrevProfitHeadID = 70001

Public Enum wis_ParentHeads
        parShareCapital = 10000
        parMemberShare = 10100
        
        parReserveFunds = 20000
        
'        parDepositLiab = 30000
        parMemberDeposit = 30000
        
        parBankLoanAccount = 40000
        parGovtLoanSubsidy = 40100
        parOtherLoans = 40200
        
        parPayAble = 50000
        parDepositIntProv = 50100
        parLoanIntProv = 50200
        
        parSuspAcc = 60000
        
        'ASSET HEADS
        parCash = 100000
        'parCashHeadID = 100001
        parBankAccount = 110000
        'parDepositAsset = 120000
        parInvestment = 120000
        
        parLoanAdvanceAsset = 130000
        parMemberLoan = 130100
        parMemDepLoan = 130200
        parSalaryAdvance = 130300
        
        parFixedAsset = 140000
        parReceivable = 150000
        
        parIncome = 180000 ' all income heads
        parTradingIncome = 180100 ' all Trading income heads
        parMemLoanIntReceived = 180200 ' interest received on loans parent head id
        parMemLoanPenalInt = 180300 ' interest received on loans parent head id
        parMemDepLoanIntReceived = 180400 'Deposit Loan interest received parent head id
        parDepIntReceived = 180500 'Deposit Loan interest received parent head id
        parBankIncome = 180600      'Incomes Like Share Fee Meber fee will add in this
        
        parExpense = 190000 ' all expense heads
        parTradingExpense = 190100 ' all Trading expense heads
        parMemDepIntPaid = 190200 'interest paid on Deposits parent head id
        
        parLoanIntPaid = 190300  'interest paid on loans parent head id
        parBankExpense = 190400      'expense Like SB Interest,CA Interest will add in this
        parSalaryExpense = 190500
        
        parPurchase = 210000
        parSales = 220000


        parProfitORLoss = 70000  'This parent Head will hace tow heads
                                 'One is Prevouis years Profit or Loss
                                 'Privious year is in profit it will be positve
                                 'otherwise it will be negative
        
End Enum
