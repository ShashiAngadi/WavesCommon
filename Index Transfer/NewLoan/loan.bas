Attribute VB_Name = "basLoan"
Option Explicit


Enum wisSeason
    wisNoSeason = 0
    wisKhariff = 1
    wisRabi = 2
    wisT_Belt = 3
    wisAnnual = 4
    wisOtherSeason = 5
End Enum

Enum wisFarmerClassification
    SmallFarmer = 1
    BigFarmer = 2
    MarginalFarmer = 3
    OtherFarmer = 4
End Enum

Public Enum wisInstallmentTypes
   Inst_No = 0
   Inst_Daily = 1
   Inst_Weekly = 2
   Inst_FortNightly = 3
   Inst_Monthly = 4
   Inst_BiMonthly = 5
   Inst_Quartery = 6
   Inst_HalfYearly = 7
   Inst_Yearly = 8
End Enum

Enum wis_LoanType
    wisCashCreditLoan = 1
    wisVehicleloan = 2
    wisCropLoan = 4
    wisIndividualLoan = 8
    wisBKCC = 16
End Enum


Public Enum wis_LoanReports
    '''Regular Reportr
    repMonthlyRegister = 1
    repMonthlyRegisterAll = 2
    repShedule_1 = 3
    repShedule_2 = 4
    repShedule_3 = 5
    repShedule_4 = 6
    repShedule_5 = 7
    repShedule_6 = 8
  ''Reports
    repLoanBalance = 11
    repLoanHolder = 12
    repLoanTransMade = 13
    repLoanIssued = 14
    repLoanInstOD = 15
    repLoanIntCol = 16
    repLoanDailyCash = 17
    repLoanGLedger = 18
    repLoanRepMade = 19
    repLoanOD = 20
    repLoanSanction = 21
    repLoanGuarantor = 22
    
    repConsBalance = 30
    repConsInstOD = 31
    repConsOD = 32
    
End Enum


Public Sub Show()

End Sub


