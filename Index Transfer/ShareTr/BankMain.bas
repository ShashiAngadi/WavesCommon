Attribute VB_Name = "basBank"
Option Explicit

Public NewIndexTrans As clsDBUtils

Public gstrDate As String
Public gFromDate As Date
Public OldPwd As String
Public NewPwd As String
Public gOnlyLedgerHeads As Boolean


'Decalration Varibles
'Public gDbTrans As clsTransact
'Public gDBUtils As clsDBUtilities
'Public gCurrUser As clsUsers

Private m_UserID As Long

Private m_FinIndianFromDate As String ' Financial year from 1/4
Private m_FinUSFromDate As Date ' Financial year from 1/4
Private m_FinIndianEndDate As String ' Financial year end on 31/3
Private m_FinUSEndDate As Date
Private m_HinduFinFromDate As String 'This is for financial Year from Deewali,Yugadi
Private m_HinduFinEndDate As String


Public Enum wis_CompanyType
   Enum_Self = 0
   Enum_Manufacturer = 1
   Enum_Customers = 2
   Enum_Stockist = 3
   Enum_Branch = 4
End Enum

Public Enum wis_StatePosition
    SameState = 1
    OtherState = 2
End Enum

Public Enum Wis_VoucherTypes
    VouNothing = 0
    Receipt = 1
    Payment = 2
    Purchase = 3
    Sales = 4
    FreePurchase = 5
    Journal = 6
    Contra = 7
    RejectionsIn = 8
    RejectionsOut = 9
    FreeSales = 10
    StockIn = 11
    StockOut = 12
    StoockSoot = 13
End Enum

Public Enum Wis_InvoiceType
    InvoiceNumber = 1
    RONumber = 2 'Release Order No
    STANumber = 3 'Stock Transfer Advice No
End Enum

Public Enum Wis_RedirectType
    NonReDirected = 0
    ReDirected = 1
    AmountPaid = 2
End Enum

Public Enum wis_AccountType
    Asset = 1
    Liability = 2
    Loss = 4
    Profit = 8
    ItemSales = 16
    ItemPurchase = 32
End Enum

Public Enum wis_FunctionReturned
    Failure = 0
    Success = 1
    FatalError = -1
End Enum

Public Enum wis_DrCrType
    enumDebit = 0
    enumCredit = 1
End Enum

Public Enum wis_PrintStatus
    PrintDetailed = 1
    NoPrintDetailed = 2
End Enum

Public Enum wis_PrintTitle
    Enum_PrintTitle = 1
    Enum_NoPrintTitle = 2
End Enum

Public Enum wis_PaymentTerm
    Enum_Cash = 1
    Enum_Credit = 2
    Enum_Cheque = 3
    Enum_DD = 4
End Enum

Public Enum Wis_ReportType
    StockIncludingBranches = 1
    StockOfManuIncBranches
    StockOfGroupIncBranches
    StockOfManuAndGroupIncBranches
    StockOfGroupAndProductIncBranches
    StockOfManuAndGroupAndProductIncBranches
    StockAsOn
    StockOfManufacturer
    StockOfGroup
    StockOfGroupAndProduct
    StockOfManuAndGroups
    StockOfManuAndGroupAndProducts
    
    PurchaseOfBranches
    PurchaseOfManufacturer
    PurchaseOfGroup
    PurchaseOfGroupAndProduct
    PurchaseOfManuAndGroups
    PurchaseOfManuAndGroupAndProducts
    
    SalesIncludingBranches
    SalesOfManuIncBranches
    SalesOfGroupIncBranches
    SalesOfManuAndGroupIncBranches
    SalesOfGroupAndProductIncBranches
    SalesOfManuAndGroupAndProductIncBranches
    SalesOfBranches
    SalesOfManufacturer
    SalesOfGroup
    SalesOfGroupAndProduct
    SalesOfManuAndGroups
    SalesOfManuAndGroupAndProducts
    
    ExpireDateAllProducts
    ExpireDateManufacturer
    ExpireDateGroup
    ExpireDateGroupAndProduct
    ExpireDateManuAndGroups
    ExpireDateManuAndGroupAndProducts
    
    CustomerSalesAllProductsIncAllranches
    CustomerSalesForVendorIncBranches
    CustomerSalesForGroupIncAllranches
    CustomerSalesForVendorAndGroupIncAllranches
    CustomerSalesForGroupAndProductIncAllranches
    CustomerSalesForVendorGroupAndProductIncAllranches
    CustomerSalesAllProducts
    CustomerSalesForVendor
    CustomerSalesForVendorAndGroup
    CustomerSalesForGroup
    CustomerSalesForGroupAndProduct
    CustomerSalesForVendorGroupAndProduct
    
    ShowSalesInvoices
    ShowPurchaseInvoices
End Enum
Public Function GetDBPath() As String
Dim strYearKey As String
Dim USFromDate As String
Dim USToDate As String


USFromDate = FormatDate(FinIndianFromDate)
USToDate = FormatDate(FinIndianEndDate)

strYearKey = "Mat"
strYearKey = strYearKey & Right$(Str(Year(USFromDate)), 2)

strYearKey = strYearKey & Right$(Str(Year(USToDate)), 2)

GetDBPath = App.Path & "\" & strYearKey

End Function
'------------------------------------------------------------
'this sub does some cleanup and shuts down
'------------------------------------------------------------
Sub ShutDownInventory()
On Error Resume Next

'Un Load all the forms wich are opened so for
UnloadAllForms

'Close the Database
'gDbTrans.CloseDB

End

End Sub

'------------------------------------------------------------
'this sub unloads all forms except for the
'SQL, Tables and MDI form
'------------------------------------------------------------
Sub UnloadAllForms()
  On Error Resume Next
  
  Dim i As Integer
  
  'close all forms except for the Tables and SQL forms
  For i = Forms.Count - 1 To 1 Step -1
    Unload Forms(i)
  Next i
End Sub
Private Sub Initialise()
'If gDbTrans Is Nothing Then Set gDbTrans = New clsTransact

End Sub

'Sub Main()
'
''Iniitlaise the transct class
'Call Initialise
'
''Show the main form
'MDIMain.Show
''frmMainMenu.Show
'
''Get the financual Year and users
'If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers
'
'gCurrUser.MaxRetries = 3
'gCurrUser.CancelError = True
'gCurrUser.ShowLoginDialog
'
'If Not gCurrUser.LoginStatus Then
'    MsgBox constAPPLICATION_NAME & " Could not log you on", vbInformation, wis_MESSAGE_TITLE
'    Set MDIMain = Nothing
'    'set frmMainMenu = nothing
'    End
'End If
'
'm_UserID = gCurrUser.UserID
'
''Initialse Kannada
'Call KannadaInitialize
'
''Insert the company
'Call InsertCompany
'
'
'
'End Sub
'
Public Function DecryptData(TheString As String) As String
    Dim lcount As Integer
    Dim lStr1 As String
    Dim lchar As String
    
    For lcount = 1 To Len(TheString)
        lchar = Mid$(TheString, lcount, 1)
        lStr1 = lStr1 & Chr((Asc(lchar) Xor Asc("~")) / 2)
    Next lcount
    DecryptData = lStr1
End Function

Public Function EncryptData(ByVal TheString As String) As String
    Dim lcount As Integer
    Dim lStr1 As String
    Dim lchar As String
    
    For lcount = 1 To Len(TheString)
        lchar = Mid$(TheString, lcount, 1)
         lStr1 = lStr1 & Chr((Asc(lchar) * 2) Xor Asc("~"))
    Next lcount
    EncryptData = lStr1
End Function


Public Property Get FinUSEndDate() As String
    FinUSEndDate = m_FinUSEndDate
End Property

Public Property Get FinUSFromDate() As String
    FinUSFromDate = m_FinUSFromDate
End Property

Public Property Get FinIndianFromDate() As String
FinIndianFromDate = m_FinIndianFromDate
End Property

Public Property Let FinIndianFromDate(ByVal vNewValue As String)
m_FinIndianFromDate = vNewValue
m_FinUSFromDate = FormatDate(vNewValue)
End Property

Public Property Get FinIndianEndDate() As String
FinIndianEndDate = m_FinIndianEndDate
End Property

Public Property Let FinIndianEndDate(ByVal vNewValue As String)

m_FinIndianEndDate = vNewValue
m_FinUSEndDate = FormatDate(vNewValue)
End Property

Public Property Get HinduFinFromDate() As String
HinduFinFromDate = m_HinduFinFromDate
End Property

Public Property Let HinduFinFromDate(ByVal vNewValue As String)
m_HinduFinFromDate = vNewValue
End Property

Public Property Get HinduFinEndDate() As String
HinduFinEndDate = m_HinduFinEndDate
End Property

Public Property Let HinduFinEndDate(ByVal vNewValue As String)
m_HinduFinEndDate = vNewValue
End Property

Public Property Get UserID() As Long
UserID = m_UserID
End Property

Public Property Let UserID(ByVal vNewValue As Long)
m_UserID = vNewValue
End Property
