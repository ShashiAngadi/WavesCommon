VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_frmWeb As frmWeb
Attribute m_frmWeb.VB_VarHelpID = -1

Private WithEvents m_frmRptDt As frmRptDt
Attribute m_frmRptDt.VB_VarHelpID = -1
Private WithEvents m_frmAccReports As frmAccReports
Attribute m_frmAccReports.VB_VarHelpID = -1
Private WithEvents m_WebDoc As HTMLDocument
Attribute m_WebDoc.VB_VarHelpID = -1

Private m_WebTable As HTMLTable
Private m_WebTBody As HTMLTableSection
Private m_WebRow As HTMLTableRow
Private m_WebCell As HTMLTableCell
Private m_DOmElement

' To Handle Report Dates
Private m_FromRpDate As String
Private m_ToRpDate As String

Private m_RpHeadId As Long
Private m_RpParentID As Long

Private m_FromRpAmount As Currency
Private m_ToRpAmount As Currency

Private m_AccountReportList As Wis_AccountReportList
' This will fill the imablanced row of the Balance Sheet
Private Sub BALFillEmptyRows()

Dim CurrentRow As HTMLTableRow
Dim CurrentTable As HTMLTable

Dim LiabilityRows As Integer
Dim AssetRows As Integer
Dim RowsToAdd As Integer
Dim Row As Integer

LiabilityRows = m_WebRow.cells(0).All.Liability.Rows.Length - 1
AssetRows = m_WebRow.cells(1).All.Asset.Rows.Length - 1

If LiabilityRows > AssetRows Then

    RowsToAdd = LiabilityRows - AssetRows
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    
Else
    
    RowsToAdd = AssetRows - LiabilityRows
    Set CurrentTable = m_WebRow.cells(0).All.Liability

End If

For Row = 1 To RowsToAdd

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    
Next Row
    

End Sub

Private Sub BALONstartReport()

Dim DOMElement

With m_WebDoc
   
    Set m_WebTable = m_WebDoc.All("HeaderTbl")
   
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = " Balance as on " & ToRpDate
        .Style.fontWeight = "Bold"
        .Style.FontSize = "11"
        .Align = "Center"
    End With
    
    Call StartHtmlTable("Balance", .body, 4)
    
End With

With m_WebTable
    .Width = "500"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = "Sl No"
        .Width = "10%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = "Head Name"
        .Width = "65%"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = "Amount"
        .Width = "20%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = " "
        .Width = "5%"
        .Align = "Left"
    End With
    
End With

End Sub
Private Sub TRIALBALstartReport()

Dim DOMElement

With m_WebDoc
   
    Set m_WebTable = m_WebDoc.All("HeaderTbl")
   
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = " Trial Balance as on " & ToRpDate
        .Style.fontWeight = "Bold"
        .Style.FontSize = "11"
        .Align = "Center"
    End With
    
    Call StartHtmlTable("Balance", .body, 4)
    
End With

With m_WebTable
    .Width = "500"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = "Sl No"
        .Width = "10%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = "Head Name"
        .Width = "50%"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = "DR Amount"
        .Width = "20%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = "CR Amount "
        .Width = "20%"
        .Align = "Left"
    End With
    
End With

End Sub

Private Sub GLStartReport()

Dim DOMElement

With m_WebDoc
    
    Set m_WebTable = .body.All("HeaderTbl")
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    With m_WebRow.cells(0)
        .innerText = "General Ledger For : " & ToRpDate
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
            
    Call StartHtmlTable("GLedger", .body, 6)
    
End With

With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = "Sl No"
        .Width = "7%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = "Particulars"
        .Width = "38%"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = "OP Balance"
        .Width = "13%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = "Receipts"
        .Width = "13%"
        .Align = "Right"
    End With
    
    With .cells(4)
        .innerText = "Payments"
        .Width = "13%"
        .Align = "Right"
    End With
    
    With .cells(5)
        .innerText = "CL Balance"
        .Width = "13%"
        .Align = "Right"
    End With
    
End With

End Sub

Private Sub LDCheckPageLength(ByRef SlNo As Long, _
                              ByVal ClosingBalance As Currency, _
                              ByVal AccountType As wis_AccountType)
If SlNo = 56 Then

    ' This will print the Closing Balance
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
    
    Call LDSetHeaderFooter(ClosingBalance, AccountType, "Amount Carried Forward")
    
    ' This will start the Header
   
    LDPrintHeader (True)
   
    Call LDSetHeaderFooter(ClosingBalance, AccountType, "Amount Brought Forward")
   
    SlNo = 0
    
End If

End Sub

Private Sub LDCreateViews(ByVal HeadID As Long, _
                         rstTrans As ADODB.Recordset, _
                         rstLedger As ADODB.Recordset)
                          

' Set the Sql Statement

gDbTrans.SQLStmt = " SELECT TransID,Debit,Credit,TransDate" & _
                    " FROM AccTrans " & _
                    " WHERE HeadID = " & HeadID & _
                    " AND TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                    " AND #" & FormatDate(ToRpDate) & "#" & _
                    " ORDER BY TransDate,TransID"

If Not gDbTrans.CreateView("QryLedger") Then Exit Sub
    
gDbTrans.SQLStmt = " SELECT * " & _
                  " FROM QryLedger"
                
'Fetch the Data to Recordset based on the Condition
Call gDbTrans.Fetch(rstTrans, adOpenForwardOnly)

gDbTrans.SQLStmt = " SELECT a.*,HeadName" & _
                    " FROM AccTrans a, QryLedger b,Heads c " & _
                    " WHERE a.HeadID <> " & HeadID & _
                    " AND a.TransID=b.TransID" & _
                    " AND a.HeadID=c.HeadID"
                    
Call gDbTrans.Fetch(rstLedger, adOpenStatic)

End Sub

Private Sub LDSetRowDetails(ByVal TransDate As String, _
                           ByVal HeadName As String, _
                           ByVal VoucherName As String, _
                           ByVal Debit As Currency, _
                           ByVal Credit As Currency, _
                           ByVal ClosingBalance As Currency)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow.cells(0)
    .innerText = FormatDate(TransDate)
    .Align = "Right"
End With
With m_WebRow.cells(1)
    .innerText = HeadName
    .Align = "Left"
End With
With m_WebRow.cells(2)
    .innerText = VoucherName
    .Align = "Left"
End With

If Debit > 0 Then
    With m_WebRow.cells(3)
        .innerText = FormatCurrency(Debit)
        .Align = "Right"
    End With
End If

If Credit > 0 Then
    With m_WebRow.cells(4)
        .innerText = FormatCurrency(Credit)
        .Align = "Right"
    End With
End If

With m_WebRow.cells(5)
    .innerText = FormatCurrency(ClosingBalance)
    .Align = "Right"
End With

End Sub

Private Sub LoadFrmWeb()

If Not m_frmWeb Is Nothing Then Exit Sub

Set m_frmWeb = New frmWeb
Load m_frmWeb

End Sub

Private Sub SubDaystartReport()

Dim DOMElement

With m_WebDoc

    Set m_WebTable = .body.All("HeaderTbl")
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = " Sub Day Book from " & FromRpDate & " to " & ToRpDate
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    
    Call StartHtmlTable("SubDayBook", .body, 4)
    
End With

With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = "Sl No"
        .Width = "7%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = "Particulars"
        .Width = "45%"
        .Align = "Left"
    End With
    
    
    With .cells(2)
        .innerText = "Receipts"
        .Width = "12%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = "Payments"
        .Width = "12%"
        .Align = "Right"
    End With
    
    
End With

End Sub

Private Sub DBStartReport()

Dim DOMElement

With m_WebDoc
        
    Set m_WebTable = .body.All("HeaderTbl")
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = " Day Book For : " & ToRpDate
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    
    Call StartHtmlTable("DayBook", .body, 6)
    
End With

With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 6)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = "Sl No"
        .Width = "7%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = "Date"
        .Width = "12%"
        .Align = "Right"
    End With
    
    
    With .cells(2)
        .innerText = "Particulars"
        .Width = "45%"
        .Align = "Left"
    End With
    
    With .cells(3)
        .innerText = "Vouh Type"
        .Width = "12%"
        .Align = "Left"
    End With
    
    With .cells(4)
        .innerText = "Debit Amt"
        .Width = "12%"
        .Align = "Right"
    End With
    
    With .cells(5)
        .innerText = "Credit Amt"
        .Width = "12%"
        .Align = "Right"
    End With
    
    
End With

End Sub

' This sub will process the selected report in the from
Public Function ProcessSelectedReport() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations
ProcessSelectedReport = False

' Selection
Select Case m_AccountReportList
    Case AccountLedger
        If Not ShowLedgerToHtml Then Exit Function
     
    Case AccountLedgerOnDate
        If Not ShowLedgerToHtml Then Exit Function
    
    Case AccountsClosed
    
    Case BalancesAsON
        If Not ShowBalancesAsON Then Exit Function
        
    Case DayBook
        If Not ShowDayBook Then Exit Function
        
    Case SubDayBook
        If Not ShowSubDayBook Then Exit Function
        
    Case GeneralLedger
        If Not ShowGeneralLedger Then Exit Function
        
    Case ProfitandLossTrans
    
    Case ReportNothing
    
    Case TotalTransActionsMade
    
End Select

ProcessSelectedReport = True
Exit Function

Hell:
    If Err Then MsgBox "ProcessSelectedReport : " & vbCrLf & Err.Description
   
End Function

' This function will show the Transactions of the day selected
' This will read torpDate property
Private Function ShowDayBook() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstDayBook As ADODB.Recordset

Dim HeadID As Long
Dim SlNo As Long
Dim TransID As Long

Dim HeadName As String
Dim TransDate As String
Dim VoucherName As String

Dim Debit As Currency
Dim Credit As Currency

Dim VoucherType As Wis_VoucherTypes

Dim LedgerClass As clsLedger

' Intialise the Function
ShowDayBook = False

' Check the Inputs
If ToRpDate = "" Then Exit Function

' Set the Sql Stmt

gDbTrans.SQLStmt = "SELECT a.TransID, a.HeadID, a.Debit, a.Credit, a.TransDate, " & _
                  " a.VoucherType,b.HeadName " & _
                  " FROM AccTrans a,Heads b " & _
                  " WHERE Transdate = # " & FormatDate(ToRpDate) & "#" & _
                  " and a.HeadID=b.HeadID " & _
                  " ORDER BY a.TransID"

' Fetch Recordset
'If gDbTrans.Fetch(rstDayBook, adOpenForwardOnly) < 0 Then Exit Function
Call gDbTrans.Fetch(rstDayBook, adOpenForwardOnly)
' Start Report and its Headers
DBStartReport

' Start Loop

Set LedgerClass = New clsLedger

SlNo = 0

Do While Not rstDayBook.EOF
    
    HeadName = FormatField(rstDayBook.Fields("HeadName"))
    TransDate = FormatField(rstDayBook.Fields("TransDate"))
    VoucherType = FormatField(rstDayBook.Fields("VoucherType"))
    Debit = FormatField(rstDayBook.Fields("Debit"))
    Credit = FormatField(rstDayBook.Fields("Credit"))
    TransID = FormatField(rstDayBook.Fields("TransID"))
    SlNo = SlNo + 1
    VoucherName = LedgerClass.GetVoucherName(VoucherType)
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With
    With m_WebRow.cells(1)
        .innerText = TransDate
        .Align = "Right"
    End With
    With m_WebRow.cells(2)
        .innerText = HeadName
        .Align = "Left"
    End With
    With m_WebRow.cells(3)
        .innerText = VoucherName
        .Align = "Left"
    End With
    If Debit > 0 Then
        With m_WebRow.cells(4)
            .innerText = FormatCurrency(Debit)
            .Align = "Right"
        End With
    End If
    
    If Credit > 0 Then
        With m_WebRow.cells(5)
            .innerText = FormatCurrency(Credit)
            .Align = "Right"
        End With
    End If
    
    Do
        If rstDayBook.EOF Then Exit Do
        If FormatField(rstDayBook.Fields("TransID")) <> TransID Then Exit Do
        rstDayBook.MoveNext
    Loop
Loop

Set LedgerClass = Nothing
Set rstDayBook = Nothing

ShowDayBook = True

Exit Function

Hell:

    If Err Then MsgBox "DayBook : " & vbCrLf & Err.Description
    
    Set LedgerClass = Nothing
    Set rstDayBook = Nothing

    

End Function

Private Sub IntialiseWebDocument()

If m_frmWeb Is Nothing Then Exit Sub

If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

m_WebDoc.body.Style.fontFamily = gFontName

SetCompanyTitle

End Sub


Private Sub LDPrintHeader(isInsertPageBreak As Boolean)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

If isInsertPageBreak Then m_WebRow.Style.pageBreakBefore = "Always"
    
With m_WebRow
    With .cells(0)
        .innerText = "Date"
        .Align = "Right"
        .Width = "11%"
    End With
    With .cells(1)
        .innerText = "Ledger Name"
        .Align = "Left"
        .Width = "35%"
    End With
    With .cells(2)
        .innerText = "Vouch Type"
        .Align = "Left"
        .Width = "14%"
    End With
    With .cells(3)
        .innerText = "Debit"
        .Align = "Right"
        .Width = "14%"
    End With
    With .cells(4)
        .innerText = "Credit"
        .Align = "Right"
        .Width = "14%"
    End With
    With .cells(5)
        .innerText = "Balance"
        .Align = "Right"
        .Width = "15%"
    End With

End With

End Sub

Private Sub LDStartLedger()

Dim DOMElement
Dim HeadName As String

Dim AccTransClass As clsAccTrans


Set m_WebDoc = m_frmWeb.web.document


With m_WebDoc

    
    Set m_WebTable = .body.All("HeaderTbl")
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    Set AccTransClass = New clsAccTrans
    
    HeadName = AccTransClass.GetHeadName(RpHeadID)
    
    Set AccTransClass = Nothing
    
    With m_WebRow.cells(0)
        .innerText = HeadName & " Ledger from " & FromRpDate & " to " & ToRpDate
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    

    Call StartHtmlTable("LedgerTable", .body, 6)
    
    
        
End With

LDPrintHeader (False)

End Sub

'This will fill the imablanced row of the Balance Sheet
Private Sub PLFillEmptyRows()

Dim CurrentRow As HTMLTableRow
Dim CurrentTable As HTMLTable

Dim ExpenseRows As Integer
Dim IncomeRows As Integer
Dim Row As Integer
Dim RowsToAdd As Integer

ExpenseRows = m_WebRow.cells(0).All.Expense.Rows.Length - 1
IncomeRows = m_WebRow.cells(1).All.Income.Rows.Length - 1

If ExpenseRows > IncomeRows Then

    RowsToAdd = ExpenseRows - IncomeRows
    Set CurrentTable = m_WebRow.cells(1).All.Income
    
Else
    RowsToAdd = IncomeRows - ExpenseRows
    Set CurrentTable = m_WebRow.cells(0).All.Expense

End If

For Row = 1 To RowsToAdd

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

Next Row
    
End Sub

'This function will show the balances of the specified heads to the web control
Private Function ShowBalancesAsON() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstBalances As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstOpBalance As ADODB.Recordset

Dim AccTransClass As clsAccTrans
Dim DRCRType As wis_DrCrType

Dim Balance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim OpBalance As Currency

Dim HeadID As Long
Dim ParentID As Long
Dim PrevParentID As Long

Dim SlNo As Long

Dim HeadName As String
Dim ParentName As String
Dim FromDate As String
Dim ToDate As String

Dim fldHeadID As ADODB.Field
Dim fldHeadName As ADODB.Field
Dim fldParentName As ADODB.Field
Dim fldParentID As ADODB.Field
Dim fldAccountType As ADODB.Field

Dim isHeadID As Boolean

Dim AccountType As wis_AccountType


' Intialise the Function
ShowBalancesAsON = False

Set AccTransClass = New clsAccTrans

' Start the report
BALONstartReport

ParentID = RpParentID
HeadID = RpHeadID

FromDate = FromRpDate
ToDate = ToRpDate

' Set the Sql
If RpHeadID = 0 Then

    gDbTrans.SQLStmt = " SELECT HeadID,OpAmount" & _
                       " FROM OpBalance" & _
                       " WHERE OpDate=" & "#" & FormatDate(FromRpDate) & "#"
                     
    Call gDbTrans.Fetch(rstOpBalance, adOpenDynamic)
    
    gDbTrans.SQLStmt = " SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal,HeadId" & _
                       " FROM AccTrans a" & _
                       " WHERE TransDate BETWEEN #" & FormatDate(FromDate) & "#" & _
                       " AND #" & FormatDate(ToDate) & "#" & _
                       " GROUP BY HeadID"
                       
    Call gDbTrans.Fetch(rstTrans, adOpenDynamic)
    
    isHeadID = False
    
End If

If RpHeadID > 0 Then isHeadID = True

If RpParentID = 0 And RpHeadID = 0 Then _
    gDbTrans.SQLStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                       " FROM Heads a,ParentHeads b" & _
                       " WHERE a.ParentID=b.ParentID" & _
                       " ORDER BY b.ParentID"

If RpParentID > 0 And RpHeadID = 0 Then _
    gDbTrans.SQLStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                       " FROM Heads a,ParentHeads b" & _
                       " WHERE a.ParentID=b.ParentID" & _
                       " AND b.ParentID=" & ParentID

If RpParentID > 0 And RpHeadID > 0 Then _
    gDbTrans.SQLStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                       " FROM Heads a,ParentHeads b" & _
                       " WHERE a.HeadID=" & HeadID & _
                       " AND a.ParentID=b.ParentID"

' Get the Recordset

Call gDbTrans.Fetch(rstBalances, adOpenForwardOnly)

Set fldHeadID = rstBalances.Fields("HeadID")
Set fldHeadName = rstBalances.Fields("HeadName")
Set fldParentName = rstBalances.Fields("ParentName")
Set fldParentID = rstBalances.Fields("ParentID")
Set fldAccountType = rstBalances.Fields("AccountType")

' Start the Loop

SlNo = 0
Do While Not rstBalances.EOF
    
    HeadID = fldHeadID.value
    HeadName = fldHeadName.value
    ParentName = fldParentName.value
    ParentID = fldParentID.value
    AccountType = fldAccountType.value
    
    ' fill the parentname
    If PrevParentID <> ParentID Then
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
        m_WebRow.Style.fontWeight = "Bold"
        m_WebRow.cells(1).innerText = ParentName
        PrevParentID = ParentID
    End If
    
    If Not isHeadID Then
        OpBalance = 0
        If gDbTrans.FindRecord(rstOpBalance, "HeadID=" & HeadID) Then _
                             OpBalance = rstOpBalance.Fields("OpAmount")
                        
        DebitTotal = 0
        CreditTotal = 0
    
        If gDbTrans.FindRecord(rstTrans, "HeadID=" & HeadID) Then
            DebitTotal = FormatField(rstTrans.Fields("DebitTotal"))
            CreditTotal = FormatField(rstTrans.Fields("CreditTotal"))
        End If
                    
        Balance = AccTransClass.GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)

    Else
        Balance = AccTransClass.GetOpBalance(HeadID, ToDate)
    End If
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

    SlNo = SlNo + 1

    With m_WebRow
        With .cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With

        With .cells(1)
            .innerText = HeadName
            .Align = "Left"
        End With

        With .cells(2)
            .innerText = FormatCurrency(Abs(Balance))
            .Align = "Right"
        End With
        
        DRCRType = GetDRCRType(Balance, AccountType)
        
        If DRCRType = enumDebit Then
            With .cells(3)
                .innerText = "DR"
                .Align = "Right"
            End With
        End If
        
        If DRCRType = enumCredit Then
            With .cells(3)
                .innerText = "CR"
                .Align = "Right"
            End With
        End If
        
    End With
        
    rstBalances.MoveNext
Loop


' Exit function
Set AccTransClass = Nothing

Set rstBalances = Nothing
Set rstTrans = Nothing
Set rstOpBalance = Nothing

ShowBalancesAsON = True

Exit Function

Hell:

    If Err Then MsgBox "ShowBalancesAsOn : " & vbCrLf & Err.Description
    
    Set AccTransClass = Nothing
    
    Set rstBalances = Nothing
    Set rstTrans = Nothing
    Set rstOpBalance = Nothing

    Exit Function
    
End Function
'This function will show the balances of the specified heads to the web control
Public Function ShowTrialBalance() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstBalances As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstOpBalance As ADODB.Recordset

Dim AccTransClass As clsAccTrans
'Dim MaterialClass As clsMaterial

Dim Balance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim StockValue As Currency
Dim GrDebitTotal As Currency
Dim GrCreditTotal As Currency

Dim OpBalance As Currency

Dim HeadID As Long
Dim ParentID As Long
Dim PrevParentID As Long

Dim SlNo As Long

Dim HeadName As String
Dim ParentName As String
Dim FromDate As String
Dim ToDate As String

Dim DRCRType As wis_DrCrType

Dim fldHeadID As ADODB.Field
Dim fldHeadName As ADODB.Field
Dim fldParentName As ADODB.Field
Dim fldParentID As ADODB.Field
Dim fldAccountType As ADODB.Field

Dim isHeadID As Boolean

Dim AccountType As wis_AccountType

' Intialise the Function
ShowTrialBalance = False

Set AccTransClass = New clsAccTrans

' Start the report
TRIALBALstartReport

ParentID = RpParentID
HeadID = RpHeadID

FromDate = FromRpDate
ToDate = ToRpDate


gDbTrans.SQLStmt = " SELECT HeadID,OpAmount" & _
                   " FROM OpBalance" & _
                   " WHERE OpDate=" & "#" & FormatDate(FromRpDate) & "#"
                 
Call gDbTrans.Fetch(rstOpBalance, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal,HeadId" & _
                   " FROM AccTrans a" & _
                   " WHERE TransDate BETWEEN #" & FormatDate(FromDate) & "#" & _
                   " AND #" & FormatDate(ToDate) & "#" & _
                   " GROUP BY HeadID"
                   
Call gDbTrans.Fetch(rstTrans, adOpenDynamic)
    

gDbTrans.SQLStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                   " FROM Heads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " ORDER BY b.ParentID"


' Get the Recordset

Call gDbTrans.Fetch(rstBalances, adOpenForwardOnly)

Set fldHeadID = rstBalances.Fields("HeadID")
Set fldHeadName = rstBalances.Fields("HeadName")
Set fldParentName = rstBalances.Fields("ParentName")
Set fldParentID = rstBalances.Fields("ParentID")
Set fldAccountType = rstBalances.Fields("AccountType")

' Start the Loop

SlNo = 0
Do While Not rstBalances.EOF
    
    HeadID = fldHeadID.value
    HeadName = fldHeadName.value
    ParentName = fldParentName.value
    ParentID = fldParentID.value
    AccountType = fldAccountType.value
    
    ' fill the parentname
    If PrevParentID <> ParentID Then
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
        m_WebRow.Style.fontWeight = "Bold"
        m_WebRow.cells(1).innerText = ParentName
        PrevParentID = ParentID
    End If
    
    OpBalance = 0
    If gDbTrans.FindRecord(rstOpBalance, "HeadID=" & HeadID) Then _
                         OpBalance = rstOpBalance.Fields("OpAmount")
                    
    DebitTotal = 0
    CreditTotal = 0
    
    If gDbTrans.FindRecord(rstTrans, "HeadID=" & HeadID) Then
        DebitTotal = FormatField(rstTrans.Fields("DebitTotal"))
        CreditTotal = FormatField(rstTrans.Fields("CreditTotal"))
    End If
                
    Balance = AccTransClass.GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)
    
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

    SlNo = SlNo + 1

    With m_WebRow
        With .cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With

        With .cells(1)
            .innerText = HeadName
            .Align = "Left"
        End With
        
        
        DRCRType = GetDRCRType(Balance, AccountType)
        
        If DRCRType = enumDebit Then
        
            With .cells(2)
                .innerText = FormatCurrency(Abs(Balance))
                .Align = "Right"
            End With
            
            GrDebitTotal = GrDebitTotal + Abs(Balance)
            
        End If
        
        If DRCRType = enumCredit Then
        
            With .cells(3)
                .innerText = FormatCurrency(Abs(Balance))
                .Align = "Right"
            End With
            
            GrCreditTotal = GrCreditTotal + Abs(Balance)
            
        End If
        
    End With
        
    rstBalances.MoveNext
Loop

'Set MaterialClass = New clsMaterial

'StockValue = MaterialClass.GetOnDateClosingStockValue(ToRpDate)

'Set MaterialClass = Nothing

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

SlNo = SlNo + 1

With m_WebRow

    With .cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With

    With .cells(1)
        .innerText = "Stock In Hand"
        .Align = "Left"
    End With
    
    DRCRType = GetDRCRType(StockValue, AccountType)
    
    If DRCRType = enumDebit Then
    
        With .cells(2)
            .innerText = FormatCurrency(Abs(StockValue))
            .Align = "Right"
        End With
        
        GrDebitTotal = GrDebitTotal + Abs(StockValue)
        
    End If
    
    If DRCRType = enumCredit Then
    
        With .cells(3)
            .innerText = FormatCurrency(Abs(StockValue))
            .Align = "Right"
        End With
        
        GrCreditTotal = GrCreditTotal + Abs(StockValue)
        
    End If
    
End With


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow

    .Style.fontWeight = "Bold"
    
    With .cells(0)
        .innerText = ""
        .Align = "Right"
    End With

    With .cells(1)
        .innerText = "Grand Total"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = FormatCurrency(GrDebitTotal)
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = FormatCurrency(GrCreditTotal)
        .Align = "Right"
    End With
    
    
End With


' Exit function
Set AccTransClass = Nothing

Set rstBalances = Nothing
Set rstTrans = Nothing
Set rstOpBalance = Nothing

ShowTrialBalance = True

Exit Function

Hell:

    If Err Then MsgBox "ShowTrialBalance : " & vbCrLf & Err.Description
    
    Set AccTransClass = Nothing
    
    Set rstBalances = Nothing
    Set rstTrans = Nothing
    Set rstOpBalance = Nothing

    Exit Function
    
End Function

' This will show the General Ledger for the Given Period
' This will read the torpDate,fromrpdate,rpParentid,rpHeadid properties

Private Function ShowGeneralLedger() As Boolean

' Handle Error
On Error GoTo Hell:

Dim rstParents As ADODB.Recordset
Dim rstRPDetails As ADODB.Recordset

Dim Receipts As Currency
Dim Payments As Currency
Dim OpBalance As Currency
Dim CLBalance As Currency
Dim SumReceipts As Currency
Dim OpeningCash As Currency
Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim HeadID As Long
Dim ParentID As Long

Dim AsOnDate As String
Dim HeadName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans

Const TableName = "GLedger"
Const QRYRP = "QryRPDetails"
Const TOTALCOLS = 6

ShowGeneralLedger = False

' Validataions
If FromRpDate = "" Then Exit Function
If ToRpDate = "" Then Exit Function

' This is the date to query
AsOnDate = ToRpDate

' Start the Receipts and Payments
GLStartReport

' Put the Opening Cash to the Html
Set AccTransClass = New clsAccTrans

OpeningCash = AccTransClass.GetOpBalance(wis_CashHeadID, AsOnDate)

Call RPPutCashinField(TableName, ToRpDate, 2, "Opening Balance", OpeningCash)

' create views for the receipts and payments
GLCreateViews (QRYRP)

gDbTrans.SQLStmt = " SELECT ParentName,ParentID,AccountType" & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SQLStmt = " SELECT Sum(a.Receipt) AS Receipts, SUM(a.Payment) AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName" & _
                   " FROM " & QRYRP & " AS a, Heads AS b " & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " GROUP BY b.ParentID,a.RPHeadID,b.HeadName"
                   
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)

Set CurrentTable = m_WebTable
Set AccTransClass = New clsAccTrans

SlNo = 0

' Main Loop
Do While Not rstParents.EOF

    ParentID = rstParents.Fields("ParentID")
    ParentName = rstParents.Fields("ParentName")
    AccountType = rstParents.Fields("AccountType")
    
    With rstRPDetails

        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID

        If rstRPDetails.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, TOTALCOLS)

            CurrentRow.Style.fontWeight = "Bold"

            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
            End With
        End If

        Do While Not .EOF

            HeadID = FormatField(.Fields("RPHeadID"))
            HeadName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
 
            OpBalance = AccTransClass.GetOpBalance(HeadID, AsOnDate)
            
            CLBalance = AccTransClass.GetClosingBalanceonAccountType(OpBalance, Payments, Receipts, AccountType)

            SlNo = SlNo + 1
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, TOTALCOLS)
            With CurrentRow.cells(0)
                .innerText = SlNo
                .Align = "Right"
            End With

            With CurrentRow.cells(1)
                .innerText = HeadName
                .Align = "Left"
            End With

            With CurrentRow.cells(2)
                .innerText = FormatCurrency(OpBalance)
                .Align = "Right"
            End With

            With CurrentRow.cells(3)
                .innerText = FormatCurrency(Receipts)
                .Align = "Right"
            End With

            With CurrentRow.cells(4)
                .innerText = FormatCurrency(Payments)
                .Align = "Right"
            End With

            With CurrentRow.cells(5)
                .innerText = FormatCurrency(Abs(CLBalance))
                .Align = "Right"
            End With

            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments

            .MoveNext
        Loop
    End With

    rstParents.MoveNext

Loop

' Now Put Closing Balance
Call RPPutCashinField(TableName, ToRpDate, 5, "Closing Balance", (OpeningCash + SumReceipts) - SumPayments)
 
Set rstRPDetails = Nothing
Set rstParents = Nothing
Set AccTransClass = Nothing

ShowGeneralLedger = True

Exit Function

Hell:

    MsgBox "ShowGeneralLedger: " & vbCrLf & Err.Description
    
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    Set AccTransClass = Nothing
    
End Function

' This will show the General Ledger for the Given Period
' This will read the torpDate,fromrpdate,rpParentid,rpHeadid properties

Private Function ShowSubDayBook() As Boolean

Dim rstParents As ADODB.Recordset
Dim rstRPDetails As ADODB.Recordset

Dim Receipts As Currency
Dim Payments As Currency
Dim OpBalance As Currency
Dim CLBalance As Currency
Dim SumReceipts As Currency
Dim OpeningCash As Currency
Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim HeadID As Long
Dim ParentID As Long

Dim AsOnDate As String
Dim HeadName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans

Const TableName = "SubDayBook"
Const QRYRP = "QryRPDetails"
Const TOTALCOLS = 6

' Handle Error
On Error GoTo Hell:

ShowSubDayBook = False

' Validataions
If FromRpDate = "" Then Exit Function
If ToRpDate = "" Then Exit Function
If RpParentID = 0 Then Exit Function

' This is the date to query
AsOnDate = ToRpDate

' Start the Receipts and Payments
SubDaystartReport

' Put the Opening Cash to the Html

OpeningCash = RPPutOpeningCash(TableName, AsOnDate)

' create views for the receipts and payments
SubDayCreateViews (QRYRP)

gDbTrans.SQLStmt = " SELECT ParentName,ParentID,AccountType" & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SQLStmt = " SELECT Sum(a.Receipt) AS Receipts, SUM(a.Payment) AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName" & _
                   " FROM " & QRYRP & " a, Heads b ,ParentHeads c" & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " AND b.ParentID=c.ParentID" & _
                   " AND c.ParentID=" & RpParentID & _
                   " GROUP BY b.ParentID,a.RPHeadID,b.HeadName"
                   
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)

Set CurrentTable = m_WebTable
Set AccTransClass = New clsAccTrans

SlNo = 0

' Main Loop
Do While Not rstParents.EOF

    ParentID = rstParents.Fields("ParentID")
    ParentName = rstParents.Fields("ParentName")
    AccountType = rstParents.Fields("AccountType")
    
    With rstRPDetails

        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID

        If rstRPDetails.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

            CurrentRow.Style.fontWeight = "Bold"

            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
            End With
        End If

        Do While Not .EOF

            HeadID = FormatField(.Fields("RPHeadID"))
            HeadName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
 
            SlNo = SlNo + 1
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)
            With CurrentRow.cells(0)
                .innerText = SlNo
                .Align = "Right"
            End With

            With CurrentRow.cells(1)
                .innerText = HeadName
                .Align = "Left"
            End With

            With CurrentRow.cells(2)
                .innerText = FormatCurrency(Receipts)
                .Align = "Right"
            End With

            With CurrentRow.cells(3)
                .innerText = FormatCurrency(Payments)
                .Align = "Right"
            End With

            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments

            .MoveNext
        Loop
    End With

    rstParents.MoveNext

Loop

' Now Put Closing Balance
Call RPputClosingCash((OpeningCash + SumReceipts) - SumPayments, TableName)

Set rstRPDetails = Nothing
Set rstParents = Nothing

ShowSubDayBook = True

Exit Function

Hell:
    
    MsgBox "ShowSubDayBook: " & vbCrLf & Err.Description
    
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    
End Function

Public Sub SetCompanyTitle()

Dim rstCompany As ADODB.Recordset
Dim eCompanyType As wis_CompanyType

Dim DOMElement

eCompanyType = Enum_Self

gDbTrans.SQLStmt = " SELECT CompanyName,Address " & _
                   " FROM CompanyCreation " & _
                   " WHERE CompanyType = " & eCompanyType

If gDbTrans.Fetch(rstCompany, adOpenForwardOnly) < 0 Then Exit Sub

If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc
    
    Call StartHtmlTable("HeaderTbl", m_WebDoc.body, 1)
    
    m_WebTable.border = 0
        
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = rstCompany.Fields("CompanyName").value
        .Style.FontSize = "17"
        If gLangOffSet <> 0 Then .Style.FontSize = "19"
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = rstCompany.Fields("Address").value
        .Style.FontSize = "11"
        .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    
End With



End Sub

' This fetch the data from Base table and load it to web control
'
Private Function ShowLedgerToHtml() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations
Dim rstLedger As ADODB.Recordset
Dim rstTrans As ADODB.Recordset

Dim HeadName As String
Dim TransDate As String
Dim VoucherName As String

Dim AccountType As wis_AccountType
Dim VoucherType As Wis_VoucherTypes

Dim Debit As Currency
Dim Credit As Currency
Dim OpBalance As Currency
Dim DebitTotal As Currency
Dim HeadIDDebit As Currency
Dim HeadIDCredit As Currency
Dim CreditTotal As Currency
Dim ClosingBalance As Currency

Dim SlNo As Long
Dim HeadID As Long
Dim TransID As Long
Dim ToHeadID As Long

Dim VoucherNames

Dim LedgerClass As clsLedger
Dim AccTransClass As clsAccTrans

ShowLedgerToHtml = False

' Validations
If RpHeadID = 0 Then Exit Function

HeadID = RpHeadID

Set AccTransClass = New clsAccTrans
Set LedgerClass = New clsLedger

VoucherNames = AccTransClass.GetVoucherNametoArray

LDStartLedger
 
AccountType = AccTransClass.GetAccountType(HeadID)

OpBalance = AccTransClass.GetOpBalance(HeadID, FinIndianFromDate)

Call LDSetHeaderFooter(OpBalance, AccountType, "Opening Balance")

Call LDCreateViews(HeadID, rstTrans, rstLedger)

SlNo = 0

DebitTotal = 0
CreditTotal = 0

Do While Not rstTrans.EOF
    
    TransID = rstTrans.Fields("TransID")
    
    HeadIDDebit = rstTrans.Fields("Debit")
    HeadIDCredit = rstTrans.Fields("Credit")
    
    rstLedger.Filter = adFilterNone
    If HeadIDDebit > 0 Then
        rstLedger.Filter = "TransID=" & TransID & " AND Credit > " & 0
    Else
        rstLedger.Filter = "TransID=" & TransID & " AND Debit > " & 0
    End If
    
    Do While Not rstLedger.EOF
            
        ' This will check the page and reallocate the slno
        Call LDCheckPageLength(SlNo, ClosingBalance, AccountType)
    
        Debit = 0
        Credit = 0
    
        If HeadIDDebit > 0 Then
            If HeadIDDebit < FormatField(rstLedger.Fields("Credit")) Then
                Debit = HeadIDDebit
            Else
                Debit = FormatField(rstLedger.Fields("Credit"))
            End If
            Credit = FormatField(rstLedger.Fields("debit"))
        End If
        
        If HeadIDCredit > 0 Then
            If HeadIDCredit < FormatField(rstLedger.Fields("Debit")) Then
                Credit = HeadIDCredit
            Else
                Credit = FormatField(rstLedger.Fields("Debit"))
            End If
            Debit = FormatField(rstLedger.Fields("Credit"))
        End If
        
        Credit = rstLedger.Fields("Debit")
        ToHeadID = rstLedger.Fields("HeadID")
        HeadName = rstLedger.Fields("HeadName")
        TransDate = rstLedger.Fields("TransDate")
        VoucherType = rstLedger.Fields("VoucherType")
                
        VoucherName = LedgerClass.GetVoucherName(VoucherType)
        
        DebitTotal = DebitTotal + Debit
        CreditTotal = CreditTotal + Credit
        
        ClosingBalance = AccTransClass.GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)
        
        SlNo = SlNo + 1
        
        Call LDSetRowDetails(TransDate, HeadName, VoucherName, Debit, Credit, ClosingBalance)
            
        rstLedger.MoveNext
        
    Loop
   
   
   rstTrans.MoveNext
    
Loop

ClosingBalance = AccTransClass.GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)

Call LDSetHeaderFooter(ClosingBalance, AccountType, "Closing Balance")


Set AccTransClass = Nothing
Set LedgerClass = Nothing
Set rstLedger = Nothing
Set rstTrans = Nothing


' Validations
ShowLedgerToHtml = True
Exit Function

Hell:
    
    MsgBox "ShowLedgerToHtml:" & vbCrLf & Err.Description
    
    Set rstLedger = Nothing
    Set AccTransClass = Nothing
    Set LedgerClass = Nothing
  
End Function
Private Sub LDSetHeaderFooter(ByVal Balance As Currency, AccountType As wis_AccountType, MsgString As String)

Dim ColToPrint As Integer
Dim ColorToPrint As String

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow
        
    .cells(1).innerText = MsgString
    .cells(1).Style.fontWeight = "Bold"
        
      Select Case AccountType
        Case Asset
        
            If Balance >= 0 Then
                ColToPrint = 3
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 4
                ColorToPrint = "Red"
            End If
            
        Case Liability
            
            If Balance >= 0 Then
                ColToPrint = 4
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 3
                ColorToPrint = "Red"
            End If

        Case Loss
        
            If Balance >= 0 Then
                ColToPrint = 3
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 4
                ColorToPrint = "Red"
            End If
            
        Case Profit
        
            If Balance >= 0 Then
                ColToPrint = 4
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 3
                ColorToPrint = "Red"
            End If

            
        Case ItemSales
        
            If Balance >= 0 Then
                ColToPrint = 4
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 3
                ColorToPrint = "Red"
            End If
            
        Case ItemPurchase
        
            If Balance >= 0 Then
                ColToPrint = 3
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 4
                ColorToPrint = "Red"
            End If

    End Select
               
    With .cells(ColToPrint)
        .innerText = FormatCurrency(Abs(Balance))
        .Style.Color = ColorToPrint
        .Style.fontWeight = "Bold"
        .Align = "Right"
    End With
    
End With

End Sub
'
'This will return the current dr or cr balance for the accounttype
Private Function GetDRCRType(ByVal Balance As Currency, _
                              ByVal AccountType As wis_AccountType) As wis_DrCrType
                              
  Select Case AccountType
    Case Asset
    
        If Balance >= 0 Then
            GetDRCRType = enumDebit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumCredit
        End If
        
    Case Liability
        
        If Balance >= 0 Then
            GetDRCRType = enumCredit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumDebit
        End If

    Case Loss
    
        If Balance >= 0 Then
            GetDRCRType = enumDebit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumCredit
        End If
        
    Case Profit
    
        If Balance >= 0 Then
            GetDRCRType = enumCredit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumDebit
        End If

        
    Case ItemSales
    
        If Balance >= 0 Then
            GetDRCRType = enumCredit
        End If
          
        If Balance < 0 Then
            GetDRCRType = enumDebit
        End If
        
    Case ItemPurchase
    
        If Balance >= 0 Then
            GetDRCRType = enumDebit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumCredit
        End If
    
End Select
                   

End Function
Private Sub BALPrintGrandTotals(GrandTotal As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim PrintString As String

PrintString = "Grand Total"

Set CurrentTable = m_WebRow.cells(0).All.Liability

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal)
    .Align = "Right"
End With


Set CurrentTable = m_WebRow.cells(1).All.Asset

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    
CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal)
    .Align = "Right"
End With

End Sub
Private Sub PLPrintGrandTotals(GrandTotal As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim PrintString As String

PrintString = "Net  Total"

Set CurrentTable = m_WebRow.cells(0).All.Expense

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal)
    .Align = "Right"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With


Set CurrentTable = m_WebRow.cells(1).All.Income

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    
CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal)
    .Align = "Right"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub

Private Sub BALPrintProfitorLoss(Profit As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim PrintString As String

If Profit >= 0 Then
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    PrintString = "Profit"
End If

If Profit < 0 Then
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    PrintString = "Loss"
End If
    
CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(Abs(Profit))
    .Align = "Right"
End With

End Sub

Private Sub BALPrintClosingStockValue(ByVal StockValue As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow
Dim PrintString As String

Set CurrentTable = m_WebRow.cells(1).All.Asset
Set ParentRow = InsertRowAndCellSToTable(CurrentTable, 3)

PrintString = "Stock in Hand"

ParentRow.Style.fontWeight = "Bold"

With ParentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With
With ParentRow.cells(2)
    .innerText = FormatCurrency(StockValue)
    .Align = "Right"
    .Style.fontWeight = "Bold"
End With

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

With CurrentRow.cells(0)
    .innerText = "1. " & PrintString
    .Align = "Left"
End With

With CurrentRow.cells(1)
    .innerText = FormatCurrency(StockValue)
    .Align = "Right"
End With

End Sub

Private Sub PLPrintProfitorLoss(Profit As Currency, PrintString As String, isRight As Boolean)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

If isRight Then
    Set CurrentTable = m_WebRow.cells(1).All.Income
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
Else
    Set CurrentTable = m_WebRow.cells(0).All.Expense
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
End If
    
With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(Abs(Profit))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub

Private Sub BALSetBalanceSheetDifference(AssetTotal As Currency, LiabilityTotal As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim ParentAmount As Currency

' If same then not necessary to publish difference
If AssetTotal = LiabilityTotal Then Exit Sub

If AssetTotal > LiabilityTotal Then

    ParentAmount = AssetTotal - LiabilityTotal
    
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    
ElseIf LiabilityTotal > AssetTotal Then

    ParentAmount = LiabilityTotal - AssetTotal
    
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    
End If
    
If Not CurrentTable Is Nothing Then

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    
    CurrentRow.Style.fontWeight = "bold"
    
    With CurrentRow.cells(0)
        .innerText = "Difference in BalanceSheet"
        .Align = "Left"
    End With
    
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(ParentAmount)
        .Align = "Right"
    End With
    
End If
    
' This Equalise the rows of both the tables
BALFillEmptyRows

End Sub

Private Function BALCreateViewsForBalanceSheet() As Boolean

On Error GoTo Hell:

BALCreateViewsForBalanceSheet = False

gDbTrans.SQLStmt = " SELECT a.HeadID as OpHeadID,OpAmount,ParentID,HeadName" & _
                   " FROM OpBalance a,Heads b " & _
                   " WHERE OpDate = " & "#" & FormatDate(FinIndianFromDate) & "#" & _
                   " AND a.HeadID=b.HeadID" & _
                   " ORDER BY a.HeadID"
                   
gDbTrans.CreateView ("qryOPHeads")

gDbTrans.SQLStmt = " SELECT (sum(Credit)-sum(Debit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID, c.ParentName " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=2 " & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("qryLiabilities")

gDbTrans.SQLStmt = " SELECT (sum(Debit)-sum(Credit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID, c.ParentName " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=1 " & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("qryAssets")

BALCreateViewsForBalanceSheet = True

Exit Function

Hell:
    
End Function
Private Function BALCreateViewsForConditionalBalanceSheet() As Boolean

On Error GoTo Hell:

BALCreateViewsForConditionalBalanceSheet = False

gDbTrans.SQLStmt = " SELECT a.HeadID as OpHeadID,OpAmount,ParentID,HeadName" & _
                   " FROM OpBalance a,Heads b " & _
                   " WHERE OpDate = " & "#" & FormatDate(FinIndianFromDate) & "#" & _
                   " AND a.HeadID=b.HeadID" & _
                   " ORDER BY a.HeadID"
                   
gDbTrans.CreateView ("qryOPHeads")

gDbTrans.SQLStmt = " SELECT (sum(Credit)-sum(Debit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID" & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=2 " & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY c.ParentID,a.HeadID, Headname"

                   
gDbTrans.CreateView ("qryLiabilities")

gDbTrans.SQLStmt = " SELECT (sum(Debit)-sum(Credit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID" & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=1 " & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY c.ParentID, a.HeadID, Headname"
                   
gDbTrans.CreateView ("qryAssets")

BALCreateViewsForConditionalBalanceSheet = True

Exit Function

Hell:
    
End Function

Private Function PLCreateViewsForProfitAndLoss() As Boolean

On Error GoTo Hell:

PLCreateViewsForProfitAndLoss = False

gDbTrans.SQLStmt = " SELECT (sum(Credit)-sum(Debit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID, c.ParentName " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=8" & _
                   " AND C.ParentID <> " & wis_TradingIncParentID & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("QryIncomes")

gDbTrans.SQLStmt = " SELECT (sum(Debit)-sum(Credit)) AS Amount, a.Headid , " & _
                   " HeadName, c.ParentID, c.ParentName " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=4 " & _
                   " AND C.ParentID <> " & wis_TradingExpParentID & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("QryExpenses")

PLCreateViewsForProfitAndLoss = True

Exit Function

Hell:
    
End Function
Private Function NetProfitViews() As Boolean

On Error GoTo Hell:

NetProfitViews = False

gDbTrans.SQLStmt = " SELECT (sum(Credit)-sum(Debit)) As NetIncome " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=8" & _
                   " AND C.ParentID <> " & wis_TradingIncParentID & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#"
                   
gDbTrans.CreateView ("QryIncomes")

gDbTrans.SQLStmt = " SELECT (sum(Debit)-sum(Credit)) As NetExpenses " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=4 " & _
                   " AND C.ParentID <> " & wis_TradingExpParentID & _
                   " AND a.TransDate BETWEEN #" & FormatDate(FromRpDate) & "#" & _
                   " AND #" & FormatDate(ToRpDate) & "#"
                   
gDbTrans.CreateView ("QryExpenses")

                   
NetProfitViews = True

Exit Function

Hell:
    
End Function

Private Function InsertRowAndCellSToTable(Table As HTMLTable, Optional CellsCount As Integer) As HTMLTableRow

Dim Cell As Integer
Dim Row As HTMLTableRow

With m_WebDoc

    'Set Row = Table.tBodies(0).insertRow(-1)
    
    Set Row = .createElement("TR")
    Row.Height = 18
    Table.tBodies(0).appendChild Row
    
    ' if user has not specified the cols then get the columns from the table
    If CellsCount = 0 Then CellsCount = Table.Cols
    
    For Cell = 0 To CellsCount - 1
    
        Set m_WebCell = .createElement("TD")
        With Row
            .appendChild m_WebCell
            .Style.fontFamily = gFontName
            .Style.FontSize = 13
            If gLangOffSet <> 0 Then .Style.FontSize = 15
        End With
    Next Cell
    
    m_WebCell.innerText = " "

End With

Set InsertRowAndCellSToTable = Row

End Function
Private Function BALPutAssetorLiability(rstSumParents As ADODB.Recordset, _
                                        rstOBHeads As ADODB.Recordset, _
                                        rstOBParents As ADODB.Recordset, _
                                        rstHeads As ADODB.Recordset, _
                                        isAsset As Boolean, _
                                        PrintStatus As wis_PrintStatus _
                                        ) As Currency
On Error GoTo Hell:

Dim SlNo As Integer

Dim ParentID As Long
Dim HeadID As Long
Dim HeadsCount As Long

Dim OpBalance As Currency
Dim ParentAmount As Currency
Dim Amount As Currency

Dim ParentName As String
Dim HeadName As String

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow

Const cnstHEAD = 0
Const cnstHEAD_AMOUNT = 1
Const cnstAMOUNT = 2

Const cnstLIABILITY = 0
Const cnstASSET = 1

Const cnstTOTALCOLS = 3

BALPutAssetorLiability = 0

ParentID = FormatField(rstOBParents.Fields("ParentID"))
OpBalance = FormatField(rstOBParents.Fields("ParentOB"))
ParentName = FormatField(rstOBParents.Fields("ParentName"))

' get the parent details
ParentAmount = 0
        
If rstSumParents.RecordCount > 0 Then If gDbTrans.FindRecord(rstSumParents, "ParentID=" & ParentID) Then _
        ParentAmount = FormatField(rstSumParents.Fields("SumAmount"))

ParentAmount = ParentAmount + OpBalance

If ParentAmount >= 0 Then
    
    If isAsset Then Set CurrentTable = m_WebRow.cells(cnstASSET).All.Asset
    
    If Not isAsset Then Set CurrentTable = m_WebRow.cells(cnstLIABILITY).All.Liability
    
Else
    
    If isAsset Then Set CurrentTable = m_WebRow.cells(cnstLIABILITY).All.Liability
    
    If Not isAsset Then Set CurrentTable = m_WebRow.cells(cnstASSET).All.Asset
    
End If
                        
' Start Putting Parentid's heads
Set ParentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)

ParentRow.Style.fontWeight = "bold"

With ParentRow.cells(cnstHEAD)
    .innerText = ParentName
    .Align = "Left"
End With

' Filter the table
rstOBHeads.Filter = adFilterNone
rstOBHeads.Filter = "ParentID=" & ParentID
                        
' This is total count
HeadsCount = rstHeads.RecordCount

SlNo = 0
Do While Not rstOBHeads.EOF
    
    OpBalance = FormatField(rstOBHeads.Fields("OpAmount"))
    HeadID = FormatField(rstOBHeads.Fields("OpHeadID"))
    HeadName = FormatField(rstOBHeads.Fields("HeadName"))
        
    Debug.Assert HeadID <> 93
    Amount = 0
        
    If HeadsCount > 0 Then If gDbTrans.FindRecord(rstHeads, "HeadID=" & HeadID) Then _
            Amount = FormatField(rstHeads.Fields("Amount"))
            
    Amount = Amount + OpBalance
          
    ' It Will Print Only if Detailed Print is Enabled
    
    If PrintStatus = PrintDetailed Then
        
        If Amount <> 0 Then
        
            SlNo = SlNo + 1
            
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)
            
            With CurrentRow.cells(cnstHEAD)
                .innerText = SlNo & "." & HeadName
                .Align = "Left"
            End With
            
            With CurrentRow.cells(cnstHEAD_AMOUNT)
                .innerText = FormatCurrency(Amount)
                .Align = "Right"
            End With
        End If
        
    End If
    
    rstOBHeads.MoveNext
Loop

' Now put the parenthead Total

With ParentRow.cells(cnstAMOUNT)
    .innerText = FormatCurrency(Abs(ParentAmount))
    .Align = "Right"
    .Style.fontWeight = "bold"
End With

BALPutAssetorLiability = ParentAmount

Exit Function

Hell:
    
End Function
Private Sub PrintTagsFromWebDocument()

Dim TagsCount As Integer
Dim Tag As Integer

With m_frmWeb.web.document
    TagsCount = .All.Length
    For Tag = 0 To TagsCount - 1
        Debug.Print .All(Tag).tagName
    Next Tag
End With

End Sub

Private Sub PutHtmlStyle()

With m_frmWeb.web.document

    .Write "<style>"
    .Write ".clsContentText         {font-family:verdana;font-size:x-small;padding-left:11px;text-align:center;}"
    
    .Write ".clsRowBold         {font-weight: bold}"
    
    .Write " P  {font: 10pt/12pt; verdana;color: black;"
    .Write "    font-weight: bold;"
    .Write "    text-align:center;"
    .Write "}"
    
    .Write " table {"
    .Write " font-weight: bold;"
    '.Write "TABLE-LAYOUT:fixed;"
    .Write "border-collapse:collapse;"
    .Write "cellPadding=0;width=80% "
    .Write "border=1;"
    .Write "}"
    
    .Write "tr {"
    .Write "font-weight: normal;"
    .Write "height=20;"
    .Write "align=center;"
    .Write "}"
    
    .Write "td {"
    .Write "font-weight: normal;"
    .Write "height=20;"
    .Write "width=50;"
    .Write "}"
    
    .Write "</style>"

End With

End Sub

Private Function PLPutIncomeorExpenses(rstSum As ADODB.Recordset, _
                               rstHeads As ADODB.Recordset, isExpense As Boolean) As Currency

On Error GoTo Hell:

Dim SlNo As Integer

Dim ParentID As Long
Dim HeadID As Long

Dim ParentAmount As Currency
Dim Amount As Currency

Dim ParentName As String
Dim HeadName As String

Dim isIncome As Boolean

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Const cnstHEAD = 0
Const cnstHEAD_AMOUNT = 1
Const cnstAMOUNT = 2

Const cnstEXPENSES = 0
Const cnstINCOME = 1

Const cnstTOTALCOLS = 3

PLPutIncomeorExpenses = 0

ParentID = FormatField(rstSum.Fields("ParentID"))
ParentAmount = FormatField(rstSum.Fields("SumAmount"))
ParentName = FormatField(rstSum.Fields("ParentName"))

If ParentAmount >= 0 Then

    If isExpense Then
        Set CurrentTable = m_WebRow.cells(cnstEXPENSES).All.Expense
    Else
        Set CurrentTable = m_WebRow.cells(cnstINCOME).All.Income
    End If
    
Else

    If isExpense Then
        Set CurrentTable = m_WebRow.cells(cnstINCOME).All.Income
    Else
        Set CurrentTable = m_WebRow.cells(cnstEXPENSES).All.Expense
    End If
    
End If
                        
' Start Putting parentid's heads
        
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)

        
With CurrentRow.cells(cnstHEAD)
    .innerText = ParentName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With


' Filter the table
rstHeads.Filter = adFilterNone
rstHeads.Filter = "ParentID=" & ParentID

SlNo = 0

Do While Not rstHeads.EOF

    
    Amount = FormatField(rstHeads.Fields("Amount"))
    HeadID = FormatField(rstHeads.Fields("HeadID"))
    HeadName = FormatField(rstHeads.Fields("HeadName"))
    
    If Amount <> 0 Then
    
        SlNo = SlNo + 1
        
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)
        
        With CurrentRow.cells(cnstHEAD)
            .innerText = SlNo & "." & HeadName
            .Align = "Left"
            .Style.FontSize = 13
            If gLangOffSet <> 0 Then .Style.FontSize = 15
        End With
        
        With CurrentRow.cells(cnstHEAD_AMOUNT)
            .innerText = FormatCurrency(Amount)
            .Align = "Right"
            .Style.FontSize = 13
            If gLangOffSet <> 0 Then .Style.FontSize = 15
        End With
        
    End If
    
    rstHeads.MoveNext
    
Loop

' Now put the parenthead Total

With CurrentRow.cells(cnstAMOUNT)
    .innerText = FormatCurrency(Abs(ParentAmount))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

PLPutIncomeorExpenses = ParentAmount

Exit Function

Hell:
    
End Function

Private Sub PLPutPandLTableStructure()

Dim CurrentRow As HTMLTableRow

Set CurrentRow = InsertRowAndCellSToTable(m_WebTable, 4)

With CurrentRow.cells(0)
    .innerText = "Expenses"
    .Align = "Left"
    .Width = "40%"
End With

With CurrentRow.cells(1)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "10%"
End With

With CurrentRow.cells(2)
    .innerText = "Income"
    .Align = "Left"
    .Width = "40%"
End With

With CurrentRow.cells(3)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "10%"
End With

End Sub



Private Sub PutTradingProfitORLoss(ByVal Purchase As Currency, _
                                   ByVal Sales As Currency)

Dim Profit As Currency

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

With m_frmWeb.web.document

    If Purchase >= Sales Then
    
        Profit = Purchase - Sales
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
                
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(2)
            .innerText = "Gross Loss"
            .Align = "Left"
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Profit)
            .Align = "Right"
        End With
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
                
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(0)
            .innerText = "Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Purchase)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(2)
            .innerText = "Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Purchase)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17

        End With
                
    Else
    
        Profit = Sales - Purchase
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(0)
            .innerText = "Gross Profit"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Profit)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(0)
            .innerText = "Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Sales)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(2)
            .innerText = "Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Sales)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FontSize = 15
            If gLangOffSet <> 0 Then .Style.FontSize = 17
        End With
    End If
    
End With


End Sub

Private Sub PLPutProfitORLoss(ByVal Income As Currency, _
                              ByVal Expenses As Currency)
                            
Dim Profit As Currency
Dim Total As Currency
Dim Cell As Integer

Const TotalCells = 4
If m_WebTable Is Nothing Then Exit Sub

With m_WebTable

    'just an extra line
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
    
    If Expenses >= Income Then
    
        Profit = Expenses - Income
                    
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
        m_WebRow.Style.fontWeight = "Bold"
        With m_WebRow.cells(2)
            .innerText = "Net Loss"
            .Align = "Left"
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Profit)
            .Align = "right"
        End With
                    
        Total = Expenses
    Else
    
        Profit = Income - Expenses
        Total = Income
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
        m_WebRow.Style.fontWeight = "Bold"
        With m_WebRow.cells(0)
            .innerText = "Net Profit"
            .Align = "Left"
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Profit)
            .Align = "right"
        End With
        
    End If
    
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)

' print the total
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
m_WebRow.Style.fontWeight = "Bold"
With m_WebRow.cells(0)
    .innerText = "Net Total"
    .Align = "Left"
End With

With m_WebRow.cells(1)
    .innerText = FormatCurrency(Total)
    .Align = "right"
End With

With m_WebRow.cells(2)
    .innerText = "Net Total"
    .Align = "Left"
End With

With m_WebRow.cells(3)
    .innerText = FormatCurrency(Total)
    .Align = "right"
End With

End Sub


Private Sub BALSetAssetHead()

Dim CurrentRow As HTMLTableRow
Dim AssetTable As HTMLTable

Set AssetTable = m_WebRow.cells(1).All.Asset
Set CurrentRow = InsertRowAndCellSToTable(AssetTable, 3)

AssetTable.Style.BorderColor = "Black"
CurrentRow.Style.fontWeight = "bold"

With CurrentRow.cells(0)
    .innerText = "Asset"
    .Align = "Left"
    .Width = "55%"
End With

With CurrentRow.cells(1)
    .innerText = "Head Total"
    .Align = "Left"
    .Width = "22%"
End With

With CurrentRow.cells(2)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "23%"
End With

End Sub
Private Sub PLSetIncomeHead()

Dim CurrentRow As HTMLTableRow
Dim IncomeTable As HTMLTable

Set IncomeTable = m_WebRow.cells(1).All.Income
Set CurrentRow = InsertRowAndCellSToTable(IncomeTable, 3)

'IncomeTable.Style.BorderColor = "Black"
With CurrentRow.cells(0)
    .innerText = "Income"
    .Align = "Left"
    .Width = "57%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(1)
    .innerText = "Head Total"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(2)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "23%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub

Private Sub BALSetLiabilityHeader()

Dim CurrentRow As HTMLTableRow
Dim LiabilityTable As HTMLTable

Set LiabilityTable = m_WebRow.cells(0).All.Liability
Set CurrentRow = InsertRowAndCellSToTable(LiabilityTable, 3)

LiabilityTable.Style.BorderColor = "Black"

CurrentRow.Style.fontWeight = "bold"

With CurrentRow.cells(0)
    .innerText = "Liability"
    .Align = "Left"
    .Width = "55%"
End With

With CurrentRow.cells(1)
    .innerText = "Head Total"
    .Align = "Left"
    .Width = "22%"
End With

With CurrentRow.cells(2)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "23%"
End With

End Sub
Private Sub PLSetExpenseHeader()

Dim CurrentRow As HTMLTableRow
Dim ExpenseTable As HTMLTable

Set ExpenseTable = m_WebRow.cells(0).All.Expense
Set CurrentRow = InsertRowAndCellSToTable(ExpenseTable, 3)

CurrentRow.Style.fontWeight = "bold"

With CurrentRow.cells(0)
    .innerText = "Expenses"
    .Align = "Left"
    .Width = "57%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(1)
    .innerText = "Head Total"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(2)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "23%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub

'' This will create the views required for the Receipts and Payments
Private Sub RPCreateViews(ByVal QryName As String)

' This Qurey will only load the Receipts and Payments details of the given period

gDbTrans.SQLStmt = " SELECT a.HeadID AS RPHeadID, b.debit AS Receipt, " & _
                   " b.Credit AS Payment" & _
                   " FROM QryAccTransMerge  " & _
                   " WHERE TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                   " AND " & "#" & FormatDate(ToRpDate) & "#" & _
                   " AND  a.HeadID<>b.Headid And a.Headid<>" & wis_CashHeadID
                 
gDbTrans.CreateView (QryName)


End Sub
'' This will create the views required for the Receipts and Payments
Private Sub SubDayCreateViews(ByVal QryName As String)

' This Qurey will only load the Receipts and Payments details of the given period

gDbTrans.SQLStmt = " SELECT a.HeadID AS RPHeadID, b.debit AS Receipt, " & _
                   " b.Credit AS Payment" & _
                   " FROM QryAccTransMerge,ParentHeads c  " & _
                   " WHERE TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                   " AND " & "#" & FormatDate(ToRpDate) & "#" & _
                   " AND  a.HeadID<>b.Headid And a.Headid<>" & wis_CashHeadID

gDbTrans.CreateView (QryName)


End Sub
'' This will create the views required for the Receipts and Payments
Private Sub GLCreateViews(ByVal QryName As String)

' This Qurey will only load the Receipts and Payments details of the given period

'gDbTrans.SqlStmt = " SELECT a.HeadID AS RPHeadID, b.debit AS Receipt, " & _
                   " b.Credit AS Payment" & _
                   " FROM QryAccTransMerge  " & _
                   " WHERE TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                   " AND " & "#" & FormatDate(ToRpDate) & "#" & _
                   " AND  a.HeadID<>b.Headid And a.Headid<>" & wis_CashHeadID
                 
gDbTrans.SQLStmt = " SELECT a.HeadID AS RPHeadID, a.Credit AS Receipt, " & _
                   " a.Debit AS Payment" & _
                   " FROM AccTrans a  " & _
                   " WHERE TransDate BETWEEN " & "#" & FormatDate(ToRpDate) & "#" & _
                   " AND " & "#" & FormatDate(ToRpDate) & "#" & _
                   " AND a.Headid<>" & wis_CashHeadID

gDbTrans.CreateView (QryName)


End Sub

Private Sub RPprintCash(ByVal PrintString, ByVal CashAmount As Currency, ByVal TableName As String)

If TableName = "" Then Exit Sub
If PrintString = "" Then Exit Sub

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Set CurrentTable = m_WebDoc.All(TableName)
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 4)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = PrintString
    .Align = "Left"
End With

If CashAmount >= 0 Then
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
    End With
Else
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
    End With
End If

End Sub

'This will print the cash Closing Balance
Private Sub RPputClosingCash(ByVal CashAmount As Currency, ByVal TableName As String)

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

' Validations

If TableName = "" Then Exit Sub

Set CurrentTable = m_WebTable

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

With CurrentRow.cells(1)
    .innerText = "Closing Balance"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

If CashAmount >= 0 Then
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FontSize = 15
        If gLangOffSet <> 0 Then .Style.FontSize = 17
    End With
Else
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FontSize = 15
        If gLangOffSet <> 0 Then .Style.FontSize = 17
    End With
End If

End Sub
Private Function RPPutOpeningCash(ByVal TableName As String, AsOnDate As String) As Currency

' Handle Error
On Error GoTo Hell:

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim OpeningCash As Currency

Dim AccTransClass As clsAccTrans

RPPutOpeningCash = 0

Set AccTransClass = New clsAccTrans

OpeningCash = AccTransClass.GetOpBalance(wis_CashHeadID, AsOnDate)

Set AccTransClass = Nothing

Set CurrentTable = m_WebTable
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = "Opening Cash Balance"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

If OpeningCash >= 0 Then
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(OpeningCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FontSize = 15
        If gLangOffSet <> 0 Then .Style.FontSize = 17
    End With
Else
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(OpeningCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FontSize = 15
        If gLangOffSet <> 0 Then .Style.FontSize = 17
    End With
End If

RPPutOpeningCash = OpeningCash

Exit Function

Hell:

End Function
Private Sub RPPutCashinField(ByVal TableName As String, _
                                 ByVal AsOnDate As String, _
                                 ByVal FieldNo As Long, _
                                 ByVal MsgStr As String, _
                                 ByVal Amount As Currency)

' Handle Error
On Error GoTo Hell:

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Set CurrentTable = m_WebTable
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = MsgStr
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With CurrentRow.cells(FieldNo)
    .innerText = FormatCurrency(Amount)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With


Exit Sub

Hell:

End Sub
Public Function FormatCurrency(ByVal Curr As Currency) As String
    FormatCurrency = Format(Abs(Curr), "##########0.00")
    If Curr < 0 Then FormatCurrency = "(-)" & FormatCurrency
    
End Function

Public Function ShowBalanceSheet() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations

Dim rstOBHeads As ADODB.Recordset

Dim rstLiability As ADODB.Recordset
Dim rstAsset As ADODB.Recordset

Dim rstSumLiability As ADODB.Recordset
Dim rstSumAsset As ADODB.Recordset

Dim rstOBLIabilityParents As ADODB.Recordset
Dim rstOBAssetParents As ADODB.Recordset

Dim AssetTotal As Currency
Dim LiabilityTotal As Currency
Dim ParentAmount As Currency
Dim Profit As Currency
Dim StockValue As Currency

Dim PrintStatus As wis_PrintStatus

Dim MaterialClass As clsMaterial

Const cnstLIABILITY = 0
Const cnstASSET = 1

' Intialiase the Function
ShowBalanceSheet = False

Profit = GetNetProfit

 
Call BALCreateViewsForBalanceSheet

' Opening Balances

gDbTrans.SQLStmt = " SELECT * FROM qryOPHeads"

If gDbTrans.Fetch(rstOBHeads, adOpenDynamic) < 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT SUM(OPAmount) as ParentOB,a.ParentID,b.ParentName" & _
                   " FROM qryOPHeads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType=2" & _
                   " GROUP BY a.ParentID,b.ParentName"
                 
If gDbTrans.Fetch(rstOBLIabilityParents, adOpenDynamic) < 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT SUM(OPAmount) as ParentOB,a.ParentID,b.ParentName" & _
                   " FROM qryOPHeads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType=1" & _
                   " GROUP BY a.ParentID,b.ParentName"
                 
If gDbTrans.Fetch(rstOBAssetParents, adOpenDynamic) < 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM qryLiabilities"
                   
Call gDbTrans.Fetch(rstLiability, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(Amount) as SumAmount,ParentID,ParentName" & _
                   " FROM QryLiabilities" & _
                   " GROUP BY ParentID,ParentName"
                   
Call gDbTrans.Fetch(rstSumLiability, adOpenDynamic)
                   
' Assets
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM qryAssets"
                   
Call gDbTrans.Fetch(rstAsset, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(Amount) as SumAmount,ParentID,ParentName" & _
                   " FROM QryAssets" & _
                   " GROUP BY ParentID,ParentName"
                   
Call gDbTrans.Fetch(rstSumAsset, adOpenDynamic)


' Start the Html
BALStartBalanceSheet

' Start the Both Headers
BALSetAssetHead
BALSetLiabilityHeader


' Main Loop
Do

    If rstOBAssetParents.EOF And rstOBLIabilityParents.EOF Then Exit Do
    
    If Not rstOBLIabilityParents.EOF Then
                
        PrintStatus = PrintDetailed
        
        ParentAmount = BALPutAssetorLiability(rstSumLiability, rstOBHeads, _
                                      rstOBLIabilityParents, rstLiability, False, PrintStatus)
                                              
        If ParentAmount >= 0 Then
            LiabilityTotal = LiabilityTotal + Abs(ParentAmount)
        Else
            AssetTotal = AssetTotal + Abs(ParentAmount)
        End If
        
    End If
    
    If Not rstOBAssetParents.EOF Then
        
        PrintStatus = PrintDetailed
        
        ParentAmount = BALPutAssetorLiability(rstSumAsset, rstOBHeads, _
                                    rstOBAssetParents, rstAsset, True, PrintStatus)
                
        If ParentAmount >= 0 Then
            AssetTotal = AssetTotal + Abs(ParentAmount)
        Else
            LiabilityTotal = LiabilityTotal + Abs(ParentAmount)
        End If
    End If
    
    If Not rstOBLIabilityParents.EOF Then rstOBLIabilityParents.MoveNext
    If Not rstOBAssetParents.EOF Then rstOBAssetParents.MoveNext
    
Loop


'Get the Closing Stock value from the database
Set MaterialClass = New clsMaterial

StockValue = MaterialClass.GetOnDateClosingStockValue(ToRpDate)

'print the closing stock value
BALPrintClosingStockValue (StockValue)

'Add the Stockvalue to the asset Total
AssetTotal = AssetTotal + StockValue

' Print the the Profit
BALPrintProfitorLoss (Profit)

If Profit >= 0 Then
    LiabilityTotal = LiabilityTotal + Abs(Profit)
Else
    AssetTotal = AssetTotal + Abs(Profit)
End If

' This will equalise the rows of both the tables
BALFillEmptyRows

Call BALSetBalanceSheetDifference(AssetTotal, LiabilityTotal)
    
If AssetTotal > LiabilityTotal Then

    Call BALPrintGrandTotals(AssetTotal)
    
ElseIf LiabilityTotal >= AssetTotal Then

    Call BALPrintGrandTotals(LiabilityTotal)
    
End If

Set rstOBHeads = Nothing
Set rstLiability = Nothing
Set rstAsset = Nothing
Set rstSumLiability = Nothing
Set rstSumAsset = Nothing
Set rstOBLIabilityParents = Nothing
Set rstOBAssetParents = Nothing
Set MaterialClass = Nothing


' Exit Function
ShowBalanceSheet = True

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "Balance Sheet :" & vbCrLf & Err.Description
    
    
    Set rstOBHeads = Nothing
    Set rstLiability = Nothing
    Set rstAsset = Nothing
    Set rstSumLiability = Nothing
    Set rstSumAsset = Nothing
    Set rstOBLIabilityParents = Nothing
    Set rstOBAssetParents = Nothing
    
End Function
Public Function ShowConditionBalanceSheet() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations

Dim rstOBHeads As ADODB.Recordset

Dim rstLiability As ADODB.Recordset
Dim rstAsset As ADODB.Recordset

Dim rstSumLiability As ADODB.Recordset
Dim rstSumAsset As ADODB.Recordset

Dim rstOBLIabilityParents As ADODB.Recordset
Dim rstOBAssetParents As ADODB.Recordset

Dim AssetTotal As Currency
Dim LiabilityTotal As Currency
Dim ParentAmount As Currency
Dim Profit As Currency
Dim StockValue As Currency

Dim MaterialClass As clsMaterial
Dim PrintStatus As wis_PrintStatus

Const cnstLIABILITY = 0
Const cnstASSET = 1

' Intialiase the Function
ShowConditionBalanceSheet = False

Profit = GetNetProfit

 
Call BALCreateViewsForConditionalBalanceSheet

' Opening Balances

gDbTrans.SQLStmt = " SELECT * FROM qryOPHeads"

If gDbTrans.Fetch(rstOBHeads, adOpenDynamic) < 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT SUM(OPAmount) as ParentOB,a.ParentID,b.ParentName," & _
                   " b.PrintOrder,b.PrintDetailed" & _
                   " FROM qryOPHeads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType=2" & _
                   " GROUP BY b.PrintOrder,a.ParentID,b.ParentName,b.PrintDetailed"
                 
If gDbTrans.Fetch(rstOBLIabilityParents, adOpenDynamic) < 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT SUM(OPAmount) as ParentOB,a.ParentID,b.ParentName," & _
                   " b.PrintOrder,b.PrintDetailed" & _
                   " FROM qryOPHeads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType=1" & _
                   " GROUP BY b.PrintOrder,a.ParentID,b.ParentName,b.PrintDetailed"
                 
If gDbTrans.Fetch(rstOBAssetParents, adOpenDynamic) < 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM qryLiabilities"
                   
Call gDbTrans.Fetch(rstLiability, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(Amount) as SumAmount,ParentID" & _
                   " FROM QryLiabilities" & _
                   " GROUP BY ParentID"
                   
Call gDbTrans.Fetch(rstSumLiability, adOpenDynamic)
                   
' Assets
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM qryAssets"
                   
Call gDbTrans.Fetch(rstAsset, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(Amount) as SumAmount,ParentID" & _
                   " FROM QryAssets" & _
                   " GROUP BY ParentID"
                   
Call gDbTrans.Fetch(rstSumAsset, adOpenDynamic)


' Start the Html
BALStartBalanceSheet

' Start the Both Headers
BALSetAssetHead
BALSetLiabilityHeader

' Main Loop
Do

    If rstOBAssetParents.EOF And rstOBLIabilityParents.EOF Then Exit Do
    
    If Not rstOBLIabilityParents.EOF Then
        
        PrintStatus = rstOBLIabilityParents.Fields("PrintDetailed")
        
        ParentAmount = BALPutAssetorLiability(rstSumLiability, rstOBHeads, _
                                      rstOBLIabilityParents, rstLiability, False, PrintStatus)
                                              
        If ParentAmount >= 0 Then
            LiabilityTotal = LiabilityTotal + Abs(ParentAmount)
        Else
            AssetTotal = AssetTotal + Abs(ParentAmount)
        End If
        
    End If
    
    If Not rstOBAssetParents.EOF Then
        
        PrintStatus = rstOBAssetParents.Fields("PrintDetailed")
        
        ParentAmount = BALPutAssetorLiability(rstSumAsset, rstOBHeads, _
                                    rstOBAssetParents, rstAsset, True, PrintStatus)
                
        If ParentAmount >= 0 Then
            AssetTotal = AssetTotal + Abs(ParentAmount)
        Else
            LiabilityTotal = LiabilityTotal + Abs(ParentAmount)
        End If
    End If
    
    If Not rstOBLIabilityParents.EOF Then rstOBLIabilityParents.MoveNext
    If Not rstOBAssetParents.EOF Then rstOBAssetParents.MoveNext
    
Loop


'Get the Closing Stock value from the database
Set MaterialClass = New clsMaterial

StockValue = MaterialClass.GetOnDateClosingStockValue(ToRpDate)

'print the closing stock value
BALPrintClosingStockValue (StockValue)

'Add the Stockvalue to the asset Total
AssetTotal = AssetTotal + StockValue

' Print the the Profit
BALPrintProfitorLoss (Profit)

If Profit >= 0 Then
    LiabilityTotal = LiabilityTotal + Abs(Profit)
Else
    AssetTotal = AssetTotal + Abs(Profit)
End If

' This will equalise the rows of both the tables
BALFillEmptyRows

Call BALSetBalanceSheetDifference(AssetTotal, LiabilityTotal)
    
If AssetTotal > LiabilityTotal Then

    Call BALPrintGrandTotals(AssetTotal)
    
ElseIf LiabilityTotal >= AssetTotal Then

    Call BALPrintGrandTotals(LiabilityTotal)
    
End If

Set rstOBHeads = Nothing
Set rstLiability = Nothing
Set rstAsset = Nothing
Set rstSumLiability = Nothing
Set rstSumAsset = Nothing
Set rstOBLIabilityParents = Nothing
Set rstOBAssetParents = Nothing
Set MaterialClass = Nothing


' Exit Function
ShowConditionBalanceSheet = True

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "ShowConditionBalanceSheet :" & vbCrLf & Err.Description
        
    Set rstOBHeads = Nothing
    Set rstLiability = Nothing
    Set rstAsset = Nothing
    Set rstSumLiability = Nothing
    Set rstSumAsset = Nothing
    Set rstOBLIabilityParents = Nothing
    Set rstOBAssetParents = Nothing
    
End Function

'This will equaliase the row heights
Private Sub BALEqualRowHeights()

Dim CurrentRow As HTMLTableRow
Dim AssetTable As HTMLTable
Dim LiabilityTable As HTMLTable

Dim SpaceToAdd As Integer

Dim Rows As Integer
Dim Row As Integer

With m_WebRow
    
    If .cells(0).All.Liability.Rows.Length <> .cells(1).All.Asset.Rows.Length Then Exit Sub
    
    Set LiabilityTable = .cells(0).All.Liability
    Set AssetTable = .cells(1).All.Asset
    
    Rows = .cells(0).All.Liability.Rows.Length
    
End With


For Row = 0 To Rows - 1
    
    'If Len(LiabilityTable.Rows(Row).cells(1).innerText) <> Len(AssetTable.Rows(Row).cells(1).innerText) Then
                    
    SpaceToAdd = Len(LiabilityTable.Rows(Row).cells(0).innerText) - Len(AssetTable.Rows(Row).cells(0).innerText)
    
    If SpaceToAdd <> 0 Then
        If SpaceToAdd < 0 Then
            LiabilityTable.Rows(Row).cells(0).innerText = LiabilityTable.Rows(Row).cells(0).innerText & Space(Abs(SpaceToAdd))
        Else
            AssetTable.Rows(Row).cells(0).innerText = AssetTable.Rows(Row).cells(0).innerText & Space(Abs(SpaceToAdd))
        End If
    End If
Next Row
    
Set m_WebRow = InsertRowAndCellSToTable(LiabilityTable, 3)

End Sub

Public Function ShowPandLAccount() As Currency

' Handle Error
On Error GoTo Hell:

' Declarations

Dim rstExpense As ADODB.Recordset
Dim rstIncome As ADODB.Recordset

Dim rstSumExpense As ADODB.Recordset
Dim rstSumIncome As ADODB.Recordset

Dim IncomeTotal As Currency
Dim ExpenseTotal As Currency
Dim ParentAmount As Currency
Dim Profit As Currency
Dim GrossProfit As Currency

Const cnstExpense = 0
Const cnstINCOME = 1

' Intialiase the Function
ShowPandLAccount = 0

Set m_WebDoc = Nothing

' Create Required Views for the Balance Sheet Report

'GrossProfit = ShowTradingAccount
GrossProfit = GetTradingProfit

Call PLCreateViewsForProfitAndLoss

' Opening Balances

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM QryExpenses"
                   
Call gDbTrans.Fetch(rstExpense, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(Amount) as SumAmount,ParentID,ParentName" & _
                   " FROM QryExpenses" & _
                   " GROUP BY ParentID,ParentName"
                   
Call gDbTrans.Fetch(rstSumExpense, adOpenDynamic)

                   
' Assets

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM QryIncomes"
                   
Call gDbTrans.Fetch(rstIncome, adOpenDynamic)

gDbTrans.SQLStmt = " SELECT SUM(Amount) as SumAmount,ParentID,ParentName" & _
                   " FROM QryIncomes" & _
                   " GROUP BY ParentID,ParentName"
                   
Call gDbTrans.Fetch(rstSumIncome, adOpenDynamic)


' Start the Html
PLStartProfitAndLoss

' Start the Both Headers
PLSetExpenseHeader
PLSetIncomeHead

If GrossProfit <= 0 Then
    
    ' If profit is loss
       
    Call PLPrintProfitorLoss(Abs(GrossProfit), "Opening Balance", False)
    
    ExpenseTotal = ExpenseTotal + Abs(GrossProfit)
    
Else

    'if profit is Profit
    Call PLPrintProfitorLoss(Abs(GrossProfit), "Opening Balance", True)
    
    IncomeTotal = IncomeTotal + Abs(GrossProfit)

End If

' This will equalise the rows of both the tables
PLFillEmptyRows

' Main Loop
Do

    If rstSumIncome.EOF And rstSumExpense.EOF Then Exit Do
    
    If Not rstSumExpense.EOF Then
    
        ParentAmount = PLPutIncomeorExpenses(rstSumExpense, rstExpense, True)
        
        If ParentAmount > 0 Then
            ExpenseTotal = ExpenseTotal + Abs(ParentAmount)
        Else
            IncomeTotal = IncomeTotal + Abs(ParentAmount)
        End If
        
    End If
    
    If Not rstSumIncome.EOF Then
    
        ParentAmount = PLPutIncomeorExpenses(rstSumIncome, rstIncome, False)
        If ParentAmount > 0 Then
            IncomeTotal = IncomeTotal + Abs(ParentAmount)
        Else
            ExpenseTotal = ExpenseTotal + Abs(ParentAmount)
        End If
    End If
    
    
    If Not rstSumExpense.EOF Then rstSumExpense.MoveNext
    If Not rstSumIncome.EOF Then rstSumIncome.MoveNext
    
Loop
  
'incometotal=grossprofit+
Profit = IncomeTotal - ExpenseTotal

' Print the the Profit
If Profit <= 0 Then
    ' If profit is loss
    Call PLPrintProfitorLoss(Abs(Profit), "Net Loss", True)
Else
    'if profit is Profit
    Call PLPrintProfitorLoss(Abs(Profit), "Net Profit", False)
End If


' This will equalise the rows of both the tables
PLFillEmptyRows

If IncomeTotal > ExpenseTotal Then

    Call PLPrintGrandTotals(IncomeTotal)
    
ElseIf ExpenseTotal >= IncomeTotal Then

    Call PLPrintGrandTotals(ExpenseTotal)
    
End If

Set rstExpense = Nothing
Set rstIncome = Nothing

Set rstSumExpense = Nothing
Set rstSumIncome = Nothing


' Exit Function
ShowPandLAccount = Profit


Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "ShowPandLAccount :" & vbCrLf & Err.Description
        
    Set rstExpense = Nothing
    Set rstIncome = Nothing
    
    Set rstSumExpense = Nothing
    Set rstSumIncome = Nothing
    
End Function
'This will return the net proft or net loss
Private Function GetNetProfit() As Currency

' Handle Error
On Error GoTo Hell:

' Declarations
Dim rstExpense As ADODB.Recordset
Dim rstIncome As ADODB.Recordset

Dim IncomeTotal As Currency
Dim ExpenseTotal As Currency
Dim Profit As Currency
Dim GrossProfit As Currency

' Intialiase the Function
GetNetProfit = 0

' Create Required Views for the Balance Sheet Report

'GrossProfit = ShowTradingAccount
GrossProfit = GetTradingProfit

Call NetProfitViews

' Opening Balances

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM QryExpenses"
                   
Call gDbTrans.Fetch(rstExpense, adOpenDynamic)

                   
' Assets

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM QryIncomes"
                   
Call gDbTrans.Fetch(rstIncome, adOpenDynamic)

ExpenseTotal = FormatField(rstExpense.Fields("NetExpenses"))
IncomeTotal = FormatField(rstIncome.Fields("NetIncome"))

If GrossProfit <= 0 Then
    
    ' If profit is loss
    
        
    ExpenseTotal = ExpenseTotal + Abs(GrossProfit)
    
Else

    'if profit is Profit
    
    IncomeTotal = IncomeTotal + Abs(GrossProfit)

End If
  
'incometotal=grossprofit+
Profit = IncomeTotal - ExpenseTotal

Set rstExpense = Nothing
Set rstIncome = Nothing

' Exit Function
GetNetProfit = Profit

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "ShowPandLAccount :" & vbCrLf & Err.Description
       
    Set rstExpense = Nothing
    Set rstIncome = Nothing
    
    
End Function

Public Function ShowReportDate() As Boolean

On Error GoTo Hell:

ShowReportDate = False

Set m_frmRptDt = New frmRptDt

m_frmRptDt.txtStDate = FinIndianFromDate
m_frmRptDt.txtEndDate = FinIndianEndDate

' User should not change the opening date
'm_frmRptDt.txtStDate.Locked = True

m_frmRptDt.Show vbModal

Set m_frmRptDt = Nothing

' If no date is entered then
If FromRpDate = "" And ToRpDate = "" Then Exit Function

LoadFrmWeb
IntialiseWebDocument

ShowReportDate = True

Exit Function

Hell:

    Set m_frmRptDt = Nothing
    
End Function
Public Function ShowReportForm() As Boolean

On Error GoTo Hell:

ShowReportForm = False

If m_frmAccReports Is Nothing Then Set m_frmAccReports = New frmAccReports

FromRpDate = ""
ToRpDate = ""

m_frmAccReports.txtDate1 = FinIndianFromDate
m_frmAccReports.txtDate2 = FinIndianEndDate

LoadFrmWeb

m_frmAccReports.Show vbModal
    
' If No Date is Entered Then
If FromRpDate = "" And ToRpDate = "" Then
    Set m_frmAccReports = Nothing
    Exit Function
End If

' This Will Initialise the Web Document
IntialiseWebDocument

' this will process the selected report and will feed it to webcontrol

If Not ProcessSelectedReport Then Exit Function

ShowReportForm = True

ShowCurrentReport

Set m_WebDoc = Nothing
Set m_frmWeb = Nothing


Exit Function

Hell:
    MsgBox " There is error"
    
    'Set m_frmAccReports = Nothing
        
End Function

Private Sub PauseApplication()

Dim Start

Start = Timer
Do
    'it will be in the loop
    If Timer >= Start + 1 Then Exit Do
Loop



End Sub

Public Sub ShowCurrentReport()

With m_frmWeb
    .Show vbModal
End With


End Sub

Public Function ShowRPReport() As Boolean

Dim rstRPDetails As ADODB.Recordset
Dim rstParents As ADODB.Recordset

Dim OpeningCash As Currency
Dim ClosingStock As Currency
Dim Receipts As Currency
Dim Payments As Currency
Dim SumReceipts As Currency
Dim SumPayments As Currency

Dim ParentID As Long
Dim HeadID As Long
Dim SlNo As Long

Dim ParentName As String
Dim HeadName As String

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Const TableName = "RPTable"
Const QRYRP = "QryRPDetails"
Const TOTALCOLS = 4

' Handle Error
On Error GoTo Hell:

ShowRPReport = False

' Start the Receipts and Payments
RPStartRP

' Put the Opening Cash to the Html
OpeningCash = RPPutOpeningCash(TableName, FromRpDate)

' create views for the receipts and payments
RPCreateViews (QRYRP)

gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SQLStmt = " SELECT Sum(a.Receipt) AS Receipts, SUM(a.Payment) AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName " & _
                   " FROM " & QRYRP & " AS a, Heads AS b " & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " GROUP BY b.ParentID,a.RPHeadID,b.HeadName"

gDbTrans.SQLStmt = "  SELECT SUM(Debit) As Payments,SUM(Credit) as Receipts,a.Headid as RpHeadID,b.HeadName,b.ParentID" & _
                   " FROM AccTrans a,Heads b,ParentHeads c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND a.HeadID <> " & wis_CashHeadID & _
                   " AND TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                   " AND " & "#" & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY a.HeadId,b.HeadName,b.ParentID " & _
                   " ORDER BY b.ParentID,a.HeadID"
                   
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)

Set CurrentTable = m_WebTable

SlNo = 0

' Main Loop
Do While Not rstParents.EOF
    
    ParentID = FormatField(rstParents.Fields("ParentID"))
    ParentName = FormatField(rstParents.Fields("ParentName"))
    
    
    With rstRPDetails
        
        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID
        
        If rstRPDetails.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, TOTALCOLS)
            
            CurrentRow.Style.fontWeight = "Bold"
            
            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
                .Style.fontWeight = "Bold"
                .Style.FontSize = 14
                If gLangOffSet <> 0 Then .Style.FontSize = 16
            End With
        End If
        Do While Not .EOF
            
            HeadID = FormatField(.Fields("RPHeadID"))
            HeadName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
            
            'If Receipts <> 0 And Payments <> 0 Then
                SlNo = SlNo + 1
                Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, TOTALCOLS)
                With CurrentRow.cells(0)
                    .innerText = SlNo
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FontSize = 13
                    If gLangOffSet <> 0 Then .Style.FontSize = 15
                End With
                
                With CurrentRow.cells(1)
                    .innerText = HeadName
                    .Align = "Left"
                    .Style.fontWeight = "Normal"
                    .Style.FontSize = 13
                    If gLangOffSet <> 0 Then .Style.FontSize = 15
                End With
                
                With CurrentRow.cells(2)
                    .innerText = FormatCurrency(Receipts)
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FontSize = 13
                    If gLangOffSet <> 0 Then .Style.FontSize = 15
                End With
                
                With CurrentRow.cells(3)
                    .innerText = FormatCurrency(Payments)
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FontSize = 13
                    If gLangOffSet <> 0 Then .Style.FontSize = 15
                End With
            'End If
            
            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments
            
            .MoveNext
        Loop
    End With
    
    rstParents.MoveNext
    
Loop


' Now Put Closing Balance
Call RPputClosingCash((OpeningCash + SumReceipts) - SumPayments, TableName)


Set rstRPDetails = Nothing
Set rstParents = Nothing

ShowRPReport = True

Exit Function

Hell:
    MsgBox "ShowRPReport: " & vbCrLf & Err.Description
    
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    
End Function


' This function will show the TradingAccount
'Lingappa Sindhanur
Public Function ShowTradingAccount() As Currency

' Handle Error
On Error GoTo Hell:

Dim rstTradingIncome As ADODB.Recordset
Dim rstTradingExpense As ADODB.Recordset

Dim Profit As Currency
Dim Purchase As Currency
Dim Sales As Currency
Dim OpeningStock As Currency
Dim ClosingStock As Currency

Dim TotalPurchase As Currency
Dim TotalSales As Currency

Dim IncomeHeadName As String
Dim ExpenseHeadName As String

Dim MaterialClass As clsMaterial

' Intialiase the Function
ShowTradingAccount = 0


gDbTrans.SQLStmt = " SELECT (SUM(Debit)-SUM(Credit)) AS Purchase,A.HeadID,HeadName" & _
                    " FROM AccTrans AS a, Heads AS b " & _
                    " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_PurchaseParentID & _
                    " OR b.ParentID  = " & wis_TradingExpParentID & " ) " & _
                    " AND  a.TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                    " AND # " & FormatDate(ToRpDate) & "#" & _
                    " GROUP BY A.HeadID, HeadName "

Call gDbTrans.Fetch(rstTradingExpense, adOpenForwardOnly)

gDbTrans.SQLStmt = " SELECT (SUM(Credit)-SUM(Debit)) AS Sales ,A.HeadID,HeadName" & _
                   " FROM AccTrans AS a, Heads AS b " & _
                   " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_SalesParentID & _
                   " OR b.ParentID  = " & wis_TradingIncParentID & " ) " & _
                   " AND  a.TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                   " AND # " & FormatDate(ToRpDate) & "#" & _
                   " GROUP BY A.HeadID, HeadName "

Call gDbTrans.Fetch(rstTradingIncome, adOpenForwardOnly)


Set MaterialClass = New clsMaterial

OpeningStock = MaterialClass.GetOpeningStockValue
ClosingStock = MaterialClass.GetOnDateClosingStockValue(ToRpDate)

Set MaterialClass = Nothing

' Start the Page and set the Headers
Call TRStartHtmlPage

Call TRPutOpeningStock(OpeningStock)

Do

    If rstTradingExpense.EOF And rstTradingIncome.EOF Then Exit Do
    
    
    'Get expense Details
    If Not rstTradingExpense.EOF Then
        Purchase = FormatField(rstTradingExpense.Fields("Purchase"))
        ExpenseHeadName = FormatField(rstTradingExpense.Fields("HeadName"))
    End If
    
    'Get Income  Details
    If Not rstTradingIncome.EOF Then
        Sales = FormatField(rstTradingIncome.Fields("Sales"))
        IncomeHeadName = FormatField(rstTradingIncome.Fields("HeadName"))
    End If
    
    Call TRPutSalesPurchases(Purchase, Sales, ExpenseHeadName, IncomeHeadName)
    ExpenseHeadName = ""
    IncomeHeadName = ""
    TotalPurchase = TotalPurchase + Purchase: Purchase = 0
    TotalSales = TotalSales + Sales: Sales = 0 '+ ClosingStock

    If Not rstTradingExpense.EOF Then rstTradingExpense.MoveNext
    If Not rstTradingIncome.EOF Then rstTradingIncome.MoveNext
    
Loop


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

Call TRPutClosingStock(ClosingStock)


TotalPurchase = TotalPurchase + OpeningStock
TotalSales = TotalSales + ClosingStock
Call PutTradingProfitORLoss(TotalPurchase, TotalSales)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

ShowTradingAccount = TotalSales - TotalPurchase

Exit Function

Hell:

    If Err.Number > 0 Then MsgBox "Trading Account:" & vbCrLf & Err.Description
    
    
End Function
' This function will return the Trading Profit or Loss
'Lingappa Sindhanur
Private Function GetTradingProfit() As Currency

' Handle Error
On Error GoTo Hell:

Dim rstTradingIncome As ADODB.Recordset
Dim rstTradingExpense As ADODB.Recordset

Dim OpeningStock As Currency
Dim ClosingStock As Currency
Dim TotalPurchase As Currency
Dim TotalSales As Currency

Dim MaterialClass As clsMaterial

' Intialiase the Function
GetTradingProfit = 0

gDbTrans.SQLStmt = " SELECT (SUM(Debit)-SUM(Credit)) As TradingExpense" & _
                    " FROM AccTrans AS a, Heads AS b " & _
                    " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_PurchaseParentID & _
                    " OR b.ParentID  = " & wis_TradingExpParentID & " ) " & _
                    " AND  a.TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                    " AND # " & FormatDate(ToRpDate) & "#"
                    
Call gDbTrans.Fetch(rstTradingExpense, adOpenForwardOnly)

gDbTrans.SQLStmt = " SELECT (SUM(Credit)-SUM(Debit)) AS TradingIncome" & _
                   " FROM AccTrans AS a, Heads AS b " & _
                   " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_SalesParentID & _
                   " OR b.ParentID  = " & wis_TradingIncParentID & " ) " & _
                   " AND  a.TransDate BETWEEN " & "#" & FormatDate(FromRpDate) & "#" & _
                   " AND # " & FormatDate(ToRpDate) & "#"

Call gDbTrans.Fetch(rstTradingIncome, adOpenForwardOnly)

TotalPurchase = FormatField(rstTradingExpense.Fields("TradingExpense"))
TotalSales = FormatField(rstTradingIncome.Fields("TradingIncome"))

Set MaterialClass = New clsMaterial

OpeningStock = MaterialClass.GetOpeningStockValue
ClosingStock = MaterialClass.GetOnDateClosingStockValue(ToRpDate)

Set MaterialClass = Nothing

TotalPurchase = TotalPurchase + OpeningStock
TotalSales = TotalSales + ClosingStock

GetTradingProfit = TotalSales - TotalPurchase

Exit Function

Hell:

    If Err.Number > 0 Then MsgBox "Trading Account:" & vbCrLf & Err.Description
    
    
End Function

Private Sub BALStartBalanceSheet()

Dim strHeader As String
Dim DOMElement
Dim DOMTbody

' set the web document to htmldoc

If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document


With m_WebDoc
    
    Set m_WebRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    strHeader = "Balance Sheet for the Duration"
    With m_WebRow.cells(0)
        .innerText = strHeader
        .Style.FontSize = "14"
        .Style.fontWeight = "Bold"
        .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    strHeader = "From : " & FromRpDate & "  To : " & ToRpDate
    With m_WebRow.cells(0)
        .innerText = strHeader
        .Style.FontSize = "14"
        .Style.fontWeight = "Bold"
        .Align = "Center"
    End With
    
    Set m_WebTable = .createElement("Table")
    
    Set DOMElement = .createElement("Tbody")
    m_WebTable.appendChild DOMElement
        

End With

With m_WebTable
    
    .border = 0
    .Width = "900"
    .cellPadding = 0
    .cellSpacing = 0
    .id = "BalanceSheet"
    .Align = "Center"
    
End With

m_WebDoc.body.appendChild m_WebTable

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 2)

With m_WebRow.cells(0)
    .innerText = ""
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
    
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .id = "Liability"
        
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .borderRightStyle = "none"
            .FontSize = 13
            .paddingLeft = 5
            .paddingRight = 5
        End With
        
    End With
    
     .vAlign = "Top"
     .Width = "50%"
End With

With m_WebRow.cells(1)

    .innerText = ""
    
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .id = "Asset"
        
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .borderLeftStyle = "none"
            .FontSize = 13
            .paddingLeft = 5
            .paddingRight = 5
        End With
        
    End With
     
     .vAlign = "Top"
     .Width = "50%"
     
End With

End Sub
Private Sub RPStartRP()

Dim strHeader As String
Dim HtmlItem

Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc

    Set HtmlItem = .createElement("H3")
   .body.appendChild HtmlItem
    strHeader = "Receipts and Payments for the Duration"
    With HtmlItem
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FontSize = "17"
        If gLangOffSet <> 0 Then .Style.FontSize = 19
    End With
    
    Set HtmlItem = .createElement("H3")
   
    strHeader = "From : " & FromRpDate & "  To : " & ToRpDate
    .body.appendChild HtmlItem
    With HtmlItem
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FontSize = "17"
        If gLangOffSet <> 0 Then .Style.FontSize = 19
    End With
    
    
    Set m_WebTable = .createElement("Table")
   
    With m_WebTable
        .id = "RPTable"
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FontSize = 10
        .Width = "650"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .Cols = 4
    End With
    
    Set m_WebTBody = .createElement("TBody")
    
    m_WebTable.appendChild m_WebTBody
    
    .body.appendChild m_WebTable
    
End With
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

With m_WebRow.cells(0)
    .innerText = "Sl No"
    .Align = "Left"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With
With m_WebRow.cells(1)
    .innerText = "Particualrs"
    .Align = "Left"
    .Width = "50%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With
With m_WebRow.cells(2)
    .innerText = "Receipts"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With
With m_WebRow.cells(3)
    .innerText = "Payments"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub

Private Sub PLStartProfitAndLoss()

Dim strHeader As String
Dim DOMElement
Dim DOMTbody

' set the web document to htmldoc

Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc
    
    Set DOMElement = .createElement("H3")
    strHeader = "Profit and Loss for the Duration"
    .body.appendChild DOMElement
    With DOMElement
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FontSize = "17"
        If gLangOffSet <> 0 Then .Style.FontSize = 19
    End With
    Set DOMElement = .createElement("H3")
    strHeader = "From : " & FromRpDate & "  To : " & ToRpDate
    .body.appendChild DOMElement
    With DOMElement
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FontSize = "17"
        If gLangOffSet <> 0 Then .Style.FontSize = 19
    End With
    Set m_WebTable = .createElement("Table")
    
    Set DOMElement = .createElement("Tbody")
    m_WebTable.appendChild DOMElement
        
End With

With m_WebTable

    .border = 0
    .Width = "900"
    .cellPadding = 0
    .cellSpacing = 0
    .id = "PLTable"
    .Align = "Center"

End With

m_WebDoc.body.appendChild m_WebTable

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 2)

With m_WebRow.cells(0)
    
    .innerText = ""
    
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .id = "Expense"
        .Align = "Center"
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .FontSize = 13
            .paddingLeft = 3
            .paddingRight = 3
        End With
        
    End With
        
     .vAlign = "Top"
     .Width = "50%"
End With

With m_WebRow.cells(1)
    .innerText = ""
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .id = "Income"
        .Align = "Center"
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .FontSize = 13
            .paddingLeft = 3
            .paddingRight = 3
        End With
    End With
        
     .vAlign = "Top"
     .Width = "50%"
End With

End Sub
Private Sub StartHtmlPage()

With m_frmWeb.web
        .document.Write "<Html>" & "<Body>"
End With

End Sub

Private Sub EndHtmlPage()

With m_frmWeb.web
        .document.Write "</Html>" & "</Body>"
End With

End Sub

Private Sub StartHtmlTable(ByVal TableID As String, ObjectToAppend, ByVal ColsCount As Integer)

With m_WebDoc

    Set m_WebTable = .createElement("Table")
    
    With m_WebTable
        .id = TableID
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FontSize = 10
        .Width = "650"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .Cols = ColsCount
    End With
    
    ObjectToAppend.appendChild m_WebTable
    
    Set m_WebTBody = .createElement("TBody")
    m_WebTable.appendChild m_WebTBody
    
    
    
End With

End Sub
Private Sub EndHtmlTable()

m_frmWeb.web.document.Write "</TABLE>"

End Sub

Private Sub TRPutOpeningStock(OpeningStock As Currency)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(0)
    .innerText = "Opening Stock"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With m_WebRow.cells(1)
    .innerText = FormatCurrency(OpeningStock)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub
Private Sub TRPutClosingStock(ClosingStock As Currency)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
m_WebRow.Style.fontWeight = "Bold"
With m_WebRow.cells(2)
    .innerText = "Closing Stock"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With m_WebRow.cells(3)
    .innerText = FormatCurrency(ClosingStock)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
    If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub

Private Sub TRPutSalesPurchases(Purchase As Currency, Sales As Currency, _
                                                    ByVal PurchaseHead As String, ByVal SalesHead As String)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

If PurchaseHead <> "" Then
    With m_WebRow.cells(0)
        .innerText = PurchaseHead
        .Align = "Left"
        .Style.fontWeight = "Normal"
        .Style.FontSize = 13
        If gLangOffSet <> 0 Then .Style.FontSize = 15
    End With
End If
If Purchase > 0 Then
    With m_WebRow.cells(1)
        .innerText = FormatCurrency(Purchase)
        .Align = "Right"
        .Style.fontWeight = "Normal"
        .Style.FontSize = 13
        If gLangOffSet <> 0 Then .Style.FontSize = 15
    End With
End If

If SalesHead <> "" Then
    With m_WebRow.cells(2)
        .innerText = SalesHead
        .Align = "Left"
        .Style.fontWeight = "Normal"
        .Style.FontSize = 13
        If gLangOffSet <> 0 Then .Style.FontSize = 15
    End With
End If

If Sales > 0 Then
    With m_WebRow.cells(3)
        .innerText = FormatCurrency(Sales)
        .Align = "Right"
        .Style.fontWeight = "Normal"
        .Style.FontSize = 13
        If gLangOffSet <> 0 Then .Style.FontSize = 15
    End With
End If
End Sub
Private Sub TRStartHtmlPage()

' This will put the trading header

Dim eTbody
Dim DOMElement

Set m_WebDoc = m_frmWeb.web.document

m_WebDoc.body.Style.fontFamily = gFontName

With m_WebDoc
    
    Set DOMElement = .createElement("H3")
    .body.appendChild DOMElement
    With DOMElement
        .innerText = "Trading Account From " & FromRpDate & "  To " & ToRpDate
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FontSize = "17"
        If gLangOffSet <> 0 Then .Style.FontSize = 19
    End With
    
    Set m_WebTable = .createElement("Table")
    Set eTbody = .createElement("TBody")
    
End With

Call StartHtmlTable("Trading", m_WebDoc.body, 4)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(0)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "35%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With m_WebRow.cells(1)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "15%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With m_WebRow.cells(2)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "35%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

With m_WebRow.cells(3)
    .innerText = "Amount"
    .Align = "Left"
    .Width = "15%"
    .Style.fontWeight = "Bold"
    .Style.FontSize = 15
     If gLangOffSet <> 0 Then .Style.FontSize = 17
End With

End Sub


' This function is used to test the html functions
Public Sub WebHtml()

Dim rstLedger As ADODB.Recordset
Dim HeadID As Long
Dim FldCount As Integer
Dim FldLength As Integer

gDbTrans.SQLStmt = " SELECT * FROM AccTrans"

'Fetch the Data to Recordset based on the Condition
Call gDbTrans.Fetch(rstLedger, adOpenForwardOnly)

StartTimer

FldLength = rstLedger.Fields.Count - 1

Do While Not rstLedger.EOF
    For FldCount = 0 To FldLength
        HeadID = FormatField(rstLedger.Fields(0))
    Next FldCount
    rstLedger.MoveNext
Loop

StopTimer

Call rstLedger.Save("c:\windows\desktop\tr", adPersistXML)

MsgBox "Hoho"
End Sub

' This function is used to test the html functions

Public Sub WebHtmlthatworked()

Dim Ele
Dim EleCell
Dim EleTable
Dim EleTbody

With m_frmWeb.web.document

    .Clear
           
    Set Ele = .createElement("Html")
    .body.appendChild Ele
    
    Set Ele = .createElement("Body")
    .body.appendChild Ele
    
    Set Ele = .createElement("H1")
    .body.appendChild Ele
    
    Ele.innerText = "This is test"
    
    Set Ele = .createElement("H1")
    .body.appendChild Ele
    
    Ele.innerText = "This is test 1"
    
    Set EleTable = .createElement("Table")

    EleTable.id = "PLTable"
    EleTable.border = 1
    EleTable.cellSpacing = 0
    EleTable.cellPadding = 0

    Set EleTbody = .createElement("TBody")
    EleTable.appendChild EleTbody

    Dim i As Integer
    Dim j As Integer

    For i = 0 To 4
        Set Ele = .createElement("TR")
        For j = 0 To 4

            Set EleCell = .createElement("TD")
            EleCell.innerText = "Test" & j
            Ele.appendChild (EleCell)
        Next j
        EleTbody.appendChild (Ele)
    Next i

    .body.appendChild EleTable
    
End With

    
End Sub





Private Sub Class_Initialize()

Set m_frmWeb = New frmWeb

Load m_frmWeb

End Sub

Private Sub Class_Terminate()

On Error Resume Next

Set m_frmWeb = Nothing

End Sub




Private Sub m_frmAccReports_OKClick(StDate As String, EndDate As String, ParentID As Long, HeadID As Long, ReportSelected As Wis_AccountReportList)

FromRpDate = StDate
ToRpDate = EndDate
RpParentID = ParentID
RpHeadID = HeadID
m_AccountReportList = ReportSelected

End Sub


Private Sub m_frmRptDt_OKClick(StDate As String, EndDate As String)
    FromRpDate = StDate
    ToRpDate = EndDate
End Sub
Public Property Get FromRpDate() As String
    FromRpDate = m_FromRpDate
End Property
Public Property Let FromRpDate(ByVal NewValue As String)
    m_FromRpDate = ""
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_FromRpDate = NewValue
End Property
Public Property Get ToRpDate() As String
    ToRpDate = m_ToRpDate
End Property
Public Property Get RpHeadID() As Long
    RpHeadID = m_RpHeadId
End Property
Public Property Let RpHeadID(ByVal NewValue As Long)
    m_RpHeadId = NewValue
End Property
Public Property Get RpParentID() As Long
    RpParentID = m_RpParentID
End Property
Public Property Let RpParentID(ByVal NewValue As Long)
    m_RpParentID = NewValue
End Property
Public Property Get FromRpAmount() As Currency
    FromRpAmount = m_FromRpAmount
End Property
Public Property Let FromRpAmount(ByVal NewValue As Currency)
    m_FromRpAmount = 0
    If Not CurrencyValidate(Str(NewValue), True) Then Exit Property
    m_FromRpAmount = NewValue
End Property
Public Property Get ToRpAmount() As Currency
    ToRpAmount = m_ToRpAmount
End Property
Public Property Let ToRpAmount(ByVal NewValue As Currency)
    m_ToRpAmount = 0
    If Not CurrencyValidate(Str(NewValue), True) Then Exit Property
    m_ToRpAmount = NewValue
End Property

Public Property Let ToRpDate(ByVal NewValue As String)
    m_ToRpDate = ""
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_ToRpDate = NewValue
End Property




