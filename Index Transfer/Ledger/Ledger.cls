VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ===================================
' This is PDL used throughout the Class
' ===================================
' Handle Error
' On error Goto Hell:
' Declarations
' Validations
' Exit
' Hell:

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmLedger As frmLedger
Attribute m_frmLedger.VB_VarHelpID = -1
'Private WithEvents m_frmLedger As frmLedger1

Private m_HeadID As Long
Private m_OpBalance As Currency
Private m_LedgerName As String
Private m_ParentID As Long
Private m_DBOperation As wis_DBOperation

' To Handle Grid Functions
Private m_GrdFunctions As clsGrdFunctions
'This function Will fetch the Particulars of the TransID Given
Public Function GetTransIDParticulars(TransID As Long) As String

' Handle Error
On Error GoTo Hell:

GetTransIDParticulars = ""

' Declarations
Dim rstParticulars As ADODB.Recordset

If TransID = 0 Then Exit Function

gDbTrans.SQLStmt = " SELECT Particulars " & _
                   " FROM TransParticulars " & _
                   " WHERE TransID=" & TransID
                   
If gDbTrans.Fetch(rstParticulars, adOpenForwardOnly) < 0 Then Exit Function
                
GetTransIDParticulars = FormatField(rstParticulars.Fields("Particulars"))

Set rstParticulars = Nothing

Exit Function

Hell:

    MsgBox "GetTransIDParticulars : " & vbCrLf & Err.Description
    Set rstParticulars = Nothing
    
End Function
'This will return the VoucherName of the given VoucherType
Public Function GetVoucherName(VoucherType As Wis_VoucherTypes) As String

GetVoucherName = ""

Select Case VoucherType
    Case Sales
        GetVoucherName = "Sales"
    Case Payment
        GetVoucherName = "Payment"
    Case Purchase
        GetVoucherName = "Purchase"
    Case Receipt
        GetVoucherName = "Receipt"
    Case Contra
        GetVoucherName = "Contra"
    Case FreePurchase
        GetVoucherName = "Free Purchase"
    Case FreeSales
        GetVoucherName = "Free Sales"
    Case Journal
        GetVoucherName = "Journal"
    Case RejectionsIn
        GetVoucherName = "RejectionsIn"
    Case RejectionsOut
        GetVoucherName = "RejectionsOut"
End Select

End Function

'This will load all Heads to the given Recordset from the Base Table

Public Sub LoadHeadstoRecordset(rstHeads As ADODB.Recordset)

' Handle Error
On Error GoTo Hell:

gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Heads"
                   
Call gDbTrans.Fetch(rstHeads, adOpenDynamic)

Exit Sub

Hell:
    
    
End Sub

Private Sub SetHeaderFooter(ByVal Balance As Currency, grd As MSFlexGrid, AccountType As wis_AccountType, MsgString As String)

With grd

'    m_GrdFunctions.fMoreRows (2)
    
    '.Row = .Row + 2
    .Col = 1: .Text = MsgString: .CellFontBold = True
               
      Select Case AccountType
        Case Asset, Loss, ItemPurchase
        
            If Balance >= 0 Then .Col = 3: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 4: .CellForeColor = vbRed
        
        Case Liability, Profit, ItemSales
        
            If Balance >= 0 Then .Col = 4: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 3: .CellForeColor = vbRed

        Case Loss
        
            If Balance >= 0 Then .Col = 3: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 4: .CellForeColor = vbRed

            
        Case Profit
        
            If Balance >= 0 Then .Col = 4: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 3: .CellForeColor = vbRed

            
        Case ItemSales
        
            If Balance >= 0 Then .Col = 4: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 3: .CellForeColor = vbRed
            
        Case ItemPurchase
        
            If Balance >= 0 Then .Col = 3: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 4: .CellForeColor = vbRed

    End Select
               
    .Text = FormatCurrency(Abs(Balance)): .CellFontBold = True
    
End With

End Sub
Public Function IsHeadIDBalanceDefault(ByVal Balance As Currency, HeadID As Long) As Boolean

On Error GoTo Hell:

' Declarations
Dim AccountType As wis_AccountType
Dim AccTransClass As clsAccTrans

' get the AccountTYpe

Set AccTransClass = New clsAccTrans

AccountType = AccountType = AccTransClass.GetAccountType(HeadID)

Set AccTransClass = Nothing

' Here default is True
IsHeadIDBalanceDefault = True

' Now set the Variable

Select Case AccountType

    Case Asset
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False
    
    Case Liability
        
        If Balance < 0 Then IsHeadIDBalanceDefault = False

    Case Loss
        
        If Balance < 0 Then IsHeadIDBalanceDefault = False
        
    Case Profit
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False

    Case ItemSales
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False
        
    Case ItemPurchase
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False

End Select
               
Exit Function

Hell:
        
End Function




Private Function AddRecords() As wis_FunctionReturned

AddRecords = Failure

If Validated <> Success Then Exit Function

If m_DBOperation = Insert Then SaveLedger
If m_DBOperation = Update Then UpdateLedger
  

AddRecords = Success

End Function

Public Sub ShowLedger()

' Check the Form's Status
If m_frmLedger Is Nothing Then Set m_frmLedger = New frmLedger

m_DBOperation = Insert

Load m_frmLedger
If ParentID Then
    With m_frmLedger
        Dim Count As Integer
        For Count = 0 To .cmbParent.ListCount - 1
            If .cmbParent.ItemData(Count) = m_ParentID Then
                .cmbParent.ListIndex = Count
                Exit For
            End If
        Next
    End With
End If
m_frmLedger.Show vbModal


End Sub

Private Function SaveLedger() As wis_FunctionReturned

On Error GoTo NoSaveError:

Dim rstHeads As ADODB.Recordset

Dim curOpBalance As Currency
Dim HeadID As Long
Dim OpAmericanDate As String

SaveLedger = Failure

' check the form's status
If m_frmLedger Is Nothing Then Exit Function

With m_frmLedger

    If .cmbParent.ListIndex = -1 Then Exit Function
    
    OpAmericanDate = FormatDate(FinIndianFromDate)
    curOpBalance = CCur(Val(.txtOpBalance.Text))
    SaveLedger = False
    
    ParentID = .cmbParent.ItemData(.cmbParent.ListIndex)
    
    'Get the Maximum Head From the database
    gDbTrans.SQLStmt = " SELECT MAX(HeadID) FROM Heads " & _
                       " WHERE ParentID = " & ParentID
    
    Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)
    
    HeadID = FormatField(rstHeads.Fields(0)) + 1
    
    If HeadID < ParentID Then HeadID = HeadID + ParentID
    
    'Insert the heads inot the database
    gDbTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID) " & _
                       " VALUES ( " & _
                       HeadID & "," & _
                       AddQuotes(.txtLedgerName.Text, True) & "," & _
                       ParentID & ")"
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    'Insert the opening balance
    gDbTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     HeadID & "," & _
                     "#" & OpAmericanDate & "#," & _
                     curOpBalance & ")"
                     
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.CommitTrans
    
End With

SaveLedger = Success

' Clear the Controls

ClearControls
Set rstHeads = Nothing

Exit Function

NoSaveError:
        
    ' Clear up the  Transactions if any
    Set rstHeads = Nothing
    SaveLedger = FatalError

End Function


' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it will create the
' Head with given HeadID & parentID
' Inputs :
'           HeadID as long
'           ParentID as long
' OutPut :
Public Function GetHeadIDCreated(ByVal ParentID As Long, ByVal HeadName As String, ByVal OpBalance As Currency) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset
Dim USOpDate As String
Dim HeadID As Long

'initialise the function
GetHeadIDCreated = -1

'Validate the inputs
If ParentID = 0 Then Exit Function
If HeadName = "" Then Exit Function

'Check the given Heads in the database
gDbTrans.SQLStmt = " SELECT HeadID" & _
                   " FROM Heads " & _
                   " WHERE HeadName = " & AddQuotes(HeadName) & _
                   " AND ParentID = " & ParentID
                
'if exists then exit function
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then
    GetHeadIDCreated = FormatField(rstHeads.Fields(0))
    Exit Function
End If

'Get the Maximum Head From the database
gDbTrans.SQLStmt = " SELECT MAX(HeadID)" & _
                   " FROM Heads " & _
                   " WHERE ParentID = " & ParentID

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

HeadID = FormatField(rstHeads.Fields(0)) + 1
    
If HeadID < ParentID Then HeadID = HeadID + ParentID

' Else create the head with given headID & parentiD
'Insert the heads into the database
gDbTrans.SQLStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID) " & _
                  " VALUES ( " & _
                  HeadID & "," & _
                  AddQuotes(HeadName) & "," & _
                  ParentID & ")"

If Not gDbTrans.SQLExecute Then Exit Function


USOpDate = FormatDate(FinIndianFromDate)
'Insert the opening balance
gDbTrans.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 "#" & USOpDate & "#," & _
                 OpBalance & ")"
If Not gDbTrans.SQLExecute Then Exit Function

GetHeadIDCreated = HeadID

Set rstHeads = Nothing

Exit Function

ErrLine:
    MsgBox "CreateHead: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    GetHeadIDCreated = -1
    
    Set rstHeads = Nothing
    
    Exit Function
        
    
End Function


Private Function UpdateLedger() As wis_FunctionReturned

On Error GoTo NoUpdateError:

Dim curOpBalance As Currency
Dim LedgerName As String

Dim OpAmericanDate As String

UpdateLedger = Failure

' check the form's status

If m_frmLedger Is Nothing Then Exit Function

With m_frmLedger

    If .cmbParent.ListIndex = -1 Then Exit Function
    
    OpAmericanDate = FormatDate(FinIndianFromDate)
    curOpBalance = CCur(Val(.txtOpBalance.Text))
    LedgerName = Trim$(.txtLedgerName.Text)
    ParentID = .cmbParent.ItemData(.cmbParent.ListIndex)
    
    gDbTrans.SQLStmt = " UPDATE Heads SET " & _
                       " HeadName = " & AddQuotes(LedgerName, True) & _
                       " WHERE HeadID = " & m_HeadID
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    'Insert the opening balance
    gDbTrans.SQLStmt = " UPDATE OpBalance SET " & _
                       " OpAmount = " & curOpBalance & _
                       " WHERE OpDate = #" & OpAmericanDate & "#" & _
                       " AND HeadID = " & m_HeadID
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.CommitTrans
    
End With

UpdateLedger = Success

'Clear the controls
ClearControls

Exit Function

NoUpdateError:

    UpdateLedger = FatalError

End Function

Private Sub ClearControls()

' check the form's status
If m_frmLedger Is Nothing Then Exit Sub

With m_frmLedger

    '.cmbParent.ListIndex = -1
    .txtLedgerName.Text = ""
    .txtOpBalance.Text = ""
    '.lvwLedger.ColumnHeaders.Clear
    .cmdOk.Caption = LoadResString(gLangOffSet + 10)
    .cmbParent.Locked = False
    Call LoadHeadsToListView(ParentID, FinIndianFromDate)
    ActivateTextBox .txtLedgerName
End With

m_DBOperation = Insert

End Sub

Private Sub LoadHeadsToListView(ByVal ParentID As Long, ByVal AsOnDate As String)

Dim rstHeads As ADODB.Recordset

' check the form's status
If m_frmLedger Is Nothing Then Exit Sub

If ParentID = 0 Then Exit Sub
If Not DateValidate(AsOnDate, "/", True) Then Exit Sub

m_frmLedger.lvwLedger.ListItems.Clear

gDbTrans.SQLStmt = " SELECT a.HeadID,HeadName,OpAmount " & _
                   " FROM Heads a,OpBalance b " & _
                   " WHERE a.ParentID =  " & ParentID & _
                   " AND a.HeadID=b.HeadID" & _
                   " AND b.OpDate=" & "#" & FormatDate(AsOnDate) & "#"
                   
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) < 0 Then Exit Sub

If Not FillViewNew(m_frmLedger.lvwLedger, rstHeads, "HeadID", True) Then Exit Sub

Set rstHeads = Nothing

End Sub

Private Function Validated() As wis_FunctionReturned

On Error GoTo ValidError:

Dim Rst As ADODB.Recordset

Validated = Failure

' Check the Form's Status
If m_frmLedger Is Nothing Then Exit Function

With m_frmLedger

    If .cmbParent.ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select Parent Name "
    
    If Not TextBoxCurrencyValidate(.txtOpBalance, True, True) Then Err.Raise vbObjectError + 513, , "Invalid Opening Balance Specified"
    
    Validated = Success
    
    If m_DBOperation = Update Then Exit Function
    
    Validated = Failure
    
    'Check for the existence of the head
    gDbTrans.SQLStmt = " SELECT HeadName FROM Heads " & _
                       " WHERE HeadName = " & AddQuotes(.txtLedgerName.Text, True) & _
                       " AND ParentID = " & .cmbParent.ItemData(.cmbParent.ListIndex)
    
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then Err.Raise vbObjectError + 513, , "Duplicate Entry "
        
    Set Rst = Nothing

End With

Validated = Success
Set Rst = Nothing

Exit Function

ValidError:
    
    MsgBox "Ledger Creation :" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Validated = FatalError
    Set Rst = Nothing
   
End Function

Private Sub Class_Initialize()

Set m_GrdFunctions = New clsGrdFunctions

End Sub

Private Sub Class_Terminate()

On Error Resume Next

Set m_frmLedger = Nothing
Set m_GrdFunctions = Nothing

End Sub


Private Sub m_frmLedger_CancelClick()

On Error Resume Next

Set m_frmLedger = Nothing
    
End Sub

Private Sub m_frmLedger_LookupClick(ParentID As Long)

Call LoadHeadsToListView(ParentID, FinIndianFromDate)
    
End Sub
Private Sub m_frmLedger_LvwLedgerClick(HeadID As Long)

m_HeadID = HeadID

With m_frmLedger

    LedgerName = .txtLedgerName.Text
    OpBalance = Val(.txtOpBalance.Text)
    
    m_DBOperation = Update
End With

End Sub

Private Sub m_frmLedger_OKClick()

Call AddRecords

End Sub



Private Sub m_frmLookUp_SelectClick(strSelection As String)

m_HeadID = CLng(strSelection)

m_DBOperation = Update

End Sub
Public Property Get OpBalance() As Currency
 OpBalance = m_OpBalance
End Property

Public Property Let OpBalance(ByVal NewValue As Currency)
    m_OpBalance = NewValue
End Property

Public Property Get LedgerName() As String
    LedgerName = m_LedgerName
End Property

Public Property Let LedgerName(ByVal NewValue As String)
    m_LedgerName = NewValue
End Property

Public Property Get ParentID() As Long
    ParentID = m_ParentID
End Property

Public Property Let ParentID(ByVal NewValue As Long)
    m_ParentID = NewValue
End Property
