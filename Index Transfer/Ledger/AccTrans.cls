VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CrHeads"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"DrHeads"
Option Explicit



'Private m_frmAccTrans As frmAccTrans
Private m_AllTrans As AllTrans
'Private m_GrdFunctions As clsGrdFunctions

' Local Variable(s) to Hold Property Value(s)
Private m_SlNo As Integer

Private m_HeadName As String
Private m_TransDate As String

Private m_Debit As Currency
Private m_Credit As Currency
Private m_Balance As Currency
Private m_HeadID As Long
Private m_PrevHeadID As Long
Private m_TransID As Long

Private m_isUpdatingGrid As Boolean
Private m_UpdatingTrans As Boolean

Public Enum wis_LoadLedgers
    LoadALL = 0
    LoadExceptCash = 1
    LoadOnlyCash = 2
    LoadCashCreditros = 4
End Enum

' For the Combos
Private Const VoucherCombo = 0
Private Const ParentCombo = 1
Private Const LedgerCombo = 2
'Sda 17/12/02
Private Const TransGrid = 0
Private Const LedgerGrid = 1

' For the Options
Private Const FromOpt = 0
Private Const ToOpt = 1

' For the Add Commands
Private Const AddToGrid = 0
Private Const DeleteGrid = 1
'
Private Sub CheckTransId()

Dim rstTrans As ADODB.Recordset

NewIndexTrans.SQLStmt = " SELECT Sum(Debit),Sum(Credit)" & _
                   " FROM AccTrans " & _
                   " WHERE TransID = " & TransID
                 
Call NewIndexTrans.Fetch(rstTrans, adOpenForwardOnly)

If FormatField(rstTrans(0)) <> FormatField(rstTrans(1)) Then _
    MsgBox "CheckTransID :" & TransID & vbCrLf & "Total Not Same !!!"
    
End Sub

' This will Convert Existing  ParentHeads data to New data which is in the Parents Table
Private Function ConvertParentIDtoNewParentID() As Boolean

On Error GoTo Hell:

ConvertParentIDtoNewParentID = False

Dim rstParents As ADODB.Recordset

Dim Parent As Long
Dim ParentID As Long
Dim NewParentID As Long

NewIndexTrans.SQLStmt = " SELECT ParentID,Parent,NewParentID " & _
                   " FROM Parents"

Call NewIndexTrans.Fetch(rstParents, adOpenForwardOnly)

With rstParents

    Do While Not .EOF
        
        Parent = .Fields("Parent")
        ParentID = .Fields("ParentID")
        NewParentID = .Fields("NewParentID")
        
        NewIndexTrans.SQLStmt = " UPDATE ParentHeads " & _
                         " SET ParentID=" & NewParentID & "," & _
                         " Parent = " & Parent & _
                         " WHERE ParentID = " & ParentID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
                
        NewIndexTrans.SQLStmt = "UPDATE Heads " & _
                         " SET ParentID = " & NewParentID & _
                         " WHERE ParentID = " & ParentID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
                        
        .MoveNext
    Loop
    
End With

ConvertParentIDtoNewParentID = True

Exit Function

Hell:
    ConvertParentIDtoNewParentID = False
    
End Function
' This will Convert Existing  HeadsID data to New Heads ID
Private Function ConvertHeadsIDtoNewHeadsID() As Boolean

On Error GoTo Hell:

ConvertHeadsIDtoNewHeadsID = False

Dim rstHeads As ADODB.Recordset

Dim HeadID As Long
Dim NewHeadID As Long

NewIndexTrans.SQLStmt = "SELECT HeadID" & _
                  " FROM Heads" & _
                  " ORDER BY ParentID"

Call NewIndexTrans.Fetch(rstHeads, adOpenForwardOnly)

NewHeadID = 5

With rstHeads

    Do While Not .EOF
        
        HeadID = .Fields("HeadID")
        NewHeadID = NewHeadID + 1
        
        NewIndexTrans.SQLStmt = "UPDATE Heads" & _
                           " SET HeadID = " & NewHeadID & _
                           " WHERE HeadID = " & HeadID
        If Not NewIndexTrans.SQLExecute Then Exit Function
                
        NewIndexTrans.SQLStmt = "UPDATE AccTrans " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
        
        
        NewIndexTrans.SQLStmt = "UPDATE CompanyCreation" & _
                         " SET HeadID = " & NewHeadID & _
                         " WHERE HeadID = " & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
        
        
        NewIndexTrans.SQLStmt = " UPDATE InvoiceDetails" & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
               
        
        NewIndexTrans.SQLStmt = " UPDATE OpBalance " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack
            
        NewIndexTrans.SQLStmt = " UPDATE Purchase " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
        
                
        NewIndexTrans.SQLStmt = " UPDATE RelationMaster " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
                        
        NewIndexTrans.SQLStmt = " UPDATE Sales " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & HeadID
                         
        If Not NewIndexTrans.SQLExecute Then Exit Function
                
        .MoveNext
    Loop
    
End With

ConvertHeadsIDtoNewHeadsID = True

Exit Function

Hell:
    ConvertHeadsIDtoNewHeadsID = False
    
End Function
'
Public Function GetVoucherNametoArray()

Dim VoucherName(0 To 10) As String
Dim Vouchers As Wis_VoucherTypes
Dim Item As Integer

'GetVoucherNametoArray = VoucherName
'ReDim GetVoucherNametoArray(0 To 10)

Vouchers = Contra
VoucherName(Vouchers) = "Contra"
Vouchers = FreePurchase
VoucherName(Vouchers) = "Free Purchase"
Vouchers = FreeSales
VoucherName(Vouchers) = "Free Sales"
Vouchers = Journal
VoucherName(Vouchers) = "Journal"
Vouchers = Payment
VoucherName(Vouchers) = "Payment"
Vouchers = Purchase
VoucherName(Vouchers) = "Purchase"
Vouchers = Receipt
VoucherName(Vouchers) = "Receipt"
Vouchers = RejectionsIn
VoucherName(Vouchers) = "Rejections In"
Vouchers = RejectionsOut
VoucherName(Vouchers) = "Rejections Out"
Vouchers = Sales
VoucherName(Vouchers) = "Sales"
Vouchers = VouNothing
VoucherName(Vouchers) = "Nothing"

GetVoucherNametoArray = VoucherName(Vouchers)


End Function

' Here starts the debit entries
' Start the Loop and start inserting the Variables
Private Function AllTransCRInsertToTable(ByVal VoucherType As Wis_VoucherTypes, _
                                      ByVal USTransDate As String _
                                      ) As wis_FunctionReturned

On Error GoTo Hell:

Dim ObjKeys As Variant
    
Dim MaxCount As Integer
Dim Item As Integer

AllTransCRInsertToTable = Failure

MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .CrAmount > 0 Then
            ' Insert the Variables
            NewIndexTrans.SQLStmt = " INSERT INTO AccTrans " & _
                     " (HeadID,TransID,VoucherType,TransDate,Debit,Credit)" & _
                     " VALUES (" & _
                     .HeadID & "," & _
                     TransID & "," & _
                     VoucherType & "," & _
                     "#" & USTransDate & "#," & _
                      .DrAmount & "," & _
                      .CrAmount & _
                       ")"
            If Not NewIndexTrans.SQLExecute Then Exit Function
        End If
    End With
Next Item

AllTransCRInsertToTable = Success

Exit Function

Hell:
     
    AllTransCRInsertToTable = FatalError
    
End Function

' Here starts the debit entries
' Start the Loop and start inserting the Variables
Private Function AllTransDRInsertToTable(ByVal VoucherType As Wis_VoucherTypes, _
                                      ByVal USTransDate As String _
                                      ) As wis_FunctionReturned

On Error GoTo Hell:

Dim ObjKeys As Variant
    
Dim MaxCount As Integer
Dim Item As Integer

AllTransDRInsertToTable = Failure

MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .DrAmount > 0 Then
            ' Insert the Variables
            NewIndexTrans.SQLStmt = "INSERT INTO AccTrans " & _
                     " (HeadID,TransID,VoucherType,TransDate,Debit,Credit)" & _
                     " VALUES (" & _
                     .HeadID & "," & _
                     TransID & "," & _
                     VoucherType & "," & _
                     "#" & USTransDate & "#," & _
                      .DrAmount & "," & _
                      .CrAmount & _
                      ")"
            If Not NewIndexTrans.SQLExecute Then Exit Function
        End If
    End With
Next Item


AllTransDRInsertToTable = Success

Exit Function

Hell:
     
    AllTransDRInsertToTable = FatalError
    
End Function
Private Function IsTransIDFromInventory(ByVal TransID As Long) As Boolean

On Error GoTo Hell:

Dim rstID As ADODB.Recordset

IsTransIDFromInventory = False

NewIndexTrans.SQLStmt = " SELECT * " & _
                   " FROM IDFromInventory " & _
                   " WHERE TransID=" & TransID

If NewIndexTrans.Fetch(rstID, adOpenForwardOnly) < 1 Then Exit Function

IsTransIDFromInventory = True

Set rstID = Nothing

Exit Function

Hell:

End Function

' This sub is check all transid and its treatment,along with vouchers
' This sub checks the transactions with current validations
Public Sub MyExternalValidations()

Dim rstTrans As ADODB.Recordset
Dim VoucherType As Wis_VoucherTypes

NewIndexTrans.SQLStmt = " SELECT DISTINCT(TransID),VoucherType" & _
                    " FROM AccTrans"
       
Call NewIndexTrans.Fetch(rstTrans, adOpenForwardOnly)

Do While Not rstTrans.EOF
    
    Call MyLoadtoCollection(rstTrans.Fields("TransID"))
    
    VoucherType = rstTrans.Fields("VoucherType")
    
    If AllTransVoucherCheck(VoucherType) <> Success Then Debug.Print rstTrans.Fields("TransID")
    
    ClearCollections
    
    rstTrans.MoveNext
    
Loop


End Sub

Private Sub MyLoadtoCollection(ByVal lngTransID As Long)

Dim rstTrans As ADODB.Recordset

If lngTransID = 0 Then Exit Sub '

' Get the Records to Recordset
NewIndexTrans.SQLStmt = " SELECT * " & _
                   " FROM AccTrans " & _
                   " WHERE TransID=" & lngTransID
                   
If NewIndexTrans.Fetch(rstTrans, adOpenDynamic) < 0 Then Exit Sub

With rstTrans

    ' Start the Loop
    Do While Not .EOF
    
        'Feed data to Controls
        HeadID = FormatField(.Fields("HeadID"))
        Debit = FormatField(.Fields("Debit"))
        Credit = FormatField(.Fields("Credit"))
        TransDate = FormatField(.Fields("TransDate"))
        
        Balance = GetBalance(HeadID, .Fields(TransDate), .Fields(TransID))
        
        'Update the Collections
        If UpdateCollections <> Success Then Exit Sub
        
        .MoveNext
        
    Loop
End With

End Sub


Public Function NewDbConvertions() As Boolean

NewDbConvertions = False

If Not ConvertParentIDtoNewParentID Then Exit Function
If Not ConvertHeadsIDtoNewHeadsID Then Exit Function

NewDbConvertions = True

End Function
'This will get the Transdate from collection
' and will check the amount
' if exists it will add the amount
Private Function TransDateDataExists(ByVal USTransDate As Date) As Long
' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim rstTrans As ADODB.Recordset
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim TransID As Long

Dim DrTotal As Currency
Dim CrTotal As Currency

Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

' Intialise the Function
TransDateDataExists = 0

If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total

MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .DrAmount > 0 Then ToHeadID = .HeadID
        If .CrAmount > 0 Then FromHeadID = .HeadID
    End With
    
Next Item

NewIndexTrans.SQLStmt = " SELECT * " & _
                   " FROM QryAccTransMerge" & _
                   " WHERE TransDate = #" & USTransDate & "#" & _
                   " AND A.HEADID=" & ToHeadID & _
                   " AND B.HEADID=" & FromHeadID
                   
If NewIndexTrans.Fetch(rstTrans, adOpenStatic) < 1 Then Exit Function

DrTotal = rstTrans.Fields("a.Debit").Value
CrTotal = rstTrans.Fields("b.Credit").Value

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .DrAmount > 0 Then
            .DrAmount = .DrAmount + DrTotal
            DrTotal = .DrAmount
        End If
        
        If .CrAmount > 0 Then
            .CrAmount = .CrAmount + CrTotal
            CrTotal = .CrAmount
        End If
        
    End With
    
Next Item

TransID = rstTrans.Fields("TransID")

NewIndexTrans.BeginTrans

NewIndexTrans.SQLStmt = " UPDATE AccTrans " & _
                  " SET Debit= " & DrTotal & _
                  " WHERE HeadID=" & ToHeadID & _
                  " AND TransID=" & TransID
                  
If Not NewIndexTrans.SQLExecute Then
    NewIndexTrans.RollBack
    Exit Function
End If

NewIndexTrans.SQLStmt = " UPDATE AccTrans " & _
                  " SET Credit= " & CrTotal & _
                  " WHERE HeadID=" & FromHeadID & _
                  " AND TransID=" & TransID
                  
If Not NewIndexTrans.SQLExecute Then
    NewIndexTrans.RollBack
    Exit Function
End If

NewIndexTrans.CommitTrans

TransDateDataExists = TransID

Exit Function

Hell:
    TransDateDataExists = 0
    
End Function
'This will get the Transdate from collection
' and will check the amount
' if exists it will add the amount
Private Function UndoTransDateData(ByVal USTransDate As String) As Long
' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim rstTrans As ADODB.Recordset
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim TransID As Long

Dim DrTotal As Currency
Dim CrTotal As Currency

Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

' Intialise the Function
UndoTransDateData = 0

If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total

MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        
        If .DrAmount > 0 Then ToHeadID = .HeadID
        
        If .CrAmount > 0 Then FromHeadID = .HeadID
        
    End With
    
Next Item

NewIndexTrans.SQLStmt = " SELECT * " & _
                   " FROM QryAccTransMerge" & _
                   " WHERE TransDate = #" & USTransDate & "#" & _
                   " AND A.HEADID=" & ToHeadID & _
                   " AND B.HEADID=" & FromHeadID
                   
If NewIndexTrans.Fetch(rstTrans, adOpenStatic) < 0 Then Exit Function

DrTotal = rstTrans.Fields("a.Debit").Value
CrTotal = rstTrans.Fields("b.Credit").Value

For Item = 0 To MaxCount - 1
    
    With AllTrans(ObjKeys(Item))
        
        If .DrAmount > 0 Then
            .DrAmount = DrTotal - .DrAmount
            DrTotal = .DrAmount
        End If
        
        If .CrAmount > 0 Then
            .CrAmount = CrTotal - .CrAmount
            CrTotal = .CrAmount
        End If
        
    End With
    
Next Item

TransID = rstTrans.Fields("TransID")

NewIndexTrans.BeginTrans

If DrTotal > 0 And CrTotal > 0 Then
    NewIndexTrans.SQLStmt = " UPDATE AccTrans " & _
                      " SET Debit= " & DrTotal & _
                      " WHERE HeadID=" & ToHeadID & _
                      " AND TransID=" & TransID
                      
    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack: Exit Function
    
    NewIndexTrans.SQLStmt = " UPDATE AccTrans " & _
                      " SET Credit= " & CrTotal & _
                      " WHERE HeadID=" & FromHeadID & _
                      " AND TransID=" & TransID
                      
    If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack: Exit Function
Else
    ' Becasue we have Reached o we are simply deleting
    m_TransID = TransID
    NewIndexTrans.BeginTrans
    If DeleteTransID = Success Then NewIndexTrans.CommitTrans Else NewIndexTrans.RollBack
    
End If

NewIndexTrans.CommitTrans

UndoTransDateData = TransID

Exit Function

Hell:
    
    UndoTransDateData = 0
    
End Function

Private Property Get UpdatingTrans() As Boolean
UpdatingTrans = m_UpdatingTrans
End Property

' This Will Check the Vouchers ans recpective entries
Private Function ValidateVouchers(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle error
On Error GoTo Hell:

ValidateVouchers = Failure

' This Will Fetch The Current Voucher Type Selected by the User

If DrVoucherCheck(VoucherType) <> Success Then Exit Function
If CrVoucherCheck(VoucherType) <> Success Then Exit Function

ValidateVouchers = Success

Exit Function

Hell:
    ValidateVouchers = FatalError
    
End Function

Private Sub LoadParentHeadsForVoucher(ctrlComboBox As ComboBox, ByVal LoadLedgers As wis_LoadLedgers)

On Error GoTo NoLoadParents:

Dim rstParent As ADODB.Recordset
Dim SqlStr As String

ctrlComboBox.Clear

SqlStr = " SELECT ParentName,ParentID " & _
        " FROM ParentHeads WHERE UserCreated <= 2 "

Select Case LoadLedgers

    Case LoadALL
    
        NewIndexTrans.SQLStmt = SqlStr & _
                           " ORDER BY ParentName "
                    
    Case LoadExceptCash
    
        'NewIndexTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                           " FROM ParentHeads " & _
                           " WHERE ParentID <> " & wis_BanksParentID & _
                           " AND ParentID <> " & wis_BanksODParentID & _
                           " AND ParentID <> " & wis_CashParentID & _
                           " ORDER BY ParentName "
      '''Cahnged For INdex 2000 NEw
      
       NewIndexTrans.SQLStmt = SqlStr & " AND ParentID <> " & wis_CashParentID & _
                           " ORDER BY ParentName "
      
    Case LoadOnlyCash
    
         NewIndexTrans.SQLStmt = SqlStr & " AND ParentID = " & wis_BanksParentID & _
                            " OR ParentID = " & wis_BanksODParentID & _
                            " OR ParentID = " & wis_CashParentID & _
                            " ORDER BY ParentName "
                
    Case LoadCashCreditros
    
         NewIndexTrans.SQLStmt = SqlStr & " AND ParentID = " & wis_BanksParentID & _
                            " OR ParentID = " & wis_BanksODParentID & _
                            " OR ParentID = " & wis_CashParentID & _
                            " OR ParentID=" & wis_CreditorsParentID & _
                            " OR ParentID=" & wis_DebitorsParentID & _
                            " ORDER BY ParentName "
        
End Select

SqlStr = ""

Call NewIndexTrans.Fetch(rstParent, adOpenForwardOnly)

Do While Not rstParent.EOF
    With ctrlComboBox
        .AddItem FormatField(rstParent.Fields("ParentName"))
        .ItemData(.NewIndex) = FormatField(rstParent.Fields("ParentID"))
    End With
    'Move to the next record
    rstParent.MoveNext
Loop

Set rstParent = Nothing

Exit Sub

NoLoadParents:
    ctrlComboBox.Clear
    Set rstParent = Nothing
End Sub

' This will read the Property TransID and Delete the Records
' Dependencies : TransID
Private Function DeleteTransID() As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Initialise Function
DeleteTransID = Failure

' Should Never Happen
If TransID = 0 Then Exit Function

' Delete from AccTrans
NewIndexTrans.SQLStmt = " DELETE * " & _
                   " FROM AccTrans " & _
                   " WHERE TransID=" & TransID
'NewIndexTrans.BeginTrans
If Not NewIndexTrans.SQLExecute Then NewIndexTrans.RollBack

NewIndexTrans.SQLStmt = " DELETE * " & _
                   " FROM TransParticulars " & _
                   " WHERE TransID=" & TransID

If Not NewIndexTrans.SQLExecute Then Exit Function

DeleteTransID = Success

Exit Function

Hell:
    DeleteTransID = FatalError
    
End Function

'This function will check the Debit side total and Credit side total
Private Function DrCrTotalsSame(isHandleVouchers As Boolean) As Currency

' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim DrTotal As Currency
Dim CrTotal As Currency

Dim Item As Integer
Dim MaxCount As Integer

Dim ObjKeys As Variant

' Intialise the Function
DrCrTotalsSame = 0

'Check the Collection Dr Amount and Cr Amount Total
MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys
For Item = 0 To MaxCount - 1
    DrTotal = DrTotal + AllTrans(ObjKeys(Item)).DrAmount
    CrTotal = CrTotal + AllTrans(ObjKeys(Item)).CrAmount
Next Item

' If Equal Then make function's value to true

If DrTotal = CrTotal Then DrCrTotalsSame = 0
If DrTotal > CrTotal Then DrCrTotalsSame = DrTotal - CrTotal
If DrTotal < CrTotal Then DrCrTotalsSame = CrTotal - DrTotal

If Not isHandleVouchers Then Exit Function

' This will handle the Vocuhers and Option Buttons
'Call HandleVouchersAfterAdd(DrTotal, CrTotal)

Exit Function

Hell:

    MsgBox "DrCrTotalSame :" & vbCrLf & Err.Description
    
    
End Function

'This function will check the validation of the Voucher
Private Function AllTransVoucherCheck(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Intialise the Function
AllTransVoucherCheck = Failure

' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function

If DrVoucherCheck(VoucherType) <> Success Then Exit Function
If CrVoucherCheck(VoucherType) <> Success Then Exit Function

AllTransVoucherCheck = Success

Exit Function

Hell:
    
    AllTransVoucherCheck = FatalError
    
    
End Function

Private Function DrVoucherCheck(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim HeadID As Long
Dim ParentID As Long
Dim Item As Long
Dim MaxCount As Long

Dim isCashFound As Boolean
Dim isPurchaseFound As Boolean

Dim ObjKeys As Variant

Dim objTrans As Trans

Const CURRENTSIDE = "Debit Side: " & vbCrLf

' Intialise the Function
DrVoucherCheck = Failure

' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total
isCashFound = False
isPurchaseFound = False

Set objTrans = New Trans
MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys
For Item = 0 To MaxCount - 1

    Set objTrans = AllTrans.Item(ObjKeys(Item))
    
    'Validation Goes Hear
    ' Please Don't Change the Case Creiteria
    If objTrans.DrAmount > 0 Then
    
        HeadID = objTrans.HeadID
        ParentID = objTrans.ParentID
    
        Select Case VoucherType
        
            Case Payment
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) _
                 Then _
                Err.Raise vbObjectError + 513, , CURRENTSIDE & "Payment Voucher Can not Contain Cash or Bank !!"
                
            Case Receipt
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then isCashFound = True
                
                            
            Case Sales
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Or _
                    (ParentID = wis_CreditorsParentID) Or (ParentID = wis_DebitorsParentID) Then
                                     
                     isCashFound = True
                End If
                
                
            Case Purchase
            
               If ParentID = wis_PurchaseParentID Then isPurchaseFound = True
                
            Case FreePurchase
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
                    
            Case FreeSales
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
                
            Case Contra
            
                If (ParentID <> wis_CashParentID) And (ParentID <> wis_BanksParentID) And (ParentID <> wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Contra Voucher Can not Contain other than Cash or Bank !!"
            
            Case Journal
        
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Journal Voucher Can not Contain Cash or Bank !!"
    
            Case RejectionsIn
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsIN Voucher Can not Contain Cash or Bank !!"
                
            Case RejectionsOut
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsOut Voucher Can not Contain Cash or Bank !!"
                
        End Select
        
    End If
    
Next Item

Set objTrans = Nothing

Select Case VoucherType

    Case Sales

        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Sales Voucher Must have Creditros or Cash or Bank"
                
    Case Purchase
        
        If Not isPurchaseFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Purchase Voucher Must have Purchase Account"
    
    Case Receipt
        
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Payment or Receipt Voucher Should contain Cash or Bank !!"
        
End Select

DrVoucherCheck = Success

Exit Function


Hell:
    
    If Err Then MsgBox "Voucher Check : " & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    DrVoucherCheck = FatalError
    'Resume
    
End Function

'This function will check the validation of the Voucher
Private Function CrVoucherCheck(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim HeadID As Long
Dim ParentID As Long
Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

Dim objTrans As Trans

Dim isCashFound As Boolean
Dim isSalesFound As Boolean

Const CURRENTSIDE = "Credit Side: " & vbCrLf

' Intialise the Function
CrVoucherCheck = Failure

' Validate the Inputs and Classes
'THe Below condition never occurs 'SDA 16/12/2002
If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total

isCashFound = False
isSalesFound = False

Set objTrans = New Trans
MaxCount = AllTrans.Count
ObjKeys = AllTrans.GetKeys
For Item = 0 To MaxCount - 1

    Set objTrans = AllTrans.Item(ObjKeys(Item))
    
    'Validation Goes Hear
    If objTrans.CrAmount > 0 Then
    
        HeadID = objTrans.HeadID
        ParentID = objTrans.ParentID
        
        ' Please Don't Change the Case Creiteria
        Select Case VoucherType
        
            Case Payment
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then isCashFound = True
                    
            Case Receipt
            
                'If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Receipt Voucher Can not Contain Cash or Bank !!"
                
            Case Sales
                
                If (ParentID = wis_SalesParentID) Then isSalesFound = True
                
            Case Purchase
                
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Or _
                    (ParentID = wis_CreditorsParentID) Or (ParentID = wis_DebitorsParentID) Then
                                     
                     isCashFound = True
                     
                 
                End If
                
            Case FreePurchase
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
                    
            Case FreeSales
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
            Case Contra
            
                If (ParentID <> wis_CashParentID) And (ParentID <> wis_BanksParentID) And (ParentID <> wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Contra Voucher Can not Contain Other than Cash or Bank !!"
                
            Case Journal
        
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Journal Voucher Can not Contain Cash or Bank !!"
                
            Case RejectionsIn
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsIN Voucher Can not Contain Cash or Bank !!"
                
            Case RejectionsOut
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsOut Voucher Can not Contain Cash or Bank !!"
                
        End Select
        
    End If
    
Next Item

Set objTrans = Nothing

Select Case VoucherType
    Case Payment
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Payment Voucher Should contain Cash or Bank !!"
    Case Purchase
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Purchase Voucher Should contain Cash or Bank or Debitor!!"
    Case Sales
        If Not isSalesFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Sales Voucher Should contain Sales Account !"
End Select

CrVoucherCheck = Success

Exit Function

Hell:
    If Err Then MsgBox "Voucher Check : " & Err.Description, vbCritical, wis_MESSAGE_TITLE
    CrVoucherCheck = FatalError
    Set objTrans = Nothing
    
End Function


Private Property Let Credit(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Credit = 5
    m_Credit = vData
End Property

Public Property Let UpdatingGrid(ByVal Data As Boolean)
    m_isUpdatingGrid = Data
End Property

Private Property Get Credit() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Credit
    Credit = m_Credit
End Property

Public Property Get UpdatingGrid() As Boolean
    UpdatingGrid = m_isUpdatingGrid
End Property

Private Property Let Debit(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Debit = 5
    m_Debit = vData
End Property

Private Property Get Debit() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Debit
    Debit = m_Debit
End Property
Private Property Let Balance(ByVal vData As Currency)
    m_Balance = vData
End Property
Private Property Get Balance() As Currency
    Balance = m_Balance
End Property
Private Property Let TransDate(ByVal vData As String)
    m_TransDate = vData
End Property
Private Property Get TransDate() As String
    TransDate = m_TransDate
End Property

Private Property Let HeadName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadName = 5
    m_HeadName = vData
End Property

Private Property Get HeadName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadName
    HeadName = m_HeadName
End Property
Private Property Let HeadID(ByVal vData As Long)
    m_HeadID = vData
End Property

Private Property Get HeadID() As Long
    HeadID = m_HeadID
End Property
Private Property Let PrevHeadID(ByVal vData As Long)
    m_PrevHeadID = vData
End Property
Private Property Get PrevHeadID() As Long
    PrevHeadID = m_PrevHeadID
End Property

Public Property Let TransID(ByVal vData As Long)
    m_TransID = vData
End Property
Public Property Get TransID() As Long
    TransID = m_TransID
End Property

Private Property Let SlNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlNo = 5
    m_SlNo = vData
End Property


Private Property Get SlNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlNo
    SlNo = m_SlNo
End Property

Private Sub ClearCollections()

On Error Resume Next


Set m_AllTrans = Nothing

'Shashi 21/12/2002
'Clear the Updation process if it is doing it
UpdatingGrid = False

End Sub
Private Function ALlTransHeadIDExists(ByVal HeadID As Long) As Boolean
   
On Error GoTo Hell:

    ALlTransHeadIDExists = False
    ALlTransHeadIDExists = AllTrans.Exists("Tr" & HeadID)
                
Exit Function

Hell:
    
End Function
Private Property Get AllTrans() As AllTrans
    If m_AllTrans Is Nothing Then Set m_AllTrans = New AllTrans
    Set AllTrans = m_AllTrans
End Property
Private Function AllTransGetKeys() As Variant

    If m_AllTrans Is Nothing Then Set m_AllTrans = New AllTrans
        
    Set AllTransGetKeys = m_AllTrans.GetKeys
    
    
End Function
Private Property Set AllTrans(ByVal vData As AllTrans)
    Set m_AllTrans = vData
End Property

' Returns the MaxTransID from the Base Table
' Inputs :
'           HeadID as long
' Retrurns TransID as long
Public Function GetMaxTransID(ByVal HeadID As Long) As Long
' Handle Error
On Error GoTo TransIDFailed:

' Declare Variable
Dim rstTransID As ADODB.Recordset

Dim TableName As String

' Intialise the Function
GetMaxTransID = 0

' Validate the inputs
If HeadID = 0 Then Exit Function

' Default Name
TableName = "AccTrans"

'If HeadID = wis_CashHeadID Then TableName = "CashTrans"

' Fetch the MaxTransID from Acctrans AND FROM Cash Trans
NewIndexTrans.SQLStmt = " Select Max(TransID) as AccTransID " & _
                   " FROM " & TableName & _
                   " WHERE Headid=" & HeadID
                   
If NewIndexTrans.Fetch(rstTransID, adOpenForwardOnly) < 0 Then Exit Function

' Set getMaxtransid
GetMaxTransID = FormatField(rstTransID.Fields("AccTransID"))

Set rstTransID = Nothing

Exit Function

TransIDFailed:
    
    Set rstTransID = Nothing
    
End Function


'
' Returns the MaxTransID from the Base Table
' Returns TransID as long
'
Public Function GetMaxTransIDFromTable() As Long
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim rstTransID As ADODB.Recordset
' Intialise the Function
GetMaxTransIDFromTable = 0
' Fetch the MaxTransID from Acctrans AND FROM Cash Trans
NewIndexTrans.SQLStmt = " Select Max(TransID) as AccTransID " & _
                   " FROM AccTrans"

If NewIndexTrans.Fetch(rstTransID, adOpenForwardOnly) < 0 Then Exit Function
' Set getMaxtransid
GetMaxTransIDFromTable = FormatField(rstTransID.Fields(0))
Set rstTransID = Nothing
Exit Function
TransIDFailed:
    Set rstTransID = Nothing
End Function

'Returns the Opening Balance from the Base Table
' Inputs :
'           HeadID as long
'           TransID as long
'
' Retrurns Balance
Public Function GetOpBalance(ByVal HeadID As Long, ByVal AsOnIndainDate As String) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim rstBalance As ADODB.Recordset
' Intialise the Function
GetOpBalance = 0
' Validate the inputs
If HeadID = 0 Then Exit Function
If Not DateValidate(AsOnIndainDate, "/", True) Then Exit Function
Dim AsOnDate As Date
AsOnDate = FormatDate(AsOnIndainDate)

' Fetch the Balance from Acctrans AND FROM Cash Trans
If CDate(AsOnDate) = CDate(FinUSFromDate) Then
    NewIndexTrans.SQLStmt = " SELECT OpAmount " & _
                       " FROM OpBalance " & _
                       " WHERE Headid=" & HeadID & _
                       " AND Opdate=" & "#" & AsOnDate & "#"
    
    If NewIndexTrans.Fetch(rstBalance, adOpenForwardOnly) < 0 Then Exit Function
    GetOpBalance = FormatField(rstBalance.Fields(0))
Else
    AsOnDate = DateAdd("d", -1, AsOnDate)
    GetOpBalance = GetBalance(HeadID, AsOnDate)
End If

Set rstBalance = Nothing
Exit Function

TransIDFailed:
        Set rstBalance = Nothing
End Function

' Returns the Balance from the Base Table
' Inputs :
'           HeadID as long
'
'
' Retrurns Balance
'
Private Function GetBalance(ByVal HeadID As Long, _
                           ByVal AsOnDate As Date, _
                           Optional TransID As Long) As Currency
                           
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim rstBalance As ADODB.Recordset
Dim TableName As String
Dim OpBalance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim AccountType As wis_AccountType
' Intialise the Function
GetBalance = 0
' Validate the inputs
If HeadID = 0 Then Exit Function

Dim FromDate As Date

FromDate = FinUSFromDate

' Get the Opening Balance
OpBalance = GetOpBalance(HeadID, FinIndianFromDate)

' Get the Sum of Debit and Sum of Credit as on date
NewIndexTrans.SQLStmt = " SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal" & _
                   " FROM AccTrans a" & _
                   " WHERE TransDate BETWEEN #" & FromDate & "#" & _
                   " AND #" & AsOnDate & "#" & _
                   " AND a.HeadID = " & HeadID
                   
If NewIndexTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then Exit Function
DebitTotal = FormatField(rstBalance.Fields("DebitTotal"))
CreditTotal = FormatField(rstBalance.Fields("CreditTotal"))
Set rstBalance = Nothing

If TransID Then
'If there are more transaction on the same day then the
'result of the balance will include the balances after theis transction
'so to get the balance after a particular tranction this is usefull
'Sda 7/1/03
    NewIndexTrans.SQLStmt = "SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal" & _
                   " FROM AccTrans a" & _
                   " WHERE TransDate = #" & AsOnDate & "#" & _
                   " AND a.HeadID = " & HeadID & _
                   " AND TransID >= " & TransID
    If NewIndexTrans.Fetch(rstBalance, adOpenForwardOnly) > 0 Then
        DebitTotal = DebitTotal - FormatField(rstBalance.Fields("DebitTotal"))
        CreditTotal = CreditTotal - FormatField(rstBalance.Fields("CreditTotal"))
        Set rstBalance = Nothing
    End If
End If

' Get AccountType
AccountType = GetAccountType(HeadID)
'Depending opon the AccountType calculate the Balance
GetBalance = GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)

Exit Function

TransIDFailed:
        Set rstBalance = Nothing
End Function
' Returns the Balance from the Base Table
' Inputs :
'           HeadID as long
'
'
' Retrurns Balance
'
Public Function GetBalanceFromRecordset(ByVal HeadID As Long, rstBalance As ADODB.Recordset, ByVal AccountType As wis_AccountType) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim TransID As Long
Dim OpBalance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
' Intialise the Function
GetBalanceFromRecordset = 0
' Validate the inputs
If HeadID = 0 Then Exit Function
' Get the Opening Balance
OpBalance = GetOpBalance(HeadID, FinIndianFromDate)
' Get the Sum of Debit and Sum of Credit as on date

DebitTotal = 0
CreditTotal = 0
If NewIndexTrans.FindRecord(rstBalance, "HeadID=" & HeadID) Then
    DebitTotal = FormatField(rstBalance.Fields("DebitTotal"))
    CreditTotal = FormatField(rstBalance.Fields("CreditTotal"))
End If
'Depending opon the AccountType calculate the Balance
GetBalanceFromRecordset = GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)
Exit Function
TransIDFailed:
End Function

' Returns the ClosingBalance after Calculation
' Retrurns ClosingBalance
Public Function GetClosingBalance(ByVal HeadID As Long, _
                                  ByVal OpBalance As Currency, _
                                  ByVal DebitTotal As Currency, _
                                  ByVal CreditTotal As Currency) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim AccountType As wis_AccountType
' Intialise the Function
GetClosingBalance = 0
' Validate the inputs
If HeadID = 0 Then Exit Function
' Get AccountType
AccountType = GetAccountType(HeadID)
'Depending opon the AccountType calculate the Balance
GetClosingBalance = _
GetClosingBalanceonAccountType(OpBalance, DebitTotal, CreditTotal, AccountType)
Exit Function
TransIDFailed:
End Function
' Returns the ClosingBalance after Calculation
'
' Retrurns ClosingBalance
'
Public Function GetClosingBalanceonAccountType(ByVal OpBalance As Currency, _
                                    ByVal DebitTotal As Currency, _
                                    ByVal CreditTotal As Currency, _
                                    ByVal AccountType As wis_AccountType) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
' Intialise the Function
GetClosingBalanceonAccountType = 0
'Depending opon the AccountType calculate the Balance
Select Case AccountType
    Case Asset, Loss, ItemPurchase
        GetClosingBalanceonAccountType = OpBalance - CreditTotal + DebitTotal
    Case Liability, Profit, ItemSales
        GetClosingBalanceonAccountType = OpBalance + CreditTotal - DebitTotal
    Case Loss
        GetClosingBalanceonAccountType = OpBalance - CreditTotal + DebitTotal
    Case Profit
        GetClosingBalanceonAccountType = OpBalance + CreditTotal - DebitTotal
    Case ItemPurchase
        GetClosingBalanceonAccountType = OpBalance - CreditTotal + DebitTotal
    Case ItemSales
        GetClosingBalanceonAccountType = OpBalance + CreditTotal - DebitTotal
End Select
Exit Function
TransIDFailed:
End Function


' Returns the AccountType from the Base Table
' Inputs :
'           HeadID as long
'
' Returns AccountType
'
Public Function GetAccountType(ByVal HeadID As Long) As wis_AccountType

' Handle Error
On Error GoTo Hell:
' Declare Variable
Dim rstAccountType As ADODB.Recordset
' Intialise the Function
GetAccountType = 0
' Validate the inputs
If HeadID = 0 Then Exit Function
' Fetch the MaxTransID from Acctrans AND FROM Cash Trans
NewIndexTrans.SQLStmt = " Select AccountType " & _
                   " FROM Heads a,ParentHeads b " & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND a.HeadID=" & HeadID
If NewIndexTrans.Fetch(rstAccountType, adOpenForwardOnly) < 0 Then Exit Function
' Set getMaxtransid
GetAccountType = FormatField(rstAccountType.Fields(0))
Set rstAccountType = Nothing
Exit Function

Hell:
        Set rstAccountType = Nothing
End Function
' Returns the HeadName from the Base Table
' Inputs :
'           HeadID as long
'
' Returns HeadName as string
'
Public Function GetHeadName(ByVal HeadID As Long) As String

' Handle Error
On Error GoTo Hell:
' Declare Variable
Dim rstHeadName As ADODB.Recordset
' Intialise the Function
GetHeadName = ""
' Validate the inputs
If HeadID = 0 Then Exit Function
' Fetch the Head Name from Heads
NewIndexTrans.SQLStmt = " SELECT HeadName " & _
                   " FROM Heads" & _
                   " WHERE HeadID=" & HeadID
If NewIndexTrans.Fetch(rstHeadName, adOpenForwardOnly) < 0 Then Exit Function
' Set HeadName
GetHeadName = FormatField(rstHeadName.Fields(0))
Set rstHeadName = Nothing
Exit Function
Hell:
        Set rstHeadName = Nothing
End Function

'
'This function Saves the given data in the table
Public Function SaveVouchers(ByVal VoucherType As Wis_VoucherTypes, _
                             ByVal TransDate As Date, _
                             Optional ByVal Particulars As String, _
                             Optional ByVal isDataFromInventory As Boolean) _
                             As wis_FunctionReturned
    
' Handle Errors
On Error GoTo CouldNotSave:
' Declare Variables
Dim DrTotal As Currency
Dim CrTotal As Currency
Dim Balance As Currency
Dim UpdatedBalance As Currency
Dim Item As Long
Dim MaxCount As Long
Dim HeadID As Long
'Dim TransID As Long
Dim ObjKeys As Variant
Dim USTransDate As String
' Intialise the Function
SaveVouchers = Failure
' Validate the Inputs and Classes
'Need Not to validate 'SDA
If AllTrans Is Nothing Then Exit Function
If ValidateVouchers(VoucherType) <> Success Then Exit Function
' Check the Collection Dr Amount and Cr Amount Total and it should be same
If DrCrTotalsSame(False) > 0 Then Exit Function
' Convert the TransDdate to US format

'USTransDate = FormatDate(TransDate)
USTransDate = TransDate

' Get the MaxTransID from the Table
If m_TransID < 1 Then TransID = GetMaxTransIDFromTable + 1
'TransID = IIf(TransID, TransID, GetMaxTransIDFromTable + 1)

' Please Don't Change the Case Creiteria
Select Case VoucherType
    Case Payment
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case Receipt
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case Sales
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case Purchase
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case FreePurchase
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case FreeSales
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case Contra
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case Journal
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case RejectionsIn
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
    Case RejectionsOut
        If AllTransCRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, USTransDate) <> Success Then Exit Function
End Select
' Particulars is given it will be inserted in the different table
If Len(Particulars) > 0 Then
    NewIndexTrans.SQLStmt = " INSERT INTO TransParticulars" & _
                     " (TransID,Particulars)" & _
                     " VALUES(" & _
                     TransID & "," & _
                     AddQuotes(Particulars, True) & ")"
    If Not NewIndexTrans.SQLExecute Then Exit Function
End If
If isDataFromInventory Then ' This will Keep Track of the Data's From Inventory Module
    NewIndexTrans.SQLStmt = " INSERT INTO IdFromInventory" & _
                     " (TransID) " & _
                     " VALUES ( " & TransID & ")"
    If Not NewIndexTrans.SQLExecute Then Exit Function
End If
SaveVouchers = Success
Call ClearCollections
'If IsCashTransaction <> Success Then ClearCollections
Exit Function
CouldNotSave:
    SaveVouchers = FatalError
    Call ClearCollections
End Function

' This function Update the given data in the table
Public Function UpdateVouchers(ByVal VoucherType As Wis_VoucherTypes, _
        ByVal TransDate As Date, _
        Optional ByVal Particulars As String) As wis_FunctionReturned
        
        
' Handle Errors
On Error GoTo CouldNotSave:
' Declare Variables
Dim DrTotal As Currency
Dim CrTotal As Currency
Dim Balance As Currency
Dim UpdatedBalance As Currency
Dim Item As Long
Dim MaxCount As Long
Dim HeadID As Long
'Dim TransID As Long
Dim ObjKeys As Variant
Dim USTransDate As String
' Intialise the Function
UpdateVouchers = Failure
' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function
If ValidateVouchers(VoucherType) <> Success Then Exit Function
' Convert the TransDdate to US format
'USTransDate = FormatDate(TransDate)
USTransDate = TransDate

'TransID = TransDateDataExists(USTransDate)
If TransID Then DeleteTransID
UpdateVouchers = SaveVouchers(VoucherType, TransDate)


ClearCollections
Exit Function

CouldNotSave:
    UpdateVouchers = FatalError
End Function
' This function Update the given data in the table
Public Function UndoVouchers(ByVal VoucherType As Wis_VoucherTypes, ByVal TransDate As String, Optional ByVal Particulars As String) As wis_FunctionReturned
' Handle Errors
On Error GoTo CouldNotSave:
' Declare Variables
Dim USTransDate As String
Dim TransID As Long
' Intialise the Function
UndoVouchers = Failure
' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function
If ValidateVouchers(VoucherType) <> Success Then Exit Function
' Convert the TransDdate to US format
USTransDate = FormatDate(TransDate)
TransID = UndoTransDateData(USTransDate)
UndoVouchers = Success
ClearCollections
Exit Function
CouldNotSave:
    UndoVouchers = FatalError
End Function



'Returns the UpdatedBalance
' Inputs :
'           HeadID as long
'           Balance as currency
'           DrCrType as wis_DRCRType
'
' Retrurns Updated Balance
'
Private Function GetUpdatedBalance(ByVal HeadID As Long, ByVal Amount As Currency, ByVal Balance As Currency, DRCRType As wis_DrCrType) As Currency
' Handle Error
On Error GoTo Hell:
' Declare Variable
Dim rstHeadType As ADODB.Recordset
Dim AccountType As wis_AccountType
' Intialise the Function
GetUpdatedBalance = 0
' Validate the inputs
If HeadID = 0 Then Exit Function
'If Balance = 0 Then Exit Function
AccountType = GetAccountType(HeadID)
Select Case AccountType
    Case Loss
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance + Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance - Amount
    Case Liability
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance - Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance + Amount
    Case ItemSales
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance - Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance + Amount
    Case Asset
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance + Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance - Amount
    Case Profit
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance - Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance + Amount
    Case ItemPurchase
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance + Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance - Amount
End Select
Set rstHeadType = Nothing
Exit Function
Hell:
Set rstHeadType = Nothing
End Function

'This sub will add the given data to DrHeads Collection
Public Function AllTransHeadsAdd(ByVal HeadID As Long, ByVal DrAmount As Currency, ByVal CrAmount As Currency) As wis_FunctionReturned
On Error GoTo CouldNotAdd:
Dim ParentID As Long
AllTransHeadsAdd = Failure
' check the collection
If AllTrans Is Nothing Then Exit Function
    ' check if the headid exists in the both collection
    If ALlTransHeadIDExists(HeadID) Then
        MsgBox "Head is already Exists in Collection !"
        Exit Function
    End If
ParentID = GetParentID(HeadID)
Call AllTrans.Add(HeadID, DrAmount, CrAmount, ParentID, "Tr" & HeadID)
AllTransHeadsAdd = Success
Exit Function
CouldNotAdd:
    If Err.Number = 457 Then MsgBox " Head is already Exists in Collection !"
    AllTransHeadsAdd = FatalError
End Function
'This sub will add the given data to DrHeads Collection
Public Function AllTransHeadsUpdate(ByVal HeadID As Long, ByVal DrAmount As Currency, ByVal CrAmount As Currency) As wis_FunctionReturned
On Error GoTo CouldNotAdd:
Dim ParentID As Long
AllTransHeadsUpdate = Failure
' check the collection
If AllTrans Is Nothing Then Exit Function
If PrevHeadID = 0 Then Exit Function
ParentID = GetParentID(HeadID)
Call AllTrans.Update(HeadID, DrAmount, CrAmount, ParentID, PrevHeadID, "Tr" & HeadID)
AllTransHeadsUpdate = Success
Exit Function
CouldNotAdd:
    If Err.Number = 457 Then MsgBox " Head is already Exists in Collection !"
    AllTransHeadsUpdate = FatalError
End Function





'This sub will remove the based the given data given data to DrHeads Collection
Public Sub AllTransHeadIDRemove(ByVal HeadID As Long)
On Error GoTo CouldNotRemove:
If AllTrans Is Nothing Then Exit Sub
    Call AllTrans.Remove("Tr" & HeadID)
Exit Sub
CouldNotRemove:
    If Err.Number = 5 Then MsgBox "Head does not Exist in the Collection!"
End Sub

'
'This function will update the base Collections
Private Function UpdateCollections() As wis_FunctionReturned
On Error GoTo Hell:
' Declarations
Dim Balance As Currency
Dim UpdatedBalance As Currency
Dim TransID As Long
Dim Amount As Currency
Dim DrCrEntry As wis_DrCrType
' Initialise the Function
UpdateCollections = Failure
' Check the Form
'If m_frmAccTrans Is Nothing Then Exit Function
Call AllTransHeadsAdd(HeadID, Debit, Credit)
UpdateCollections = Success
Exit Function
Hell:
    UpdateCollections = FatalError
End Function

' This function returns the ParentName from the given Headid
' Input is Headid as long
' Returns ParentName string
'
' Pradeep
'
Public Function GetParentName(ByVal HeadID As Long) As String
' Handle Error
On Error GoTo NoParentName:
' Declare Variables
Dim rstParentName As ADODB.Recordset
' Intialiase the Variable
GetParentName = ""
' Check the Input Received if Zero then Exit
If HeadID = 0 Then Exit Function
' set the sqlstmt
NewIndexTrans.SQLStmt = " SELECT ParentName " & _
                   " FROM Heads a,ParentHeads b " & _
                   " WHERE a.HeadID=" & HeadID & _
                   " AND a.ParentID=b.ParentID"
' Now fetch the record
If NewIndexTrans.Fetch(rstParentName, adOpenForwardOnly) < 0 Then Exit Function
' Here is the ParentID!
GetParentName = FormatField(rstParentName.Fields("ParentName"))
Set rstParentName = Nothing
Exit Function
NoParentName:
    Set rstParentName = Nothing
End Function

'
Private Sub Class_Terminate()
On Error Resume Next
    Set m_AllTrans = Nothing
    'Set m_frmAccTrans = Nothing
End Sub



' This will return adodb.recordset if transid has mulitple credit and debit heads
' takes TransID as Long as Argument

Private Function GetMultiHeads() As ADODB.Recordset

On Error GoTo Hell:

Dim rstMultiHeads As ADODB.Recordset

Dim QryNametoFetch As String

Set GetMultiHeads = Nothing

' This will prepare queries
QryNametoFetch = CreateViewsforMultiHeads

' This means creation of qry is failed
If QryNametoFetch = "" Then Exit Function
    
NewIndexTrans.SQLStmt = " SELECT * FROM " & QryNametoFetch

If NewIndexTrans.Fetch(rstMultiHeads, adOpenStatic) = -1 Then Exit Function

Set GetMultiHeads = rstMultiHeads

Set rstMultiHeads = Nothing

Exit Function

Hell:
    
    Set rstMultiHeads = Nothing

End Function


Private Function CreateViewsforShowLedger(ByVal DateToConsider As Boolean, _
                                          ByRef rstLedger As ADODB.Recordset, _
                                          ByVal FromDate As String, _
                                          ByVal ToDate As String, _
                                          ByVal HeadID As Long) As Long
' Set the Sql Statement
#If Junk Then
If DateToConsider Then
        
        
    NewIndexTrans.SQLStmt = " SELECT a.Headid, Debit, Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM qryAccTransMerge AS a, Heads AS b " & _
                      " WHERE a.HeadID = " & HeadID & _
                      " AND a.HeadID = b.HeadID " & _
                      " AND a.Debit=0" & _
                      " AND TransDate BETWEEN #" & FormatDate(FromDate) & "#" & _
                      " AND #" & FormatDate(ToDate) & "#"
        
    NewIndexTrans.CreateView ("QryCrHeads")
    
    NewIndexTrans.SQLStmt = " SELECT a.Headid, Debit, Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM qryAccTransMerge AS a, Heads AS b " & _
                      " WHERE a.HeadID = " & HeadID & _
                      " AND a.HeadID = b.HeadID " & _
                      " AND a.Credit=0" & _
                      " AND TransDate BETWEEN #" & FormatDate(FromDate) & "#" & _
                      " AND #" & FormatDate(ToDate) & "#"
        
    NewIndexTrans.CreateView ("QryDrHeads")
    
End If

If Not DateToConsider Then
        
    NewIndexTrans.SQLStmt = " SELECT CrHeadid AS HeadID, 0 AS Debit, Amount AS Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM AccTrans AS a, Heads AS b " & _
                      " WHERE DrHeadID = " & HeadID & _
                      " AND a.CrHeadID = b.HeadID "
        
    NewIndexTrans.CreateView ("QryCrHeads")
    
    
    NewIndexTrans.SQLStmt = " SELECT DrHeadid AS HeadID, Amount AS Debit, 0 AS Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM AccTrans AS a, Heads AS b " & _
                      " WHERE CrHeadID = " & HeadID & _
                      " AND a.DrHeadID = b.HeadID "

    NewIndexTrans.CreateView ("QryDrHeads")
    
End If
                        
NewIndexTrans.SQLStmt = " SELECT * FROM qryCrHeads " & _
                   " UNION ALL " & _
                   " SELECT * FROM qryDrHeads "
                   
If Not NewIndexTrans.CreateProcedures("QryLedger") Then Exit Function
    

NewIndexTrans.SQLStmt = " SELECT * " & _
                   " FROM QryLedger" & _
                   " ORDER BY TransDate,TransID"
                
'Fetch the Data to Recordset based on the Condition
CreateViewsforShowLedger = NewIndexTrans.Fetch(rstLedger, adOpenStatic)
#End If
End Function

'This prepare the database to return the recordset of multiple heads
' it will create the conserned quiries
Private Function CreateViewsforMultiHeads() As String

On Error GoTo Hell:

CreateViewsforMultiHeads = ""

NewIndexTrans.SQLStmt = " SELECT HeadID as CrHeadID, TransID " & _
                 " FROM AccTrans " & _
                 " WHERE Credit > 0 " & _
                 " ORDER BY TRANSID"
                 
If Not NewIndexTrans.CreateView("QryCreditHeads") Then Exit Function

NewIndexTrans.SQLStmt = " SELECT HeadID as DrHeadID, TransID " & _
                 " FROM AccTrans " & _
                 " WHERE Debit > 0 " & _
                 " ORDER BY TRANSID"
                 
If Not NewIndexTrans.CreateView("QryDebitHeads") Then Exit Function

NewIndexTrans.SQLStmt = " SELECT DISTINCT * " & _
                   " FROM  QryCreditHeads A ,QryDebitHeads B " & _
                   " WHERE A.TransID = B.TransID " '& _
                   " AND A.CrHeadID=B.DrHeadID "

If Not NewIndexTrans.CreateView("QryMultiHeads") Then Exit Function

CreateViewsforMultiHeads = "QryMultiHeads"

Exit Function

Hell:
    
End Function



