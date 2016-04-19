VERSION 5.00
Begin VB.UserControl CashIndex 
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "CashIdx.ctx":0000
   PropertyPages   =   "CashIdx.ctx":0442
   ScaleHeight     =   1050
   ScaleWidth      =   1350
   ToolboxBitmap   =   "CashIdx.ctx":0455
End
Attribute VB_Name = "CashIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum PaymentModeValue
    Cash
    Cheque
End Enum

Const def_PaymentMode = 0
Const def_DialogTitle = " "

Private WithEvents m_CashIndex As frmCashIndex
Attribute m_CashIndex.VB_VarHelpID = -1
Private WithEvents m_ChequeIndex As frmCheque
Attribute m_ChequeIndex.VB_VarHelpID = -1
Public DenominationReceived As Denomination
Public DenominationRefunded As Denomination
Dim m_ExpectedCash As Currency
Dim m_DialogTitle As String
Dim m_Caption As String
Dim m_CancelError As Boolean
Dim m_CancelClicked As Boolean
Dim m_CashReceiveError As Boolean
Dim m_CashMismatch As Boolean
Dim m_PaymentMode As PaymentModeValue
Public Property Let CancelError(NewValue As Boolean)
    m_CancelError = NewValue
    PropertyChanged ("CancelError")
End Property
Public Property Get CancelError() As Boolean
    CancelError = m_CancelError
End Property
Public Property Let CashReceiveError(NewVal As Boolean)
m_CashReceiveError = NewVal
PropertyChanged ("CashreceiveError")
End Property
Public Property Get CashReceiveError() As Boolean
    CashReceiveError = m_CashReceiveError
End Property

Public Property Let DialogTitle(ByVal NewVal As String)

    If NewVal <> "" Then
        m_DialogTitle = NewVal
    Else
        m_DialogTitle = def_DialogTitle
    End If
        PropertyChanged ("DialogTitle")
End Property
Public Property Get DialogTitle() As String
    DialogTitle = m_DialogTitle
End Property
Public Property Let ExpectedCash(ByVal NewVal As Currency)
    m_ExpectedCash = NewVal
    PropertyChanged ("ExpectedCash")
End Property

Public Property Get ExpectedCash() As Currency
    ExpectedCash = m_ExpectedCash
End Property
Public Property Get PaymentMode() As PaymentModeValue
    PaymentMode = m_PaymentMode
End Property

Public Property Let PaymentMode(NewValue As PaymentModeValue)
    m_PaymentMode = NewValue
    PropertyChanged ("PaymentMode")
End Property

Public Sub Show()
    If PaymentMode = Cash Then
        Load m_CashIndex
        
        m_CashIndex.Caption = m_DialogTitle
        m_CashIndex.txtExpectedCash.Text = CStr(m_ExpectedCash)
        m_CashIndex.Show vbModal
    ElseIf PaymentMode = Cheque Then
        Load m_ChequeIndex
        m_ChequeIndex.Caption = m_DialogTitle
        m_ChequeIndex.txtExpectedCash.Text = CStr(m_ExpectedCash)
        m_ChequeIndex.Show vbModal
    End If
    If CancelError And m_CancelClicked Then
        Err.Raise 32755, , "Cancel was selected"
    End If
    If CashReceiveError And m_CashMismatch Then
        Err.Raise 520, , "Cash Mismatch"
    End If
End Sub
Private Sub m_CashIndex_CancelClicked()
    m_CancelClicked = True
End Sub
Private Sub m_CashIndex_OKClicked()
    'This event will be raised before you unload the frmidx form
    'First check if the Cash Receive Error is on
    If m_CashReceiveError And Val(m_CashIndex.txtExpectedCash.Text) <> Val(m_CashIndex.txtNetAmount.Text) Then
        m_CashMismatch = True
        Exit Sub
    End If
    

    With DenominationReceived
        .Rs500 = Val(m_CashIndex.txtCashIn(0).Text)
        .Rs100 = Val(m_CashIndex.txtCashIn(1).Text)
        .Rs50 = Val(m_CashIndex.txtCashIn(2).Text)
        .Rs20 = Val(m_CashIndex.txtCashIn(3).Text)
        .Rs10 = Val(m_CashIndex.txtCashIn(4).Text)
        .Rs5 = Val(m_CashIndex.txtCashIn(5).Text)
        .Rs2 = Val(m_CashIndex.txtCashIn(6).Text)
        .Rs1 = Val(m_CashIndex.txtCashIn(7).Text)
        .Coins = Val(m_CashIndex.txtCashIn(8).Text)
    End With
    
    With DenominationRefunded
        .Rs500 = Val(m_CashIndex.txtCashOut(0).Text)
        .Rs100 = Val(m_CashIndex.txtCashOut(1).Text)
        .Rs50 = Val(m_CashIndex.txtCashOut(2).Text)
        .Rs20 = Val(m_CashIndex.txtCashOut(3).Text)
        .Rs10 = Val(m_CashIndex.txtCashOut(4).Text)
        .Rs5 = Val(m_CashIndex.txtCashOut(5).Text)
        .Rs2 = Val(m_CashIndex.txtCashOut(6).Text)
        .Rs1 = Val(m_CashIndex.txtCashOut(7).Text)
        .Coins = Val(m_CashIndex.txtCashOut(8).Text)
    End With
    m_CancelClicked = False
    m_CashMismatch = False
End Sub

Private Sub UserControl_Initialize()
Set m_CashIndex = New frmCashIndex
Set m_ChequeIndex = New frmCheque
m_PaymentMode = Cash
Set DenominationReceived = New Denomination
Set DenominationRefunded = New Denomination
End Sub
Private Sub UserControl_InitProperties()
m_CancelError = False
m_ExpectedCash = 0
m_CancelClicked = False
m_CashReceiveError = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DialogTitle = PropBag.ReadProperty("DialogTitle", def_DialogTitle)
    CancelError = PropBag.ReadProperty("CancelError", False)
    ExpectedCash = PropBag.ReadProperty("ExpectedCash", 0)
    CashReceiveError = PropBag.ReadProperty("CashReceiveError", False)
    PaymentMode = PropBag.ReadProperty("PaymentMode", def_PaymentMode)
End Sub
Private Sub UserControl_Resize()
UserControl.Height = 400
UserControl.Width = 400
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DialogTitle", m_DialogTitle, def_DialogTitle)
    Call PropBag.WriteProperty("CancelError", m_CancelError, False)
    'Call PropBag.WriteProperty("ReceivedDenomination")
    Call PropBag.WriteProperty("ExpectedCash", m_ExpectedCash, "0")
    Call PropBag.WriteProperty("CashReceiveError", m_CashReceiveError, False)
    Call PropBag.WriteProperty("PaymentMode", m_PaymentMode, def_PaymentMode)
End Sub
