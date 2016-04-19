VERSION 5.00
Begin VB.Form frmPDInterest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pigmy Interest"
   ClientHeight    =   3435
   ClientLeft      =   2865
   ClientTop       =   1950
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   2025
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   3555
      Begin VB.TextBox txtDepositAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox txtInterestAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1950
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Height          =   345
         Left            =   1950
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtInterestRate 
         Height          =   345
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   285
         Left            =   3180
         TabIndex        =   7
         Top             =   180
         Width           =   285
      End
      Begin VB.Label lblDepositAmount 
         Caption         =   "Deposit amount:"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   1050
         Width           =   1695
      End
      Begin VB.Label lblInterestAmount 
         Caption         =   "Interest accrued : "
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   1470
         Width           =   1725
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblInterestRate 
         Caption         =   "Rate of interest:"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   1695
      End
   End
   Begin VB.TextBox txtPaidAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   2100
      Width           =   1095
   End
   Begin VB.TextBox txtBalanceAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   2490
      Width           =   1125
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   3030
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   3030
      Width           =   945
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   3630
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label lblPaidAmount 
      Caption         =   "Repaid amount : "
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   2130
      Width           =   1725
   End
   Begin VB.Label lblBalanceAmount 
      Caption         =   "Balance loan amount : "
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   2460
      Width           =   1725
   End
End
Attribute VB_Name = "frmPDInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Accid As Long
Private m_UserID As Integer

Private Sub SetKannadaCaption()
Dim ctrl As Control
If Not gLangOffSet = wis_KannadaOffset Then Exit Sub
For Each ctrl In Me
    ctrl.Font.Name = gFontName
    If Not TypeOf ctrl Is ComboBox Then
      ctrl.Font.Size = gFontSize
    End If
Next ctrl
On Error GoTo 0
lblDate.Caption = LoadResString(gLangOffSet + 37)   '
lblInterestRate.Caption = LoadResString(gLangOffSet + 186)
lblDepositAmount.Caption = LoadResString(gLangOffSet + 243)
lblInterestAmount.Caption = LoadResString(gLangOffSet + 252)
lblPaidAmount.Caption = LoadResString(gLangOffSet + 487)
lblBalanceAmount = LoadResString(gLangOffSet + 243)
cmdAccept.Caption = LoadResString(gLangOffSet + 4)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)

End Sub

Private Sub UpdateDetails()
Dim TransType As wisTransactionTypes
Dim LoanAmt As Currency
Dim RepayAmt As Currency
Dim Balance As Currency
Dim TransDate As String
Dim Days As Long
Dim InterestRate As Double
Dim Loan As Boolean
Dim Rst As Recordset

'Get the last date of transaction
    Loan = False
    gDbTrans.SQLStmt = "Select TOP 1 TransDate, Balance from PDTrans where " & _
                        " AccID = " & m_Accid & _
                        "And UserID = " & m_UserID & _
                        " and Loan = " & Loan & " Order by TransID desc "

    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo ErrLine
    
    TransDate = FormatField(Rst("TransDate"))
    Balance = CCur(FormatField(Rst("Balance")))
    Me.txtBalanceAmount.Text = FormatCurrency(Balance)
    
'Get Rate of Interest For This deposit
    gDbTrans.SQLStmt = "Select RateOfInterest From PDMaster Where " & _
                        " AccID = " & m_Accid
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo ErrLine
    
    InterestRate = FormatField(Rst("RateOfInterest"))
    txtInterestRate.Text = Format(InterestRate, "#0.00")
    
'Calculate the interest till date given
   Dim IntAmount As Currency
    If Not DateValidate(Trim(txtDate.Text), "/", True) Then
        'Days = WisDateDiff(Transdate, FormatDate(gStrDate))
        IntAmount = FormatCurrency(ComputePDInterest(Balance, CSng(InterestRate)))
    Else
       IntAmount = FormatCurrency(ComputePDInterest(Balance, CSng(InterestRate)))
    End If
    txtInterestAmount.Text = FormatCurrency(IntAmount \ 1)
    Me.txtPaidAmount = txtInterestAmount
    Me.txtDepositAmount = FormatCurrency(Balance)
    
'Total payable amount
    'txtTotalAmount.Text = FormatCurrency(Balance)
    txtBalanceAmount.Text = FormatCurrency(Balance)


Exit Sub

ErrLine:
'MsgBox "No loans have been issued on this deposit !", vbExclamation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 582), vbExclamation, gAppName & " - Error"

End Sub

Private Sub cmdAccept_Click()
Dim TransDate As Date
Dim TransType As wisTransactionTypes
Dim TransID As Long
Dim Balance As Currency
Dim Loan As Boolean

'Check the date
    If Not DateValidate(txtDate.Text, "/", True) Then
        'MsgBox "Date not specified in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Sub
    End If
    TransDate = FormatDate(txtDate)
    
'Check last date of transaction
    Dim Rst As Recordset
    gDbTrans.SQLStmt = "Select TOP 1 TransDate from PDTrans where " & _
                        " AccID = " & m_Accid & _
                        " And UserID = " & m_UserID & _
                        " order by TransID desc"
    
    Call gDbTrans.Fetch(Rst, adOpenDynamic)
    TransDate = Rst("TransDate")
    
If DateDiff("d", TransDate, FormatDate(txtDate.Text)) < 0 Then
    'MsgBox "Date specified should be later than the date of last transaction on this account !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
    
If Not CurrencyValidate(txtPaidAmount.Text, False) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPaidAmount
    Exit Sub
End If

If Val(txtPaidAmount.Text) < 0 Then
    'MsgBox "Amount repaid is greater that total amount to be paid !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 585), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

If Not CurrencyValidate(txtBalanceAmount.Text, False) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtBalanceAmount
    Exit Sub
End If

If Val(txtBalanceAmount.Text) < 0 Then
    'MsgBox "Amount repaid is greater that total amount to be paid !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 585), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtBalanceAmount
    Exit Sub
End If
 
'Get new transID
    Loan = False
    gDbTrans.SQLStmt = "Select TOP 1 TransID ,Balance from PDTrans where " & _
                        " Loan = " & Loan & " And AccID = " & m_Accid & _
                        "And UserID = " & m_UserID & _
                        " order by TransID desc"
    Call gDbTrans.Fetch(Rst, adOpenForwardOnly)
    TransID = FormatField(Rst("TransID")) + 1
    Balance = CCur(FormatField(Rst("Balance")))

'Check whether The Deposit balance is correct or not
If Val(txtBalanceAmount.Text) < Balance Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtBalanceAmount
    Exit Sub
End If

If Val(txtBalanceAmount.Text) > Balance Then
    'if MsgBox ("You have entered loan deposit balance more than prevous balance !"
   '     " do You want to continue ?", vbExclamation+vbyesno, gAppName & " - Error") = vbno then
    If MsgBox(LoadResString(gLangOffSet + 584) & " " & LoadResString(gLangOffSet + 541), _
          vbYesNo + vbInformation + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
         ActivateTextBox txtBalanceAmount
         Exit Sub
   End If
End If

Dim Amount As Currency
Amount = txtPaidAmount.Text
Balance = txtBalanceAmount.Text
'CHECK fOR THE TRANSACTION
gDbTrans.SQLStmt = "sELECT * FROM PDTrans Where AccId = " & m_Accid & _
                " And UserID = " & m_UserID & _
                " And TransDate > #" & FormatDate(txtDate.Text) & "#"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    MsgBox LoadResString(gLangOffSet + 572), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
'Start database operations
gDbTrans.BeginTrans
    'Insert Interest first
    MsgBox "Check the Code"
    TransType = wWithdraw
    gDbTrans.SQLStmt = "Insert into PDTrans (AccID, Loan, TransID, TransType, " & _
                        " TransDate, Amount, Balance," & _
                        " Particulars, UserID) values ( " & _
                        m_Accid & "," & _
                        Loan & ", " & _
                        TransID & "," & _
                        TransType & "," & _
                        "#" & FormatDate(txtDate.Text) & "#," & _
                        Amount & "," & _
                        Balance & ", " & _
                        "'" & "int Paid" & "', " & _
                        m_UserID & _
                        ")"
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 535), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Sub
    End If
    
gDbTrans.CommitTrans
txtDate.Text = FormatDate(gStrDate)
txtPaidAmount.Text = "0.00"
Call UpdateDetails
'MsgBox "Repayment made successfully !", vbInformation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 586), vbInformation, gAppName & " - Error"
Unload Me
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDate_Click()
With Calendar
    .Left = Me.Left + cmdDate.Left - .Width / 2
    .Top = Me.Top + cmdDate.Top
    .SelDate = txtDate.Text
    .Show vbModal
    txtDate.Text = .SelDate
End With
End Sub


Private Sub Form_Load()
    
'Center the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
'Set kannada fonts
Call SetKannadaCaption
'Load values to the text box
    txtDate.Text = FormatDate(gStrDate)
    m_Accid = frmPDAcc.m_Accid
    'm_UserID = frmPDAcc.m_UserID
'Todays date
'    txtDate.Text = Formatdate(gStrDate)
    
Call UpdateDetails
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPDInterest = Nothing
End Sub


Private Sub lblTotalAmount_Click()

End Sub

Private Sub txtDate_LostFocus()
If Not DateValidate(txtDate.Text, "/", True) Then
    Exit Sub
End If
Call UpdateDetails

End Sub
