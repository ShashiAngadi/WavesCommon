VERSION 5.00
Begin VB.Form frmPDRepay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repayment"
   ClientHeight    =   3360
   ClientLeft      =   4065
   ClientTop       =   2925
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   0
      TabIndex        =   11
      Top             =   -30
      Width           =   3045
      Begin VB.TextBox txtLoanAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtInterestAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1170
         Width           =   975
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtInterestRate 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Total loan amount:"
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label lblInterestAmount 
         Caption         =   "Interest accrued : "
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label lblTotalAmount 
         Caption         =   "Net repayable amount : "
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   1530
         Width           =   1725
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label lblInterestRate 
         Caption         =   "Rate of interest:"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   90
      TabIndex        =   10
      Top             =   2700
      Width           =   2835
   End
   Begin VB.TextBox txtRepaidAmount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtBalanceAmount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   2910
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   570
      TabIndex        =   7
      Top             =   2910
      Width           =   1125
   End
   Begin VB.Label lblRepaidAmount 
      Caption         =   "Repaid amount : "
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1950
      Width           =   1725
   End
   Begin VB.Label lblBalanceAmount 
      Caption         =   "Balance loan amount : "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2310
      Width           =   1725
   End
End
Attribute VB_Name = "frmPDRepay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_AccID As Long
Private m_UserID As Integer


Private Sub UpdateDetails()
Dim TransType As wisTransactionTypes
Dim LoanAmt As Currency
Dim RepayAmt As Currency
Dim Transdate As String
Dim Days As Integer
Dim InterestRate As Double

'Get the last date of transaction
    gDBTrans.SQLStmt = "Select TOP 1 TransDate from PDTrans where " & _
                        " AccID = " & m_AccID & _
                        " and UserId = " & m_UserID & _
                        " and Loan = True order by TransID desc "

    If gDBTrans.SQLFetch <> 1 Then
        GoTo ErrLine
    End If
    Transdate = FormatField(gDBTrans.Rst("TransDate"))
    
    InterestRate = GetPDLoanInterest(100, gStrDate)
        txtInterestRate.Text = Format(InterestRate, "#0.00")
    
'Get Loan Balance as on today
    TransType = wDeposit
    gDBTrans.SQLStmt = "Select top 1 Balance from PDTrans where " & _
                        " AccId = " & m_AccID & _
                        " and UserId = " & m_UserID & _
                        " And Loan = True Order by transid desc "
                        
    If gDBTrans.SQLFetch <> 1 Then
        GoTo ErrLine
    End If
    LoanAmt = Val(FormatField(gDBTrans.Rst(0)))
    txtLoanAmount.Text = FormatCurrency(LoanAmt)
                    

'Calculate the interest accrued
    Days = WisDateDiff(Transdate, txtDate.Text) '
    'txtInterestAmount.Text = FormatCurrency(frmPDLoans.ComputePDLoanInterest((LoanAmt - RepayAmt), Days, InterestRate))
    
'Total payable amount
    txtTotalAmount.Text = FormatCurrency(Val(txtLoanAmount.Text) + Val(txtInterestAmount.Text))
    txtBalanceAmount.Text = txtTotalAmount.Text

'txtRepaidAmount.Text = "0.00"
Exit Sub

ErrLine:
'MsgBox "No loans have been issued on this deposit !", vbExclamation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 582), vbExclamation, gAppName & " - Error"

End Sub

Private Sub cmdAccept_Click()
Dim Transdate As String
Dim TransType As wisTransactionTypes
Dim TransID As Long

'Check the date
    If Not DateValidate(txtDate.Text, "/", True) Then
        'MsgBox "Date not specified in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Sub
    End If
    
'Check last date of transaction
    gDBTrans.SQLStmt = "Select max(TransDate) from PDTrans where " & _
                        " Loan = False And AccID = " & m_AccID & _
                        " and UserId = " & m_UserID
    Call gDBTrans.SQLFetch
    Transdate = FormatField(gDBTrans.Rst(0))
    
If WisDateDiff(Transdate, txtDate.Text) < 0 Then
    'MsgBox "Date specified should be later than the date of last transaction on this account !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
    
If Not CurrencyValidate(txtRepaidAmount.Text, True) Or Val(txtBalanceAmount.Text) < 0 Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtRepaidAmount
    Exit Sub
End If

If Val(txtRepaidAmount.Text) < 0 Then
    'MsgBox "Amount repaid is greater that total amount to be paid !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 585), vbExclamation, gAppName & " - Error"
    Exit Sub
End If
    
'Get new transID
Dim Balance As Currency
    gDBTrans.SQLStmt = "Select TOP 1 TransID,Balance from PDTrans where " & _
                        " AccID = " & m_AccID & _
                        " and UserId = " & m_UserID & " And Loan = True " & _
                        " order by TransID desc"
    Call gDBTrans.SQLFetch
    TransID = FormatField(gDBTrans.Rst("TransID")) + 1
    Balance = FormatField(gDBTrans.Rst("Balance"))
    
'Start database operations
gDBTrans.BeginTrans
    'Insert Interest first
    TransType = wCharges
    gDBTrans.SQLStmt = "Insert into PDTrans (AccID, UserId, TransID, TransType," & _
                        " TransDate, Amount, Balance, Loan," & _
                        " Particulars ) values ( " & _
                        m_AccID & ", " & _
                        m_UserID & ", " & _
                        TransID & ", " & _
                        TransType & ", " & _
                        "#" & FormatDate(txtDate.Text) & "#, " & _
                        Val(txtInterestAmount.Text) & ", " & _
                        Balance & ", " & True & ", " & _
                        "'" & "To Loans" & "')"

    If Not gDBTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 535), vbExclamation, gAppName & " - Error"
        gDBTrans.RollBack
        Exit Sub
    End If

    'Now insert repaid amount
    TransID = TransID + 1
    Balance = Val(txtBalanceAmount.Text)
    TransType = wDeposit
    gDBTrans.SQLStmt = "Insert into PDTrans (AccID, UserId, TransID, TransType, " & _
                        " TransDate, Amount,Balance, Loan, " & _
                        " Particulars ) values ( " & _
                        m_AccID & "," & _
                        m_UserID & "," & _
                        TransID & "," & _
                        TransType & "," & _
                        "#" & FormatDate(txtDate.Text) & "#," & _
                        Val(txtRepaidAmount.Text) - Val(txtInterestAmount.Text) & "," & _
                        Balance & ", True ," & _
                        "'" & "To Loans" & "')"

    If Not gDBTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 535), vbExclamation, gAppName & " - Error"
        gDBTrans.RollBack
        Exit Sub
    End If
    
gDBTrans.CommitTrans
txtDate.Text = FormatDate(gStrDate)
txtRepaidAmount.Text = "0.00"
Call UpdateDetails
'MsgBox "Repayment made successfully !", vbInformation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 586), vbInformation, gAppName & " - Error"
Unload Me
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub Form_Load()
    
'Center the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'Set kannada fonts
Call SetKannadaCaption

'Load values to the text box
    m_AccID = frmPDLoans.m_AccID
    m_UserID = frmPDLoans.m_UserID
        
'Todays date
    txtDate.Text = FormatDate(gStrDate)
    
Call UpdateDetails
End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
Set frmPDRepay = Nothing
End Sub


Private Sub txtDate_LostFocus()
If Not DateValidate(txtDate.Text, "/", True) Then
    Exit Sub
End If
Call UpdateDetails

End Sub

Private Sub txtInterestAmount_Change()
txtTotalAmount = FormatCurrency(Val(txtLoanAmount.Text) + Val(txtInterestAmount.Text))
End Sub

Private Sub txtLoanAmount_Change()
txtTotalAmount = FormatCurrency(Val(txtLoanAmount.Text) + Val(txtInterestAmount.Text))
End Sub


Private Sub txtRepaidAmount_Change()
If Not CurrencyValidate(txtRepaidAmount.Text, True) Then
    txtBalanceAmount.Text = FormatCurrency(Val(txtTotalAmount.Text))
    Exit Sub
End If

txtBalanceAmount.Text = FormatCurrency(Val(txtTotalAmount.Text) - Val(txtRepaidAmount.Text))

End Sub

Private Sub SetKannadaCaption()
Dim Ctrl As Control
If gLangOffSet = wis_KannadaOffset Then
    For Each Ctrl In Me
        Ctrl.Font.Name = gFontName
        If Not TypeOf Ctrl Is ComboBox Then
          Ctrl.Font.Size = gFontSize
        End If
    Next Ctrl
End If
On Error GoTo 0
Me.lblDate.Caption = LoadResString(gLangOffSet + 37)   ' "ŠÂñ®°ð"
Me.lblInterestRate.Caption = LoadResString(gLangOffSet + 186)   ' "ÄˆÜ®õðô ÁðÇð"
Me.lblLoanAmount.Caption = LoadResString(gLangOffSet + 235)   ' "«»ôÚ ÍñÈÁð Æú÷ÀðÞ"
Me.lblInterestAmount.Caption = LoadResString(gLangOffSet + 252)    '"ÄˆÜ"
Me.lblTotalAmount.Caption = LoadResString(gLangOffSet + 253)    '"«»ôÚ Æú÷ÀðÞ"
Me.lblRepaidAmount.Caption = LoadResString(gLangOffSet + 254)    '"ÆðôÇðôÃñÆð‰ Æú÷ÀðÞ"
Me.lblBalanceAmount = LoadResString(gLangOffSet + 255)   '"ÍñÈÁð ¥’Áð Æú÷ÀðÞ"
Me.cmdAccept.Caption = LoadResString(gLangOffSet + 4)    '"¡®‚ó°ðÍðô"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 2)    '"ÇðÁðôà"

End Sub

