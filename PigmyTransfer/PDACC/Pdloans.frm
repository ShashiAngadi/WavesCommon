VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPDLoans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loans"
   ClientHeight    =   6345
   ClientLeft      =   1830
   ClientTop       =   1095
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   6660
      TabIndex        =   9
      Top             =   5850
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   5625
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7665
      Begin VB.CommandButton cmdRepay 
         Caption         =   "Repay"
         Height          =   315
         Left            =   5070
         TabIndex        =   7
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   150
         TabIndex        =   27
         Top             =   2970
         Width           =   7365
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Undo"
         Height          =   315
         Left            =   3810
         TabIndex        =   8
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   120
         TabIndex        =   24
         Top             =   630
         Width           =   7365
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Height          =   315
         Left            =   6330
         TabIndex        =   6
         Top             =   3120
         Width           =   1125
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2550
         TabIndex        =   1
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox txtAvailable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1770
         Width           =   1065
      End
      Begin VB.TextBox txtLoan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2100
         Width           =   1065
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox txtSanctioned 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   780
         Width           =   1065
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   150
         TabIndex        =   12
         Top             =   3510
         Width           =   7365
      End
      Begin VB.TextBox txtInterest 
         Height          =   285
         Left            =   2550
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox txtInterestAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox txtIssuedAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   1440
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   1815
         Left            =   150
         TabIndex        =   23
         Top             =   3630
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   3201
         _Version        =   393216
         ScrollBars      =   2
         AllowUserResizing=   1
      End
      Begin VB.Label lblcaption1 
         Caption         =   "New loans will be issued only after deducting interest prevailing on previous loans first."
         Height          =   495
         Left            =   3960
         TabIndex        =   26
         Top             =   1890
         Width           =   3435
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCaption 
         Caption         =   "Total loan amount drawn on this deposit:"
         Height          =   225
         Left            =   1920
         TabIndex        =   25
         Top             =   270
         Width           =   4875
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   780
         Width           =   2025
      End
      Begin VB.Label lblLoanAmtAvail 
         Caption         =   "Available loan amount : "
         Height          =   285
         Left            =   180
         TabIndex        =   21
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label lblPrevLoanAmt 
         Caption         =   "Previous loan amount : "
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   2100
         Width           =   2085
      End
      Begin VB.Label lblDepAmount 
         Caption         =   "Deposit amount : "
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Top             =   1110
         Width           =   2145
      End
      Begin VB.Label lblLoanSanctioned 
         Caption         =   "Sanctioned amount : "
         Height          =   255
         Left            =   3990
         TabIndex        =   18
         Top             =   810
         Width           =   2265
      End
      Begin VB.Label lblRateofIntForLoans 
         Caption         =   "Rate of interest for loans:"
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label lblDepositNo 
         Caption         =   "Deposit No: "
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblLessIntOnPrevLoan 
         Caption         =   "Less interest on previous loans: "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3990
         TabIndex        =   15
         Top             =   1140
         Width           =   2295
      End
      Begin VB.Label lblTotAmtIssued 
         Caption         =   "Total amount to be issued:"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Top             =   1470
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmPDLoans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_AccID As Long
Public m_UserID As Integer
Private M_setUp As New clsSetup

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
Me.lblDepositNo.Caption = LoadResString(gLangOffSet + 241)   ' "«œ˙Û∆æÚ ÕÆ≤˙Â"
Me.lblCaption.Caption = LoadResString(gLangOffSet + 242)   '"§ «œ˙Û∆æÚÆıÙ ∆˙ÙÛë¬ ÕÒ»¡ ∆˙˜¿ﬁ"
Me.lblDate.Caption = LoadResString(gLangOffSet + 37)    '"ä¬ÒÆ∞"
Me.lblDepAmount.Caption = LoadResString(gLangOffSet + 243)    '"«œ˙Û∆æÚ ∆˙˜¿ﬁ"
Me.lblRateofIntForLoans.Caption = LoadResString(gLangOffSet + 244)    '"ÕÒ»¡ ∆˙ÙÛë¬ ƒà‹ÆıÙ ¡«"
Me.lblLoanAmtAvail.Caption = LoadResString(gLangOffSet + 245)    '"ÕÒ»¡ ∆˙˜¿ﬁ"
Me.lblPrevLoanAmt.Caption = LoadResString(gLangOffSet + 246)    '"ñÆä¬ ÕÒ»¡ ∆˙˜¿ﬁ"
Me.lblLoanSanctioned.Caption = LoadResString(gLangOffSet + 247)    '"∆ÙÆ∏˜«Ò¡ ÕÒ»"
Me.lblLessIntOnPrevLoan.Caption = LoadResString(gLangOffSet + 248)    '"ñÆä¬ ÕÒ»¡ ƒà‹ "
Me.lblTotAmtIssued.Caption = LoadResString(gLangOffSet + 249)    '"´ªÙ⁄ ∆˙˜¿ﬁ ∞˙˜Ω≈˙Û∞ÒÇ¡Ù‡"
Me.lblcaption1.Caption = LoadResString(gLangOffSet + 250)    '"ñÆä¬ ÕÒ»¡ ƒà‹ÆıÙ¬Ù· ∞ ˙¡Ù Œ˙˜Õ ÕÒ»∆¬Ù· ∞˙˜Ω≈˙Û∞Ù"
Me.cmdUndo.Caption = LoadResString(gLangOffSet + 19)    '"°íÕÙ(∞˙˜¬˙ÆıÙ)"
Me.cmdRepay.Caption = LoadResString(gLangOffSet + 20)    '"∆Ù«Ù√Ò∆â"
Me.cmdAccept.Caption = LoadResString(gLangOffSet + 4)    '"°ÆÇÛ∞êÕÙ"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 11)    '"∆ÙÙ∂Ù’"
End Sub

Private Sub cmdAccept_Click()

Dim TransType As wisTransactionTypes

'Check out the date
    If Not DateValidate(txtDate.Text, "/", True) Then
        'MsgBox "Date of transaction not in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Sub
    End If

'See if account has already been matured or closed
    gDBTrans.SQLStmt = "Select * from PDMaster where AccID = " & m_AccID & _
                        " and UserId = " & m_UserID '& " order by TransID"
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Error accessing data base !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - Error"
        Exit Sub
    End If

    If WisDateDiff(FormatField(gDBTrans.Rst("MaturityDate")), txtDate.Text) >= 0 Then
        'MsgBox "You have specified a date that is later than the maturity date i.e " & FormatField(gDBTrans.Rst("MaturityDate")) & vbCrLf & "This means that you are trying to issue loans on a matured deposit !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 578) & FormatField(gDBTrans.Rst("MaturityDate")) & vbCrLf & "This means that you are trying to issue loans on a matured deposit !", vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Sub
    End If
    If FormatField(gDBTrans.Rst("ClosedDate")) <> "" Then
        'MsgBox "This deposit has already been closed !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 524), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

'Check date range w.r.t to loan
    gDBTrans.SQLStmt = "Select TOP 1 TransDate from PDTrans where AccID = " & m_AccID & _
                        " and UserId = " & m_UserID & " And Loan = True order by TransID desc"
    If gDBTrans.SQLFetch > 0 Then
        If WisDateDiff(FormatField(gDBTrans.Rst("TransDate")), txtDate.Text) < 0 Then
            'MsgBox "Date of transaction is lesser than the previous transaction date", vbExclamation, gAppName & " - Error"
            MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtDate
            Exit Sub
        End If
    End If
    
'Check date range w.r.t to Deposit
    gDBTrans.SQLStmt = "Select TOP 1 TransDate from PDTrans where AccID = " & m_AccID & _
                        " and UserId = " & m_UserID & " And Loan = False order by TransID desc"
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Error accessing data base !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - Error"
        Exit Sub
    End If
    If WisDateDiff(FormatField(gDBTrans.Rst("TransDate")), txtDate.Text) < 0 Then
        'MsgBox "Date of transaction is lesser than the previous transaction date", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
'''        Exit Sub
    End If


'Check out if the interest rate is valid
    If Val(txtInterest.Text) <= 0 Then
        'MsgBox "Interest rate has not been specified for this period." & vbCrLf & vbCrLf & "Please set the value of interest for this period in the properties of this account !", vbInformation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 579) & vbCrLf & vbCrLf & LoadResString(gLangOffSet + 659), vbInformation, gAppName & " - Error"
        Exit Sub
    End If

'Check out the sanctioned amount
    If Not CurrencyValidate(txtSanctioned.Text, False) Then
        'MsgBox "Invalid amount sanctioned !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtSanctioned
        Exit Sub
    End If
    
    If Val(txtSanctioned.Text) > Val(txtAvailable.Text) Then
        'MsgBox "Loan amount sanctioned is greater than the available loan amount !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 581), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtSanctioned
'''        Exit Sub
    End If
    
'Get New transaction ID
    Dim TransID As Long
    Dim Balance As Currency
    
    gDBTrans.SQLStmt = "Select TOP 1 TransID,Balance from PDTrans where AccID = " & _
            m_AccID & " and UserId = " & m_UserID & " And Loan = True order by TransID desc"
    If gDBTrans.SQLFetch <= 0 Then  'There has to be atleast one transaction
        TransID = 100
        Balance = Val(CCur(Me.txtSanctioned.Text))
    Else
        Balance = Val(FormatField(gDBTrans.Rst("Balance"))) + Val(CCur(Me.txtSanctioned.Text))
        TransID = FormatField(gDBTrans.Rst("TransID")) + 1
    End If
    
'Start data base transactions
gDBTrans.BeginTrans
    'First insert any interest of previous loans
    'If Val(txtInterestAmount.Text) > 0 Then
    'TransType = wCharges
    TransType = wContraWithdraw
        gDBTrans.SQLStmt = "Insert into PDTrans (AccID, UserId, TransID, TransType, " & _
                            " TransDate,Amount,Balance,Loan, " & _
                            " Particulars) values ( " & _
                            m_AccID & "," & _
                            m_UserID & "," & _
                            TransID & "," & _
                            TransType & "," & _
                            "#" & FormatDate(txtDate.Text) & "#," & _
                            Val(txtInterestAmount.Text) & "," & _
                            Balance & ", True, " & _
                            "'" & "By interest" & "')"
                            
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Sub
        End If
        TransID = TransID + 1
    'End If
    
    'Now insert record of the new loan
    TransType = wWithDraw
    gDBTrans.SQLStmt = "Insert into PDTrans (AccID, UserId, TransID, TransType, " & _
                        " TransDate,  Amount,Balance, Loan , " & " Particulars) values ( " & _
                        m_AccID & "," & _
                        m_UserID & "," & _
                        TransID & "," & _
                        TransType & "," & _
                        "#" & FormatDate(txtDate.Text) & "#," & _
                        Val(txtSanctioned.Text) & "," & _
                        Balance & ",True ," & _
                        "'" & "To Loans" & "')"
                        
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Sub
    End If

'COmmit transactions
gDBTrans.CommitTrans

'Udate date with todays date (By default)
    txtDate.Text = FormatDate(gStrDate)
    Me.txtSanctioned.Text = "0.00"
'Update the details on the UI
    Call UpdateUserInterface
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdRepay_Click()
'frmPDRepay.Show vbModal
Call UpdateUserInterface
End Sub

Private Sub cmdUndo_Click()

Dim TransID As Long
'Get the last transaction ID
    gDBTrans.SQLStmt = "Select Top 1 TransID, TransType from PDTrans where AccID = " & _
                        m_AccID & " and UserId = " & m_UserID & " And Loan = True " & _
                        " order by TransID desc"
    'Call gDBTrans.SQLFetch
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "You do not have any loans on this deposit !", vbInformation, gAppName & " - Message"
        MsgBox LoadResString(gLangOffSet + 582), vbInformation, gAppName & " - Message"
        Exit Sub
    End If
    TransID = FormatField(gDBTrans.Rst("TransID"))

'Check out the transaction before the last transaction, because it may the interest
'added. Since we are performing interest charges automatically, we've got to remove
'this also automatically
'May Not Necessary

gDBTrans.SQLStmt = "Select TransType from PDTrans where " & _
                    " AccID = " & m_AccID & _
                    " and UserId = " & m_UserID & _
                    " and Loan = True " & _
                    " and TransID = " & TransID - 1
Call gDBTrans.SQLFetch
Dim TransType As wisTransactionTypes
TransType = FormatField(gDBTrans.Rst("TransType"))

'COnfirm about transaction
'If MsgBox("Are you sure you want to undo a previous loan transaction ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
If MsgBox(LoadResString(gLangOffSet + 583), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
    Exit Sub
End If

gDBTrans.BeginTrans
    'Remove the last transaction
    gDBTrans.SQLStmt = "Delete from PDTrans where AccID = " & m_AccID & _
                        " and UserId = " & m_UserID & _
                        " and Loan = True " & _
                        " and TransID = " & TransID
    If Not gDBTrans.SQLExecute Then
        'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
        MsgBox LoadResString(gLangOffSet + 609), vbCritical, gAppName & " - Critical Error"
        gDBTrans.RollBack
        Exit Sub
    End If

    'Remove the transaction previous to the one removed only if it is of type
    'CHARGES levied. b'cause this record would have been added automatically
    'If TransType = wCharges Then
        gDBTrans.SQLStmt = "Delete from PDTrans where AccID = " & m_AccID & _
                            " and UserId = " & m_UserID & _
                            " And Loan = True " & _
                            " and TransID = " & TransID - 1
        If Not gDBTrans.SQLExecute Then
            'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
            MsgBox LoadResString(gLangOffSet + 609), vbCritical, gAppName & " - Critical Error"
            gDBTrans.RollBack
            Exit Sub
        End If
    'End If
gDBTrans.CommitTrans

'Udate date with todays date (By default)
    txtDate.Text = FormatDate(gStrDate)

Call UpdateUserInterface
End Sub
Private Sub Form_Load()

'Centre the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    txtDate.Text = FormatDate(gStrDate)
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'set kannada fonts
Call SetKannadaCaption
'Initialize the grid
    grd.Rows = 6
    grd.Cols = 4
    grd.FixedRows = 1
    grd.FixedCols = 0
    grd.Row = 0
    grd.Col = 0: grd.Text = LoadResString(gLangOffSet + 37): grd.ColWidth(0) = (grd.Width / 4) '- (grd.Width / 12) 'Some shit adjustment,"Date"
    grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 235): grd.ColWidth(1) = grd.Width / 4  '"Loan Amount"
    grd.Col = 2: grd.Text = LoadResString(gLangOffSet + 216): grd.ColWidth(2) = grd.Width / 4 '"Repayment"
    grd.Col = 3: grd.Text = LoadResString(gLangOffSet + 274): grd.ColWidth(3) = grd.Width / 4 '"Interest"
    
'Fill up the two module level variables
    m_AccID = frmPDAcc.m_AccID
    m_UserID = frmPDAcc.m_UserID
    
'Obtain the rate of interest as applicable to this deposit
    Dim Days As Long
    gDBTrans.SQLStmt = "Select * from PDMaster where " & _
                        " AccID = " & m_AccID & " and UserId = " & m_UserID

    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Error accessing data base !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
    'Check if deposit is closed
    If FormatField(gDBTrans.Rst("ClosedDate")) <> "" Then
        cmdUndo.Enabled = False: cmdAccept.Enabled = False: cmdRepay.Enabled = False
    Else
        cmdUndo.Enabled = True: cmdAccept.Enabled = True: cmdRepay.Enabled = True
    End If
    
    Days = WisDateDiff(FormatField(gDBTrans.Rst("CreateDate")), FormatField(gDBTrans.Rst("MaturityDate")))
    
'''Read The Interest Rate from Set Up
    txtInterest = GetPDLoanInterest(Days, FormatDate(gStrDate))
    If Val(txtInterest.Text) = 0 Then
        Dim Intclass As New clsInterest
        txtInterest.Text = Intclass.InterestRate(wis_PDAcc, "6_12_Loan", (gStrDate))
    End If
    txtInterest.Text = Format(txtInterest.Text, "#0.00")
    
Call UpdateUserInterface
End Sub

Private Sub UpdateUserInterface()

Dim TransType As wisTransactionTypes

'Get transaction details of this account and deposit
    
                
'Get the Maturity date & Close Date
    Dim MatDate As String
    Dim ClosedDate As String
    gDBTrans.SQLStmt = "Select * from PDMaster where AccID = " & _
                    m_AccID & " and UserId = " & _
                    m_UserID '& _
                    '" order by TransID"
    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    Else
        MatDate = FormatField(gDBTrans.Rst("MaturityDate"))
        ClosedDate = FormatField(gDBTrans.Rst("ClosedDate"))
    End If
    'Update command buttons only if deposit is not closed
    'If FormatField(gDBTrans.Rst("ClosedDate")) = "" Then
    If ClosedDate = "" Then
            cmdRepay.Enabled = True
            cmdUndo.Enabled = True
    Else
        cmdRepay.Enabled = False: cmdUndo.Enabled = False: cmdAccept.Enabled = False
    End If
    
    'Get The Total Depited Amount
    TransType = wisTransactionTypes.wDeposit
    gDBTrans.SQLStmt = "Select Sum(Amount) from PDTrans where AccID = " & _
                    m_AccID & " and UserId = " & m_UserID & _
                    " And Loan = False and TransType = " & TransType
                    
    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    Else
        txtDeposit.Text = FormatField(gDBTrans.Rst(0))
    End If
    
    gDBTrans.SQLStmt = "Select * from PDTrans where AccID = " & _
                    m_AccID & " and UserId = " & m_UserID & _
                    " And Loan = True order by TransID"
                    
    If gDBTrans.SQLFetch < 1 Then
    
    End If
    
    Dim LoanAmount As Currency
    Dim Transdate As String
    Dim i As Integer
    LoanAmount = 0
    grd.Rows = 1: grd.Rows = 7
    grd.Row = 0
    
    While Not gDBTrans.Rst.EOF
        
        'Register the TransType first
        TransType = FormatField(gDBTrans.Rst("TransTYpe"))
        
        'Check out if field is displayable
        If gDBTrans.Rst("Amount") = 0 Then
'            If TransType = wCharges Then
'                GoTo NextRecord
'            End If
        End If
        
        
        'Set new row number for displaying the record
        If grd.Rows = grd.Row + 2 Then
            grd.Rows = grd.Rows + 2
        End If
        grd.Row = grd.Row + 1
        
        If TransType = wWithDraw Then  'Loans Drawn
            grd.Col = 0: grd.Text = FormatField(gDBTrans.Rst("TransDate"))
            grd.Col = 1: grd.Text = FormatField(gDBTrans.Rst("Amount"))
            LoanAmount = LoanAmount + FormatField(gDBTrans.Rst("Amount"))
        ElseIf TransType = wDeposit Then
            grd.Col = 0: grd.Text = FormatField(gDBTrans.Rst("TransDate"))
            grd.Col = 2: grd.Text = FormatField(gDBTrans.Rst("Amount"))
             LoanAmount = LoanAmount - FormatField(gDBTrans.Rst("Amount"))
        'ElseIf TransType = wCharges Then        'Interest charged
            grd.Col = 0: grd.Text = FormatField(gDBTrans.Rst("TransDate"))
            grd.Col = 3: grd.Text = FormatField(gDBTrans.Rst("Amount"))
        End If
        
        Transdate = FormatField(gDBTrans.Rst("TransDate"))
NextRecord:
        gDBTrans.Rst.MoveNext
    Wend
    txtLoan.Text = FormatCurrency(LoanAmount)
    
'Calculate the available amount form loan ( 80 % of deposit - Total loans drawn till date)
                                            'is what is available. Take this from setup
    Dim LoanLimit As Single
    If M_setUp Is Nothing Then Set M_setUp = New clsSetup
    LoanLimit = Val(M_setUp.ReadSetupValue("PDAcc", "MaxLoanPercent", "80"))
    txtAvailable.Text = FormatCurrency((Val(txtDeposit.Text) * LoanLimit / 100) - Val(txtLoan.Text))
    

'Calculate the interest for loan if a loan has been drawn previously
    Dim Days As Integer
    'Dim Days1 As Integer, Days2 As Integer
    On Error Resume Next
    
    'See if deposit has matured
    If LoanAmount > 0 Then
        If WisDateDiff(txtDate.Text, MatDate) <= 0 Then
            Days = WisDateDiff(Transdate, MatDate)
        Else
            Days = WisDateDiff(Transdate, txtDate.Text)
        End If
        If Days >= 0 Then
            txtInterestAmount.Text = FormatCurrency(ComputePDLoanInterest(LoanAmount, Days, Val(txtInterest.Text)))
        End If
    Else
        txtInterestAmount.Text = "0.00"
    End If
    lblCaption.Caption = LoadResString(gLangOffSet + 242) & " " & LoadResString(gLangOffSet + 242) & " " & LoanAmount


End Sub

Public Function ComputePDLoanInterest(Principle As Currency, NumOfDays As Integer, RateOfInterest As Double) As Currency
    ComputePDLoanInterest = Principle * (NumOfDays / 365) * (RateOfInterest / 100)
End Function

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
Set frmPDLoans = Nothing
End Sub



Private Sub txtDate_LostFocus()
If Not DateValidate(txtDate.Text, "/", True) Then
    Exit Sub
End If
Call UpdateUserInterface

End Sub


Private Sub txtInterestAmount_Change()
'COmpute the total amount to be actully issued
txtIssuedAmount.Text = FormatCurrency(Val(txtSanctioned.Text) - Val(txtInterestAmount.Text))

End Sub

Private Sub txtSanctioned_Change()
'COmpute the total amount to be actully issued
txtIssuedAmount.Text = FormatCurrency(Val(txtSanctioned.Text) - Val(txtInterestAmount.Text))
End Sub


