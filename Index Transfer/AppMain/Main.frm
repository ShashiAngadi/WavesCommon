VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Data Transfer of Index2000 v2 TO v3"
   ClientHeight    =   7050
   ClientLeft      =   1800
   ClientTop       =   1785
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7230
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6705
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   6795
      Begin ComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   6240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CheckBox chkOnlyAccHeads 
         Caption         =   "To Transfer only  ledger heads check this  (no account details will  be transferred)"
         Height          =   225
         Left            =   240
         TabIndex        =   21
         Top             =   1830
         Width           =   6255
      End
      Begin VB.CommandButton cdmKannnada 
         Caption         =   "."
         Height          =   285
         Left            =   6570
         TabIndex        =   20
         Top             =   1470
         Width           =   195
      End
      Begin VB.CommandButton cmdCompact 
         Caption         =   "Compact New Data base"
         Height          =   525
         Left            =   150
         TabIndex        =   15
         Top             =   5220
         Width           =   3165
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "BankAccounts"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3480
         TabIndex        =   14
         Top             =   4590
         Width           =   3105
      End
      Begin VB.TextBox txtDate 
         Height          =   345
         Left            =   5340
         TabIndex        =   18
         Text            =   "31/3/YYYY"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdRepair 
         Caption         =   "Repair the date base"
         Height          =   525
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   3195
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   525
         Left            =   3480
         TabIndex        =   16
         Top             =   5250
         Width           =   3105
      End
      Begin VB.CommandButton cmdBkcc 
         Caption         =   "Transfer BKCC Loans"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3450
         TabIndex        =   13
         Top             =   3990
         Width           =   3165
      End
      Begin VB.CommandButton CMDLoan 
         Caption         =   "Transfer Loans"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3450
         TabIndex        =   12
         Top             =   3390
         Width           =   3165
      End
      Begin VB.CommandButton cmdMem 
         Caption         =   "Transfer Member Details"
         Enabled         =   0   'False
         Height          =   525
         Left            =   150
         TabIndex        =   5
         Top             =   2190
         Width           =   3165
      End
      Begin VB.CommandButton cmdOldDb 
         Caption         =   "Select the Old Database"
         Height          =   525
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   3165
      End
      Begin VB.CommandButton cmdNewDb 
         Caption         =   "Create New Data base"
         Height          =   525
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   3165
      End
      Begin VB.CommandButton cmdCA 
         Caption         =   "Transfer Current Account"
         Enabled         =   0   'False
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   3390
         Width           =   3165
      End
      Begin VB.CommandButton cmdSb 
         Caption         =   "Transfer Saving Account"
         Enabled         =   0   'False
         Height          =   525
         Left            =   150
         TabIndex        =   6
         Top             =   2790
         Width           =   3165
      End
      Begin VB.CommandButton cmdPD 
         Caption         =   "Transfer Pigmy Deposit"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3450
         TabIndex        =   11
         Top             =   2790
         Width           =   3165
      End
      Begin VB.CommandButton cmdRD 
         Caption         =   "Transfer Recurring Deposit"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3450
         TabIndex        =   10
         Top             =   2190
         Width           =   3165
      End
      Begin VB.CommandButton cmdFD 
         Caption         =   "Transfer Fixed Deposit"
         Enabled         =   0   'False
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   3990
         Width           =   3165
      End
      Begin VB.CommandButton cmdDL 
         Caption         =   "Transfer Cash Certificate (DL)"
         Enabled         =   0   'False
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   4590
         Width           =   3165
      End
      Begin VB.CommandButton cmdName 
         Caption         =   "Name Transfer"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3450
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the date from where the data has to be transferred"
         Height          =   315
         Left            =   300
         TabIndex        =   19
         Top             =   1500
         Width           =   4500
      End
      Begin VB.Label lblProgress 
         Caption         =   "progreass information"
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   5850
         Width           =   6270
      End
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   6750
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_OldDbName As String
Private m_NewDBName As String

Private m_Shift As Boolean
Private m_Ctrl As Boolean
Sub CheckControls()

If m_OldDbName = "" Then GoTo Exit_line
If m_NewDBName = "" Then GoTo Exit_line
If Not DateValidate(txtDate, "/", True) Then GoTo Exit_line

gFromDate = FormatDate(txtDate)

Dim NewVal As Boolean
NewVal = True
If Dir(m_OldDbName) = "" Then NewVal = False 'GoTo Exit_line
If Dir(m_NewDBName) = "" Then NewVal = False 'GoTo Exit_line


cmdName.Enabled = NewVal
cmdSb.Enabled = NewVal
cmdCA.Enabled = NewVal
CMDLoan.Enabled = NewVal
cmdMem.Enabled = NewVal
cmdBkcc.Enabled = NewVal
cmdRD.Enabled = NewVal
cmdFD.Enabled = NewVal
cmdPD.Enabled = NewVal
cmdDL.Enabled = NewVal
cmdRepair.Enabled = NewVal
cmdBank.Enabled = NewVal


Exit_line:

End Sub

Private Sub cdmKannnada_Click()
If m_Shift Then _
    m_Shift = False: OldPwd = InputBox("Enter the PassWord for Old Data Base", "OLD DB PassWord", "PRAGMANS")

If m_Ctrl Then _
    m_Ctrl = False: NewPwd = InputBox("Enter the PassWord for New DataBase", "New Database PassWord", "WIS!@#")


gLangOffSet = Val(InputBox("Enter the Language offset", "Select Language", gLangOffSet))

End Sub

Private Sub chkOnlyAccHeads_Click()
If chkOnlyAccHeads.Value = vbChecked Then _
    gOnlyLedgerHeads = True Else gOnlyLedgerHeads = False
End Sub

Private Sub cmdBank_Click()
cmdBank.Enabled = False

If TransferBank(m_OldDbName, m_NewDBName) Then
    MsgBox "Bank Account details transferred", vbInformation, wis_MESSAGE_TITLE
Else
    MsgBox "Unable to transfer Bank Account details", vbInformation, wis_MESSAGE_TITLE
End If

End Sub

Private Sub cmdBkcc_Click()
cmdBkcc.Enabled = False
If TransferBKCC(m_OldDbName, m_NewDBName) Then
    MsgBox "BKCC Details Transferred"
Else
    MsgBox "Unable to transfer the BKCC details"
End If


End Sub

Private Sub cmdCA_Click()

cmdCA.Enabled = False
If TransferCA(m_OldDbName, m_NewDBName) Then
    MsgBox "CA Details Transferred"
Else
    MsgBox "unable to transfer the CA details"
End If


End Sub

Private Sub cmdClose_Click()
Unload Me
End
End Sub

Private Sub cmdCompact_Click()

Dim DbCompTrans As clsDBUtils

Set DbCompTrans = New clsDBUtils

If Not DbCompTrans.OpenDB(m_NewDBName, NewPwd) Then
    MsgBox "Unable to compact the data base"
    Exit Sub
End If

If Not DbCompTrans.WISCompactDB(m_NewDBName, NewPwd, NewPwd) Then
    MsgBox "Unable to compact the data base"
    Exit Sub
End If

Call DbCompTrans.CloseDB

Set DbCompTrans = Nothing

End Sub

Private Sub cmdDL_Click()
cmdDL.Enabled = False
If TransferDL(m_OldDbName, m_NewDBName) Then
    MsgBox "DL Details Transferred"
Else
    MsgBox "Unable to transfer the DL details"
End If


End Sub

Private Sub cmdFD_Click()
cmdFD.Enabled = False
If TransferFD(m_OldDbName, m_NewDBName) Then
    MsgBox "FD Details Transferred"
Else
    MsgBox "Unable to transfer the FD details"
End If


End Sub

Private Sub CMDLoan_Click()
CMDLoan.Enabled = False
If TransferLoan(m_OldDbName, m_NewDBName) Then
    MsgBox "Loan Details Transferred"
Else
    MsgBox "unable to transfer the Loan details"
End If
End Sub

Private Sub cmdMem_Click()
cmdMem.Enabled = False
If MemberTransfer(m_OldDbName, m_NewDBName) Then
    MsgBox "Member detials Tranferred"
Else
    MsgBox "Unable to transfer the member details"
End If
End Sub


Private Sub cmdName_Click()
If Not TransferNameTab(m_OldDbName, m_NewDBName) Then
    MsgBox "Unable to Transfer the Name Details"
Else
    MsgBox "Name details transferred"
End If
End Sub


Private Sub cmdNewDb_Click()
With cdb
    .CancelError = False
    .FileName = ""
    .InitDir = "C:\Program Files\Index 2000"
    .Filter = "Data Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
    .DialogTitle = "Create/Select the new database"
    .ShowSave
    m_NewDBName = .FileName
    Me.Refresh
End With

Dim DbClass As New clsDBUtils

If Dir(m_NewDBName) = "" Then
    Me.Refresh
    lblProgress = "Creating Database"
    prg.Value = 0
    Me.Refresh
    Screen.MousePointer = vbHourglass
    'Check whethere Database is alread exists.
        If Not DbClass.CreateDB(App.Path & "\Indx2000.tab", NewPwd, m_NewDBName) Then
            Screen.MousePointer = vbDefault
            MsgBox "Unable to Create the Database", vbInformation, wis_MESSAGE_TITLE
            Exit Sub
        End If
    Screen.MousePointer = vbDefault
End If
Set DbClass = Nothing

Call CheckControls
End Sub


Private Sub cmdOldDb_Click()
With cdb
    .CancelError = False
    .FileName = ""
    .InitDir = "C:\Program Files\Index 2000"
    .Filter = "Data Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
    .DialogTitle = "Select the Old Data Base"
    .ShowOpen
    m_OldDbName = .FileName
    
End With
  Call CheckControls
End Sub

Private Sub cmdPD_Click()
cmdPD.Enabled = False
If TransferPD(m_OldDbName, m_NewDBName) Then
    MsgBox "Pigmy Details Transferred"
Else
    MsgBox "unable to transfer the Pigmy"
End If


End Sub

Private Sub cmdRD_Click()
cmdRD.Enabled = False
If TransferRD(m_OldDbName, m_NewDBName) Then
    MsgBox "RD Details Transferred"
Else
    MsgBox "Unable to transfer the RD details"
End If


End Sub

Private Sub cmdRepair_Click()

    Call RepairOldDB(m_OldDbName)

End Sub

Private Sub cmdSb_Click()
cmdSb.Enabled = False
If TransferSB(m_OldDbName, m_NewDBName) Then
    MsgBox "Sb Details Transferred"
Else
    MsgBox "unable to transfer the Sb details"
End If

End Sub


Private Sub cndClose_Click()

End Sub


Private Sub Command1_Click()

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift And vbShiftMask Then m_Shift = True
If Shift And vbCtrlMask Then m_Ctrl = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift And vbShiftMask Then m_Shift = False
If Shift And vbCtrlMask Then m_Ctrl = False

End Sub


Private Sub Form_Load()
NewPwd = "WIS!@#"
OldPwd = "PRAGMANS"

End Sub

Private Sub txtDate_LostFocus()
Call CheckControls
End Sub

