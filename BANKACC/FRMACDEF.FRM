VERSION 5.00
Begin VB.Form frmAcDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Accounts..."
   ClientHeight    =   2115
   ClientLeft      =   2130
   ClientTop       =   3525
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3870
         TabIndex        =   4
         Top             =   1245
         Width           =   1200
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "C&reate"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   3870
         TabIndex        =   2
         Top             =   285
         Width           =   1200
      End
      Begin VB.TextBox txtAccName 
         Height          =   345
         Left            =   240
         TabIndex        =   1
         Top             =   1185
         Width           =   3225
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3870
         TabIndex        =   3
         Top             =   765
         Width           =   1200
      End
      Begin VB.TextBox txtAccHead 
         Height          =   345
         Left            =   255
         TabIndex        =   5
         Top             =   525
         Width           =   3225
      End
      Begin VB.Label lblAccName 
         AutoSize        =   -1  'True
         Caption         =   "Specify the account name :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   930
         Width           =   1950
      End
      Begin VB.Label lblAccHead 
         AutoSize        =   -1  'True
         Caption         =   "Account head category :"
         Height          =   195
         Left            =   255
         TabIndex        =   6
         Top             =   270
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmAcDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public p_HeadId  As Long
Private Sub SetKannadaCaption()
Dim Ctrl As Control
On Error Resume Next
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next
Me.lblAccHead.Caption = LoadResString(gLangOffSet + 112)
Me.lblAccName.Caption = LoadResString(gLangOffSet + 113)
cmdCreate.Caption = LoadResString(gLangOffSet + 15) '"Çð„Íðô"
Me.cmdDelete.Caption = LoadResString(gLangOffSet + 14)  '"¡’Íðô"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 11)  '"ÇðÁðôà"


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()

'Check For The Duplicate entry

gDBTrans.SQLStmt = "Select * From AccMAster Where AccName = '" & Me.txtAccName & "'" & _
        " And accid > " & p_HeadId & " And AccId < " & p_HeadId + wis_BankHeadOffSet
If gDBTrans.SQLFetch > 0 Then
    'MsgBox "This Head already exist" & vbCrLf & " Can't Create duplicate Heads"
    MsgBox LoadResString(gLangOffSet + 671), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Dim AccId As Long

'Get the next HeadId For new Account
'get the HeadId First

gDBTrans.SQLStmt = "Select max(AccId) From AccMaster Where AccID > " & p_HeadId & _
                                " And AccId < " & p_HeadId + wis_BankHeadOffSet
AccId = 0
If gDBTrans.SQLFetch > 0 Then
    AccId = Val(FormatField(gDBTrans.Rst(0))) + 1
End If
If AccId < wis_BankHeadOffSet Then
    AccId = p_HeadId + 1
End If

    gDBTrans.SQLStmt = " Insert Into AccMaster (AccId,AccName,CreateDate) " & _
                    " Values(" & AccId & "," & _
                    " '" & txtAccName.Text & "'," & _
                    " #" & FormatDate(frmBankAcc.txtDate) & "# )"
    gDBTrans.BeginTrans
    If Not gDBTrans.SQLExecute Then
        'MsgBox " Unable to add new Head ", vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 672), vbInformation, wis_MESSAGE_TITLE
        gDBTrans.RollBack
    End If
    gDBTrans.CommitTrans
    txtAccName.Text = ""
    
End Sub


Private Sub cmdDelete_Click()
'Check Whethe Head exists are not
Dim AccId As Integer
gDBTrans.SQLStmt = "Select AccId From AccMaster Where AccName = '" & _
                txtAccName & "' And Accid > " & p_HeadId & " And Accid < " & p_HeadId
If gDBTrans.SQLFetch < 1 Then
    'MsgBox "Account Type is not existing"
    MsgBox LoadResString(gLangOffSet + 673), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
Else
    AccId = Val(gDBTrans.Rst(0))
End If
gDBTrans.SQLStmt = "Select * From AccTrans Where AccId = " & AccId
If gDBTrans.SQLFetch > 1 Then
    'MsgBox "Account Type is Having Transacion can not be deleted"
    MsgBox LoadResString(gLangOffSet + 674)
    Exit Sub
End If
'Confirmation Mesage
'If MsgBox("Are you sure you want to delete Account ?", vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
If MsgBox(LoadResString(gLangOffSet + 675), vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
    Exit Sub
End If

gDBTrans.SQLStmt = "Delete * From AccMaster Where AccId = " & AccId
gDBTrans.BeginTrans
If Not gDBTrans.SQLExecute Then
    'MsgBox "Unable to delete Account", vbCritical, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 676), vbCritical, wis_MESSAGE_TITLE
    gDBTrans.RollBack
    Exit Sub
End If
gDBTrans.CommitTrans
txtAccName.Text = ""
End Sub


Private Sub Form_Load()
If p_HeadId < wis_BankHeadOffSet Then
    Err.Raise 5001, , " Invalid Head Id assigned"
    cmdCreate.Enabled = False
    cmdDelete.Enabled = False
    Unload Me
    Exit Sub
End If

Call SetKannadaCaption

End Sub

Private Sub txtAccName_Change()
    If Trim$(txtAccName.Text) <> "" Then
        cmdCreate.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdCreate.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub


