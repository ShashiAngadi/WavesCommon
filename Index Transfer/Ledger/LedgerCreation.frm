VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Ledger"
   ClientHeight    =   5340
   ClientLeft      =   2940
   ClientTop       =   2025
   ClientWidth     =   5385
   Icon            =   "LedgerCreation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5385
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtOpBalance 
      Height          =   395
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtLedgerName 
      Height          =   395
      Left            =   1800
      TabIndex        =   4
      Top             =   810
      Width           =   3285
   End
   Begin ComctlLib.ListView lvwLedger 
      Height          =   2745
      Left            =   60
      TabIndex        =   8
      Top             =   1830
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4842
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox cmbParent 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   3285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4080
      TabIndex        =   1
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   2910
      TabIndex        =   0
      Top             =   4920
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5295
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Label lblOpBalance 
      Caption         =   "Opening Balance"
      Height          =   390
      Left            =   60
      TabIndex        =   6
      Top             =   1350
      Width           =   1695
   End
   Begin VB.Label lblLedger 
      Caption         =   "Ledger Name"
      Height          =   390
      Left            =   60
      TabIndex        =   5
      Top             =   810
      Width           =   1695
   End
   Begin VB.Label lblParentName 
      Caption         =   "Select Parent Ledger"
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   1695
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_HeadID As Long

' are used in the class

Public Event OKClick()
Public Event CancelClick()
Public Event LookupClick(ParentID As Long)
Public Event LvwLedgerClick(HeadID As Long)

Private m_DBOperation As wis_DBOperation
'set the Kannada option here.
Private Sub SetKannadaCaption()

Dim ctrl As VB.Control
On Error Resume Next
For Each ctrl In Me
 
 ctrl.Font.Name = gFontName
 If Not TypeOf ctrl Is ComboBox Then
    ctrl.Font.Size = gFontSize
 End If
Next

'set the Kannada for all controls
lblParentName.Caption = LoadResString(gLangOffSet + 160) & " " & LoadResString(gLangOffSet + 36)
lblLedger.Caption = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 35)
lblOpBalance.Caption = LoadResString(gLangOffSet + 284)
cmdOk.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 11)



End Sub



Private Sub ClearControls()

cmbParent.ListIndex = -1
txtLedgerName.Text = ""
txtOpBalance.Text = ""
lvwLedger.ColumnHeaders.Clear
cmdOk.Caption = LoadResString(gLangOffSet + 10)

End Sub






Private Function Validated() As Boolean

Validated = False

If cmbParent.ListIndex = -1 Then
    MsgBox "Select Parent Name ", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

If Not CurrencyValidate(txtOpBalance.Text, True) Then
    MsgBox "Invalid opening balance specified", vbInformation, wis_MESSAGE_TITLE
    txtOpBalance.SetFocus
    Exit Function
End If

If Len(txtLedgerName.Text) = 0 Then
    MsgBox "No LedgerName specified", vbInformation, wis_MESSAGE_TITLE
    txtLedgerName.SetFocus
    Exit Function
End If

If Val(txtLedgerName.Text) > 0 Then
    MsgBox "Invalid LedgerName specified", vbInformation, wis_MESSAGE_TITLE
    txtLedgerName.SetFocus
    Exit Function
End If

Validated = True

End Function


Private Sub cmbParent_Click()

Dim ParentID As Long
Dim rstHead As Recordset

If cmbParent.ListIndex = -1 Then Exit Sub

ParentID = cmbParent.ItemData(cmbParent.ListIndex)

Me.lblLedger.Refresh

RaiseEvent LookupClick(ParentID)


End Sub


Private Sub cmdCancel_Click()
Unload Me
RaiseEvent CancelClick

End Sub


Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

RaiseEvent OKClick

m_DBOperation = Insert

End Sub

Private Sub Form_Load()

'Center the form
CenterMe Me

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)


If gLangOffSet <> 0 Then SetKannadaCaption

Call LoadParentHeads(cmbParent)

m_DBOperation = Insert


End Sub

Private Sub LoadParentHeads(ctrlComboBox As ComboBox)

Dim rstParent  As ADODB.Recordset

ctrlComboBox.Clear

gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads WHERE UserCreated <= 2 " & _
                   " ORDER BY ParentName "

Call gDbTrans.Fetch(rstParent, adOpenForwardOnly)

Do While Not rstParent.EOF

    ctrlComboBox.AddItem FormatField(rstParent.Fields("ParentName"))
    ctrlComboBox.ItemData(ctrlComboBox.NewIndex) = FormatField(rstParent.Fields("ParentID"))
    
    'Move to the next record
    rstParent.MoveNext
    
Loop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLedger = Nothing
End Sub

Private Sub lvwLedger_DblClick()
On Error Resume Next

Dim Count As Integer

' selected item key will be like this - "A2001"
' so we have to fetch only value

With lvwLedger.SelectedItem
    m_HeadID = Val(Mid(.Key, 2))
    txtLedgerName.Text = .Text
    txtOpBalance.Text = .SubItems(1)
    cmbParent.Locked = True
End With

m_DBOperation = Update
cmdOk.Caption = LoadResString(gLangOffSet + 171)

RaiseEvent LvwLedgerClick(m_HeadID)


End Sub



Private Sub txtLedgerName_LostFocus()
'txtLedgerName = ConvertToProperCase(txtLedgerName.Text)
End Sub


