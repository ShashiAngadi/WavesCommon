VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLedger1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Ledger1"
   ClientHeight    =   4875
   ClientLeft      =   450
   ClientTop       =   2235
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11130
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox lstLedger 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   5760
      TabIndex        =   9
      Top             =   570
      Width           =   5115
   End
   Begin VB.TextBox txtOpBalance 
      Height          =   345
      Left            =   1830
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtLedgerName 
      Height          =   345
      Left            =   1830
      TabIndex        =   4
      Top             =   570
      Width           =   3285
   End
   Begin ComctlLib.ListView lvwLedger 
      Height          =   2745
      Left            =   30
      TabIndex        =   8
      Top             =   1530
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
         Name            =   "Nudi B-Akshar"
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
      Left            =   1830
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   60
      Width           =   3285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   2910
      TabIndex        =   0
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label lblOpBalance 
      Caption         =   "Opening Balance"
      Height          =   345
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLedger 
      Caption         =   "Ledger Name"
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblParentName 
      Caption         =   "Select Parent Ledger"
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1695
   End
End
Attribute VB_Name = "frmLedger1"
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

Const LB_FINDSTRING = &H18F

Private Declare Function SendMessage Lib "User32" _
       Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long

'set the Kannada option here.
Private Sub SetKannadaCaption()

Dim Ctrl As VB.Control
On Error Resume Next
For Each Ctrl In Me
' If Not TypeOf Ctrl Is ComboBox Then
    Ctrl.FontName = gFontName
 '   If Not TypeOf Ctrl Is ListView Then
       
       Ctrl.FontSize = gFontSize
  '  End If
 'End If
Next

'set the Kannada for all controls
lblParentName.Caption = LoadResString(gLangOffSet + 128) & " " & LoadResString(gLangOffSet + 101) & " " & LoadResString(gLangOffSet + 19)
lblLedger.Caption = LoadResString(gLangOffSet + 101) & " " & LoadResString(gLangOffSet + 102)
lblOpBalance.Caption = LoadResString(gLangOffSet + 144)
cmdOK.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)



End Sub



Private Sub ClearControls()

cmbParent.ListIndex = -1
txtLedgerName.Text = ""
txtOpBalance.Text = ""
lvwLedger.ColumnHeaders.Clear
cmdOK.Caption = "OK"

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

Call LoadHeadsToListBox(ParentID, FinFromDate)

End Sub


Private Sub LoadHeadsToListBox(ByVal ParentID As Long, ByVal AsOnDate As String)

Dim rstHeads As ADODB.Recordset

' Check the Form's Status

If ParentID = 0 Then Exit Sub
If Not DateValidate(AsOnDate, "/", True) Then Exit Sub

lstLedger.Clear

gDbTrans.SqlStmt = " SELECT a.HeadID,HeadName,OpAmount " & _
                   " FROM Heads a,OpBalance b " & _
                   " WHERE a.ParentID =  " & ParentID & _
                   " AND a.HeadID=b.HeadID" & _
                   " AND b.OpDate=" & "#" & FormatDate(AsOnDate) & "#" & _
                   " ORDER BY HeadName"
                   
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) < 0 Then Exit Sub

With lstLedger
    
    Do While Not rstHeads.EOF
        .AddItem rstHeads.Fields("HeadName")
        .ItemData(.NewIndex) = rstHeads.Fields("HeadID")
        rstHeads.MoveNext
    Loop
    
End With

Set rstHeads = Nothing

End Sub

Private Sub cmdCancel_Click()

Unload Me
RaiseEvent CancelClick

End Sub


Private Sub cmdOK_Click()

If Not Validated Then Exit Sub

RaiseEvent OKClick

m_DBOperation = Insert

End Sub

Private Sub Form_Load()
'Center the Form

CenterMe Me

If gLangOffSet <> 0 Then SetKannadaCaption

Call LoadParentHeads(cmbParent)

m_DBOperation = Insert

End Sub

Private Sub LoadParentHeads(ctrlComboBox As ComboBox)

Dim rstParent  As ADODB.Recordset

ctrlComboBox.Clear

gDbTrans.SqlStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
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

' Selected Item key will be like this - "A2001"
' so we have to fetch only value

With lvwLedger.SelectedItem
    m_HeadID = Val(Mid(.Key, 2))
    txtLedgerName.Text = .Text
    txtOpBalance.Text = .SubItems(1)
    cmbParent.Locked = True
End With

m_DBOperation = Update
cmdOK.Caption = "&Update"

RaiseEvent LvwLedgerClick(m_HeadID)


End Sub



Private Sub txtLedgerName_Change()

Dim PrevText As String

PrevText = txtLedgerName.Text

If PrevText = "" Then
    lstLedger.ListIndex = -1
    Exit Sub
End If

With lstLedger
    
    .ListIndex = SendMessage(.hWnd, LB_FINDSTRING, -1, _
       ByVal PrevText)
          
    PrevText = Left$(PrevText, Len(PrevText) - 1)
    
    If .ListIndex = -1 Then .ListIndex = SendMessage(.hWnd, LB_FINDSTRING, -1, _
       ByVal PrevText)
       
End With

End Sub



