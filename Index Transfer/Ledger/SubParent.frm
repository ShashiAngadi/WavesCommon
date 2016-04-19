VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSubParent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Sub Parent"
   ClientHeight    =   6390
   ClientLeft      =   2955
   ClientTop       =   1395
   ClientWidth     =   5415
   Icon            =   "SubParent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   5415
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtLedgerName 
      Height          =   390
      Left            =   1860
      TabIndex        =   4
      Top             =   630
      Width           =   3285
   End
   Begin ComctlLib.ListView lvwLedger 
      Height          =   4305
      Left            =   120
      TabIndex        =   6
      Top             =   1170
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7594
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1860
      TabIndex        =   3
      Top             =   120
      Width           =   3285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4170
      TabIndex        =   1
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   5310
      Y1              =   5640
      Y2              =   5640
   End
   Begin ComctlLib.ImageList img1 
      Left            =   570
      Top             =   5910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubParent.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubParent.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLedger 
      Caption         =   "Ledger Name"
      Height          =   390
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   1695
   End
   Begin VB.Label lblParentName 
      Caption         =   "Select Parent Ledger"
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   1695
   End
End
Attribute VB_Name = "frmSubParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_SubParentID As Long

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
    If TypeName(ctrl) <> "ListView" Then
        ctrl.Font.Name = gFontName
        If Not TypeOf ctrl Is ComboBox Then
            ctrl.FontSize = gFontSize
        End If
    End If
Next

'set the Kannada for all controls
lblParentName.Caption = LoadResString(gLangOffSet + 160) & " " & LoadResString(gLangOffSet + 36)
lblLedger.Caption = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 35)
cmdOk.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)

lvwLedger.Font.Name = gFontName

End Sub


'
Private Sub ClearControls()
cmbParent.ListIndex = -1
txtLedgerName.Text = ""
lvwLedger.ColumnHeaders.Clear
cmdOk.Caption = "OK"
cmbParent.Locked = False
End Sub


'
Private Function Validated() As Boolean

Validated = False
If cmbParent.ListIndex = -1 Then
    MsgBox "Select Parent Name ", vbInformation, wis_MESSAGE_TITLE
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
'
Private Sub cmbParent_Click()

Dim ParentID As Long
Dim rstHead As Recordset

If cmbParent.ListIndex = -1 Then Exit Sub

ParentID = cmbParent.ItemData(cmbParent.ListIndex)

Me.lblLedger.Refresh

LoadSubHeadsToListView (ParentID)


End Sub


'
Private Sub LoadSubHeadsToListView(ParentID As Long)

Dim rstSubHeads As ADODB.Recordset
Dim MaxParentID As Long
Dim FillViewClass As clsFillView

Me.lvwLedger.ListItems.Clear

If ParentID = 0 Then Exit Sub

MaxParentID = ParentID + HEAD_OFFSET

gDbTrans.SQLStmt = " SELECT ParentID,ParentName,IIF(AccountType=2," & "'Liability'," & _
                        " IIF(Accounttype=1," & "'Asset'," & _
                        " IIF(Accounttype=4," & "'Expenses'," & _
                        " IIF(Accounttype=8," & "'Income'," & _
                        " IIF(Accounttype=16," & "'Sales'," & _
                        " 'Purchases'" & " )))))AS VoucherName" & _
                   " FROM ParentHeads" & _
                   " WHERE ParentID > " & ParentID & _
                   " AND ParentID < " & MaxParentID & _
                   " ORDER BY ParentName"
                                   
If gDbTrans.Fetch(rstSubHeads, adOpenForwardOnly) < 0 Then Exit Sub

Set FillViewClass = New clsFillView

If Not FillViewClass.FillViewWithSlno(Me.lvwLedger, rstSubHeads, "ParentID", False) Then Exit Sub

Set FillViewClass = Nothing

End Sub

'
Private Sub cmdCancel_Click()
Unload Me
RaiseEvent CancelClick

End Sub


'
Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

If m_DBOperation = Insert Then SaveSubParentHead
If m_DBOperation = Update Then UpdateSubParentHead


m_DBOperation = Insert

End Sub

'
Private Function SaveSubParentHead() As wis_FunctionReturned

On Error GoTo NoSaveError:

Dim ParentID As Long
Dim SubParentID As Long
Dim MaxParentID As Long

Dim AccountType As wis_AccountType

Dim rstHeads As ADODB.Recordset

SaveSubParentHead = Failure
' check the form's status
With Me
    If .cmbParent.ListIndex = -1 Then Exit Function
    ParentID = .cmbParent.ItemData(.cmbParent.ListIndex)
    MaxParentID = ParentID + HEAD_OFFSET
    
    'Get the Maximum Head From the database
    gDbTrans.SQLStmt = " SELECT MAX(ParentID) as MaxParentID,AccountType FROM ParentHeads " & _
                       " WHERE ParentID > " & ParentID & _
                       " AND ParentID < " & MaxParentID & _
                       " GROUP BY AccountType"
    
    Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)
    SubParentID = FormatField(rstHeads.Fields("MaxParentID")) + 100
    AccountType = FormatField(rstHeads.Fields("AccountType"))
    
    If SubParentID < ParentID Then SubParentID = SubParentID + ParentID
    'Insert the heads inot the database
    gDbTrans.SQLStmt = " INSERT INTO ParentHeads (ParentID,ParentName,AccountType,UserCreated) " & _
                      " VALUES ( " & _
                      SubParentID & "," & _
                      AddQuotes(.txtLedgerName.Text, True) & "," & _
                      AccountType & "," & _
                      2 & ")" ' this is user creared so value is 2
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    gDbTrans.CommitTrans
End With
MsgBox " New Sub ParentHead is Saved !!"

LoadSubHeadsToListView (ParentID)

SaveSubParentHead = Success

'Clear the controls
ClearControls

Exit Function

NoSaveError:
    ' Clear up the  Transactions if any
    If gDbTrans.BeginTrans Then gDbTrans.RollBack
        SaveSubParentHead = FatalError
End Function

'
Private Function UpdateSubParentHead() As wis_FunctionReturned

On Error GoTo NoSaveError:

Dim SubParentID As Long

UpdateSubParentHead = Failure

' check the form's status

With Me

    If .cmbParent.ListIndex = -1 Then Exit Function
    
    SubParentID = m_SubParentID
    
    'Insert the heads inot the database
    gDbTrans.SQLStmt = " UPDATE ParentHeads " & _
                       " SET ParentName=" & AddQuotes(.txtLedgerName.Text, True) & _
                       " WHERE ParentID=" & SubParentID
    
    
    gDbTrans.BeginTrans
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
        
    
    gDbTrans.CommitTrans
    
End With


MsgBox "Sub ParentHead is Updated !!"

UpdateSubParentHead = Success

'Clear the controls
ClearControls

Exit Function

NoSaveError:
        
    'Clear up the  Transactions if any
    If gDbTrans.BeginTrans Then gDbTrans.RollBack
    
    UpdateSubParentHead = FatalError

End Function

Private Sub Form_Load()

'Center the form
CenterMe Me

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

If gLangOffSet <> 0 Then SetKannadaCaption

Call LoadParentHeads(cmbParent)

Set lvwLedger.SmallIcons = img1

m_DBOperation = Insert

End Sub

'
Private Sub LoadParentHeads(ctrlComboBox As ComboBox)
Dim rstParent  As ADODB.Recordset

ctrlComboBox.Clear

gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
                   " WHERE ParentID mod " & HEAD_OFFSET & "=0" & _
                   " ORDER BY ParentName "

Call gDbTrans.Fetch(rstParent, adOpenForwardOnly)

Do While Not rstParent.EOF
    ctrlComboBox.AddItem FormatField(rstParent("ParentName"))
    ctrlComboBox.ItemData(ctrlComboBox.NewIndex) = FormatField(rstParent("ParentID"))
    'Move to the next record
    rstParent.MoveNext
Loop

End Sub

'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLedger = Nothing
End Sub


Private Sub lvwLedger_DblClick()

On Error Resume Next

Dim Count As Integer
' Selected Item Key will be like this - "A2001"
' So We have to Fetch only value
With lvwLedger.SelectedItem
    m_SubParentID = Val(Mid(.Key, 2))
    txtLedgerName.Text = .SubItems(1)
    'txtLedgerName.Text = .Text
    cmbParent.Locked = True
End With

m_DBOperation = Update
cmdOk.Caption = "&Update"

End Sub





Private Sub txtLedgerName_LostFocus()
'txtLedgerName = ConvertToProperCase(txtLedgerName.Text)
End Sub


