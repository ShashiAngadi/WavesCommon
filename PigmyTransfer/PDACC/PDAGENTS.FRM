VERSION 5.00
Begin VB.Form frmPDAgents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1905
   ClientLeft      =   1860
   ClientTop       =   1785
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4815
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   990
      Width           =   2130
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   3645
      TabIndex        =   7
      Top             =   1485
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Register"
      Height          =   330
      Left            =   1980
      TabIndex        =   6
      Top             =   1485
      Width           =   1020
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   345
      Left            =   180
      TabIndex        =   5
      Top             =   1485
      Width           =   1020
   End
   Begin VB.TextBox txtArea 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   2805
   End
   Begin VB.CommandButton cmdCustReg 
      Caption         =   "..."
      Height          =   285
      Left            =   4050
      TabIndex        =   3
      Top             =   990
      Width           =   555
   End
   Begin VB.ComboBox cmbNames 
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   165
      Width           =   2850
   End
   Begin VB.Label lblSelectedAgent 
      Caption         =   "Selcted Agent :"
      Height          =   240
      Left            =   225
      TabIndex        =   9
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label lblWorkArea 
      Caption         =   "Work Area"
      Height          =   300
      Left            =   225
      TabIndex        =   2
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label lblAgentNames 
      Caption         =   "Agent Name List :"
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   150
      Width           =   1470
   End
End
Attribute VB_Name = "frmPDAgents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CustReg As New clsCustReg
Dim m_CustomerID As Long

Private Sub LoadAgentNames()
gDBTrans.SQLStmt = "Select UserId,Title,FirstName,MiddleName, LastName " & _
        " from Agents,NameTab Where Agents.CustomerId=NameTab.CustomerId"
If gDBTrans.SQLFetch < 1 Then Exit Sub
Dim AgentName As String
cmbNames.Clear
'Add One Blank entry To Add new Agent
cmbNames.AddItem " "
While Not gDBTrans.Rst.EOF
    AgentName = FormatField(gDBTrans.Rst("Title")) & " " & FormatField(gDBTrans.Rst("FirstName")) & " " & _
        FormatField(gDBTrans.Rst("MiddleName")) & " " & FormatField(gDBTrans.Rst("LastName"))
    cmbNames.AddItem AgentName
    cmbNames.ItemData(cmbNames.NewIndex) = FormatField(gDBTrans.Rst("UserId"))
    gDBTrans.Rst.MoveNext
Wend
End Sub

Private Sub SetKannadaCaption()
'declare variables
Dim Ctrl As Control

On Error Resume Next
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
     End If
Next
Me.lblAgentNames.Caption = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35)
Me.lblSelectedAgent.Caption = LoadResString(gLangOffSet + 468)
Me.lblWorkArea.Caption = LoadResString(gLangOffSet + 329)
Me.cmdAdd.Caption = LoadResString(gLangOffSet + 24)
Me.cmdClose.Caption = LoadResString(gLangOffSet + 11)
Me.cmdRemove.Caption = LoadResString(gLangOffSet + 12)
End Sub


Private Sub cmbNames_Click()

If cmbNames.ListIndex < 1 Then
    cmdAdd.Caption = "Register"
    cmdAdd.Tag = "Register"
    m_CustReg.NewCustomer
    txtName.Text = ""
    txtArea.Text = ""
    m_CustomerID = 0
Else
    cmdAdd.Enabled = True
    gDBTrans.SQLStmt = "Select UserId, AreaOfWork ,Title, FirstName, MiddleName, LastName " & _
        " From Agents,NameTab Where UserId = " & cmbNames.ItemData(cmbNames.ListIndex) & _
        " and Agents.CustomerId = NameTab.CustomerId "
    txtName.Text = cmbNames.List(cmbNames.ListIndex)
    If gDBTrans.SQLFetch > 0 Then
        txtArea.Text = FormatField(gDBTrans.Rst("AreaOfWork"))
        cmdAdd.Caption = "&Update"
        cmdAdd.Tag = "Update"
        m_CustomerID = cmbNames.ItemData(cmbNames.ListIndex)
    End If
End If
End Sub


Private Sub cmdAdd_Click()
Dim UserID As Integer
Dim LstIndx As Integer
LstIndx = cmbNames.ListIndex
If UCase(cmdAdd.Tag) = "UPDATE" Then
    If m_CustomerID = 0 Then
        'MsgBox "You have not selected the Agent", , wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 783), , wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    gDBTrans.BeginTrans
    If m_CustReg.CustomerLoaded Then
        If Not m_CustReg.SaveCustomer Then Exit Sub
    End If
    gDBTrans.SQLStmt = "Update Agents Set AreaOfWork = '" & Trim$(txtArea.Text) & "'" & _
        " Where CustomerID = " & m_CustomerID
    If Not gDBTrans.SQLExecute Then
        'MsgBox "Can not Update Agent detail"
        MsgBox LoadResString(gLangOffSet + 593)
        gDBTrans.RollBack
        Exit Sub
    End If
    gDBTrans.CommitTrans
Else
    If m_CustomerID = 0 Then
        'MsgBox "You have not not entered the Agent's detail", , wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 594), , wis_MESSAGE_TITLE
        Call cmdCustReg_Click
        If Trim$(m_CustReg.FullName) = "" Then
            Exit Sub
        End If
    End If
    
    gDBTrans.SQLStmt = "Select * From Agents where UserId = " & m_CustomerID
    If gDBTrans.SQLFetch > 0 Then
        Exit Sub
    End If
    
    
    gDBTrans.SQLStmt = "Select * From Agents order by UserId desc"
    If gDBTrans.SQLFetch > 0 Then
        UserID = Val(FormatField(gDBTrans.Rst(0))) + 1
    Else
        UserID = 1
    End If
    gDBTrans.BeginTrans
    If Not m_CustReg.SaveCustomer Then Exit Sub
    gDBTrans.SQLStmt = "Insert Into Agents(UserId,CustomerId,AreaOfWork) Values (" & _
        UserID & ", " & m_CustomerID & ", '" & Trim$(txtArea.Text) & "')"
    If Not gDBTrans.SQLExecute Then
        'MsgBox "Can not Add Agents detail"
        MsgBox LoadResString(gLangOffSet + 595)
        gDBTrans.RollBack
        Exit Sub
    End If
    gDBTrans.CommitTrans
    'MsgBox "New Agent created"
    MsgBox LoadResString(gLangOffSet + 596)
End If
Call LoadAgentNames
On Error Resume Next
cmbNames.ListIndex = LstIndx
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub cmdCustReg_Click()
Dim strName As String
Dim Count As Integer

If cmbNames.ListIndex > 0 Then
    m_CustReg.LoadCustomerInfo (cmbNames.ItemData(cmbNames.ListIndex))
Else
    m_CustReg.NewCustomer
End If

m_CustReg.ShowDialog
strName = m_CustReg.FullName
m_CustomerID = m_CustReg.CustomerId
For Count = 0 To cmbNames.ListCount - 1
    If m_CustomerID = cmbNames.ItemData(Count) Then
        cmdAdd.Caption = "Update"
        cmdAdd.Tag = "Update"
        Exit Sub
    End If
Next Count

If Trim(strName) = "" Then
    m_CustomerID = 0
    cmdAdd.Enabled = False
    cmbNames.ListIndex = -1
Else
    txtName.Text = strName
    'cmbNames.AddItem strName
    'cmbNames.ItemData(cmbNames.NewIndex) = m_CustomerId
End If

End Sub

Private Sub cmdRemove_Click()
If txtName.Text = "" Or cmbNames.ListIndex < 1 Then
    'MsgBox "You Have not selectd the agent to remove"
    MsgBox LoadResString(gLangOffSet + 597)
    Exit Sub
End If

gDBTrans.SQLStmt = " Select AccId From PDMaster Where UserId = " & _
        cmbNames.ItemData(cmbNames.ListIndex) & " And ClosedDate = NULL"
If gDBTrans.SQLFetch > 0 Then
    'MsgBox "This agent is having some accounts with him " & vbCrLf & vbCrLf & _
    '   " Transfer his accounts to other agents , then continue"
    MsgBox LoadResString(gLangOffSet + 598) & vbCrLf & vbCrLf & _
        LoadResString(gLangOffSet + 599)
    Exit Sub
End If
gDBTrans.BeginTrans
gDBTrans.SQLStmt = "Delete * From Agents where UserId = " & cmbNames.ItemData(cmbNames.ListIndex)
If Not gDBTrans.SQLExecute Then
    'MsgBox "Unable to remove the Agent"
    MsgBox LoadResString(gLangOffSet + 600)
    gDBTrans.RollBack
    Exit Sub
End If
gDBTrans.CommitTrans
Call LoadAgentNames
End Sub

Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

cmdAdd.Enabled = True
'Load the Agents Name in the Combo
' set kannada Captions for the  controls
Call SetKannadaCaption

Call LoadAgentNames
End Sub

