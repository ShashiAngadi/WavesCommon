VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheques"
   ClientHeight    =   3105
   ClientLeft      =   3465
   ClientTop       =   1950
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4170
      TabIndex        =   7
      Top             =   2640
      Width           =   960
   End
   Begin VB.Frame fraRemove 
      Caption         =   "Current leaf set"
      Height          =   1845
      Left            =   270
      TabIndex        =   8
      Top             =   540
      Width           =   4755
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clea&r"
         Height          =   315
         Left            =   3750
         TabIndex        =   15
         Top             =   1080
         Width           =   945
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   315
         Left            =   3750
         TabIndex        =   14
         Top             =   660
         Width           =   945
      End
      Begin VB.CommandButton cmdInvert 
         Caption         =   "Invert Selection"
         Height          =   345
         Left            =   3720
         TabIndex        =   6
         Top             =   1470
         Width           =   1395
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   3750
         TabIndex        =   5
         Top             =   270
         Width           =   915
      End
      Begin VB.ListBox lstCheque 
         Height          =   1185
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   270
         Width           =   3465
      End
   End
   Begin VB.Frame fraAdd 
      Caption         =   "Issue new cheque book"
      Height          =   1860
      Left            =   300
      TabIndex        =   4
      Top             =   540
      Width           =   4665
      Begin VB.TextBox txtSeriesNo 
         Height          =   315
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.TextBox txtLeaves 
         Height          =   315
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1050
         Width           =   705
      End
      Begin VB.TextBox txtStartNo 
         Height          =   315
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   1
         Top             =   660
         Width           =   1995
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Add"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   315
         Left            =   3765
         TabIndex        =   3
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblSeries 
         Caption         =   "Series No"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label lblLeaves 
         Caption         =   "No of leaves:"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lblStartNo 
         Caption         =   "Start no:"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   720
         Width           =   1245
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2505
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4419
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add cheque book"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Add new cheque book"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Remove cheque leaves"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public p_AccId As Long
Public p_ModuleId As wisModules

'Public Event OKClicked(Series As String, StartNo As Long, Leaves As Long, Cancel As Boolean)
'Public Event RemoveLeaves(LeafArr() As Long, Opeation As wis_ChequeTrans)


Private Function FillChequeList()
Dim I As Integer
Dim SqlStr As String
Dim rst As ADODB.Recordset
lstCheque.Clear

SqlStr = "SELECT * FROM ChequeMaster WHERE AccID = " & p_AccId & _
        " AND AccType = '" & p_ModuleId & "'"
gDBTrans.SQLStmt = SqlStr
If gDBTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

While Not rst.EOF
    If FormatField(rst("Trans")) <> chqPay Then
        lstCheque.AddItem FormatField(rst("ChequeNo"))
        If rst("Trans") = chqStop Or rst("Trans") = chqStop Then
            lstCheque.Selected(lstCheque.NewIndex) = True
        End If
    End If
    rst.MoveNext
Wend


    
End Function

Private Sub SetKannadaCaption()
Dim Ctrl As Control
    For Each Ctrl In Me
        Ctrl.Font.Name = gFontName
        If Not TypeOf Ctrl Is ComboBox Then
            Ctrl.Font.Size = gFontSize
        End If
    Next
'with general form.
TabStrip1.Tabs(1).Caption = LoadResString(gLangOffSet + 141)
TabStrip1.Tabs(2).Caption = LoadResString(gLangOffSet + 142)
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 11)
'with tabstrip1 or frame add
Me.fraAdd.Caption = LoadResString(gLangOffSet + 141)
Me.lblStartNo.Caption = LoadResString(gLangOffSet + 144)  '"¢«Æ≈˛¡ ÕÆ≤˙Â"
Me.lblLeaves.Caption = LoadResString(gLangOffSet + 145)     '"∂˙∞˝ ®…˙≥  ÕÆ≤˙Â"
Me.cmdOK.Caption = LoadResString(gLangOffSet + 10)     '"Õ˙ÛêÕÙ"
'with tabstrip2 or Frame Remove
fraRemove.Caption = LoadResString(gLangOffSet + 142)   '"∂ÒëﬁèÙ«Ù∆ ∂˙∞Ù–≥ Ù"
Me.cmdRemove.Caption = LoadResString(gLangOffSet + 12)    '"¿˙≥˙"
Me.cmdInvert.Caption = LoadResString(gLangOffSet + 21)    '"•…Ò⁄ Õ˙…˙∞Î¬˝ "
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
Dim I As Integer
Dim SqlStr As String

Dim Found As Boolean

gDBTrans.BeginTrans
For I = 0 To lstCheque.ListCount - 1
    If lstCheque.Selected(I) = True Then
        Found = True
        SqlStr = "UPDATE ChequeMaster Set Trans = " & chqIssue & _
            " WHERE  ChequeNO = " & lstCheque.List(I) & _
            " AND AccId = " & p_AccId & " AND AccType = '" & p_ModuleId & "'"
        gDBTrans.SQLStmt = SqlStr
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Sub
        End If
    End If
Next I
gDBTrans.CommitTrans


If Found Then MsgBox "Specified leaves stopped successfully.", vbInformation, gAppName & " - Error"
Call FillChequeList

End Sub

Private Sub cmdInvert_Click()
Dim I As Integer
For I = 0 To lstCheque.ListCount - 1
    lstCheque.Selected(I) = Not lstCheque.Selected(I)
Next I


End Sub

Private Sub cmdOK_Click()

'iF DEVELOPER HAS NOT SPEICIFIED THE
'ACCOUNT NO AND MODULE ID THEN RAISE A ERROR
If p_AccId = 0 Or p_ModuleId = 0 Then
    Err.Raise 50012, , "Account id or Module id not set"
    Exit Sub
End If

Dim LeafCount As Long
Dim Cancel As Boolean


'Validate the Series No
    If Trim(txtSeriesNo.Text) <= "" Then
        MsgBox "You have not specified the series no", vbExclamation, gAppName & " - Error"
        ActivateTextBox txtSeriesNo
        Exit Sub
    End If

'Validate the Cheque Start No
    If Val(txtStartNo.Text) <= 0 Then
        MsgBox LoadResString(gLangOffSet + 503), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtStartNo
        Exit Sub
    End If
Dim StartNo As Long
StartNo = Val(txtStartNo)
'Validate the no of leaves
    If Val(txtLeaves.Text) <= 0 Then
        MsgBox LoadResString(gLangOffSet + 504), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtLeaves
        Exit Sub
    Else
        LeafCount = Val(txtLeaves.Text)
    End If
Dim Leaves As Integer
Leaves = Val(txtLeaves.Text)
'Validate the number of leaves
    If LeafCount <> 1 And LeafCount <> 10 And LeafCount <> 25 And LeafCount <> 50 And LeafCount <> 100 Then
         'MsgBox "Invalid number of leaves specified !" & vbCrLf & "Number of leaves should be 10, 25, 50 , 100", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 504), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtLeaves
        Exit Sub
    End If


'Now insert into the data base
Dim SqlStr As String
Dim Series As String
Dim J As Integer
Dim rst As ADODB.Recordset
Series = Trim(txtSeriesNo.Text)

'Now Cheque For the Existance same cheque no
SqlStr = "SELECT * FROM ChequeMaster WHERE SeriesNo = " & AddQuotes(Series, True) & _
        " AND (ChequeNO >= " & StartNo & " AND ChequeNo < " & StartNo + Leaves & ")"
gDBTrans.SQLStmt = SqlStr
If gDBTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    MsgBox "This cheque already issued", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Now Insert each leaves record as saperate record
gDBTrans.BeginTrans
For J = 0 To Leaves - 1
    SqlStr = "INSERT INTO ChequeMaster (SeriesNo,ChequeNo,AccID,AccType,IssuedDate,Trans)" & _
        " VALUES (" & AddQuotes(Series, True) & "," & _
        StartNo + J & "," & p_AccId & "," & _
        "'" & p_ModuleId & "'," & _
        "#" & gStrDate & "#," & chqIssue & " )"
        
    gDBTrans.SQLStmt = SqlStr
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Sub
    End If
Next
gDBTrans.CommitTrans




'RaiseEvent OKClicked(Val(txtStartNo.Text), Val(txtLeaves.Text), Cancel)
Call FillChequeList

'MsgBox "Cheque Book added successfully !", vbExclamation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 637), vbExclamation, gAppName & " - Error"
txtStartNo.Text = ""
txtLeaves.Text = ""

End Sub

Private Sub cmdRemove_Click()
Dim I As Integer
Dim SqlStr As String

Dim Found As Boolean

gDBTrans.BeginTrans
For I = 0 To lstCheque.ListCount - 1
    If lstCheque.Selected(I) = True Then
        Found = True
        'sqlstr="UPDATE ChequeMaster Set Trans = " & chq
        SqlStr = "DELETE * FROM ChequeMaster Where ChequeNo = " & lstCheque.List(I) & _
            " AND AccId = " & p_AccId & " AND AccType = '" & p_ModuleId & "'"
        gDBTrans.SQLStmt = SqlStr
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Sub
        End If
    End If
Next I
gDBTrans.CommitTrans


'RaiseEvent RemoveLeaves(LeafArr)

'If Found Then MsgBox "Specified leaves remove successfully.", vbInformation, gAppName & " - Error"
If Found Then MsgBox LoadResString(gLangOffSet + 557), vbInformation, gAppName & " - Error"

Call FillChequeList

End Sub

Private Sub cmdStop_Click()
Dim I As Integer
Dim SqlStr As String

Dim Found As Boolean

gDBTrans.BeginTrans
For I = 0 To lstCheque.ListCount - 1
    If lstCheque.Selected(I) = True Then
        Found = True
        SqlStr = "UPDATE ChequeMaster Set Trans = " & chqStop & _
            " WHERE  ChequeNO = " & lstCheque.List(I) & _
            " AND AccId = " & p_AccId & " AND AccType = '" & p_ModuleId & "'"
        gDBTrans.SQLStmt = SqlStr
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Sub
        End If
    End If
Next I
gDBTrans.CommitTrans


If Found Then MsgBox "Specified leaves stopped successfully.", vbInformation, gAppName & " - Error"
Call FillChequeList

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
      If Me.TabStrip1.SelectedItem.Index = TabStrip1.Tabs.Count Then
            TabStrip1.Tabs(1).Selected = True
      Else
            TabStrip1.Tabs(TabStrip1.SelectedItem.Index + 1).Selected = True
      End If
End If

End Sub

Private Sub Form_Load()
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
fraAdd.Visible = True
fraRemove.Visible = False
fraAdd.ZOrder 0

'Fill Remove list box
Call FillChequeList
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCheque = Nothing
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = 1 Then
        fraAdd.Visible = True
        fraRemove.Visible = False
    End If
    If TabStrip1.SelectedItem.Index = 2 Then
        fraAdd.Visible = False
        fraRemove.Visible = True
    End If

End Sub

Private Sub txtLeaves_Change()
If Trim$(txtStartNo.Text) <> "" And Trim$(txtLeaves.Text) <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If

End Sub

Private Sub txtStartNo_Change()
If Trim$(txtStartNo.Text) <> "" And Trim$(txtLeaves.Text) <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
End Sub


