VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAccs 
   Caption         =   "Bank Accounts..."
   ClientHeight    =   6675
   ClientLeft      =   1020
   ClientTop       =   1530
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   8145
   Begin VB.Frame fra 
      Height          =   5835
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   570
      Width           =   7245
      Begin VB.CommandButton cmdAccNames 
         Caption         =   "..."
         Height          =   315
         Left            =   6645
         TabIndex        =   28
         Top             =   675
         Width           =   315
      End
      Begin VB.ComboBox cmbAccNames 
         Height          =   315
         Left            =   2235
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   660
         Width           =   4260
      End
      Begin VB.ComboBox cmbAccHeads 
         Height          =   315
         ItemData        =   "frmAccs.frx":0000
         Left            =   2235
         List            =   "frmAccs.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   270
         Width           =   4260
      End
      Begin VB.Frame fraPassBook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2265
         Left            =   270
         TabIndex        =   14
         Top             =   2850
         Width           =   6675
         Begin VB.CommandButton cmdNextTrans 
            Caption         =   ">"
            Height          =   315
            Left            =   6270
            TabIndex        =   16
            Top             =   690
            Width           =   375
         End
         Begin VB.CommandButton cmdPrevTrans 
            Caption         =   "<"
            Height          =   315
            Left            =   6270
            TabIndex        =   15
            Top             =   135
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2205
            Left            =   90
            TabIndex        =   17
            Top             =   90
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3889
            _Version        =   327680
            Rows            =   5
            AllowUserResizing=   1
         End
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Left            =   4830
         TabIndex        =   13
         Top             =   1515
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton cmdCheque 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   12
         Top             =   1500
         Width           =   345
      End
      Begin VB.ComboBox cmbCheque 
         Height          =   315
         Left            =   4830
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "cmbCheque"
         Top             =   1515
         Width           =   1725
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   6795
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1545
         Width           =   1965
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   4830
         TabIndex        =   8
         Top             =   1185
         Width           =   2085
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   180
         TabIndex        =   7
         Top             =   2340
         Width           =   6825
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5820
         TabIndex        =   6
         Top             =   5400
         Width           =   1305
      End
      Begin VB.CheckBox chkBackLog 
         Caption         =   "Clear backlog"
         Height          =   225
         Left            =   2640
         TabIndex        =   5
         Top             =   5430
         Width           =   1395
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1140
         TabIndex        =   4
         Top             =   1215
         Width           =   1965
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   5400
         Width           =   1425
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   1905
         Width           =   5805
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2745
         Left            =   165
         TabIndex        =   18
         Top             =   2505
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   4842
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Pass book"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Select the account name :"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Account head type :"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Transaction : "
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Amount (Rs) : "
         Height          =   255
         Left            =   3810
         TabIndex        =   22
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Instument no:"
         Height          =   195
         Left            =   3750
         TabIndex        =   21
         Top             =   1605
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Particulars : "
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1995
         Width           =   945
      End
      Begin VB.Label Label11 
         Caption         =   "Date : "
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1245
         Width           =   735
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   6390
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   11271
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Account Transactions"
            Key             =   "Transactions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAccs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_AccHead  As String
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub fraTransact_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub cmbAccHeads_Click()
    If m_AccHead = cmbAccHeads.Text Then Exit Sub
    m_AccHead = cmbAccHeads.Text
    Me.cmbTrans.Clear
    If cmbAccHeads.ListIndex = 0 Then
        cmbTrans.AddItem "Deposit"   ' TransType = wwithdraw
        cmbTrans.AddItem "WithDrawn"   'TransType= wdeposit
        cmbTrans.AddItem "Interest Recieved"     'TransType = wcharges
        cmbTrans.AddItem "Charges Paid"   'TransType=Winterest
    ElseIf cmbAccHeads.ListIndex = 1 Then
        cmbTrans.AddItem "Amount Recieved"  'TransType=wCnterest
    ElseIf cmbAccHeads.ListIndex = 2 Then
        cmbTrans.AddItem "Amount Issued"  'TransType=wInterest
    End If
    
End Sub


Private Sub cmdAccNames_Click()
If cmbAccHeads.List < 0 Then Exit Sub
frmAcDef.txtAccHead.Text = m_AccHead
frmAcDef.Show vbModal, Me
End Sub


