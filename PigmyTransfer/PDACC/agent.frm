VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAgentTrans 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   1950
   ClientTop       =   1635
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   7470
   Begin VB.Frame fraAgent 
      Height          =   5865
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7260
      Begin VB.CommandButton cmdAmount 
         Caption         =   "...."
         Height          =   255
         Left            =   6570
         TabIndex        =   31
         Top             =   1410
         Width           =   315
      End
      Begin VB.Frame fraAgentPassbook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2205
         Left            =   270
         TabIndex        =   17
         Top             =   3060
         Width           =   6645
         Begin VB.CommandButton cmdAgentPrevTrans 
            Caption         =   "<"
            Height          =   315
            Left            =   6270
            TabIndex        =   20
            Top             =   135
            Width           =   375
         End
         Begin VB.CommandButton cmdAgentNextTrans 
            Caption         =   ">"
            Height          =   315
            Left            =   6270
            TabIndex        =   19
            Top             =   690
            Width           =   375
         End
         Begin VB.CommandButton cmdAgentPrint 
            Height          =   345
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1860
            Width           =   405
         End
         Begin MSFlexGridLib.MSFlexGrid grdAgent 
            Height          =   1995
            Left            =   90
            TabIndex        =   21
            Top             =   150
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3519
            _Version        =   393216
            Rows            =   5
            Cols            =   3
            WordWrap        =   -1  'True
            AllowUserResizing=   1
         End
      End
      Begin VB.ComboBox cmbAgentList 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "agent.frx":0000
         Left            =   1275
         List            =   "agent.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   810
         Width           =   5700
      End
      Begin VB.ComboBox cmbAgentParticulars 
         Height          =   315
         Left            =   1215
         TabIndex        =   15
         Top             =   2115
         Width           =   5730
      End
      Begin VB.CommandButton cmdAgentLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2505
         TabIndex        =   14
         Top             =   270
         Width           =   930
      End
      Begin VB.CommandButton cmdAgentUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4620
         TabIndex        =   13
         Top             =   5460
         Width           =   1335
      End
      Begin VB.TextBox txtAgentDate 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1230
         TabIndex        =   12
         Top             =   1395
         Width           =   1425
      End
      Begin VB.CommandButton cmdAgentAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   5460
         Width           =   1125
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   180
         TabIndex        =   10
         Top             =   2490
         Width           =   6825
      End
      Begin VB.TextBox txtAgentAmount 
         Height          =   285
         Left            =   4830
         TabIndex        =   9
         Top             =   1395
         Width           =   1635
      End
      Begin VB.ComboBox cmbAgentTrans 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1740
         Width           =   1875
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   1275
         Width           =   6795
      End
      Begin VB.TextBox txtAgentNo 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   6
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox txtAgentCheque 
         Height          =   300
         Left            =   4830
         TabIndex        =   5
         Top             =   1740
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Frame fraAgentInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2205
         Left            =   300
         TabIndex        =   2
         Top             =   3075
         Width           =   6615
         Begin VB.CommandButton cmdAgentNote 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6090
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   90
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   1995
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   3519
            _Version        =   393217
            TextRTF         =   $"agent.frx":0004
         End
      End
      Begin VB.CommandButton cmdAgentTransactDate 
         Caption         =   "...."
         Height          =   255
         Left            =   2730
         TabIndex        =   1
         Top             =   1410
         Width           =   315
      End
      Begin ComctlLib.TabStrip TabAgentStrip2 
         Height          =   2745
         Left            =   150
         TabIndex        =   22
         Top             =   2625
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   4842
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Instructions"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Pass book"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAgent 
         Caption         =   "Agents :"
         Height          =   225
         Left            =   180
         TabIndex        =   30
         Top             =   870
         Width           =   975
      End
      Begin VB.Label lblAgentDate 
         Caption         =   "Date : "
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblAgentParticular 
         Caption         =   "Particulars : "
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label lblAgentBalance 
         Caption         =   "Balance : Rs. 00.00"
         Height          =   285
         Left            =   5175
         TabIndex        =   27
         Top             =   390
         Width           =   1755
      End
      Begin VB.Label lblAgentInstrNo 
         Caption         =   "Instument no:"
         Height          =   195
         Left            =   3750
         TabIndex        =   26
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label lblAgentAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   255
         Left            =   3810
         TabIndex        =   25
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label lblAgentTrans 
         Caption         =   "Transaction : "
         Height          =   285
         Left            =   150
         TabIndex        =   24
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label lblAgentNo 
         Caption         =   "Account No. : "
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   330
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmAgentTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_AccID As Long
Public m_UserID As Integer
Private m_AccClosed As Boolean
Private m_rstPassBook As Recordset
Private m_CustReg As New clsCustReg
Private m_Notes As New clsNotes
Private M_setUp As New clsSetup

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmPDReport As frmPDReport
Attribute m_frmPDReport.VB_VarHelpID = -1
'Private WithEvents m_frmSearch As frmPDSearch
Const CTL_MARGIN = 15
Private m_accUpdatemode As Integer

Private Sub SetKannadaCaption()
Dim Ctrl As Control
On Error Resume Next
'Now Assign the Kannada fonts to the All controls
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next Ctrl
'Now Assign The Names to the Controls
'The Below Code load From The the resource file



End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
'Centre the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
cmdAgentPrint.Picture = LoadResPicture(120, vbResBitmap)

'set kannada caption
Call SetKannadaCaption
 
'Intialize the custreg Calss
m_CustReg.ModuleId = wis_PDAcc


'Fill up transaction Types
   With cmbAgentTrans
        .AddItem LoadResString(gLangOffSet + 271)
        .AddItem LoadResString(gLangOffSet + 272)
        .AddItem LoadResString(gLangOffSet + 273)
        .AddItem LoadResString(gLangOffSet + 274)
 End With
     
'Fill up particulars with default values from PDAgent.INI
    Dim Particulars As String
    Dim I As Integer
    Do
        Particulars = ReadFromIniFile("Particulars", _
                "Key" & I, gAppPath & "\PDAgent.INI")
        If Trim$(Particulars) <> "" Then
            cmbAgentParticulars.AddItem Particulars
        End If
        I = I + 1
    Loop Until Trim$(Particulars) = ""


'Adjust the Grid for Pass book
With grdAgent
    .Clear
    .Rows = 11
    .Cols = 4
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = LoadResString(gLangOffSet + 37): .ColWidth(0) = 1000 ' "Date"
    .Col = 1: .Text = LoadResString(gLangOffSet + 39): .ColWidth(1) = 1000 '"Particulars"
    .Col = 2: .Text = LoadResString(gLangOffSet + 276): .ColWidth(3) = 1000 '"Debit"
    .Col = 3: .Text = LoadResString(gLangOffSet + 328): .ColWidth(4) = 1000 '"Pigmy commission"
End With

'Load Agent Name
    Call LoadAgentNames(cmbAgentList)
    txtAgentNo.Locked = True

'Set Report Frame
    optDepositBalance.value = True
    Call optDepositBalance_Click
TabStrip2.Tabs(1).Selected = True
fraAgent.ZOrder 0
fraAgentInstructions.ZOrder 0

Screen.MousePointer = vbDefault

End Sub


