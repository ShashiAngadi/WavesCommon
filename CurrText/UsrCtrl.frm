VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   2325
   ClientTop       =   2280
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5625
   Begin WIS_Currency_Text_Box.CurrText CurrText2 
      Height          =   315
      Left            =   450
      TabIndex        =   8
      Top             =   600
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin WIS_Currency_Text_Box.CurrText CurrText1 
      Height          =   345
      Left            =   450
      TabIndex        =   7
      Top             =   150
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   609
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3630
      TabIndex        =   6
      Top             =   630
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   315
      Left            =   3630
      TabIndex        =   4
      Top             =   90
      Width           =   1785
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Rs 1234578"
      Top             =   990
      Width           =   2715
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3300
      TabIndex        =   0
      Top             =   1590
      Width           =   1845
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   390
         TabIndex        =   1
         Text            =   "Rs"
         Top             =   330
         Width           =   1155
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   180
      TabIndex        =   5
      Top             =   2490
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   300
      TabIndex        =   3
      Top             =   1410
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label1.Caption = Me.CurrText1.TextInFigure
    'Debug.Print CurrText1.NumberInFigure(CurrText1.Value)
End Sub

Private Sub CurrTextBox1_Click()

End Sub


Private Sub CurrTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
Private Sub Command2_Click()
CurrText1.Enabled = False

'Debug.Print CurrText2.Value
CurrText2.Value = -12342
Exit Sub
Label2.Caption = "..* "
Label2.Caption = Label2.Caption & Me.CurrText1.SelStart
Label2.Caption = Label2.Caption & vbCrLf & Me.CurrText1.SelLength
Label2.Caption = Label2.Caption & vbCrLf & Me.CurrText1.SelText
End Sub


Private Sub CurrText1_GotFocus()
'CurrText1.SelStart = 0
'CurrText1.SelLength = 1
Label2.Caption = CurrText1.SelText
With CurrText1
    .SelStart = 0
    .SelLength = Len(.Text)
End With
Label2.Caption = CurrText1.SelText
End Sub


Private Sub CurrText2_GotFocus()
With CurrText2
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


Private Sub Form_Load()
 ' = AlignConstants.vbAlignRight

 CurrText2 = 1234567890
End Sub


