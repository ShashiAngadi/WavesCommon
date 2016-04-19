VERSION 5.00
Begin VB.Form frmInt 
   Caption         =   "Interst Test"
   ClientHeight    =   4845
   ClientLeft      =   2115
   ClientTop       =   3465
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7245
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   555
      Left            =   330
      TabIndex        =   24
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   5400
      TabIndex        =   23
      Top             =   2490
      Width           =   1485
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   405
      Left            =   5400
      TabIndex        =   22
      Top             =   2970
      Width           =   1485
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "Int Date"
      Height          =   405
      Left            =   3750
      TabIndex        =   21
      Top             =   2490
      Width           =   1425
   End
   Begin VB.CommandButton cmdInterest 
      Caption         =   "Interest Rate"
      Height          =   405
      Left            =   3780
      TabIndex        =   20
      Top             =   3000
      Width           =   1395
   End
   Begin VB.TextBox txtIntDate 
      Height          =   345
      Left            =   5370
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1830
      Width           =   1665
   End
   Begin VB.TextBox txtIntRate 
      Height          =   285
      Left            =   750
      TabIndex        =   16
      Text            =   "Text9"
      Top             =   1770
      Width           =   765
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   7
      Left            =   5340
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   1350
      Width           =   1905
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   6
      Left            =   5340
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   960
      Width           =   1905
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   5
      Left            =   5340
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   570
      Width           =   1905
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   4
      Left            =   5340
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   180
      Width           =   1905
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   3
      Left            =   1710
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1350
      Width           =   1935
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   2
      Left            =   1710
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   570
      Width           =   1935
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Index           =   0
      Left            =   1710
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label lblIntdate 
      Caption         =   "Interest Date :"
      Height          =   315
      Left            =   4080
      TabIndex        =   18
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Label lblIntRate 
      Caption         =   "Interest"
      Height          =   225
      Left            =   60
      TabIndex        =   17
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label8 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3690
      TabIndex        =   15
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3690
      TabIndex        =   14
      Top             =   1020
      Width           =   1545
   End
   Begin VB.Label Label6 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3690
      TabIndex        =   11
      Top             =   630
      Width           =   1545
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3690
      TabIndex        =   10
      Top             =   240
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1020
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   630
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   240
      Width           =   1545
   End
End
Attribute VB_Name = "frmInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_intClass As New clsInterest



Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdDate_Click()
Dim intrate As Single

Static Count As Integer
If Count = 0 Then
   Dim module As wisModules
   module = wis_FDAcc
   Dim SchemeID As Integer
   SchemeID = 1
   m_intClass.ModuleID = module
   m_intClass.SchemeName = SchemeID
      intrate = m_intClass.InterestRate("14/12/1998", "12/12/2010")
End If
Me.txtIntDate.Text = m_intClass.NextInterestDate
On Error GoTo Errline
If Trim(Me.txtIntDate.Text) <> "" Then
   Me.txtInterest(Count).Text = m_intClass.SchemeID
End If


Errline:
Count = Count + 1


End Sub





Private Sub cmdInterest_Click()
Dim intrate As Single

Static Count As Integer
If Count = 0 Then
   Dim module As wisModules
   module = wis_FDAcc
   Dim SchemeID As Integer
   SchemeID = 1
   m_intClass.ModuleID = module
   m_intClass.SchemeID = SchemeID
      intrate = m_intClass.InterestRate("14/12/1998", "12/12/2010")
      Me.txtIntRate.Text = intrate
      Count = Count + 1
      Exit Sub
End If
Me.txtIntRate.Text = m_intClass.NextInterestRate

On Error GoTo Errline
If Me.txtIntRate.Text <> "0" Then
      Me.txtInterest(Count).Text = m_intClass.SchemeID
End If



Errline:
Count = Count + 1



End Sub


Private Sub cmdReset_Click()

Me.txtIntRate.Text = CStr(m_intClass.InterestRate("12/12/1999", "12/12/2000"))
m_intClass.ClearInterest
Call Form_Load
End Sub

Private Sub cmdSave_Click()
Dim SchemeID As Integer
Dim intrate As Single
Dim IntDate As String
Dim module As wisModules

intrate = CSng(txtIntRate.Text)
IntDate = CStr(txtIntDate.Text)
SchemeID = 1
module = wis_FDAcc
If Not m_intClass.SaveInterest(module, SchemeID, intrate, IntDate) Then
      MsgBox "Unable To Save Interest"
End If


End Sub


Private Sub Form_Load()

Dim ctrl As Control
For Each ctrl In Me
   If TypeOf ctrl Is TextBox Then
      ctrl.Text = ""
   End If
   
Next ctrl

End Sub


Private Sub Form_Unload(Cancel As Integer)
gDBTrans.CloseDB
End
End Sub


