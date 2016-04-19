VERSION 5.00
Begin VB.Form frmCashIndexOLD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Index"
   ClientHeight    =   6384
   ClientLeft      =   -1128
   ClientTop       =   1056
   ClientWidth     =   6216
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   6216
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   396
      Left            =   3000
      ScaleHeight     =   348
      ScaleWidth      =   2604
      TabIndex        =   39
      Top             =   0
      Width           =   2655
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   396
      Left            =   3240
      TabIndex        =   19
      Top             =   5952
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   396
      Left            =   1440
      TabIndex        =   18
      Top             =   5952
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   1236
      Left            =   120
      TabIndex        =   31
      Top             =   4584
      Width           =   5895
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3336
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4296
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Net Amount :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Refund :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   3504
         TabIndex        =   36
         Top             =   276
         Width           =   780
      End
      Begin VB.Label Label13 
         Caption         =   "Total  Received :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   280
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cash Particulars"
      Height          =   4236
      Left            =   120
      TabIndex        =   20
      Top             =   384
      Width           =   5895
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   0
         Left            =   2328
         TabIndex        =   0
         Tag             =   "500"
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   1
         Left            =   2328
         TabIndex        =   1
         Tag             =   "100"
         Top             =   1095
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   2
         Left            =   2328
         TabIndex        =   2
         Tag             =   "50"
         Top             =   1470
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   3
         Left            =   2328
         TabIndex        =   3
         Tag             =   "20"
         Top             =   1845
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   4
         Left            =   2328
         TabIndex        =   4
         Tag             =   "10"
         Top             =   2220
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   5
         Left            =   2328
         TabIndex        =   5
         Tag             =   "5"
         Top             =   2595
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   6
         Left            =   2328
         TabIndex        =   6
         Tag             =   "2"
         Top             =   2970
         Width           =   1050
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   7
         Left            =   2328
         TabIndex        =   7
         Tag             =   "1"
         Top             =   3345
         Width           =   1050
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   0
         Left            =   4416
         TabIndex        =   9
         Tag             =   "500"
         Top             =   720
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   1
         Left            =   4416
         TabIndex        =   10
         Tag             =   "100"
         Top             =   1095
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   2
         Left            =   4416
         TabIndex        =   11
         Tag             =   "50"
         Top             =   1470
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   3
         Left            =   4416
         TabIndex        =   12
         Tag             =   "20"
         Top             =   1845
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   4
         Left            =   4416
         TabIndex        =   13
         Tag             =   "10"
         Top             =   2220
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   5
         Left            =   4416
         TabIndex        =   14
         Tag             =   "5"
         Top             =   2595
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   6
         Left            =   4416
         TabIndex        =   15
         Tag             =   "2"
         Top             =   2970
         Width           =   1044
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   7
         Left            =   4416
         TabIndex        =   16
         Tag             =   "1"
         Top             =   3345
         Width           =   1044
      End
      Begin VB.TextBox txtRecv 
         Height          =   288
         Index           =   8
         Left            =   2328
         TabIndex        =   8
         Tag             =   "1"
         Top             =   3720
         Width           =   1050
      End
      Begin VB.TextBox txtRefund 
         Height          =   288
         Index           =   8
         Left            =   4416
         TabIndex        =   17
         Tag             =   "1"
         Top             =   3720
         Width           =   1044
      End
      Begin VB.Label Label12 
         Caption         =   "Amount in Coins   ( in rupees  )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   30
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Rs    1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   3250
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Rs    2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   2920
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Rs    5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   2560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Rs   10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   2200
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Rs   20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   1840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Rs   50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Rs   100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1120
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Rs   500"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   760
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   -240
         X2              =   5784
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Denomination                  Received                           Refund"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   21
         Top             =   336
         Width           =   4932
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Cash Expected  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   38
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmCashIndexOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public WithEvents MyCtrl As ctlCashIndex
Private Function GetTotalAmountRecieved() As Long
Dim I As Integer
On Error GoTo ErrorLine
    For I = 0 To 8
        GetTotalAmountRecieved = GetTotalAmountRecieved + Val(txtRecv(I).Tag) * Val(txtRecv(I).Text)
    Next I
    Exit Function
ErrorLine:
    GetTotalAmountRecieved = 0
    MsgBox Err.Description
End Function


Private Function GetTotalAmountRefunded() As Long
Dim I As Integer
On Error GoTo ErrorLine
    For I = 0 To 8
        GetTotalAmountRefunded = GetTotalAmountRefunded + Val(txtRefund(I).Tag) * Val(txtRefund(I).Text)
    Next I
    Exit Function
ErrorLine:
    GetTotalAmountRefunded = 0
    MsgBox Err.Description
End Function


Private Sub cmdAccept_Click()
    If Val(Label2.Caption) <> Val(Text2(2).Text) Then
        MsgBox "Net amount not equal to expected amount", vbInformation
        Exit Sub
    End If
    gReceivedDenomination.Rs500 = Val(txtRecv(0).Text)
    gReceivedDenomination.Rs100 = Val(txtRecv(1).Text)
    gReceivedDenomination.Rs50 = Val(txtRecv(2).Text)
    gReceivedDenomination.Rs20 = Val(txtRecv(3).Text)
    gReceivedDenomination.Rs10 = Val(txtRecv(4).Text)
    gReceivedDenomination.Rs5 = Val(txtRecv(5).Text)
    gReceivedDenomination.Rs2 = Val(txtRecv(6).Text)
    gReceivedDenomination.Rs1 = Val(txtRecv(7).Text)
    gReceivedDenomination.Coins = Val(txtRecv(8).Text)
    
    gRefundDenomination.Rs500 = Val(txtRefund(0).Text)
    gRefundDenomination.Rs100 = Val(txtRefund(1).Text)
    gRefundDenomination.Rs50 = Val(txtRefund(2).Text)
    gRefundDenomination.Rs20 = Val(txtRefund(3).Text)
    gRefundDenomination.Rs10 = Val(txtRefund(4).Text)
    gRefundDenomination.Rs5 = Val(txtRefund(5).Text)
    gRefundDenomination.Rs2 = Val(txtRefund(6).Text)
    gRefundDenomination.Rs1 = Val(txtRefund(7).Text)
    gRefundDenomination.Coins = Val(txtRefund(8).Text)
    
    gStatus = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    gStatus = False
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCashIndex = Nothing
End Sub


Private Sub Text2_Change(Index As Integer)
    Text2(2).Text = Val(Text2(0).Text) - Val(Text2(1).Text)
End Sub

Private Sub txtRecv_Change(Index As Integer)
    Text2(0).Text = GetTotalAmountRecieved()
End Sub


Private Sub txtRecv_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyBack) Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtRefund_Change(Index As Integer)
    Text2(1).Text = GetTotalAmountRefunded()
End Sub
Private Sub txtRefund_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyBack) Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

