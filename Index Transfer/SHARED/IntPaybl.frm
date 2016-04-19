VERSION 5.00
Begin VB.Form frmIntPayble 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   2040
   ClientTop       =   1995
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   7755
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   5850
      TabIndex        =   12
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   6810
      TabIndex        =   13
      Top             =   5760
      Width           =   855
   End
   Begin VB.PictureBox picOut 
      Height          =   5625
      Left            =   150
      ScaleHeight     =   5565
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   30
      Width           =   7515
      Begin VB.TextBox txtBefore 
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   930
         Width           =   465
      End
      Begin VB.TextBox txtAfter 
         Height          =   285
         Left            =   1500
         ScrollBars      =   1  'Horizontal
         TabIndex        =   10
         Top             =   930
         Width           =   435
      End
      Begin VB.VScrollBar vScr 
         Height          =   5565
         Left            =   7200
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame fraTitle 
         BackColor       =   &H80000009&
         Height          =   555
         Left            =   30
         TabIndex        =   1
         Top             =   -60
         Width           =   7125
         Begin VB.Label lblCustTitle 
            BackColor       =   &H80000009&
            Caption         =   "Name of the Customer"
            Height          =   315
            Left            =   60
            TabIndex        =   2
            Top             =   150
            Width           =   2775
         End
         Begin VB.Label lblBalance 
            BackColor       =   &H80000009&
            Caption         =   "BAlance"
            Height          =   315
            Left            =   6300
            TabIndex        =   4
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label lblIntTitle 
            BackColor       =   &H80000009&
            Caption         =   "Interest Payble"
            Height          =   315
            Left            =   4680
            TabIndex        =   3
            Top             =   120
            Width           =   1425
         End
      End
      Begin VB.Frame fraIN 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   60
         TabIndex        =   5
         Top             =   600
         Width           =   7095
         Begin VB.TextBox txtBox 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2640
            TabIndex        =   9
            Top             =   390
            Width           =   1035
         End
         Begin VB.Label txtAmount 
            BackColor       =   &H80000005&
            Height          =   285
            Index           =   0
            Left            =   4950
            TabIndex        =   7
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label lblCustName 
            BackColor       =   &H80000009&
            Caption         =   "Name Of The Customer"
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   90
            Width           =   4935
         End
         Begin VB.Label txtBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   11
            Top             =   90
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frmIntPayble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Gap As Integer
Public Function LoadContorls(ControlLoopCount As Integer, CtrlGap As Integer)

On Error Resume Next
If ControlLoopCount < 1 Then Exit Function

Dim LoopCount As Integer
Dim Gap As Integer

LoopCount = ControlLoopCount
If CtrlGap > 0 Then
    m_Gap = CtrlGap
Else
    m_Gap = 50
End If

Gap = m_Gap

Dim CustLeft As Integer
Dim CustWidth As Integer
Dim AmountLeft As Integer
Dim AmountWidth As Integer
Dim BalanceLeft As Integer
Dim BalanceWidth As Integer

Dim TopPos As Single
Dim Ht As Single
Dim IndxTab As Integer

LoopCount = 0
Ht = txtAmount(0).Height
TopPos = txtBalance(LoopCount).Top + txtBalance(LoopCount).Height + Gap
IndxTab = txtBalance(LoopCount).TabIndex + 1

CustLeft = lblCustName(0).Left
CustWidth = lblCustName(0).Width

AmountLeft = txtAmount(0).Left
AmountWidth = txtAmount(0).Width

BalanceLeft = txtBalance(0).Left
BalanceWidth = txtBalance(0).Width


For LoopCount = 1 To ControlLoopCount - 1
    Load lblCustName(LoopCount)
    With lblCustName(LoopCount)
        .Top = TopPos
        .Left = CustLeft
        .Width = CustWidth
        .Height = Ht
        .Visible = True
        .TabIndex = IndxTab: IndxTab = IndxTab + 1
    End With
    Load txtAmount(LoopCount)
    
    With txtAmount(LoopCount)
        .Top = TopPos
        .Left = AmountLeft
        .Width = AmountWidth
        .Height = Ht
        .Visible = True
        .TabIndex = IndxTab: IndxTab = IndxTab + 1
    End With
    Load txtBalance(LoopCount)
    With txtBalance(LoopCount)
        .Top = TopPos
        .Left = BalanceLeft
        .Width = BalanceWidth
        .Height = Ht
        .Visible = True
        .TabIndex = IndxTab: IndxTab = IndxTab + 1
    End With
    
    'Now Set the top Position for next control
    TopPos = txtAmount(LoopCount).Top + txtAmount(LoopCount).Height + Gap
Next
'Now Set the Tab Of Text box
'txtBefore.TabIndex = IndxTab
'txtBox.TabIndex = IndxTab + 1
'txtAfter.TabIndex = txtBox.TabIndex + 2


fraIN.Height = TopPos
Dim CtrlsPerPage As Integer
With vScr
    .Min = 0
    .Max = LoopCount
    .SmallChange = 1
    CtrlsPerPage = 14 'picOut.Height / (txtAmount(0).Height + Gap) - 1
    If ControlLoopCount >= CtrlsPerPage And CtrlsPerPage > 1 Then
        .Max = LoopCount - CtrlsPerPage + 2
        .LargeChange = CtrlsPerPage - 1
        .Visible = True
    Else
        .Visible = False
        picOut.Width = picOut.Width - .Width
        Me.Width = Me.Width - .Width
    End If
End With

LoadContorls = True

End Function


Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()
Me.Hide

End Sub


Private Sub Form_Activate()
    Call txtAmount_Click(Val(txtBox.Tag))
    fraIN.ZOrder 0
End Sub



Private Sub Form_Load()
Dim Ctrl As Control
On Error Resume Next
'Now Assign the Kannada fonts to the All controls
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next Ctrl
lblCustTitle.Caption = LoadResString(gLangOffSet + 446)
lblIntTitle.Caption = LoadResString(gLangOffSet + 450) 'Interest Payble
'lblBalance.Caption = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 450) 'Interest Payble
lblBalance.Caption = LoadResString(gLangOffSet + 42) 'Balance

End Sub




Private Sub txtAfter_GotFocus()
'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the next amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)
    If txtNo < txtAmount.Count - 1 Then
    txtNo = txtNo + 1
    Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If

End Sub


Private Sub txtAmount_Change(Index As Integer)

txtBalance(Index).Caption = FormatCurrency(Val(txtBalance(Index).Tag) + _
        Val(txtAmount(Index)))

If Index = txtAmount.Count - 1 Then Exit Sub

'If Val(txtAmount(Index)) > 0 Then
    txtAmount(txtAmount.Count - 1) = Val(txtAmount(txtAmount.Count - 1).Tag) - _
            Val(txtAmount(Index).Tag) + Val(txtAmount(Index))
            
'End If

End Sub

Private Sub txtBefore_GotFocus()
'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the previous amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)

If txtNo > 0 Then
    txtNo = txtNo - 1
    Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If




End Sub


Private Sub txtbox_GotFocus()
Dim TopPos As Single

    TopPos = txtAmount(Val(txtBox.Tag)).Top

On Error Resume Next

If TopPos + fraIN.Top > picOut.Height - fraTitle.Height Then
    vScr.value = vScr.value + 1
    fraIN.Top = fraIN.Top - txtAmount(0).Height + m_Gap
End If

If TopPos - m_Gap + fraIN.Top < fraTitle.Height Then
    vScr.value = vScr.value - 1
'    fraIN.Top = fraIN.Top + txtAmount(0).Height + m_Gap
    fraIN.Top = fraIN.Top + txtAmount(0).Height - 250
End If

Err.Clear

End Sub

Private Sub txtAmount_GotFocus(Index As Integer)
Dim TopPos As Single

TopPos = txtAmount(Index).Top

On Error Resume Next
If TopPos + fraIN.Top > picOut.Height - fraTitle.Height Then
    vScr.value = vScr.value + 1
    fraIN.Top = fraIN.Top - txtAmount(0).Height + m_Gap
End If

If TopPos - m_Gap + fraIN.Top < fraTitle.Height Then
    vScr.value = vScr.value - 1
    fraIN.Top = fraIN.Top + txtAmount(0).Height + m_Gap
End If

Err.Clear

End Sub


Private Sub txtAmount_LostFocus(Index As Integer)
    
If Index = txtAmount.Count - 1 Then Exit Sub

If Me.ActiveControl.Name = vScr.Name Then _
    vScr.TabIndex = txtAmount(Index).TabIndex

If Not CurrencyValidate(txtAmount(Index), True) Then
    txtAmount(Index) = "0.00"
    Exit Sub
Else
    txtAmount(Index) = FormatCurrency(txtAmount(Index))
End If

If Val(txtAmount(Index)) > 0 Then
    txtAmount(txtAmount.Count - 1) = Val(txtAmount(txtAmount.Count - 1).Tag) - _
            Val(txtAmount(Index).Tag) + Val(txtAmount(Index))
    txtAmount(Index).Tag = txtAmount(Index)
    txtAmount(txtAmount.Count - 1).Tag = txtAmount(txtAmount.Count - 1)
    
End If

End Sub




Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode & " DOWN " & Shift
If KeyCode <> vbKeyTab Then Exit Sub

Dim txtNo As Integer
txtNo = Val(txtBox.Tag)
    txtNo = txtNo + IIf(Shift, -1, 1)
    If txtNo < 0 Then Exit Sub
    If txtNo > txtAmount.Count Then Exit Sub
    
    Call txtAmount_Click(txtNo)
    
End Sub

Private Sub txtBox_KeyUp(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode & " UP " & Shift
End Sub


Private Sub txtBox_LostFocus()
    
If Val(txtBox.Tag) = txtAmount.Count - 1 Then Exit Sub

'If Me.ActiveControl.Name = vScr.Name Then _
    vScr.TabIndex = txtAmount(Index).TabIndex

If Not CurrencyValidate(txtBox.Text, True) Then
    txtBox.Text = "0.00"
    Exit Sub
Else
    txtAmount(txtBox.Tag) = FormatCurrency(txtBox.Text)
End If

If Val(txtBox) > 0 Then
    txtAmount(txtAmount.Count - 1) = Val(txtAmount(txtAmount.Count - 1).Tag) - _
            Val(txtAmount(Val(txtBox.Tag)).Tag) + Val(txtAmount(Val(txtBox.Tag)))
    txtAmount(Val(txtBox.Tag)).Tag = txtAmount(Val(txtBox.Tag))
    txtAmount(txtAmount.Count - 1).Tag = txtAmount(txtAmount.Count - 1)
End If

End Sub



Private Sub txtAmount_Click(Index As Integer)
    Call txtBox_LostFocus
    txtBox.Tag = Index
    'txtBox.Text = ""
    txtBox.Top = txtAmount(Index).Top
    txtBox.Left = txtAmount(Index).Left
    txtBox.Text = txtAmount(Index).Caption
    txtBox.ZOrder 0
    txtBox.SetFocus
End Sub

Private Sub txtBox_Change()
    
    If txtBox.Tag = "" Then Exit Sub
    txtAmount(txtBox.Tag) = txtBox.Text
    
End Sub


Private Sub txtDummy_Change()

End Sub



Private Sub vScr_Change()
    If Not IsNumeric(vScr.Tag) Then vScr.Tag = "0"
    If Abs(vScr.Tag - vScr.value) = vScr.SmallChange Then
        fraIN.Top = fraIN.Top + txtBalance(0).Height * (vScr.Tag - vScr.value)
    ElseIf Abs(vScr.Tag - vScr.value) = vScr.LargeChange Then
        fraIN.Top = fraIN.Top + (txtBalance(0).Height + m_Gap) * (picOut.Height / (txtBalance(0).Height + m_Gap) - 1) * IIf(vScr.Tag > vScr.value, 1, -1)
    Else
        fraIN.Top = vScr.value * txtBalance(0).Height * -1 + fraTitle.Height + fraTitle.Top
        If fraIN.Top < -fraIN.Height Then fraIN.Top = picOut.Height - fraIN.Height
    End If
    vScr.Tag = vScr.value
End Sub


