VERSION 5.00
Begin VB.Form frmCashIndex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   2280
   ClientTop       =   1920
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      Left            =   2610
      TabIndex        =   17
      Top             =   5490
      Width           =   1188
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3990
      TabIndex        =   16
      Top             =   5490
      Width           =   1188
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cash flow details "
      Height          =   5280
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   5175
      Begin VB.TextBox txtCashOut 
         Height          =   315
         Index           =   0
         Left            =   3690
         TabIndex        =   12
         Top             =   510
         Width           =   1200
      End
      Begin VB.TextBox txtCashIn 
         Height          =   315
         Index           =   0
         Left            =   1110
         TabIndex        =   3
         Top             =   480
         Width           =   1185
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   3210
         TabIndex        =   15
         Top             =   4770
         Width           =   1188
      End
      Begin VB.TextBox txtExpectedCash 
         Height          =   315
         Left            =   1608
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4815
         Width           =   1116
      End
      Begin VB.TextBox txtTotalOut 
         Height          =   288
         Left            =   3312
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3870
         Width           =   1200
      End
      Begin VB.TextBox txtTotalIn 
         BackColor       =   &H00FFFFFF&
         Height          =   288
         Left            =   792
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3870
         Width           =   1200
      End
      Begin VB.TextBox txtNetAmount 
         Height          =   315
         Left            =   1608
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4395
         Width           =   1116
      End
      Begin VB.Label lblCashIn 
         Caption         =   "Rs. 1000"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   705
      End
      Begin VB.Label lblExpectedCash 
         Alignment       =   1  'Right Justify
         Caption         =   "Expected Amount: "
         Height          =   240
         Left            =   165
         TabIndex        =   8
         Top             =   4830
         Width           =   1380
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   2580
         X2              =   2580
         Y1              =   150
         Y2              =   4182
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   45
         X2              =   5373
         Y1              =   3750
         Y2              =   3750
      End
      Begin VB.Label lblTotalOut 
         Alignment       =   1  'Right Justify
         Caption         =   "Total "
         Height          =   180
         Left            =   2610
         TabIndex        =   13
         Top             =   3915
         Width           =   615
      End
      Begin VB.Label lblTotalIn 
         Alignment       =   1  'Right Justify
         Caption         =   "Total "
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   3900
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   45
         X2              =   5385
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblNetAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount: "
         Height          =   240
         Left            =   30
         TabIndex        =   6
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Returned "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3405
         TabIndex        =   10
         Top             =   150
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Received "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   915
         TabIndex        =   1
         Top             =   150
         Width           =   990
      End
      Begin VB.Label lblCashOut 
         Alignment       =   1  'Right Justify
         Caption         =   "Rs. 1000"
         Height          =   180
         Index           =   0
         Left            =   2580
         TabIndex        =   11
         Top             =   630
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmCashIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OKClicked()
Public Event CancelClicked()

Private Sub ClearControls()

Dim COunt As Integer
For COunt = 0 To txtCashIn.COunt - 1
    txtCashIn(COunt).Text = "0"
    txtCashOut(COunt).Text = "0"
Next COunt

End Sub

'This function returns the string that was typed before it is displayed on the object
'Thus one can check the string and validate it accordingly before it will be displayed
Function PreviewKeyAsciiOLD(txt As Object, Key As Integer) As String
#If JUNK Then
Dim Start As Integer
Dim Length As Integer
Dim Part1 As String, Part2 As String, Part3 As String

    Start = txt.SelStart
    Length = txt.SelLength
   
    If Start < 1 Then
        Part1 = ""
    Else
        Part1 = Left$(txt.Text, Start)
    End If
    
   If (Len(txt.Text) - Start - Length) > 0 Then
        Part3 = Right(txt.Text, Len(txt.Text) - Start - Length)
   Else
        Part3 = ""
   End If
    
   If Key = 22 Then      'Ctrl - V
    Part2 = Clipboard.GetText
   Else
    Part2 = Chr$(Key)
   End If
   
   If Key = 24 Then  ' Ctrl - X
    Part2 = ""
   End If
   
#If OLD Then
'To take care of the Delete Key
    If KeyCod = 46 Then
        Part2 = ""
    End If
#End If
    
   If Key = 3 Then
     PreviewKeyAscii = txt.Text
     Exit Function
   End If
   
   If Key = 8 Then
      Part2 = ""
      If Len(Part1) > 0 Then
      Part1 = Left(Part1, Len(Part1) - 1)
      End If
   End If

PreviewKeyAscii = Part1 & Part2 & Part3

'MsgBox PreviewKeyAscii
#End If
End Function

Private Sub AlignControls()

Dim I As Integer

    'lblCashIn(0).Top = 350
    'txtCashIn(0).Top = 300
    
    lblCashOut(0).Visible = True
    txtCashOut(0).Visible = True
    lblCashIn(0).Visible = True
    txtCashIn(0).Visible = True
    
    lblCashOut(0).Top = lblCashIn(0).Top
    txtCashOut(0).Top = txtCashIn(0).Top
    lblCashOut(0).Height = lblCashIn(0).Height
    txtCashOut(0).Height = txtCashIn(0).Height
    
Dim Cap As Integer
Cap = 1000
    
I = 0
'lblCashOut(0).TabIndex = 4
'txtCashOut(0).TabIndex = 5

Do
    If Cap = 0 Then Exit Do
    If Cap = 1 Then Cap = 0
    If Cap = 2 Then Cap = 1
    If Cap = 5 Then Cap = 2
    If Cap = 10 Then Cap = 5
    If Cap = 20 Then Cap = 10
    If Cap = 50 Then Cap = 20
    If Cap = 100 Then Cap = 50
    If Cap = 500 Then Cap = 100
    If Cap = 1000 Then Cap = 500
    
    I = I + 1
    Load lblCashIn(I)
    Load txtCashIn(I)
    Load lblCashOut(I)
    Load txtCashOut(I)
    
    lblCashOut(I).Visible = True
    txtCashOut(I).Visible = True
    lblCashIn(I).Visible = True
    txtCashIn(I).Visible = True
    
    txtCashIn(I).Top = txtCashIn(I - 1).Top + txtCashIn(I - 1).Height + 40
    txtCashIn(I).Left = txtCashIn(I - 1).Left
    lblCashIn(I).Top = txtCashIn(I).Top 'lblCashIn(i - 1).Top + lblCashIn(i - 1).Height + 150
    lblCashIn(I).Left = lblCashIn(I - 1).Left
    lblCashIn(I).Caption = IIf(Cap, "Rs. " & Cap, "Coins")
    
    'For the Cash Out part
    txtCashOut(I).Top = txtCashIn(I).Top 'txtCashOut(i - 1).Top + txtCashOut(i - 1).Height + 40
    txtCashOut(I).Left = txtCashOut(I - 1).Left
    lblCashOut(I).Top = lblCashIn(I).Top 'lblCashout(- 1).Top + lblCashOut(i - 1).Height + 150
    lblCashOut(I).Left = lblCashOut(I - 1).Left
    lblCashOut(I).Caption = lblCashIn(I).Caption
    
    'now set the tab index
    lblCashOut(I).TabIndex = lblCashOut(I - 1).TabIndex + 2
    txtCashOut(I).TabIndex = lblCashOut(I).TabIndex + 1
    lblCashIn(I).TabIndex = lblCashIn(I - 1).TabIndex + 2
    txtCashIn(I).TabIndex = lblCashIn(I).TabIndex + 1
Loop

'Align the Total Boxes
txtTotalIn.Left = txtCashIn(0).Left
txtTotalOut.Left = txtCashOut(0).Left
lblTotalIn.Left = lblCashIn(0).Left
lblTotalOut.Left = lblCashOut(0).Left

txtTotalOut.TabIndex = txtCashOut(I).TabIndex + 1
lblTotalOut.TabIndex = txtCashOut(I).TabIndex + 1
txtTotalIn.TabIndex = txtCashOut(I).TabIndex + 1
lblTotalIn.TabIndex = txtCashOut(I).TabIndex + 1

lblCashOut(0).Visible = True
txtCashOut(0).Visible = True
    

I = lblCashIn.COunt - 1
With lblCashIn(I)
    Line1.Y1 = .Top + txtCashIn(I).Height + 100
    Line1.Y2 = Line1.Y1  '.Top + .Height + 50
End With
    
txtTotalIn.Top = Line1.Y1 + 100
txtTotalOut.Top = txtTotalIn.Top

lblTotalIn.Top = txtTotalIn.Top
lblTotalOut.Top = txtTotalIn.Top

Line2.Y1 = txtTotalIn.Top + txtTotalIn.Height + 125
Line2.Y2 = txtTotalOut.Top + txtTotalOut.Height + 125
Line3.Y2 = Line2.Y2
txtNetAmount.Top = Line2.Y1 + 100
lblNetAmount.Top = Line2.Y1 + 100
With txtNetAmount
    txtExpectedCash.Top = .Top + .Height + 50
    lblExpectedCash.Top = .Top + .Height + 50
    cmdClear = .Top + .Height + 50
End With
Me.Frame1.Height = txtExpectedCash.Top + txtExpectedCash.Height + 100
cmdCancel.Top = Frame1.Top + Frame1.Height + 100
cmdOK.Top = cmdCancel.Top

Me.Height = cmdOK.Top + cmdOK.Height + 500


End Sub

'Function checks for a valid money value entered

Function CashValidateKeyAsciiOLD(txt As TextBox, Key As Integer) As Boolean
#If JUNK Then
Dim TextPrev As String
Dim Pos As Integer
Dim COunt As Integer
Dim CHar As String * 1
Dim AscVal As Integer
CashValidateKeyAscii = False

#If JUNK Then
'First check for the valid key set
    If Key < Asc("0") Or Key > Asc("9") Then
        If Key <> Asc(".") And Key <> 8 Then GoTo lastline
    End If
#End If

'Preview the text
    TextPrev = PreviewKeyAscii(txt, Key)
    
'Now check all the characters if there is any invalid character
    For COunt = 1 To Len(TextPrev)
        AscVal = Asc(Right(Left(TextPrev, COunt), 1))
        If AscVal < Asc("0") Or AscVal > Asc("9") Then
            If AscVal <> Asc(".") Then GoTo lastline
        End If
    Next COunt
    
'Now check if there are more than two decimals
    Pos = InStr(1, TextPrev, ".", vbBinaryCompare)
    If Pos <> 0 Then 'There is a dot(.)
        If Len(Right(TextPrev, Len(TextPrev) - Pos)) > 2 Then
            GoTo lastline
        End If
    End If
'Check if the left part of the decimal number is within range of currency
    If Len(Left(TextPrev, Len(TextPrev) - Pos)) > 14 Then 'Gosh there is a lot of money here !!!
        GoTo lastline
    End If
        
    
CashValidateKeyAscii = True
Exit Function
lastline:
Key = 0
CashValidateKeyAscii = False
#End If
End Function

Private Sub cmdCancel_Click()
RaiseEvent CancelClicked
Unload Me

End Sub

Private Sub cmdClear_Click()

Call ClearControls
End Sub

Private Sub cmdOK_Click()
RaiseEvent OKClicked

Unload Me
End Sub

Private Sub Form_Load()
    Call AlignControls
    Call ClearControls
    cmdOK.Enabled = False
End Sub

'
'   This function allows only the chars present in the ValidSet passed to it.
'   AllowOtherCase allows the other case also.
'   Eg. If your valid set contains A and you want to allow "a" also,
'   then pass AllowOtherCase as TRUE
'
Function AllowKeyAsciiOLD(txt As Object, ValidSet As String, Key As Integer, Optional AllowOtherCase As Boolean) As Integer
#If JUNK Then
Dim COunt As Integer, I As Integer
Dim Flag As Boolean
Dim TempBuf As String

    ReDim InvalidArr(0)
    
    If Not IsMissing(AllowOtherCase) Then
        If AllowOtherCase Then       'We have to consider the case
            ValidSet$ = UCase(ValidSet$) & LCase(ValidSet)
        End If
    End If

    Flag = 0
    For COunt = 1 To Len(ValidSet)
        If Key = Asc(Mid(ValidSet, COunt, 1)) Then
            Flag = True
        End If
    Next COunt
    

    If Key = 22 Then
        TempBuf = Clipboard.GetText
        For COunt = 1 To Len(TempBuf)
            Flag = False
            For I = 1 To Len(ValidSet)
                If Asc(Mid(TempBuf, COunt, 1)) = Asc(Mid(ValidSet, I, 1)) Then
                    Flag = True
                    Exit For
                End If
            Next I
           If Flag = False Then
                Exit For
           End If

        Next COunt
    End If
    
    If Not Flag Then
        Key = 0
    End If
#End If
End Function

Private Sub txtCashIn_Change(Index As Integer)

Dim Factor As Integer
Dim TotalCashIn As Currency
Dim COunt As Integer
On Error GoTo ErrLine

TotalCashIn = 0
For COunt = 0 To txtCashIn.COunt - 1
    Select Case COunt
        Case 0:
            Factor = 1000
        Case 1:
            Factor = 500
        Case 2:
            Factor = 100
        Case 3:
            Factor = 50
        Case 4:
            Factor = 20
        Case 5:
            Factor = 10
        Case 6:
            Factor = 5
        Case 7:
            Factor = 2
        Case 8, 9:
            Factor = 1
    End Select
    
    TotalCashIn = TotalCashIn + Factor * CCur(Val(txtCashIn(COunt).Text))
Next COunt
    
    txtTotalIn.Text = Format(TotalCashIn, "#############0.00")
    
    Exit Sub
ErrLine:
    If Err.Number = 6 Then MsgBox "Overflow"
    
End Sub

Private Sub txtCashIn_GotFocus(Index As Integer)
txtCashIn(Index).SelStart = 0
txtCashIn(Index).SelLength = Len(txtCashIn(Index).Text)
End Sub
Private Sub txtCashIn_KeyPress(Index As Integer, KeyAscii As Integer)
Dim RetBool As Boolean
Dim Factor As Integer

'RetBool = CashValidateKeyAscii(txtCashIn(Index), KeyAscii)
RetBool = AllowKeyAscii(txtCashIn(Index), "1234567890" & Chr(8), KeyAscii, False)
    
End Sub
Private Sub txtCashIn_LostFocus(Index As Integer)

If Trim$(txtCashIn(Index).Text) = "" Then txtCashIn(Index).Text = "0"

If Val(txtTotalIn) > Val(txtExpectedCash) Then txtCashOut(0).SetFocus
If Val(txtTotalIn) = Val(txtExpectedCash) Then cmdOK.SetFocus

'txtCashIn(Index).Text = Format(txtCashIn(Index).Text, "#############0.00")

End Sub

Private Sub txtCashOut_Change(Index As Integer)
Dim Factor As Integer
Dim TotalCashOut As Currency
Dim COunt As Integer
On Error GoTo ErrLine

TotalCashOut = 0
For COunt = 0 To txtCashOut.COunt - 1
    Select Case COunt
        Case 0:
            Factor = 1000
        Case 1:
            Factor = 500
        Case 2:
            Factor = 100
        Case 3:
            Factor = 50
        Case 4:
            Factor = 20
        Case 5:
            Factor = 10
        Case 6:
            Factor = 5
        Case 7:
            Factor = 2
        Case 8, 9:
            Factor = 1
    End Select
    
    TotalCashOut = TotalCashOut + Factor * CCur(Val(txtCashOut(COunt).Text))
Next COunt

    txtTotalOut.Text = Format(TotalCashOut, "###############0.00")
    Exit Sub
ErrLine:
    If Err.Number = 6 Then
        MsgBox "Overflow"
    End If
End Sub
Private Sub txtCashOut_GotFocus(Index As Integer)

txtCashOut(Index).SelStart = 0
txtCashOut(Index).SelLength = Len(txtCashOut(Index).Text)

End Sub


Private Sub txtCashOut_KeyPress(Index As Integer, KeyAscii As Integer)
Dim RetBool As Boolean
Dim Factor As Integer
'RetBool = CashValidateKeyAscii(txtCashOut(Index), KeyAscii)
RetBool = AllowKeyAscii(txtCashOut(Index), "1234567890" & Chr(8), KeyAscii, False)

End Sub

Private Sub txtCashOut_LostFocus(Index As Integer)

If Trim$(txtCashOut(Index).Text) = "" Then txtCashOut(Index).Text = "0"

If Val(txtNetAmount) < Val(txtExpectedCash) Then txtCashIn(0).SetFocus
If Val(txtNetAmount) = Val(txtExpectedCash) Then cmdOK.SetFocus


End Sub

Private Sub txtNetAmount_Change()

If Val(txtNetAmount.Text) <> Val(txtExpectedCash.Text) Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If

End Sub

Private Sub txtTotalIn_Change()
'Change the Net Amount
On Error GoTo ErrLine
txtNetAmount.Text = Format(CCur(Val(txtTotalIn.Text)) - CCur(Val(txtTotalOut.Text)), "#############0.00")

Exit Sub
ErrLine:
If Err.Number = 6 Then MsgBox "Overflow"

End Sub

Private Sub txtTotalOut_Change()
'Change the Net Amount
On Error GoTo ErrLine
txtNetAmount.Text = Format(CCur(Val(txtTotalIn.Text)) - CCur(Val(txtTotalOut.Text)), "#############0.00")
Exit Sub

ErrLine:

If Err.Number = 6 Then MsgBox "Overflow"

End Sub

