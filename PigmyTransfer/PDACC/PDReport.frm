VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmPDReport 
   Caption         =   "Pigmy Deposit Reports ..."
   ClientHeight    =   6000
   ClientLeft      =   1725
   ClientTop       =   1845
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1320
      TabIndex        =   1
      Top             =   5280
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3780
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.CheckBox chkAgent 
         Caption         =   "Show Agent Name"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4785
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8440
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   " Report Title "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   5
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "frmPDReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_FromIndianDate As String
Dim m_ToIndianDate As String
Dim m_FromDate As Date
Private m_ToDate As Date
Private m_AccID As Long
Private m_AgentID As Integer

Dim m_FromAmt As Currency
Dim m_ToAmt As Currency
Dim m_Gender As Integer
Dim m_Caste As String
Dim m_Place As String
Dim m_AgentNameShow As Boolean
Dim m_ReportOrder As wis_ReportOrder
Dim m_ReportType As wis_PDReports
Dim m_AccGroup As Integer


Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)

Private WithEvents m_grdPrint As GridPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private WithEvents m_frmCancel As frmCancel
Attribute m_frmCancel.VB_VarHelpID = -1



Public Property Let AccountGroup(NewValue As Integer)
    m_AccGroup = NewValue
End Property


Public Property Let AgentID(NewValue As Integer)
    m_AgentID = NewValue
End Property

Public Property Let Caste(NewCaste As String)
m_Caste = NewCaste
End Property

Public Property Let DisplayAgentName(NewValue As Boolean)
    m_AgentNameShow = NewValue
End Property

Public Property Let FromAmount(newAmount As Currency)
    m_FromAmt = newAmount
End Property

Public Property Let Gender(NewGender As Integer)
    m_Gender = NewGender
End Property

Public Property Let Place(NewPlace As String)
    m_Place = NewPlace
End Property

Public Property Let ReportOrder(NewRP As wis_ReportOrder)
    m_ReportOrder = NewRP
End Property

Public Property Let ReportType(newRT As wis_PDReports)
    m_ReportType = newRT
End Property

Public Property Let ToAmount(newAmount As Currency)
    m_ToAmt = newAmount
End Property

Public Property Let ToIndianDate(NewStrdate As String)
    If Not DateValidate(NewStrdate, "/", True) Then
        Err.Raise 5002, , "Invalid Date"
        Exit Property
    End If
    m_ToIndianDate = NewStrdate
    m_ToDate = GetSysFormatDate(NewStrdate)
    'm_ToIndianDate = GetAppFormatDate(m_ToDate)
End Property

Public Property Let FromIndianDate(NewStrdate As String)
    If Not DateValidate(NewStrdate, "/", True) Then Exit Property
    m_FromIndianDate = NewStrdate
    m_FromDate = GetSysFormatDate(NewStrdate)
    'm_FromIndianDate = GetAppFormatDate(m_FromDate)
    
End Property

Private Sub ShowAgentTransaction()

Dim SQLStmt As String
Dim Rst As Recordset
Dim Count As Integer
Dim TransType As wisTransactionTypes
Dim PigmyCommission As Single
    
    RaiseEvent Processing("Reading & Verifying the data ", 0)
    TransType = wDeposit
    gDbTrans.SQLStmt = "Select Sum(Amount) as TotalAmount,AgentId,TransDate  " & _
            " From AgentTrans Where TransDate >= #" & m_FromDate & "# " & _
            " and TransDate <= #" & m_ToDate & "# " & _
            " Group By AgentId, TransDate "
    
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
    chkAgent.Enabled = False
Call InitGrid


RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Dim l_AgentID As Integer
Dim AgentAmount As Currency
Dim TotalAmount As Currency
Dim PigmyAmount As Currency

Dim SetupClass As New clsSetup
PigmyCommission = SetupClass.ReadSetupValue("PDAcc", "PigmyCommission", "03")
If PigmyCommission > 1 Then PigmyCommission = PigmyCommission / 100
Set SetupClass = Nothing

Dim rowNo As Long
rowNo = grd.Row
While Not Rst.EOF
    With grd
        If .Rows <= rowNo + 2 Then .Rows = .Rows + 2
        rowNo = rowNo + 1
        If l_AgentID <> FormatField(Rst("AgentID")) Then
            .Row = rowNo
            If .Rows <= rowNo + 1 Then .Rows = .Rows + 1
            If l_AgentID <> 0 Then
                .Col = 0: .Text = LoadResString(gLangOffSet + 304) '"Sub Total"
                .CellAlignment = 7: .CellFontBold = True
                .Col = 2: .Text = FormatCurrency(PigmyAmount)
                .CellAlignment = 7: .CellFontBold = True
                .Col = 3: .Text = FormatCurrency(AgentAmount)
                .CellAlignment = 7: .CellFontBold = True
                TotalAmount = TotalAmount + PigmyAmount: PigmyAmount = 0
                AgentAmount = 0
            Else
                .Row = 0: rowNo = 0
            End If
            If .Rows = rowNo + 1 Then .Rows = .Rows + 2
            rowNo = rowNo + 1
            .Row = rowNo
            l_AgentID = Val(FormatField(Rst("AgentId")))
            .Col = 0: .Text = GetAgentName(CLng(l_AgentID))
            .CellFontBold = True
        End If
        .TextMatrix(rowNo, 1) = FormatField(Rst("TransDate"))
        .TextMatrix(rowNo, 2) = FormatField(Rst("TotalAmount"))
        .TextMatrix(rowNo, 3) = FormatCurrency(FormatField(Rst("TotalAmount")) * PigmyCommission)
    End With
    
    AgentAmount = AgentAmount + Val(grd.Text)
    PigmyAmount = PigmyAmount + FormatField(Rst("totalAmount"))
    
NextRecord:
    Rst.MoveNext
    
    DoEvents
    Me.Refresh

    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
Wend

With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .Col = 0: .Text = LoadResString(gLangOffSet + 52)
    .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(PigmyAmount): .CellFontBold = True: .CellAlignment = 7
    .Col = 3: .Text = FormatCurrency(AgentAmount): .CellFontBold = True: .CellAlignment = 7
    TotalAmount = TotalAmount + PigmyAmount
        
  If PigmyAmount <> TotalAmount Then
    If .Rows <= .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 2
    .Col = 0: .Text = LoadResString(gLangOffSet + 286) '"Grand Total"
    .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(TotalAmount): .CellAlignment = 7
    .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalAmount * PigmyCommission): .CellAlignment = 7
    .CellFontBold = True
  End If
End With

End Sub
Private Sub InitGrid()
gCancel = 0

Dim Count As Integer
Dim ColWid As Single
Dim I As Integer
    'ColWid = (grd.Width - 200) / grd.Cols + 1
With grd
    .Clear
    .Rows = 20
    .Cols = 2
    .FixedCols = 0
End With
        
On Error Resume Next
If m_ReportType = repPDMonTrans Then
    
    With grd
        .Cols = 5
        'If chkAgent.Value = vbChecked Then .Cols = .Cols + 1
        Count = 0
        .Row = 0
        .FixedRows = 1
        .FixedCols = 1
        .Col = 0: .Text = LoadResString(gLangOffSet + 33) '"sL No"
        .Col = 1: .Text = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35)
        .Col = 2: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)  '"Account No"
        .Col = 3: .Text = LoadResString(gLangOffSet + 35) '"Name"
        Dim TransDate As Date
        TransDate = m_FromDate
        Do
            .Cols = .Cols + 1
            .Col = .Col + 1
            .Text = GetMonthString(Month(TransDate))
            TransDate = DateAdd("m", 1, TransDate)
            If TransDate > m_ToDate Then Exit Do
        Loop
    End With
    
    GoTo BoldLine
End If

If m_ReportType = repPDBalance Then
    With grd
        .Cols = 4
        If chkAgent.Value = vbChecked Then .Cols = .Cols + 1
        Count = 0
        .Row = 0
        .FixedRows = 1
        .FixedCols = 1
        .Col = Count: .Text = LoadResString(gLangOffSet + 33): Count = Count + 1 '"sL No"
        .Col = Count: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60): Count = Count + 1 '"Account No"
        .Col = Count: .Text = LoadResString(gLangOffSet + 35): Count = Count + 1 '"Name"
        If m_AgentNameShow Then
            .Col = Count
            .Text = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35): Count = Count + 1  '"Agent Name"
        End If
        .Col = Count: .Text = LoadResString(gLangOffSet + 42): Count = Count + 1    '"Balance"
    End With
    GoTo BoldLine
End If
    
If m_ReportType = repPDLedger Then
    With grd
        .Clear
        .Cols = 6: .Rows = 10
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = LoadResString(gLangOffSet + 33)  '"Slno":
        .Col = 1: .Text = LoadResString(gLangOffSet + 37)  '"Date":
        .Col = 2: .Text = LoadResString(gLangOffSet + 284) 'Opening balance
        .Col = 3: .Text = LoadResString(gLangOffSet + 271) 'Withdraw
        .Col = 4: .Text = LoadResString(gLangOffSet + 272) '"Repayment"
        .Col = 5: .Text = LoadResString(gLangOffSet + 285) '"Closing balance"
    End With
    GoTo BoldLine
End If

If m_ReportType = repPDDayBook Then
    With grd
        .Clear
        .Cols = 10
        .MergeCells = flexMergeFree
        .FixedCols = 1
        .FixedRows = 2
        Dim TmpStr As String
        .Row = 0: Count = 0
        .MergeRow(0) = True
        .Col = Count: .Text = LoadResString(gLangOffSet + 33): Count = Count + 1 ' "Sl NO"
        .Col = Count: .Text = LoadResString(gLangOffSet + 37): Count = Count + 1 ' "Date"
        .Col = Count: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60): Count = Count + 1 '"Acc NO"
        .Col = Count: .Text = LoadResString(gLangOffSet + 35): Count = Count + 1 '"Name"
        If chkAgent Then
            .Cols = .Cols + 1
            .Col = Count: .Text = LoadResString(gLangOffSet + 330) & " " & _
                LoadResString(gLangOffSet + 35): Count = Count + 1 '"Agent Name"
        End If
        .Col = Count: .Text = LoadResString(gLangOffSet + 271): Count = Count + 1 '"Deposit"
        .Col = Count: .Text = LoadResString(gLangOffSet + 271): Count = Count + 1 '"Deposit"
        .Col = Count: .Text = LoadResString(gLangOffSet + 272): Count = Count + 1  '"Payment"
        .Col = Count: .Text = LoadResString(gLangOffSet + 272): Count = Count + 1  '"Payment"
        .Col = Count: .Text = LoadResString(gLangOffSet + 274): Count = Count + 1 '"Interest"
        .Col = Count: .Text = LoadResString(gLangOffSet + 274): Count = Count + 1 '"Interest"
        .Row = 1
        .MergeRow(2) = True
        For Count = 0 To .Cols - 1
            .MergeCol(Count) = True
            .Row = 0
            .Col = Count
            .CellAlignment = 4: .CellFontBold = True
            TmpStr = .Text
            .Row = 1
            .Text = TmpStr
            .Col = Count
            .CellAlignment = 4: .CellFontBold = True
        Next
        
        I = 0: .Row = 1
        For Count = .Cols - 1 To .Cols - 6 Step -1
            I = I + 1
            .Col = Count
            .MergeCol(Count) = False
            .Text = LoadResString(gLangOffSet + 269 + I Mod 2)
        Next
    End With
    GoTo BoldLine
End If

If m_ReportType = repPDCashBook Then
    With grd
        .Clear
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0: Count = 0
        .MergeRow(0) = True
        .Col = 0: .Text = LoadResString(gLangOffSet + 33)   ' "Sl NO"
        .Col = 1: .Text = LoadResString(gLangOffSet + 37)   ' "Date"
        .Col = 2: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)   '"Acc NO"
        .Col = 3: .Text = LoadResString(gLangOffSet + 35)   '"Name"
        .Col = 4: .Text = LoadResString(gLangOffSet + 41):  'Voucher No
        .Col = 5: .Text = LoadResString(gLangOffSet + 271)  '"Deposit"
        .Col = 6: .Text = LoadResString(gLangOffSet + 272)  '"Payment"
        .Row = 1
        .MergeRow(2) = True
        For Count = 0 To .Cols - 1
            .MergeCol(Count) = True
            .Row = 0
        Next
    End With
    
    GoTo BoldLine
End If
    
    
If m_ReportType = repPDAccClose Then
    With grd
        .Cols = 5
        .FixedCols = 1
        .Row = 0: Count = 0
        .Col = Count: .Text = LoadResString(gLangOffSet + 33): Count = Count + 1 '"Sl No"
        .Col = Count: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60): Count = Count + 1 '"AccNo"
        .Col = Count: .Text = LoadResString(gLangOffSet + 35): Count = Count + 1 '"Name"
        If chkAgent.Value = vbChecked Then
            .Cols = .Cols + 1
            .Col = Count: .Text = LoadResString(gLangOffSet + 330) & _
                " " & LoadResString(gLangOffSet + 35): Count = Count + 1 '"AgentName"
        End If
        .Col = Count: .Text = LoadResString(gLangOffSet + 282): Count = Count + 1 '"Closed Date"
        .Col = Count: .Text = LoadResString(gLangOffSet + 292): Count = Count + 1 '"MaturedAmount"
        .Row = 0
    End With
    
    GoTo BoldLine

End If
    
If m_ReportType = repPDAccOpen Then
    With grd
        .Rows = 25
        .Cols = 4
        If chkAgent Then .Cols = .Cols + 1
        .FixedCols = 0
        .WordWrap = True
        .Row = 0: Count = 0
        .Col = Count: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60): Count = Count + 1: .ColWidth(Count) = .Width / .Cols  '"AccNo"
        .Col = Count: .Text = LoadResString(gLangOffSet + 35): Count = Count + 1: .ColWidth(Count) = .Width / .Cols '"Name"
        If chkAgent.Value = vbChecked Then
            .Col = Count: .Text = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35): Count = Count + 1: .ColWidth(Count) = .Width / .Cols '"Agent Name"
        End If
        .Col = Count: .Text = LoadResString(gLangOffSet + 281): Count = Count + 1: .ColWidth(Count) = .Width / .Cols '"CreateDate"
        .Col = Count: .Text = LoadResString(gLangOffSet + 226): Count = Count + 1: .ColWidth(Count) = .Width / .Cols '"Deposited Amount"
    End With
    GoTo BoldLine
End If
        
If m_ReportType = repPDAgentTrans Then
    With grd
        .Cols = 4
        .Row = 0
        .Col = 0: .Text = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35) '"Agent Name"
        .Col = 1: .Text = LoadResString(gLangOffSet + 38) + LoadResString(gLangOffSet + 37) '"Transaction Date"
        .Col = 2: .Text = LoadResString(gLangOffSet + 40)   '"Amount Collected"
        .Col = 3: .Text = LoadResString(gLangOffSet + 328)   '"Pigmy Commission"
    End With
    GoTo BoldLine
End If

If m_ReportType = repPDMonBal Then
    With grd
        .Clear
        .Rows = 5: .Cols = 4
        .FixedRows = 2: .FixedCols = 1
        .Cols = 4 + DateDiff("M", m_FromDate, m_ToDate) * 2
        If .Cols = 4 Then .Cols = 6
        .Row = 0
        .Col = 0: .Text = LoadResString(gLangOffSet + 33)
        .Col = 1: .Text = "Agent ID"
        .Col = 2: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)
        .Col = 3: .Text = LoadResString(gLangOffSet + 35)
        .Row = 1
        .Col = 0: .Text = LoadResString(gLangOffSet + 33)
        .Col = 1: .Text = "Agent ID"
        .Col = 2: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)
        .Col = 3: .Text = LoadResString(gLangOffSet + 35)
        Count = Month(m_FromDate)
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        Do
            .Col = .Col + 1
            .Row = 0: .Text = GetMonthString(Count)
            .Row = 1: .Text = LoadResString(gLangOffSet + 271) 'Deposit
            .Col = .Col + 1
            
            .Row = 0: .Text = GetMonthString(Count)
            .Row = 1: .Text = LoadResString(gLangOffSet + 289) 'With draw
            If .Col = .Cols - 1 Then Exit Do
            Count = Count + 1
        Loop
    End With
End If
    
BoldLine:

With grd
    .Row = 0
    Do
        If .Row = .FixedRows Then Exit Do
        For Count = 0 To .Cols - 1
            .Col = Count
            .CellAlignment = 4
            .CellFontBold = True
        Next Count
        .Row = .Row + 1
    Loop
End With

    Exit Sub

ExitLine:

With grd
    ColWid = 0
    For Count = 0 To .Cols - 2
        ColWid = ColWid + .ColWidth(Count)
        '.CellFontBold = True
    Next Count
    .ColWidth(grd.Cols - 1) = .Width - ColWid - TextWidth(.ScrollBars) - 250
End With

End Sub

Private Sub MaturedDeposits()
Dim Count As Integer
Dim Rst As Recordset
'
RaiseEvent Processing("Reading & Verifying the records ", 0)

Dim strClause As String
If m_FromAmt > 0 Then strClause = " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then strClause = strClause & " And Balance < " & m_ToAmt
If Len(m_Place) Then strClause = strClause & " And B.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And B.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And B.Gender= " & m_Gender
If m_AgentID Then strClause = strClause & " And A.AgentID = " & m_AgentID

gDbTrans.SQLStmt = "Select AccId,AccNum, AgentId,CreateDate,MaturityDate," & _
    " ClosedDate,RateOfInterest, B.Name from PDMaster A Inner join " & _
    " QryName B On A.CustomerId= B.CustomerId " & _
    " where MaturityDate Between #" & m_FromDate & "# " & _
    " and #" & m_ToDate & "# " & strClause

If chkAgent.Value Then
        
    gDbTrans.SQLStmt = "Select AccId,A.AccNum,A.AgentId,A.CreateDate,MaturityDate," & _
        " A.ClosedDate,RateOfInterest, B.Name, C.Name  as AgentName " & _
        " From QryName B Inner join (PDMaster A " & _
            " Inner join (UserTab D inner join QryName C " & _
        " On C.CustomerId = D.CustomerId )On A.AgentId = D.UserId) " & _
        " On A.CustomerId= B.CustomerId " & _
        " Where MaturityDate Between #" & m_FromDate & "#  " & _
        " And #" & m_ToDate & "# " & strClause
End If

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub


'initialise the grid.
With grd
    Count = 0
    .Clear
    .Rows = 2: .Rows = 25
    .Cols = 5
    If chkAgent Then .Cols = .Cols + 1
    .Row = 0
    .FixedCols = 0
    .Row = 0
    .Col = Count: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60): Count = Count + 1  ' "Account No"
    .Col = Count: .Text = LoadResString(gLangOffSet + 35): Count = Count + 1 '"Name"
    If chkAgent.Value = vbChecked Then
        .Col = Count: .Text = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35)
        Count = Count + 1 '"Agent Name "
    End If
    .Col = Count: .Text = LoadResString(gLangOffSet + 291): Count = Count + 1 ' "Maturity Date"
    .Col = Count: .Text = LoadResString(gLangOffSet + 186): Count = Count + 1 '"RateOfInterest"
    .Col = Count: .Text = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 40): Count = Count + 1 '"Deposited Amount"
    .Row = 0
    For Count = 0 To .Cols - 1
        .Col = Count
        .CellAlignment = 4: .CellFontBold = True
    Next
End With


Dim SecondRst As Recordset
Dim Days As Integer
Dim DepAmt As Currency, MatAmt As Currency
Dim Interest As Double
Dim DepTotal As Currency, MatTotal As Currency
Dim DepDate As String, MatDate As String

    
    RaiseEvent Initialise(0, Rst.RecordCount)
    RaiseEvent Processing("Aligning  the data ", 0)

Dim rowNo As Long
grd.Row = 0

While Not Rst.EOF
    gDbTrans.SQLStmt = "Select Sum(Amount) as TotalAmount From PDTrans " & _
                " Where AccId = " & FormatField(Rst("Accid"))
    If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo NextRecord
    With grd
        'Set next row
        If .Rows <= rowNo + 1 Then .Rows = .Rows + 1
        rowNo = rowNo + 1: Count = 0
        DepDate = FormatField(Rst("CreateDate"))
        MatDate = FormatField(Rst("MaturityDate"))
        Days = WisDateDiff(DepDate, MatDate)
        DepAmt = Val(FormatField(SecondRst("TotalAmount")))
        Interest = Val(FormatField(Rst("RateOfInterest")))
        MatAmt = FormatCurrency(DepAmt + ComputePDInterest(DepAmt, Interest))
        MatTotal = MatTotal + MatAmt
        DepTotal = DepTotal + DepAmt
        .Col = Count: .TextMatrix(rowNo, Count) = FormatField(Rst("AccNUM")): Count = Count + 1
        .Col = Count: .TextMatrix(rowNo, Count) = FormatField(Rst("Name")): Count = Count + 1
        If chkAgent.Value = vbChecked Then
            .Col = Count: .TextMatrix(rowNo, Count) = FormatField(Rst("AgentName")) ''GetAgentName(FormatField(Rst("UserId")))
            Count = Count + 1
        End If
        .Col = Count: .TextMatrix(rowNo, Count) = MatDate: Count = Count + 1
        .Col = Count: .TextMatrix(rowNo, Count) = Interest: Count = Count + 1
        .Col = Count: .TextMatrix(rowNo, Count) = FormatCurrency(DepAmt): Count = Count + 1
    End With
    
NextRecord:
       
    DoEvents
    Me.Refresh
    If gCancel Then Rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext

Wend

'Set last
With grd
    .Row = rowNo
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    
    .Col = 1: .Text = LoadResString(gLangOffSet + 52) '"Totals"
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(DepTotal)
    .CellAlignment = 7: .CellFontBold = True
End With
    
lblReportTitle.Caption = LoadResString(gLangOffSet + 72) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
    
   
End Sub

'
Private Sub ShowDepositBalances()
Dim I As Integer
Dim Rst As Recordset
Dim SQLStmt As String
Dim StrAmount As String

'
RaiseEvent Processing("Reading & Verifying the records ", 0)

SQLStmt = "Select Max(TransId) AS MaxTransID, A.AccID" & _
    " From PDTrans B Inner Join PDMaster A On B.AccId = A.AccId " & _
    " Where TransDate <= #" & m_ToDate & "#" & _
    " GROUP BY A.AccID"
     
gDbTrans.SQLStmt = SQLStmt
If Not gDbTrans.CreateView("QryTemp") Then Exit Sub

SQLStmt = "Select  B.Balance, A.AgentId, A.AccID,A.AccNum,A.CustomerId, Name " & _
    " From QryName C Inner join (PDMaster A inner join " & _
    " (PDtrans B Inner join QryTemp D ON B.TransId = D.MaxTransID AND D.AccID = B.AccID )" & _
        " On A.AccID = B.AccId )" & _
    " ON C.CustomerId = A.CustomerId "

StrAmount = ""
If m_FromAmt > 0 Then StrAmount = " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then StrAmount = StrAmount & " And Balance < " & m_ToAmt
If Len(m_Place) Then StrAmount = StrAmount & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then StrAmount = StrAmount & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then StrAmount = StrAmount & " And C.Gender= " & m_Gender
If m_AccGroup Then StrAmount = StrAmount & " And AccGroupID = " & m_AccGroup
If m_AgentID > 0 Then StrAmount = StrAmount & " And A.AgentID = " & m_AgentID

If Len(StrAmount) Then
    StrAmount = " WHERE " & Mid(Trim$(StrAmount), 4)
End If

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SQLStmt = SQLStmt & StrAmount & " Order by A.AgentID,A.AccNum"
Else
    gDbTrans.SQLStmt = SQLStmt & StrAmount & " Order by A.AgentID,C.IsciName"
End If

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
RaiseEvent Initialise(0, Rst.RecordCount + 1)
RaiseEvent Processing("Aligning the data ", 0)

Dim TotalAmount As Currency
Dim AgentID As Long
Dim AgentName As String
Dim Total As Currency

Call InitGrid
Dim SlNo As Integer

grd.Row = 0
Rst.MoveFirst
AgentID = Val(FormatField(Rst("AgentID")))
AgentName = GetAgentName(AgentID)
I = 0
If m_AgentNameShow Then I = 1
Dim rowNo As Long

While Not Rst.EOF
    With grd
        If AgentID <> Val(FormatField(Rst("AgentID"))) And TotalAmount > 0 Then
            If .Rows <= rowNo + 2 Then .Rows = .Rows + 2
            rowNo = rowNo + 1
            .Row = rowNo: .Col = .Cols - 1
            .Text = FormatCurrency(TotalAmount): .CellAlignment = 7: .CellFontBold = True
            Total = TotalAmount
            .Col = .Cols - 2
            .Text = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 42)
            .CellFontBold = True
            AgentID = Val(FormatField(Rst("AgentID"))): TotalAmount = 0
            AgentName = GetAgentName(AgentID)
        End If
        
        If FormatField(Rst("Balance")) = 0 Then GoTo NextRecord
        If .Rows = rowNo + 2 Then .Rows = .Rows + 2
        rowNo = rowNo + 1
        SlNo = SlNo + 1
        .TextMatrix(rowNo, 0) = Format(SlNo, "00")
        .TextMatrix(rowNo, 1) = FormatField(Rst("AccNum"))
        .TextMatrix(rowNo, 2) = FormatField(Rst("Name"))
        If chkAgent.Value = vbChecked Then
          .TextMatrix(rowNo, 3) = AgentName
        End If
        .TextMatrix(rowNo, 3 + I) = FormatField(Rst("Balance"))
    End With
    TotalAmount = TotalAmount + Val(FormatField(Rst("Balance")))

NextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then Rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext

Wend
    
With grd
    rowNo = rowNo + 2
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    .Row = rowNo
    .Col = 2: .Text = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 42)    ' "Totals Balance "
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(TotalAmount)
    .CellAlignment = 7: .CellFontBold = True
    .Row = .Row: .Text = FormatCurrency(TotalAmount + Total)
    .CellAlignment = 7: .CellFontBold = True
End With

lblReportTitle.Caption = LoadResString(gLangOffSet + 70)

End Sub

Private Sub ShowDepositsOpened()

Dim Dt As Boolean
Dim Amt As Boolean
Dim AmtStr As String
Dim Rst As Recordset
Dim Total As Currency
Dim SqlStr As String
'
RaiseEvent Processing("Reading & Verifying the records ", 0)

Dim strClause As String
If Len(m_Place) Then strClause = strClause & " And B.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And B.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And B.Gender= " & m_Gender
If m_AccGroup Then strClause = strClause & " And AccGroupID = " & m_AccGroup
If m_AgentID Then strClause = strClause & " And A.AgentID= " & m_AgentID

'Build the FINAL SQL
SqlStr = " Select AccId,AgentID,CreateDate,MaturityDate, " & _
    " Closeddate,RateOfInterest, Name" & _
    " From PDMaster A Inner join QryName B " & _
    " ON B.CustomerID = A.CustomerID " & _
    " where CreateDate <= #" & m_ToDate & "#" & _
    " And CreateDate >= #" & m_FromDate & "# " & strClause & _
    " order by val(AccNum)"
 
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
        
Dim Count As Integer
RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Alignign the data ", 0)

Call InitGrid
Dim AccID As Long
Dim Amount As Currency
Dim SecondRst As Recordset
Dim rowNo As Long
rowNo = grd.Row
While Not Rst.EOF
    With grd
        gDbTrans.SQLStmt = "Select Sum(Amount) as TotalAmount " & _
            " From PDTrans Where AccId = " & FormatField(Rst("AccId"))
        If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo NextRecord
        'Set next row
        rowNo = rowNo + 1
        If .Rows < rowNo + 1 Then .Rows = rowNo + 1
        Count = 0
        AccID = FormatField(Rst("AccID"))
        .TextMatrix(rowNo, Count) = AccID: Count = Count + 1
        .TextMatrix(rowNo, Count) = FormatField(Rst("Name")): Count = Count + 1
        If chkAgent.Value = vbChecked Then
            .TextMatrix(rowNo, Count) = FormatField(Rst("AgentName")): Count = Count + 1
        End If
        .TextMatrix(rowNo, Count) = FormatField(Rst("CreateDate")): Count = Count + 1
        .TextMatrix(rowNo, Count) = FormatField(SecondRst("TotalAmount")): Count = Count + 1
        Total = Total + Val(FormatField(SecondRst("TotalAmount")))
    End With
NextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then Rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
Wend

'Set next row
With grd
    rowNo = rowNo + 2
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    .Row = rowNo
    
    .Col = 0: .Text = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 43)    '"Total Deposits "
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: grd.Text = FormatCurrency(Total): .CellAlignment = 4: .CellFontBold = True
End With

End Sub

Private Sub ShowDepositsClosed()

Dim SqlStr As String
Dim Rst As Recordset
Dim Total As Currency


RaiseEvent Processing("Reading & Verifying the records", 0)

SqlStr = "Select A.AccId,A.Amount as PrinAmount,B.Amount as IntAmount," & _
        " A.TransID,B.TransType " & _
        " From PDTrans A Left Join PDIntTrans B" & _
        " ON A.AccID = B.ACCID and A.TransID = B.TransID" & _
        " Where A.TransDate >= #" & m_FromDate & "#" & _
        " AND A.TransDate <= #" & m_ToDate & "#"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.CreateView("QryPDClose1")

SqlStr = "Select A.*,B.Amount as PayableAmount,  " & _
        " PrinAmount + IntAmount + Amount as MatAmount" & _
        " From qryPDClose1 A Left Join PDIntPayable B" & _
        " ON A.AccID = B.ACCID and A.TransID = B.TransID " & _
        " Where A.TransID = (Select Max(TransID) From PDTrans C " & _
            " Where C.AccId = A.AccID )"
        
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.CreateView("QryPDClose")

'Build the SQL
SqlStr = "Select AccNum,AgentID,MaturityDate, " & _
        " PigmyAmount,ClosedDate, c.*, B.Name " & _
        " From QryName B Inner join (PDMaster A " & _
        " Inner join qryPdClose C ON C.AccId = A.AccID) " & _
            " On B.CustomerId = A.CustomerID " & _
        " WHERE ClosedDate >= #" & m_FromDate & "#" & _
        " AND ClosedDate <= #" & m_ToDate & "#"

If Len(m_Place) Then SqlStr = SqlStr & " And B.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And B.Caste = " & AddQuotes(m_Caste)
If m_Gender Then SqlStr = SqlStr & " And B.Gender= " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroup
If m_AgentID Then SqlStr = SqlStr & " And A.AgentID = " & m_AgentID

If m_ReportOrder = wisByName Then
    SqlStr = SqlStr & " ORDER BY ClosedDate,A.AgentID,B.IsciName"
Else
    SqlStr = SqlStr & " ORDER BY ClosedDate,A.AgentID,val(A.AccNum)"
End If

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

'Initialize the Grid
Call InitGrid

Dim I As Integer
Dim SlNo As Integer

Dim AccID As Long
Dim AgentID As Integer
Dim AgentName As String
Dim TransID As Integer
Dim Amount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim rstTemp As Recordset
Dim rowNo As Long

I = IIf(chkAgent.Value = vbChecked, 1, 0)
rowNo = grd.Row
While Not Rst.EOF
    'Get Returned Amount
    If AgentID <> Rst("AgentID") Then AgentName = GetAgentName(Rst("AgentID"))
    AgentID = Rst("AgentID")
    
    SlNo = SlNo + 1
    gDbTrans.SQLStmt = "select * From qryPDClose " & _
                        " Where AccID = " & Rst("AccId")
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then GoTo NextRecord
    
    Amount = 0: IntAmount = 0
    PayableAmount = 0: TransID = 0
    
    Amount = FormatField(Rst("PrinAmount"))
    'Get the the INterest Paid Amount
    IntAmount = FormatField(Rst("IntAmount"))
    If Rst("TransType") = wWithdraw Or Rst("TransType") = wContraWithdraw _
                                Then IntAmount = IntAmount * -1
    PayableAmount = FormatField(Rst("PayableAmount"))
    If IntAmount < 0 Then IntAmount = IntAmount + PayableAmount: PayableAmount = 0
    
    'Checkthe condition of the minimum amount
    If (Amount + IntAmount + PayableAmount) < m_FromAmt Then GoTo NextRecord
    'Check the condition of the maximum amount is given
    If m_ToAmt > 0 And (Amount + IntAmount + PayableAmount) > m_ToAmt Then GoTo NextRecord
    
    With grd
        'Set next row
        SlNo = SlNo + 1
        If .Rows < rowNo + 2 Then .Rows = rowNo + 2
        rowNo = rowNo + 1:
        AccID = FormatField(Rst("AccId"))
        .TextMatrix(rowNo, 0) = SlNo
        .TextMatrix(rowNo, 1) = FormatField(Rst("AccNum"))
        .TextMatrix(rowNo, 2) = FormatField(Rst("Name"))
        If chkAgent.Value = vbChecked Then .TextMatrix(rowNo, 3) = AgentName
        
        .TextMatrix(rowNo, 3 + I) = FormatField(Rst("ClosedDate"))
        .TextMatrix(rowNo, 4 + I) = FormatCurrency(Amount + IntAmount + PayableAmount)
        Total = Total + Val(.TextMatrix(rowNo, 4 + I))
    End With
    
NextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then Rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
Wend

'Set next row
With grd
    rowNo = rowNo + 2
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    .Row = rowNo + 1
    .Col = 1: .Text = LoadResString(gLangOffSet + 52) & " " & _
                    LoadResString(gLangOffSet + 43) ' "Total Deposits "
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(Total)
    .CellAlignment = 7: .CellFontBold = True
End With

    lblReportTitle.Caption = LoadResString(gLangOffSet + 78) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        
End Sub

Private Sub ShowLiabilities()
Dim SQLStmt As String
Dim SecondRst As Recordset
Dim Rst As Recordset
Dim Count As Integer
'
RaiseEvent Processing("Reading & Verifying the records ", 0)

                    
SQLStmt = "Select A.AgentID,A.AccID,MaturityDate,CreateDate,RateOfInterest," & _
        " CLosedDate, Name " & _
        " From PDMaster A Inner join QryName B ON B.CustomerId = A.CustomerId" & _
        " Where A.AccId not In (Select AccId From PDMaster" & _
        " Where ClosedDate < #" & m_ToDate & "# )"

If chkAgent.Value Then
SQLStmt = "Select A.AgentID,A.AccID,MaturityDate,CreateDate,RateOfInterest," & _
        " ClosedDate, B.Name, D.Name as AgentName " & _
        " From QryName B Inner join (PDMaster A Inner join " & _
            " (UserTab C Inner join QryName D ON C.CustomerID = D.CustomerID)" & _
        " ON A.AgentID = C.UserID) ON B.CustomerId = A.CustomerId" & _
        " Where A.AccId not In (Select AccId From PDMaster" & _
            " Where ClosedDate < #" & m_ToDate & "#" & ")  "

End If

If m_FromAmt > 0 Then SQLStmt = " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then SQLStmt = SQLStmt & " And Balance < " & m_ToAmt
If Len(m_Place) Then SQLStmt = SQLStmt & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SQLStmt = SQLStmt & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then SQLStmt = SQLStmt & " And C.Gender= " & m_Gender
If m_AccGroup Then SQLStmt = SQLStmt & " And AccGroupID = " & m_AccGroup
If m_AgentID > 0 Then SQLStmt = SQLStmt & " And A.AgentID= " & m_AgentID


If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SQLStmt = SQLStmt & " Order by A.UserId,A.AccID"
Else
    gDbTrans.SQLStmt = SQLStmt & " Order by IsciName "
End If
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    Count = 100
    Count = 0
    Exit Sub
End If
    
'Init the grid
Call InitGrid

Dim Days As Integer
Dim Liability As Currency
Dim GrandTotal As Currency

Dim CustName As String
grd.Row = 0
Dim rowNo As Long

RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)
With grd
  While Not Rst.EOF
    gDbTrans.SQLStmt = "Select sum(Amount) as TotalAmount " & _
        " from PDTrans where TransDate <=" & _
        " #" & m_FromDate & "# and AccId = " & Rst("AccId")
    If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo NextRecord
    'Set next row
    rowNo = rowNo + 1
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    Count = 0
    .TextMatrix(rowNo, Count) = FormatField(Rst("AccID")): Count = Count + 1
    .TextMatrix(rowNo, Count) = FormatField(Rst("Name")): Count = Count + 1
    If chkAgent.Value = vbChecked Then
        .TextMatrix(rowNo, Count) = FormatField(Rst("AgentName")) ''GetAgentName(Val(Rst("UserId")))
        Count = Count + 1
    End If
    .TextMatrix(rowNo, Count) = FormatField(Rst("CreateDate")): Count = Count + 1
    .TextMatrix(rowNo, Count) = FormatField(Rst("MaturityDate")): Count = Count + 1
    .TextMatrix(rowNo, Count) = FormatField(Rst("RateOfInterest")): Count = Count + 1
    .TextMatrix(rowNo, Count) = FormatField(SecondRst("TotalAmount")): Count = Count + 1
    Liability = Val(FormatField(SecondRst("TotalAmount"))) + _
            ComputePDInterest(Val(FormatField(SecondRst("TotalAmount"))), Val(FormatField(Rst("RateOfInterest"))))
    .TextMatrix(rowNo, Count) = FormatCurrency(Liability): Count = Count + 1
    GrandTotal = GrandTotal + Liability
NextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then Rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
  Wend
'Fill In total Liability
    rowNo = rowNo + 2
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    
    .Row = rowNo
    .Col = 1: .Text = LoadResString(gLangOffSet + 52) + " " + LoadResString(gLangOffSet + 405) '"TOTAL LIABILITIES":
    .CellAlignment = 4: .CellFontBold = True
    
    .Col = IIf(chkAgent.Value = vbChecked, 7, 6)
    .Text = FormatCurrency(GrandTotal)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub


Private Sub ShowDayBook()

Dim SqlStr As String
Dim Rst As Recordset
Dim I As Integer
Dim TransDep As wisTransactionTypes
Dim TransType As wisTransactionTypes
Dim Count As Integer

'Report title
lblReportTitle.Caption = LoadResString(gLangOffSet + 71)

'To Get Deposits & Payments of of PD Account
TransDep = wDeposit
TransType = wWithdraw

Dim strClause As String
If m_FromAmt > 0 Then strClause = " And Amount > " & m_FromAmt
If m_ToAmt > 0 Then strClause = strClause & " And Amount < " & m_ToAmt
If Len(m_Place) Then strClause = strClause & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And C.Gender= " & m_Gender
If m_AccGroup Then strClause = strClause & " And AccGroupID = " & m_AccGroup
If m_AgentID > 0 Then strClause = strClause & " And A.AgentID = " & m_AgentID

RaiseEvent Processing("Reading & Verifyig the records ", 0)
On Error GoTo ErrLine

SqlStr = "Select C.Name, A.AccId,A.AgentId,A.AccNum, TransID,TransDate, " & _
    " Amount, TransType,IsciName from  QryName C Inner join " & _
    " (PDMaster A Inner Join PDTrans B On B.AccId = A.AccId) " & _
    " ON C.CustomerId = A.CustomerID " & _
    " where TransDate >= #" & m_FromDate & "# " & _
    " And TransDate <= #" & m_ToDate & "#" & _
 strClause

SqlStr = SqlStr & " UNION " & "Select  'INTEREST', A.AccId,A.AgentId,A.AccNum, " & _
    " TransID,TransDate, Amount, TransType,IsciName from QryName C" & _
    " Inner join (PDMaster A Inner join PDIntTrans B On B.AccId = A.AccId)" & _
    " ON C.CustomerId = A.CustomerID " & _
    " Where TransDate >= #" & m_FromDate & "# " & _
    " And TransDate <= #" & m_ToDate & "#"

SqlStr = SqlStr & strClause

If m_ReportOrder = wisByAccountNo Then
    SqlStr = SqlStr & " Order By TransDate,TransID,A.AgentId,A.AccNum"
Else
    SqlStr = SqlStr & " Order By TransDate,TransID,A.AgentId,C.IsciName" 'Isciname was not in the above query(Included)
End If

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
    
RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid
grd.Row = grd.FixedRows
Dim CustName As String
Dim TransDate As Date
Dim AgentName As String
Dim AgentDepositTotal As Currency
Dim AgentID As Long
Dim AccID As Long

Dim SubTotal() As Currency
Dim GrandTotal() As Currency

ReDim SubTotal(4 To grd.Cols - 1)
ReDim GrandTotal(4 To grd.Cols - 1)

AgentDepositTotal = 0
AgentID = Val(FormatField(Rst("AgentId")))

Dim PRINTTotal As Boolean
Dim SlNo As Integer
Dim blInt As Boolean

TransDate = Rst("TransDate")

'Put All The Agents name in a StringArray
grd.Row = grd.FixedRows
I = 0
If chkAgent.Value = vbChecked Then I = 1
AgentName = GetAgentName(AgentID)
Dim rowNo As Long, colNo As Byte

While Not Rst.EOF
    'See if you have to calculate sub totals
    With grd
        If AgentID <> Val(FormatField(Rst("AgentId"))) Then
            If .Rows <= rowNo + 2 Then .Rows = .Rows + 4
            rowNo = rowNo + 1
            .Row = rowNo
            .Col = 5: .Text = FormatCurrency(AgentDepositTotal)
            .CellAlignment = 7: .CellFontBold = True
            AgentID = Val(FormatField(Rst("AgentId")))
            AgentName = GetAgentName(AgentID)
            If .Rows <= rowNo + 2 Then .Rows = .Rows + 4
            rowNo = rowNo + 1
            AgentDepositTotal = 0
        End If
        If TransDate <> Rst("TransDate") Then
            PRINTTotal = True
            'Set next row
            AccID = 0: SlNo = 0
            If .Rows <= rowNo + 2 Then .Rows = .Rows + 2
            rowNo = rowNo + 1: Count = 0
            .Row = rowNo
            .Col = 3: .Text = LoadResString(gLangOffSet + 304) & _
                " " & GetIndianDate(TransDate)
            .CellAlignment = 4: .CellFontBold = True
            For Count = IIf(I, 5, 4) To .Cols - 1
                .Col = Count
                .Text = FormatCurrency(SubTotal(Count))
                .CellAlignment = 7: .CellFontBold = True
                GrandTotal(Count) = GrandTotal(Count) + SubTotal(Count)
                SubTotal(Count) = 0
            Next
            If .Rows <= rowNo + 2 Then .Rows = .Rows + 1
            rowNo = rowNo + 1: Count = 0
            .Row = rowNo
            TransDate = Rst("transDate")
        End If
        'Set next row
        If .Rows <= rowNo + 2 Then .Rows = rowNo + 2
        rowNo = rowNo + 1: Count = 0
        'if He has paid the interest amount then
        'need not to write into the grid 'so moveback one row
        If AccID = Rst("AccId") Then rowNo = rowNo - 1 Else blInt = False
        
        If chkAgent.Value = vbChecked Then .TextMatrix(rowNo, 4) = AgentName
        
        TransType = FormatField(Rst("TransType"))
        If FormatField(Rst(0)) = "INTEREST" Then
            blInt = True
            If TransType = wWithdraw Then colNo = 8 + I
            If TransType = wContraWithdraw Then colNo = 9 + I
        Else
            If AccID = Rst("AccId") And Not blInt Then rowNo = rowNo + 1
            SlNo = SlNo + 1
            .TextMatrix(rowNo, 0) = Format(SlNo, "00")
            .TextMatrix(rowNo, 3) = FormatField(Rst(0))
            If TransType = wDeposit Then colNo = 4 + I
            If TransType = wContraDeposit Then colNo = 5 + I
            If TransType = wWithdraw Then colNo = 6 + I
            If TransType = wContraWithdraw Then colNo = 7 + I
        End If
        
        .TextMatrix(rowNo, colNo) = FormatField(Rst("Amount"))
        SubTotal(.Col) = SubTotal(.Col) + Val(.Text)
        .TextMatrix(rowNo, 1) = GetIndianDate(TransDate)
        .TextMatrix(rowNo, 2) = FormatField(Rst("AccNum"))
        
    End With
        
        AccID = Rst("AccId")
        Rst.MoveNext
        DoEvents
        Me.Refresh
        RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
Wend

'Now Print the Subtotal Of the Last day
'Set next row
With grd
    AccID = 0
    If .Rows <= rowNo + 2 Then .Rows = .Rows + 2
    rowNo = rowNo + 1: Count = 0
    .Row = rowNo
    .Col = 2: .Text = LoadResString(gLangOffSet + 304) & _
        " " & GetIndianDate(TransDate): Count = Count + 1
    .CellAlignment = 4: .CellFontBold = True
    If chkAgent.Value = vbChecked Then Count = Count + 1
    For Count = IIf(chkAgent.Value = vbChecked, 5, 4) To .Cols - 1
        .Col = Count
        .Text = FormatCurrency(SubTotal(Count))
        .CellAlignment = 7: .CellFontBold = True
        GrandTotal(Count) = GrandTotal(Count) + SubTotal(Count)
    Next
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1: Count = 0
End With
      
'Now Print the grand total Of the Last day
If PRINTTotal = True Then
    With grd
        'Set next row
        If .Rows <= rowNo + 3 Then .Rows = rowNo + 3
        rowNo = rowNo + 2: Count = 0
        .Row = rowNo
        .Col = 2: .Text = LoadResString(gLangOffSet + 286): Count = Count + 1
        .CellAlignment = 4: .CellFontBold = True
        If chkAgent.Value = vbChecked Then Count = Count + 1
        For Count = IIf(chkAgent.Value = vbChecked, 5, 6) To .Cols - 1
            .Col = Count
            .Text = FormatCurrency(GrandTotal(Count))
            .CellAlignment = 7: .CellFontBold = True
        Next
    End With

End If
  
lblReportTitle.Caption = LoadResString(gLangOffSet + 390) & " " & _
        LoadResString(gLangOffSet + 63) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
  
Exit Sub
ErrLine:
'Resume
Exit Sub

End Sub

Private Sub ShowSubCashBook()

Dim SqlStr As String
Dim Rst As Recordset
Dim I As Integer
Dim TransDep As wisTransactionTypes
Dim TransType As wisTransactionTypes
Dim Count As Integer

'Report title
lblReportTitle.Caption = LoadResString(gLangOffSet + 71)

'To Get Deposits & Payments of of PD Account
TransDep = wDeposit
TransType = wWithdraw

Dim strClause As String
If m_FromAmt > 0 Then strClause = " And Amount > " & m_FromAmt
If m_ToAmt > 0 Then strClause = strClause & " And Amount < " & m_ToAmt
If Len(m_Place) Then strClause = strClause & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And C.Gender= " & m_Gender
If m_AccGroup Then strClause = strClause & " And AccGroupID = " & m_AccGroup

RaiseEvent Processing("Reading & Verifyig the records ", 0)
On Error GoTo ErrLine

SqlStr = "Select Name, A.AccId,A.AgentId,A.AccNum, TransID,TransDate, VoucherNo," & _
    " Amount, TransType,IsciName from QryName C" & _
    " Inner join (PDMaster A Inner join PDTrans B ON B.AccId = A.AccId)" & _
    " ON C.CustomerId = A.CustomerID  " & _
    " where TransDate >= #" & m_FromDate & "# " & _
    " And TransDate <= #" & m_ToDate & "#"
SqlStr = SqlStr & strClause

If m_ReportOrder = wisByAccountNo Then
    SqlStr = SqlStr & " Order By TransDate,A.AgentId,A.AccNum,TransID"
Else
    SqlStr = SqlStr & " Order By TransDate,A.AgentId,C.IsciName,TransID" 'Isciname was not in the above query(Included)
End If

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
SqlStr = ""

RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid
grd.Row = grd.FixedRows
Dim CustName As String
Dim TransDate As Date
Dim AgentName As String
Dim AgentDepositTotal As Currency
Dim AgentID As Long
Dim AccID As Long

Dim SubTotal() As Currency
Dim GrandTotal() As Currency

ReDim SubTotal(5 To grd.Cols - 1)
ReDim GrandTotal(5 To grd.Cols - 1)

AgentDepositTotal = 0

Dim PRINTTotal As Boolean
Dim SlNo As Integer
Dim blInt As Boolean
Dim rowNo As Long, colNo As Byte

TransDate = Rst("TransDate")

'Put All The Agents name in a StringArray
grd.Row = grd.FixedRows - 1
rowNo = grd.Row
I = 0
AgentName = GetAgentName(AgentID)

While Not Rst.EOF
    'See if you have to calculate sub totals
    With grd
        If AgentID <> Val(FormatField(Rst("AgentId"))) Then
            If .Rows <= rowNo + 2 Then .Rows = rowNo + 4
            rowNo = rowNo + 1
            .Row = rowNo
            .Col = 5: .Text = FormatCurrency(AgentDepositTotal)
            .CellAlignment = 7: .CellFontBold = True
            AgentDepositTotal = 0
        End If
        If TransDate <> Rst("TransDate") Then
            PRINTTotal = True: AgentID = 0
            'Set next row
            AccID = 0: SlNo = 0
            If .Rows < rowNo + 2 Then .Rows = rowNo + 2
            rowNo = rowNo + 1: Count = 0
            .Row = rowNo
            .Col = 3: .Text = LoadResString(gLangOffSet + 304) & _
                            " " & GetIndianDate(TransDate)
            .CellAlignment = 4: .CellFontBold = True
            For Count = IIf(I, 6, 5) To .Cols - 1
                .Col = Count
                .Text = FormatCurrency(SubTotal(Count))
                .CellAlignment = 7: .CellFontBold = True
                GrandTotal(Count) = GrandTotal(Count) + SubTotal(Count)
                SubTotal(Count) = 0
            Next
            If .Rows <= rowNo + 2 Then .Rows = rowNo + 2
            rowNo = rowNo + 1: Count = 0
            TransDate = Rst("transDate")
        End If
        If AgentID <> Val(FormatField(Rst("AgentId"))) Then
            AgentID = Val(FormatField(Rst("AgentId")))
            AgentName = GetAgentName(AgentID)
            If .Rows <= rowNo + 2 Then .Rows = rowNo + 4
            rowNo = rowNo + 1
            .Row = rowNo
            .Col = 3: .Text = AgentName: .CellFontBold = True
        End If
        
        'Set next row
        rowNo = rowNo + 1: Count = 0
        TransType = FormatField(Rst("TransType"))
        If AccID = Rst("AccId") And Not blInt Then rowNo = rowNo + 1
        If .Rows < rowNo + 1 Then .Rows = rowNo + 1
        
        SlNo = SlNo + 1
        .TextMatrix(rowNo, 0) = Format(SlNo, "00")
        .TextMatrix(rowNo, 1) = GetIndianDate(TransDate)
        .TextMatrix(rowNo, 2) = FormatField(Rst("AccNum"))
        .TextMatrix(rowNo, 3) = FormatField(Rst("Name"))
        .TextMatrix(rowNo, 4) = FormatField(Rst("VoucherNo"))
        
        colNo = 6
        If TransType = wDeposit Or TransType = wContraDeposit Then colNo = 5
        .TextMatrix(rowNo, colNo) = FormatField(Rst("Amount"))
        
        SubTotal(colNo) = SubTotal(colNo) + Val(.TextMatrix(rowNo, colNo))
        
    End With
        
    AccID = Rst("AccId")
    Rst.MoveNext
    DoEvents
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
Wend

'Now Print the Subtotal Of the Last day
'Set next row
With grd
    AccID = 0
    rowNo = rowNo + 1
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    .Row = rowNo: Count = 0
    .Col = 3: .Text = LoadResString(gLangOffSet + 304) & _
        " " & GetIndianDate(TransDate): Count = Count + 1
    .CellAlignment = 4: .CellFontBold = True
    If chkAgent.Value = vbChecked Then Count = Count + 1
    For Count = IIf(I, 6, 5) To .Cols - 1
        .Col = Count
        .Text = FormatCurrency(SubTotal(Count))
        .CellAlignment = 7: .CellFontBold = True
        GrandTotal(Count) = GrandTotal(Count) + SubTotal(Count)
    Next
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1: Count = 0
    rowNo = .Row
End With
      
'Now Print the grand total Of the Last day
If PRINTTotal = True Then
    With grd
        'Set next row
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1: Count = 0
        If .Rows = .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1: Count = 0
        .Col = 3: .Text = LoadResString(gLangOffSet + 286): Count = Count + 1
        .CellAlignment = 4: .CellFontBold = True
        If chkAgent.Value = vbChecked Then Count = Count + 1
        For Count = IIf(I, 6, 5) To .Cols - 1
            .Col = Count
            .Text = FormatCurrency(GrandTotal(Count))
            .CellAlignment = 7: .CellFontBold = True
        Next
    End With

End If
  
lblReportTitle.Caption = LoadResString(gLangOffSet + 390) & " " & _
        LoadResString(gLangOffSet + 85) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
  
Exit Sub
ErrLine:
'Resume
Exit Sub

End Sub

'
Private Sub ShowMonthlyTransaction(Optional Loan As Boolean)
Dim Rst As Recordset
Dim TransType As wisTransactionTypes
Dim Count As Integer

'To Get Deposits & Payments of of PD Account
TransType = wWithdraw

'Now Set the Date as on date
'
RaiseEvent Processing("Reading & Verifyig the records ", 0)
Err.Clear
grd.Clear
grd.Cols = 3
grd.Row = 0
grd.Col = 0: grd.Text = "Agent ID": grd.ColWidth(0) = 0.1
grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)
grd.Col = 2: grd.Text = LoadResString(gLangOffSet + 35)

'First Insert the AgentId, AccId  to the grid
gDbTrans.SQLStmt = "Select AgentID,AccID, AccNum, Name  as CustName " & _
    " From PDMaster A Inner join QryName B ON A.CustomerID = B.CustomerID " & _
    " Where (ClosedDate >= #" & m_ToDate & "# OR ClosedDate is NULL )" & _
    " Order By AgentId, val(AccNum)"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
grd.Row = 0

' Intially set the No col
Dim FirstDate As Date
Dim LastDate As Date
Dim colNo As Integer
Dim AgentID As Long
Dim AccNum As String

FirstDate = GetSysFirstDate(m_FromDate)
grd.Cols = grd.Cols + DateDiff("m", m_FromDate, m_ToDate) * 2

'Now onwardss alll dates are in "mm/dd/yyyyyy" format
LastDate = DateAdd("m", 1, FirstDate)
Dim rowNo As Long
rowNo = grd.Row
While Not Rst.EOF
    With grd
        rowNo = rowNo + 1
        If .Rows < rowNo + 1 Then .Rows = rowNo + 1
        .TextMatrix(rowNo, 0) = FormatField(Rst("AgentId"))
        .TextMatrix(rowNo, 1) = FormatField(Rst("AccNum"))
        .TextMatrix(rowNo, 2) = FormatField(Rst("CustName"))
    End With
    Rst.MoveNext
Wend


' Now start to fill the grid
colNo = 3
Do
        'Condition
        'Get One month transaction details of all accounts
        If DateDiff("M", FirstDate, m_ToDate) < 0 Then Exit Do
        Set Rst = Nothing
        gDbTrans.SQLStmt = "Select Sum(Amount) as TotalAmount,AgentId," & _
            " AccNum,TransType From PdTrans A Inner join PDMaster B ON " & _
            " A.AccID = B.accId WHERE Transdate >= #" & FirstDate & "# " & _
            " And TransDate < #" & LastDate & "# " & _
            " GROUP BY AgentId,AccNum,TransType" & _
            " ORDER BY AgentID,val(AccNum) "
    
        If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo NextMonth
        ' now Insert the Transaction
    With grd
        .Row = 0: rowNo = 0
        If .Cols <= colNo Then .Cols = colNo + 2
        .Col = colNo
        .Text = GetMonthString(Month(FirstDate)) & " " & LoadResString(gLangOffSet + 271)
        .Col = colNo + 1
        .Text = GetMonthString(Month(FirstDate)) & " " & LoadResString(gLangOffSet + 272)
    End With
    
    While Not Rst.EOF
        With grd
            TransType = FormatField(Rst("TransType"))
            'Now Get the Propre grid row to fit the Values
            Do
                AgentID = Val(.TextMatrix(rowNo, 0))
                AccNum = .TextMatrix(rowNo, 1)
                If AgentID = Rst("AgentId") And AccNum = Rst("AccNum") Then Exit Do
                If rowNo = .Rows - 1 Then
                    rowNo = 1
                    GoTo NextRecord
                End If
                If rowNo = .Rows - 1 Then Exit Do
                rowNo = rowNo + 1
            Loop
            If TransType = wDeposit Or TransType = wContraDeposit Then
                .TextMatrix(rowNo, colNo) = FormatField(Rst("TotalAmount"))
            Else 'If TransType = wWithDraw Then
                .TextMatrix(rowNo, colNo + 1) = FormatField(Rst("TotalAmount"))
            End If
        End With
NextRecord:
        Rst.MoveNext
    Wend
    
NextMonth:
    colNo = colNo + 2
    FirstDate = LastDate
    LastDate = DateAdd("m", 1, CDate(FirstDate))
Loop

End Sub

Private Sub ShowMonthlyTransactionNew(Optional Loan As Boolean)
Dim Rst As Recordset
Dim TransType As wisTransactionTypes

RaiseEvent Processing("Reading & Verifyig the records ", 0)
Err.Clear

Call InitGrid

'First Insert the AgentId, AccId  to the grid
gDbTrans.SQLStmt = "Select AgentID,AccID, AccNum, Name  as CustName " & _
    " From PDMaster A Inner join QryName B ON A.CustomerID = B.CustomerID " & _
    " WHERE (ClosedDate >= #" & m_ToDate & "# OR ClosedDate is NULL )" & _
    " Order By AgentId, val(AccNum)"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
grd.Row = 0

' Intially set the No col
Dim AccNum As String
Dim MonthNo As Integer
Dim SlNo As Integer
Dim rowNo As Long

grd.Row = grd.FixedRows - 1
rowNo = grd.Row
While Not Rst.EOF
    rowNo = rowNo + 1
    If grd.Rows < rowNo + 1 Then grd.Rows = rowNo + 1
    SlNo = SlNo + 1
    grd.TextMatrix(rowNo, 0) = Format(SlNo, "00")
    grd.TextMatrix(rowNo, 1) = FormatField(Rst("AgentId"))
    grd.TextMatrix(rowNo, 2) = FormatField(Rst("AccNum"))
    grd.TextMatrix(rowNo, 3) = FormatField(Rst("CustName"))
    Rst.MoveNext
Wend


' Now start to fill the grid
Dim DepAmount As Currency
Dim WithdrawAmount As Currency

rowNo = grd.FixedRows
While rowNo < grd.Rows
    With grd
    AccNum = grd.TextMatrix(rowNo, 2)
    
    gDbTrans.SQLStmt = "Select Sum(Amount) as TotalAmount,month(Transdate) as MonthNo," & _
            " TransType From PdTrans A Inner Join PDMaster B " & _
            " ON A.AccID = B.accId WHERE Transdate >= #" & m_FromDate & "# " & _
            " And TransDate < #" & m_ToDate & "# " & _
            " AND AccNum = " & AddQuotes(AccNum, True) & _
            " GROUP BY month(Transdate), TransType" & _
            " Order By month(Transdate)"
            
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) <= 0 Then GoTo NextAccount
    MonthNo = Rst("MonthNo")
    grd.Col = 4
    Do
        On Error Resume Next
        If MonthNo <> Rst("MonthNo") Or Rst.EOF Then
            Do
                .Row = 0
                If .Text = GetMonthString(MonthNo) Then
                    .Row = rowNo
                    .Text = FormatCurrency(DepAmount)
                    .Col = .Col + 1
                    .Text = FormatCurrency(WithdrawAmount)
                    Exit Do
                End If
                If .Col = .Cols - 1 Then Exit Do
                .Col = .Col + 1
            Loop
            DepAmount = 0: WithdrawAmount = 0
            MonthNo = Rst("MonthNo")
            If Rst.EOF Then Exit Do
        End If
        TransType = FormatField(Rst("TransType"))
        If TransType > 0 Then DepAmount = DepAmount + FormatField(Rst("totalAmount"))
        If TransType < 0 Then WithdrawAmount = WithdrawAmount + FormatField(Rst("totalAmount"))
        
        RaiseEvent Processing("Writing Data", rowNo / .Rows)
        Rst.MoveNext
        
    Loop

NextAccount:
    
    End With
    
    rowNo = rowNo + 1
Wend

End Sub

Private Sub cmdOk_Click()
chkAgent.Enabled = True

Unload Me
End Sub



Private Sub cmdPrint_Click()

Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
        
    Set m_frmCancel = New frmCancel
    Load m_frmCancel
    
    With m_frmCancel
       .Visible = True
       '.Show
    End With
    
    .CompanyName = gCompanyName
    .Font.Name = gFontName
    .ReportTitle = lblReportTitle
    .GridObject = grd
    .PrintGrid
Unload m_frmCancel
End With


End Sub

Private Sub cmdWeb_Click()
Dim clswebGrid As New clsgrdWeb
With clswebGrid
    Set .GridObject = grd
    .CompanyAddress = ""
    .CompanyName = gCompanyName
    .ReportTitle = lblReportTitle
    Call clswebGrid.ShowWebView '(grd)

End With

End Sub

Private Sub Form_Click()
Call grd_LostFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
gCancel = 0
'set kannada fonts
Call SetKannadaCaption
'Init the grid
With grd
    .Clear
    .Rows = 20
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    .Text = "No Records Available"
    .CellAlignment = 4: .CellFontBold = True
End With

'Show report
    chkAgent.Value = IIf(m_AgentNameShow, vbChecked, vbUnchecked)
    
    If m_ReportType = repPDBalance Then Call ShowDepositBalances
    If m_ReportType = repPDDayBook Then Call ShowDayBook
    If m_ReportType = repPDCashBook Then Call ShowSubCashBook
    If m_ReportType = repPDMat Then Call MaturedDeposits
    If m_ReportType = repPDAccClose Then Call ShowDepositsClosed
    If m_ReportType = repPDAccOpen Then
        lblReportTitle.Caption = LoadResString(gLangOffSet + 64) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDepositsOpened
    End If
    
    If m_ReportType = repPDLedger Then Call ShowDepositGeneralLedger
    
    If m_ReportType = repPDAgentTrans Then
        lblReportTitle.Caption = LoadResString(gLangOffSet + 330) & " " & _
                LoadResString(gLangOffSet + 38) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowAgentTransaction
    End If
    
    If m_ReportType = repPDMonTrans Then   'This Report to take individua recipts & Pay ments of account Holders
        lblReportTitle.Caption = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 38) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMonthlyTransactionNew
    End If
    If m_ReportType = repPDMonBal Then Call ShowMonthlyBalances
    
    lblReportTitle.FontSize = 14

'Set the Caption here
'Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 85)
End Sub

Private Sub ShowMonthlyBalances()
gCancel = 0
Dim Count As Long
Dim TotalCount As Long
Dim ProcCount As Long

Dim rstMain As Recordset
Dim SQLStmt As String

Dim FromDate As Date
Dim ToDate As Date

'Get the Last day of the given month
ToDate = GetSysLastDate(m_ToDate)
'Get the Last day of first month to get the balance of that month
FromDate = GetSysLastDate(m_FromDate)

'Set the Title for the Report.
lblReportTitle.Caption = LoadResString(gLangOffSet + 463) & " " & _
        LoadResString(gLangOffSet + 67) & " " & _
        LoadResString(gLangOffSet + 42) & " " & _
        GetFromDateString(GetMonthString(Month(FromDate)), GetMonthString(Month(ToDate)))

SQLStmt = "SELECT A.AccNum,A.AccID, A.CustomerID, Name as CustNAme " & _
        " From QryName B Inner join (PDMaster A inner join" & _
        " PDMaster C ON C.AccID = A.AccID) On B.CustomerID = A.CustomerID" & _
        " WHERE A.CreateDate <= #" & ToDate & "#" & _
        " AND (A.ClosedDate Is NULL OR A.ClosedDate >= #" & FromDate & "#)"

SQLStmt = SQLStmt & " Order By A.AgentID," & _
        IIf(m_ReportOrder = wisByAccountNo, "val(a.ACCNUM)", "IsciName")
        
gDbTrans.SQLStmt = SQLStmt
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub

Count = DateDiff("M", FromDate, ToDate) + 2
TotalCount = (Count + 1) * rstMain.RecordCount
RaiseEvent Initialise(0, TotalCount)

Dim prmAccId As Parameter
Dim prmdepositId As Parameter

With grd
    .Clear
    .Cols = 3
    .Rows = 5
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = LoadResString(gLangOffSet + 33) 'Sl No
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60) 'AccountNo
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = LoadResString(gLangOffSet + 35) 'Name
    .CellAlignment = 4: .CellFontBold = True
End With

grd.Row = 0: Count = 0
Dim rowNo As Long

While Not rstMain.EOF
    With grd
        rowNo = rowNo + 1
        If .Rows < rowNo + 1 Then .Rows = rowNo + 1
        Count = Count + 1
        .TextMatrix(rowNo, 0) = Count
        .TextMatrix(rowNo, 1) = FormatField(rstMain("AccNum"))
        .TextMatrix(rowNo, 2) = FormatField(rstMain("CustNAme"))
        .RowData(rowNo) = 0
    End With
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / TotalCount)
    rstMain.MoveNext
Wend
    
    With grd
        rowNo = rowNo + 2
        If .Rows < rowNo + 1 Then .Rows = rowNo + 1
        .Row = rowNo
        .Col = 2: .Text = LoadResString(gLangOffSet + 286) 'Grand Total
        .CellFontBold = True
    End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

Do
    If DateDiff("d", FromDate, ToDate) < 0 Then Exit Do
    
    SQLStmt = "SELECT AccId, Max(TransID) AS MaxTransID" & _
            " FROM PDTrans Where TransDate <= #" & FromDate & "# " & _
            " GROUP BY AccID"
    gDbTrans.SQLStmt = SQLStmt
    gDbTrans.CreateView ("PDMonBal")
    SQLStmt = "SELECT A.AccId,Balance From PDTrans A,PDMonBal B " & _
        " Where B.AccId = A.AccID ANd  TransID =MaxTransID"
    gDbTrans.SQLStmt = SQLStmt
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        .Row = 0: rowNo = 0
        .Col = .Cols - 1: .Text = GetMonthString(Month(FromDate)) & _
                " " & LoadResString(gLangOffSet + 42)
        .CellAlignment = 4: .CellFontBold = True
    End With
    
    rstMain.MoveFirst
    TotalBalance = 0
    
    While Not rstMain.EOF
        rowNo = rowNo + 1
        
        rstBalance.MoveFirst
        rstBalance.Find "ACCID = " & rstMain("AccID")
        If rstBalance.EOF Then GoTo NextAccount
        If rstBalance("Balance") = 0 Then GoTo NextAccount
        With grd
            .TextMatrix(rowNo, .Col) = FormatField(rstBalance("Balance"))
            .RowData(rowNo) = 1
        End With
        Balance = rstBalance("Balance")
        TotalBalance = TotalBalance + Balance
        
        DoEvents
        If gCancel Then rstMain.MoveLast
        RaiseEvent Processing("Calculating deposit balance", ProcCount / TotalCount)
                
NextAccount:
        rstMain.MoveNext
        ProcCount = ProcCount + 1
    Wend
    
    With grd
        .Row = .Rows - 1
        .Text = FormatCurrency(TotalBalance)
        .CellFontBold = True
        .RowData(rowNo) = 1
    End With
    
NextMonth:

    FromDate = DateAdd("D", 1, FromDate)
    FromDate = DateAdd("m", 1, FromDate)
    FromDate = DateAdd("D", -1, FromDate)
Loop

''Now Checkall the accounts
'Delete the account from grid which are not having any balance
With grd
    Count = 0
    Do
        Count = Count + 1
        If Count >= .Rows Then Exit Do
        If .RowData(Count) = 0 Then .RemoveItem (Count): Count = Count - 1
    Loop
    
End With

Exit Sub
ErrLine:
    MsgBox "Error MonBalance", vbExclamation, wis_MESSAGE_TITLE
    Err.Clear
End Sub


Private Sub Form_Resize()
Screen.MousePointer = vbDefault
On Error Resume Next
lblReportTitle.Top = 0
lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
grd.Left = 0
grd.Top = lblReportTitle.Top + lblReportTitle.Height
grd.Width = Me.Width - 150
fra.Top = Me.ScaleHeight - fra.Height
fra.Left = Me.Width - fra.Width
grd.Height = Me.ScaleHeight - fra.Height - lblReportTitle.Height
cmdOK.Left = fra.Width - cmdOK.Width - (cmdOK.Width / 4)
cmdPrint.Left = cmdOK.Left - cmdPrint.Width - (cmdPrint.Width / 4)
cmdWeb.Top = cmdPrint.Top
cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)

Dim Wid As Single
Dim I As Integer
Wid = (grd.Width - 185) / grd.Cols
    For I = 0 To grd.Cols - 1
        Wid = GetSetting(App.EXEName, "PDReport" & m_ReportType, "ColWidth" & I, 1 / grd.Cols) * grd.Width
        If Wid > grd.Width * 0.9 Then Wid = grd.Width / grd.Cols
        If Wid <= 15 Then Wid = 20
        grd.ColWidth(I) = Wid
    Next I

End Sub

Private Sub ShowDepositGeneralLedger()
Dim Count As Integer
Dim SqlStr As String
Dim Rst As Recordset
Dim TransDate As Date
Dim OpeningBalance As Currency
'
RaiseEvent Processing("Reading & Verifying the records", 0)

SqlStr = "Select 'PRINCIPAL',Sum(Amount) as TotalAmount,TransDate,TransType From PDTrans " & _
        " WHERE TransDate >= #" & GetSysFormatDate(m_FromIndianDate) & "# " & _
        " And TransDate <= #" & GetSysFormatDate(m_ToIndianDate) & "#" & _
        " Group By TransDate,TransType"

gDbTrans.SQLStmt = SqlStr & " ORDER BY TransDate"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

chkAgent.Enabled = False
RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid

With grd
    .Row = .FixedRows
    .Col = 0: .Text = LoadResString(gLangOffSet + 284) '"Opening Balnce"
    .CellAlignment = 4: .CellFontBold = True
    
    OpeningBalance = GetPDBalance(GetIndianDate(DateAdd("D", -1, m_FromDate)))
    .Col = 2: .Text = FormatCurrency(OpeningBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

Dim TransType As wisTransactionTypes
Dim DepositAmount As Currency
Dim WithdrawAmount As Currency
Dim TotalDepositAmount As Currency
Dim TotalWithdrawAmount As Currency
Dim PRINTTotal As Boolean
Dim SlNo As Integer
Dim rowNo As Long
TransDate = Rst("TransDate")
While Not Rst.EOF
    If TransDate <> Rst("TransDate") <> 0 Then
        With grd
            PRINTTotal = True
            rowNo = rowNo + 1
            If .Rows = rowNo + 2 Then .Rows = .Rows + 2
            .Row = rowNo
            SlNo = SlNo + 1
            .TextMatrix(rowNo, 0) = SlNo
            .TextMatrix(rowNo, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowNo, 2) = FormatCurrency(OpeningBalance)
            .TextMatrix(rowNo, 3) = FormatCurrency(DepositAmount)
            .TextMatrix(rowNo, 4) = FormatCurrency(WithdrawAmount)
            
            OpeningBalance = OpeningBalance + DepositAmount - WithdrawAmount
            .TextMatrix(rowNo, 5) = FormatCurrency(OpeningBalance)
            TotalDepositAmount = TotalDepositAmount + DepositAmount
            TotalWithdrawAmount = TotalWithdrawAmount + WithdrawAmount
            WithdrawAmount = 0: DepositAmount = 0
            TransDate = Rst("TransDate")
        End With
    End If
    
    TransType = FormatField(Rst("TransType"))
    If TransType = wDeposit Or TransType = wContraDeposit Then
        DepositAmount = DepositAmount + Rst("TotalAmount")
    Else
        WithdrawAmount = WithdrawAmount + Rst("TotalAmount")
    End If
    
    DoEvents
    Me.Refresh
    If gCancel Then Rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
Wend
    
With grd
    rowNo = rowNo + 1
    If .Rows < rowNo + 1 Then .Rows = rowNo + 1
    .Row = rowNo
    SlNo = SlNo + 1
    .TextMatrix(rowNo, 0) = SlNo
    .TextMatrix(rowNo, 1) = GetIndianDate(TransDate)
    .TextMatrix(rowNo, 2) = FormatCurrency(OpeningBalance)
    .TextMatrix(rowNo, 3) = FormatCurrency(DepositAmount)
    .TextMatrix(rowNo, 4) = FormatCurrency(WithdrawAmount)
    
    OpeningBalance = OpeningBalance + DepositAmount - WithdrawAmount
    .TextMatrix(rowNo, 5) = FormatCurrency(OpeningBalance): .CellAlignment = 7
    TotalDepositAmount = TotalDepositAmount + DepositAmount
    TotalWithdrawAmount = TotalWithdrawAmount + WithdrawAmount
    WithdrawAmount = 0: DepositAmount = 0

    If PRINTTotal Then
        rowNo = rowNo + 2
        If .Rows < rowNo + 1 Then .Rows = rowNo + 1
        .Row = rowNo
        .Col = 3: .Text = FormatCurrency(TotalDepositAmount)
        .CellAlignment = 4: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(TotalWithdrawAmount)
        .CellAlignment = 4: .CellFontBold = True
            
        If .Rows = .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 4
        .CellAlignment = 4: .CellFontBold = True
        .Text = LoadResString(gLangOffSet + 285)  '"Totals Amount"
        .Col = 5
        .CellAlignment = 7: .CellFontBold = True
        .Text = OpeningBalance
    Else
        .RemoveItem .FixedRows
    End If

End With

If DateDiff("D", m_FromDate, m_ToDate) = 0 Or Not PRINTTotal Then
    lblReportTitle.Caption = LoadResString(gLangOffSet + 425) & " " & _
        LoadResString(gLangOffSet + 93) '"Deposit GeneralLegder
Else
    Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 425) & " " & _
        LoadResString(gLangOffSet + 93) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
'Set mfrmPDReport = Nothing

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdPrint.Caption = LoadResString(gLangOffSet + 23)
Me.cmdOK.Caption = LoadResString(gLangOffSet + 11)
End Sub


Private Sub grd_LostFocus()
Dim ColCount As Integer
    
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "PDReport" & m_ReportType, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount

End Sub




Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
m_TotalCount = MaxCount
Set m_frmCancel = New frmCancel
m_frmCancel.PicStatus.Visible = True
m_frmCancel.PicStatus.ZOrder 0

End Sub

Private Sub m_grdPrint_Message(strMessage As String)
m_frmCancel.lblMessage = strMessage
End Sub


Private Sub m_grdPrint_ProcessCount(Count As Long)
On Error Resume Next

If (Count / m_TotalCount) > 0.95 Then
    Unload m_frmCancel
    Exit Sub
End If
UpdateStatus m_frmCancel.PicStatus, Count / m_TotalCount
Err.Clear


End Sub


