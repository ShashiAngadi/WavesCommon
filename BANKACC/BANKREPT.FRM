VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBankReport 
   Caption         =   "Bank Accounts Reports"
   ClientHeight    =   5790
   ClientLeft      =   1875
   ClientTop       =   2115
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   6495
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1170
      TabIndex        =   1
      Top             =   5070
      Width           =   5205
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         Height          =   345
         Left            =   1530
         TabIndex        =   4
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   345
         Left            =   3780
         TabIndex        =   3
         Top             =   210
         Width           =   1305
      End
      Begin VB.CheckBox chkDetails 
         Caption         =   "Show details"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   3285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4425
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   7805
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
      Left            =   2490
      TabIndex        =   5
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "frmBankReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_FromDate As String
Dim m_ToDate As String
Dim m_FromAmt As Currency
Dim m_ToAmt As Currency

Dim m_MinHeadID As Long
Dim m_MaxHeadID As Long
Dim m_HeadID

Dim m_BankTrans As wisBankHeads

Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)
Private Function BankBalance(OpeningDate As String) As Currency
'Declare the variables
Dim OpenBalance As Currency
Dim Rst As Recordset
Dim MinAccountHead As Integer
Dim MaxAccountHead As Integer
Dim TransID As Integer

MinAccountHead = frmBankAcc.cmbRepAccHead.ItemData(frmBankAcc.cmbRepAccHead.ListIndex)
MaxAccountHead = frmBankAcc.cmbRepAccHead.ItemData(frmBankAcc.cmbRepAccHead.ListIndex) + wis_BankHeadOffSet
RaiseEvent Processing("Reading & Verifying the records", 0)
gDBTrans.SQLStmt = "Select Max (transID) from Acctrans  where Transdate <= # " & FormatDate(OpeningDate) & _
                                "# and  accId >" & MinAccountHead & " And AccId < " & MaxAccountHead
If gDBTrans.SQLFetch < 1 Then
    Exit Function
Else
    Set Rst = gDBTrans.Rst.Clone
End If

On Error Resume Next
If Val(FormatField(Rst(0))) = 0 Then
    TransID = 0
Else
    TransID = FormatField(Rst(0))
End If
gDBTrans.SQLStmt = "Select  AccID,Balance,TransId from AccTrans A  where A.TransID  = " & TransID & _
                                 " and Transdate <= #" & FormatDate(OpeningDate) & "# and " & _
                                 " accId  >  " & MinAccountHead & " And AccId < " & MaxAccountHead
'gDBTrans.SQLStmt = "Select  AccID,Balance,TransId from AccTrans A  where A.TransID  =  " & _
                                 " (Select MAX  (TransID) from AccTrans B  where Transdate <= #" & FormatDate(OpeningDate) & "#) and " & _
                                 " accId  >  " & MinAccountHead & " And AccId < " & MaxAccountHead '& _
                                 " Order by  AccId ,Transdate desc "
If gDBTrans.SQLFetch < 1 Then
    Exit Function
Else
    Set Rst = gDBTrans.Rst.Clone
End If
RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Veifying the balance tobe calculated on   " & OpeningDate, 0)
' Show the balance
While Not Rst.EOF
    Dim AccId As Long
    If AccId = FormatField(Rst("AccID")) Then
        GoTo NextRecord
    End If
    AccId = FormatField(Rst("AccID"))
    OpenBalance = OpenBalance + FormatCurrency(Abs(FormatField(Rst("Balance"))))
NextRecord:
    DoEvents
    If gCancel Then Exit Function
    RaiseEvent Processing("Calculating the openning balance ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
    
Wend
BankBalance = OpenBalance
End Function

Private Sub InitGrid(Optional Resize As Boolean)

Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        grd.ColWidth(ColCount) = GetSetting(App.EXEName, lblReportTitle.Caption, "ColWidth" & ColCount, 1 / grd.Cols) * grd.Width
    Next ColCount
grd.ColWidth(0) = 25

If Resize Then Exit Sub

Dim ColWid As Single
If Not Resize Then
    grd.Clear
    grd.Rows = 30
End If
With frmBankAcc
        
        If .optReports(0).value Then
            If Not Resize Then
                grd.Cols = 3
            End If
            If grd.Cols = 1 Then
                grd.ColWidth(0) = Me.Width
                Exit Sub
            End If
            ColWid = (grd.Width - 100) / grd.Cols
            grd.FixedCols = 1
            grd.Row = 0
            grd.Col = 0: grd.ColWidth(0) = ColWid * 0.5: grd.Text = LoadResString(gLangOffSet + 33)  ' "Sl No"
            grd.Col = 1: grd.ColWidth(1) = ColWid * 1.4: grd.Text = LoadResString(gLangOffSet + 34)  '"Account Name"
            grd.Col = 2: grd.ColWidth(2) = ColWid * 1.1: grd.Text = LoadResString(gLangOffSet + 42)  '"Balance"
            grd.Row = 0
        ElseIf .optReports(1).value Then
            If Not Resize Then
                grd.Cols = 4
            End If
            If grd.Cols = 1 Then
                grd.ColWidth(0) = Me.Width
                Exit Sub
            End If
            ColWid = grd.Width / grd.Cols
            grd.FixedCols = 1
            grd.Row = 0
            grd.Col = 0: grd.ColWidth(0) = ColWid * 0.5: grd.Text = LoadResString(gLangOffSet + 33) '"Sl No"
            grd.Col = 1: grd.ColWidth(1) = ColWid * 1.7: grd.Text = LoadResString(gLangOffSet + 34) '"Account Name"    '
            grd.Col = 2: grd.ColWidth(2) = ColWid * 0.9: grd.Text = LoadResString(gLangOffSet + 37)
            grd.Col = 3: grd.ColWidth(3) = ColWid * 0.9: grd.Text = LoadResString(gLangOffSet + 42)
      ElseIf .optReports(2).value Then
            If Not Resize Then
                grd.Cols = 7
                grd.FixedCols = 1
            End If
            If grd.Cols = 1 Then
                grd.ColWidth(0) = Me.Width
                Exit Sub
            End If
            ColWid = grd.Width / grd.Cols
            grd.Row = 0
            grd.Col = 0: grd.ColWidth(0) = ColWid * 0.5: grd.Text = LoadResString(gLangOffSet + 33) '"Sl No"
            grd.Col = 1: grd.ColWidth(1) = ColWid * 1.3: grd.Text = LoadResString(gLangOffSet + 37)    '"dATE"
            grd.Col = 2: grd.ColWidth(2) = ColWid * 1.7: grd.Text = LoadResString(gLangOffSet + 34)      '" Account Name"
            grd.Col = 3: grd.ColWidth(2) = ColWid * 1.7: grd.Text = LoadResString(gLangOffSet + 39)      '" Account Name"
            grd.Col = 4: grd.ColWidth(3) = ColWid * 0.8: grd.Text = LoadResString(gLangOffSet + 279)   '"Withdrawn"
            grd.Col = 5: grd.ColWidth(4) = ColWid * 0.8: grd.Text = LoadResString(gLangOffSet + 271)   '"Deposited"
            grd.Col = 6: grd.ColWidth(1) = ColWid * 0.8: grd.Text = LoadResString(gLangOffSet + 274) 'interest
        ElseIf .optReports(3).value Then
            If Not Resize Then
                grd.Cols = 5
                grd.FixedCols = 1
            End If
            If grd.Cols = 1 Then
                grd.ColWidth(0) = Me.Width
                Exit Sub
            End If
            ColWid = grd.Width / grd.Cols
            grd.Row = 0
            grd.Col = 0: grd.ColWidth(0) = ColWid * 0.5: grd.Text = LoadResString(gLangOffSet + 33) '"Sl No"
            grd.Col = 1: grd.ColWidth(1) = ColWid * 1: grd.Text = LoadResString(gLangOffSet + 37)   '"Date"
            grd.Col = 2: grd.ColWidth(2) = ColWid * 1.4: grd.Text = LoadResString(gLangOffSet + 34) '"Account Name"
            grd.Col = 3: grd.ColWidth(3) = ColWid * 0.9: grd.Text = LoadResString(gLangOffSet + 403) '"Profit"
            grd.Col = 4: grd.ColWidth(4) = ColWid * 0.9: grd.Text = LoadResString(gLangOffSet + 404) '"Loss"
            
        ElseIf .optReports(4).value Then
            If Not Resize Then
                grd.Cols = 3
                grd.FixedCols = 1
            End If
            If grd.Cols = 1 Then
                grd.ColWidth(0) = Me.Width
                Exit Sub
            End If
            grd.Row = 0
            grd.Col = 0: grd.Text = LoadResString(gLangOffSet + 281)    '"Create Date"
            grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 36)   '"Acc No"
            grd.Col = 2: grd.Text = LoadResString(gLangOffSet + 35)   '"Name"
          
        ElseIf .optReports(5).value Then  ' General Leder
            If Not Resize Then
                grd.Cols = 4
                grd.FixedCols = 1
            End If
            If grd.Cols = 1 Then
                grd.ColWidth(0) = Me.Width
                Exit Sub
            End If
            ColWid = grd.Width / grd.Cols
            grd.Row = 0
            grd.Col = 0: grd.ColWidth(0) = ColWid * 0.5: grd.Text = LoadResString(gLangOffSet + 33)      '"Sl No"
            grd.Col = 1: grd.ColWidth(1) = ColWid * 1.1: grd.Text = LoadResString(gLangOffSet + 37)    '"Date"
            grd.Col = 2: grd.ColWidth(2) = ColWid * 1.1: grd.Text = LoadResString(gLangOffSet + 276)    ' "Debit "
            grd.Col = 3: grd.ColWidth(3) = ColWid * 1.1: grd.Text = LoadResString(gLangOffSet + 277)   ' "Credit"
        
        End If
    End With
    grd.Row = 0
    ColWid = 0
    Dim ColNo As Integer
    For ColNo = 0 To grd.Cols - 1
        grd.Col = ColNo
        grd.CellAlignment = 4
        grd.CellFontBold = True
        ColWid = ColWid + grd.ColWidth(ColNo)
    Next
      grd.ColWidth(grd.Cols - 1) = grd.ColWidth(grd.Cols - 1) + (grd.Width - ColWid) - (0.05 * Me.ScaleWidth)
End Sub

Private Sub SetKannadaCaption()
Dim Ctrl As Control
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next
chkDetails.Caption = LoadResString(gLangOffSet + 147)
cmdOK.Caption = LoadResString(gLangOffSet + 11)
cmdPrint.Caption = LoadResString(gLangOffSet + 23)
End Sub

Private Sub ShowAccountsClosed()

Dim SQLStmt As String
Dim Rst As Recordset
Dim NoAnd As Boolean   'Stupid variable

RaiseEvent Processing("Reading & Verifying the records", 0)
'Fire SQL
    SQLStmt = "Select * from CAmaster, NameTab "
    If m_FromDate <> "" And m_ToDate <> "" Then
        SQLStmt = SQLStmt & " where ClosedDate between #" & _
                    FormatDate(m_FromDate) & "# and #" & _
                    FormatDate(m_ToDate) & "#"
    ElseIf m_FromDate = "" And m_ToDate <> "" Then
        SQLStmt = SQLStmt & " where ClosedDate  <= #" & FormatDate(m_ToDate) & "#"
    ElseIf m_FromDate <> "" And m_ToDate = "" Then
        SQLStmt = SQLStmt & " where ClosedDate >= #" & FormatDate(m_FromDate) & "#"
    ElseIf m_FromDate = "" And m_ToDate = "" Then
        SQLStmt = SQLStmt & " where ClosedDate <> NULL "
    End If

    SQLStmt = SQLStmt & IIf(NoAnd = True, "", " and ") & " NameTab.CustomerID = CAmaster.CustomerID "
    
'    If frmBankAcc.optAccId Then
'        gDBTrans.SQLStmt = SQLStmt & " order by CreateDate, AccID"
'    Else
'        gDBTrans.SQLStmt = SQLStmt & "order by CreateDate, IsciName"
'    End If
'
    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    Else
        Set Rst = gDBTrans.Rst.Clone
    End If
    
'Initialize the grid
    Set Rst = gDBTrans.Rst.Clone
Call InitGrid
RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the records to write into the grid ", 0)
'Fill the grid
    While Not Rst.EOF
        'Set next row
        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
        
        grd.Col = 0: grd.Text = " " & FormatField(Rst("ClosedDate"))
        grd.Col = 1: grd.Text = " " & FormatField(Rst("AccID"))
        grd.Col = 2: grd.Text = " " & FormatField(Rst("Title")) & _
                                " " & FormatField(Rst("FirstName")) & _
                                " " & FormatField(Rst("MiddleName")) & _
                                " " & FormatField(Rst("LastName"))
NextRecord:
        DoEvents
        If gCancel Then Exit Sub
        RaiseEvent Processing("Writing the data to the grid ", Rst.AbsolutePosition / Rst.RecordCount)
        Rst.MoveNext
    
    Wend
End Sub

Private Sub ShowAccountsCreated()

Dim SQLStmt As String
Dim Rst As Recordset
Dim NoAnd As Boolean 'Stupid variable

RaiseEvent Processing("Reading & Verifying the records ", 0)
'Fire SQL
    SQLStmt = "Select * from CAmaster, NameTab "
    If m_FromDate <> "" And m_ToDate <> "" Then
        SQLStmt = SQLStmt & " where CreateDate between #" & _
                    FormatDate(m_FromDate) & "# and #" & _
                    FormatDate(m_ToDate) & "#"
    ElseIf m_FromDate = "" And m_ToDate <> "" Then
        SQLStmt = SQLStmt & " where CreateDate  <= #" & FormatDate(m_ToDate) & "#"
    ElseIf m_FromDate <> "" And m_ToDate = "" Then
        SQLStmt = SQLStmt & " where CreateDate >= #" & FormatDate(m_FromDate) & "#"
    ElseIf m_FromDate = "" And m_ToDate = "" Then
        SQLStmt = SQLStmt & " where "
        NoAnd = True
    End If

    SQLStmt = SQLStmt & IIf(NoAnd = True, "", " and ") & " NameTab.CustomerID = CAmaster.CustomerID "
    
'    If frmBankAcc.optAccId Then
'        gDBTrans.SQLStmt = SQLStmt & " order by CreateDate, AccID"
'    Else
'        gDBTrans.SQLStmt = SQLStmt & " order by CreateDate, IsciName"
'    End If

    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    Else
        Set Rst = gDBTrans.Rst.Clone
    End If
    
    Set Rst = gDBTrans.Rst.Clone
'Initialize the grid
    InitGrid
    
    RaiseEvent Initialise(0, Rst.RecordCount)
    RaiseEvent Processing("Aligning the data to write into the grid ", 0)
'Fill the grid
    While Not Rst.EOF
        'Set next row
        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
        
        grd.Col = 0: grd.Text = " " & FormatField(Rst("CreateDate"))
        grd.Col = 1: grd.Text = " " & FormatField(Rst("AccID"))
        grd.Col = 2: grd.Text = " " & FormatField(Rst("Title")) & _
                                " " & FormatField(Rst("FirstName")) & _
                                " " & FormatField(Rst("MiddleName")) & _
                                " " & FormatField(Rst("LastName"))
NextRecord:
        DoEvents
        If gCancel Then Exit Sub
        RaiseEvent Processing("Writing the data into the grid .", Rst.AbsolutePosition / Rst.RecordCount)
        Rst.MoveNext
    Wend
End Sub
Private Sub ShowBalances()
Dim Rst As Recordset
Dim i As Long
Dim Total As Currency
Dim SQL As String

RaiseEvent Processing("Reading & Verifying the records ", 0)
'Set the SQL Query
SQL = "Select A.AccId, Balance,AccName From AccTrans A, ACCMaster  where TransID = " & _
                        "(Select MAX(TransID) from AccTrans B " & _
                        " where A.AccID = B.AccID)  And A.AccId = AccMaster.AccId And A.AccId >= " & _
                        m_MinHeadID & " And A.AccId <= " & m_MaxHeadID
If m_FromAmt > 0 And m_ToAmt > 0 Then
    SQL = SQL & " And Balance between " & m_FromAmt & " and " & m_ToAmt

ElseIf m_FromAmt > 0 And m_ToAmt = 0 Then
    SQL = SQL & " And Balance >= " & m_FromAmt

ElseIf m_FromAmt = 0 And m_ToAmt > 0 Then
    SQL = SQL & " And Balance <=  " & m_ToAmt

ElseIf m_FromAmt = 0 And m_ToAmt = 0 Then
    SQL = SQL
End If
 gDBTrans.SQLStmt = SQL & " Order By A.AccId"

If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
End If

'Initialize the grid
Set Rst = gDBTrans.Rst.Clone
Call InitGrid

RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data to write into yhe grid. ", 0)

Dim SlNo As Integer
SlNo = 1
While Not Rst.EOF
    'See if you have to show this record
    If FormatField(Rst("Balance")) = 0 Then
        GoTo NextRecord
    End If
    
    'Set next row
    If grd.Rows = grd.Row + 2 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
  
    grd.Col = 0: grd.Text = " " & Format(SlNo, "00"): SlNo = SlNo + 1
    grd.Col = 1: grd.Text = FormatField(Rst("AccName"))
    If m_BankTrans = wis_AdvanceHead Then
        grd.Col = 2: grd.Text = FormatCurrency(Abs(FormatField(Rst("Balance"))))
    Else
        grd.Col = 2: grd.Text = FormatCurrency(Abs(FormatField(Rst("Balance"))))
    End If
    Total = Total + FormatField(Rst("Balance"))
NextRecord:
    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Writing the data into the grid .", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext

Wend

'Move next row
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 283) '"Total Balances"
    grd.CellAlignment = 4: grd.CellFontBold = True
    
    grd.Col = 2: grd.Text = FormatCurrency(Total): grd.CellFontBold = True

End Sub

Private Sub ShowGeneralLedger()
'Declare the variables
Dim SQLStmt As String
Dim OpeningDate As String
Dim OpeningBalance As Currency
Dim ClosingBalance As Currency
Dim Rst As Recordset
Dim Transdate As String

'Prelim check
If m_FromDate <> "" And m_ToDate = "" Then
    m_ToDate = m_FromDate
End If
RaiseEvent Processing("Reading & Verifying the records ", 0)

'Get liability on a day before m_fromdate
OpeningDate = FormatDate(DateAdd("d", -1, FormatDate(m_FromDate)))

'Build the SQL
SQLStmt = "SELECT * FROM AccTrans "
If m_FromDate <> "" And m_ToDate <> "" Then
    SQLStmt = SQLStmt & " WHERE TransDate BETWEEN #" & FormatDate(m_FromDate) & _
                         "# AND #" & FormatDate(m_ToDate) & "#"
ElseIf m_FromDate = "" And m_ToDate <> "" Then
    SQLStmt = SQLStmt & " WHERE TransDate <= #" & FormatDate(m_ToDate) & "#"
ElseIf m_FromDate <> "" And m_ToDate = "" Then
    SQLStmt = SQLStmt & " WHERE TransDate >= #" & FormatDate(m_FromDate) & "#"
End If

SQLStmt = SQLStmt & " AND AccId = " & m_MinHeadID '''& " And AccId <= " & m_MaxHeadID & " order by TransDate, AccID"
gDBTrans.SQLStmt = SQLStmt

'Fire the query
If gDBTrans.SQLFetch <= 0 Then
    Exit Sub
End If
Set Rst = gDBTrans.Rst.Clone
RaiseEvent Processing("Computing the openning balance", 0)

'COmpute liability (Opening Balance) as on this date
OpeningBalance = BankBalance(OpeningDate) 'ComputeTotalCALiability(OpeningDate)

'Initialize the grid
Dim SubTotal As Currency, GrandTotal As Currency
Dim WithDraw As Currency, Deposit As Currency, Interest As Currency, Charges As Currency
Dim TotalWithDraw As Currency, TotalDeposit As Currency, TotalInterest As Currency, TotalCharges As Currency

Call InitGrid
RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data to write into the grid. ", 0)
grd.Row = 0
SubTotal = 0: GrandTotal = 0
WithDraw = 0: Deposit = 0: Interest = 0: Charges = 0
Transdate = ""

If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
grd.Row = grd.Row + 1
grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 284): grd.CellFontBold = True
grd.Col = 2: grd.Text = FormatCurrency(OpeningBalance): grd.CellFontBold = True
'''grd.Row = grd.Row + 1
'Fill the grid
While Not Rst.EOF
   
    If Transdate <> "" And Transdate <> FormatField(Rst("TransDate")) Then
        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
        grd.Col = 1: grd.Text = Transdate
        grd.Col = 2: grd.Text = FormatCurrency(Interest)
        grd.Col = 3: grd.Text = FormatCurrency(Deposit)
        WithDraw = 0: Deposit = 0: Interest = 0: Charges = 0
    End If
    
    Dim TransType As wisTransactionTypes
    TransType = FormatField(Rst("TransType"))
    
    If TransType = wWithDraw Then
'''        grd.Col = 1
        WithDraw = WithDraw + FormatField(Rst("Amount"))
    ElseIf TransType = wDeposit Then
'''        grd.Col = 2:
        Deposit = Deposit + FormatField(Rst("Amount"))
    ElseIf TransType = wInterest Then
'''        grd.Col = 3:
        Interest = Interest + FormatField(Rst("Amount"))
    ElseIf TransType = wCharges Then
'''        grd.Col = 4:
        Charges = Charges + FormatField(Rst("Amount"))
    End If
'''    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
'''    grd.Row = grd.Row + 1
'''    grd.Text = FormatField(Rst("Amount"))
    Transdate = FormatField(Rst("TransDate"))
NextRecord:
    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Writing the data to the grid.", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext

Wend

If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
grd.Row = grd.Row + 1
'OpeningBalance = OpeningBalance + Deposit + Interest - WithDraw - Charges
'OpeningBalance = ComputeTotalCALiability(TransDate)
grd.Col = 1: grd.Text = Transdate
grd.Col = 2: grd.Text = FormatCurrency(Interest)
grd.Col = 3: grd.Text = FormatCurrency(Deposit)
If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
grd.Row = grd.Row + 1
grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 285): grd.CellFontBold = True  '"Closing Balance"

'compute closing balance
ClosingBalance = BankBalance(Transdate) 'ComputeTotalCALiability(OpeningDate)
grd.Col = 2: grd.Text = FormatCurrency(ClosingBalance): grd.CellFontBold = True
End Sub
Private Sub ShowLiability()
Dim Rst As Recordset
Dim i As Long
Dim Total As Currency

If m_FromDate = "" Then
    m_FromDate = gStrDate
End If

RaiseEvent Processing("Reading & Verifying the records ", 0)
'Fire SQL Query
gDBTrans.SQLStmt = "Select A.AccId, Balance,TransDate, AccName " & _
                        " From AccTrans A, ACCMaster where TransDate = " & _
                        "(Select MAX(TransDate) from AccTrans B " & _
                        " where A.AccID = B.AccID and TransDate <= # " & _
                        FormatDate(m_FromDate) & "#)  And A.AccId = AccMaster.AccId " & _
                        " And A.AccId >= " & m_MinHeadID & " And  A.AccId <= " & m_MaxHeadID & " Order By A.AccId"

    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    End If

'Initialize the Grid
    Set Rst = gDBTrans.Rst.Clone
    Call InitGrid
    grd.Row = 0

RaiseEvent Initialise(0, Rst.RecordCount)
RaiseEvent Processing("Aligning the data to write into the grid,", 0)
Dim AccId As Long
Dim SlNo As Long
SlNo = 1
While Not Rst.EOF
    'See if account is closed
    If FormatField(Rst("Balance")) = 0 Then  'Closed account
        AccId = FormatField(Rst("AccID"))
        GoTo NextRecord
    End If
    
    'See if you have to show this record
    If AccId = FormatField(Rst("AccID")) Then
        GoTo NextRecord
    End If
    'Set next row
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    
    grd.Col = 0: grd.Text = " " & Format(SlNo, "00"): SlNo = SlNo + 1
    grd.Col = 1: grd.Text = FormatField(Rst("AccName"))
    grd.Col = 2: grd.Text = FormatField(Rst("TransDate"))
    grd.Col = 3: grd.Text = FormatCurrency(Abs(FormatField(Rst("Balance"))))
    AccId = FormatField(Rst("AccID"))
    
    Total = Total + FormatCurrency(Abs(FormatField(Rst("Balance"))))

NextRecord:
    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Writing the data into the grid ", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
Wend

'Set next row and print grand total
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 286) ' "Grand Total"
    grd.CellFontBold = True: grd.CellAlignment = 4
    grd.Col = 3: grd.Text = FormatCurrency(Total)
    grd.CellFontBold = True:
End Sub
Private Sub ShowProductsAndInterests()
Dim AccIDs() As Long
Dim Products() As Currency
Dim Mn As Integer
Dim yr As Integer
Dim i As Integer
Dim This_date As String
Dim Rate As Double

'Prelim checks
    If Not DateValidate(m_FromDate, "/", True) Then
        Exit Sub
    End If
    If Not DateValidate(m_ToDate, "/", True) Then
        Exit Sub
    End If
    
'Initialize the grid
Call InitGrid
'Get the month start and number of months
    If WisDateDiff(m_FromDate, m_ToDate) < 0 Then
        Exit Sub
    End If
RaiseEvent Initialise(0, 12)
'Get the interest rate from setup
'    Dim l_Setup As New clsSetup
'    Rate = Val(l_Setup.ReadSetupValue("SBAcc", "RateOfInterest", 0))

'Loop through all the months till To_Date starting with From_Date
    This_date = m_FromDate
    While DateDiff("m", FormatDate(This_date), FormatDate(m_ToDate)) >= 0
        ReDim AccIDs(0)
        ReDim Products(0)
        Mn = Month(FormatDate(This_date))
        RaiseEvent Processing("Computing the interest for the month  " & GetMonthString(Mn), Mn / 12)
        yr = Year(FormatDate(This_date))
'        Call ComputeCAProducts(AccIDs(), Mn, Yr, Products())
        
        'NOTE:
        'I have deliberately put the foll 3 separate loops to minimize switching
        'between columns and to make code more readable. GIRISH
        
        'LOOP1:     Print the AccIDs the first time
        If m_FromDate = This_date Then
            'Reset the row
            With grd
                .Row = 0:
                .Col = 0: .Text = LoadResString(gLangOffSet + 36) ' "Account No."
                
                For i = 0 To UBound(AccIDs) - 1
                    If .Rows = .Row + 1 Then .Rows = .Rows + 1
                    .Row = .Row + 1
                    .Text = AccIDs(i)
                Next i
            End With
        End If
        
        'LOOP 2:    Print the products
        grd.Row = 0
        If grd.Cols = grd.Col + 1 Then grd.Cols = grd.Cols + 1
        grd.Col = grd.Col + 1
        grd.Text = GetMonthString(Mn) & " " & yr
        For i = 0 To UBound(AccIDs) - 1
            grd.Row = grd.Row + 1
            grd.Text = FormatCurrency(Products(i))
        Next i
        
        'LOOP 3:    Print the interest values
        grd.Row = 0
        If grd.Cols = grd.Col + 1 Then grd.Cols = grd.Cols + 1
        grd.Col = grd.Col + 1
        grd.Text = GetMonthString(Mn) & " " & yr
        For i = 0 To UBound(AccIDs) - 1
            grd.Row = grd.Row + 1
            'grd.Text = FormatCurrency(ComputeCAInterest(Products(I), Rate))
        Next i
        
        'Move to next month
        This_date = FormatDate(DateAdd("m", 1, FormatDate(This_date)))
    Wend
End Sub


Private Sub ShowTransactions()
'Declare the variables
Dim SQLStmt As String
Dim TmpStr As String
Dim Rst As Recordset
Dim Transdate As String

RaiseEvent Processing("Reading & Verifying the records ", 0)
'Build the Query
SQLStmt = "Select A.AccId,TransDate,TransType,Amount,Balance ,Particulars," & _
                " AccName From AccTrans A,AccMaster Where A.Accid = AccMaster.AcciD " & _
                " And A.AccId >= " & m_MinHeadID & " And A.AccId <= " & m_MaxHeadID
        
If m_FromDate <> "" And m_ToDate <> "" Then
    SQLStmt = SQLStmt & " And TransDate between #" & FormatDate(m_FromDate) & _
                         "# and #" & FormatDate(m_ToDate) & "#"
ElseIf m_FromDate = "" And m_ToDate <> "" Then
    SQLStmt = SQLStmt & " And TransDate <= #" & FormatDate(m_ToDate) & "#"
ElseIf m_FromDate <> "" And m_ToDate = "" Then
    SQLStmt = SQLStmt & " where TransDate >= #" & FormatDate(m_FromDate) & "#"
End If

If m_FromAmt <> 0 And m_ToAmt <> 0 Then
    TmpStr = "Amount >= " & m_FromAmt & " and Amount <= " & m_ToAmt
ElseIf m_FromAmt <> 0 And m_ToAmt = 0 Then
    TmpStr = "Amount >= " & m_FromAmt
ElseIf m_FromAmt = 0 And m_ToAmt <> 0 Then
    TmpStr = "Amount <= " & m_ToAmt
End If

If TmpStr <> "" Then
        SQLStmt = SQLStmt & " and " & TmpStr
End If

    SQLStmt = SQLStmt & " order by TransDate, A.AccID,Transid Asc"

'Assign The Querry to DBClass
gDBTrans.SQLStmt = SQLStmt

If gDBTrans.SQLFetch <= 0 Then
    Exit Sub
End If
Set Rst = gDBTrans.Rst.Clone
'Initialize the grid
    Dim SubTotal As Currency, GrandTotal As Currency
    Dim WithDraw As Currency, Deposit As Currency
    Dim TotalWithDraw As Currency, TotalDeposit As Currency
    Dim Interest As Currency, TotalInterest As Currency
    
    Call InitGrid
    grd.Row = 0
        
    RaiseEvent Initialise(0, Rst.RecordCount)
    RaiseEvent Processing("Aligning the data to write into the grid.", 0)
    
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0: Interest = 0
    TotalWithDraw = 0: TotalDeposit = 0: TotalInterest = 0
Transdate = ""
Dim SlNo As Long
SlNo = 1
While Not Rst.EOF
    'Set next row
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    
    If Transdate <> "" And Transdate <> FormatField(Rst("TransDate")) Then
        grd.Col = 2: grd.Text = "Sub Total ": grd.CellFontBold = True
        grd.Col = 4: grd.Text = FormatCurrency(WithDraw): grd.CellFontBold = True
        grd.Col = 5: grd.Text = FormatCurrency(Deposit): grd.CellFontBold = True
        grd.Col = 6: grd.Text = FormatCurrency(Interest): grd.CellFontBold = True
        
        TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
        TotalInterest = TotalInterest + Interest
        WithDraw = 0: Deposit = 0: Interest = 0
        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
'        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
'        grd.Row = grd.Row + 1

    End If
    grd.Col = 0: grd.Text = " " & Format(SlNo, "00"):   SlNo = SlNo + 1
    grd.Col = 1: grd.Text = FormatField(Rst("TransDate"))
    grd.Col = 2: grd.Text = FormatField(Rst("AccName"))
    grd.Col = 3: grd.Text = FormatField(Rst("Particulars")): grd.CellAlignment = 1
    Dim TransType As wisTransactionTypes
    TransType = FormatField(Rst("TransType"))
    
    If TransType = wContraWithdraw Or TransType = wWithDraw Or TransType = wContraInterest Then
        grd.Col = 4:         WithDraw = WithDraw + FormatField(Rst("Amount"))
       'TotalWithDraw = TotalWithDraw + WithDraw
    ElseIf TransType = wDeposit Or TransType = wContraDeposit Or TransType = wStock Then
        grd.Col = 5:         Deposit = Deposit + FormatField(Rst("Amount"))
        'TotalDeposit = TotalDeposit + Deposit
    ElseIf TransType = wInterest Or TransType = wContraCharges Then
        grd.Col = 6: Interest = Interest + FormatField(Rst("Amount"))
        'TotalInterest = TotalInterest + Interest
    End If
    grd.Text = FormatField(Rst("Amount"))
    Transdate = FormatField(Rst("TransDate"))
   
NextRecord:
    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Writing the data to the grid.", Rst.AbsolutePosition / Rst.RecordCount)
    
    Rst.MoveNext
Wend

'Show grand Total
    If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Rows + 2
    grd.Row = grd.Row + 1
    grd.Col = 2: grd.CellFontBold = True: grd.Text = LoadResString(gLangOffSet + 304) '"Sub Total "
    grd.Col = 4: grd.CellFontBold = True: grd.Text = FormatCurrency(WithDraw)
    grd.Col = 5: grd.CellFontBold = True: grd.Text = FormatCurrency(Deposit)
    grd.Col = 6: grd.CellFontBold = True: grd.Text = FormatCurrency(Interest)
    TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
    TotalInterest = TotalInterest + Interest
    WithDraw = 0: Deposit = 0: Interest = 0
    
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    grd.Col = 2: grd.CellFontBold = True: grd.Text = LoadResString(gLangOffSet + 286) '"Grand Total "
    grd.Col = 4: grd.CellFontBold = True: grd.Text = FormatCurrency(TotalWithDraw)
    grd.Col = 5: grd.CellFontBold = True: grd.Text = FormatCurrency(TotalDeposit)
    grd.Col = 6: grd.CellFontBold = True: grd.Text = FormatCurrency(TotalInterest)
    TotalWithDraw = 0: TotalDeposit = 0: TotalInterest = 0
End Sub

Private Sub ShowProfitLossTransaction()
'Declare the variables
Dim SQLStmt As String
Dim TmpStr As String
Dim Rst As Recordset
Dim Transdate As String

RaiseEvent Processing("Reading & Verifying the records ", 0)
'Build the Query
SQLStmt = "Select A.AccId,TransDate,TransType,Amount,Balance ," & _
                " AccName From AccTrans A,AccMaster Where A.Accid = AccMaster.AcciD " & _
                " And A.AccId >= " & m_MinHeadID & " And A.AccId <= " & m_MaxHeadID
        
If m_FromDate <> "" And m_ToDate <> "" Then
    SQLStmt = SQLStmt & " And TransDate between #" & FormatDate(m_FromDate) & _
                         "# and #" & FormatDate(m_ToDate) & "#"
ElseIf m_FromDate = "" And m_ToDate <> "" Then
    SQLStmt = SQLStmt & " And TransDate <= #" & FormatDate(m_ToDate) & "#"
ElseIf m_FromDate <> "" And m_ToDate = "" Then
    SQLStmt = SQLStmt & " where TransDate >= #" & FormatDate(m_FromDate) & "#"
End If

If m_FromAmt <> 0 And m_ToAmt <> 0 Then
    TmpStr = "Amount >= " & m_FromAmt & " and Amount <= " & m_ToAmt
ElseIf m_FromAmt <> 0 And m_ToAmt = 0 Then
    TmpStr = "Amount >= " & m_FromAmt
ElseIf m_FromAmt = 0 And m_ToAmt <> 0 Then
    TmpStr = "Amount <= " & m_ToAmt
End If

If TmpStr <> "" Then
        SQLStmt = SQLStmt & " and " & TmpStr
End If

    SQLStmt = SQLStmt & " order by TransDate, A.AccID"

'Assign The Querry to DBClass
gDBTrans.SQLStmt = SQLStmt

If gDBTrans.SQLFetch <= 0 Then
    Exit Sub
End If
Set Rst = gDBTrans.Rst.Clone

'Initialize the grid
    Dim SubTotal As Currency ', GrandTotal As Currency
    Dim Interest As Currency, Charges As Currency
    Dim TotalInterest As Currency, TotalCharges As Currency
    Dim SlNo As Long
    
    Call InitGrid
    grd.Row = 0
    RaiseEvent Initialise(0, Rst.RecordCount)
    RaiseEvent Processing("Aligning the data to write into the grid.", 0)
    
    Interest = 0: Charges = 0
Transdate = ""
SlNo = 1

While Not Rst.EOF
    
    'Set next row
    'If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    'grd.Row = grd.Row + 1
    
    If Transdate <> "" And Transdate <> FormatField(Rst("TransDate")) Then
        'Set next row
        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
        grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 304) '"Sub Total "
        grd.CellFontBold = True
        grd.Col = 3: grd.Text = FormatCurrency(TotalInterest): grd.CellFontBold = True
        grd.Col = 4: grd.Text = FormatCurrency(TotalCharges): grd.CellFontBold = True
       ' TotalInterest = TotalInterest + Interest: TotalCharges = TotalCharges + Charges
        'Interest = 0: Charges = 0
        
       ' If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        'grd.Row = grd.Row + 1
        'If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        'grd.Row = grd.Row + 1
    End If
    Dim TransType As wisTransactionTypes
    TransType = FormatField(Rst("TransType"))
    
    
    If TransType = wInterest Or TransType = wCharges Then
        'Set next row
        If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
        grd.Col = 0: grd.Text = " " & Format(SlNo, "00"):   SlNo = SlNo + 1
        grd.Col = 1: grd.Text = FormatField(Rst("TransDate"))
        grd.Col = 2: grd.Text = FormatField(Rst("AccName"))
    End If
    'Dim TransType As wisTransactionTypes
    'TransType = FormatField(Rst("TransType"))
    
    If TransType = wInterest Then
        grd.Col = 3:         Interest = FormatField(Rst("Amount"))
         grd.Text = FormatField(Rst("Amount"))
        Transdate = FormatField(Rst("TransDate"))
        TotalInterest = TotalInterest + Interest
    ElseIf TransType = wCharges Then
        grd.Col = 4:         Charges = FormatField(Rst("Amount"))
        grd.Text = FormatField(Rst("Amount"))
        Transdate = FormatField(Rst("TransDate"))
        TotalCharges = TotalCharges + Charges
    End If
    'TotalInterest = TotalInterest + Interest: TotalCharges = TotalCharges + Charges
    ' Interest = 0: Charges = 0
NextRecord:
    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Writing the data to the grid.", Rst.AbsolutePosition / Rst.RecordCount)
    Rst.MoveNext
Wend

'Show grand Total
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 304): grd.CellFontBold = True '"Sub Total "
    grd.Col = 3: grd.Text = FormatCurrency(Interest): grd.CellFontBold = True
    grd.Col = 4: grd.Text = FormatCurrency(Charges): grd.CellFontBold = True
    'grd.Col = 7: grd.Text = FormatCurrency(ComputeTotalCALiability(TransDate))
    Interest = 0: Charges = 0
    
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 286): grd.CellFontBold = True
    grd.Col = 3: grd.Text = FormatCurrency(TotalInterest): grd.CellFontBold = True
    grd.Col = 4: grd.Text = FormatCurrency(TotalCharges): grd.CellFontBold = True
    
End Sub


Private Sub cmdOK_Click()
Unload Me
End Sub



Private Sub Command1_Click()
Call Form_Load
End Sub

Private Sub cmdPrint_Click()
' Call the print class services...
Dim my_printClass As New clsPrint
Set my_printClass.DataSource = grd
my_printClass.ReportDestination = "PREVIEW"
my_printClass.ReportTitle = Me.lblReportTitle.Caption
my_printClass.HeaderRectangle = True
my_printClass.CompanyName = gCompanyName

'See the Code - Ravi
frmPrint.picPrint.Cls   'vinay

my_printClass.PrintReport

End Sub

Private Sub Form_Click()
    Call grd_LostFocus
End Sub

Private Sub Form_Load()
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption
'Initialize the grid

   grd.Clear
    grd.Rows = 50
    grd.Cols = 1
    grd.ColWidth(0) = grd.Width - 100
    grd.FixedCols = 0
    grd.Row = 1
    
    grd.Text = LoadResString(gLangOffSet + 278)   '"No Records Available"
    grd.CellFontBold = True: grd.CellAlignment = 4
    Screen.MousePointer = vbHourglass
    With frmBankAcc
        'Assign the Bank Head Const to varaible
        m_BankTrans = .cmbRepAccHead.ItemData(.cmbRepAccHead.ListIndex)
        'Assign the HeadId to Module level variables
        If .cmbRepAccHead.ListIndex >= 0 Then
            m_MinHeadID = Val(.cmbRepAccHead.ItemData(.cmbRepAccHead.ListIndex))
            m_MaxHeadID = m_MinHeadID + wis_BankHeadOffSet
            If .cmbRepAccName.ListIndex >= 0 Then
                m_HeadID = .cmbRepAccName.ItemData(.cmbRepAccName.ListIndex)
                m_MinHeadID = m_HeadID
                m_MaxHeadID = m_HeadID
            End If
        End If
        
        If .optReports(0).value Then
            If Not Trim(.txtAmt1.Text) = "" Then
               m_FromAmt = CCur((.txtAmt1.Text))
            End If
            If Not Trim(.txtAmt2.Text) = "" Then
               m_ToAmt = CCur((.txtAmt2.Text))
            End If
            If m_FromAmt = 0 And m_ToAmt = 0 Then
                Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 61)
            ElseIf m_FromAmt > 0 Then
                 Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 61) & " " & LoadResString(gLangOffSet + 42) & "  > or = " & m_FromAmt
            Else
                Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 61) & " " & m_FromAmt & " " + LoadResString(gLangOffSet + 277) & " " & m_ToAmt
            End If
            Call ShowBalances
        ElseIf .optReports(1).value Then
            m_FromDate = .txtDate1.Text
            Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 67) & " " & m_FromDate
            Call ShowLiability
        ElseIf .optReports(2).value Then
            m_FromDate = .txtDate1.Text
            m_ToDate = .txtDate2.Text
            Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 62) & " " & m_FromDate & " " + LoadResString(gLangOffSet + 277) & " " & m_ToDate
            Call ShowTransactions
        ElseIf .optReports(3).value Then
            m_FromDate = .txtDate1.Text
            m_ToDate = .txtDate2.Text
            Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 403) + " " + LoadResString(gLangOffSet + 277) + " " + LoadResString(gLangOffSet + 404) + " " + LoadResString(gLangOffSet + 28)
            Call ShowProfitLossTransaction
        ElseIf .optReports(4).value Then
            m_FromDate = .txtDate1.Text
            m_ToDate = .txtDate2.Text
            Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 65)
            Call ShowAccountsClosed
        ElseIf .optReports(5).value Then
            m_FromDate = .txtDate1.Text
            m_ToDate = .txtDate2.Text
            Me.lblReportTitle.Caption = LoadResString(gLangOffSet + 63)
            Call ShowGeneralLedger
        End If
        
    End With
Screen.MousePointer = vbNormal
Me.lblReportTitle.FontBold = True
Me.lblReportTitle.FontSize = 14

End Sub

Private Sub Form_Resize()
Screen.MousePointer = vbDefault
On Error Resume Next
lblReportTitle.Top = 0
lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
grd.Left = 0
grd.Top = lblReportTitle.Top + lblReportTitle.Height
grd.Width = Me.Width - 150
Fra.Top = Me.ScaleHeight - Fra.Height
Fra.Left = Me.Width - Fra.Width
grd.Height = Me.ScaleHeight - Fra.Height - lblReportTitle.Height

Dim wid As Single
Dim i As Integer
wid = (grd.Width - 185) / grd.Cols
For i = 0 To grd.Cols - 1
    grd.ColWidth(i) = wid - 30
Next i
InitGrid (True)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmBankReport = Nothing
   
End Sub




Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, lblReportTitle.Caption, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount

End Sub


