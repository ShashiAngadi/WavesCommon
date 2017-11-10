VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPig2PC 
   Caption         =   "Amount Collected"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8705
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPig2PC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim m_LastAccID As Long
Private m_PDHeadId As Long
Private m_TotalAmount As Currency

Dim m_RstPigmy As Recordset

Private Sub SavePrathinidhiTrans()

Dim strFileName As String
strFileName = App.Path & "\Pig_2_PC.Dat"

If Dir(strFileName) = "" Then
    MsgBox "Input does not exists", vbOKOnly, "Index 2000"
    Exit Sub
End If
    
'Now Read the file
Dim iFileNo As Integer
Dim strData As String
Dim lineCount As Integer
Dim strArr() As String
Dim lastAccId As Long
Dim NumOfRecord As Integer
Dim AmountCollected As Currency
Dim NumOfDays As Single
Dim AccountNotFound(0) As Integer

Dim AccID As Integer
Dim AccNum As Integer
Dim TransAmount() As Double
Dim TotalAmount() As Double
Dim AccountTotal As Double
Dim loopCount As Integer
Dim TransDate As Date
Dim LastTransDate As Date

Dim Balance As Currency
Dim TransID As Long

TransDate = GetSysFormatDate(gStrDate)
lineCount = 0
iFileNo = FreeFile
Open strFileName For Input As #iFileNo

gDbTrans.BeginTrans

Do While Not EOF(iFileNo)
  'Input #iFileNo, strData
  Line Input #iFileNo, strData
  If lineCount > 0 Then
    'Reading the Data
    strArr = Split(strData, ",")
    'Get the AccID
    'Search the details in AcciD
    AccID = GetAccountIDByAccountNumber(strArr(0))
    'AccID = CInt(strArr(0))
    'm_RstPigmy.MoveFirst
    'm_RstPigmy.Find "AccId = " & AccID
    'If account ID not found
    If m_RstPigmy.EOF Then GoTo NextRecord
    
    For loopCount = 1 To NumOfDays
        TransAmount(loopCount - 1) = CDbl(strArr(loopCount))
    Next loopCount
    
    Call InsertPigmyAmount(AccID, TransDate, TransAmount)
    'Update the Grid
    
  Else
    'reading the header
    strArr = Split(strData, ",")
    ''Get the AgenID as
    If gAgentID <> CInt(strArr(1)) Then
        MsgBox "The file does not belongs to this pigmy agent", vbOKOnly, "Index 2000"
        Exit Sub
    End If
    
    lastAccId = CLng(strArr(2))
    TransDate = GetSysFormatDate(strArr(3))
    AmountCollected = CCur(strArr(5))
    NumOfDays = CSng(strArr(8))
    NumOfRecord = CSng(strArr(7))
    Set m_RstPigmy = GetRecordSet(gAgentID)
    ReDim Preserve TransAmount(NumOfDays)
    ReDim Preserve TotalAmount(NumOfDays)
  End If
NextRecord:
  lineCount = lineCount + 1
Loop
Close #iFileNo

'Now Update the Pigmy Transaction for Agent
If Not AgentTransaction(GetSysFormatDate(gStrDate), AmountCollected) Then
    GoTo ErrLine:
End If

gDbTrans.CommitTrans

On Error Resume Next

'Move teh FIle to the Archive folder
''Check for the archive folder
If Not (GetAttr(App.Path & "\Archive") And vbDirectory) Then _
    MkDir (App.Path & "\Archive")
If Err.Number = 53 Then MkDir (App.Path & "\Archive")
'Now Move the file
MkDir (App.Path & "\Archive")
Dim target As String
target = App.Path & "\Archive\Pig_2_PC" & "_" & CStr(gAgentID) & "_" & Format(Now, "DD-MM-YYYY") & ".dat"
Name strFileName As target

Exit Sub

ErrLine:
    
     If gDbTrans.isInTransaction Then
        MsgBox "unable to transfer the pigmy details", vbOKOnly, wis_MESSAGE_TITLE
        gDbTrans.RollBack
     End If
End Sub

Private Sub SaveBalajiTrans()

Dim strFileName As String
Dim agentFileName As String
strFileName = App.Path & "\PCRX.TXT"

If gAgentID > 0 Then
    agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat.txt"
    If Dir(agentFileName) <> "" Then
        strFileName = agentFileName
    Else
        agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat"
        If Dir(agentFileName) <> "" Then strFileName = agentFileName
    End If
End If

If Dir(strFileName) = "" Then
    MsgBox "Input does not exists", vbOKOnly, "Index 2000"
    Exit Sub
End If
    
'Now Read the file
Dim iFileNo As Integer
Dim strData As String
Dim lineCount As Integer
Dim strArr() As String
Dim strArrWhole() As String
Dim lastAccId As Long
Dim NumOfRecord As Integer
Dim AmountCollected As Currency
Dim NumOfDays As Single
Dim AccountNotFound(0) As Integer

Dim AccID As Integer
Dim TransAmount() As Double
Dim TotalAmount() As Double
Dim AccountTotal As Double
Dim loopCount As Integer
Dim TransDate As Date
Dim LastTransDate As Date
Dim ArrayCount As Integer

Dim Balance As Currency
Dim TransID As Long

TransDate = GetSysFormatDate(gStrDate)
lineCount = 0
iFileNo = FreeFile
Open strFileName For Input As #iFileNo

ReDim Preserve TransAmount(0)

gDbTrans.BeginTrans

Do While Not EOF(iFileNo)
  'Input #iFileNo, strData
  Line Input #iFileNo, strData
NextLine:
  If lineCount > 0 Then
    'Reading the Data
    strArr = Split(strData, ",")
    
    'Get the AccID
    'Search the details in AcciD
    AccID = GetAccountIDByAccountNumber(strArr(0))
    
    'If account ID not found
    If m_RstPigmy.EOF Then GoTo NextRecord
    
    'TransAmount(0) = CDbl(strArr(5))  ' Accountwise Data
    TransAmount(0) = CDbl(strArr(1))  ' Reciept wise
    
    'For loopCount = 1 To NumOfDays
    '    TransAmount(loopCount - 1) = CDbl(strArr(4))
    'Next loopCount
    
    TransDate = GetSysFormatDate(Replace(strArr(4), ".", "/"))
    Call InsertPigmyAmount(AccID, TransDate, TransAmount)
    'Update the Grid
    
  Else 'READING THE HEADR
    
    strArr = Split(strData, ",")
    ''Get the AgenID as
    If gAgentID <> CInt(Right(strArr(3), 3)) Then
        MsgBox "The file does not belongs to this pigmy agent", vbOKOnly, "Index 2000"
        Exit Sub
    End If
    
    TransDate = GetSysFormatDate(Replace(strArr(4), ".", "/"))
    AmountCollected = CCur(strArr(2))
    
    NumOfDays = 1
    NumOfRecord = CSng(strArr(1))
    Set m_RstPigmy = GetRecordSet(gAgentID)
    ReDim Preserve TransAmount(NumOfDays)
    ReDim Preserve TotalAmount(NumOfDays)
    If (UBound(strArr)) > 6 Then
        strArrWhole = Split(strData, vbLf)
        ArrayCount = UBound(strArrWhole)
    End If

  End If
NextRecord:
  lineCount = lineCount + 1
  If ArrayCount > lineCount Then
    strData = strArrWhole(lineCount)
    GoTo NextLine
  End If
Loop
Close #iFileNo

'Now Update the Pigmy Transaction for Agent
If Not AgentTransaction(GetSysFormatDate(gStrDate), AmountCollected) Then
    GoTo ErrLine:
End If

gDbTrans.CommitTrans

On Error Resume Next

'Move teh FIle to the Archive folder
''Check for the archive folder
If Not (GetAttr(App.Path & "\Archive") And vbDirectory) Then _
    MkDir (App.Path & "\Archive")
If Err.Number = 53 Then MkDir (App.Path & "\Archive")
'Now Move the file
MkDir (App.Path & "\Archive")
Dim target As String
target = App.Path & "\Archive\Pig_2_PC" & "_" & CStr(gAgentID) & "_" & Format(Now, "DD-MM-YYYY") & ".dat"
Name strFileName As target

Exit Sub

ErrLine:
    
     If gDbTrans.isInTransaction Then
        MsgBox "unable to transfer the pigmy details", vbOKOnly, wis_MESSAGE_TITLE
        gDbTrans.RollBack
     End If
     
End Sub

Private Sub SaveBalajiTransOLD()

Dim strFileName As String
Dim agentFileName As String
strFileName = App.Path & "\PCRX.TXT"

If gAgentID > 0 Then
    agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat.txt"
    If Dir(agentFileName) <> "" Then
        strFileName = agentFileName
    Else
        agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat"
        If Dir(agentFileName) <> "" Then strFileName = agentFileName
    End If
End If

If Dir(strFileName) = "" Then
    MsgBox "Input does not exists", vbOKOnly, "Index 2000"
    Exit Sub
End If
    
'Now Read the file
Dim iFileNo As Integer
Dim strData As String
Dim lineCount As Integer
Dim strArr() As String
Dim strArrWhole() As String
Dim lastAccId As Long
Dim NumOfRecord As Integer
Dim AmountCollected As Currency
Dim NumOfDays As Single
Dim AccountNotFound(0) As Integer

Dim AccID As Integer
Dim TransAmount() As Double
Dim TotalAmount() As Double
Dim AccountTotal As Double
Dim loopCount As Integer
Dim TransDate As Date
Dim LastTransDate As Date
Dim ArrayCount As Integer

Dim Balance As Currency
Dim TransID As Long

TransDate = GetSysFormatDate(gStrDate)
lineCount = 0
iFileNo = FreeFile
Open strFileName For Input As #iFileNo

gDbTrans.BeginTrans

Do While Not EOF(iFileNo)
  'Input #iFileNo, strData
  Line Input #iFileNo, strData
NextLine:
  If lineCount > 0 Then
    'Reading the Data
    strArr = Split(strData, ",")
    
    'Get the AccID
    'Search the details in AcciD
    AccID = GetAccountIDByAccountNumber(strArr(2))
    'AccID = CInt(strArr(2))
    'm_RstPigmy.MoveFirst
    'm_RstPigmy.Find "AccId = " & AccID
    
    'If account ID not found
    If m_RstPigmy.EOF Then GoTo NextRecord
    
    For loopCount = 1 To NumOfDays
        TransAmount(loopCount - 1) = CDbl(strArr(4))
    Next loopCount
    
    TransDate = GetSysFormatDate(Replace(strArr(5), ".", "/"))
    Call InsertPigmyAmount(AccID, TransDate, TransAmount)
    'Update the Grid
    
  Else
    'reading the header
    strArr = Split(strData, ",")
    ''Get the AgenID as
    If gAgentID <> CInt(strArr(4)) Then
        MsgBox "The file does not belongs to this pigmy agent", vbOKOnly, "Index 2000"
        Exit Sub
    End If
    
'    lastAccId = CLng(strArr(2))
    TransDate = GetSysFormatDate(Replace(strArr(5), ".", "/"))
    AmountCollected = CCur(strArr(2))
    'NumOfDays = CSng(strArr(8))
    NumOfDays = 1
    NumOfRecord = CSng(strArr(3))
    Set m_RstPigmy = GetRecordSet(gAgentID)
    ReDim Preserve TransAmount(NumOfDays)
    ReDim Preserve TotalAmount(NumOfDays)
    If (UBound(strArr)) > 6 Then
        strArrWhole = Split(strData, vbLf)
        ArrayCount = UBound(strArrWhole)
    End If

  End If
NextRecord:
  lineCount = lineCount + 1
  If ArrayCount > lineCount Then
    strData = strArrWhole(lineCount)
    GoTo NextLine
  End If
Loop
Close #iFileNo

'Now Update the Pigmy Transaction for Agent
If Not AgentTransaction(GetSysFormatDate(gStrDate), AmountCollected) Then
    GoTo ErrLine:
End If

gDbTrans.CommitTrans

On Error Resume Next

'Move teh FIle to the Archive folder
''Check for the archive folder
If Not (GetAttr(App.Path & "\Archive") And vbDirectory) Then _
    MkDir (App.Path & "\Archive")
If Err.Number = 53 Then MkDir (App.Path & "\Archive")
'Now Move the file
MkDir (App.Path & "\Archive")
Dim target As String
target = App.Path & "\Archive\Pig_2_PC" & "_" & CStr(gAgentID) & "_" & Format(Now, "DD-MM-YYYY") & ".dat"
Name strFileName As target

Exit Sub

ErrLine:
    
     If gDbTrans.isInTransaction Then
        MsgBox "unable to transfer the pigmy details", vbOKOnly, wis_MESSAGE_TITLE
        gDbTrans.RollBack
     End If
     
End Sub


Private Sub cmdSave_Click()

Dim strPigmyType As String

If gDEVICE = "BALAJI_OLD" Then
    SaveBalajiTransOLD
ElseIf gDEVICE = "BALAJI" Then
    Call SaveBalajiTrans
Else
    Call SavePrathinidhiTrans
End If

Call gDbTrans.CloseDB

Unload Me

End Sub


Private Sub Form_Load()
Call gDbTrans.OpenDB(gDBFileName, constDBPWD)
    SetKannadaCaption
    'Get theData file
    If gDEVICE = "BALAJI_OLD" Then
        GetBalajiDataFromPigOLD
    ElseIf gDEVICE = "BALAJI" Then
        GetBalajiDataFromPig
    Else
        GetPrathinidhiData
    End If
    
End Sub
Private Sub GetPrathinidhiData()
On Error Resume Next

Dim strFileName As String
strFileName = ReadFromIniFile("Pigmy", "Machine", App.Path & "\" & constFINYEARFILE)

strFileName = App.Path & "\Pig_2_PC.Dat"


If Dir(strFileName) = "" Then
    Dim appid
    appid = Shell("Prati-Nidhi66.exe 2", vbNormalFocus)
    AppActivate appid, True
End If


If Dir(strFileName) = "" Then
    MsgBox "Input file does not exists", vbOKOnly, "Index 2000"
    'Unload Me
    Exit Sub
End If

'Now Read the file
Dim iFileNo As Integer
Dim strData As String
Dim lineCount As Integer
Dim strArr() As String
Dim lastAccId As Long
Dim NumOfRecord As Integer
Dim NumOfDays As Single
Dim AccountNotFound(0) As Integer
'Dim m_RstPigmy As Recordset

Dim AccID As Integer
Dim TransAmount() As Double
Dim TotalAmount() As Double
Dim AccountTotal As Double
Dim loopCount As Integer
Dim TransDate As Date
Dim LastTransDate As Date

lineCount = 0
iFileNo = FreeFile
Open strFileName For Input As #iFileNo
Do While Not EOF(iFileNo)
  'Input #iFileNo, strData
  Line Input #iFileNo, strData
  If lineCount > 0 Then
    'Reading the Data
    strArr = Split(strData, ",")
    'Get the AccID
    AccID = CInt(strArr(0))
    AccID = GetAccountIDByAccountNumber(strArr(0))
    'Search the details in AcciD
    'm_RstPigmy.MoveFirst
    'm_RstPigmy.Find "AccId = " & AccID

    'If account ID not found
    If m_RstPigmy.EOF Then
        'ReDim AccountNotFound(UBound(AccountNotFound) + 1)
        'AccountNotFound(UBound(AccountNotFound) - 1) = AccID
        GoTo NextRecord
    End If
    'Balance = FormatField(m_RstPigmy("Balance"))
    'TransId = FormatField(m_RstPigmy("TransID"))
    ''Now add this to Database
    
    'TransDate = GetSysFormatDate(strArr(10))
    For loopCount = 1 To NumOfDays
        TransAmount(loopCount - 1) = CDbl(strArr(loopCount))
    Next loopCount
    'Call InsertPigmyAmount(AccID, TransDate, TransAmount)
    'Update the Grid
    With grd
        If .rows = .Row + 1 Then .rows = .rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = CStr(.Row)
        .Col = 1: .Text = FormatField(m_RstPigmy("AccNum"))
        .Col = 2: .Text = FormatField(m_RstPigmy("Name"))
        .Col = 3: .Text = FormatField(m_RstPigmy("FullName"))
        AccountTotal = 0
        For loopCount = 4 To (3 + NumOfDays)
            .Col = loopCount: .Text = FormatCurrency(CCur(TransAmount(loopCount - 4)))
            AccountTotal = AccountTotal + CCur(TransAmount(loopCount - 4))
        Next
        .Col = loopCount: .Text = FormatCurrency(AccountTotal)
    End With
    
  Else
    'reading the header
    strArr = Split(strData, ",")
    ''Get the AgenID as
    cmdSave.Enabled = True
    If gAgentID <> CInt(strArr(1)) Then
        MsgBox "The file does not belongs to this pigmy agent", vbOKOnly, "Index 2000"
        cmdSave.Enabled = False
        'Exit Sub
    End If
    
    lastAccId = CLng(strArr(2))
    TransDate = GetSysFormatDate(strArr(3))
    NumOfDays = CSng(strArr(8))
    NumOfRecord = CSng(strArr(7))
    Set m_RstPigmy = GetRecordSet(CInt(strArr(1)))
    ReDim Preserve TransAmount(NumOfDays)
    ReDim Preserve TotalAmount(NumOfDays)
    Call InitGrid(TransDate, NumOfDays)
    grd.Row = 0
  End If
NextRecord:
  lineCount = lineCount + 1
Loop
Close #iFileNo
    
End Sub

Private Sub SetKannadaCaption()
    Call SetFontToControls(Me)
    
End Sub

Private Sub InitGrid(FirstDate As Date, NumOfDays As Single)
    Dim ColCount As Integer
    Dim LastDate As Date
    LastDate = DateAdd("d", NumOfDays - 1, FirstDate)
    With grd
        .rows = 2
        .cols = 6 + DateDiff("d", FirstDate, LastDate)
        'If chkAgent.Value = vbChecked Then .Cols = .Cols + 1
        .Row = 0
        .FixedRows = 1
        .FixedCols = 1
        .Col = 0: .Text = LoadResString(gLangOffSet + 33): .ColWidth(0) = 400 '"sL No"
        .Col = 1: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60): .ColWidth(1) = 800 '"Account No"
        .Col = 2: .Text = LoadResString(gLangOffSet + 35): .ColWidth(2) = 2400 '"Name"
        .Col = 3: .Text = "English Name": .ColWidth(3) = 2400
         ColCount = 4
         Do While DateDiff("d", LastDate, FirstDate) < 1
            .Col = ColCount: .Text = GetIndianDate(FirstDate): .ColWidth(ColCount) = 1100
            FirstDate = DateAdd("d", 1, FirstDate)
            ColCount = ColCount + 1
         Loop
        .Col = ColCount: .Text = LoadResString(gLangOffSet + 42): .ColWidth(ColCount) = 1800
    End With
End Sub

Private Function AgentTransaction(TransDate As Date, Amount As Currency) As Boolean

Dim PrevAmount As Currency
Dim LastDate As Date
Dim Trans As wisTransactionTypes

Dim InTrans As Boolean

InTrans = gDbTrans.isInTransaction

'TransDate = GetSysFormatDate(txtAgentDate)
'Amount = txtAgentAmount

'Get Th LAst TransCtion Date
Dim Balance As Currency
Dim IntBalance As Currency
Dim TransID As Long
Dim Rst As Recordset

'Get the Balance and new transid
LastDate = "1/1/100"

gDbTrans.SQLStmt = "Select TOP 1 * from AgentTrans " & _
            " Where AgentID = " & gAgentID & " order by TransID desc"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    'Check The Transaction date w.r.t to Account CreateDate
    TransID = 100
    LastDate = "1/1/100"
    Balance = Val(InputBox("Please enter a balance to start with as this account has not transaction performed", "Initial Balance", "0.00"))
    If Balance < 0 Then
        'MsgBox "Invalid initial balance specified !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 517), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
Else
    Balance = FormatField(Rst("Balance"))
    TransID = FormatField(Rst("TransID")) + 1
    LastDate = Rst.Fields("TransDate")
    
    'See if the date is earlier than last date of transaction
    If DateDiff("D", TransDate, LastDate) > 0 Then
        'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 568), vbExclamation, gAppName & " - Error"
        Exit Function
    ElseIf DateDiff("D", TransDate, LastDate) = 0 Then
        PrevAmount = FormatField(Rst.Fields("Amount"))
        'If MsgBox("Transaction of " & Me.cmbAgentList.Text & _
                " On " & txtAgentDate & " already made " & vbCrLf & _
                " Do you want to update this transaction", vbQuestion + vbYesNo, _
                wis_MESSAGE_TITLE) = vbNo Then Exit Function
    End If
End If

Trans = wDeposit
'Get the Particulars
    Balance = Balance + Amount
    If Not InTrans Then gDbTrans.BeginTrans
    
    gDbTrans.SQLStmt = "Insert into AgentTrans (AgentID, TransID, " & _
            " TransDate, Amount,Balance, Particulars," & _
            " TransType, VoucherNo) values ( " & _
            gAgentID & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            Amount & "," & _
            Balance & "," & _
            "'From Pigmy colection'," & _
            Trans & ",'Pigmy')"
    
    If DateDiff("D", TransDate, LastDate) = 0 Then
        Balance = Rst("Balance") - PrevAmount + Amount
        TransID = Rst("TransID")
        gDbTrans.SQLStmt = "UPDATE AgentTrans " & _
            " SET Amount = " & Amount & "" & _
            " WHERE AgentID = " & gAgentID & _
            " AND TransID = " & TransID & _
            " AND #" & TransDate & "#"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
        
        gDbTrans.SQLStmt = "UPDATE AgentTrans " & _
            " SET Balance = Balance + " & "(" & Amount - PrevAmount & ")" & _
            " WHERE AgentID = " & gAgentID & "" & _
            " AND TransID >= " & TransID & _
            " AND TransDate >= #" & TransDate & "#"
    End If
    
    If Not gDbTrans.SQLExecute Then
        If Not InTrans Then gDbTrans.RollBack
        Exit Function
    End If

Dim BankClass As clsBankAcc
Set BankClass = New clsBankAcc

'Get the Pigmy headId
'get HeadID in the HeadsAccTrans Table(PigmyHeadID)

'Get the Pigmy HeadID
m_PDHeadId = BankClass.GetHeadIDCreated(LoadResString(gLangOffSet + 425), _
        parMemberDeposit, 0, wis_PDAcc)

'Perform the tranaction in the Bank Head
If Not BankClass.UpdateCashDeposits(m_PDHeadId, Amount - PrevAmount, TransDate) Then
    If Not InTrans Then gDbTrans.RollBack
    Set BankClass = Nothing
    Exit Function
End If

Set BankClass = Nothing
MsgBox "The pigmy transfer has done", vbOKOnly, "Inex 2000"
AgentTransaction = True
If Not InTrans Then gDbTrans.CommitTrans

End Function

Private Sub Form_Terminate()
If Not gDbTrans Is Nothing Then gDbTrans.CloseDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not gDbTrans Is Nothing Then gDbTrans.CloseDB
End Sub

Private Function GetAccountIDByAccountNumber(AccNum As String) As Integer
GetAccountIDByAccountNumber = 0
    Dim newAccNum As String
    newAccNum = Trim$(AccNum)
    If gDEVICE <> "BALAJI" And gDEVICE <> "BALAJI_OLD" Then
        Do While Mid$(newAccNum, 1, 1) = "0"
            ''
            m_RstPigmy.MoveFirst
            m_RstPigmy.Find "AccNum = '" & newAccNum & "'"
            If Not m_RstPigmy.EOF Then GetAccountIDByAccountNumber = FormatField(m_RstPigmy("AccID")): Exit Function
            ''
            newAccNum = Mid$(newAccNum, 2)
        Loop
    End If
    'Search the details in AcciD
    m_RstPigmy.MoveFirst
    m_RstPigmy.Find "AccNum = '" & newAccNum & "'"
    If Not m_RstPigmy.EOF Then GetAccountIDByAccountNumber = FormatField(m_RstPigmy("AccID"))
End Function


Private Sub GetBalajiDataFromPig()
On Error Resume Next

Dim strFileName As String
Dim agentFileName As String
strFileName = App.Path & "\PCRX.TXT"

If gAgentID > 0 Then
    agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat.txt"
    
    If Dir(agentFileName) <> "" Then
        strFileName = agentFileName
    Else
        agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat"
        If Dir(agentFileName) <> "" Then strFileName = agentFileName
    End If

End If

If Dir(strFileName) = "" Then
    MsgBox "Input file does not exists", vbOKOnly, "Index 2000"
    'Unload Me
    Exit Sub
End If

'Now Read the file
Dim iFileNo As Integer
Dim strData As String
Dim lineCount As Integer
Dim strArr() As String
Dim strArrWhole() As String
Dim lastAccId As Long
Dim NumOfRecord As Integer
Dim NumOfDays As Single
Dim AccountNotFound(0) As Integer
'Dim m_RstPigmy As Recordset

Dim AccID As Integer
Dim TransAmount As Double
Dim TotalAmount() As Double
Dim AccountTotal As Double
Dim loopCount As Integer
Dim TransDate As Date
Dim LastTransDate As Date
Dim ArrayCount As Integer

lineCount = 0: ArrayCount = 0

iFileNo = FreeFile
Open strFileName For Input As #iFileNo
Do While Not EOF(iFileNo)
  'Input #iFileNo, strData
  Line Input #iFileNo, strData
NextLine:
  If lineCount > 0 Then
    'Reading the Data
    strArr = Split(strData, ",")
    'Get the AccID
    AccID = CInt(strArr(0))
    
    'Search the details in AcciD
    'm_RstPigmy.MoveFirst
    'm_RstPigmy.Find "AccId = " & AccID
    AccID = GetAccountIDByAccountNumber(Trim$(strArr(0)))
    'If account ID not found
    If m_RstPigmy.EOF Then
        'ReDim AccountNotFound(UBound(AccountNotFound) + 1)
        'AccountNotFound(UBound(AccountNotFound) - 1) = AccID
        GoTo NextRecord
    End If
    'Balance = FormatField(m_RstPigmy("Balance"))
    'TransId = FormatField(m_RstPigmy("TransID"))
    ''Now add this to Database
    
    TransDate = GetSysFormatDate(Replace(strArr(4), ".", "/"))
    'TransAmount = CDbl(strArr(5))  ' Accountwise Data
    TransAmount = CDbl(strArr(1))  ' Reciept wise
    
    'TransAmount = CDbl(strArr(4))
    'Call InsertPigmyAmount(AccID, TransDate, TransAmount)
    'Update the Grid
    With grd
        If .rows = .Row + 1 Then .rows = .rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = CStr(.Row)
        .Col = 1: .Text = FormatField(m_RstPigmy("AccNum"))
        .Col = 2: .Text = FormatField(m_RstPigmy("Name"))
        .Col = 3: .Text = FormatField(m_RstPigmy("FullName"))
        AccountTotal = 0
        For loopCount = 4 To (3 + NumOfDays)
            '.Col = loopCount: .Text = FormatCurrency(CCur(TransAmount(loopCount - 4)))
            'AccountTotal = AccountTotal + CCur(TransAmount(loopCount - 4))
            .Col = loopCount: .Text = FormatCurrency(CCur(TransAmount))
            AccountTotal = AccountTotal + CCur(TransAmount)
        Next
        .Col = loopCount: .Text = FormatCurrency(AccountTotal)
    End With
    
  Else
    'reading the header
    strArr = Split(strData, ",")
    ''Get the AgenID as
    cmdSave.Enabled = True
    If gAgentID <> CInt(strArr(3)) Then
        MsgBox "The file does not belongs to this pigmy agent", vbOKOnly, "Index 2000"
        cmdSave.Enabled = False
        'Exit Sub
    End If
    
    'lastAccId = CLng(strArr(2))
    If InStr(1, strArr(4), ".") Then
        strArr(4) = Replace(strArr(4), ".", "/")
    End If
    TransDate = GetSysFormatDate(strArr(4))
    
    'NumOfDays = CSng(strArr(8))
    NumOfDays = 1
    NumOfRecord = CSng(strArr(1))
    Set m_RstPigmy = GetRecordSet(CInt(Right(strArr(3), 3)))
    'ReDim Preserve TransAmount(NumOfDays)
    ReDim Preserve TotalAmount(NumOfDays)
    Call InitGrid(TransDate, NumOfDays)
    grd.Row = 0
    If (UBound(strArr)) > 5 Then
        strArrWhole = Split(strData, vbLf)
        ArrayCount = UBound(strArrWhole)
    End If
  End If
NextRecord:
  lineCount = lineCount + 1
  If ArrayCount > lineCount Then
    strData = strArrWhole(lineCount)
    GoTo NextLine
  End If
Loop
Close #iFileNo
    
End Sub


Private Sub GetBalajiDataFromPigOLD()
On Error Resume Next

Dim strFileName As String
Dim agentFileName As String
strFileName = App.Path & "\PCRX.TXT"

If gAgentID > 0 Then
    agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat.txt"
    
    If Dir(agentFileName) <> "" Then
        strFileName = agentFileName
    Else
        agentFileName = App.Path + "\" + Format(gAgentID, "0000") + "-pcrx.dat"
        If Dir(agentFileName) <> "" Then strFileName = agentFileName
    End If

End If

If Dir(strFileName) = "" Then
    MsgBox "Input file does not exists", vbOKOnly, "Index 2000"
    'Unload Me
    Exit Sub
End If

'Now Read the file
Dim iFileNo As Integer
Dim strData As String
Dim lineCount As Integer
Dim strArr() As String
Dim strArrWhole() As String
Dim lastAccId As Long
Dim NumOfRecord As Integer
Dim NumOfDays As Single
Dim AccountNotFound(0) As Integer
'Dim m_RstPigmy As Recordset

Dim AccID As Integer
Dim TransAmount() As Double
Dim TotalAmount() As Double
Dim AccountTotal As Double
Dim loopCount As Integer
Dim TransDate As Date
Dim LastTransDate As Date
Dim ArrayCount As Integer

lineCount = 0: ArrayCount = 0

iFileNo = FreeFile
Open strFileName For Input As #iFileNo
Do While Not EOF(iFileNo)
  'Input #iFileNo, strData
  Line Input #iFileNo, strData
NextLine:
  If lineCount > 0 Then
    'Reading the Data
    strArr = Split(strData, ",")
    'Get the AccID
    AccID = CInt(strArr(2))
    
    'Search the details in AcciD
    'm_RstPigmy.MoveFirst
    'm_RstPigmy.Find "AccId = " & AccID
    AccID = GetAccountIDByAccountNumber(strArr(2))
    'If account ID not found
    If m_RstPigmy.EOF Then
        'ReDim AccountNotFound(UBound(AccountNotFound) + 1)
        'AccountNotFound(UBound(AccountNotFound) - 1) = AccID
        GoTo NextRecord
    End If
    'Balance = FormatField(m_RstPigmy("Balance"))
    'TransId = FormatField(m_RstPigmy("TransID"))
    ''Now add this to Database
    
    'TransDate = GetSysFormatDate(strArr(10))
    For loopCount = 1 To NumOfDays
        TransAmount(loopCount - 1) = CDbl(strArr(4))
    Next loopCount
    'TransAmount = CDbl(strArr(4))
    'Call InsertPigmyAmount(AccID, TransDate, TransAmount)
    'Update the Grid
    With grd
        If .rows = .Row + 1 Then .rows = .rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = CStr(.Row)
        .Col = 1: .Text = FormatField(m_RstPigmy("AccNum"))
        .Col = 2: .Text = FormatField(m_RstPigmy("Name"))
        .Col = 3: .Text = FormatField(m_RstPigmy("FullName"))
        AccountTotal = 0
        For loopCount = 4 To (3 + NumOfDays)
            .Col = loopCount: .Text = FormatCurrency(CCur(TransAmount(loopCount - 4)))
            AccountTotal = AccountTotal + CCur(TransAmount(loopCount - 4))
        Next
        .Col = loopCount: .Text = FormatCurrency(AccountTotal)
    End With
    
  Else
    'reading the header
    strArr = Split(strData, ",")
    ''Get the AgenID as
    cmdSave.Enabled = True
    If gAgentID <> CInt(strArr(4)) Then
        MsgBox "The file does not belongs to this pigmy agent", vbOKOnly, "Index 2000"
        cmdSave.Enabled = False
        'Exit Sub
    End If
    
    'lastAccId = CLng(strArr(2))
    If InStr(1, strArr(5), ".") Then
        strArr(5) = Replace(strArr(5), ".", "/")
    End If
    TransDate = GetSysFormatDate(strArr(5))
    
    'NumOfDays = CSng(strArr(8))
    NumOfDays = 1
    NumOfRecord = CSng(strArr(3))
    Set m_RstPigmy = GetRecordSet(CInt(strArr(4)))
    ReDim Preserve TransAmount(NumOfDays)
    ReDim Preserve TotalAmount(NumOfDays)
    Call InitGrid(TransDate, NumOfDays)
    grd.Row = 0
    If (UBound(strArr)) > 6 Then
        strArrWhole = Split(strData, vbLf)
        ArrayCount = UBound(strArrWhole)
    End If
  End If
NextRecord:
  lineCount = lineCount + 1
  If ArrayCount > lineCount Then
    strData = strArrWhole(lineCount)
    GoTo NextLine
  End If
Loop
Close #iFileNo
    
End Sub



