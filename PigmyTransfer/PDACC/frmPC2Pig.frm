VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPC2Pig 
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   10
      Cols            =   5
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Transfer"
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   6360
      Width           =   2055
   End
End
Attribute VB_Name = "frmPC2Pig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_LastAccountId As String
Dim m_NoOFAccounts As Integer

Private Sub InitGrid()
With grd
        .rows = 2
        .cols = 6
        'If chkAgent.Value = vbChecked Then .Cols = .Cols + 1
        .Row = 0
        .FixedRows = 1
        .FixedCols = 1
        .Col = 0: .Text = LoadResString(gLangOffSet + 33) '"sL No"
        '.Col = 1: .Text = LoadResString(gLangOffSet + 330) & " " & LoadResString(gLangOffSet + 35) ' Agent Name
        .Col = 1: .Text = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)  '"Account No"
        .Col = 2: .Text = LoadResString(gLangOffSet + 35) '"Name"
        .Col = 3: .Text = "English Name"
        .Col = 4: .Text = LoadResString(gLangOffSet + 281)  ' Opening Date
        .Col = 5: .Text = LoadResString(gLangOffSet + 42)  ' Balance
    End With
    

Dim Rst As Recordset
Set Rst = GetRecordSet(gAgentID)

grd.Visible = False
Dim SlNo As Integer
SlNo = 1
With grd
    
    While Not Rst.EOF
        .rows = .rows + 1
        .Row = SlNo
        .Col = 0: .Text = Format(SlNo, "00"): .ColWidth(0) = 400
        .Col = 1: .Text = FormatField(Rst("AccNum")): .ColWidth(1) = 800
        .Col = 2: .Text = FormatField(Rst("Name")): .ColWidth(2) = 2400
        .Col = 3: .Text = FormatField(Rst("FullName")): .ColWidth(3) = 2400
        .Col = 4: .Text = FormatField(Rst("CreateDate")): .ColWidth(4) = 1200
        .Col = 5: .Text = FormatField(Rst("Balance")): .ColWidth(5) = 1200
         m_LastAccountId = FormatField(Rst("AccNum"))
         SlNo = SlNo + 1
         Rst.MoveNext
    Wend
End With
m_NoOFAccounts = SlNo - 1
grd.Visible = True
End Sub

Private Sub cmdExport_Click()

Dim strPigmyType As String

If gDEVICE = "BALAJI_OLD" Then
    CreateBalajiOLDOutput
ElseIf gDEVICE = "BALAJI" Then
    CreateBalajiOutput
Else
    CreatePrathinidhi
End If
Call gDbTrans.CloseDB

Unload Me

End Sub

Private Sub Form_Load()
    Call gDbTrans.OpenDB(gDBFileName, constDBPWD)
    SetKannadaCaption
    InitGrid
End Sub

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)
End Sub
Private Sub CreateBalajiOLDOutput()

Dim sFileText As String
Dim iFileNo As Integer
Dim strData As String
Dim strTemp As String
Dim rowNo As Integer
Dim m_PigmyType As String
Dim strName As String



'Print the Header
'1) Branch Code 4 digit
'2) Agent code 3 digit
'3) Last A/c no 6 digits
'4) Start Date 8 char
'5) 0000 4 digit
'6) 00000000 8 digit
'7) Password 8 digit
'8) No Of Entries 3 digits
'9) No Of Days 1 digit
    strTemp = m_LastAccountId

 strData = "$" + "C" + "0" + "," + "0" + "," + "00000" + "," + "0000000000000000000" + "," + "0000000.00" + "," + "00000000000000000000" + "," + IIf(gAgentID < 10, "000000000", "00000000") + CStr(gAgentID) + ","

strData = strData + Format(Now, "DD.MM.YY") + "," + "1234512345" + "@"
strTemp = CStr(m_NoOFAccounts)
iFileNo = FreeFile

Open App.Path & "\PCTX.DAT" For Output As #iFileNo

'Write the header of file
Print #iFileNo, strData

Dim Rst As Recordset
Set Rst = GetRecordSet(gAgentID)
    
While Not Rst.EOF
'1) A/c no 6 digit
'2) 000000 6 digit
'3) 000000 6 digit
'4) 000000 6 digit
'5) 000000 6 digit
'6) 000000 6 digit
'7) 000000 6 digit
'8) 000000 6 digit
'9) Name 16 char
'10) Opening bal. 7 digits
'11) Opn.date 8 digits


strName = FormatField(Rst("FullName"))
If Len(Trim$(strName)) < 1 And gLangOffSet <> wis_KannadaOffset Then strName = FormatField(Rst("Name"))
If Len(Trim$(strName)) < 1 Then strName = "Account No:" + FormatField(Rst("AccNum"))

strTemp = CStr(FormatField(Rst("AccNum")))

m_PigmyType = FormatField(Rst("PigmyType"))
If m_PigmyType = "LOAN" Then
    strData = "000" + "," + "-" + "," + Mid(m_PigmyType, 1, 5) + String(5 - Len(Mid(m_PigmyType, 1, 5)), " ") + ","
Else
   strData = "000" + "," + "+" + "," + Mid(m_PigmyType, 1, 5) + String(5 - Len(Mid(m_PigmyType, 1, 5)), " ") + ","
End If

'strData = String(6 - Len(strTemp), "0") + strTemp + ","
'strData = strData + String(6, "0") + ","
'strData = strData + String(6, "0") + ","
'strData = strData + String(6, "0") + ","
'strData = strData + String(6, "0") + ","
'strData = strData + String(6, "0") + ","
'strData = strData + String(6, "0") + ","
'strData = strData + String(6, "0") + ","

''Name or Account nUm
    If Len(Trim$(strTemp)) < 1 Then strTemp = FormatField(Rst("AccID"))
    strTemp = Left$(strTemp, 19)
''Name or Account num
strData = strData + strTemp + String(19 - Len(strTemp), " ") + "," + "0000000.00" + "," + Mid(strName, 1, 20) + String(20 - Len(Mid(strName, 1, 20)), " ") + ","
    strTemp = CStr(FormatField(Rst("Balance")) \ 1)
    strTemp = Right(strTemp, 7)
strData = strData + String(7 - Len(strTemp), "0") + strTemp + ".00" + ","
strData = strData + Format(Rst("CreateDate"), "dd.mm.yy") + ","
strData = strData + "0000000.00" + "@"
     
''Move to the next recrod
     Rst.MoveNext
'Update the file
'Write #iFileNo, strData
Print #iFileNo, strData

Wend
Print #iFileNo, Chr(4)

Close #iFileNo

MsgBox "File for the pigmy export has created", vbOKOnly, "Index 2000"
On Error Resume Next

End Sub


Private Sub CreateBalajiOutput()

Dim sFileText As String
Dim iFileNo As Integer
Dim strData As String
Dim strTemp As String
Dim rowNo As Integer
Dim m_PigmyType As String
Dim strName As String
Dim strAccNum As String


'LINE LENGHT  56 chars

'Print the Header
'1) Agent's last account        6 digit
'2) Total No Of Receipet        6 digit
'3) Total Collect Amount        16 digit Last A/c no 6 digits
'4) Bank(3) & agent(3) code     6 digit
'5) Collection(current) date    8 digit
'6) Password 8 digit(4 key pad & 4 holidya) 8 digit


'7) Password 8 digit
    

strData = ""
strTemp = m_LastAccountId
If Len(strTemp) > 6 Then strTemp = strTemp & Left$(1, 6)
If Len(strTemp) < 6 Then strTemp = String(6 - Len(strTemp), " ") & strTemp
strData = strTemp                                           '1) 6 digit Agent's last account
strData = strData & "," & Format(m_NoOFAccounts, "000000")  '2) 6 digit Total No Of Receipet
strData = strData & "," & "000000          "                '3) 16 digit Total Collect Amount        16 digit Last A/c no 6 digits
strData = strData & ",000" & Format(gAgentID, "000")        '4) 6 digit Bank(3) & agent(3) code     6 digit
strData = strData & "," & Format(Now, "DD.MM.YY")           '5) Collection(current) date    8 digit
strData = strData & "," & "12341234"                        '6) Password 8 digit(4 key pad & 4 holidya) 8 digit

iFileNo = FreeFile

Open App.Path & "\PCTX.DAT" For Output As #iFileNo

'Write the header of file
Print #iFileNo, strData

Dim Rst As Recordset
Set Rst = GetRecordSet(gAgentID)
    
While Not Rst.EOF
    strTemp = CStr(FormatField(Rst("AccNum")))
    
    m_PigmyType = FormatField(Rst("PigmyType"))
    
    strAccNum = CStr(FormatField(Rst("AccNum")))
    
    'Replace any Comma's with #
    strName = FormatField(Rst("FullName"))
    strName = Replace(strName, ",", "#")
    
    If Len(Trim$(strName)) < 1 And gLangOffSet <> wis_KannadaOffset Then strName = FormatField(Rst("Name"))
    If Len(Trim$(strName)) < 1 Then strName = "Account No:" + String(7 - Len(strAccNum), " ") & strAccNum
    strName = Left(strName, 16)
    If Len(strName) < 16 Then strName = strName & String(16 - Len(strName), " ")
    
    strData = String(6 - Len(strAccNum), " ") & strAccNum   '1)6 digit A/c no
    strData = strData & "," & "000000"                      '2) 6 digit 2,3 days collection
    strData = strData & "," & strName '3) 16 digit Customer Name
    
    strTemp = CStr(FormatField(Rst("Balance")) \ 1)
    strTemp = Right(strTemp, 6)
    If Len(strTemp) < 6 Then strTemp = String(6 - Len(strTemp), "0") & strTemp
    strData = strData & "," & strTemp   '4) 6 digit Balance after collection
    
    strTemp = IIf(IsNull(Rst("TransDate").Value), Format(Now, "DD.MM.YY"), Format(Rst("TransDate"), "DD.MM.YY"))
    strData = strData & "," & strTemp   '5) 6 digit Collection Date
    strData = strData & "," & "000000"  '6) 6 digit Collection amount of a day
    If Len(strData) < 56 Then _
     strData = strData & String(56 - Len(strData), " ") '7) 1 digit Last one should be space
    
    ''Move to the next recrod
    Rst.MoveNext
    'Update the file
    'Write #iFileNo, strData
    Print #iFileNo, strData

Wend
Print #iFileNo, Chr(4)

Close #iFileNo

MsgBox "File for the pigmy export has created", vbOKOnly, "Index 2000"
On Error Resume Next

End Sub



Private Sub CreatePrathinidhi()

Dim maxRecords As Integer
Dim sFileText As String
Dim iFileNo As Integer
Dim strData As String
Dim strTemp As String
Dim recordNo As Integer

strTemp = String$(512, 0)
strTemp = ReadFromIniFile("Pigmy", "PigmyCount", App.Path & "\" & constFINYEARFILE)
maxRecords = CInt(strTemp)
If maxRecords <= 0 Then maxRecords = 600
strTemp = ""
'Print the Header
'1) Branch Code 4 digit
'2) Agent code 3 digit
'3) Last A/c no 6 digits
'4) Start Date 8 char
'5) 0000 4 digit
'6) 00000000 8 digit
'7) Password 8 digit
'8) No Of Entries 3 digits
'9) No Of Days 1 digit
    strTemp = m_LastAccountId
strData = "0000" + "," + IIf(gAgentID < 10, "00", "0") + CStr(gAgentID) + ","
strData = strData + String(6 - Len(strTemp), "0") + strTemp + ","
strData = strData + Format(Now, "DD/MM/YY") + ","
strData = strData + "0000" + "," + "00000000" + "," + "24081406" + ","
strTemp = IIf(m_NoOFAccounts > maxRecords, CStr(maxRecords), CStr(m_NoOFAccounts))
strData = strData + String(3 - Len(strTemp), "0") + strTemp + "," + "7,"

strData = strData + String(89 - Len(strData), "0")
iFileNo = FreeFile

Open App.Path & "\PC_2_PIG.DAT" For Output As #iFileNo
'Write the header of file
Print #iFileNo, strData

Dim Rst As Recordset
Set Rst = GetRecordSet(gAgentID)
recordNo = 0
While Not Rst.EOF
'Write #iFileNo "\r\n"
recordNo = recordNo + 1
'1) A/c no 6 digit
'2) 000000 6 digit
'3) 000000 6 digit
'4) 000000 6 digit
'5) 000000 6 digit
'6) 000000 6 digit
'7) 000000 6 digit
'8) 000000 6 digit
'9) Name 16 char
'10) Opening bal. 7 digits
'11) Opn.date 8 digits
strTemp = CStr(FormatField(Rst("AccNum")))
strData = String(6 - Len(strTemp), "0") + strTemp + ","
strData = strData + String(6, "0") + ","
strData = strData + String(6, "0") + ","
strData = strData + String(6, "0") + ","
strData = strData + String(6, "0") + ","
strData = strData + String(6, "0") + ","
strData = strData + String(6, "0") + ","
strData = strData + String(6, "0") + ","
''Name or Account nUm
    strTemp = FormatField(Rst("FullName"))
    If Len(Trim$(strTemp)) < 1 And gLangOffSet <> wis_KannadaOffset Then strTemp = FormatField(Rst("Name"))


    If Len(Trim$(strTemp)) < 1 Then strTemp = "Account No:" + FormatField(Rst("AccNum"))
    strTemp = Left$(strTemp, 16)
''Name or Account num
strData = strData + strTemp + String(16 - Len(strTemp), " ") + ","
    strTemp = CStr(FormatField(Rst("Balance")) \ 1)
    strTemp = Right(strTemp, 7)
strData = strData + String(7 - Len(strTemp), "0") + strTemp + ","
strData = strData + Format(Rst("CreateDate"), "dd/mm/yy")
     
''Move to the next recrod
     Rst.MoveNext
'Update the file
'Write #iFileNo, strData
Print #iFileNo, strData

If recordNo >= maxRecords Then GoTo CloseFile
Debug.Assert recordNo <> 357
Wend

CloseFile:
Close #iFileNo

MsgBox "File for the pigmy export has created", vbOKOnly, "Index 2000"
On Error Resume Next

Dim appid
appid = Shell("Prati-Nidhi66.exe 1", vbNormalFocus)
AppActivate appid, True

End Sub

Private Sub Form_Terminate()
If Not gDbTrans Is Nothing Then gDbTrans.CloseDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not gDbTrans Is Nothing Then gDbTrans.CloseDB
End Sub
