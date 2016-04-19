VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCreatDb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4290
   ClientLeft      =   2100
   ClientTop       =   1860
   ClientWidth     =   6915
   Icon            =   "CreatDb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6915
   Begin VB.Timer Timer1 
      Left            =   -300
      Top             =   -150
   End
   Begin ComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   20
      Top             =   3765
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   926
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   90
      TabIndex        =   4
      Top             =   360
      Width           =   6735
      Begin VB.OptionButton optTabFile 
         Caption         =   "Create Tab File"
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   150
         Width           =   1455
      End
      Begin VB.OptionButton optRunQuery 
         Caption         =   "Run Query.."
         Height          =   255
         Left            =   2940
         TabIndex        =   5
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton optCreate 
         Caption         =   "CreateDataBase"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   450
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optDesign 
         Caption         =   "Compare && Re-Design DataBase"
         Height          =   225
         Left            =   2940
         TabIndex        =   0
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2865
      Left            =   90
      TabIndex        =   3
      Top             =   930
      Width           =   6735
      Begin VB.CommandButton cmdoptions 
         Caption         =   "O&ptions"
         Height          =   345
         Left            =   30
         TabIndex        =   23
         Top             =   2490
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   5940
         TabIndex        =   19
         Top             =   2520
         Width           =   765
      End
      Begin VB.TextBox txtQuery 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "CreatDb.frx":030A
         Top             =   1230
         Width           =   4815
      End
      Begin VB.CheckBox chkRelation 
         Caption         =   "Check &Relation"
         Height          =   315
         Left            =   4980
         TabIndex        =   13
         Top             =   2175
         Width           =   1875
      End
      Begin VB.CheckBox chkIndex 
         Caption         =   "Check &Indexes"
         Height          =   315
         Left            =   3000
         TabIndex        =   12
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox txtDBPath 
         Height          =   285
         Left            =   1695
         TabIndex        =   11
         Top             =   585
         Width           =   3960
      End
      Begin VB.CommandButton cmdDbPath 
         Caption         =   "..."
         Height          =   300
         Left            =   5790
         TabIndex        =   10
         Top             =   555
         Width           =   330
      End
      Begin VB.CommandButton cmdTabPath 
         Caption         =   "..."
         Height          =   300
         Left            =   5790
         TabIndex        =   9
         Top             =   210
         Width           =   330
      End
      Begin VB.TextBox txtTabPath 
         Height          =   285
         Left            =   1695
         TabIndex        =   8
         Top             =   240
         Width           =   3960
      End
      Begin VB.CommandButton cmdDbFile 
         Caption         =   "..."
         Height          =   300
         Left            =   5775
         TabIndex        =   7
         Top             =   900
         Width           =   330
      End
      Begin VB.TextBox txtDBFile 
         Height          =   285
         Left            =   1695
         TabIndex        =   6
         Top             =   915
         Width           =   3960
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   5070
         TabIndex        =   2
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblQuery 
         Caption         =   "Write your Query"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1365
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Data Base Path"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblTab 
         Caption         =   "Tab file "
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lblDbFile 
         Caption         =   "Data Base name "
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   1005
         Width           =   1245
      End
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   450
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDataBase 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   22
      Top             =   -30
      Width           =   5655
   End
End
Attribute VB_Name = "frmCreatDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CreaDBClass As clsTransact
Attribute m_CreaDBClass.VB_VarHelpID = -1
Private m_LblQueryTitle As String

Private WithEvents m_frmReport As frmQueryRes
Attribute m_frmReport.VB_VarHelpID = -1

Private m_TestProp As clsCommon
Private Function DropTable() As Boolean

Dim tabName As String
   
   'open the mdb
   If Trim$(txtDBFile.Text) = "" Then
        MsgBox "Please specify the database name", vbInformation, wis_MESSAGE_TITLE
        Exit Function
   End If
   
   If Not gDBTrans.OpenDB(txtDBFile, "WIS!@#") Then
       MsgBox ""
       Exit Function
   End If
   
   'Declare variables
    Dim Sqlstr As String * 100
    Dim ExtraChar As String
    Dim Fetch As Boolean
    Dim Rst As Recordset
    
    Me.MousePointer = vbHourglass
    Sqlstr = Me.txtQuery.Text
    
    If InStr(1, Sqlstr, "SELECT", vbTextCompare) Then Fetch = True
    
    Sqlstr = RemoveString(Sqlstr, ExtraChar)
    
    'Set m_frmReport = New frmqueryRes
    'Load m_frmReport
    If Fetch Then
        gDBTrans.SQLStmt = Sqlstr  'Before this remove vbcrlf
        If gDBTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
           MsgBox "No Records to Display"
           Exit Function
        End If
    Else
        gDBTrans.BeginTrans
        gDBTrans.SQLStmt = Sqlstr
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Function
        End If
        gDBTrans.CommitTrans
    End If
    
DropTable = True
End Function
Private Sub PutintoGrd()
Dim Rst As Recordset
' 'Here Show the result set in a grid
' 'Load frmqueryRes
' Dim Count As Integer
' 'Dim RowCount As Integer
'
'    With frmQueryRes.grd
'        .Clear
'        .Rows = Rst.RecordCount + 1
'        .FixedRows = 1
'        .Row = 0
'        .Cols = Rst.fields.Count
'
'         For Count = 0 To Rst.fields.Count - 1
'            .Col = Count
'            .Text = Rst.fields(Count).Name
'            .CellAlignment = 6
'            .CellFontBold = True
'        Next
'
'        While Not Rst.EOF
'          .Row = .Row + 1
'           For Count = 0 To Rst.fields.Count - 1
'                .Col = Count
'                .Text = FormatField(Rst(Count))
'            Next
'           Rst.MoveNext
'        Wend
'    End With
End Sub
Private Function Redesign_Comapre() As Boolean
Dim IniFile As String

If optDesign Then
    If Trim$(txtTabPath) = "" Then
        MsgBox " Please Specify the Tab File to check Database"
        Exit Function
    End If
    
    IniFile = txtTabPath
    'IniFile = InputBox("Enter the Name file to create Database", "Create Access DataBase", App.Path & "\Index 2000.tab")
    If Dir(IniFile, vbNormal) = "" Then
        MsgBox "Invalid file name", vbCritical, "Check DataBase"
        Exit Function
    End If
    
    If m_CreaDBClass Is Nothing Then
        Set m_CreaDBClass = New clsTransact
    End If
    'Check for the mdb file for comparision...(Should not be empty)
    If Trim$(txtDBFile) = "" Then
        MsgBox " Please Specify the mdb file to compare with the Tab file", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
    'Check for the strcture of database
    If m_CreaDBClass.CheckDbStructure(IniFile, txtDBFile, "WIS!@#", chkIndex.Value, chkRelation.Value) Then
        MsgBox "DataBase Rectified"
    Else
         MsgBox "Error in Database checking Check the database properly", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
 End If
End Function
Private Function RemoveString(strString As String, stringToRemove As String) As String

Dim strTemp As String
Dim strRemove As String

strTemp = strString
strRemove = stringToRemove
Dim Pos As Integer
Do
    Pos = InStr(Pos + 1, strTemp, vbCrLf)
    If Pos = 0 Then Exit Do
    strTemp = Left(strTemp, Pos) & Mid(strTemp, Pos + Len(strRemove))
Loop

RemoveString = strTemp
End Function

Private Function RunQuery() As Boolean

Dim tabName As String
   'frmqueryRes.Show vbModal
   If gDBTrans Is Nothing Then Set gDBTrans = New clsTransact
   
   'open the mdb
   If Trim$(txtDBFile.Text) = "" Then
        MsgBox "Please specify the database name", vbInformation, wis_MESSAGE_TITLE
        Exit Function
   End If
   
   If Not gDBTrans.OpenDB(txtDBFile, "WIS!@#") Then
       MsgBox ""
       Exit Function
   End If
   
   'Declare variables
    Dim Sqlstr As String * 100
    Dim ExtraChar As String
    Dim Fetch As Boolean
    Dim Rst As Recordset
    
    Sqlstr = Me.txtQuery.Text
    
    If InStr(1, Sqlstr, "SELECT", vbTextCompare) Then Fetch = True
    
    Sqlstr = RemoveString(Sqlstr, ExtraChar)
    'Set m_frmReport = New frmqueryRes
    'Load m_frmReport
    
    If Fetch Then
        gDBTrans.SQLStmt = Sqlstr  'Before this remove vbcrlf
        If gDBTrans.SQLFetch < 1 Then
           MsgBox "No Records to Display"
           Exit Function
        End If
        Set Rst = gDBTrans.Rst.Clone
    Else
        gDBTrans.BeginTrans
        gDBTrans.SQLStmt = Sqlstr
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Function
        End If
        gDBTrans.CommitTrans
    End If
    
'Check for the query
If InStr(1, Sqlstr, "SELECT", vbTextCompare) Then
    frmQueryRes.Show vbModal
End If

'Set mousepointer as default
Me.MousePointer = vbDefault

End Function

Private Sub cmdCancel_Click()
Unload Me
'gCancel = True
End Sub


Private Sub cmdDbFile_Click()
    cdb.InitDir = App.Path
    cdb.Filter = " Database Files (*.mdb) | *.mdb"
    cdb.DialogTitle = "Open the mdb file"
    cdb.CancelError = False
    cdb.ShowOpen
    txtDBFile = cdb.Filename
End Sub

Private Sub cmdDbPath_Click()
frmpath.Show vbModal
txtDBPath = frmpath.txtPath
Unload frmpath

End Sub



Private Sub cmdok_Click()
Dim IniFile As String

If gDBTrans Is Nothing Then Set gDBTrans = New clsTransact

'if option is creating the tabfile.
If optTabFile Then
    If gDBTrans.OpenDB(txtDBFile, "WIS!@#") Then
        'For Checking Purpose
        If Trim(Me.txtTabPath.Text) = "" Then
            MsgBox "you have not specified the path for Tab File Creation ", vbExclamation, "PATH ERROR"
            Exit Sub
        End If
        If Not gDBTrans.CreateTabFile(txtTabPath) Then
            MsgBox "Error in creating the tab file"
            Else
            MsgBox "Tab file created succsesfully", vbInformation, wis_MESSAGE_TITLE
            Exit Sub
        End If
     End If
End If

'if Option is Creating the database
Dim DBPath As String

If m_CreaDBClass Is Nothing Then Set m_CreaDBClass = New clsTransact

If optCreate Then
        'Check for the DataBase Path
        If Trim(Me.txtDBPath.Text) = "" Then
            MsgBox "you have not specified the path for database file Creation ", vbExclamation, "PATH ERROR"
            Exit Sub
         ElseIf Dir(txtDBPath.Text, vbDirectory) = "" Then
           If MsgBox("Specified Does not exists " & vbCrLf & "Do you want to Create the path ", vbInformation + vbYesNo, _
                       "DB Path Error") = vbNo Then Exit Sub
            End If
        
        If Not MakeDirectories(txtDBPath.Text) Then
             MsgBox "Error in creating the path " & txtDBPath _
             & " for " & "DBName", vbCritical
             Exit Sub
                 'GoTo dbCreate_err
        End If
 
    DBPath = Trim$(txtDBPath.Text)
    ' Now Set the PathOf DataBase In tab File
    Dim strRet As String
    Dim NewStr As String
    Dim I As Integer
    
    If Trim$(txtTabPath) = "" Then
        MsgBox " Please Specify the Tab File to create Database"
        Exit Sub
    End If
    IniFile = txtTabPath
    'IniFile = InputBox("Enter the Name file to create Database", "Create Access DataBase", App.Path & "\Index 2000.tab")
    If Dir(IniFile, vbNormal) = "" Then
        MsgBox "Invalid file name", vbCritical, "Create DataBase"
        End
    End If
    Do
        I = I + 1
        strRet = ReadFromIniFile("DataBases", "database" & I, IniFile)
        If strRet = "" Then Exit Do
        'Now set the DataBase Path To then Ini file
        NewStr = putToken(strRet, "DBPath", Trim(DBPath))
        Call WriteToIniFile("DataBases", "DataBase" & I, NewStr, IniFile)
    Loop
    
    If m_CreaDBClass.CreateDB(IniFile, "WIS!@#") Then
           ' msgbox "Database Created"
        Else
            'msgbox Not Created"
    End If
End If
'Added Every functios in the one structure
'if gdbtrans is nothing then set gdbtrans=new clsTransact
If optDesign Then
    If Not Redesign_Comapre Then
        'MsgBox "Succsefully Completed ", vbInformation, wis_MESSAGE_TITLE
    End If
End If

'If Option is for Running the query then Do the Query Operations
If optRunQuery Then
   If Not RunQuery Then Exit Sub
End If

On Error Resume Next

Call m_CreaDBClass.CloseDB
Set m_CreaDBClass = Nothing

MsgBox "Completed"

End Sub
Private Sub cmdDrop_Click()
Dim tabName As String
Dim DBName As String

If gDBTrans Is Nothing Then Set gDBTrans = New clsTransact

'Open the Database
If gDBTrans.OpenDB(txtDBFile, "WIS!@#") Then
    If Trim$(txtDBFile.Text) = "" Then
            MsgBox "Specify the Database name to drop the table", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    DBName = txtDBFile
    tabName = Me.txtQuery.Text
    
    If Not gDBTrans.DropTable(DBName, tabName) Then
            MsgBox "Unable to drop the table from the MDB"
        Else
            MsgBox "Table dropped succsesfully", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    End If
End Sub

Private Sub CmdEnter_Click()
Frame1.Visible = True
Frame2.Visible = True
'CmdEnter.Enabled = False
End Sub


Private Sub cmdoptions_Click()
frmOptions.Show vbModal
End Sub





Private Sub cmdTabPath_Click()
        
    Dim TabFile As String
    

    cdb.InitDir = App.Path
    cdb.Filter = "Tab Files (*.Tab)|*.tab|All Files(*.*)|*.*"
    cdb.DefaultExt = "tab"
    cdb.FilterIndex = 1
    cdb.DialogTitle = "Open the Tab File"
    cdb.CancelError = False
    cdb.ShowOpen
        
    TabFile = cdb.Filename
    
    If TabFile = "" Then Exit Sub
    
    txtTabPath = TabFile
End Sub


Private Sub Form_Load()

Dim LocalTest As String

Set m_TestProp = New clsCommon

LocalTest = m_TestProp.GetProperty

lblDataBase.Caption = LocalTest

cmdoptions.Enabled = False
'When Loaded Make All form UnVisible
stBar.SimpleText = ""
'Make option write to enabled
cmdOk.Enabled = False
Call optCreate_Click

End Sub
Private Sub m_CreaDBClass_CreateDBStatus(strMsg As String, CreatedDBRatio As Single)
If CreatedDBRatio > 0 And CreatedDBRatio <= 1 Then
End If
Me.Label1.Caption = strMsg
Me.Refresh
End Sub

Private Sub m_frmReport_Initailise(Min As Long, Max As Long)
If Max <> 0 And Max > Min Then
    frmCancel.prg.Visible = True
    frmCancel.prg.Min = Min
    If Max > 32000 Then
        frmCancel.prg.Max = 32000
    Else
        frmCancel.prg.Max = Max
    End If
End If

End Sub

Private Sub m_frmReport_Processing(strMessage As String, Ratio As Single)
frmCancel.lblMessage = "PROCESS :" & vbCrLf & strMessage
If Ratio > 0 Then
    If Ratio > 1 Then
        frmCancel.prg.Value = Ratio
    Else
        frmCancel.prg.Value = frmCancel.prg.Max * Ratio
    End If
End If
End Sub
Private Sub optCreate_Click()
frmCreatDb.Caption = "Creating the DataBase "
    If optCreate Then
        stBar.SimpleText = "Select Tabfile, form that tabfile" & _
                           vbCrLf & "create Database"
        cmdOk.Default = True
        txtDBPath.Enabled = True
        cmdDbPath.Enabled = True
        txtDBFile.Enabled = False
        cmdDbFile.Enabled = False
        cmdOk.Enabled = True
        chkIndex.Enabled = False
        chkRelation.Enabled = False
        txtQuery.Enabled = False
    End If
End Sub
Private Sub optDesign_Click()
frmCreatDb.Caption = "Re-designing the database file with tab File"
    If optDesign Then
        stBar.SimpleText = "Select Database Name and tabfile name" & _
                           "to compare design & structure"
                           
        txtDBFile.Enabled = True
        cmdDbFile.Enabled = True
        txtTabPath.Enabled = True
        cmdTabPath.Enabled = True
        
        txtDBPath.Enabled = False
        cmdDbPath.Enabled = False
        
        cmdOk.Enabled = True
        chkIndex.Enabled = False
        chkRelation.Enabled = False
        
        txtQuery.Enabled = False
        
    End If
End Sub

Private Sub optRunQuery_Click()
frmCreatDb.Caption = "Write your query "
    If optRunQuery Then
        stBar.SimpleText = "Select Database name and write" & _
                           "Your query in the QueryTextBox"
        
        txtQuery.Enabled = True
        cmdOk.Enabled = True
        txtDBFile.Enabled = True
        cmdDbFile.Enabled = True
        
        cmdDbPath.Enabled = False
        txtTabPath.Enabled = False
        cmdTabPath.Enabled = False
        txtDBPath.Enabled = False
        cmdDbPath.Enabled = False
        chkIndex.Enabled = False
        chkRelation.Enabled = False
    End If
End Sub

Private Sub optTabfile_Click()
frmCreatDb.Caption = "Creating the tab File"
    If optTabFile Then
        stBar.SimpleText = "Select existing Database name" & _
                           "and create tab file"
        cmdOk.Default = True
        txtTabPath.Enabled = True
        cmdTabPath.Enabled = True
        
        txtDBPath.Enabled = False
        cmdDbPath.Enabled = False
        cmdOk.Enabled = True
        cmdDbFile.Enabled = True
        txtDBFile.Enabled = True
        chkRelation.Enabled = False
        chkIndex.Enabled = False
        
        txtQuery.Enabled = False
        
    End If
End Sub

