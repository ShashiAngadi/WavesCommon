VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTab 
   Caption         =   "             Creating TabFiles From DataBase"
   ClientHeight    =   2460
   ClientLeft      =   2415
   ClientTop       =   1845
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdClick 
      Caption         =   "click"
      Height          =   345
      Left            =   3180
      TabIndex        =   9
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Creating Tab Files"
      Height          =   2025
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.CommandButton cmdCretae 
         Caption         =   "&Create"
         Height          =   285
         Left            =   3930
         TabIndex        =   6
         Top             =   1620
         Width           =   705
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   1620
         Width           =   765
      End
      Begin VB.CommandButton cmdTabPath 
         Caption         =   "....."
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Top             =   1140
         Width           =   345
      End
      Begin VB.TextBox txtTabFilePath 
         Height          =   285
         Left            =   1530
         TabIndex        =   3
         Top             =   1140
         Width           =   2025
      End
      Begin VB.CommandButton cmdFileName 
         Caption         =   "....."
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   420
         Width           =   345
      End
      Begin VB.TextBox txtDbFileName 
         Height          =   285
         Left            =   1530
         TabIndex        =   1
         Top             =   420
         Width           =   2025
      End
      Begin VB.Label Label2 
         Caption         =   "Tab File Name  :"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Database Path  :"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   450
         Width           =   1845
      End
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   1620
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_CreaDBClass As clsTransact
Attribute m_CreaDBClass.VB_VarHelpID = -1
Private Sub cmdCancel_Click()
'gCancel = True
Unload Me
End Sub
Private Sub cmdClick_Click()
If gDBTrans Is Nothing Then Set gDBTrans = New clsTransact
If gDBTrans.OpenDB(txtDbFileName, "WIS!@#") Then
    If Not gDBTrans.CreateTabFile(txtTabFilePath) Then
        MsgBox "NOt done"
    Else
        MsgBox "DONE"
    End If
End If
End Sub
Private Sub cmdCretae_Click()
Dim IniFile As String
'Dim strDbName As String
'Check for the DataBase existance
If Trim(Me.txtDbFileName.Text) = "" Then
    MsgBox "you have not specified the database", vbExclamation, "DB ERROR"
    Exit Sub
ElseIf Dir(txtTabFilePath.Text, vbDirectory) = "" Then
    If MsgBox("Specified Does not exists " & vbCrLf & "Do you want to Create the path ", vbInformation + vbYesNo, _
                "DB Path Error") = vbNo Then Exit Sub
            If Not MakeDirectories(txtTabFilePath.Text) Then
        MsgBox "Error in creating the path " & txtTabFilePath _
            & " for " & "DBName", vbCritical
        Exit Sub
        'GoTo dbCreate_err
    End If
End If

IniFile = txtDbFileName

'Set Db = OpenDatabase(IniFile, True, False)

''''''''Code Of Shashi Starts Here
'Ins this code i've not made any validations

Dim ws As Workspace

'Open the DataBase

'Set Db = OpenDatabase(IniFile, True, False)
'Set Db = ws.OpenDatabase(IniFile, True, False, "WIS!@#")

'Call the Funcion
'Call GenerateTabFile
  
'''Code of shashi ends here
  
  Dim DBPath As String
  DBPath = Trim$(txtDbFileName.Text)
    If m_CreaDBClass Is Nothing Then
        Set m_CreaDBClass = New clsTransact
    End If
   
    ' Now Set the PathOf DataBase In tab File
    Dim strRet As String
    Dim NewStr As String
    Dim I As Integer
    If Trim$(txtDbFileName) = "" Then
        MsgBox " Please Specify the mdb File to create TabFiles"
        Exit Sub
    End If
    IniFile = txtDbFileName
    
    Do
        I = I + 1
        strRet = ReadFromIniFile("DataBases", "DataBase" & I, IniFile)
        If strRet = "" Then Exit Do
       
        'Now set the DataBase Path To then Ini file
        NewStr = putToken(strRet, "DBPath", Trim(DBPath))
        
        '(To the Database)call WriteToMdbFile("TabFiles","TabFile" & i(KeyDAta),StringValue,MdbFile)
        Call WriteToIniFile("DataBases", "DataBase" & I, NewStr, IniFile)
     Loop
    If m_CreaDBClass.CreateDB(IniFile, "WIS!@#") Then
        'Label2.Caption = "Created DataBase"
    Else
        'Label2.Caption = "Error in Creating DataBase"
    End If

On Error Resume Next

Call m_CreaDBClass.CloseDB
Set m_CreaDBClass = Nothing
End Sub
Private Sub cmdFileName_Click()
    cdb.InitDir = App.Path
    cdb.Filter = " mdb Files (*.mdb) | *.mdb"
    cdb.DialogTitle = "Choose mdbfile for the Tabfile creation"
    cdb.CancelError = False
    cdb.ShowOpen
    txtDbFileName = cdb.FileName
End Sub
Private Sub cmdTabPath_Click()
'HERE Better to show  CaommonaDailogbox(ShoeSavae)
'Beacause here you have selecting the file not the
Dim TabFile As String
With cdb
    .CancelError = False
    .Filter = "Tab File (*.Tab) | *.* |All Files (*.*)| *.*"
    .FilterIndex = 1
    .DefaultExt = "tab"
    .DialogTitle = "Save the Tab File"
    .ShowOpen
        TabFile = .FileName
End With
txtTabFilePath = ""
If TabFile = "" Then Exit Sub

txtTabFilePath = TabFile

'frmpath.Show vbModal
'txtTabFilePath = frmpath.txtPath
Unload frmpath
End Sub

Private Sub Form_Load()
cmdCretae.Enabled = False
End Sub


