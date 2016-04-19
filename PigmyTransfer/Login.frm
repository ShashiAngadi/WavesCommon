VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please identify yourself"
   ClientHeight    =   2790
   ClientLeft      =   2775
   ClientTop       =   2955
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4890
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   90
      TabIndex        =   8
      Top             =   60
      Width           =   4665
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   1560
         Width           =   2265
      End
      Begin VB.ComboBox cmbFinancialYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   5
         Text            =   "cmbFinancialYear"
         ToolTipText     =   "Select the Finanicial Year You want to Explore"
         Top             =   1100
         Width           =   2265
      End
      Begin VB.TextBox txtUserPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   670
         Width           =   2265
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label lblDate 
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Image img 
         Height          =   480
         Left            =   150
         Picture         =   "Login.frx":0000
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblUserDate 
         Caption         =   "Finanacial Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   1095
         Width           =   1455
      End
      Begin VB.Label lblUserPassword 
         Caption         =   "User password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label lblUserName 
         Caption         =   "User name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   2370
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2310
      TabIndex        =   6
      Top             =   2370
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LoginClicked(UserName As String, _
                            Userpassword As String, _
                            LoginDate As String, _
                            UnloadDialog As Boolean)

Public Event CancelClicked()
Public Event FinYearSelected(ByVal YearID As Integer)
Public Event FinYearChanged(ByVal YearID As Integer)

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

lblUserName.Caption = LoadResString(gLangOffSet + 151) & " " & LoadResString(gLangOffSet + 35)
lblUserPassword.Caption = LoadResString(gLangOffSet + 151) & " " & LoadResString(gLangOffSet + 153)
lblUserDate.Caption = LoadResString(gLangOffSet + 37)
lblDate.Caption = LoadResString(gLangOffSet + 38) & " " & LoadResString(gLangOffSet + 37)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
cmdLogin.Caption = LoadResString(gLangOffSet + 151)

End Sub
'
'This subroutine will add the Financial Year to the ComboBox
'From the External file FinYear.fin File
'WRITTEN By Lingappa Sindhanur
'DATED   "June 19, 2002
Private Sub GetFinancialYear()
Dim YearID As Long

YearID = cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)
   
Call LoadFinYearData(App.Path & "\" & constFINYEARFILE, YearID)
Dim Rst As Recordset
If DayBeginDate = "" Then DayBeginDate = GetSysFormatDate(gStrDate)

'Now Initilise the Global varible
'Now get the Name of the Bank /Society from DataBase
gDbTrans.SQLStmt = "Select * from CompanyCreation"
If gDbTrans.Fetch(Rst, adOpenStatic) > 0 Then _
            gCompanyName = FormatField(Rst("CompanyName"))
    
Dim retstr As String
retstr = ReadSetupValue("General", "ONLINE", "False")
gOnLine = IIf(UCase(retstr) = "FALSE", False, True)
DateFormat = UCase(ReadSetupValue("General", "DateFormat", "dd/mm/yyyy"))


End Sub
Private Sub LoadAdmin()
txtUserName.Text = "admin"
txtUserPassword.Text = "admin"
If cmbFinancialYear.ListIndex < 0 Then cmbFinancialYear.ListIndex = cmbFinancialYear.ListCount - 1

Call cmdLogin_Click

End Sub

Private Function Validated() As Boolean
Validated = False

If txtUserName.Text = "" Then
    MsgBox "Please enter the User Name"
    Exit Function
End If
If txtUserPassword.Text = "" Then
    MsgBox "Please enter the password"
    Exit Function
End If

If cmbFinancialYear.ListIndex = -1 Then
   MsgBox "Please Select the Current financial Year"
   Exit Function
End If
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date of transaction not in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If

Validated = True
End Function

Private Sub cmdCancel_Click()

RaiseEvent CancelClicked

Unload Me
End Sub

Private Sub cmdLogin_Click()

Dim UnloadDialog As Boolean
Dim YearID As Integer
Dim DBPath As String

If Not Validated Then Exit Sub

Me.MousePointer = vbHourglass

YearID = cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)
gStrDate = txtDate.Text

If txtUserName.Enabled = True Then
    RaiseEvent FinYearSelected(YearID)
    RaiseEvent LoginClicked(Trim$(txtUserName.Text), Trim$(txtUserPassword.Text), "", UnloadDialog)
    If UnloadDialog Then
        GetFinancialYear
        
        DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "Server")
        
        If DBPath = "" Then
            DBPath = FilePath(GetDBNameWithPath(App.Path & "\" & constFINYEARFILE, YearID))
        End If
        
        Me.MousePointer = vbDefault
        Unload Me
        
    End If
Else
    
    RaiseEvent FinYearChanged(YearID)
    
    Call GetFinancialYear

    Me.MousePointer = vbDefault
    Unload Me
    
End If

Me.MousePointer = vbDefault

End Sub
Private Sub Form_Load()

Call SetKannadaCaption


Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

Me.Caption = "Please identify yourself"

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

Call GetFinYearData
cmbFinancialYear.ListIndex = cmbFinancialYear.ListCount - 1
cmbFinancialYear.Enabled = False
txtDate = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing
End Sub

Private Sub img_Click()
'LoadAdmin
End Sub

Private Sub txtUserName_GotFocus()
ActivateTextBox txtUserName
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If KeyAscii = 17 Then LoadAdmin
End Sub

Private Sub txtUserPassword_GotFocus()
ActivateTextBox txtUserPassword
End Sub

Public Function GetFinYearData() As Boolean

'Trap an error
On Error GoTo ErrLine

'Declare Variables
Dim strFinYearFile As String
Dim I As Integer

Dim FromYear As String
Dim ToYear As String
Dim DBPath As String
Dim strRet As String

Dim encrKey As String
Dim encrSection As String

Const strFromYear = "FromYear#"
Const strToYear = "ToYear#"
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"

GetFinYearData = False

strFinYearFile = App.Path & "\" & constFINYEARFILE

encrKey = strYear & 1 'EncryptData(strYear & 1)
encrSection = strFinYearSection 'EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

'If strRet = "" Then CreateDBFirstTime
    
'Call SaveFinYear(strFinYearFile, True)

Me.cmbFinancialYear.Clear

I = 1
Do
   
    ' Read the dbname from datafile.
    encrKey = strYear & I 'EncryptData(strYear & i)
    encrSection = strFinYearSection 'EncryptData(strFinYearSection)
    
    strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)
    
    If strRet = "" Then Exit Do
    
    GetFinYearData = True
    
    FromYear = ExportExtractToken(strRet, strFromYear, , ",")
    DBPath = ExportExtractToken(strRet, strDBPath, , ",")
    
    ToYear = GetSysFormatDate("31/3/" & CStr(Year(FromYear) + 1))
    cmbFinancialYear.AddItem "April " & Year(CDate(FromYear)) & " TO March " & Year(CDate(ToYear))
    cmbFinancialYear.ItemData(cmbFinancialYear.NewIndex) = I
    
    I = I + 1
Loop

Exit Function

ErrLine:
    MsgBox "GetFinYearData()" & vbCrLf & Err.Description
    'Resume
End Function
Public Sub LoadFinYearData(ByVal strFinYearFile As String, ByVal YearID As Integer)
'Declare the constansts
Const strFromYear = "FromYear#"
Const strToYear = "ToYear#"
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"


'Declare the variables
Dim encrKey As String
Dim encrSection As String
Dim strRet As String

encrKey = strYear & YearID ' EncryptData(strYear & YearID)
encrSection = strFinYearSection 'EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then Exit Sub

'strRet = DecryptData(strRet)

FinIndianFromDate = ExportExtractToken(strRet, strFromYear, , ",")
FinIndianEndDate = ExportExtractToken(strRet, strToYear, , ",")
gStrDate = txtDate
Exit Sub

ErrLine:
    MsgBox "GetDBNameWithPath()" & vbCrLf & Err.Description

End Sub

