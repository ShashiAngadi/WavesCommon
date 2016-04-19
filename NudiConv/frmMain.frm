VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSuchiNudi 
      Caption         =   "Convert to Nudi"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   2160
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDatabase 
      Caption         =   "Database Convert"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmbNUm 
      Caption         =   "Num to Text"
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdNudiSuchi 
      Caption         =   "Convert to Suchi"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdKeyboard 
      Caption         =   "keyboard"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtNudi 
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Text            =   "¤Ð·ÐÔì"
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox txtSuchi 
      BeginProperty Font 
         Name            =   "SUCHI-KAN-0850"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   5415
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Convert Files"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtTarget 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "SUCHI-KAN-0850"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.CheckBox chkKannada 
      Caption         =   "English to Kannada"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblFont 
      Caption         =   "Fonts"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strfontname As String
Dim lRetVal As Long
Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002 'HKEY_LOCAL_MACHINE
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Function ConvertToSuchi(strNudiInput As String) As String
'Dim strData As String
'Dim strOutData As String
Dim strInput As String
Dim strOutput As String

    strOutput = Space$(200)
    strInput = Space$(200)
    strInput = Trim$(strNudiInput)
    Call CONVERTDATA(Pass1, Pass2, strInput, strOutput, 7, 61, 18)     ' SUCHI to NUDI Conversion
    'strOutData = strOutData & ";" & Trim$(strOutput)

ConvertToSuchi = Trim$(strOutput)

End Function
Private Function ConvertToNudi(strSuchiInput As String) As String
'Dim strData As String
'Dim strOutData As String
Dim strInput As String
Dim strOutput As String

    strOutput = Space$(200)
    strInput = Space$(200)
    strInput = Trim$(strSuchiInput)
    Call CONVERTDATA(Pass1, Pass2, strInput, strOutput, 61, 7, 18)    ' SUCHI to NUDI Conversion
    'strOutData = strOutData & ";" & Trim$(strOutput)

ConvertToNudi = Trim$(strOutput)

End Function

Private Sub chkKannada_Click()
    If chkKannada.value Then
        Call SHREE_SETSCRIPT(Pass1, Pass2, 0) ' set script as english
        txtSource.FontName = strEnglishFont
        txtTarget.FontName = strKannadaFont
    Else
        Call SHREE_SETSCRIPT(Pass1, Pass2, glScriptCode)
        txtSource.FontName = strKannadaFont
        txtTarget.FontName = strEnglishFont
    End If
End Sub

Private Sub cmbFont_Change()
    If chkKannada.value = vbUnchecked Then txtSource.FontName = cmbFont.Text
    
    strKannadaFont = cmbFont.Text
    
End Sub

Private Sub cmbFont_Click()
    If chkKannada.value = vbUnchecked Then txtSource.FontName = cmbFont.Text
    
    strKannadaFont = cmbFont.Text
    
End Sub

Private Sub cmbNUm_Click()
    If chkKannada And Not IsNumeric(Trim$(txtSource.Text)) Then Exit Sub
    
Dim inte As Double
Dim str1 As String
  str1 = Space$(200)
  inte = Val(txtSource.Text)
  'API to set script
  Call SHREE_SETSCRIPT(Pass1, Pass2, glScriptCode)
  'API to set Shree font type
  Call SHREE_SETFONTTYPE(Pass1, Pass2, glScriptCode, 18)
  
  'API call to convert number to words
  Call SHREE2000_NUM_TO_WORDS(Pass1, Pass2, inte, str1, glScriptCode, 0, 0)
  txtTarget = str1
End Sub

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdConvert_Click()
If chkKannada Then
   txtTarget.Text = ConvertToKannada(txtSource.Text)
Else
    txtTarget.Text = ConvertToEnglish(txtSource.Text)
End If

End Sub

Private Sub cmdDatabase_Click()

Dim dbName As String

With cdb
    .CancelError = False
    .FileName = ""
    .InitDir = "C:\Program Files\Index 2000"
    .Filter = "Data Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
    .DialogTitle = "Select the Old Data Base"
    .ShowOpen
    dbName = .FileName
End With
    Call ChangeDBFont(dbName)
End Sub

Private Sub cmdKeyboard_Click()
    Call SAMHITA.SHREE_KBD_SETUP(Pass1, Pass2)
End Sub

Private Sub Command2_Click()
    txtSuchi.Text = ConvertNudiToSuchita(txtNudi.Text)
End Sub

Private Sub ConvertFile(strFileName As String)
    
'out file
Dim opFIleName As String

Dim inFileNo As Integer
Dim outFileNo As Integer
Dim strData As String
Dim strOutData As String
Dim Pos As Integer
Dim subStr() As String
Dim loopCount As Integer

inFileNo = FreeFile

Open strFileName For Input As #inFileNo

outFileNo = FreeFile

opFIleName = Replace$(strFileName, "temp", "temp1", , , vbTextCompare)
'Open "c:\temp\output.txt" For Output As #outFileNo
Open opFIleName For Output As #outFileNo

Dim strInput As String
Dim strOutput As String
    
    strOutput = Space$(200)
    strInput = Space$(200)

Call INIT_CONVERT
        
Do While Not EOF(inFileNo)
  'Input #iFileNo, strData
  Line Input #inFileNo, strData
  If Len(Trim$(strData)) = 0 Then
    Print #outFileNo, strData
    GoTo NextLine
  End If
    Pos = InStr(1, strData, ";", vbTextCompare)
    If Pos > 0 Then
        subStr = Split(strData, ";")
        strOutData = ""
        For loopCount = 0 To UBound(subStr)
            strOutput = Space$(200)
            strInput = Space$(200)
            strInput = Trim$(subStr(loopCount))
            Call CONVERTDATA(Pass1, Pass2, strInput, strOutput, 7, 61, 18)     ' SUCHI to NUDI Conversion
            strOutData = strOutData & ";" & Trim$(strOutput)
        Next
        strOutData = Mid(strOutData, 2)
    Else
        strOutput = Space$(500)
        strInput = Space$(500)
        strInput = Trim$(strData)
        strOutData = ""
        
    'Call CONVERTDATA(PASS1, PASS2, str_Renamed1, str2, 7, 18, 61)      ' SUCHI to NUDI Conversion
    
        Call CONVERTDATA(Pass1, Pass2, strInput, strOutput, 7, 61, 18)     ' SUCHI to NUDI Conversion
        strOutData = Trim$(strOutput)
    End If
   Print #outFileNo, strOutData

NextLine:
 
Loop
Close #inFileNo
Close #outFileNo
End Sub


Private Sub CheckMergeFile(strFileName As String)
    
'out file
Dim opFIleName As String

Dim inFileNo As Integer
Dim inFileNo_2 As Integer
Dim outFileNo As Integer
Dim strData As String
Dim strResData As String
Dim strOutData As String
Dim Pos As Integer
Dim Pos_2 As Integer
Dim strInt As String
Dim resInt As Integer

Dim subStr() As String
Dim loopCount As Integer

inFileNo = FreeFile
Open strFileName For Input As #inFileNo

        
Dim check As Boolean

Do While Not EOF(inFileNo)
  'Input #iFileNo, strData
    Line Input #inFileNo, strData
    If Len(Trim$(strData)) = 0 Then GoTo NextLine
    If Trim$(strData) = "BEGIN" Then check = True
    If Not check Then GoTo NextLine
  
    strData = Trim$(strData)
    Pos = InStr(1, strData, " ", vbTextCompare)
    If Pos > 0 Then
        strInt = Mid(strData, 1, Pos)
        resInt = CInt(Val(strInt))
        If resInt > 0 Then
            Line Input #inFileNo, strResData
            Line Input #inFileNo, strOutData
            strResData = Mid(Trim$(strResData), 2, 3)
            strOutData = Mid(Trim$(strOutData), 2, 3)
            
            If CInt(strResData) = CInt(strOutData) And CInt(strOutData) = resInt Then
                GoTo NextLine
            Else
                MsgBox "ERROR"
            End If
            
        End If
    
    End If
   
NextLine:
 
Loop
Close #inFileNo

End Sub




Private Sub MergeFile(strFileName As String, strFileName_2 As String)
    
'out file
Dim opFIleName As String

Dim inFileNo As Integer
Dim inFileNo_2 As Integer
Dim outFileNo As Integer
Dim strData As String
Dim strResData As String
Dim strOutData As String
Dim Pos As Integer
Dim Pos_2 As Integer
Dim strInt As String
Dim resInt As Integer

Dim subStr() As String
Dim loopCount As Integer

inFileNo = FreeFile
Open strFileName For Input As #inFileNo

inFileNo_2 = FreeFile
Open strFileName_2 For Input As #inFileNo_2

outFileNo = FreeFile

opFIleName = Replace$(strFileName, ".rc", "_Merge.rc", , , vbTextCompare)
'Open "c:\temp\output.txt" For Output As #outFileNo
Open opFIleName For Output As #outFileNo

Dim strInput As String
Dim strOutput As String
strOutput = Space$(200)
strInput = Space$(200)
        

Do While Not EOF(inFileNo)
  'Input #iFileNo, strData
    Line Input #inFileNo, strData
    If Len(Trim$(strData)) = 0 Then
      Print #outFileNo, strData
      GoTo NextLine
    End If
    
    'Check whether the line has Kannada stuff '5000
    strOutData = Trim$(strData)
    
    Pos = InStr(1, strData, " ", vbTextCompare)
    If Pos > 0 Then
        strInt = Mid(strData, 1, Pos)
        resInt = CInt(Val(strInt))
        If resInt >= 5000 Then
            If resInt > 5349 And resInt < 5351 Then
            resInt = 5350
            End If
            ''FInd the Same Value in other Input FIle
            Do While Not EOF(inFileNo_2)
                Line Input #inFileNo_2, strResData
                strResData = Trim$(strResData)
                Pos_2 = InStr(1, strResData, " ", vbTextCompare)
                If Pos_2 > 0 Then
                    strInt = Mid(strResData, 1, Pos_2)
                    If resInt = CInt(Val(strInt)) Then
                        strResData = CStr((resInt - 3000)) & Mid(strResData, 5)
                        Print #outFileNo, strResData
                        Exit Do
                    End If
                End If
            Loop
            strOutData = CStr((resInt - 1000)) & Mid(strData, 5)
        End If
    
    End If
    
   Print #outFileNo, strOutData

NextLine:
 
Loop
Close #inFileNo
Close #inFileNo_2
Close #outFileNo
End Sub



Private Function GetStringsFile() As String()
On Error Resume Next
Dim strFileName As String
Dim retValue() As String

Dim inFileNo As Integer

Dim strData As String
Dim subStr() As String
Dim loopCount As Integer

inFileNo = FreeFile

strFileName = "C:\Temp1\Indx2000.txt"

Open strFileName For Input As #inFileNo


Dim strInput As String
Dim strOutput As String
    
    strOutput = Space$(200)
    strInput = Space$(200)

Call INIT_CONVERT
loopCount = -1
Dim cmbSort As ComboBox
cmbSort.Clear
Do While Not EOF(inFileNo)
    
  'Input #iFileNo, strData
  Line Input #inFileNo, strData
  If Len(Trim$(strData)) = 0 Then GoTo NextLine
  
    subStr = Split(strData, " ")
    If CInt(subStr(0)) > 5000 Then
        loopCount = loopCount + 1
        ReDim Preserve retValue(loopCount)

        strOutput = Space$(200)
        strInput = Space$(200)
        strInput = Trim$(subStr(1))
        'Call CONVERTDATA(Pass1, Pass2, strInput, strOutput, 7, 61, 18)     ' SUCHI to NUDI Conversion
        
        retValue(loopCount) = Trim$(strOutput)
        cmbSort.AddItem (Trim$(strInput))
        cmbSort.ItemData(cmbSort.NewIndex) = Trim$(strInput)
    End If

NextLine:
 
Loop
Close #inFileNo


cmbSort.FontName = strKannadaFont

GetStringsFile = retValue()



End Function



Private Sub cmdFile_Click()
Dim fso As New FileSystemObject

Dim fld As Folder
Dim fil As File
Set fld = fso.GetFolder("C:\temp")
Dim fileNames() As String
Dim loopCount As Integer
For Each fil In fld.Files
    loopCount = loopCount + 1
    ReDim Preserve fileNames(loopCount)
    fileNames(loopCount) = fld.Path & "\" & fil.Name
Next

Set fil = Nothing
Set fld = Nothing
Set fso = Nothing
For loopCount = 1 To UBound(fileNames)
    Call ConvertFile(fileNames(loopCount))
    MsgBox "File " & fileNames(loopCount) & " converted", vbOKOnly, "Conversion"
Next
MsgBox "All files converted", vbOKOnly, "Conversion"
End Sub

Private Sub cmdNudiSuchi_Click()
    Dim strInput As String
    strInput = Trim$(txtNudi.Text)
    If Len(strInput) > 0 Then txtSuchi.Text = ConvertToSuchi(Trim$(txtNudi.Text))
    
End Sub

Private Sub Command1_Click()
    Dim file1 As String
    Dim file2 As String
    file1 = "C:\Projects\Index 2000 Total New\RESOURCE\indx2000.rc"
    file2 = "C:\Projects\Index 2000 Total New\RESOURCE\indx2000_Nudi.rc"
    'Call MergeFile(file1, file2)
    
    file2 = "C:\Projects\Index 2000 Total New\RESOURCE\indx2000_Merge.rc"
    Call CheckMergeFile(file2)
    
    MsgBox "File " & file1 & " merged", vbOKOnly, "Conversion"
End Sub

Private Sub cmdSuchiNudi_Click()

    If Len(Trim$(txtSuchi.Text)) > 0 Then txtNudi.Text = ConvertToSuchi(txtSuchi.Text)
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    
  Dim strSamhitaPath As String
    
  Dim strTmpScriptList As String
  Dim strScriptList As String
  Dim strCommaPos As String
  Dim strScriptName As String
  Dim strScriptCode As String
    
  Dim lSize As Long
  Dim Cnt As Integer
        
 'Allocate space for variable
  strTmpScriptList = Space$(500)
  strScriptList = Space$(500)
  strCommaPos = Space$(500)
  strScriptName = Space$(500)
  strScriptCode = Space$(500)
  strfontname = Space$(100)
  lSize = 11111111


  
Call InitFonts
Call InitializeSamhita

    
  'To read the registry path to know where actually samhita get installed
  'strSamhitaPath = GetRegString("SOFTWARE\Modular InfoTech\Shree Samhita\2.0", "SamhitaPath", HKEY_LOCAL_MACHINE)
  strSamhitaPath = "C:\Smhita20\"
  'Procedure to be called to set the main dictionary for transliteration.
  'SET_TRANS_MAINDICT (strSamhitaPath + "kandict.dbf")
    
  'API Call to get the list of Scripts present
  'Call GET_SHREE_INSTALLED_SCRIPTS(PASS1, PASS2, strTmpScriptList, lSize)
      
  'Loop to add List of scripts in Combo box
  'strScriptList = Trim(strTmpScriptList)
   'strScriptCode = Mid(strScriptList, 1, Len(strScriptList))
  ' Call SHREE_SCRIPT_TO_STR(strScriptCode, strScriptName)

    
 'API call for getting script code of the first default script
 'lScriptCode = SHREE_STR_TO_SCRIPT(cmbSetScript.Text)
 'lScript = 7
'MsgBox "Changed", vbOKOnly

 
'load Fonts
AddBilingualFont cmbFont
strKannadaFont = cmbFont.Text
txtSource.FontName = strKannadaFont

'Call GetStringsFile

 Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    'In case of error, Deactivate samhita by calling Close_Shree API call
    CLOSE_SHREE
    End
End Sub


Private Sub Form_Load1()
'Set Kannada script code
lScript = 7
glScriptCode = 7
lFontType = 0
Pass1 = 73412761 'CLng(Trim$(InputBox("Password 1", "Samhita Passwords")))
Pass2 = 651917425 'CLng(Trim$(InputBox("Password 2", "Samhita Passwords")))
  
Call InitFonts

txtSuchi.FontName = strKannadaFont

lRetVal = START_SHREE2(Pass1, Pass2)
If lRetVal <> 0 Then
     MsgBox "Error Initialising Shree-Samhita : " + CStr(lRetVal)
End If

'Loading Transliteration Dll
  lRetVal = LOADTRANSLITERATION(Pass1, Pass2)
  If lRetVal <> 0 Then
     MsgBox "Error Initialising transliteration : " + CStr(lRetVal)
  End If
  
  'To read the registry path to know where actually samhita get installed
  Dim strSamhitaPath As String
  'strSamhitaPath = GetRegString("SOFTWARE\Modular InfoTech\Shree Samhita\2.0", "SamhitaPath", HKEY_LOCAL_MACHINE)
  
  'Procedure to be called to set the main dictionary for transliteration.
  'SET_TRANS_MAINDICT (strSamhitaPath + "kandict.dbf")
  
Call INIT_CONVERT

'Loading Transliteration Dll
'lRetval = LOADTRANSLITERATION(Pass1, Pass2)
    
'Set the Kannada font Lay out to KGP
'API call to set font layout
Call SAMHITA.SHREE_SETFONTTYPE(Pass1, Pass2, glScriptCode, SUCHI2000)
        
txtSource.FontName = strKannadaFont

'set the Keyboard to ENG passing Language Script :7
 'Call SHREE_SET_KEYBOARD(Pass1, Pass2, glScriptCode, "ENG")
 lRetVal = SHREE_SET_KEYBOARD(Pass1, Pass2, glScriptCode, "KGP")
 
 
    'API Call to Set script
    lRetVal = SHREE_SETSCRIPT(Pass1, Pass2, glScriptCode)
    'Set the FontType to SUCHI
    Call SHREE_SETFONTTYPE(Pass1, Pass2, glScriptCode, 15)
    

End Sub

Private Sub txtNudi_Change()
txtSuchi.Text = ""
End Sub

