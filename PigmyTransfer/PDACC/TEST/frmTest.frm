VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   1995
   ClientTop       =   1875
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdAkToNu 
      Caption         =   "Convert"
      Height          =   315
      Left            =   3870
      TabIndex        =   12
      Top             =   1620
      Width           =   1815
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "?"
      TabIndex        =   10
      Top             =   1620
      Width           =   2205
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   300
      Left            =   1290
      TabIndex        =   8
      Top             =   1110
      Width           =   4335
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "..."
      Height          =   300
      Left            =   5700
      TabIndex        =   7
      Top             =   1110
      Width           =   465
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   300
      Left            =   5700
      TabIndex        =   6
      Top             =   630
      Width           =   465
   End
   Begin VB.TextBox txtFilePath 
      Height          =   300
      Left            =   1290
      TabIndex        =   4
      Top             =   630
      Width           =   4335
   End
   Begin VB.CommandButton cmdAscii 
      Caption         =   "Ascii"
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   3
      Top             =   3750
      Width           =   1200
   End
   Begin VB.CommandButton cmdTestFile 
      Caption         =   "Test File"
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   2
      Top             =   3270
      Width           =   1200
   End
   Begin VB.CommandButton cmdTabs 
      Caption         =   "List Tables"
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   1
      Top             =   2790
      Width           =   1200
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Enabled         =   0   'False
      Height          =   300
      Left            =   330
      TabIndex        =   0
      Top             =   2310
      Width           =   1200
   End
   Begin VB.Label lblConversion 
      Height          =   300
      Left            =   1590
      TabIndex        =   13
      Top             =   150
      Width           =   3375
   End
   Begin VB.Label lblPassWrd 
      Caption         =   "Password :"
      Height          =   270
      Left            =   210
      TabIndex        =   11
      Top             =   1650
      Width           =   1005
   End
   Begin VB.Label lblSaveAs 
      Caption         =   "Save As :"
      Height          =   300
      Left            =   210
      TabIndex        =   9
      Top             =   1140
      Width           =   825
   End
   Begin VB.Label lblPath 
      Caption         =   "File Path :"
      Height          =   300
      Left            =   210
      TabIndex        =   5
      Top             =   660
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gDBTrans As clsDBUtils
Private Function AkrToNudFile(strSrcFile As String, strDestFile As String) As Boolean
Dim File1Number As Integer
Dim File2Number As Integer
Dim LinesFromFile As String
Dim NextLine As String
Dim strOutPut As String
Dim i As Long
Dim J As Long

File1Number = FreeFile                    ' Get unused file number
Open strDestFile For Binary As #File1Number   ' Create temp HTML file
'    Write #FileNumber, "<HTML> <\HTML>"  ' Output text
'Close #FileNumber                        ' Close file

File2Number = FreeFile                    ' Get unused file number
Open strSrcFile For Input As #File2Number    ' Create temp HTML file

J = 1
Do Until EOF(File2Number)
   Line Input #File2Number, NextLine
   LinesFromFile = NextLine + Chr(13) + Chr(10)
'   Debug.Print LinesFromFile
    strOutPut = AkrutiToNudiStr(LinesFromFile)
    If strOutPut = "" Then
'        Debug.Print LinesFromFile
    Else
        For i = 1 To Len(strOutPut)
            Put #File1Number, J, Mid$(strOutPut, i, 1)
            J = J + 1
        Next
    End If
'    Write #File1Number, LinesFromFile  ' Output text
Loop

Close #File1Number                        ' Close file
Close #File2Number                        ' Close file

AkrToNudFile = True
End Function

Private Function AkStrToNu(strSource As String) As String
Dim strRet As String
Dim StLen As Integer
Dim Item As Integer
Dim tempItem As Integer
Dim NewSt As String
Dim strDest As String
Dim strFinal As String
Dim strTemp As String
Dim strInterTemp As String
Dim Count As Long

'strRet = ReadFromIniFile("Line1", "Line1", "C:\Windows\Desktop\File1.txt")

strRet = Trim$(strSource)
StLen = Len(strSource)

Count = 1
'Debug.Assert StLen <= 2
For Item = 1 To StLen
    NewSt = Mid(strRet, 1, StLen)
    If strInterTemp = "" Then strInterTemp = NewSt
    strDest = ReadFromIniFile("Source", NewSt, "C:\Test\Akruti2.txt")
    strTemp = ReadFromIniFile("Destination", strDest, "C:\Test\Nudi2.txt")
    tempItem = Item
    Do While strTemp = "" And Item <= StLen And StLen > 1
        NewSt = Mid(strInterTemp, 1, StLen - Item)
        strDest = ReadFromIniFile("Source", NewSt, "C:\Test\Akruti2.txt")
        strTemp = ReadFromIniFile("Destination", strDest, "C:\Test\Nudi2.txt")
        Item = Item + 1
        strInterTemp = NewSt
    Loop
    If strTemp <> "" Then
        strFinal = strFinal + strTemp
    Else
        Item = tempItem
        strFinal = strFinal + strInterTemp
    End If
    Count = Len(strInterTemp)
    strRet = Right$(strRet, StLen - Count)
    StLen = Len(strRet)
    strInterTemp = ""
    Item = Count + 1
Next

'Debug.Print strFinal & vbCr; vbLf

AkStrToNu = strFinal

End Function

Private Function AkrutiToNudiStr(strSource As String) As String
Dim strRet As String
Dim StLen As Long
Dim Item As Long
Dim lngAkrAscii As Long
Dim strAkrSingle As String
Dim strNudiSingle As String
Dim lngAkrAscii1 As Long
Dim strToReturn As String
Dim lngNudiAscii As Long
Dim strArrNudi() As String
Dim i As Long
Dim chrNudiFound As Boolean

strRet = Trim$(strSource)
StLen = Len(strSource)

For Item = 1 To StLen
    strAkrSingle = Mid(strRet, Item, 1)
'    Debug.Assert Item < 300
    If strAkrSingle = "" Then Exit For
    lngAkrAscii = Asc(strAkrSingle)
    If lngAkrAscii > 128 Then
        strNudiSingle = LoadResString(lngAkrAscii)
        If Len(strNudiSingle) > 3 Then
            GetStringArray strNudiSingle, strArrNudi, "+"
            strNudiSingle = ""
            For i = 0 To UBound(strArrNudi)
                strNudiSingle = strNudiSingle + Chr(strArrNudi(i))
            Next
            chrNudiFound = True
        Else
            lngNudiAscii = CLng(strNudiSingle)
        End If
    End If
    If Not chrNudiFound Then
        If lngNudiAscii = 0 Then
            strNudiSingle = Chr(lngAkrAscii)
        Else
            strNudiSingle = Chr(lngNudiAscii)
        End If
        If lngAkrAscii = 144 Or lngAkrAscii = 138 Or _
           lngAkrAscii = 193 Or lngAkrAscii = 199 Or lngAkrAscii = 226 Or _
           lngAkrAscii = 220 Or lngAkrAscii = 227 Or lngAkrAscii = 224 _
           Then
            Item = Item + 1
            strAkrSingle = Mid(strRet, Item, 1)
            If strAkrSingle <> "" Then lngAkrAscii1 = Asc(strAkrSingle)
            If (lngAkrAscii1 = 152 Or lngAkrAscii1 = 207 Or _
                lngAkrAscii1 = 239 Or lngAkrAscii1 = 255) Then
                Select Case lngAkrAscii
                    Case 144
                        strNudiSingle = Chr(173)
                    Case 138
                        strNudiSingle = Chr(182)
                    Case 193
                        strNudiSingle = Chr(181)
                    Case 199
                        strNudiSingle = Chr(172)
                    Case 220
                        strNudiSingle = Chr(232)
                    Case 227
                        strNudiSingle = Chr(242)
                    Case 224
                        strNudiSingle = Chr(237)
                    Case 226
                        strNudiSingle = Chr(240)
                End Select
            Else
                Item = Item - 1
            End If
'        ElseIf lngAkrAscii = 244 Then
'            Item = Item + 2
'            strAkrSingle = Mid(strRet, Item, 1)
'            If strAkrSingle <> "" Then lngAkrAscii1 = Asc(strAkrSingle)
'            If lngAkrAscii1 = 255 Then
'                strNudiSingle = Chr(165)
'                Item = Item - 2
'            Else
'                Item = Item - 1
'            End If
        End If
    End If
    strToReturn = strToReturn + strNudiSingle
    lngAkrAscii = 0
    lngNudiAscii = 0
    chrNudiFound = False
Next

AkrutiToNudiStr = strToReturn
'AkrutiToNudiStr = NudiBilingual2MonolingualGlyph(strToReturn)
End Function
Private Function AkToNu(strTbl As String) As Boolean

Dim rst As ADODB.Recordset
Dim fld As ADODB.Field
Dim ArrFld() As String
Dim ArrFldVal() As String
Dim i As Long
Dim J As Long
Dim strTxtFld As String
Dim strSQLStr As String

gDBTrans.SQLStmt = "SELECT TOP 1 * FROM " & strTbl
If gDBTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    For Each fld In rst.Fields
        If fld.Type = adVarWChar Then
'            Debug.Print fld.Type & " " & _
                fld.Value & ";" '& vbCr; vbLf
            strTxtFld = AddQuotes(fld.Value, True)
        End If
        i = i + 1
        ReDim Preserve ArrFld(i)
        ReDim Preserve ArrFldVal(i)
        ArrFld(i) = fld.Name
        ArrFldVal(i) = IIf(fld.Type = adVarWChar, strTxtFld, fld.Value)
    Next
End If

strSQLStr = "UPDATE " & strTbl & " SET "
For J = 1 To i
    strSQLStr = strSQLStr & ArrFld(J) & " = " & ArrFldVal(J) & ", "
Next

strSQLStr = Left$(strSQLStr, Len(strSQLStr) - 2)
strSQLStr = strSQLStr & " WHERE "
For J = 1 To i
    strSQLStr = strSQLStr & ArrFld(J) & " = " & ArrFldVal(J) & " AND "
Next

strSQLStr = Left$(strSQLStr, Len(strSQLStr) - 5)

Debug.Print strSQLStr
AkToNu = True

End Function

Private Function AkToNuTbl(strTbl As String) As Boolean

Dim rst As ADODB.Recordset
Dim InterRst As ADODB.Recordset
Dim fld As ADODB.Field
Dim strTxtFld As String
Dim strFieldName As String
Dim strFieldValue As String

Debug.Print strTbl
gDBTrans.SQLStmt = "SELECT TOP 1 * FROM " & strTbl
If gDBTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    For Each fld In rst.Fields
        If fld.Type = adVarWChar And Not fld.Name = "Module" _
            And Not fld.Name = "Key" And Not fld.Name = "Password" Then
'            Debug.Print fld.Type & " " & _
                fld.Value & ";" '& vbCr; vbLf
            gDBTrans.SQLStmt = "SELECT DISTINCT " & fld.Name & _
                " FROM " & strTbl
            If gDBTrans.Fetch(InterRst, adOpenForwardOnly) > 0 Then
                strFieldName = fld.Name
                gDBTrans.BeginTrans
                Do While Not InterRst.EOF
                    If Not IsNull(InterRst.Fields(strFieldName)) And InterRst.Fields(strFieldName) <> " " Then
                        strFieldValue = CStr(InterRst.Fields(strFieldName))
                        strTxtFld = AddQuotes(AkrutiToNudiStr(strFieldValue), True)
                        gDBTrans.SQLStmt = "UPDATE " & strTbl & _
                            " SET " & fld.Name & " = " & strTxtFld & _
                            " WHERE " & fld.Name & " = " & AddQuotes(strFieldValue, True)
                        If Not gDBTrans.SQLExecute Then
                            gDBTrans.RollBack
                            Exit Function
                        End If
                    End If
                    InterRst.MoveNext
                Loop
                gDBTrans.CommitTrans
            End If
        End If
    Next
End If

AkToNuTbl = True

End Function
Private Function AkToNuDataBase(DBName As String, pwd As String) As Boolean

Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table

If Not gDBTrans.OpenDB(DBName, pwd) Then
    MsgBox "Please Enter Password", vbInformation
     Exit Function
End If

Set cat = New Catalog
Set cat.ActiveConnection = gDBTrans.GetActiveConnection
StartTimer
For Each tbl In cat.Tables
   If tbl.Type = "TABLE" Then
        If Not AkToNuTbl(tbl.Name) Then
            gDBTrans.CloseDB
            Exit Function
        End If
   End If
Next

StopTimer
gDBTrans.CloseDB
AkToNuDataBase = True

End Function

Private Sub cmdAkToNu_Click()
Dim strSource As String
Dim strDest As String
Dim strPassWrd As String

strSource = txtFilePath.Text
strDest = txtSaveAs.Text

strPassWrd = "PRAGMANS"
If Trim$(txtPWD.Text) <> "" Then strPassWrd = Trim$(txtPWD.Text)

If strSource = "" Or strDest = "" Then
    MsgBox "select sourcefile to convert the database"
    ActivateTextBox txtFilePath
    Exit Sub
End If
If Dir$(strDest) <> "" Then Kill (strDest)
If txtPWD.Visible Then
    Call FileCopy(strSource, strDest)
    If Not AkToNuDataBase(strDest, strPassWrd) Then
        MsgBox "Cannot Convert"
        Exit Sub
    End If
    MsgBox "Database converted Successfully", vbInformation
Else
    StartTimer
    If Not AkrToNudFile(strSource, strDest) Then
        MsgBox "Cannot Convert"
        Exit Sub
    End If
    StopTimer
    MsgBox "File converted Successfully", vbInformation
End If
End Sub

Private Sub cmdAscii_Click()
Dim FileName As String
Dim File1Number As Integer
Dim i As Long

FileName = "C:\Test\Temp.txt"
File1Number = FreeFile                    ' Get unused file number
Open FileName For Output As #File1Number   ' Create temp HTML file
For i = 0 To 255
    Write #File1Number, i & "  =  " & Chr(i)  ' Output text
Next
Close #File1Number                        ' Close file

End Sub

Private Sub cmdPath_Click()
Dim strTemp As String
Dim strExtn As String
frmpath.Show vbModal

txtFilePath.Text = frmpath.txtPath.Text

If txtFilePath.Text = "" Then Exit Sub
strTemp = Left$(frmpath.txtPath.Text, Len(frmpath.txtPath.Text) - 4)
strExtn = Right$(frmpath.txtPath.Text, 4)
strTemp = strTemp + "Temp" + strExtn
If StrComp(UCase$(strExtn), ".MDB", vbTextCompare) = 0 Then
    lblPassWrd.Visible = True
    txtPWD.Visible = True
Else
    lblPassWrd.Visible = False
    txtPWD.Visible = False
End If
txtSaveAs.Text = strTemp
Unload frmpath
End Sub
Private Sub cmdSaveAs_Click()
frmpath.Show vbModal
txtSaveAs.Text = frmpath.txtPath.Text
Unload frmpath
End Sub


Private Sub cmdTabs_Click()
'Call ListTables("C:\Indx2000\AppMain\Index 2000.mdb", "PRAGMANS")
gDBTrans.CloseDB
End Sub


Private Sub cmdTest_Click()
Dim strRet As String
Dim StLen As Integer
Dim Item As Integer
Dim tempItem As Integer
Dim NewSt As String
Dim strDest As String
Dim strFinal As String
Dim strTemp As String
Dim strInterTemp As String

strRet = ReadFromIniFile("Line1", "Line1", "C:\Windows\Desktop\File1.txt")

strRet = Trim$(strRet)
StLen = Len(strRet)

For Item = 1 To StLen
    NewSt = Mid(strRet, Item, 1)
    If strInterTemp = "" Then strInterTemp = NewSt
    strDest = ReadFromIniFile("Source", NewSt, "C:\Windows\Desktop\File2.txt")
    strTemp = ReadFromIniFile("Destination", strDest, "C:\Windows\Desktop\File2.txt")
    tempItem = Item
    Do While strTemp = "" And Item <= StLen
        Item = Item + 1
        NewSt = NewSt + Mid(strRet, Item, 1)
        strDest = ReadFromIniFile("Source", NewSt, "C:\Windows\Desktop\File2.txt")
        strTemp = ReadFromIniFile("Destination", strDest, "C:\Windows\Desktop\File2.txt")
    Loop
    If strTemp <> "" Then
        strFinal = strFinal + strTemp
    Else
        Item = tempItem
        strFinal = strFinal + strInterTemp
    End If
    strInterTemp = ""
Next

End Sub


Private Sub cmdTestFile_Click()
Dim FileName As String
Dim File1Number As Integer
Dim File2Number As Integer
Dim LinesFromFile As String
Dim NextLine As String
Dim strOutPut As String
Dim i As Long
Dim J As Long

FileName = "C:\Test\Temp.txt"
File1Number = FreeFile                    ' Get unused file number
Open FileName For Binary As #File1Number   ' Create temp HTML file
'    Write #FileNumber, "<HTML> <\HTML>"  ' Output text
'Close #FileNumber                        ' Close file

FileName = "C:\Test\TestAkr1.txt"
File2Number = FreeFile                    ' Get unused file number
Open FileName For Input As #File2Number   ' Create temp HTML file

J = 1
Do Until EOF(File2Number)
   Line Input #File2Number, NextLine
   LinesFromFile = NextLine + Chr(13) + Chr(10)
'   Debug.Print LinesFromFile
    strOutPut = AkrutiToNudiStr(LinesFromFile)
    If strOutPut = "" Then
'        Debug.Print LinesFromFile
    Else
        For i = 1 To Len(strOutPut)
            Put #File1Number, J, Mid$(strOutPut, i, 1)
            J = J + 1
        Next
    End If
'    Write #File1Number, LinesFromFile  ' Output text
Loop

Close #File1Number                        ' Close file
Close #File2Number                        ' Close file

End Sub


Private Sub Form_Load()
lblConversion.Caption = "Database Conversion Version : " & App.Major & "." _
                       & App.Minor & "." & App.Revision
lblConversion.FontBold = True

lblPassWrd.Visible = True
txtPWD.Visible = True
Set gDBTrans = New clsDBUtils
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set gDBTrans = Nothing
End Sub


