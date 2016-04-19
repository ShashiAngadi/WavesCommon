Attribute VB_Name = "submain"
Option Explicit

Public gdbtrans As New clsTransact
Public m_DataBaseObject As Database
Public gAppPath As String
Public gBankId As Long
Public gBankName As String

Public Const gappName = "Transferd Data"

Public Const gDelim = "~"
Public Const wis_MESSAGE_TITLE = "Test"

'
'This function Is Used to Get the primary key of the given table
'It also Returns sqlstring Having Field names  With the retuned sql
'Just we have insert the field value with frefix values

Function GetPrimeryKey(TableName As String, retPrimeryKeyFileds() As String, retPrimeryFieldPos() As Integer, retSqlInsert As String) As Boolean
    
    Dim RstRec As Recordset
    Dim SqlBasic As String
    Dim TheTable As TableDef
    Dim theField As Field
    Dim Found As Boolean
    Dim Indx As Index
    Dim strPrimary As String
    
    ReDim retPrimeryFieldPos(0)
    ReDim retPrimeryKeyFileds(0)
    
    Dim Count As Integer
    Dim MaxCount As Integer
    Dim Pos As Long
    Dim Rst As Recordset
    Dim SqlStr As String
    'Get The Field Details & Prepare basic SQl
    Set TheTable = gdbtrans.GetDataObject.TableDefs(TableName)
        Count = 0
        MaxCount = TheTable.fields.Count - 1
        SqlBasic = ""
    For Count = 0 To MaxCount
        SqlBasic = SqlBasic & TheTable.fields(Count).Name & ","
    Next
    
    SqlBasic = Left(SqlBasic, Len(SqlBasic) - 1)
    SqlBasic = "INSERT INTO " & TableName _
          & " (" & SqlBasic & ")"
    retSqlInsert = SqlBasic
    'Now get the Primary Keys
    Count = 0
    Do
        Count = Count + 1
        If Count > TheTable.Indexes.Count Then Exit Do
        If TheTable.Indexes(Count - 1).Primary Then
            Set Indx = TheTable.Indexes(Count - 1)
            Call GetStringArray(Mid(TheTable.Indexes(Count - 1).fields, 2), retPrimeryKeyFileds, "+")
            Found = True
            Exit Do
        End If
    Loop
    Count = 0
    
    strPrimary = retPrimeryKeyFileds(Count)
    
 'GetPrimeryKey = True
End Function



'CHECK FOR THE SIGNATURE IN THIS FUNCTION
'this function is used to check the siganture
'with the existing filename and the with the given files
Private Function CheckSignature(FileName As String, RetTableName() As String) As Boolean

Dim FSIze As Long
Dim StrData As String
Dim FileData As String
Dim FileNo As Integer
Static ReadFileNo As Integer
 
CheckSignature = False

ReadFileNo = ReadFileNo + 1
 
On Error GoTo Exit_Line
'Check For the Signatures of this table...
        '1. Ho BankID
        '2. Ho Name
        '3. DataType -Master / Data
        '4. Table Name
        '5. Fieldnames with space as delimiter
    Dim HOName As String
    Dim TheTable As TableDef
    
    gdbtrans.SQLStmt = "SELECT * FROM BankDet Where BankID = " & gBankId - (gBankId) ' Mod HO_Offset)
        Call gdbtrans.SQLFetch
        
    FileNo = FreeFile
    'OPen file in Input mode
    Open FileName For Input Lock Read Write As #FileNo
    'Checking Operations Begin Here
    'Now Check the first line (i.e. "BEGIN SIGNATURE")
    StrData = "BEGIN SIGNATURE"
    Input #FileNo, FileData
    If UCase(StrData) <> UCase(FileData) Then GoTo Exit_Line
    
    'Now Check the Second Line(i,e. "DateOfCreation")
    StrData = "DateofCreation"
    Input #FileNo, FileData
    If InStr(1, FileData, StrData, vbTextCompare) = 0 Then GoTo Exit_Line
    
    'now Check for the Third Line(i,e."FileNumber")
    StrData = "FileNumber=" & ReadFileNo
    Input #FileNo, FileData
    If InStr(1, FileData, StrData, vbTextCompare) = 0 Then GoTo Exit_Line

'    'now check for the Bankid and datatype
    'Call gDbTrans.SQLFetch
    HOName = FormatField(gdbtrans.Rst("BankName"))
    StrData = "BankID=" & gBankId & gDelim & "HOBankName=" & HOName
    StrData = StrData & gDelim & "DataType=MASTER"
    Input #FileNo, FileData

    Dim MaxCount As Integer
    Dim Count As Integer
    Dim theField As Field
    Dim StrTable As String
    Dim TableCount As Integer
    
    'Now Check for the table Structure of
    'Now Get the Table Name
        Input #FileNo, FileData

    TableCount = 0
Do
    StrTable = ExtractToken(FileData, "TableName")
    If StrTable = "" Then Exit Do
    ReDim Preserve RetTableName(TableCount)
    RetTableName(TableCount) = StrTable
    Set TheTable = gdbtrans.GetDataObject.TableDefs(StrTable)
    
    StrData = "Fields of " & TheTable.Name & " are"
    Input #FileNo, FileData
    If InStr(1, StrData, FileData, vbTextCompare) = 0 Then GoTo Exit_Line
    
    
    'Now Check for the all fields in the table
    MaxCount = TheTable.fields.Count - 1
    For Count = 0 To MaxCount
         StrData = "Field" & Count + 1 & "=" & TheTable.fields(Count).Name
         Input #FileNo, FileData
         If UCase(StrData) <> UCase(FileData) Then GoTo Exit_Line
    Next Count
    
    'Now fetch the next line
    Input #FileNo, FileData
    TableCount = TableCount + 1
Loop

StrData = "END SIGNATURE"
If UCase(StrData) <> FileData Then GoTo Exit_Line

CheckSignature = True

Exit_Line:
    Close FileNo
    Err.Clear
End Function


'
'
Function ImportData() As Boolean
'For Testing(After testing should be removed)

'This functions used to Get Details From The Head office
'Such as type of Loan schemes, Loan Purpose, Caste Details
'Because braches are not allowed to Create New Loan Types
'Or to Chengethe Name of Loans
'Because Head office has to maintain unique code for
'Braches, LOanTypes.Loan Purpose and other such imp details
    
    Err.Clear
    On Error GoTo Invalid_File
    'First check the data Base in which
    'you have to write the details
    'And then find the path where to write
    Dim strDB As String
    'Select The file
    With frmExportData.cdb
        .FileName = "WISLNDAT.wislndat"
        .CancelError = False
        .DefaultExt = ".WISLNDAT"
        .Filter = "Data Files(*.wislndat)|*.wislndat|All Fiels (*.*)|*.* "
        .InitDir = App.Path
        .ShowOpen
        If .FileName = "" Then GoTo Invalid_File
        strDB = .FileName
    End With
    Dim FileName As String
     FileName = strDB
    'Now open the Data Base
    Dim MaxCount As Integer
    Dim Count As Integer
    Dim SqlStr As String
    Dim RstRec As Recordset
   'If Signatures on not match then warn as invalid files structure
   'Now open the speicified file
    Dim FileNo As Integer
    Dim StrData As String
    Dim HOName As String
    Dim TheTable As TableDef
    Dim TableCount As Integer
    'Now check for the existnce of the file & Signature of file
    'If singanture not matches then shrow error as invalid file format
    Dim TableArray() As String
    If Not CheckSignature(FileName, TableArray) Then
         MsgBox "Invalid File Structure", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    Dim FileData As String
    Dim DataType As String
   'Check for the data type
    StrData = "DataType=Master"
    DataType = ExtractToken(StrData, "DataType")
    FileNo = FreeFile
    Open FileName For Input As #FileNo
    Do
        Input #FileNo, FileData
        If InStr(1, "Begin Data", FileData, vbTextCompare) Then Exit Do
    Loop
    'Here fetch the data & put into the database(Stuck Continue)
    Dim FieldName() As String
    Dim FieldVal() As String
    Dim NoOfFiled As Integer
    Dim Pos As Integer
    'Here Get the FieldNAmes From the file
    If gdbtrans Is Nothing Then
      Set gdbtrans = New clsTransact
    End If
    'Get The Field Details & Prepare basic SQl
    Dim SqlBasic As String
    TableCount = 0
    Dim PrimFields() As String
    Dim primFieldsPos() As Integer
    Dim strPrimary  As String
    Dim Found As Boolean
    
    Set TheTable = gdbtrans.GetDataObject.TableDefs(TableArray(TableCount))
    '
    Call GetPrimeryKey(TheTable.Name, PrimFields, primFieldsPos, SqlBasic)
    strPrimary = PrimFields(0)
    
    SqlStr = "SELECT " & strPrimary & " From " & TheTable.Name
    gdbtrans.SQLStmt = SqlStr
        
    If gdbtrans.SQLFetch > 0 Then Set RstRec = gdbtrans.Rst.Clone
    Set gdbtrans.Rst = Nothing
        
    'get the primary field position
        ReDim primfieldpos(UBound(PrimFields))
    Count = 0
    MaxCount = TheTable.fields.Count - 1
    'Get the primary field position
    Do
      If Count > MaxCount Then Exit Do
          If strPrimary = TheTable.fields(Count).Name Then Exit Do
          Count = Count + 1
    Loop
    Count = Count + 1
   
   Do
        'IF CONDITION EXIT DO
FileRead:
        Input #FileNo, FileData
        'If the file indicates "End Data" Then Exit loop
        If FileData = "END DATA" Then Exit Do
        If FileData = "Next Table" Then
            TableCount = TableCount + 1
            'Now get the primary key of this table
            Call GetPrimeryKey(TableArray(TableCount), PrimFields, primFieldsPos, SqlBasic)
            GoTo FileRead
        End If
        
        Call GetStringArray(FileData, FieldVal, gDelim)
        SqlStr = ""
        MaxCount = UBound(FieldVal)
        
         'check for the fields
         For Count = 0 To MaxCount
            If FieldVal(Count) = "" Then FieldVal(Count) = "NULL"
            SqlStr = SqlStr & FieldVal(Count) & ","
         Next
                             
        'Remove the Last Comma from the string
        SqlStr = Left(SqlStr, Len(SqlStr) - 1)
        
        'Now prepare the sql statement to insert
        SqlStr = SqlBasic & " Values (" & SqlStr & ")"
        
        'Now run the query
        gdbtrans.BeginTrans
        gdbtrans.SQLStmt = SqlStr
        
        If Not gdbtrans.SQLExecute Then
            gdbtrans.RollBack
            GoTo Invalid_File
        End If
        gdbtrans.CommitTrans
 Loop
 
 MsgBox "transfered the records into the database"

ImportData = True
Invalid_File:
    If Err Then
        MsgBox "ERROR In Import Master" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
End Function



'
'
Private Sub Initialize()
'
'If gdbtrans Is Nothing Then Set gdbtrans = New clsTransact
'Dim DbPath As String
'Dim DBName As String
'
'DbPath = App.Path 'assign the path where dartabase loacted
'DBName = "Test.mdb" 'what ever it is
'If Right(DbPath, 1) = "\" Then DBName = Left(DBName, Len(DBName) - 1)
'
''Now open the data base
'If Not gdbtrans.OpenDB(DbPath & "\" & DBName, "WIS!@#") Then
'    'Do what ever you want
'    'But do not continue to run the project
'    End
'End If
'
'Initialize the global variables
    gAppPath = App.Path
    
    If gdbtrans Is Nothing Then
        Set gdbtrans = New clsTransact
    End If

'Open the data base
    If Not gdbtrans.OpenDB(gAppPath & "\loans.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gappName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gdbtrans.CreateDB(gAppPath & "\CustReg.TAB", "") Then
            MsgBox "unable to create new DataBase", vbCritical, gappName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\CustReg.MDB"
            End
        End If
    End If

End Sub




'
'open
Private Function WriteSignature(FileName As String, ParamArray TableName()) As Boolean

Dim FileNo As Integer
Dim FSIze As Long
Dim NoOFTable As Integer
Dim Count As Integer
Dim StrData As String

Static NoOfFile As Integer

On Error GoTo Exit_Line
    
    Dim Rst As Recordset
    Dim HOName As String
    Dim BegSig As String
    Dim FileSig As String
    
    FileNo = FreeFile
    'open file in ouput mode
    Open FileName For Output As #FileNo
    'Open filename For Binary Access Write Lock Read As #FileNo
    NoOfFile = NoOfFile + 1
    'Write signatue 1
    FSIze = LOF(FileNo)
    Write #FileNo, "BEGIN SIGNATURE"
    
    StrData = "DateOfCreation=" & Format(Now, "dd/mm/yy")
    Write #FileNo, StrData
    
    Write #FileNo, "FILENUMBER=" & NoOfFile
    
'    StrData = "BankID=" & gBankId & gDelim & "HOBankName=" & HOName
'    StrData = StrData & gDelim & "DataType=MASTER"
    Write #FileNo, StrData
'
'    Here you write the which table data you are writing
'    After that Fields name of the table in next line
    
    Dim TheTable As TableDef
'Now write the table details to be transferred
For Count = 0 To UBound(TableName)
    NoOFTable = NoOFTable + 1
    Set TheTable = gdbtrans.GetDataObject.TableDefs(TableName(Count))
    StrData = "TableName=" & TheTable.Name
    Write #FileNo, StrData
      
    'Extract the Field Name
    Dim MaxCount As Integer
    Dim LoopCount As Integer
    Dim theField As Field
   
    StrData = "Fields=" & TheTable.fields.Count - 1
    MaxCount = TheTable.fields.Count - 1
    
    'Write All the field Names to the file
    Write #FileNo, "Fields of " & TheTable.Name & " are"
    For LoopCount = 0 To MaxCount
         StrData = "Field" & LoopCount + 1 & "=" & TheTable.fields(LoopCount).Name
         Write #FileNo, StrData
    Next LoopCount
Next
    
    'Now Write a that signature has completed
    Write #FileNo, "END SIGNATURE"
WriteSignature = True

Exit_Line:
    Close FileNo
    Err.Clear
End Function



'
'
Function ExportData() As Boolean
    'Commneted areument
    ''''RstTransfer()  As  Recordset
    Err.Clear
    On Error GoTo ExitLine
    Dim MaxCount As Integer
    Dim Count As Integer
    Dim SqlStr As String
    Dim Rst As Recordset
    Dim FileName As String
    Dim StrData As String
    Dim strDbName As String
    Dim TableName() As String
    Dim RstCount As Recordset
    
    'Check for the existance of record


SqlStr = "SELECT * FROM BankDet"
gdbtrans.SQLStmt = SqlStr

If gdbtrans.SQLFetch < 1 Then
    MsgBox "There are no records to transafer"
    Exit Function
End If

Set Rst = gdbtrans.Rst.Clone

'get the file name to store the data
RetryLine:
    With frmExportData.cdb
        .CancelError = False
        .DefaultExt = "WISLndat"
        .Filter = "Loan Data File(*.wisLndat)|*.wisLndat|All Files (*.*)|*.*"
        .DialogTitle = "Save the Data file as"
        .FilterIndex = 1
        .ShowSave
        FileName = .FileName
    End With
    'If he has not mentoin the File name then Exit the function
    If FileName = "" Then GoTo ExitLine
    'Check for the existance of specified file
    If Dir(FileName) <> "" Then  'if exists
        Count = MsgBox("This file already exists" & vbCrLf & "Do you want to overwrite the existing file?", _
            vbInformation + vbYesNoCancel + vbDefaultButton2)
        If Count = vbCancel Then GoTo ExitLine
        If Count = vbNo Then GoTo RetryLine
        Kill FileName
    End If
    
'Now get the tableName
    'First Write the Signature
    If Not WriteSignature(FileName, "BankDet") Then Exit Function
    
    Dim FileNo As Integer
    'Now start writing the information of data
    FileNo = FreeFile
    Open FileName For Append Access Write Lock Read As #FileNo
    
    'Put the cusro at the last position of file
    Seek #FileNo, LOF(FileNo) + 1
    Write #FileNo, vbNewLine
    Write #FileNo, "BEGIN DATA"
   MaxCount = Rst.fields.Count - 1

'''Loop of record set start here

Do
    If Rst.EOF Then Exit Do
    StrData = ""
    For Count = 0 To MaxCount
        If Rst(Count).Type = dbText Then
            StrData = StrData & "'" & Rst(Count).Value & "'" & gDelim
        ElseIf Rst(Count).Type = dbDate Then
            StrData = StrData & "#" & Rst(Count).Value & "#" & gDelim
        Else
           StrData = StrData & Rst(Count).Value & gDelim
        End If
    Next Count
    'We added one extra delimeter at the end so remove that
    StrData = Left(StrData, Len(StrData) - 1)
    Write #FileNo, StrData
    Rst.MoveNext
Loop
''Check for the Next recorset avalible
'if next recordset is there then
    'Write to the file as
    'Write #FileNo, "Next Table"
    'Now Assign the next record set to rst
    'Set Rst = rstransfer(RstCount)
'''

'''loop of record set end here

    Write #FileNo, "END DATA"
    Close #FileNo
    ExportData = True
'get
ExitLine:
    If Err Then
        MsgBox "ERROR In Export master" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
End Function

Public Sub Main()
Call Initialize
frmExportData.Show
End Sub

