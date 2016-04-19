VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------
'   Module Name     :   Transaction Class:
'   Version         :   1.0.2
'   Last modified   :    09/09/2001 (dd/mm/yyyy)
'---------------------------------------------------------

Option Explicit

Const op_COMMIT = 1
Const op_ROLLBACK = 2
Const op_BEGINTRANS = 3

Private m_PrevOp As Byte
Private m_MultiTrans As Boolean

Private m_LogFileName As String
Private m_LogFileSet As Boolean
Private m_TableCount As Integer
Private m_FieldCount As Integer
Private M_FieldsCreated As Integer
Private m_DataBaseObject As Database
Private m_DBOpen As Boolean  ' Added On 10/5/2000 'To Find Whether DB is open or Closed
Private m_SQLLog As String
Private m_TransCount As Integer

'Added On 8/8/2000 to Compact & Repair the DataBase
Private M_DbPath As String
Private M_dbName As String
Dim M_DbPwd As String
''"**** /8/8/2000
Public Records As Long
Public SQLStmt As String
Public Rst As Recordset


Public QDef As QueryDef

' Structure for holding the fields info for table.
' This is used by CreateDB function.
Private Type TabStruct
    Field As String
    Type As String
    Length As Integer
    'Index As Boolean
    Required As Boolean
    'Primary As Boolean
    AutoIncrement As Boolean
End Type

' Structure for holding the fields info for Relation.
' This is used by CreateDB function.
Private Type RelnStruct
    Name As String
    Field As String
    SourceTable As String
    ForiegnTable As String
End Type

' Index structure details.
Private Type idx
    Name As String
    fields As String
    Primary As Boolean
    'Required As Boolean
    Unique As Boolean
    IgnoreNulls As Boolean
End Type

' User defined events
    'Informs about the current activity.
    Public Event UpdateStatus(strMsg As String)
    Public Event CreateDBStatus(strMsg As String, CreatedDBRatio As Single)
' Error object.
Private m_errNum As Integer
Private m_errDesc As String

Private m_SQLFile As String



Public Function CreateQueryDef(SQLStmt As String) As Boolean
Dim qdfTemp As QueryDef
Dim qdfAll As QueryDefs

Set qdfAll = m_DataBaseObject.QueryDefs

For Each qdfTemp In qdfAll
    If qdfTemp.Name = "qryTEMP" Then
        qdfAll.Delete qdfTemp.Name
        Exit For
    End If
Next

Set qdfTemp = m_DataBaseObject.CreateQueryDef("qryTEMP", SQLStmt)
CreateQueryDef = True
End Function
Public Function CheckDbStructure(strTabFile As String, strDBName As String, Optional StrPwd As String) As Boolean

CheckDbStructure = False
' Read each database name and create the db.
Dim strRet As String
Dim DBName As String
Dim Rst As Recordset

'Get The Db Name
Dim Pos As Integer
Dim PrevPos As Integer

Do
    Pos = InStr(PrevPos + 1, strDBName, "\", vbTextCompare)
    If Pos = 0 Then
        Pos = InStr(PrevPos + 1, strDBName, ".", vbTextCompare)
        If Pos <> 0 Then
            DBName = Left(Mid(strDBName, PrevPos + 1), Pos - PrevPos - 1)
        Else
            DBName = Mid(strDBName, PrevPos + 1)
        End If
        Exit Do
    End If
    PrevPos = Pos
Loop

    Dim NewTableDet() As TabStruct
   Dim TblExist As Boolean
   Dim CreateNewTable As Boolean
   Dim FldExist As Boolean
   Dim TheTable As TableDef
   Dim theField As Field
   
    ReDim NewTableDet(0)
   

    ' Check if the file path of the database
    ' is existing.  If not create it.
    On Error Resume Next
    ' Create the database.
    Dim Db As Database
    If Trim$(StrPwd) <> "" Then
        Set Db = OpenDatabase(strDBName, False, True, ";pwd=" & StrPwd)
    Else
        Set Db = OpenDatabase(strDBName, False, True)
    End If
    
    If Err.Number Then
        MsgBox "unable to find the Database", vbInformation, "WIS"
        Exit Function
    End If
    
    'Check the specified tables for this db.
    Dim J As Byte
    J = 1
    
    CreateNewTable = False
    Do
        ' Read the table name.
        Dim strTblName As String
        Dim tblData() As TabStruct
        
        TblExist = True
        strTblName = ReadFromIniFile(StripExtn(DBName), "Table" & J, strTabFile)
        If strTblName = "" Then Exit Do
        'Now Check The DataBase for table existance
        
        'Check Whethr Tale exists or Not
        Set Rst = Db.OpenRecordset("SELECT TOP 1 * From " & strTblName)
        If Err.Number <> 0 Then 'IF Table Not Existing
            Err.Clear
            TblExist = False
        End If
        Set Rst = Nothing
        
        ' Read the field details for this table into an array.
        Dim K As Byte
        K = 0
        ReDim tblData(0)
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Field" & K + 1, strTabFile)
            If strRet = "" Then Exit Do
            
            ' Add to fields array.
            ReDim Preserve tblData(K)
            With tblData(K)
                ' Set the field name.
                .Field = ExtractToken(strRet, "FieldName")
                ' Set the field type.
                .Type = FieldTypeNum(ExtractToken(strRet, "FieldType"))

                ' Set the field length.
                .Length = Val(ExtractToken(strRet, "Length"))
                ' Check, if the required flag is set.
                .Required = IIf((UCase$(ExtractToken(strRet, _
                        "Required")) = "TRUE"), True, False)

                ' Autoincrement flag.
                .AutoIncrement = IIf((UCase$(ExtractToken(strRet, _
                        "AutoIncrement")) = "TRUE"), True, False)
            End With
            
            'Now Check whether Field Exists or not
            If TblExist Then 'If table exist then only Check The Field Detail
                Set Rst = Db.OpenRecordset("SELECT  " & tblData(K).Field & " FROM " & strTblName)
                If Err.Number <> 0 Then
                    FldExist = False
                    Err.Clear
                    Set TheTable = Db.TableDefs(strTblName)
                    Set theField = TheTable.CreateField(tblData(K).Field, tblData(K).Type)
                    If tblData(K).Type = dbText Or tblData(K).Type = dbMemo Then
                        theField.Size = tblData(K).Length
                    End If
                    If tblData(K).Required Then theField.AllowZeroLength = True
                    TheTable.fields.Append theField
                Else    'CHeck The DataType
                    Set theField = TheTable.fields(tblData(K).Field)
                    If theField.Type <> tblData(K).Type Then
                        Dim NewField As Field
                        'Then Create a Temp Field Transfer the Data To THat field
                            'Set NewField = TheTable.CreateField("WISTEMPWIS", dbText, 250)
                            'TheTable.fields.Append TheField
                        'Trnasfer the Data From to new field
                        'Rst.MoveFirst
                            'While Not Rst.EOF
                                
                            'Wend
                    End If
                End If
                Set Rst = Nothing
            Else
                CreateNewTable = True
                NewTableDet(UBound(NewTableDet)) = tblData(K)
                If Err.Number Then
                    strDBName = strDBName
                End If
                ReDim Preserve NewTableDet(UBound(NewTableDet) + 1)
            End If
            'Increment the field count variable "k"
            K = K + 1
        Loop

        ' Create the table.
        If CreateNewTable Then
            CreateNewTable = False
            ReDim Preserve NewTableDet(UBound(NewTableDet) - 1)
            If Not CreateTBL(Db, strTblName, NewTableDet()) Then
                Exit Function
                'GoTo dbCreate_err
            End If
            ReDim NewTableDet(0)
        End If

        ' If any indexes are specified, create them.
        K = 0
        Dim IndxData() As idx, IndxCount As Integer
        IndxCount = 0
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Index" & K + 1, strTabFile)
            If strRet = "" Then Exit Do
            ReDim Preserve IndxData(K)
            IndxCount = K + 1
            With IndxData(K)
                .Name = ExtractToken(strRet, "IndexName")
                .fields = ExtractToken(strRet, "Fields")
                .Primary = IIf(UCase$((ExtractToken(strRet, _
                            "Primary"))) = "TRUE", True, False)
                '.Required = IIf(UCase$((extracttoken(strRet, _
                            "Required"))) = "TRUE", True, False)
                .Unique = IIf(UCase$((ExtractToken(strRet, _
                            "Unique"))) = "TRUE", True, False)
                .IgnoreNulls = IIf(UCase$((ExtractToken(strRet, _
                            "IgnoreNulls"))) = "TRUE", True, False)
            End With
            K = K + 1
        Loop
        'RaiseEvent CreateDBStatus("Creating Index For " & strTblName, m_FieldsCreated / m_FieldCount)
        If IndxCount > 0 Then
'            If Not CreateIndexes(db, strTblName, IndxData()) Then
'                GoTo dbCreate_err
'            End If
        End If
        ' Increment the table count variable "j"
        J = J + 1
    Loop
    
    ' If any Relations are specified, create them.
    Dim L As Byte
    L = 0
    Dim RelnData() As RelnStruct, RelnCount As Integer
    RelnCount = 0
    Do
        strRet = ReadFromIniFile("Relations", _
                    "Relation" & L + 1, strTabFile)
        If strRet = "" Then Exit Do
        ReDim Preserve RelnData(L)
        RelnCount = L + 1
        With RelnData(L)
            .Name = ExtractToken(strRet, "RelationName")
            .Field = ExtractToken(strRet, "Field")
            .ForiegnTable = ExtractToken(strRet, "ForiegnTable")
            .SourceTable = ExtractToken(strRet, "SourceTable")
        End With
        L = L + 1
    Loop
    Call CreateRelation(Db, RelnData)
CheckDbStructure = True
End Function
'Function CheckPermission() As Boolean
'    On Error GoTo ErrLine
'    'Temp Code
'    CheckPermission = False
'
'    '**********IMPORTANT'************
'    ' This code is temporarily written here
'    ' later we have to write a function for this Set of code
'
'    'CHECK THE PERMISSION FROM HERE
'    Dim Perms As wis_Permissions
'    Dim Operation As wis_Operations
'
'    Perms = gCurrUser.UserPermissions
'    Dim x As Integer
'    x = 1
'
'
'    If Perms = wisFullPermissions Then GoTo SuccessLine
'    'Now find the operation Type
'    SQLStmt = Trim(SQLStmt)
'    'Check the execute staement
'    If InStr(1, SQLStmt, "Update", vbTextCompare) Then Operation = wisTransaction
'    If InStr(1, SQLStmt, "Insert", vbTextCompare) Then Operation = wisTransaction
'    If InStr(1, SQLStmt, "Delete", vbTextCompare) Then Operation = wisUndoTransaction
'
'    'Decide the permissions
'    If (Perms And wisPigmyOperator) = wisPigmyOperator Then 'he is pimy operator
'        'If (Perms And wisTransaction) <> wisTransaction Then
'          If InStr(1, SQLStmt, "PDTrans", vbTextCompare) = 0 Then
'            Perms = wisNoPermissions
'          Else
'            Perms = Perms Or wisTransact  'Perms = wisTransact
'          End If
'        'End If
'    End If
'
'    'if he is creating Account then check for account creation
'    Dim Pos As Integer
'    Dim tmpPerms As wis_Permissions
'
'    Pos = InStr(1, SQLStmt, "Master", vbTextCompare)
'    If Pos Then
'        tmpPerms = wisCreateUsers
'        If Operation = wisTransaction Then
'            If (Perms And tmpPerms) <> tmpPerms Then
'                GoTo ErrLine
'            End If
'        End If
'    End If
'
'    If Operation = wisReadTransact Then GoTo ErrLine
'
'Recheck:
'    If Perms = wisFullPermissions Then GoTo SuccessLine
'
'    If Operation = wisTransaction Then
'        If (Perms And wisTransact) = wisTransact Then
'            GoTo SuccessLine
'        Else
'            GoTo ErrLine
'        End If
'    End If
'    If Operation = wisUndoTransaction Then
'        If (Perms And wisUndo) = wisUndo Then
'            GoTo SuccessLine
'        Else
'            GoTo ErrLine
'        End If
'    End If
'
'SuccessLine:
'    CheckPermission = True
'    Exit Function
'
'ErrLine:
'    If Err Then MsgBox Err.Description
''CheckPermission = True
'
'End Function
Public Sub CloseDB()
'This Function is to Chech the Transactioon Status & Close the database
On Error GoTo ErrLine
Dim nRet As Integer
If m_PrevOp = op_BEGINTRANS Then
    nRet = MsgBox("You are trying to terminate the services " _
        & "of the transaction class while inside a transaction." _
        & vbCrLf & "Do you want to commit the pending transactions?", _
        vbQuestion + vbYesNo)
    If nRet = vbNo Then
        Workspaces(0).RollBack
    ElseIf nRet = vbYes Then
        Workspaces(0).CommitTrans
    End If
End If

m_DataBaseObject.Close
m_DBOpen = False

Exit Sub

ErrLine:
If Err.Number = 3420 Then
    Exit Sub
Else
    MsgBox Err.Number & " :" & vbCrLf & "Following error occured while closing DatatBase" & _
                vbCrLf & Err.Description, vbCritical, "Db Error"
End If

End Sub

Public Function CommitTrans() As Boolean
Dim FIleNo As Integer
Dim FSIze As Long
On Error GoTo ErrLine
CommitTrans = False
    FIleNo = FreeFile
    If m_PrevOp = op_BEGINTRANS Then
        'First open the file
        If m_LogFileName <> "" Then 'YOu should have a log
            Open m_LogFileName For Binary As #FIleNo
            FSIze = LOF(FIleNo)
            Seek #FIleNo, FSIze + 1
            Put #FIleNo, , m_SQLLog
            Close #FIleNo
            m_SQLLog = ""
        End If
        m_PrevOp = op_COMMIT
        Workspaces(0).CommitTrans
        m_TransCount = m_TransCount - 1
        CommitTrans = True
    ElseIf MultiTrans Then
        Debug.Assert m_TransCount = 10
        Workspaces(m_TransCount - 1).CommitTrans
        m_TransCount = m_TransCount - 1
        CommitTrans = True
    End If
    
Exit Function
ErrLine:
Close #FIleNo
CommitTrans = False
End Function
Public Function CompactDataBase() As Boolean
Dim strDupDb As String
Dim Prp As Property

If m_DBOpen Then
    'Set Prp = m_DataBaseObject
    Me.CloseDB
    m_DBOpen = False
End If
'now get a name for Duplicate Of Database
Dim Count As Integer
Do
    Count = InStr(Count + 1, M_dbName, ".", vbTextCompare)
    If Count = 0 Then Exit Do
    strDupDb = Left(M_dbName, Count - 1)
    
Loop
'For Destination Databace put Db name as same with suffix "dup"
strDupDb = strDupDb & "dup.MDB"

If Dir(strDupDb) <> "" Then Kill strDupDb

Call DBEngine.CompactDataBase(M_dbName, strDupDb, , dbLangGeneral & ";pwd=" & M_DbPwd)
Call DBEngine.RepairDatabase(M_dbName)
Call Me.OpenDB(M_dbName, M_DbPwd)

End Function
' Creates a Database file (.mdb) by taking the necessary data
' from a given INI file, with the given password.
'
' Returns:
'   True    -   On successful creation of the database.
'   False   -   On failure.
'
' Created by    :   Ravindranath M.
' Last modified :   05/08/1999 (dd/mm/yyyy)
'                   08/12/1999
'
Public Function CreateDB(ByVal strdataFile As String, Optional StrPwd As String) As Boolean
On Error GoTo dbCreate_err
RaiseEvent CreateDBStatus("Reading Data File " & strdataFile, 0)
' Check if the specified datafile is existing.
If Dir(strdataFile, vbNormal) = "" Then
    'MsgBox "File not found - " & strDataFile, vbCritical
    Err.Raise vbObjectError - wis_FILENOTFOUND, , _
            ErrMsg(wis_FILENOTFOUND, strdataFile)
    GoTo dbCreate_err
End If

'Get the total No Of fields & Tables to be Craete
m_TableCount = GetTableCount(strdataFile)
m_FieldCount = GetFieldCount(strdataFile)

' Read each database name and create the db.
Dim i As Byte, strRet As String
Dim DBName As String, dbFile As String
Dim DBPath As String
i = 1
Do
    ' Read the dbname from datafile.
    strRet = ReadFromIniFile("Databases", "DataBase" & i, strdataFile)
    If strRet = "" Then Exit Do

    ' Get the name of the database file.
    DBName = ExtractToken(strRet, "dbName")
    If DBName = "" Then Exit Do
    'RaiseEvent CreateDBStatus("Creating DataBase " & DBName, m_FieldsCreated / m_FieldCount)
    ' Get the file path.
    DBPath = ExtractToken(strRet, "dbPath")
    dbFile = DBPath & "\" & DBName
    ' If no path mentioned, set the path as the path of data file.
    If DBPath = "" Then
        dbFile = AppendBackSlash(FilePath(strdataFile)) & DBName
        DBPath = FilePath(dbFile)
    End If

    ' Check if the file path of the database
    ' is existing.  If not create it.
    If Dir(DBPath, vbDirectory) = "" Then
        If Not MakeDirectories(DBPath) Then
            MsgBox "Error in creating the path " & DBPath _
                & " for " & DBName, vbCritical
            'MsgBox LoadResString(gLangOffSet + 809) & DBPath, vbCritical
            GoTo dbCreate_err
        End If

    Else
        ' Check if the file is already existing.
        ' If existing, get the user action.
        dbFile = StripExtn(dbFile) & ".mdb"
        If Dir(dbFile, vbNormal) <> "" Then
        Dim nRet As Integer
            nRet = MsgBox("WARNING : " & vbCrLf & vbCrLf & "The database file '" _
                    & dbFile & "' is already existing.  If you choose to overwrite " _
                    & "this file, you will loose the existing data permanantly." _
                    & vbCrLf & vbCrLf & "Do you want overwrite this file?", _
                    vbYesNo + vbCritical + vbDefaultButton2)
            If nRet = vbYes Then
                ' Delete the existing file.
                Kill dbFile
            ElseIf nRet = vbNo Then
                GoTo dbCreate_err
            End If
        End If
    End If

    ' Create the database.
    Dim Db As Database
    Dim strLocale As String
    If Trim$(StrPwd) = "" Then
        Set Db = CreateDatabase(dbFile, dbLangGeneral)
    Else
        Set Db = CreateDatabase(dbFile, dbLangGeneral & ";pwd=" & StrPwd)
    End If

    'Create the specified tables for this db.
    Dim J As Byte
    J = 1
    Do
        ' Read the table name.
        Dim strTblName As String
        Dim tblData() As TabStruct
        strTblName = ReadFromIniFile(StripExtn(DBName), "Table" & J, strdataFile)
        If strTblName = "" Then Exit Do

        ' Read the field details for this table into an array.
        Dim K As Byte
        K = 0
        ReDim tblData(0)
        'RaiseEvent CreateDBStatus("Reading table Structure of " & DBName, m_FieldsCreated / m_FieldCount)
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Field" & K + 1, strdataFile)
            If strRet = "" Then Exit Do

            ' Add to fields array.
            ReDim Preserve tblData(K)
            With tblData(K)
                ' Set the field name.
                .Field = ExtractToken(strRet, "FieldName")
                ' Set the field type.
                .Type = FieldTypeNum(ExtractToken(strRet, "FieldType"))

                ' Set the field length.
                .Length = Val(ExtractToken(strRet, "Length"))
                ' Check, if the required flag is set.
                .Required = IIf((UCase$(ExtractToken(strRet, _
                        "Required")) = "TRUE"), True, False)

                ' Autoincrement flag.
                .AutoIncrement = IIf((UCase$(ExtractToken(strRet, _
                        "AutoIncrement")) = "TRUE"), True, False)
            End With

            ' Increment the field count variable "k"
            K = K + 1
        Loop

        ' Create the table.
        'RaiseEvent CreateDBStatus("Creating Table " & strTblName & " For " & DBName, m_FieldsCreated / m_FieldCount)
        If Not CreateTBL(Db, strTblName, tblData()) Then
            GoTo dbCreate_err
        End If

        ' If any indexes are specified, create them.
        K = 0
        Dim IndxData() As idx, IndxCount As Integer
        IndxCount = 0
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Index" & K + 1, strdataFile)
            If strRet = "" Then Exit Do
            ReDim Preserve IndxData(K)
            IndxCount = K + 1
            With IndxData(K)
                .Name = ExtractToken(strRet, "IndexName")
                .fields = ExtractToken(strRet, "Fields")
                .Primary = IIf(UCase$((ExtractToken(strRet, _
                            "Primary"))) = "TRUE", True, False)
                '.Required = IIf(UCase$((extracttoken(strRet, _
                            "Required"))) = "TRUE", True, False)
                .Unique = IIf(UCase$((ExtractToken(strRet, _
                            "Unique"))) = "TRUE", True, False)
                .IgnoreNulls = IIf(UCase$((ExtractToken(strRet, _
                            "IgnoreNulls"))) = "TRUE", True, False)
            End With
            K = K + 1
        Loop
        'RaiseEvent CreateDBStatus("Creating Index For " & strTblName, m_FieldsCreated / m_FieldCount)
        If IndxCount > 0 Then
            If Not CreateIndexes(Db, strTblName, IndxData()) Then
                GoTo dbCreate_err
            End If
        End If

        ' Increment the table count variable "j"
        J = J + 1
    Loop
    
    ' If any Relations are specified, create them.
    Dim L As Byte
    L = 0
    Dim RelnData() As RelnStruct, RelnCount As Integer
    RelnCount = 0
    Do
        strRet = ReadFromIniFile("Relations", _
                    "Relation" & L + 1, strdataFile)
        If strRet = "" Then Exit Do
        ReDim Preserve RelnData(L)
        RelnCount = L + 1
        With RelnData(L)
            .Name = ExtractToken(strRet, "RelationName")
            .Field = ExtractToken(strRet, "Field")
            .ForiegnTable = ExtractToken(strRet, "ForiegnTable")
            .SourceTable = ExtractToken(strRet, "SourceTable")
        End With
        L = L + 1
    Loop
    Call CreateRelation(Db, RelnData)
        
    
    ' Increment the DB count variable "i"
    i = i + 1
Loop

' Set the return value.
CreateDB = True

Exit_Line:
    Exit Function

dbCreate_err:
    If Err.Number = 75 Then ' Path/File access error.
       nRet = MsgBox("Error accessing the file '" _
                & strRet & "'.", vbRetryCancel + vbCritical)
        If nRet = vbRetry Then Resume
    
    ElseIf Err Then
       
        MsgBox Err.Description, vbCritical
    End If
'    Resume
End Function


Private Function CreateRelation(Db As Database, Reln() As RelnStruct) As Boolean
On Error GoTo CreateRalation_Error

' Open the specified table.
Dim TheTable As TableDef
Dim FrnTable As TableDef
Dim theField As Field
Dim TheRelation As Relation
Dim fldArr() As String
Dim i As Integer, J As Integer

'Set TheTable = Db.TableDefs(SourceTable)
'Set FrnTable = Db.TableDefs(ForiegnTable)
'If Reln() Is Nothing Then GoTo ExitLine
For i = 0 To UBound(Reln)
    With Db
        'Create Relation object,
        Set TheTable = Db.TableDefs(Reln(i).SourceTable)
        Set FrnTable = Db.TableDefs(Reln(i).ForiegnTable)

        Set TheRelation = .CreateRelation(Reln(i).Name, Reln(i).SourceTable, Reln(i).ForiegnTable, dbRelationUpdateCascade)
        
        ' the names of the two tables in the relation.
        'Set relNew = .CreateRelation("EmployeesDepartments", _
            tdfNew.Name, tdfEmployees.Name, _
            )

        ' Create Field object for the Fields collection of the
        ' new Relation object. Set the Name and ForeignName
        ' properties based on the fields to be used for the
        ' relation.
        
        
        With TheRelation
            Set theField = .CreateField(Reln(i).Field)
            '.fields.Append .CreateField(Reln(i).Field)
        End With
        TheRelation.fields.Append theField
    End With
Next

ExitLine:
CreateRelation = True

Exit Function

CreateRalation_Error:
    If Err.Number = 9 Then
        GoTo ExitLine
    ElseIf Err.Number = 3283 Then   ' Primary key already exists.
        Resume Next
    ElseIf Err Then
        'MsgBox Err.Description, vbCritical
        MsgBox Err.Description, vbCritical
    End If

End Function

Public Function GetDataObject() As Database
    
    Set GetDataObject = m_DataBaseObject
End Function

' Retrieves the value for a specified token
' in a given source string.
' The source should be of type :
'       name1=value1;name2=value2;...;name(n)=value(n)
'   similar to DSN strings maintained by ODBC manager.
'
Private Function ExtractToken(src As String, TokenName As String) As String
' If the src is empty, exit.
If Len(src) = 0 Or _
    Len(TokenName) = 0 Then Exit Function

' Search for the token name.
Dim token_pos As Integer
Dim strSearch As String
strSearch = TokenName & "="

'token_pos = InStr(src, strSearch)
'If token_pos = 0 Then
'    'Try ignoring the white space
'    strSearch = token_name & " ="
'    token_pos = InStr(src, strSearch)
'    If token_pos = 0 Then Exit Function
'End If

' Search for the token_name in the src string.
token_pos = InStr(1, src, strSearch, vbTextCompare)
Do
    ' The character before the token_name
    ' should be ";" or, it should be the first word.
    ' Else, search for the next occurance of the token.
    If token_pos = 0 Then
        If token_pos = 0 Then
            'Try ignoring the white space
            strSearch = TokenName & " ="
            token_pos = InStr(src, strSearch)
            If token_pos = 0 Then Exit Function
        End If
    ElseIf token_pos = 1 Then
        Exit Do
    ElseIf Mid$(src, token_pos - 1, 1) = ";" Then
        Exit Do
    Else
        'Get next occurance.
        token_pos = InStr(token_pos + 1, src, TokenName, vbTextCompare)
    End If
Loop

token_pos = token_pos + Len(strSearch)

' Search for the delimiter ";", after the token_pos.
Dim Delim_pos As Integer
Delim_pos = InStr(token_pos, src, ";")
If Delim_pos = 0 Then Delim_pos = Len(src) + 1

' Return the token_value.
ExtractToken = Mid$(src, token_pos, Delim_pos - token_pos)
End Function

Private Function CreateIndexes(Db As Database, Tbl As String, Indx() As idx) As Boolean
On Error GoTo CreateIndex_Error
' Open the specified table.
Dim TheTable As TableDef
Dim TheIndex As Index
Dim fldArr() As String
Dim i As Integer, J As Integer

Set TheTable = Db.TableDefs(Tbl)

For i = 0 To UBound(Indx)
    With TheTable
        Set TheIndex = .CreateIndex(Indx(i).Name)
        ' Breakup the fields string to get the field names for index.
        GetStringArray Indx(i).fields, fldArr(), "+"
        With TheIndex
            For J = 0 To UBound(fldArr)
                .fields.Append .CreateField(fldArr(J))
            Next
            .Primary = Indx(i).Primary
            '.Required = Indx(i).Required
        End With
        .Indexes.Append TheIndex
    End With
Next
CreateIndexes = True

CreateIndex_Error:
    If Err.Number = 3283 Then   ' Primary key already exists.
        Resume Next
    ElseIf Err Then
        'MsgBox Err.Description, vbCritical
        MsgBox Err.Description, vbCritical
    End If
End Function

Private Function CreateTBL(Db As Database, TblName As String, tblData() As TabStruct) As Boolean
'
'   This routine creates the specified table whose fields are
'   defined in the array tblData().  tblData is a structure type
'   which specifies the following field information.
'       1. Field Name
'       2. Field Type
'       3. Field Length
'       4. Index (True/False)
'       5. Primary (True/False)

Dim TheTable As TableDef
Dim TheFields() As Field, TheIndex As Index
Dim i As Integer
Dim PrimaryKeySet As Boolean
'On Error GoTo err_line

' Create the table.
'RaiseEvent UpdateStatus("Creating the table " & TblName & "...")
Set TheTable = Db.CreateTableDef(TblName)

' Create and add the fields.
Dim theField As Field
With TheTable
    For i = 0 To UBound(tblData)
        ' Create and append the field.
        Set theField = .CreateField(tblData(i).Field, _
                        tblData(i).Type, tblData(i).Length)
        If tblData(i).Type = dbText Or tblData(i).Type = dbMemo Then
            theField.AllowZeroLength = IIf(tblData(i).Required, False, True)
        End If
        
        theField.Required = tblData(i).Required
        If tblData(i).AutoIncrement Then theField.Attributes = dbAutoIncrField
        .fields.Append theField

    '// Indexes are now created separately.  They are not
    '// part of field specification.  To create indexes,
    '// an index key entry under the table section must exist.
            '' Create and append index, if specified.
            'If tblData(i).Index Then
             '   Set TheIndex = TheTable.CreateIndex("idx" & tblData(i).Field)
             '   TheIndex.fields.Append TheIndex.CreateField(tblData(i).Field)
             '   ' Check if primary option is specified.
             '   ' Do this only if a primary key is not already set.
             '   ' Avoid multiple Primary fields specification.
             '   If Not PrimaryKeySet And tblData(i).Primary Then
             '       TheIndex.Primary = True
             '   End If
             '   TheTable.Indexes.Append TheIndex
            'End If
    
 ' Update the Variable Fieldscreated by ine
    M_FieldsCreated = M_FieldsCreated + 1
    
    Next
End With

Db.TableDefs.Append TheTable
Set TheTable = Nothing
CreateTBL = True
Exit Function

Err_Line:
    If Err.Number = 3010 Then 'table already exists
        Exit Function
    ElseIf Err.Number > 0 Then
         'MsgBox Err.Description, vbCritical
        MsgBox Err.Description
        CreateTBL = False
    'Resume
    End If

End Function
' Returns the vb field type number, given a field type string.
Private Function FieldTypeNum(strFldType As String) As Integer

Select Case UCase$(strFldType)
    Case "TEXT", "DBTEXT"
        FieldTypeNum = dbText
    Case "NUMBER", "INTEGER", "DBNUMBER", "DBINTEGER"
        FieldTypeNum = dbInteger
    Case "DATE", "DBDATE"
        FieldTypeNum = dbDate
    Case "BOOLEAN", "DBBOOLEAN"
        FieldTypeNum = dbBoolean
    Case "BYTE", "DBBYTE"
        FieldTypeNum = dbByte
    Case "CHAR", "DBCHAR"
        FieldTypeNum = dbChar
    Case "CURRENCY", "DBCURRENCY"
        FieldTypeNum = dbCurrency
    Case "DECIMAL", "DBDECIMAL"
        FieldTypeNum = dbDecimal
    Case "DOUBLE", "DBDOUBLE"
        FieldTypeNum = dbDouble
    Case "FLOAT", "DBFLOAT"
        FieldTypeNum = dbFloat
    Case "LONG", "DBLONG"
        FieldTypeNum = dbLong
    Case "LONGBINARY", "DBLONGBINARY"
        FieldTypeNum = dbLongBinary
    Case "MEMO", "DBMEMO"
        FieldTypeNum = dbMemo
    Case "SINGLE", "DBSINGLE"
        FieldTypeNum = dbSingle
    Case Else
        FieldTypeNum = 0
End Select

End Function
Private Function FieldTypeStr(nFldType As Integer) As String
' Given a fieldtype number,
' this function returns the string litteral.

Select Case nFldType
    Case dbText
        FieldTypeStr = "Text"
    Case dbInteger
        FieldTypeStr = "Integer"
    Case dbDate
        FieldTypeStr = "Date"
    Case dbBoolean
        FieldTypeStr = "Boolean"
    Case dbByte
        FieldTypeStr = "Byte"
    Case dbChar
        FieldTypeStr = "Char"
    Case dbCurrency
        FieldTypeStr = "Currency"
    Case dbDecimal
        FieldTypeStr = "Decimal"
    Case dbDouble
        FieldTypeStr = "Double"
    Case dbFloat
        FieldTypeStr = "Float"
    Case dbLong
        FieldTypeStr = "Long"
    Case dbLongBinary
        FieldTypeStr = "LongBinary"
    Case dbMemo
        FieldTypeStr = "Memo"
    Case dbSingle
        FieldTypeStr = "Single"
    Case Else   ' Unrecognized fldtype.
        FieldTypeStr = ""
End Select

End Function

Private Function GetFieldCount(strdataFile As String, Optional TblName As String) As Integer
Dim strRet As String
Dim FieldCount As Integer
Dim TableName As String
Dim DBName As String
Dim i As Integer, J As Integer, K As Integer
Dim FieldName As String
Do
    i = i + 1
    DBName = ReadFromIniFile("DataBases", "DataBase" & i, strdataFile)
    If DBName = "" Then Exit Do
    J = 0
    Do
        J = J + 1
        TableName = ReadFromIniFile(DBName, "Table" & J, strdataFile)
        If TableName = "" Then Exit Do
        K = 0
        Do
            K = K + 1
            FieldName = ReadFromIniFile(TableName, "Field" & K, strdataFile)
             If FieldName = "" Then Exit Do
            FieldCount = FieldCount + 1
        Loop
        If Trim(TblName) <> "" Then
            If TableName = TblName Then GoTo ExitLine
            FieldCount = 0
        End If
            
    Loop
Loop
ExitLine:
GetFieldCount = FieldCount
End Function

Private Function GetTableCount(strdataFile As String) As Integer
Dim strRet As String
Dim TableCount As Integer
Dim DBName As String
Dim i As Integer, J As Integer, K As Integer
    Do
        i = i + 1
        DBName = ReadFromIniFile("DataBases", "Database" & i, strdataFile)
        If DBName = "" Then Exit Do
        J = 0
        Do
            J = J + 1
            strRet = ReadFromIniFile(DBName, "Table" & J, strdataFile)
            If strRet = "" Then Exit Do
            TableCount = TableCount + 1
        Loop
    Loop
GetTableCount = TableCount
End Function



'
'
Public Function BackUPDB(DBPath As String) As Boolean
'CAUTION: The data base that you pass to it should be closed
'Steps for Backing UP
'1. Repair the data base
'2. Compact the Data Base
'3. Split the data base ' Do this later with the SPLIT DLL
'Dim Count As Integer
Dim ch As String * 1
Dim Pos As Integer
Dim DBDir As String
    'This method first repairs the data base and then compacts the database
    If DBPath = "" Then
        Exit Function
    End If

    #If GD Then
        'Get the directory of the Data base
        For Count = 0 To Len(DBPath)   'C:\kdjfkljklsdja\dasjfdljflsd\kfdsjfdlasf.mdb
            ch = Right(Left(DBPath, Len(DBPath) - Count), 1)
            MsgBox ch
            If ch = "\" Then
                Exit For
            End If
        Next Count
        DBDir = Left(DBPath, Len(DBPath) - Count)
        DBDir = DBDir & "TMP.MDB"
    #End If

    ' Modified by Ravindra...
    DBDir = AppendBackSlash(FilePath(DBPath)) & "Tmp.mdb"

    'On Error GoTo Errline
    FileCopy DBPath, DBDir

   
    Kill DBDir
        Exit Function
ErrLine:
    Exit Function

End Function
Public Function OpenDB(strDBName As String, pwd As String) As Boolean

On Error GoTo ErrLine
If strDBName = "" Then
    Exit Function
End If
Dim Count As Integer
  
'  Debug.Assert strDbName = ""

Set m_DataBaseObject = OpenDatabase(strDBName, False, False, ";pwd=" & pwd)
OpenDB = True
m_DBOpen = True
M_dbName = strDBName
M_DbPwd = pwd
Do
    Count = InStr(Count + 1, strDBName, "\", vbBinaryCompare)
    If Count = 0 Then Exit Do
    M_DbPath = Left(strDBName, Count)
Loop

Exit Function
ErrLine:
    If Err.Number = 3031 Then
        Debug.Print "NOT A VALID PASSWORD"
        pwd = "PRAGMANS"
        Resume
    Else
        MsgBox "OpenDB: The following error occured, " _
                & "while opening the database." _
                & vbCrLf & Err.Description, vbCritical
    End If
    Exit Function

End Function

Public Function QueryDef(ByVal QueryName As String) As Boolean

Err.Clear
On Error GoTo ErrLine
QueryDef = False

If QueryName = "" Then Exit Function
Set QDef = m_DataBaseObject.QueryDefs(QueryName)

QueryDef = True
Exit Function

ErrLine:

If Err.Number = 3265 Then
    MsgBox "There is no such query in database", vbCritical, wis_MESSAGE_TITLE
    Exit Function
End If
If Err Then
    MsgBox "QueryDef : " & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Exit Function
End If
End Function

Public Function RollBack() As Boolean
    
    RollBack = False
    If m_PrevOp = op_BEGINTRANS Then
        m_PrevOp = op_ROLLBACK
        m_SQLLog = ""
        Workspaces(0).RollBack
        m_TransCount = m_TransCount - 1
        RollBack = True
    ElseIf MultiTrans Then
        Debug.Assert m_TransCount = 10
        Workspaces(m_TransCount - 1).RollBack
        m_TransCount = m_TransCount - 1
        RollBack = True
    End If

End Function

Public Function BeginTrans() As Boolean
    BeginTrans = False
    If m_PrevOp <> op_BEGINTRANS Then
        Workspaces(0).BeginTrans
        m_TransCount = m_TransCount + 1
        m_PrevOp = op_BEGINTRANS
        BeginTrans = True
    ElseIf m_PrevOp = op_BEGINTRANS And MultiTrans Then
        m_TransCount = m_TransCount + 1
        Workspaces(m_TransCount).BeginTrans
    Else
        Debug.Assert m_TransCount = 10
    End If

End Function
Public Property Get LogFileName() As String
    LogFileName = m_LogFileName
End Property

Public Property Let LogFileName(NewValue As String)
    Dim Retval As Integer
    Dim FIleNo As Integer
    Retval = DoesPathExist(NewValue)
    If Retval = 2 Or Retval = -1 Or Retval = -2 Then
        Err.Raise 10000, "DBTransact", "Invalid Log File Name specified !"
        Exit Property
    End If
    If Retval = 0 Then
        'FIle does  not exist open the file and put the header info to it
        On Error GoTo ErrLine
        FIleNo = FreeFile
        Open NewValue For Binary Access Write As #FIleNo
        Put #FIleNo, , "INDEX-2000 Log file"
        Close #FIleNo
    End If

    m_LogFileName = NewValue
    m_LogFileSet = True     ' Set the LOGFILE flag to True.
    Exit Property
ErrLine:
    Close #FIleNo
    Err.Raise Err.Number, Err.Source, Err.Description

End Property



Public Function SQLExecute() As Boolean
On Error GoTo ErrLine

' Raise the UpdateStatus event.
    RaiseEvent UpdateStatus("Executing the SQL statement...")

' Clear the error settings...
    m_errDesc = "": m_errNum = 0

' Check if the private variable m_sqlstmt is set.
    If Trim$(SQLStmt) = "" Then
        MsgBox "No SQL statement is set!", vbExclamation
        GoTo Exit_Line
    End If

'Check if the transaction has started
     If m_PrevOp <> op_BEGINTRANS Then
        'MsgBox "Transaction not started !", vbExclamation, gAppName & " - Critical Error"
        MsgBox "Transaction not started !", vbExclamation, gAppName & " - Critical Error"
        GoTo Exit_Line
    End If

'If Not CheckPermission Then GoTo AccessDeny

ExeLine:
'Execute the Action Query.
m_DataBaseObject.Execute SQLStmt$, _
        dbFailOnError Or dbSeeChanges
'm_SQLLog = m_SQLLog & vbCrLf & SQLStmt
 m_SQLLog = SQLStmt
Dim FIleNo As Integer
'FIleNo = FreeFile(1)
'''get the file of the today date

'Open m_SQLFile For Append Access Write Lock Write As FIleNo
'Write #FIleNo, SQLStmt & vbCrLf
'Close FIleNo
SQLExecute = True

Exit Function

AccessDeny:
    MsgBox "You are not permitted todo this Operation", vbInformation, wis_MESSAGE_TITLE
    'MsgBox LoadResString(gLangOffSet + 810), vbInformation, wis_MESSAGE_TITLE
    
Exit_Line:
    'Reset sqlstmt to avoid misuse
    SQLStmt = ""
    RaiseEvent UpdateStatus("")
    
    Exit Function

ErrLine:
    #If COMMENTED_BY_RAVINDRA Then
    If Err.Number = 3022 Then       'Duplicate record.
        GoTo Exit_Line
    ElseIf Err.Number = 3464 Then     ' ??? Why handle this error in this way ???
        Err.Clear
        On Error GoTo Exit_Line
        m_DataBaseObject.Execute SQLStmt
        SQLExecute = True
        'Do this only on success
        m_SQLLog = m_SQLLog & vbCrLf & SQLStmt
    ElseIf Err Then
        MsgBox "SQLExecute: The following error occurred." _
            & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    End If
    #End If
    
    ' Handle the errors here...
    If Err Then
        If Err.Number = 3375 Then
            SQLExecute = True
            Exit Function
        ElseIf Err.Number <> 3380 Then
            MsgBox "SQLExecute: The following error occurred." _
                & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
            m_errNum = Err.Number
            m_errDesc = Err.Description
        End If
    End If
End Function

Public Function SQLFetch() As Long

'Debug.Assert SqlStmt <> ""

' Clear the error settings...
m_errDesc = "": m_errNum = 0

' Raise event for status updation.
RaiseEvent UpdateStatus("Querying records...")

'If no SQL statement set, exit.
If SQLStmt = "" Then
    MsgBox "No SQL statement has been set!", vbExclamation
    GoTo Exit_Line
End If

' If the datbase object is not initialized, exit.
If m_DataBaseObject Is Nothing Then
    MsgBox "SQLFetch: No database is in use.  " _
        & "Set the database name before calling SQLFetch.", vbExclamation
    GoTo Exit_Line
End If

Records = 0
On Error GoTo ErrLine
SQLFetch = 0
Set Rst = m_DataBaseObject.OpenRecordset(SQLStmt, _
            dbOpenDynaset, dbSeeChanges, dbOptimistic)
                If Rst.RecordCount > 0 Then
    
    Rst.MoveLast
    Rst.MoveFirst
    Records = Rst.RecordCount
    SQLFetch = Records
    'Set the sqlstmt to null
    SQLStmt = ""
End If

Exit_Line:
    RaiseEvent UpdateStatus("")
    Exit Function

ErrLine:
    If Err Then
        SQLFetch = -1   ' Cannot we have some other number (0) for error return?
                    ' Because -1 in VB signifies True.
        MsgBox "SQLFetch: The following error occurred." _
                & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
        m_errNum = Err.Number
        m_errDesc = Err.Description
        Err.Clear
    End If
'Resume
End Function



Private Sub Class_Initialize()
'During the initialize get the file name
'with the path
'These files should be stored where Db back upo is keeping
Dim FolderName As String
Dim PresentDate As String

PresentDate = Format(Now, "DD/YY/MM")
'FolderName = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "BaseFolder")

'if date = 22/2/01 then
'Filename becomes wissql2201

If Trim(FolderName) = "" Then FolderName = App.Path
If Right(FolderName, 1) = "\" Then FolderName = Left(FolderName, Len(FolderName) - 1)

m_SQLFile = FolderName + "\" + "wSQL" + Left(PresentDate, 2) + Right(PresentDate, 2) & ".txt"

End Sub
Private Sub Class_Terminate()
' The Below Check Has to done in Close Db Method
Dim nRet As Integer
If m_DBOpen Then
    'Set Prp = m_DataBaseObject
    Me.CloseDB
    m_DBOpen = False
End If
If m_PrevOp = op_BEGINTRANS Then
    nRet = MsgBox("You are trying to terminate the services " _
        & "of the transaction class while inside a transaction." _
        & vbCrLf & "Do you want to commit the pending transactions?", _
        vbQuestion + vbYesNo)
    If nRet = vbNo Then
        Workspaces(0).RollBack
    ElseIf nRet = vbYes Then
        Workspaces(0).CommitTrans
    End If
End If
'Added on 10/5/2000 ' To Check The status of DataBase
If m_DBOpen Then
    
    Err.Raise 2000 + wis_FAILURE, "Transaction Class", "You are trying to terminate the services " _
        & "of the transaction class while the Data Base is open ."
 End If
    
End Sub


Public Property Get errNum() As Variant
    errNum = m_errNum
End Property


Public Property Get errDesc() As Variant
    errDesc = m_errDesc
End Property

Public Property Let MultiTrans(NewBool As Boolean)
    If m_TransCount = 0 Then
        m_MultiTrans = NewBool
    Else
        Err.Raise 10002, "DataBase", "Transaction is active"
    End If
End Property

Public Property Get MultiTrans() As Boolean
     MultiTrans = m_MultiTrans
End Property

