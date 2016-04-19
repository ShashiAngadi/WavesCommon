VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'********************************************************************
' Module Name   :   Transaction Class
' Version       :   1.0.0
' Created On    :   06/06/2002 (dd/mm/yyyy)
' Last Modified :   17/07/2002 (dd/mm/yyyy)
' Remarks       :   This global class gives the services of opening
'                   the database, fetching data, executing queries and etc.
'********************************************************************

Option Explicit

Const op_COMMIT = 1
Const op_ROLLBACK = 2
Const op_BEGINTRANS = 3

Private m_PrevOp As Byte
Private m_DBOpen As Boolean

'Error object
'Private m_errNum As Integer
'Private m_errDesc As String

'Filename to store sql strings processed on a particular date
'Private m_SQLFile As String

'Public Rst As ADODB.Recordset
'Public CommandObject As ADODB.Command

'Private m_TransCount As Integer

Private m_DataBaseObject As ADODB.Connection
Private m_Catalog As ADOX.Catalog
Public m_SQLStmt As String

'\\\\\\\\\
' This function searches the ADO recordset for the given value.
'
' Paramaeters are
'  1.  rstToFind  : adodb.recordset
'  2.  SearchString   : search string   ex : "BankID=" & bankid
'
' searches the record if found returns true else false
'
' - Pradeep , 31/01/2002
'\\\\\\\\\\
Public Function FindRecord(ByRef rstToFind As ADODB.Recordset, ByVal SearchString As String) As Boolean

' in the given string
On Error GoTo Hell

Dim SearchArray() As String
Dim ItemCount As Integer

' initialise the function to false
FindRecord = False

' if any filter is there then remove the filter

rstToFind.Filter = adFilterNone

Call GetStringArray(SearchString, SearchArray, ",")

'Comment on filtereation
For ItemCount = LBound(SearchArray) To UBound(SearchArray) - 1
    rstToFind.Filter = SearchArray(ItemCount)
Next ItemCount

SearchString = SearchArray(ItemCount)
    
'Move the recordset to first record
'Although we can start searching record from the current record
'but still right now we are moving to the first record.
If rstToFind.BOF And rstToFind.EOF Then
    FindRecord = False
    Exit Function
End If

rstToFind.MoveFirst
' seek the searchstring

rstToFind.Find SearchString

' if found it will be in the found record
If Not rstToFind.EOF Then FindRecord = True

'Else
'    FindRecord = False ' else it will on the EOF
    'move the cursor where it was
'End If

Exit Function

Hell:
    MsgBox "Field Specified or Value given is wrong !!", vbCritical, "Find Record"
    
End Function
'
Public Function BeginTrans() As Boolean
On Error GoTo ErrLine

If m_PrevOp <> op_BEGINTRANS Then
    m_PrevOp = op_BEGINTRANS
    m_DataBaseObject.BeginTrans
End If

BeginTrans = True

ErrLine:
    If Err Then _
        MsgBox "BeginTrans : " & Err.Number & " - " & Err.Description, vbCritical
    
End Function

'
'
'Thissubroutine will check for any pending Transactions
'if so it will prompt you to commit or rollback
Public Sub CloseDB()
On Error GoTo ErrLine

Dim nRet As Integer

If m_PrevOp = op_BEGINTRANS Then
    nRet = MsgBox("You are trying to terminate the services " _
        & "of the transaction class while inside a transaction." _
        & vbCrLf & "Do you want to commit the pending transactions?", _
        vbQuestion + vbYesNo)
    If nRet = vbNo Then
        m_DataBaseObject.RollbackTrans
    ElseIf nRet = vbYes Then
        m_DataBaseObject.CommitTrans
    End If
End If

If m_DBOpen Then m_DataBaseObject.Close
On Error Resume Next
Set m_DataBaseObject = Nothing
Set m_Catalog.ActiveConnection = Nothing
Set m_Catalog = Nothing

m_DBOpen = False

Exit Sub

ErrLine:
    MsgBox "Could't close the Database " & Err.Description, vbCritical

End Sub

Public Function CommitTrans() As Boolean

If m_PrevOp = op_BEGINTRANS Then
    m_PrevOp = op_COMMIT
    m_DataBaseObject.CommitTrans
    CommitTrans = True
End If

End Function

'Creates a Database file (.mdb) in Access 2000 format by taking
'the necessary data from given TAB file which is in INI file format
'with the given password.
'
'Returns
'   True    -   On successful creation of the database
'   Created on  :   07/06/2002  (dd/mm/yyyy)
'   Modified on :   11/06/2002  (dd/mm/yyyy)
'
Public Function CreateDB(ByVal strdataFile As String, Optional StrPwd As String, Optional NewDBName As String = "") As Boolean
On Error GoTo dbCreate_err
Dim nRet As Integer
'Check if the specified INI file is existing
If Dir(strdataFile, vbNormal) = "" Then
    'MsgBox "File not found - " & strDataFile, vbCritical
    MsgBox "The INI File : " & strdataFile & " not found", vbCritical, _
        "Creating DataBase"
    Exit Function
End If

Dim i As Byte
Dim strRet As String
Dim DbName As String
Dim dbFile As String
Dim DBPath As String

i = 1
Do
    'read the line containing database name
    If NewDBName = "" Then
        strRet = ReadFromIniFile("Databases", "DataBase" & i, strdataFile)
        If strRet = "" Then Exit Do
        
        'extract token containing database name
        DbName = ExtractToken(strRet, "dbName")
        If DbName = "" Then Exit Do
    
        'get the file path
        DBPath = ExtractToken(strRet, "dbPath")
        dbFile = DBPath & "\" & DbName
        ' if no path mentioned, set path as of data file
        If DBPath = "" Then
            dbFile = AppendBackSlash(FilePath(strdataFile)) & DbName
            DBPath = FilePath(dbFile)
        End If
    Else
        DBPath = FilePath(NewDBName)
        dbFile = NewDBName
        
        strRet = ReadFromIniFile("Databases", "DataBase" & i, strdataFile)
        If strRet = "" Then Exit Do
        
        'extract token containing database name
        DbName = ExtractToken(strRet, "dbName")
        If DbName = "" Then Exit Do

    End If
    'check if the database already exist
    If Dir(DBPath, vbDirectory) = "" Then
        If Not MakeDirectories(DBPath) Then
            MsgBox "Error in creating the path " & DBPath _
                & " for " & DbName, vbCritical
            'MsgBox LoadResString(gLangOffSet + 809) & DBPath, vbCritical
            Exit Function
        End If
    Else
        'check if the database exists
        dbFile = StripExtn(dbFile) & ".mdb"
        If Dir(dbFile, vbNormal) <> "" Then
            nRet = MsgBox("WARNING : " & vbCrLf & vbCrLf & "The database '" & _
                dbFile & "' already exist. If you choose to overwrite " & _
                "this file, you will loose the existing data permanantly." & _
                vbCrLf & vbCrLf & "Do you you want to overwrite this file?", _
                vbYesNo + vbCritical + vbDefaultButton2)
            If nRet = vbYes Then
                'delete the existing file
                Kill dbFile
            ElseIf nRet = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    'create the database
    Set m_Catalog = New ADOX.Catalog

    If Trim$(StrPwd) = "" Then
        m_Catalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & dbFile & ";"
    Else
        m_Catalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & dbFile & ";" & _
            "Jet OLEDB:Database Password=" & StrPwd & ";"
    End If
    
    'crate the specified tables for this database
        
    If Not CreateTBLS(StripExtn(DbName), strdataFile) Then
        MsgBox "Error in creating tables for  : " & StripExtn(DbName), vbCritical
        Exit Function
    End If

    'if any relations are specified, create them
    Dim L As Byte
    Dim RelationData() As ADOX.Key
    Dim RelnCount As Integer
    Dim SourceTableName() As String
    Dim RelnFlds() As String
    RelnCount = 0
    L = 0
    
    Do
        strRet = ReadFromIniFile("Relations", "Relation" & L + 1, strdataFile)
        If strRet = "" Then Exit Do
        ReDim Preserve RelationData(L)
        ReDim Preserve SourceTableName(L)
        ReDim Preserve RelnFlds(L)
        'SourceTableName(L) = ExtractToken(strRet, "SourceTable")
        SourceTableName(L) = ExtractToken(strRet, "ForeignTable")
        Set RelationData(L) = New ADOX.Key
        RelnCount = L + 1
        RelnFlds(L) = ExtractToken(strRet, "Field")
        With RelationData(L)
            .Name = ExtractToken(strRet, "RelationName")
            '.RelatedTable = ExtractToken(strRet, "ForeignTable")
            .RelatedTable = ExtractToken(strRet, "SourceTable")
            .Type = RelationTypeEnum(ExtractToken(strRet, "RelationType"))
        End With
        L = L + 1
    Loop
    If RelnCount > 0 Then
        If Not CreateRelation(SourceTableName(), RelationData(), RelnFlds()) Then
            MsgBox "some of the relation can not create", vbInformation, wis_MESSAGE_TITLE
        End If
    End If
    
    Dim strDBQryName As String
    Dim strQryName As String
    Dim J As Byte
    Dim ParamString As String
    L = 0
    
    strDBQryName = StripExtn(DbName) + "Query"
    Do
        strRet = ReadFromIniFile(strDBQryName, "QueryDef" & L + 1, strdataFile)
        If strRet = "" Then Exit Do
        strQryName = ExtractToken(strRet, "QueryDefName")
        SQLStmt = ExtractToken(strRet, "sql")
        J = 0
        Do
            strRet = ReadFromIniFile(strQryName, "Parameter" & J + 1, strdataFile)
            If strRet = "" Then Exit Do
            If ParamString = "" Then ParamString = "PARAMETERS "
            ParamString = ParamString + "[" + ExtractToken(strRet, "ParameterName") + "] "
            ParamString = ParamString + ExtractToken(strRet, "DataType") + ","
            J = J + 1
        Loop
        If ParamString = "" Then
            If Not CreateView(strQryName) Then
                MsgBox "Could not Create View", vbCritical
                Exit Function
            End If
        Else
            ParamString = Left(ParamString, Len(ParamString) - 1)
            ParamString = ParamString + ";"
            SQLStmt = ParamString + SQLStmt
            If Not CreateProcedure(strQryName) Then
                MsgBox "Could not Create Parameterised Procedure", vbCritical
                Exit Function
            End If
            ParamString = ""
        End If
        L = L + 1
    Loop
    
    MsgBox "Database File : " & dbFile & " Created Successfully", vbInformation, _
            "Creating Database"

    'go for next database
    i = i + 1
    Set m_Catalog = Nothing
Loop

CreateDB = True
Exit Function
dbCreate_err:
    
    MsgBox " ERROR No : " & Err.Number & vbCrLf & Err.Description & _
        " Read the help file : " & Err.HelpFile, vbInformation, wis_MESSAGE_TITLE
    Debug.Assert Err.Number
'    Resume
    Set m_Catalog = Nothing

End Function

'
Public Function CreateProcedure(ProcName As String) As Boolean

On Error GoTo ErrLine

Dim objCommand As ADODB.Command

If SQLStmt = "" Then
    MsgBox "Command Text for View is not set", vbExclamation
    Exit Function
End If

Set objCommand = New ADODB.Command
objCommand.CommandText = SQLStmt
m_Catalog.Procedures.Append ProcName, objCommand

SQLStmt = ""
CreateProcedure = True

Exit Function
ErrLine:
    MsgBox "Error in Creating Procedure " & Err.Description, vbCritical, _
        "Creating Procedure"

End Function

'
'
'This private functon creates the specified tables whose feilds are
'read from TAB file (in INI Format)
'The properties of the fields (Column Object) are,
'   1.  Jet OLEDB:Allow Zero Length
'   2.  Nullable
'   3.  Description
'   4.  AutoIncrement
'   5.  Default Value etc.
'               Created On  :   11/06/2002    (dd/mm/yyyy)
'               Modified On :   11/06/2002    (dd/mm/yyyy)
'
Private Function CreateTBLS(DbName As String, strTabFile As String) As Boolean
On Error GoTo Err_Line

Dim strFieldName As String
Dim intFieldType As Integer
Dim intLength As Integer
Dim J As Byte
Dim strTblName As String
Dim TheTables() As ADOX.Table
Dim ReqFld As Boolean
Dim K As Byte
Dim L As Byte
Dim strRet As String
Dim strValidation As String

Dim strDefault

J = 0

Do
    strTblName = ReadFromIniFile(DbName, "Table" & J + 1, strTabFile)
    If strTblName = "" Then Exit Do
'    Debug.Print strTblName & " " & J
    ReDim Preserve TheTables(J)
    Set TheTables(J) = New ADOX.Table
    TheTables(J).Name = strTblName
    Set TheTables(J).ParentCatalog = m_Catalog
    K = 0
    Do
        strRet = ReadFromIniFile(strTblName, "Field" & K + 1, strTabFile)
        If strRet = "" Then Exit Do
        With TheTables(J)
'            Debug.Assert strTblName <> "FDTrans"

            'set the field name
            strFieldName = ExtractToken(strRet, "FieldName")
            'field type
            intFieldType = FieldTypeNum(ExtractToken(strRet, "FieldType"))
            'field length
            
            intLength = Val(ExtractToken(strRet, "FieldLength"))
            strDefault = ExtractToken(strRet, "Default")
            
            If intFieldType = adVarWChar And Len(strDefault) Then strDefault = AddQuotes(CStr(strDefault), True)
            If intFieldType <> adVarWChar And Len(strDefault) Then strDefault = Val(strDefault)
            'append the field
            If intFieldType = adBoolean Then
                .Columns.Append strFieldName, intFieldType
                If Len(strDefault) Then strDefault = CBool(strDefault)
            Else
                .Columns.Append strFieldName, intFieldType, intLength
                'check if the required flag is set
                ReqFld = IIf((UCase$(ExtractToken(strRet, "Required")) _
                    = "TRUE"), True, False)
                If ReqFld Then
                    .Columns(strFieldName).Properties("Nullable") = False
                    .Columns(strFieldName).Properties("Jet OLEDB:Allow Zero Length") = False
                Else
                    .Columns(strFieldName).Properties("Nullable") = True
                    .Columns(strFieldName).Properties("Jet OLEDB:Allow Zero Length") = True
                End If
            End If

            .Columns(strFieldName).Properties("AutoIncrement") = _
                IIf((UCase$(ExtractToken(strRet, "AutoIncrement")) = "TRUE"), True, False)

            .Columns(strFieldName).Properties("Description") = ExtractToken(strRet, "Desc")
            'If Len(strDefault) Then .Columns(strFieldName).Properties("Default") = strDefault
        End With
        'go for next field
        K = K + 1
    Loop

    m_Catalog.Tables.Append TheTables(J)

    strValidation = ReadFromIniFile(strTblName, "ValidationRule", strTabFile)
    If strValidation <> "" Then _
        m_Catalog.Tables(TheTables(J).Name).Properties("Jet OLEDB:Table Validation Rule") = strValidation

'   "(IIf([Nominee] Is Not Null, [NomineeRelation] Is Not Null,[NomineeRelation] = Null)) And ([MaturityDate]>[CreateDate])"

    'create the indexes if specified
    K = 0
    Dim IndxData() As ADOX.Index
    Dim IndxCount As Integer
    Dim IndxFlds() As String
    
    IndxCount = 0
    Do
        strRet = ReadFromIniFile(strTblName, "Index" & K + 1, strTabFile)
        If strRet = "" Then Exit Do
        ReDim Preserve IndxData(K)
        ReDim Preserve IndxFlds(K)
        Set IndxData(K) = New ADOX.Index
        IndxCount = K + 1
        IndxFlds(K) = ExtractToken(strRet, "Fields")
        With IndxData(K)
            .Name = ExtractToken(strRet, "IndexName")
            .PrimaryKey = IIf(UCase$(ExtractToken(strRet, _
                "Primary")) = "TRUE", True, False)
            .Unique = IIf(UCase$(ExtractToken(strRet, _
                "Unique")) = "TRUE", True, False)
            .IndexNulls = IndexTypeNum(ExtractToken(strRet, "IndexNulls"))
        End With
        K = K + 1
    Loop

    If IndxCount > 0 Then
        If Not CreateIndexes(strTblName, IndxData(), IndxFlds()) Then
            MsgBox "can not creat index on " & strTblName & vbCrLf & Err.Description
'            Exit Function
        End If
    End If
        'go for next table
        J = J + 1
Loop

CreateTBLS = True

Exit Function

Err_Line:
    If Err.Number = 3010 Then
        Exit Function
    ElseIf Err.Number > 0 Then
        MsgBox Err.LastDllError & " :: " & Err.Description, vbCritical, _
            "Error in Creating Table : " & strTblName
        CreateTBLS = True
        'Resume
    End If

End Function
Public Function CreateView(ViewName As String) As Boolean
On Error GoTo ErrLine
Dim ViewCount As Integer
Dim objCommand As ADODB.Command

If SQLStmt = "" Then
    MsgBox "Command Text for View is not set", vbExclamation
    Exit Function
End If

Call DeleteView(ViewName)
Set objCommand = New ADODB.Command
objCommand.CommandText = SQLStmt
m_Catalog.Views.Append ViewName, objCommand
SQLStmt = ""

CreateView = True

Exit Function
ErrLine:
    MsgBox "Error in Creating View " & Err.Description, vbCritical, "Creating View"
    
End Function

 'This function will delete the specified view
' returns true after successful deletion
'Inputs : ViewName as string
Public Function DeleteView(ByVal ViewName As String) As Boolean
'Trap the error
On Error GoTo ErrLine

'Declare the variables
Dim ViewCount As Integer
Dim lpCount As Integer
Dim SqlStr As String
Dim isViewFound As Boolean

'initalise the functions
DeleteView = False

Set m_Catalog = New ADOX.Catalog

m_Catalog.ActiveConnection = m_DataBaseObject

ViewCount = m_Catalog.Views.Count - 1

isViewFound = False

For lpCount = ViewCount To 0 Step -1
    If UCase$(m_Catalog.Views.Item(lpCount).Name) = UCase$(ViewName) Then
        
        SqlStr = m_SQLStmt
        m_SQLStmt = " DROP TABLE " & ViewName
        BeginTrans
        If Not SQLExecute Then
            RollBack
            Exit Function
        End If
        CommitTrans
        m_SQLStmt = SqlStr
        isViewFound = True
        Exit For

   End If
Next lpCount

If Not isViewFound Then If Not DeleteProcedures(ViewName) Then Exit Function

DeleteView = True

Exit Function

ErrLine:
   If Err Then
      MsgBox "DeleteView: " & vbCrLf & Err.Description, vbCritical
      Exit Function
   End If

End Function


'This function will delete the specified view
' returns true after successful deletion
'Inputs : ViewName as string
Public Function DeleteProcedures(ByVal ViewName As String) As Boolean
'Trap the error
On Error GoTo ErrLine

'Declare the variables
Dim ViewCount As Integer
Dim lpCount As Integer
Dim SqlStr As String

'initalise the functions
DeleteProcedures = False

Set m_Catalog = New ADOX.Catalog

m_Catalog.ActiveConnection = m_DataBaseObject

ViewCount = m_Catalog.Procedures.Count - 1

For lpCount = ViewCount To 0 Step -1
    If UCase$(m_Catalog.Procedures.Item(lpCount).Name) = UCase$(ViewName) Then

        SqlStr = m_SQLStmt
        m_SQLStmt = " DROP TABLE " & ViewName
        BeginTrans
        If Not SQLExecute Then
            RollBack
            Exit Function
        End If
        CommitTrans
        m_SQLStmt = SqlStr
        Exit For

   End If
Next lpCount

DeleteProcedures = True

Exit Function

ErrLine:
   If Err Then
      MsgBox "DeleteProcedures: " & vbCrLf & Err.Description, vbCritical
      Exit Function
   End If

End Function

'This function fetches the data from the database assigns
'the resulted records to recordset which you have passed
'Returns number of records fetched
'Created on     -   11/06/2002      (dd/mm/yyyy)
Public Function Fetch(rstFetch As ADODB.Recordset, TheCursorType As CursorTypeEnum) As Long

'If no sqlstmt is set
If SQLStmt = "" Then
    MsgBox "No SQL Statement has been set!", vbExclamation
    Exit Function
End If

'if no database in use
If m_DataBaseObject Is Nothing Then
    MsgBox "Fetch : No connection in use!" & vbCr & vbLf & _
        "Set the database name before fetching any data", vbExclamation
    Exit Function
End If

On Error GoTo ErrLine:

Set rstFetch = New ADODB.Recordset

rstFetch.CursorLocation = adUseClient
rstFetch.Open SQLStmt, m_DataBaseObject, TheCursorType, adLockReadOnly, adCmdText
Select Case TheCursorType
    Case adOpenStatic
        Fetch = rstFetch.RecordCount
'        If Fetch = 0 Then Fetch = -1
    Case Else
        Fetch = 1
        If rstFetch.BOF And rstFetch.EOF Then Fetch = 0
End Select

ErrLine:
    If Err Then
        Fetch = 0
        MsgBox "Fetch : The following error occurred " _
            & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
        Err.Clear
        'Resume
    End If
    
    'set sqlstring to null
    m_SQLStmt = ""
    
    'Disconnect the recordset
    Set rstFetch.ActiveConnection = Nothing
    
    If Fetch = 0 Then Set rstFetch = Nothing

End Function

'
'
Public Function FetchFromProcedure(rstFetch As ADODB.Recordset, ProcName As String, _
    Param() As String) As Long
Dim ObjCmd As ADODB.Command
Dim i As Integer
' If the datbase object is not initialized, exit.
If m_DataBaseObject Is Nothing Then
    MsgBox "Fetch: No Connection is in use.  " _
        & "Set the database name before calling Fetch.", vbExclamation
    Exit Function
End If

On Error GoTo ErrLine

Set rstFetch = New ADODB.Recordset
Set ObjCmd = New ADODB.Command

Set rstFetch.ActiveConnection = m_DataBaseObject
Set ObjCmd = m_Catalog.Procedures(ProcName).Command

For i = 0 To UBound(Param())
    ObjCmd.Parameters(i).Value = Param(i)
Next

rstFetch.Open ObjCmd, , adOpenForwardOnly, adLockOptimistic, adCmdStoredProc
Set ObjCmd = Nothing
'Set rstFetch.ActiveConnection = Nothing

FetchFromProcedure = rstFetch.RecordCount

Exit Function

ErrLine:

If Err Then
    FetchFromProcedure = -1
    MsgBox "Fetch from Procedure failed " & vbCr & vbLf & Err.Number & " - " & _
        Err.Description, vbCritical
End If

End Function

'
'
'Returns the ADO field type number, given a field type string
Private Function FieldTypeNum(strFldType As String) As Integer
Select Case UCase$(strFldType)
    Case "ADTEXT", "ADVARWCHAR", "ADVARCHAR", "DBTEXT", "TEXT"
        FieldTypeNum = adVarWChar
    Case "NUMBER", "INTEGER", "DBNUMBER", "DBINTEGER", "ADINTEGER"
        FieldTypeNum = adSmallInt
    Case "DATE", "DBDATE", "ADDATE", "DATETIME"
        FieldTypeNum = adDate
    Case "ADBOOLEAN", "DBBOOLEAN", "BOOLEAN"
        FieldTypeNum = adBoolean
    Case "BYTE", "DBBYTE", "ADBYTE"
        FieldTypeNum = adUnsignedTinyInt
    Case "CHAR", "DBCHAR"
        FieldTypeNum = adVarWChar
    Case "CURRENCY", "DBCURRENCY", "ADCURRENCY"
        FieldTypeNum = adCurrency
    Case "DECIMAL", "DBDECIMAL"
        FieldTypeNum = adNumeric
    Case "DOUBLE", "DBDOUBLE", "ADDOUBLE"
        FieldTypeNum = adDouble
    Case "LONG", "DBLONG", "ADLONG"
        FieldTypeNum = adInteger
    Case "LONGBINARY", "DBLONGBINARY", "ADLONGBINARY"
        FieldTypeNum = adLongVarBinary
    Case "MEMO", "DBMEMO"
        FieldTypeNum = adLongVarWChar
    Case "SINGLE", "DBSINGLE", "ADSINGLE"
        FieldTypeNum = adSingle
    Case Else
        FieldTypeNum = 0
End Select
End Function



Private Function CreateIndexes(tbl As String, idx() As ADOX.Index, idxFlds() As String) As Boolean
On Error GoTo CreateIndex_Error
Dim TheTable As ADOX.Table
Dim TheIndex() As ADOX.Index
Dim fldArr() As String
Dim i As Byte, J As Byte

Set TheTable = m_Catalog.Tables(tbl)

For i = 0 To UBound(idx)
    With TheTable
        ReDim Preserve TheIndex(i)
        Set TheIndex(i) = New ADOX.Index
        TheIndex(i).Name = idx(i).Name
        'breakup the fields string to get the field names for index
        GetStringArray idxFlds(i), fldArr(), "+"
        
        With TheIndex(i)
            For J = 0 To UBound(fldArr)
                .Columns.Append fldArr(J)
            Next J
            .PrimaryKey = idx(i).PrimaryKey
            .IndexNulls = idx(i).IndexNulls
        End With
        .Indexes.Append TheIndex(i)
    End With
Next i

CreateIndexes = True
Exit Function

CreateIndex_Error:
    If Err.Number = 3283 Then
        Resume Next
    ElseIf Err Then
        MsgBox Err.Description, vbCritical, "Error occurred while " & _
            "creating indexes for the table : " & tbl
        Set m_Catalog = Nothing
        'Resume
    End If
End Function


Private Function CreateRelation(srcTableName() As String, Reln() As ADOX.Key, RlnFlds() As String) As Boolean
On Error GoTo CreateRelation_Error

'open the specified table
Dim TheTable As ADOX.Table
Dim FKey As ADOX.Key
Dim RelationFlds() As String
Dim i As Integer, J As Integer

For i = 0 To UBound(Reln)
    Set TheTable = m_Catalog.Tables(srcTableName(i))
    Set FKey = New ADOX.Key
    
    With FKey
        .Name = Reln(i).Name
        .Type = Reln(i).Type
        .RelatedTable = Reln(i).RelatedTable
        
        'get the flds in an array
        GetStringArray RlnFlds(i), RelationFlds, "+"
        'RelationFlds = Split(RlnFlds(i), "+")
        For J = 0 To UBound(RelationFlds)
            .Columns.Append RelationFlds(J)
        Next J
        
        J = 0
        For J = 0 To UBound(RelationFlds)
            .Columns(RelationFlds(J)).RelatedColumn = RelationFlds(J)
        Next J
    End With
    'm_Catalog.
    TheTable.Keys.Append FKey
    Set TheTable = Nothing
    Set FKey = Nothing
Next i


CreateRelation = True
Exit Function

CreateRelation_Error:
    If Err.Number = 9 Then
        Exit Function
        Set m_Catalog = Nothing
    ElseIf Err.Number = 3283 Then 'primary key already exists
        Resume Next
    ElseIf Err Then
        MsgBox Err.Description, vbCritical, Reln(i).Name
        CreateRelation = False
        'Resume
        Set m_Catalog = Nothing
    End If
End Function

'
'returns the ADO indextype number, given index type string
Private Function IndexTypeNum(strIndexType As String) As Integer
Select Case UCase$(strIndexType)
    Case "DISALLOWNULLS"
        IndexTypeNum = adIndexNullsDisallow
    Case "IGNORENULLS"
        IndexTypeNum = adIndexNullsIgnore
    Case "IGNOREANYNULLS"
        IndexTypeNum = adIndexNullsIgnoreAny
    Case Else
'        MsgBox "Invalid index type entered in the tab file"
        IndexTypeNum = adIndexNullsDisallow
'        End
End Select
End Function


'
'This Function opens a connection for a data source
'Returns true if connection is successful
'Arguments : Database name as string
'          : password
'           Created on      :   11/06/2002
'           Modified on     :   11/06/2002
Public Function OpenDB(strDBName As String, pwd As String) As Boolean
On Error GoTo ErrLine

If strDBName = "" Then Exit Function
If m_DataBaseObject Is Nothing Then Set m_DataBaseObject = New ADODB.Connection

m_DataBaseObject.Mode = adModeShareDenyNone
   
On Error GoTo ErrLine
'Open a connection using Microsoft Jet 4.0
m_DataBaseObject.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                      "Data Source=" & strDBName & ";" & _
                      "jet OLEDB:Database Password=" & pwd & ";"

m_DBOpen = True
If m_Catalog Is Nothing Then Set m_Catalog = New ADOX.Catalog
If m_Catalog.ActiveConnection Is Nothing Then Set m_Catalog.ActiveConnection = m_DataBaseObject

OpenDB = True

Exit Function

ErrLine:
    MsgBox "Cannot Open Database : " & Err.Description, vbCritical, "Opening Database"
End Function

'
'returns the ADO Relation type number, given relation type string
Private Function RelationTypeEnum(RelationType As String) As Integer
Select Case UCase$(RelationType)
    Case "ADKEYFOREIGN"
        RelationTypeEnum = adKeyForeign
    Case "ADKEYPRIMARY"
        RelationTypeEnum = adKeyPrimary
    Case "ADKEYUNIQUE"
        RelationTypeEnum = adKeyUnique
    Case Else
        RelationTypeEnum = adKeyForeign
End Select
End Function


Public Function RollBack() As Boolean

    If m_PrevOp = op_BEGINTRANS Then
        m_PrevOp = op_ROLLBACK
        m_DataBaseObject.RollbackTrans
        RollBack = True
    End If

End Function

Public Function SQLExecute() As Boolean
On Error GoTo ErrLine

Dim TheRecordsAffected As Long

'Check if the sqlstring is set
If Trim$(SQLStmt) = "" Then
    MsgBox "No SQL Statement is set!", vbExclamation
    Exit Function
End If

'Execute the action query
m_DataBaseObject.Execute SQLStmt, TheRecordsAffected, adExecuteNoRecords
m_SQLStmt = ""
SQLExecute = True

ErrLine:
    
    'Handle the errors here...
    If Err Then
        MsgBox "SQLExecute: The following error occurred." _
            & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
         SQLExecute = False
         
         'Resume
    End If
    
End Function
Public Function WISCompactDB(ByVal DbName As String, ByVal OldPwd As String, ByVal NewPwd As String) As Boolean
Dim je As JRO.JetEngine

Dim PauseTime, Start, Finish, TotalTime
    PauseTime = 1.5 ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer  ' Set end time.
    TotalTime = Finish - Start  ' Calculate total time.

CloseDB

Set je = New JRO.JetEngine
' Make sure there isn't already a file with the
' name of the compacted database.
Debug.Print FileLen(DbName)

Screen.MousePointer = vbHourglass
On Error Resume Next
Err.Clear
If Dir(App.Path & "\TempDB.mdb") <> "" Then _
   Kill App.Path & "\TempDB.mdb"
' Compact the database specifying the new database password
je.CompactDatabase "Data Source=" & DbName & ";" & _
   "Jet OLEDB:Database Password=" & OldPwd & ";", _
   "Data Source=" & App.Path & "\TempDB.mdb;" & _
   "Jet OLEDB:Database Password=" & NewPwd & ";"
 
If Err.Number Then GoTo LastLine
' Delete the original database
Kill DbName
' Rename the file back to the original name
Name App.Path & "\TempDB.mdb" As DbName
Set je = Nothing

Debug.Print FileLen(DbName)

LastLine:
If Not OpenDB(DbName, NewPwd) Then _
    MsgBox "Could Not Open The Database", vbCritical

WISCompactDB = True

End Function

Public Property Get SQLStmt() As String
    SQLStmt = m_SQLStmt
End Property

Public Property Let SQLStmt(ByVal vNewValue As String)
    m_SQLStmt = vNewValue
End Property

