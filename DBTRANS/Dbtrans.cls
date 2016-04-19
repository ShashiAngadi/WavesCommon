VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
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

' Structure for holding the fields info The querydef.
' This is used by CreateQryDefs(CreateDB) function to create queryDef object.
Private Type QryStruct
    Name As String
    Type As String
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
'
Public Function CheckDbStructure(strTabFile As String, strDbName As String, Optional strPwd As String, _
        Optional boolIndex As Boolean, Optional boolReln As Boolean) As Boolean

CheckDbStructure = False
'Read each database name and create the db.
Dim strRet As String
Dim DBName As String
Dim Rst As Recordset

'Get The Db Name
Dim Pos As Integer
Dim PrevPos As Integer
Dim Count As Integer
Dim RelnTest As Relation
Dim newVal As Variant
Dim IndxTest As Index
Dim I As Byte
Dim j As Byte

'Get the "\" from StrDbName
Do
    Pos = InStr(PrevPos + 1, strDbName, "\", vbTextCompare)
    If Pos = 0 Then
        Pos = InStr(PrevPos + 1, strDbName, ".", vbTextCompare)
        If Pos <> 0 Then
            DBName = Left(Mid(strDbName, PrevPos + 1), Pos - PrevPos - 1)
        Else
            DBName = Mid(strDbName, PrevPos + 1)
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
   Dim TheField As Field
   
   ReDim NewTableDet(0)
    'Check if the file path of the database
    'is existing.  If not create it.
   On Error Resume Next
    'Create the database.
    Dim db As Database
    If Trim$(strPwd) <> "" Then
        Set db = OpenDatabase(strDbName, False, False, ";pwd=" & strPwd)
    Else
        Set db = OpenDatabase(strDbName, False, False)
    End If
    
    If Err.Number Then
        MsgBox "unable to find the Database", vbInformation, "WIS"
        Exit Function
    End If
    If boolReln Or boolIndex Then  'Delete all the relataions
        Do
            Count = db.Relations.Count - 1
            If Count < 0 Then Exit Do
            db.Relations.Delete db.Relations(Count).Name
        Loop
        db.Relations.Refresh
    End If
   'Check the specified tables for this db.
    j = 1
    CreateNewTable = False
    Do
        'Read the table name.
        Dim strTblName As String
        Dim tblData() As TabStruct
        
        TblExist = True
        strTblName = ReadFromIniFile(StripExtn(DBName), "Table" & j, strTabFile)
        If strTblName = "" Then Exit Do
        
        'Check Whethr Tale exists or Not.
        Err.Clear
        Set Rst = db.OpenRecordset("SELECT TOP 1 * From " & strTblName)
        If Err.Number <> 0 Then  'IF Table Not Existing
            Err.Clear
            TblExist = False
        End If
        Set Rst = Nothing
          'Chenges tobe Made for existing mdb with tab file
          'Change the structure, while comparing the Tabfile with mdb file.
          'If user creates one OR more fileds in existing tab file
          'then ask for message to he has created the new Field.---siddu
        
          'MsgBox("You have added extra fields" & TheField.Name & " in the tab file which creates new fields " & vbCrLf & _
           '     "Do You want to create new fieldname", vbYesNoCancel, App.EXEName) = 0
                 
          'Read the field details for this table into an array.
        Dim K As Byte
        K = 0
        ReDim tblData(0)
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Field" & K + 1, strTabFile)
            If strRet = "" Then Exit Do
            ' Add to fields array.
            ReDim Preserve tblData(K)
            'c:\rdcc_mis\mis
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
            'now Check whether Field Exists or not
            If TblExist Then 'If table exist then only Check The Field Detail
                Set TheTable = db.TableDefs(strTblName)
                If boolIndex Then  'Delete all indexes of this table
                    Do
                        Count = TheTable.Indexes.Count - 1
                        If Count < 0 Then Exit Do
                        TheTable.Indexes.Delete TheTable.Indexes(Count).Name
                    Loop
                    Err.Clear
                    TheTable.Indexes.Refresh
                End If
                
                Set Rst = db.OpenRecordset("SELECT  " & tblData(K).Field & " FROM " & strTblName)
                If Err.Number <> 0 Then
                    Set Rst = Nothing
                    FldExist = False
                    Err.Clear
                    Set TheField = TheTable.CreateField(tblData(K).Field, tblData(K).Type)
                    If tblData(K).Type = dbText Or tblData(K).Type = dbMemo Then
                       TheField.AllowZeroLength = IIf(tblData(K).Required, False, True)
                    End If
                    TheField.Required = tblData(K).Required
                    TheTable.fields.Append TheField
                    If Err.Number Then
                        MsgBox Err.Description
                        Exit Function
                    End If
                Else
                'CHeck for the DataType in the mdb.
                'If posible change Type.like if their is date in tab file and date/time in
                'mdb file then remove or overwrite to the tabfile.-16/3/02
                
                    Set TheTable = db.TableDefs(strTblName)
                    Set TheField = TheTable.fields(tblData(K).Field)
                    Pos = 0
                    If TheField.Type <> tblData(K).Type Then
                        Pos = MsgBox("The Data type of " & tblData(K).Field & " is different from the Tab file " & vbCrLf & _
                            "Do You want to change the datafield", vbYesNoCancel, App.EXEName)
                    End If
                    If Pos = vbYes Then
                        Rst.Close
                        Set TheField = Nothing
                        'Then Create a Temp Field Transfer the Data To That field
                        Set TheField = TheTable.CreateField("WISTEMPWIS", dbText, 250)
                        TheField.AllowZeroLength = True: TheField.Required = False
                        TheTable.fields.Append TheField
                        'Db.TableDefs.Append TheTable
                        Set Rst = db.OpenRecordset(strTblName)
                        'Trnasfer the Data From to new field
                        Rst.MoveFirst
                        While Not Rst.EOF
                            Rst.Edit
                            newVal = Rst(K)
                            If Err Then Err.Clear: newVal = ""
                            Rst("WISTEMPWIS") = CStr(newVal)
                            If Err Then Err.Clear: newVal = ""
                            Rst.Update
                            Rst.MoveNext
                        Wend
                        Rst.Close
                        If Err.Number = 0 Then
                        'Now Delete the Field & Create new with new data type
                            Set TheField = Nothing
                            Set TheTable = Nothing
                            'Before deleting the field delete the relation and indexec
                            While db.Relations.Count > 0
                                db.Relations.Delete _
                                    db.Relations(db.Relations.Count - 1).Name
                            Wend
                            db.Relations.Refresh
                            'Now Delete the index
                            Set TheTable = db.TableDefs(strTblName)
                            While TheTable.Indexes.Count > 0
                                TheTable.Indexes.Delete _
                                    TheTable.Indexes(TheTable.Indexes.Count - 1).Name
                            Wend
                            TheTable.Indexes.Refresh
                            With TheTable
                                'Delete the field
                                .fields.Delete tblData(K).Field
                                
                                TheTable.fields.Refresh
                                Set TheField = .CreateField(tblData(K).Field, tblData(K).Type)
                                If tblData(K).Type = dbText Or tblData(K).Type = dbMemo Then
                                   TheField.AllowZeroLength = IIf(tblData(K).Required, False, True)
                                End If
                                TheField.Required = tblData(K).Required
                                .fields.Append TheField
                            End With
                            Set Rst = db.OpenRecordset(strTblName)
                            'Trnasfer the Data From New field to old field
                            Rst.MoveFirst
                            While Not Rst.EOF
                                Rst.Edit
                                If tblData(K).Type = dbText Then
                                    If Rst("WISTEMPWIS") <> Null Then
                                        newVal = AddQuotes(Rst("WISTEMPWIS"), True)
                                    End If
                                ElseIf tblData(K).Type = dbDate Then
                                    newVal = CDate(Rst("WISTEMPWIS"))
                                    If Err Then
                                        Err.Clear
                                        newVal = Null
                                    End If
                                Else
                                    newVal = CLng(Rst("WISTEMPWIS"))
                                End If
                                Rst(tblData(K).Field) = newVal
                                Rst.Update
                                Rst.MoveNext
                            Wend
                            Rst.Close
                            Err.Clear
                        End If
                        'Now Delete the New Temp Field
                        TheTable.fields.Delete "WISTEMPWIS"
                        Err.Clear
                    End If
                End If
                Set Rst = Nothing
            Else
                CreateNewTable = True
                NewTableDet(UBound(NewTableDet)) = tblData(K)
                If Err.Number Then
                    strDbName = strDbName
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
            If Not CreateTBL(db, strTblName, NewTableDet()) Then
                Exit Function
                'GoTo dbCreate_err
            End If
            ReDim NewTableDet(0)
        End If
        
        'Assign the table
        Set TheTable = db.TableDefs(strTblName)
        
        'If any indexes are specified, create them.
        K = 0
        
        Dim IndxData() As idx
        Count = 0
        On Error Resume Next
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Index" & K + 1, strTabFile)
            If strRet = "" Then Exit Do
            'ReDim Preserve IndxData(K)
            ReDim Preserve IndxData(Count)
            With IndxData(Count)
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
            'Check whether  this Index exists or not
            Set IndxTest = TheTable.Indexes(IndxData(Count).Name)
            If Err.Number = 3265 Then
                Count = Count + 1
                Err.Clear
            ElseIf Err.Number Then
                GoTo Err_Line
            Else
                TheTable.Indexes.Delete IndxData(Count).Name
                If Err.Number = 3281 Then
                    Err.Clear
                Else
                    Count = Count + 1
                End If
            End If
            K = K + 1
        Loop
        'RaiseEvent CreateDBStatus("Creating Index For " & strTblName, m_FieldsCreated / m_FieldCount)
        If Count > 0 Then
            If Not CreateIndexes(db, strTblName, IndxData()) Then
                Exit Function
            End If
        End If
        ' Increment the table count variable "j"
        j = j + 1
    Loop
    
    ' If any Relations are specified, create them.
    Dim L As Byte
    L = 0
    Dim RelnData() As RelnStruct
    Count = 0
    Do
        strRet = ReadFromIniFile("Relations", _
                    "Relation" & L + 1, strTabFile)
        If strRet = "" Then Exit Do
        ReDim Preserve RelnData(Count)
        With RelnData(L)
            .Name = ExtractToken(strRet, "RelationName")
            .Field = ExtractToken(strRet, "Field")
            .ForiegnTable = ExtractToken(strRet, "ForiegnTable")
            .SourceTable = ExtractToken(strRet, "SourceTable")
        End With
        Set RelnTest = db.Relations(RelnData(L).Name)
        If Err.Number = 3265 Then 'If relationship not exists
            Count = Count + 1
            Err.Clear
        ElseIf Err.Number Then
            MsgBox "Unbale to correct the relationship", vbInformation, App.EXEName
            Exit Function
        Else
            db.Relations.Delete RelnData(Count).Name
            If Err.Number Then
                Err.Clear
            Else
                db.Relations.Refresh
                Count = Count + 1
            End If
        End If
        L = L + 1
    Loop
    If Count Then
         If Not CreateRelation(db, RelnData) Then
            MsgBox "Unbale to correct the relationship", vbInformation, App.EXEName
            Exit Function
         End If
    End If
    
    Dim strQryName As String
    Dim qryData() As QryStruct
    Dim Sqlstr As String
    Dim tmpQryDef As QueryDef
    
    j = 0: Count = 0
    Do
        'Read the Query name.
        strRet = ReadFromIniFile(StripExtn(DBName) & "Query", "QueryDef" & j + 1, strTabFile)
        strQryName = ExtractToken(strRet, "QueryDefName")
        Sqlstr = ExtractToken(strRet, "Sql")
        
        If strQryName = "" Then Exit Do
    
        'Read the Prameters details for this query into an array.
        K = 0
        ReDim qryData(0)
       
        'RaiseEvent CreateDBStatus("Reading Prameters of " & DBName, m_FieldsCreated / m_FieldCount)
        Do
            strRet = ReadFromIniFile(strQryName, _
                        "Parameter" & K + 1, strTabFile)
            If strRet = "" Then Exit Do
            ' Add to fields array.
            ReDim Preserve qryData(K)
            With qryData(K)
                ' Set the Prameter name.
                .Name = ExtractToken(strRet, "ParameterName")
                ' Set the prameter data type.
                .Type = ExtractToken(strRet, "DataType")
            End With
    
            ' Increment the count variable "k"
            K = K + 1
        Loop
        
        'Create the table.
        'On Error Resume Next
        'If This Query exists the delete it and create new
        Set tmpQryDef = db.QueryDefs(strQryName)
        If Err.Number = 0 Then
            db.QueryDefs.Delete strQryName
        Else
            Err.Clear
        End If
        'On Error GoTo Err_Line
        'RaiseEvent CreateDBStatus("Creating Query " & strQryName & " For " & DBName, M_FieldsCreated / m_FieldCount)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If Not CreateQryDefs(db, strQryName, Sqlstr, qryData()) Then
                MsgBox "Error in creating Querydefs", vbInformation, App.EXEName
                Exit Function
            End If
        End If
        j = j + 1
    Loop
    
CheckDbStructure = True

Err_Line:
If Err.Number Then MsgBox "Error In Checking database" & vbCrLf & Err.Description, vbInformation, App.EXEName
End Function

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

If Not m_DataBaseObject Is Nothing Then
    m_DataBaseObject.Close
End If

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
Dim fsize As Long
On Error GoTo ErrLine
CommitTrans = False
    FIleNo = FreeFile
    If m_PrevOp = op_BEGINTRANS Then
        'First open the file
        If m_LogFileName <> "" Then 'YOu should have a log
            Open m_LogFileName For Binary As #FIleNo
            fsize = LOF(FIleNo)
            Seek #FIleNo, fsize + 1
            Put #FIleNo, , m_SQLLog
            Close #FIleNo
            m_SQLLog = ""
        End If
        m_PrevOp = op_COMMIT
        Workspaces(m_TransCount - 1).CommitTrans
        m_TransCount = m_TransCount - 1
        CommitTrans = True
    ElseIf MultiTrans Then
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
'For Distination Databace put Db name as same with suffix "dup"
strDupDb = strDupDb & "dup.MDB"

If Dir(strDupDb) <> "" Then Kill strDupDb

Call DBEngine.CompactDataBase(M_dbName, strDupDb, , dbLangGeneral & ";pwd=" & M_DbPwd)
Call DBEngine.RepairDatabase(M_dbName)
Call Me.OpenDB(M_dbName, M_DbPwd)

End Function

'
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

Public Function CreateDB(ByVal strdataFile As String, Optional strPwd As String) As Boolean
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
'm_TableCount = GetTableCount(strdataFile)
'm_FieldCount = GetFieldCount(strdataFile)

' Read each database name and create the db.
Dim I As Byte, strRet As String
Dim DBName As String, dbFile As String
Dim DBPath As String
I = 1
Do
    ' Read the dbname from datafile.
    strRet = ReadFromIniFile("Databases", "DataBase" & I, strdataFile)
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
    Dim db As Database
    Dim strLocale As String
    If Trim$(strPwd) = "" Then
        Set db = CreateDatabase(dbFile, dbLangGeneral)
    Else
        Set db = CreateDatabase(dbFile, dbLangGeneral & ";pwd=" & strPwd)
    End If

    'Create the specified tables for this db.
    Dim j As Byte
    j = 1
    Do
        ' Read the table name.
        Dim strTblName As String
        Dim tblData() As TabStruct
        strTblName = ReadFromIniFile(StripExtn(DBName), "Table" & j, strdataFile)
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
        If Not CreateTBL(db, strTblName, tblData()) Then
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
            If Not CreateIndexes(db, strTblName, IndxData()) Then
                GoTo dbCreate_err
            End If
        End If
        ' Increment the table count variable "j"
        j = j + 1
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
    
    Call CreateRelation(db, RelnData)
    
    
    Dim strQryName As String
    Dim qryData() As QryStruct
    Dim Sqlstr As String
    j = 0
    Do
        ' Read the Query name.
        strRet = ReadFromIniFile(StripExtn(DBName) & "Query", "QueryDef" & j + 1, strdataFile)
        strQryName = ExtractToken(strRet, "QueryDefName")
        Sqlstr = ExtractToken(strRet, "Sql")
        
        If strQryName = "" Then Exit Do
    
        ' Read the Prameters details for this query into an array.
        K = 0
        ReDim qryData(0)
        
        'RaiseEvent CreateDBStatus("Reading Prameters of " & DBName, m_FieldsCreated / m_FieldCount)
        Do
            strRet = ReadFromIniFile(strQryName, _
                        "Parameter" & K + 1, strdataFile)
            If strRet = "" Then Exit Do
            
            ' Add to fields array.
            ReDim Preserve qryData(K)
            With qryData(K)
                ' Set the Prameter name.
                .Name = ExtractToken(strRet, "ParameterName")
                ' Set the prameter data type.
                .Type = ExtractToken(strRet, "DataType")
            End With
    
            ' Increment the count variable "k"
            K = K + 1
        Loop
    
        ' Create the table.
        'RaiseEvent CreateDBStatus("Creating Query " & strqryName & " For " & DBName, m_FieldsCreated / m_FieldCount)
        If Not CreateQryDefs(db, strQryName, Sqlstr, qryData()) Then
            MsgBox "Error in creating Querydefs", vbInformation, App.EXEName
            Exit Function
        End If
        j = j + 1
    Loop
    
    ' Increment the DB count variable "i"
    I = I + 1
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
        MsgBox Err.Description & " in table " & strTblName, vbCritical
    End If
    'Resume
End Function
Private Function CreateQryDefs(db As Database, QueryName As String, SqlString As String, ParamStr() As QryStruct) As Boolean
'
'   This routine creates the specified Query def whose parameters are
'   defined in the array ParamStr() OF TYPE QRYSTRUCT
'   WHERE QRY STRUCT AS OF FOLLOWING TYPE
'   Name  :NAme of the paramater
'   Type  :Specifies the data type of paramente

Dim TheQuery As QueryDef
Dim TheParam() As Parameter, TheIndex As Index
Dim I As Integer
Dim strParam  As String

Err.Clear
On Error GoTo Err_Line

'First Get the Intofrmation  Of Parameters if Any
'If ParamStr Is Nothing Then ReDim ParamStr(0)
strParam = "PARAMETERS"
For I = LBound(ParamStr) To UBound(ParamStr)
    If ParamStr(0).Name = "" Then Exit For
    If ParamStr(0).Type = "" Then Exit For
    strParam = strParam & " " & ParamStr(I).Name & " " & FieldTypeStr(FieldTypeNum(ParamStr(I).Type)) & ","
Next
'Now Check for any parameters exists or not
If Right(strParam, 1) = "," Then  'There are prameters
    strParam = Left$(strParam, Len(strParam) - 1) & ";"
Else
    strParam = ""
End If

Dim Sqlstr As String

Sqlstr = strParam & SqlString
strParam = ""

' Create the table.
'RaiseEvent UpdateStatus("Creating the Query " & qryName & "...")
Set TheQuery = db.CreateQueryDef(QueryName, Sqlstr)
    'Set TheQuery = db.CreateQueryDef("", _
        "PARAMETERS AsOnDate DateTime; " & _
        "SELECT EmployeeID, COUNT(OrderID) AS NumOrders " & _
        "FROM Orders WHERE ShippedDate BETWEEN " & _
        "[dteBegin] AND [dteEnd] GROUP BY EmployeeID " & _
        "ORDER BY EmployeeID")

'Db.QueryDefs.Append TheQuery

' Create and add the fields.
Set TheQuery = Nothing
CreateQryDefs = True
Exit Function

Err_Line:
    If Err.Number = 3010 Then 'table already exists
        Exit Function
    ElseIf Err.Number > 0 Then
         'MsgBox Err.Description, vbCritical
        MsgBox Err.Description
        CreateQryDefs = False
    'Resume
    End If


End Function

Private Function CreateRelation(db As Database, Reln() As RelnStruct) As Boolean
 On Error GoTo CreateRelation_Error


If Reln(0).Name = "" Then GoTo ExitLine

' Open the specified table.
Dim TheTable As TableDef
Dim frnTable As TableDef
Dim srcField As Field
Dim frnField As Field
Dim TheRelation As Relation
Dim fldArr() As String
Dim I As Integer, j As Integer

For I = 0 To UBound(Reln)
    With db
        'Get the table reference of the source and foriegn table
        Set TheTable = db.TableDefs(Reln(I).SourceTable)
        Set frnTable = db.TableDefs(Reln(I).ForiegnTable)

        ' Create Field object for the Fields collection of the
        ' new Relation object. Set the Name and ForeignName
        ' properties based on the fields to be used for the
        ' relation.
        
        GetStringArray Reln(I).Field, fldArr(), "+"
        
        'Before Creating the Relation Check whether
        'Fields are of same data type in the Source & target tables
        For j = 0 To UBound(fldArr)
            Set srcField = TheTable.fields(fldArr(j))
            Set frnField = frnTable.fields(fldArr(j))
            If srcField.Type <> frnField.Type Then
                MsgBox "Field data type has to be same in source & foriegn Table", vbInformation, App.EXEName
                GoTo NextRelation
            End If
            Set srcField = Nothing
            Set frnField = Nothing
        Next
        
        'Create Relation object
        Set TheRelation = .CreateRelation(Reln(I).Name, _
                Reln(I).SourceTable, Reln(I).ForiegnTable, dbRelationUpdateCascade)
        
        With TheRelation
            For j = 0 To UBound(fldArr)
                .fields.Append .CreateField(fldArr(j))
                .fields(fldArr(j)).ForeignName = fldArr(j)
            Next
        End With
        
        'Append the relation object
        .Relations.Append TheRelation
    End With
NextRelation:
Next

ExitLine:
CreateRelation = True

Exit Function

CreateRelation_Error:
    If Err.Number = 9 Then
        GoTo ExitLine
    ElseIf Err.Number = 3283 Then   ' Primary key already exists.
        Resume Next
    ElseIf Err.Number = 3265 Then   ' field not exist either in source or forign table
        MsgBox "Field not exists to create relation", vbInformation, App.EXEName
        GoTo NextRelation
        'Resume Next
    ElseIf Err.Number = 3609 Then   ' field not exist either in source or forign table
        MsgBox "Indexing Required for relation " & Reln(I).Name, vbInformation, App.EXEName
        GoTo NextRelation
        'Resume Next
    ElseIf Err Then
        'MsgBox Err.Description, vbCritical
        If MsgBox(Err.Description & Reln(I).Name & vbCrLf & " Continue ?", vbYesNo + vbQuestion, App.EXEName) = vbYes Then Resume Next
    End If

End Function

'This Function is used to drop the table form the database
Public Function DropTable(trDbName As String, StrTabName As String) As Boolean
'Declare the variables
Dim TheIndex As Index
Dim TheTable As TableDef
Dim db As Database

Dim TableCount As Integer
Dim MaxIndxCount As Integer
Dim IndxCount As Integer
Dim maxTableCount As Integer

Dim Sqlstr As String
Dim strRet As String
Dim tabName As String
Dim strIndxname As String

tabName = StrTabName

''TO be modified for removing the indexes~~~sIDDU
''Get the database objects
'Set db = m_DataBaseObject
''
''Count the Tables
'MaxTableCount = db.TableDefs.Count - 1
'For tableCount = 0 To MaxTableCount
'    Set thetable = db.TableDefs(tableCount)
'
'tableCount = tableCount + 1
'StrRet = thetable.Name
'
'Set TheIndex = thetable.Indexes(IndxCount)
''
'strIndxname = TheIndex.Name
'
'Sqlstr = "Drop Index " & strIndxname & " ON " & StrRet
'gDBTrans.SQLStmt = Sqlstr
'
'gDBTrans.BeginTrans
'
''If gDBTrans.SQLFetch <= 0 Then Exit Function
''Check for the Indexes
'
'If Not gDBTrans.SQLExecute Then
'   gDBTrans.RollBack
'   Exit Function
'End If
'
'Next tableCount

'gDBTrans.CommitTrans

'Drop the Table
Sqlstr = tabName

gDBTrans.SQLStmt = Sqlstr
gDBTrans.BeginTrans

If Not gDBTrans.SQLExecute Then
   gDBTrans.RollBack
   Exit Function
End If

'Set TheIndex = thetable.Indexes(IndxCount)
'
''Line for indexes
'strIndxname = TheIndex.Name
'
'Sqlstr = "Drop Index " & strIndxname & " ON " & TabName
'gDBTrans.SQLStmt = Sqlstr
'
'gDBTrans.BeginTrans
'
''If gDBTrans.SQLFetch <= 0 Then Exit Function
''Check for the Indexes

gDBTrans.CommitTrans

DropTable = True

End Function

Public Function GetDataObject() As Database
    
    Set GetDataObject = m_DataBaseObject
End Function

' Retrieves the value for a specified token
' in a given source string.
' The source should be of type :
'       name1=value1;name2=value2;...;name(n)=value(n)
'   similar to DSN strings maintained by ODBC manager.
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

'
'
Private Function CreateIndexes(db As Database, Tbl As String, Indx() As idx) As Boolean
On Error GoTo CreateIndex_Error
' Open the specified table.
Dim TheTable As TableDef
Dim TheIndex As Index
Dim fldArr() As String
Dim I As Integer, j As Integer

Set TheTable = db.TableDefs(Tbl)

For I = 0 To UBound(Indx)
    With TheTable
        Set TheIndex = .CreateIndex(Indx(I).Name)
        ' Breakup the fields string to get the field names for index.
        GetStringArray Indx(I).fields, fldArr(), "+"
        With TheIndex
            For j = 0 To UBound(fldArr)
                .fields.Append .CreateField(fldArr(j))
            Next
            .Primary = Indx(I).Primary
            '.Required = Indx(i).Required
            .Unique = Indx(I).Unique
        End With
        .Indexes.Append TheIndex
    End With
NextIndex:
Next

CreateIndexes = True
CreateIndex_Error:
    If Err.Number = 3283 Then   ' Primary key already exists.
        Err.Clear
        Indx(I).Primary = False
        I = I - 1
        Resume Next
    ElseIf Err.Number = 3022 Then
        Err.Clear
        GoTo NextIndex
    ElseIf Err Then
        MsgBox Err.Description, vbCritical
        'Resume
       Err.Clear
    End If
End Function

'
'This function is used to Creating the tab file
'with the mdb file
'
Public Function CreateTabFile(IniFile As String) As Boolean

On Error Resume Next

Dim strSection As String
Dim strKey As String
Dim strValue As String
Dim StrField As String

Dim Count As Integer
Dim MaxCount As Integer

Dim TheTable As TableDef
Dim TheField As Field
Dim TheIndex As Index
Dim TheRelation As Relation
Dim TheParameter As Parameter

Dim TableCount As Integer
Dim FieldCount As Integer
Dim maxTableCount As Integer
Dim MaxFieldCount As Integer

Dim rel As RelnStruct
Dim frnTable As TableDef
Dim maxRelCount As Integer
Dim RelCount As Integer
Dim TheString As String
Dim strDbName As String
Dim strPwd As String

Dim Pos As Integer
Dim dbNO As Integer
Dim RetStr As String

Dim strWrite As String
Dim StrName As String
Dim db As Database

Dim ws As Workspaces
Dim DBPath As String
Dim strFile As String

Dim StrTab As String
Dim tabSection As String
Dim tabKey As String

Dim MaxIndxCount As Integer
Dim StrInd As String
Dim IndSection As String
Dim IndKey As String
Dim IndxCount As Integer

Dim TheQuery As QueryDef
Dim QueCount As Integer
Dim MaxQueCount As Integer
Dim Sqlstr As String

'Declare the variables for the parameter name
Dim ParaCount As Integer
Dim maxParaCount As Integer
Dim strParaSection As String
Dim strParaValue As String

Const conDelim = ";"
'Give the user to create as many as tab files on the
'same d/b.Ask for the user to Specify his own
'Location Where he wants for TabFile.-----Siddu
IniFile = frmCreatDb.txtTabPath
'strPwd = strPwd & "WIS!@#"
Set ws = DBEngine.Workspaces

Set m_DataBaseObject = Me.GetDataObject

'Read form the Inifile
RetStr = ReadFromIniFile("DataBases", vbNullString, IniFile)

If RetStr = "" Then
    dbNO = 1 'No d/b no has written.
Else
    MsgBox "Already database number has written once.", vbInformation
End If

'Get the Data Base Name and get the database Label and databasenumber

strSection = "DATABASES"
strKey = "DATABASE" & dbNO
strDbName = m_DataBaseObject.Name
'Check for the path here
Count = 0
Do
        Count = InStr(Count + 1, strDbName, "\")
        If Count = 0 Then Exit Do
        DBPath = Left$(strDbName, Count - 1)
Loop

If DBPath = "" Then
  DBPath = App.Path
Else
    strDbName = Mid(strDbName, Len(DBPath) + 2)
End If

'Now remove the extension from the dbanme
Count = InStr(1, strDbName, ".")

If Count Then strDbName = Left(strDbName, Count - 1)
'Specify the (DataBase) DB Path
    strWrite = "Dbname=" & strDbName & conDelim & "PATH=" & DBPath

    Call WriteToIniFile(strSection, strKey, strWrite, IniFile)

    tabSection = "TABLES"
    tabKey = "TABLE"
    Set db = m_DataBaseObject

   'Count the tables including system tables...
    maxTableCount = db.TableDefs.Count - 1
    For TableCount = 0 To maxTableCount
        Set TheTable = db.TableDefs(TableCount)
    
   'Do not Show the System files in the tab file.--modified on 22/03/02
   If db.TableDefs(TableCount).Attributes And dbSystemObject Then GoTo NextCount
     strSection = strDbName
     strKey = TheTable.Name
     strKey = "Table" & TableCount + 1
     
     If Not WriteToIniFile(strSection, strKey, TheTable.Name, IniFile) Then Exit Function
          strSection = TheTable.Name
          MaxFieldCount = TheTable.fields.Count - 1
          For FieldCount = 0 To MaxFieldCount
            Set TheField = TheTable.fields(FieldCount)
              strKey = "Field" & FieldCount + 1
              strValue = ""
              strValue = strValue & "FieldName=" & TheField.Name & conDelim
              strValue = strValue & "FieldType=" & FieldTypeStr(TheField.Type) & conDelim
              strValue = strValue & "FieldLength=" & TheField.Size & conDelim
              strValue = strValue & "Required=" & TheField.Required & conDelim
              'write 'True' Part of the AutoIncrement
              If Not TheField.Attributes And dbAutoIncrField Then
                  strValue = strValue & "AutoIncrement=" & TheField.Required
              End If
              'Find the Position of the "autoincrement=False"
              'If it is False Then do not write
              Pos = 0
              Do
                   Pos = InStr(1, strValue, ";AutoIncrement=False")
                   If Pos = 0 Then Exit Do
                   strValue = Left$(strValue, Pos - 1)
                   Debug.Print strValue
              Loop
              'Now Writ to to the ini file
              If Not WriteToIniFile(strSection, strKey, strValue, IniFile) Then Exit Function
          Next
        'Check for the Indexes in the table
         MaxIndxCount = TheTable.Indexes.Count - 1
            For IndxCount = 0 To MaxIndxCount
               Set TheIndex = TheTable.Indexes(IndxCount)
                 IndSection = TheTable.Name
                 IndKey = TheIndex.Name
                 IndKey = "Index" & IndxCount + 1
                 
                 StrInd = ""
                 StrInd = StrInd & "IndexName=" & TheIndex.Name & conDelim
                 
                 'Remove the very first '+'
                 StrField = TheIndex.fields
                 StrField = Mid(StrField, 2)
                 
                 'Remove the ";" if there are more than one index fields and substitute "+"
                 Pos = 0
                 Do
                    Pos = InStr(Pos + 1, StrField, ";")
                    If Pos = 0 Then Exit Do
                    StrField = Left(StrField, Pos - 1) & Mid(StrField, Pos + 1)
                 Loop
                 
                 StrInd = StrInd & "Fields=" & StrField & conDelim '"Fields=" & Mid$(TheIndex.fields, 2) _
                        '& Mid$(TheIndex.fields, 9, 11) '& Mid$(TheIndex.fields, 17, 19) & conDelim
                 Debug.Print StrInd
                 StrInd = StrInd & "Primary=" & TheIndex.Primary & conDelim
                 StrInd = StrInd & "Unique=" & TheIndex.Unique & conDelim
                 StrInd = StrInd & "IgnoreNulls=" & TheIndex.IgnoreNulls
                 'Now Write to the ini file
                 If Not WriteToIniFile(IndSection, IndKey, StrInd, IniFile) Then Exit Function
                 'Debug.Print StrInd
            Next
    
NextCount:

Next TableCount
'Create Relation between the Source  table and the Foregin Table
'Next Add relation for all the fields
'And after setting the relation write it to the Tab file.
    strSection = "RELATIONS"
    maxRelCount = db.Relations.Count - 1
    
    For RelCount = 0 To maxRelCount
         Set TheRelation = db.Relations(RelCount)
         ' strSection = db.Name
          strKey = TheRelation.Name
          strValue = ""
          strValue = strValue & "RelationName=" & TheRelation.Name & conDelim
          strValue = strValue & "Field=" & TheField.Name & conDelim
          strValue = strValue & "ForiegnTable=" & TheRelation.ForeignTable & conDelim
          strValue = strValue & "SourceTable=" & TheRelation.Table
         'Now Write it into the ini file
          If Not WriteToIniFile(strSection, strKey, strValue, IniFile) Then Exit Function
     Next
         
    'From Here start writing details of the query defs if any
    'DeclareVariables for the queryDef
         QueCount = 0: MaxQueCount = db.QueryDefs.Count - 1
         strSection = "QueryDefs"
         
    For QueCount = 0 To MaxQueCount
        Set TheQuery = db.QueryDefs(QueCount)
           strKey = "Query" & QueCount + 1
           strSection = strDbName & "Query"
           strKey = "QueryDef" & QueCount + 1
           Sqlstr = TheQuery.SQL
           Count = InStr(1, Sqlstr, ";")
           If Count Then Sqlstr = Mid(Sqlstr, Count + 1)
     
         'Theremay be some carriege return charectors
         'Remove those Char's
           Count = 1
           Do
               Count = InStr(Count, Sqlstr, vbCr)
               If Count = 0 Then Exit Do
               Sqlstr = Left(Sqlstr, Count - 1) & " " & Mid(Sqlstr, Count + 1)
               Debug.Print Sqlstr
           Loop
               Count = 1
           Do
               Count = InStr(Count, Sqlstr, vbLf)
               If Count = 0 Then Exit Do
               Sqlstr = Left(Sqlstr, Count - 1) & " " & Mid(Sqlstr, Count + 1)
           Loop
               strValue = "QueryDefName=" & TheQuery.Name & conDelim
               strValue = strValue & "Sql=" & Sqlstr
           If Not WriteToIniFile(strSection, strKey, strValue, IniFile) Then Exit Function
         
      'Parameter checking
         strSection = TheQuery.Name
         strKey = TheParameter.Name
         'Increment the KeyValue(i,e StrKey)
         strKey = "Parameter" & ParaCount + 1
         'strSection = strParaSection
         maxParaCount = TheQuery.Parameters.Count - 1
         
         For ParaCount = 0 To maxParaCount
              Set TheParameter = TheQuery.Parameters(ParaCount)
                   strKey = "Parameter" & ParaCount + 1
                   strValue = "ParameterName=" & TheParameter.Name & conDelim
                   strValue = strValue & "DataType=" & FieldTypeStr(TheParameter.Type)
              If Not WriteToIniFile(strSection, strKey, strValue, IniFile) Then Exit Function
          Next
     Next
'next check for the testing
CreateTabFile = True
End Function

Private Function CreateTBL(db As Database, TblName As String, tblData() As TabStruct) As Boolean
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
Dim I As Integer
Dim PrimaryKeySet As Boolean
'On Error GoTo err_line

' Create the table.
'RaiseEvent UpdateStatus("Creating the table " & TblName & "...")
Set TheTable = db.CreateTableDef(TblName)


' Create and add the fields.
Dim TheField As Field
With TheTable
    For I = 0 To UBound(tblData)
        ' Create and append the field.
        Set TheField = .CreateField(tblData(I).Field, _
                        tblData(I).Type, tblData(I).Length)
        If tblData(I).Type = dbText Or tblData(I).Type = dbMemo Then
            TheField.AllowZeroLength = IIf(tblData(I).Required, False, True)
        End If
            
        If tblData(I).AutoIncrement Then TheField.Attributes = dbAutoIncrField
           
        TheField.Required = tblData(I).Required
        .fields.Append TheField
        'TheField.ForeignName
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

db.TableDefs.Append TheTable
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
    Case "DATE", "DBDATE", "DATETIME"
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
        FieldTypeStr = "DbText"
    Case dbInteger
        FieldTypeStr = "DbInteger"
    Case dbDate
        FieldTypeStr = "DateTime"
    Case dbBoolean
        FieldTypeStr = "DbBoolean"
    Case dbByte
        FieldTypeStr = "DbByte"
    Case dbChar
        FieldTypeStr = "DbChar"
    Case dbCurrency
        FieldTypeStr = "DbCurrency"
    Case dbDecimal
        FieldTypeStr = "DbDecimal"
    Case dbDouble
        FieldTypeStr = "DbDouble"
    Case dbFloat
        FieldTypeStr = "DbFloat"
    Case dbLong
        FieldTypeStr = "DbLong"
    Case dbLongBinary
        FieldTypeStr = "DbLongBinary"
    Case dbMemo
        FieldTypeStr = "DbMemo"
    Case dbSingle
        FieldTypeStr = "DbSingle"
    Case Else   ' Unrecognized fldtype.
        FieldTypeStr = ""
End Select
End Function
Private Function GetFieldCount(strdataFile As String, Optional TblName As String) As Integer
Dim strRet As String
Dim FieldCount As Integer
Dim TableName As String
Dim DBName As String
Dim I As Integer, j As Integer, K As Integer
Dim FieldName As String
Do
    I = I + 1
    DBName = ReadFromIniFile("DataBases", "DataBase" & I, strdataFile)
    If DBName = "" Then Exit Do
    j = 0
    Do
        j = j + 1
        TableName = ReadFromIniFile(DBName, "Table" & j, strdataFile)
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
Dim I As Integer, j As Integer, K As Integer
    Do
        I = I + 1
        DBName = ReadFromIniFile("DataBases", "Database" & I, strdataFile)
        If DBName = "" Then Exit Do
        j = 0
        Do
            j = j + 1
            strRet = ReadFromIniFile(DBName, "Table" & j, strdataFile)
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
Public Function OpenDB(strDbName As String, pwd As String) As Boolean

On Error GoTo ErrLine
If strDbName = "" Then
    Exit Function
End If

Dim Count As Integer
'  Debug.Assert strDbName = ""

'If there is new Password then ask for the user

'If pwd = "WIS!@#" Then
'   pwd = InputBox(vbTab & "  Please Confirm Your Password." _
'                 & vbNewLine & vbCrLf & "If Password has been Change then enter " & _
'                   "your new Password.", "Confirm Your Password ", "WIS!@#")
'End If

Set m_DataBaseObject = OpenDatabase(strDbName, False, False, ";pwd=" & pwd)
OpenDB = True
m_DBOpen = True
M_dbName = strDbName
M_DbPwd = pwd
Do
    Count = InStr(Count + 1, strDbName, "\", vbBinaryCompare)
    If Count = 0 Then Exit Do
    M_DbPath = Left(strDbName, Count)
Loop
Exit Function
ErrLine:
    If Err Then
        MsgBox "OpenDB: The following error occured, " _
                & "while opening the database." _
                & vbCrLf & Err.Description, vbCritical
    End If
    Exit Function
End Function
Public Function PrintToHtml() As Boolean
'Initilize the Variables

End Function

'
Public Function QueryResult(StrTabName As String) As Boolean
'Initilize the variables
Dim Sqlstr As String
Dim db As Database
Dim rstData As Recordset
Dim TheTable As TableDef
Dim DualTabName As String
Dim TableCount As Integer
Dim maxTableCount As Integer
Dim Pos As Integer

Set db = m_DataBaseObject

Pos = 0
 Do
      Pos = InStr(1, StrTabName, vbCrLf)
      If Pos = 0 Then Exit Do
      StrTabName = Left$(StrTabName, Pos - 1)
      Debug.Print StrTabName
 Loop

DualTabName = StrTabName

Sqlstr = DualTabName

gDBTrans.SQLStmt = Sqlstr
gDBTrans.BeginTrans

If gDBTrans.SQLFetch < 0 Then Exit Function

If Not gDBTrans.SQLExecute Then
    gDBTrans.RollBack
    Exit Function
End If

gDBTrans.CommitTrans



QueryResult = True

End Function
Public Function RollBack() As Boolean
    
    RollBack = False
    If m_PrevOp = op_BEGINTRANS Then
        m_PrevOp = op_ROLLBACK
        m_SQLLog = ""
        Workspaces(m_TransCount - 1).RollBack
        m_TransCount = m_TransCount - 1
        RollBack = True
    ElseIf MultiTrans Then
        Workspaces(m_TransCount - 1).RollBack
        m_TransCount = m_TransCount - 1
        RollBack = True
    End If

End Function

Public Function BeginTrans() As Boolean
    BeginTrans = False
    If m_PrevOp <> op_BEGINTRANS Then
        Workspaces(m_TransCount).BeginTrans
        m_TransCount = m_TransCount + 1
        m_PrevOp = op_BEGINTRANS
        BeginTrans = True
    ElseIf m_PrevOp = op_BEGINTRANS And MultiTrans Then
        m_TransCount = m_TransCount + 1
        Workspaces(m_TransCount).BeginTrans
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

Debug.Assert SQLStmt <> ""

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


ExeLine:
'Execute the Action Query.
m_DataBaseObject.Execute SQLStmt$, _
        dbFailOnError Or dbSeeChanges
m_SQLLog = m_SQLLog & vbCrLf & SQLStmt
Dim FIleNo As Integer
FIleNo = FreeFile(1)
'''get the file of the today date

Open m_SQLFile For Append Access Write Lock Write As FIleNo
'''Open "TESTFILE" For Binary Access Read Lock Read As #1
Write #FIleNo, SQLStmt & vbCrLf
Close FIleNo
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
        Err.Clear                           ''Execution of program will be faster with AnsiNums--siddu
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
        MsgBox "SQLExecute: The following error occurred." _
            & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
        m_errNum = Err.Number
        m_errDesc = Err.Description
    End If

End Function
Public Function SQLFetch() As Long

Debug.Assert SQLStmt <> ""
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
    On Error Resume Next  'temporaly added for continue process
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





