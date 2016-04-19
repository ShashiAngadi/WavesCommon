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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'---------------------------------------------------------
'   Module Name       :   Transaction Class:
'   Version           :   1.0
'   Created on        :   22nd January 2002
'   Last modified     :   22nd January 2002
'   Last modified  by :   Lingappa Sindhanur
'---------------------------------------------------------

Const op_COMMIT = 1
Const op_ROLLBACK = 2
Const op_BEGINTRANS = 3

Public m_DataBaseObject As ADODB.Connection
Private m_Catalog As New ADOX.Catalog
Public CommandObject As ADODB.Command

Private m_DBOpen As Boolean
Private m_PrevOp As Integer
Private m_SQLStmt As String



 'This function will delete the specified view
' returns true after successful deletion
'Inputs : ViewName as string
Public Function DeleteView(ByVal ViewName As String) As Boolean
'Trap the error
On Error GoTo ErrLine

'Declare the variables
Dim ViewCount As Integer
Dim lpCount As Integer
Dim Sqlstr As String
Dim isViewFound As Boolean

'initalise the functions
DeleteView = False

Set m_Catalog = New ADOX.Catalog

m_Catalog.ActiveConnection = m_DataBaseObject

ViewCount = m_Catalog.Views.Count - 1

isViewFound = False

For lpCount = ViewCount To 0 Step -1
    If UCase$(m_Catalog.Views.Item(lpCount).Name) = UCase$(ViewName) Then
        
        'm_Catalog.Views.Delete (lpCount)
        
        Sqlstr = m_SQLStmt
        
        Call DropView(ViewName)
        
        m_SQLStmt = Sqlstr
        
        isViewFound = True
        Exit For
   End If
   
Next lpCount

If Not isViewFound Then If Not DeleteProcedures(ViewName) Then Exit Function

DeleteView = True

Exit Function

ErrLine:
   MsgBox "DeleteView: " & vbCrLf & Err.Description, vbCritical
   

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
Dim Sqlstr As String


'initalise the functions
DeleteProcedures = False

Set m_Catalog = New ADOX.Catalog

m_Catalog.ActiveConnection = m_DataBaseObject

ViewCount = m_Catalog.Procedures.Count - 1

For lpCount = ViewCount To 0 Step -1
    If UCase$(m_Catalog.Procedures.Item(lpCount).Name) = UCase$(ViewName) Then
        
        Sqlstr = m_SQLStmt
        
        Call DropView(ViewName)
        
        m_SQLStmt = Sqlstr
        
        Exit For
   End If
   
Next lpCount

DeleteProcedures = True

Exit Function

ErrLine:
   MsgBox "DeleteProcedures: " & vbCrLf & Err.Description, vbCritical
      

End Function

Public Sub ReturnActiveConnection(Connection As ADODB.Connection)

On Error GoTo Hell:

Set Connection = New ADODB.Connection

Connection = m_DataBaseObject

Exit Sub

Hell:

End Sub

Public Function CheckBeginTrans() As Boolean
'initialise the function
CheckBeginTrans = False

'Check if the transaction has started
If m_PrevOp <> op_BEGINTRANS Then
   MsgBox "CheckBeginTrans:" & vbCrLf & " Transaction not started!", vbExclamation, " - Critical Error"
   CheckBeginTrans = False
   Exit Function
End If

CheckBeginTrans = True

End Function

Public Function CreateView(ViewName As String) As Boolean

'Declare the Variables
Dim ViewCount As Integer

'Trap the error
On Error GoTo ErrLine

CreateView = False

'If no SQL statement set, exit.
If m_SQLStmt = "" Then
    MsgBox "No SQL statement has been set!", vbExclamation
    Exit Function
End If

' First delete the Qry temp if exists
DeleteView (ViewName)

Set gDBTrans.CommandObject = New ADODB.Command

gDBTrans.CommandObject.CommandText = m_SQLStmt

m_Catalog.ActiveConnection = m_DataBaseObject

m_Catalog.Views.Append ViewName, CommandObject

ViewCount = m_Catalog.Views.Count

CreateView = ViewCount

'Set the sqlstmt to null
m_SQLStmt = ""
CreateView = True

Exit Function

ErrLine:
   MsgBox "CreateView: " & vbCrLf & Err.Description, vbCritical
      
End Function


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


Public Function DeleteAllViews() As Boolean

'Declare the Variables
Dim lpCount As Integer
Dim ViewCount As Integer
'Trap the error
On Error GoTo ErrLine

DeleteAllViews = False

m_Catalog.ActiveConnection = m_DataBaseObject

ViewCount = m_Catalog.Views.Count - 1

For lpCount = ViewCount To 0 Step -1
   m_Catalog.Views.Delete (lpCount)
Next lpCount

DeleteAllViews = True

Exit Function

ErrLine:
    MsgBox "DeleteAllViews: " & vbCrLf & Err.Description, vbCritical
   

End Function

Private Sub DropView(ViewName As String)

'Trap the error
On Error GoTo ErrLine

m_SQLStmt = " DROP TABLE " & ViewName

gDBTrans.BeginTrans

If Not gDBTrans.SQLExecute Then gDBTrans.RollBacknRaiseError
    
gDBTrans.CommitTrans
        
ErrLine:
    

End Sub


'==============================================================================
' This function searches the ADO recordset for the given value.
'
' Paramaeters are
'  1.  rstToFind  : adodb.recordset
'  2.  SearchString   : search string   ex : "BankID=" & bankid
'
' searches the record if found returns true else false
'
' - Pradeep , 31/01/2002
'==============================================================================
Public Function FindRecord(ByRef rstToFind As ADODB.Recordset, ByVal SearchString As String) As Boolean

' in the given string
On Error GoTo Hell

Dim SearchArray() As String
Dim ItemCount As Integer
Const SEARCH_STRING_NO = 1
' initialise the function to false
FindRecord = False

' if any filter is there then remove the filter

rstToFind.Filter = adFilterNone

'"BankID=" & 10000500 & "TermType=" & 12 & "LoanPurpose=" & 1

Call GetStringArray(SearchString, SearchArray, ",")
'in the VB6 U can use SPLIT function  which will be faster than GetstringArray

'Comment on filtereation
For ItemCount = LBound(SearchArray) To UBound(SearchArray) - SEARCH_STRING_NO    ' WHY -1 ?
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
If Not rstToFind.EOF Then ' if found it will be in the found record
    FindRecord = True
Else
    FindRecord = False ' else it will on the EOF
End If

'If isFilter Then rstToFind.Filter = adFilterNone
Exit Function

Hell:
    MsgBox "Field Specified or Value given is wrong !!", vbCritical, wis_MESSAGE_TITLE
    
End Function

Public Function ParameterizedQuery(QryName As String) As Boolean
On Error GoTo ErrLine

ParameterizedQuery = False

' Open the catalog
m_Catalog.ActiveConnection = m_DataBaseObject

' Get the Command object from the
' Procedure
Set CommandObject = m_Catalog.Procedures(QryName).Command
   
ParameterizedQuery = True

Exit Function

ErrLine:
      MsgBox "ParameterizedQuery: " & vbCrLf & Err.Description, vbCritical
   
   
End Function

'This function will seeek the record from the specified recordset with specified criteria
'Input : Recordset Object
'        seek Criteria as string
' Returns true if it finds the record
' Retuns false otherwise
' And if seek values are more than one column
' SeekCriteria string should be delimited by , (Comma)
Public Function SeekRecord(rstSeeK As ADODB.Recordset, SeekCriteria As String) As Boolean


On Error GoTo ErrLine
Dim I As Integer

If Not rstSeeK.Supports(adIndex) Then Exit Function

I = InStr(1, SeekCriteria, ",", vbTextCompare)

rstSeeK.Index = "PrimaryKey"

If I > 0 Then
   rstSeeK.Seek Array(SeekCriteria), adSeekFirstEQ
Else
   rstSeeK.Seek SeekCriteria, adSeekFirstEQ
End If

SeekRecord = False

If Not rstSeeK.EOF Then SeekRecord = True


Exit Function

ErrLine:
   MsgBox "SeekRecord: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
   
End Function

Public Function SQLExecute() As Boolean

On Error GoTo ErrLine

Dim TheRecordsAffected As Long


SQLExecute = False

' Check if the private variable m_sqlstmt is set.
If Trim$(m_SQLStmt) = "" Then
   MsgBox "No SQL statement is set!", vbExclamation
   Exit Function
End If

'Check if the transaction has started
If m_PrevOp <> op_BEGINTRANS Then
   MsgBox "Transaction not started !", vbExclamation, " - Critical Error"
   Exit Function
End If

'Execute the Action Query.
m_DataBaseObject.Execute m_SQLStmt, TheRecordsAffected, adExecuteNoRecords
   
SQLExecute = True

'Reset sqlstmt to avoid misuse
m_SQLStmt = ""

Exit Function

ErrLine:
        MsgBox "SQLExecute: The following error occurred." _
                & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
        
        m_SQLStmt = ""
        SQLExecute = False
    
End Function





'This Function will check the transaction status
'if transaction is made if u r exiting from the application
' It will prompt u the transaction to be commited or rollbacked
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

'
Public Function BeginTrans() As Boolean
On Error GoTo ErrLine

BeginTrans = False

If m_PrevOp <> op_BEGINTRANS Then
   m_PrevOp = op_BEGINTRANS
   m_DataBaseObject.BeginTrans
   BeginTrans = True
End If

Exit Function

ErrLine:
   MsgBox "BeginTrans: " & Err.Description, vbCritical
   
End Function

'
Public Function CommitTrans() As Boolean
CommitTrans = False

If m_PrevOp = op_BEGINTRANS Then
    m_PrevOp = op_COMMIT
    m_DataBaseObject.CommitTrans
    CommitTrans = True
End If

End Function

'
Public Function RollBack() As Boolean
    
    RollBack = False
    If m_PrevOp = op_BEGINTRANS Then
        m_PrevOp = op_ROLLBACK
        m_DataBaseObject.RollbackTrans
        RollBack = True
    End If

End Function
'
Public Function RollBacknRaiseError() As Boolean
    
    RollBacknRaiseError = False
    If m_PrevOp = op_BEGINTRANS Then
        m_PrevOp = op_ROLLBACK
        m_DataBaseObject.RollbackTrans
        RollBacknRaiseError = True
    End If
        
    Err.Raise vbObjectError + 513, , "Database Error!"
    
End Function


' This function Fetches the data from the database assigns the
' the resulted records into the recordset you have passed
' Input : recordset object
'  Returns the number of records fetched from the database
'  Created on : 22nd Jan 2002
'  Created By : Lingappa Sindhanur
Public Function Fetch(rstFetch As ADODB.Recordset, TheCursorType As CursorTypeEnum) As Long

'If no SQL statement set, exit.
If m_SQLStmt = "" Then
    MsgBox "No SQL statement has been set!", vbExclamation
    Exit Function
End If

' If the datbase object is not initialized, exit.
If m_DataBaseObject Is Nothing Then
    MsgBox "Fetch: No Connection is in use.  " _
        & "Set the database name before calling Fetch.", vbExclamation
    Exit Function
End If

On Error GoTo ErrLine
Fetch = 0

Set rstFetch = New ADODB.Recordset

rstFetch.CursorLocation = adUseClient

rstFetch.Open m_SQLStmt, m_DataBaseObject, TheCursorType, adLockReadOnly, adCmdText

Select Case TheCursorType
   Case adOpenStatic
      Fetch = rstFetch.RecordCount
   Case Else
      Fetch = 1
      If rstFetch.BOF And rstFetch.EOF Then Fetch = 0
End Select

' Disconnect the Recordset
Set rstFetch.ActiveConnection = Nothing

'Set the sqlstmt to null
m_SQLStmt = ""

Exit Function

ErrLine:
      Fetch = 0
      MsgBox "SQLFetch: The following error occurred " _
            & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
          
    
End Function
' This function Fetches the data from the database assigns the
' the resulted records into the recordset you have passed
' Input : recordset object
' returns the number of records fetched from the database
'  Created on : 22nd Jan 2002
'  Created By : Lingappa Sindhanur
Public Function ParameterFetch(rstFetch As ADODB.Recordset, TheCursorType As CursorTypeEnum) As Long

'If no Command object set, exit.
If CommandObject Is Nothing Then
   MsgBox "Fetch: No command object has been set!", vbExclamation
    Exit Function
End If

' If the datbase object is not initialized, exit.
If m_DataBaseObject Is Nothing Then
    MsgBox "Fetch: No Connection is in use.  " _
        & "Set the database name before calling Fetch.", vbExclamation
    Exit Function
End If

On Error GoTo ErrLine
ParameterFetch = 0

Set rstFetch = New ADODB.Recordset

rstFetch.Open CommandObject, , TheCursorType, adLockOptimistic, adCmdStoredProc

Select Case TheCursorType
   Case adOpenStatic
      ParameterFetch = rstFetch.RecordCount
   Case Else
      ParameterFetch = 1
      If rstFetch.BOF And rstFetch.EOF Then ParameterFetch = -1
End Select

'Set the sqlstmt to null
m_SQLStmt = ""

Exit Function

ErrLine:
      ParameterFetch = -1
      MsgBox "SQLFetch: The following error occurred " _
            & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
      
   
End Function

' This function Opens the connection with database
' returns true after successful connection and
' returns false after failure
' Input       : Database name as string
'              : password if any
'  Created on : 22nd Jan 2002
'  Created By : Lingappa Sindhanur

Public Function OpenDB(strDbName As String, pwd As String) As Boolean

If strDbName = "" Then Exit Function

If m_DataBaseObject Is Nothing Then Set m_DataBaseObject = New ADODB.Connection

m_DataBaseObject.Mode = adModeShareDenyNone
   
On Error GoTo ErrLine
'Open a connection Using Microsoft.Jet Provider
m_DataBaseObject.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & strDbName & ";" & _
                       "Jet OLEDB:Database Password=" & pwd & ";"

OpenDB = True
m_DBOpen = True

Exit Function

ErrLine:
      MsgBox "OpenDB: The following error occured, " _
                & "while opening the database." _
                & vbCrLf & Err.Description, vbCritical
      
    
End Function




'
Public Property Get SQLStmt() As String
SQLStmt = m_SQLStmt
End Property

'This function will create the view
' returns true after successful creation
' inputs  : ViewName as string this the name of the view to be created
Public Function CreateProcedures(ViewName As String) As Boolean
'Declare the Variables
Dim ViewCount As Integer

'Trap the error
On Error GoTo ErrLine

CreateProcedures = False

'If no SQL statement set, exit.
If m_SQLStmt = "" Then
    MsgBox "No SQL statement has been set!", vbExclamation
    Exit Function
End If

' First delete the Qry temp if exists
Call gDBTrans.DeleteProcedures(ViewName)

Set gDBTrans.CommandObject = New ADODB.Command

gDBTrans.CommandObject.CommandText = m_SQLStmt

m_Catalog.ActiveConnection = m_DataBaseObject

m_Catalog.Procedures.Append ViewName, CommandObject

ViewCount = m_Catalog.Views.Count

'Set the sqlstmt to null
m_SQLStmt = ""

CreateProcedures = True

Exit Function

ErrLine:
    MsgBox "CreateProcedures: " & vbCrLf & Err.Description, vbCritical
      
End Function

'
Public Property Let SQLStmt(ByVal vNewValue As String)
m_SQLStmt = vNewValue
End Property



