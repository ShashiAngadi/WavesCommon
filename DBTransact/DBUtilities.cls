VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 
Option Explicit

Private m_JROEngine As JRO.JetEngine
Private m_SourceDB As ADOX.Catalog
Private m_TargetDB As ADOX.Catalog


'
Private Function CheckFieldStructure(SourceField As ADOX.Column, TargetField As ADOX.Column) As Boolean

'Setup an error handler...
On Error GoTo ErrLine


With SourceField
    On Error Resume Next
    If .Type <> TargetField.Type Then TargetField.Type = .Type
    
    If Err Then Exit Function
    
    If .DefinedSize <> TargetField.DefinedSize Then TargetField.DefinedSize = .DefinedSize
    If .Attributes <> TargetField.Attributes Then TargetField.Attributes = .Attributes
    If .Properties("Jet OLEDB:Allow Zero Length") <> TargetField.Properties("Jet OLEDB:Allow Zero Length") Then TargetField.Properties("Jet OLEDB:Allow Zero Length") = .Properties("Jet OLEDB:Allow Zero Length")
    If .Properties("Default") <> TargetField.Properties("Default") Then TargetField.Properties("Default") = .Properties("Default")
    If .Properties("Description") <> TargetField.Properties("Description") Then TargetField.Properties("Description") = .Properties("Description")
    If .Properties("Nullable") <> TargetField.Properties("Nullable") Then TargetField.Properties("Nullable") = .Properties("Nullable")
End With

CheckFieldStructure = True

Exit Function

ErrLine:
    CheckFieldStructure = False
    
    

End Function

Private Function CheckFieldsEqual(SourceField As ADOX.Column, TargetTable As ADOX.Table) As Boolean
'Declare the variables
Dim FieldCount As Integer
Dim Item As Integer
Dim SourceFieldName As String

If SourceField Is Nothing Then Exit Function

FieldCount = TargetTable.Columns.Count

SourceFieldName = SourceField.Name

For Item = 0 To FieldCount - 1
    If TargetTable.Columns(Item).Name = SourceFieldName Then
        CheckFieldsEqual = False
        If CheckFieldStructure(SourceField, TargetTable.Columns(Item)) Then CheckFieldsEqual = True
        Exit Function
    End If
Next Item


End Function

'This file will return true if file is already open
Public Function IsFileAlreadyOpen(ByVal Filename As String) As Boolean
Dim hFile As Long
Dim lastErr As Long

'Initialize file handle and error variable.
hFile = -1
lastErr = 0

'Open for for read and exclusive sharing.
hFile = lopen(Filename, &H10)

' If we couldn't open the file, get the last error.
If hFile = -1 Then
   lastErr = Err.LastDllError
' Make sure we close the file on success.
Else
   lclose (hFile)
End If

' Check for sharing violation error.
If (hFile = -1) And (lastErr = 32) Then
   IsFileAlreadyOpen = True
Else
   IsFileAlreadyOpen = False
End If
End Function
'This function will make the back up of the given file name
Public Function MakeBackUp(ByVal DataFileName As String) As Boolean
'Declare the variables
Dim BaseFolder As String
Dim PresentDate As String
Dim FolderName As String
Dim PrevFolder As String
Dim isFileOpen As Boolean

PresentDate = Format(Now, "DD/YY/MM")
BaseFolder = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "BaseFolder")
PrevFolder = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "DbFolderName")

If Trim$(BaseFolder) = "" Then BaseFolder = App.Path
If Right$(BaseFolder, 1) = "\" Then BaseFolder = Left$(BaseFolder, Len(BaseFolder) - 1)

FolderName = BaseFolder + "\" & constBAKFILEPREFIX & Left$(PresentDate, 2) & Right$(PresentDate, 2)

If PrevFolder = FolderName Then Exit Function

#If TRIALVERSION Then
    If GetTrialPeriod >= 30 Then
        MsgBox "Your Trial Period Expired" & vbCrLf & "For Details Contact Waves Information Systems Gadag", vbInformation, wis_MESSAGE_TITLE
        ShutDownInventory
    End If
#End If

On Error GoTo ErrLine

' check for the existence of path
If Dir(FolderName, vbDirectory) = "" Then MakeDirectories (FolderName)

'Close the database
gDBTrans.CloseDB


isFileOpen = False

isFileOpen = IsFileAlreadyOpen(DataFileName)
    
If Not isFileOpen Then
    FileCopy DataFileName, FolderName + "\" & constDBName
    
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "DbFolderName", FolderName)
    
    'Delete the previous folder of the backupcopy
    On Error Resume Next
    
    If FolderName <> PrevFolder Then
        If Dir(PrevFolder, vbDirectory) <> "" Then
            Kill PrevFolder & "\" & constDBName
            RmDir PrevFolder
        End If
    End If

End If

If Not gDBTrans.OpenDB(DataFileName, constDBPWD) Then
    MsgBox "Unable to Open the Database"
    End
End If

MakeBackUp = True

Exit Function
 
ErrLine:
    MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "Unable to take back up", vbInformation
End Function

'This function will compact the database
Public Function CompactTheDataBase(ByVal strDbName As String, ByVal pwd As String) As Boolean
    
On Error GoTo ErrLine

CompactTheDataBase = False

'Validate the inputs
If strDbName = "" Then Exit Function


If Dir(App.Path & "\NewMaterial.mdb") <> "" Then Kill App.Path & "\NewMaterial.mdb"

If m_JROEngine Is Nothing Then Set m_JROEngine = New JRO.JetEngine

gDBTrans.CloseDB

' Compact the database specifying the new database password
m_JROEngine.CompactDataBase "Data Source = " & strDbName & ";" & _
                            "Jet OLEDB:Database Password=" & pwd & ";", _
                            "Data Source = " & App.Path & "\NewMaterial.mdb" & ";" & _
                            "Jet OLEDB:Database Password=" & pwd & ";"


' Delete the original database
Kill strDbName

'Rename the file back to the original name
Name App.Path & "\NewMaterial.mdb" As strDbName

Call gDBTrans.OpenDB(strDbName, constDBPWD)

CompactTheDataBase = True

Exit Function

ErrLine:
    MsgBox "CompactTheDataBase:" & vbCrLf & Err.Description
    
End Function

'
'This will create the new Database using the old database
'Dependencies : Previously opened database is required
'new database will be created using the structure of the old database
' Inputs : DataFile name as string
'           StrPwd as string password
'
Public Function CreateDBFromDB(DbFile As String, strPwd As String) As Boolean
On Error GoTo Hell:

Dim SourceDB As ADOX.Catalog
Dim TargetDB As ADOX.Catalog
Dim TheTable As ADOX.Table

Dim TotalTable As Integer
Dim TableCount As Integer

Dim SourceConnection As ADODB.Connection

'validate the inputs
If DbFile = "" Then Exit Function

'Check if the existance of the directory path
If Not ValidateDBPath(DbFile) Then Exit Function

Set SourceDB = New ADOX.Catalog
Set TargetDB = New ADOX.Catalog

Set SourceConnection = New ADODB.Connection

'Get the opened connection
Call gDBTrans.ReturnActiveConnection(SourceConnection)

SourceDB.ActiveConnection = gDBTrans.m_DataBaseObject

If Trim$(strPwd) = "" Then
   TargetDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                     "Data Source=" & DbFile & ";"
Else
   TargetDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                     "Data Source=" & DbFile & ";" & _
                     "Jet OLEDB:Database Password=" & strPwd & ";"
End If

TotalTable = SourceDB.Tables.Count

For TableCount = 0 To TotalTable - 1
    Set TheTable = SourceDB.Tables(TableCount)
    If TheTable.Type = "TABLE" Then Call CreateTableFromDB(TheTable, TargetDB)
Next TableCount

CreateDBFromDB = True

Exit Function
Hell:

End Function
'
'This will create the new Database using the old database
'Dependencies : Previously opened database is required
'new database will be created using the structure of the old database
' Inputs : DataFile name as string
'           StrPwd as string password
' NOTE : SOURCEDB IS THE BLANK DB WHCIH IS PASSED
'        TARGETDB IS CURRENT DB WHHICH IS OPEN
Public Function CompareDBFromDB(ByVal BlankDbFile As String, ByVal strPwd As String) As Boolean

On Error GoTo Hell:

Dim TheTable As ADOX.Table
Dim TotalTable As Integer
Dim TableCount As Integer

Dim SourceConnection As ADODB.Connection
Dim TargetConnection As ADODB.Connection

'validate the inputs
If BlankDbFile = "" Then Exit Function

Set m_SourceDB = New ADOX.Catalog
Set m_TargetDB = New ADOX.Catalog
Set SourceConnection = New ADODB.Connection
Set TargetConnection = New ADODB.Connection

m_TargetDB.ActiveConnection = gDBTrans.m_DataBaseObject

If Trim$(strPwd) = "" Then
   SourceConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                     "Data Source=" & BlankDbFile & ";"
Else
   SourceConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                     "Data Source=" & BlankDbFile & ";" & _
                     "Jet OLEDB:Database Password=" & strPwd & ";"
End If

m_SourceDB.ActiveConnection = SourceConnection

TotalTable = m_SourceDB.Tables.Count

For TableCount = 0 To TotalTable - 1

    Set TheTable = m_SourceDB.Tables(TableCount)
    
    If TheTable.Type = "TABLE" Then
        If TableExists(TheTable.Name, m_TargetDB) Then
            If Not UpdateTableField(TheTable, m_TargetDB.Tables(TheTable.Name)) Then Exit Function
        Else
            If Not CreateTableFromDB(TheTable, m_TargetDB) Then Exit Function
        End If
    End If
    
Next TableCount

CompareDBFromDB = True

Exit Function

Hell:
    MsgBox Err.Description, vbCritical
    
End Function
'This will create the table structere from the old database
Private Function CreateTableFromDB(TheTable As ADOX.Table, TargetDB As ADOX.Catalog) As Boolean
Dim TheField As ADOX.Column
Dim tblData() As ADOX.Column
Dim strTblName As String
Dim ReqFld As Boolean
Dim TotalField As Long
Dim K As Long
Dim j As Byte

j = 1
'Get the Table name
strTblName = TheTable.Name

K = 0

ReDim tblData(0)

TotalField = TheTable.Columns.Count

Set tblData(K) = New ADOX.Column

Do

   If K = TotalField Then Exit Do

    'Add to fields array.
   ReDim Preserve tblData(K)

   Set tblData(K) = New ADOX.Column

   Set TheField = New ADOX.Column

   TheField = TheTable.Columns(K)

   With tblData(K)
    .Name = TheTable.Columns(K).Name
    .Type = TheTable.Columns(K).Type
    .DefinedSize = TheTable.Columns(K).DefinedSize
    
    Set .ParentCatalog = TargetDB
    
    ReqFld = False
    If TheTable.Columns(K).Properties("Nullable") = True Then ReqFld = True

    If ReqFld Then
        .Attributes = adColFixed
        .Properties("Nullable") = False
        .Properties("Jet OLEDB:Allow Zero Length") = False
    Else
        .Attributes = adColNullable
        .Properties("Nullable") = True
        .Properties("Jet OLEDB:Allow Zero Length") = True
    End If
    .Properties("Description") = TheTable.Columns(K).Properties("Description")
    .Properties("Default") = TheTable.Columns(K).Properties("Default")
   End With

  '  Increment the field count variable "k"
   K = K + 1
Loop

 'Create the table.
If Not CreateTBL(TargetDB, strTblName, tblData) Then
    MsgBox "Error in creating Table", vbCritical
    Exit Function
End If
   
If Not CreateTablIndx(TheTable, TargetDB) Then Exit Function

CreateTableFromDB = True

Exit Function

ErrLine:
    MsgBox "CreateTableFromDB: " & Err.Description
    

End Function
'This will create the table structere from the old database
Private Function CreateTableFromDB_NEW(TheTable As ADOX.Table, TargetDB As ADOX.Catalog) As Boolean
Dim TheField As ADOX.Column
Dim tblData() As ADOX.Column
Dim strTblName As String
Dim ReqFld As Boolean
Dim TotalField As Long
Dim K As Long
Dim j As Byte

j = 1
'Get the Table name
strTblName = TheTable.Name

K = 0

ReDim tblData(0)

TotalField = TheTable.Columns.Count

Set tblData(K) = New ADOX.Column

Do

   If K = TotalField Then Exit Do

    'Add to fields array.
   ReDim Preserve tblData(K)

   Set tblData(K) = New ADOX.Column

   Set TheField = New ADOX.Column

   TheField = TheTable.Columns(K)

   With tblData(K)
    .Name = TheTable.Columns(K).Name
    .Type = TheTable.Columns(K).Type
    .DefinedSize = TheTable.Columns(K).DefinedSize
    
    ReqFld = False
    If TheTable.Columns(K).Properties("Nullable") = True Then ReqFld = True

    Set .ParentCatalog = TargetDB

    If ReqFld Then
        .Attributes = adColFixed
        .Properties("Nullable") = False
        .Properties("Jet OLEDB:Allow Zero Length") = False
    Else
        .Attributes = adColNullable
        .Properties("Nullable") = True
        .Properties("Jet OLEDB:Allow Zero Length") = True
    End If
    .Properties("Description") = TheTable.Columns(K).Properties("Description")
    .Properties("Default") = TheTable.Columns(K).Properties("Default")
   End With

  '  Increment the field count variable "k"
   K = K + 1
Loop

 'Create the table.
If Not CreateTBL(TargetDB, strTblName, tblData) Then
    MsgBox "Error in creating Table", vbCritical
    Exit Function
End If
   
Call CreateTablIndx(TheTable, TargetDB)

End Function

'This will compare the table structere from the old database
Private Function CompareTableFromDB(SourceTable As ADOX.Table, TargetTable As ADOX.Table) As Boolean
Dim SourceField As ADOX.Column
Dim TargetField As ADOX.Column

Dim tblData() As ADOX.Column
Dim strTblName As String
Dim ReqFld As Boolean

Dim TargetFieldCount As Integer
Dim SourceFieldCount As Integer
Dim K As Long
Dim j As Byte

j = 1
'Get the Table name
strTblName = SourceTable.Name

K = 0

ReDim tblData(0)

SourceFieldCount = SourceTable.Columns.Count
TargetFieldCount = TargetTable.Columns.Count

Set tblData(K) = New ADOX.Column

End Function
'
Private Function CreateTablIndx(TheTable As ADOX.Table, TargetDB As ADOX.Catalog) As Boolean
' If any indexes are specified, create them.
Dim IndxData As ADOX.Index
Dim IndxTotal As Integer
Dim IndxFlds As String
Dim IndxCols As Integer
Dim K As Integer

'Setup an error handler...
On Error GoTo ErrLine

K = 0

IndxTotal = TheTable.Indexes.Count

Do
   
   If K = IndxTotal Then Exit Do
   
   Set IndxData = New ADOX.Index
   
   For IndxCols = 0 To TheTable.Indexes(K).Columns.Count - 1
      If IndxCols = 0 Then
         IndxFlds = TheTable.Indexes(K).Columns(IndxCols).Name
      Else
          IndxFlds = IndxFlds & "+" & TheTable.Indexes(K).Columns(IndxCols).Name
      End If
      
   Next IndxCols
   
   With IndxData
       .Name = TheTable.Indexes(K).Name
       
       .PrimaryKey = TheTable.Indexes(K).PrimaryKey
       
       .Unique = TheTable.Indexes(K).Unique
       .IndexNulls = TheTable.Indexes(K).IndexNulls
   End With

    If IndxTotal > 0 Then
       If Not CreateIndex(TargetDB, TheTable.Name, IndxData, IndxFlds) Then
          MsgBox "CreateIndexes error:" & Err.Description
          Exit Function
       End If
    End If
   
   
   K = K + 1
Loop

CreateTablIndx = True

Exit Function

ErrLine:
    MsgBox Err.Description
    
      
End Function
Private Function CreateIndex(Cat As ADOX.Catalog, tbl As String, idx As ADOX.Index, idxFlds As String) As Boolean

On Error GoTo CreateIndex_Error

Dim TheTable As ADOX.Table
Dim TheIndex As ADOX.Index
Dim fldArr() As String
Dim I As Integer, j As Integer

Set TheTable = Cat.Tables(tbl)
Set TheIndex = New ADOX.Index
I = 0

With TheTable
    'Set TheIndex = New ADOX.Index
    TheIndex.Name = idx.Name
    ' Breakup the fields string to get the field names for index.
    GetStringArray idxFlds, fldArr(), "+"
    
    With TheIndex
        For j = 0 To UBound(fldArr)
           .Columns.Append fldArr(j)
        Next j
        .PrimaryKey = idx.PrimaryKey
        .IndexNulls = idx.IndexNulls
           '.Unique = idx.Unique
    End With
    .Indexes.Append TheIndex
End With

CreateIndex = True

Exit Function

CreateIndex_Error:
    If Err.Number = 3283 Then   ' Primary key already exists.
        Resume Next
    ElseIf Err Then
        MsgBox Err.Description, vbCritical
    End If
End Function

' Creates a Database file (.mdb) by taking the necessary data
' from a given INI file, with the given password.
'
' Returns:
' True    -   On successful creation of the database.
' False   -   On failure.
'
' Created by    :   Lingappa Sindhanur
' Created on    :   23rd Jan 2002
' Last modified :   23rd Jan 2002
'
Public Function CreateDB(ByVal strdataFile As String, Optional strPwd As String) As Boolean
On Error GoTo ErrLine

Dim I As Byte
Dim j As Byte
Dim K As Byte

Dim strRet As String
Dim DBName As String
Dim DbFile As String
Dim DBPath As String
Dim strTblName As String
Dim CatalogObj As ADOX.Catalog
Dim tblData() As ADOX.Column
Dim ReqFld As Boolean

Dim IndxData() As ADOX.Index
Dim IndxCount As Integer
Dim IndxFlds As String

Dim RelationData() As ADOX.Key
Dim RelnCount As Integer
Dim SourceTableName As String
Dim RelnFlds As String
      

' Check if the specified datafile is existing.
If Dir(strdataFile, vbNormal) = "" Then
    MsgBox "File not found - " & strdataFile, vbCritical
    Exit Function
End If

I = 1

Do
   ' Read the dbname from datafile.
   strRet = ReadFromIniFile("Databases", "DataBase" & I, strdataFile)
   If strRet = "" Then Exit Do
   
   ' Get the name of the database file.
   DBName = ExtractToken(strRet, "dbName")
   If DBName = "" Then Exit Do
   
   ' Get the file path.
   DBPath = ExtractToken(strRet, "dbPath")
   DbFile = DBPath & "\" & DBName
   ' If no path mentioned, set the path as the path of data file.
   If DBPath = "" Then
       DbFile = AppendBackSlash(FilePath(strdataFile)) & DBName
       DBPath = FilePath(DbFile)
   End If

   ' Check if the file path of the database
   ' is existing.  If not create it.
   If Dir(DBPath, vbDirectory) = "" Then
       If Not MakeDirectories(DBPath) Then
           MsgBox "Error in creating the path " & DBPath _
               & " for " & DBName, vbCritical
           'MsgBox LoadResString(gLangOffSetNew + 809) & DBPath, vbCritical
           Exit Function
       End If

   Else
       ' Check if the file is already existing.
       ' If existing, get the user action.
       DbFile = StripExtn(DbFile) & ".mdb"
       If Dir(DbFile, vbNormal) <> "" Then
           Dim nRet As Integer
           nRet = MsgBox("WARNING : " & vbCrLf & vbCrLf & "The database file '" _
                   & DbFile & "' is already existing.  If you choose to overwrite " _
                   & "this file, you will loose the existing data permanantly." _
                   & vbCrLf & vbCrLf & "Do you want overwrite this file?", _
                   vbYesNo + vbCritical + vbDefaultButton2)
           If nRet = vbYes Then
               ' Delete the existing file.
               Kill DbFile
           ElseIf nRet = vbNo Then
               Exit Function
           End If
       End If
   End If
   
   Set CatalogObj = New ADOX.Catalog
   
   If Trim$(strPwd) = "" Then
      CatalogObj.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & DbFile & ";"
   Else
      CatalogObj.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & DbFile & ";" & _
                        "Jet OLEDB:Database Password=" & strPwd & ";"
   End If

   'Create the specified tables for this db.
   j = 1
   Do
      ' Read the table name.
      'Get the Table name
      strTblName = ReadFromIniFile(StripExtn(DBName), "Table" & j, strdataFile)
      If strTblName = "" Then Exit Do
      K = 0
      
      ReDim tblData(0)
      
      Set tblData(K) = New ADOX.Column
      Do
         strRet = ReadFromIniFile(strTblName, _
                     "Field" & K + 1, strdataFile)
         If strRet = "" Then Exit Do
         
         ' Add to fields array.
         ReDim Preserve tblData(K)
      
         Set tblData(K) = New ADOX.Column
         
         With tblData(K)
             ' Set the field name.
             .Name = ExtractToken(strRet, "FieldName")
             ' Set the field type.
             .Type = FieldTypeNum(ExtractToken(strRet, "FieldType"))
            ' Set the field length.
             .DefinedSize = Val(ExtractToken(strRet, "FieldLength"))
             ' Check, if the required flag is set.
              ReqFld = IIf((UCase$(ExtractToken(strRet, _
                     "Required")) = "TRUE"), True, False)
            
            Set .ParentCatalog = CatalogObj
            
            If ReqFld Then
               '.Attributes = adColFixed
               .Properties("Nullable") = False
               .Properties("Jet OLEDB:Allow Zero Length") = False
            Else
               '.Attributes = adColNullable
              .Properties("Nullable") = True
              .Properties("Jet OLEDB:Allow Zero Length") = True
            End If
            .Properties("Description") = ExtractToken(strRet, "Desc")
            
         End With
                  
         ' Increment the field count variable "k"
         K = K + 1
      Loop

      ' Create the table.
      If Not CreateTBL(CatalogObj, strTblName, tblData) Then
          MsgBox "Error in creating Table", vbCritical
          Exit Function
      End If

      ' If any indexes are specified, create them.
      K = 0
      IndxCount = 0
      
      Do
         strRet = ReadFromIniFile(strTblName, "Index" & K + 1, strdataFile)
         If strRet = "" Then Exit Do
         ReDim Preserve IndxData(K)
         Set IndxData(K) = New ADOX.Index
         IndxCount = K + 1
         
         IndxFlds = ExtractToken(strRet, "Fields")
         With IndxData(K)
             .Name = ExtractToken(strRet, "IndexName")
             .PrimaryKey = IIf(UCase$((ExtractToken(strRet, _
                         "Primary"))) = "TRUE", True, False)
             .Unique = IIf(UCase$((ExtractToken(strRet, _
                         "Unique"))) = "TRUE", True, False)
             .IndexNulls = IndexTypeNum(ExtractToken(strRet, "IndexNulls"))
         End With
         K = K + 1
      Loop
      
      If IndxCount > 0 Then
         If Not CreateIndexes(CatalogObj, strTblName, IndxData(), IndxFlds) Then
            MsgBox "CreateIndexes error:" & Err.Description
            Exit Function
         End If
      End If

      ' Increment the table count variable "j"
      j = j + 1
   Loop
    
   ' If any Relations are specified, create them.
   Dim L As Byte
   L = 0
   
   RelnCount = 0
   Do
      strRet = ReadFromIniFile("Relations", _
                  "Relation" & L + 1, strdataFile)
      If strRet = "" Then Exit Do
      SourceTableName = ExtractToken(strRet, "SourceTable")
      ReDim Preserve RelationData(L)
      Set RelationData(L) = New ADOX.Key
      RelnCount = L + 1
      RelnFlds = ExtractToken(strRet, "Field")
      With RelationData(L)
          .Name = ExtractToken(strRet, "RelationName")
          .RelatedTable = ExtractToken(strRet, "ForeignTable")
          .Type = RelationTypeEnum(ExtractToken(strRet, "RelationType")) 'adKeyForeign,adKeyPrimary,adKeyUnique
      End With
      
      L = L + 1
   Loop
   If RelnCount > 0 Then
      If Not CreateRelation(CatalogObj, SourceTableName, RelationData, RelnFlds) Then
         Exit Function
      End If
   End If
   ' Increment the DB count variable "i"
   I = I + 1
Loop

' Set the return value.
CreateDB = True


Exit Function

ErrLine:
    If Err.Number = 75 Then ' Path/File access error.
       nRet = MsgBox("File is already opened '" _
                & strRet & "'.", vbRetryCancel + vbCritical)
        If nRet = vbRetry Then Resume
    
    ElseIf Err Then
       MsgBox Err.Description, vbCritical
    End If

End Function

'
Private Function CreateIndexes(Cat As ADOX.Catalog, tbl As String, idx() As ADOX.Index, idxFlds As String) As Boolean

On Error GoTo CreateIndex_Error

Dim TheTable As ADOX.Table
Dim TheIndex As ADOX.Index
Dim fldArr() As String
Dim I As Integer, j As Integer

Set TheTable = Cat.Tables(tbl)
Set TheIndex = New ADOX.Index

For I = 0 To UBound(idx)
    With TheTable
        TheIndex.Name = idx(I).Name
        ' Breakup the fields string to get the field names for index.
        GetStringArray idxFlds, fldArr(), "+"
        
        With TheIndex
            For j = 0 To UBound(fldArr)
               .Columns.Append fldArr(j)
            Next j
            .PrimaryKey = idx(I).PrimaryKey
            .IndexNulls = idx(I).IndexNulls
'            .Unique = idx(I).Unique
        End With
        .Indexes.Append TheIndex
    End With
Next I

CreateIndexes = True
'Set cat = Nothing
Exit Function

CreateIndex_Error:
    If Err.Number = 3283 Then   ' Primary key already exists.
        Resume Next
    ElseIf Err Then
        MsgBox Err.Description, vbCritical
        Set Cat = Nothing
    End If
End Function


'
Private Function CreateRelation(Cat As ADOX.Catalog, srcTableName As String, Reln() As ADOX.Key, RlnFlds As String) As Boolean

On Error GoTo CreateRalation_Error

' Open the specified table.
Dim TheTable As ADOX.Table
Dim FKey As ADOX.Key
Dim RelationFlds() As String
Dim I As Integer, j As Integer

Set TheTable = Cat.Tables(srcTableName)
Set FKey = New ADOX.Key

For I = 0 To UBound(Reln)
   With FKey
      .Name = Reln(I).Name
      .Type = Reln(I).Type
      .RelatedTable = Reln(I).RelatedTable
      
      'Get the flds in an array
      GetStringArray RlnFlds, RelationFlds, "+"
      For j = 0 To UBound(RelationFlds)
         .Columns.Append RelationFlds(j)
      Next j
      
      j = 0
      For j = 0 To UBound(RelationFlds)
         .Columns(RelationFlds(j)).RelatedColumn = RelationFlds(j)
      Next
   End With
Next I

TheTable.Keys.Append FKey

Set Cat = Nothing

CreateRelation = True

Exit Function

CreateRalation_Error:
    If Err.Number = 9 Then
        Exit Function
    ElseIf Err.Number = 3283 Then   ' Primary key already exists.
        Resume Next
    ElseIf Err Then
        MsgBox Err.Description, vbCritical
        CreateRelation = False
        Set Cat = Nothing
    End If

End Function



'   This function creates the specified table whose fields are
'   defined in the array tblData().  tblData is a structure type
'   which specifies the following field information.
'       1. Field Name
'       2. Field Type
'       3. Field Length (Defined size) and
'       4. properties(required True/False)
Private Function CreateTBL(Cat As ADOX.Catalog, TblName As String, tblData() As ADOX.Column) As Boolean

Dim TheTable As ADOX.Table
Dim I As Integer
Dim AllowZero As Integer

On Error GoTo Err_Line

Set TheTable = New ADOX.Table

' Create the table.
With TheTable
   
   .Name = TblName
   
   Set .ParentCatalog = Cat
   
   For I = 0 To UBound(tblData)
      .Columns.Append tblData(I).Name, tblData(I).Type, tblData(I).DefinedSize
      
      .Columns(tblData(I).Name).Properties("Jet OLEDB:Allow Zero Length") = tblData(I).Properties("Jet OLEDB:Allow Zero Length")
      .Columns(tblData(I).Name).Properties("Description") = tblData(I).Properties("Description")
      .Columns(tblData(I).Name).Properties("Default") = tblData(I).Properties("Default")
      .Columns(tblData(I).Name).Properties("Nullable") = tblData(I).Properties("Nullable")
   
   Next I
   
End With

Cat.Tables.Append TheTable


CreateTBL = True

Exit Function

Err_Line:
    If Err.Number = 3010 Then 'table already exists
        Exit Function
    ElseIf Err.Number > 0 Then
        MsgBox Err.Description, vbCritical
        CreateTBL = False
    End If

End Function

' Returns the vb field type number, given a field type string.
Private Function FieldTypeNum(strFldType As String) As Integer
Select Case UCase$(strFldType)
    Case "ADTEXT", "ADVARWCHAR", "ADVARCHAR"
        FieldTypeNum = adVarWChar
    Case "NUMBER", "INTEGER", "DBNUMBER", "DBINTEGER", "ADINTEGER"
        FieldTypeNum = adSmallInt
    Case "DATE", "DBDATE", "ADDATE"
        FieldTypeNum = adDate
    Case "ADBOOLEAN", "DBBOOLEAN"
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


'
Private Function RelationTypeEnum(RelationType As String) As Integer
Select Case UCase$(RelationType)
   Case "ADKEYFOREIGN"
       RelationTypeEnum = adKeyForeign
   Case "ADKEYPRIMARY"
       RelationTypeEnum = adKeyPrimary
   Case "ADKEYUNIQUE"
        RelationTypeEnum = adKeyUnique
   Case Else
        RelationTypeEnum = 0
   End Select
End Function

' Returns the vb field type number, given a field type string.
Private Function IndexTypeNum(strIndexType As String) As Integer

Select Case UCase$(strIndexType)
    Case "DISALLOWNULLS"
        IndexTypeNum = adIndexNullsDisallow
    Case "NULLSIGNORE"
        IndexTypeNum = adIndexNullsIgnore
    Case "NULLSIGNOREANY"
        IndexTypeNum = adIndexNullsIgnoreAny
    Case Else
        MsgBox "Invalid index type entered in the tab file"
        End
End Select

End Function


'
Private Function AddNewField(SourceField As ADOX.Column, TargetTable As ADOX.Table) As Boolean

On Error GoTo Hell:

Dim TargetField As ADOX.Column

Set TargetField = New ADOX.Column

With TargetField
    .Name = SourceField.Name
    .Type = SourceField.Type
    .DefinedSize = SourceField.DefinedSize
    .Attributes = SourceField.Attributes

End With

TargetTable.Columns.Append TargetField

With TargetField
    
    .Properties("Default").Value = SourceField.Properties("Default").Value
    .Properties("Jet OLEDB:Allow Zero Length").Value = SourceField.Properties("Jet OLEDB:Allow Zero Length").Value
    .Properties("Description").Value = SourceField.Properties("Description").Value
    
End With

AddNewField = True

Exit Function

Hell:
    AddNewField = True
    
End Function
'
Private Function TableExists(ByVal TableName As String, Cat As ADOX.Catalog) As Boolean

Dim TableCount As Integer
Dim Item As Integer

TableCount = Cat.Tables.Count

TableExists = False

For Item = 0 To TableCount - 1
    If UCase$(Cat.Tables(Item).Name) = UCase$(TableName) Then
        TableExists = True
        Exit Function
    End If
Next Item


End Function
'
Private Function FieldExists(ByVal FieldName As String, TheTable As ADOX.Table) As Boolean

Dim FieldCount As Integer
Dim Item As Integer

FieldExists = False

If FieldName = "" Then Exit Function

FieldCount = TheTable.Columns.Count

For Item = 0 To FieldCount - 1
    If UCase$(TheTable.Columns(Item).Name) = UCase$(FieldName) Then
        FieldExists = True
        Exit Function
    End If
Next Item



End Function
'
Private Function FieldTypeSame(ByVal SourceField As ADOX.Column, ByVal TargetField As ADOX.Column) As Boolean

FieldTypeSame = False

If SourceField Is Nothing Then Exit Function
If TargetField Is Nothing Then Exit Function

If SourceField.Type = TargetField.Type Then FieldTypeSame = True

End Function
'
Private Function UpdateTableField(SourceTable As ADOX.Table, TargetTable As ADOX.Table) As Boolean
'Declare the variables
Dim SourceItem As Integer
Dim SourceFieldCount As Integer
Dim SourceFieldName As String
Dim FieldFound As Boolean


SourceFieldCount = SourceTable.Columns.Count

For SourceItem = 0 To SourceFieldCount - 1
    
    SourceFieldName = SourceTable.Columns(SourceItem).Name
    
    FieldFound = False
    
    FieldFound = FieldExists(SourceFieldName, TargetTable)
    If Not FieldFound Then If Not AddNewField(SourceTable.Columns(SourceItem), TargetTable) Then Exit Function
    
Next SourceItem

UpdateTableField = True

End Function

'
Private Function ValidateDBPath(ByVal DbFile As String) As Boolean
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim DBPath As String
Dim nRet As Integer

ValidateDBPath = False

If DbFile = "" Then Exit Function
' If no path mentioned, set the path as the path of data file.
DBPath = FilePath(DbFile)

' Check if the file path of the database
' is existing.  If not create it.
If Dir(DBPath, vbDirectory) = "" Then
    If Not MakeDirectories(DBPath) Then
        MsgBox "Error in creating the path " & DBPath _
            & " for " & constDBName, vbCritical
        Exit Function
    End If

Else
    ' Check if the file is already existing.
    ' If existing, get the user action.
    If Dir(DbFile, vbNormal) <> "" Then
        nRet = MsgBox("WARNING : " & vbCrLf & vbCrLf & "The database file '" _
                & DbFile & "' is already existing.  If you choose to overwrite " _
                & "this file, you will loose the existing data permanantly." _
                & vbCrLf & vbCrLf & "Do you want overwrite this file?", _
                vbYesNo + vbCritical + vbDefaultButton2)
        If nRet = vbYes Then
            ' Delete the existing file.
            Kill DbFile
        ElseIf nRet = vbNo Then
            Exit Function
        End If
    End If
End If

ValidateDBPath = True

Exit Function

ErrLine:
    MsgBox "ValidateDBPath()" & vbCrLf & Err.Description
    
End Function
'copy the default shipped database to the respective directory
'
Public Function CopyDefaultDatabase(ByVal SourceFile As String, TargetFile As String) As Boolean
'Trap an error
On Error GoTo ErrLine
Dim DBPath As String

'dbpath =
CopyDefaultDatabase = False

' check for the existence of path
If Not ValidateDBPath(TargetFile) Then Exit Function

'copy the default database to the required directory
FileCopy SourceFile, TargetFile
CopyDefaultDatabase = True

Exit Function
ErrLine:
    MsgBox "CopyDefaultDatabase():" & vbCrLf & Err.Description
    
End Function

