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



Public Property Get DataBaseName() As String
    On Error GoTo LastLine:
    If m_DataBaseObject Is Nothing Then Exit Sub
    DataBaseName = m_DataBaseObject.Properties("Data Source Name").Value
LastLine:
    Err.Clear
End Property

Private Sub FormatSQL()
'This FUnction Coverts the Sql Statments Date Part into any other formate
'Here we are converting date format from dd/mm/yyyy to 'dd MMM YYYY'
'Check the the Data parameteres And Change their format to 'dd MMM yyyy'
Dim Pos As Integer
Dim SecPos As Integer
Dim strDate As String

Do
    Pos = InStr(SecPos + 1, m_SQLStmt, "#", vbBinaryCompare)
    If Pos = 0 Then Exit Do
SplChar:
    SecPos = InStr(Pos + 1, m_SQLStmt, "#", vbBinaryCompare)
    If SecPos = 0 Then Exit Do
    If InStr(Pos, m_SQLStmt, "/") < SecPos Then
        On Error Resume Next
        strDate = Format(Mid(m_SQLStmt, Pos + 1, SecPos - Pos - 1), "DD MMM YYYY")
        If Err Then
            Err.Clear
            Pos = SecPos
            GoTo SplChar
        Else
            m_SQLStmt = Left(m_SQLStmt, Pos) & _
                Format(Mid(m_SQLStmt, Pos + 1, SecPos - Pos - 1), "DD MMM YYYY") & _
                Mid(m_SQLStmt, SecPos)
        End If
    End If
    SecPos = InStr(Pos + 1, m_SQLStmt, "#", vbBinaryCompare)
Loop

End Sub

Public Property Get isInTransaction() As Boolean
    If m_PrevOp = op_BEGINTRANS Then
        isInTransaction = True
    Else
        isInTransaction = False
    End If
End Property

Public Property Get ObjectName()
    ObjectName = "TransactionClass"
End Property


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
        
        'm_Catalog.Views.Delete (lpCount)
        
        SqlStr = m_SQLStmt
        
        isViewFound = True
        
        If Not DropView(ViewName) Then Exit Function
        
        m_SQLStmt = SqlStr
        
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
Dim SqlStr As String


'initalise the functions
DeleteProcedures = False

Set m_Catalog = New ADOX.Catalog

m_Catalog.ActiveConnection = m_DataBaseObject

ViewCount = m_Catalog.Procedures.Count - 1

For lpCount = ViewCount To 0 Step -1
    If UCase$(m_Catalog.Procedures.Item(lpCount).Name) = UCase$(ViewName) Then
        
        SqlStr = m_SQLStmt
        
        If Not DropView(ViewName) Then Exit Function
        
        m_SQLStmt = SqlStr
        
        Exit For
   End If
   
Next lpCount

DeleteProcedures = True

Exit Function

ErrLine:
   MsgBox "DeleteProcedures: " & vbCrLf & Err.Description, vbCritical
      

End Function

Public Sub ReturnActiveConnection(Connection As ADODB.Connection)

On Error GoTo hell:

Set Connection = New ADODB.Connection

Connection = m_DataBaseObject

Exit Sub

hell:

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

Set gDbTrans.CommandObject = New ADODB.Command


'format the sql
FormatSQL

gDbTrans.CommandObject.CommandText = m_SQLStmt

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

Private Function DropView(ViewName As String) As Boolean

'Trap the error
On Error GoTo ErrLine

m_SQLStmt = " DROP TABLE " & ViewName

Dim transBegun As Boolean
transBegun = isInTransaction
If Not transBegun Then gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then _
    DropView = False: If Not transBegun Then gDbTrans.RollBacknRaiseError
    

If Not transBegun Then gDbTrans.CommitTrans
DropView = True

ErrLine:
    

End Function


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
On Error GoTo hell

Dim SearchArray() As String
Dim ItemCount As Integer
Const SEARCH_STRING_NO = 1
' initialise the function to false
FindRecord = False

' if any filter is there then remove the filter

rstToFind.Filter = adFilterNone


Call GetStringArray(SearchString, SearchArray, ",")
'in the VB6 U can use SPLIT function  which will be faster
'than GetstringArray

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

hell:
    MsgBox "Field Specified or Value given is wrong !!", vbCritical, wis_MESSAGE_TITLE
    
End Function

Public Function ParameterizedQuery(qryName As String) As Boolean
On Error GoTo ErrLine

ParameterizedQuery = False

' Open the catalog
m_Catalog.ActiveConnection = m_DataBaseObject

' Get the Command object from the
' Procedure
Set CommandObject = m_Catalog.Procedures(qryName).Command
   
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

'Format the Date parts of SQL
FormatSQL

'Execute the Action Query.
m_DataBaseObject.Execute m_SQLStmt ', TheRecordsAffected, adExecuteNoRecords
   
SQLExecute = True

'Reset sqlstmt to avoid misuse
m_SQLStmt = ""

Exit Function

ErrLine:
        MsgBox "SQLExecute: The following error occurred." _
                & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
        
        'm_SQLStmt = ""
        SQLExecute = False
    
End Function

'This Function will check the transaction status
'if transaction is made if u r exiting from the application
' It will prompt u the transaction to be commited or rollbacked
Public Sub CloseDB()

If Not m_DBOpen Then Exit Sub

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

If Not m_DataBaseObject Is Nothing Then m_DataBaseObject.Close
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
    'error nav
    'when clicking the cancel button it rise the
    'data base error
    
        
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

'Check the the Data parameteres And Change their format to 'dd MMM yyyy'
FormatSQL

rstFetch.Open m_SQLStmt, m_DataBaseObject, TheCursorType, adLockReadOnly, adCmdText

Select Case TheCursorType
   Case adOpenStatic
      Fetch = rstFetch.RecordCount
   Case Else
      Fetch = IIf(rstFetch.RecordCount < 0, 1, rstFetch.RecordCount)
      If rstFetch.BOF And rstFetch.EOF Then Fetch = 0 Else Fetch = 1
End Select

' Disconnect the Recordset
Set rstFetch.ActiveConnection = Nothing

'Set the sqlstmt to nullm_SQLStmt = ""

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

Public Function OpenDB(strDBName As String, pwd As String) As Boolean

Dim Retry As Boolean

If m_DBOpen Then
    OpenDB = True
    Exit Function
End If
Retry = False
If strDBName = "" Then Exit Function

If m_DataBaseObject Is Nothing Then Set m_DataBaseObject = New ADODB.Connection

m_DataBaseObject.Mode = adModeShareDenyNone
RetryLine:
On Error GoTo ErrLine
'Open a connection Using Microsoft.Jet Provider
m_DataBaseObject.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & strDBName & ";" & _
                       "Jet OLEDB:Database Password=" & pwd & ";"

OpenDB = True
m_DBOpen = True

Exit Function

ErrLine:
If Retry = False Then
    Retry = True
    pwd = "WIS!@#"
    GoTo RetryLine
End If

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
Call gDbTrans.DeleteProcedures(ViewName)

Set gDbTrans.CommandObject = New ADODB.Command

'Format the SQl
FormatSQL

gDbTrans.CommandObject.CommandText = m_SQLStmt

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

Public Property Let SQLStmt(ByVal vNewValue As String)
m_SQLStmt = vNewValue
End Property

