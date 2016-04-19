VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private M_DataBase As Database
Private m_WorkSpace As Workspaces
Private m_Rst As Recordset

' User defined events
    'Informs about the current activity.
    Public Event UpdateStatus(strMsg As String)
    Public Event CreateDBStatus(strMsg As String, CreatedDBRatio As Single)
' Error object.
Private m_errNum As Integer
Private m_errDesc As String

Private m_SQLFile As String

Public Property Get DBObject() As Database
        Set newDB = M_DataBase
End Property
Public Property Let DBObject(newDB As Database)
        Set M_DataBase = newDB
End Property

Public Function SQLFetch(SQLStmt As String) As Recordset

' Raise event for status updation.
RaiseEvent UpdateStatus("Querying records...")

' If the datbase object is not initialized, exit.
If M_DataBase Is Nothing Then
    MsgBox "SQLFetch: No database is in use.  " _
        & "Set the database name before calling SQLFetch.", vbExclamation
    GoTo Exit_Line
End If

On Error GoTo ErrLine
Set SQLFetch = Nothing

Set SQLFetch = M_DataBase.OpenRecordset(SQLStmt, _
            dbOpenDynaset, dbSeeChanges, dbOptimistic)

Exit_Line:
    RaiseEvent UpdateStatus("")
    Exit Function

ErrLine:
    If Err Then
        Set SQLFetch = Nothing
        MsgBox "SQLFetch: The following error occurred." _
                & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
        m_errNum = Err.Number
        m_errDesc = Err.Description
    End If

End Function




