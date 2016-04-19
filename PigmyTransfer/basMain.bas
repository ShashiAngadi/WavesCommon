Attribute VB_Name = "basMain"
Public gDbTrans As clsTransact
Public gDBFileName As String
Public gStrDate As String
Public gCompanyName As String
Public gAppName As String
Public gAgentID As Integer
Dim gCurrUser As clsUsers
Public gDEVICE As String
Public gAppPath As String

Public Sub Main()
gAppName = "Pigmy Index"
'Get Date Form
DateFormat = "dd/mm/yyyy"

'Call Initialize
'Initialize the global variables
If gDbTrans Is Nothing Then Set gDbTrans = New clsTransact

gAppPath = App.Path

Call KannadaInitialize

Set gCurrUser = New clsUsers

With gCurrUser
    .MaxRetries = 3
    .ShowLoginDialog
    If Not .LoginStatus Then
        Do
            Call ExitApplication(False, 1)
        Loop
        End
    End If
End With

Call KannadaInitialize

gDEVICE = UCase$(ReadFromIniFile("Pigmy", "Machine", App.Path & "\" & constFINYEARFILE))


'Temprary code for a year
'Now Crete the Tab Main

gDbTrans.SQLStmt = "SELECT CustomerID,Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,Place,Caste,Gender,IsciName From NameTab"
gDbTrans.CreateView ("QryName")

gDbTrans.SQLStmt = "SELECT CustomerID,Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,IsciName From NameTab"
gDbTrans.CreateView ("QryOnlyName")

''''Temp code ends here

'If It is online then Show theBegin date
'If gOnLine Then Call BeginDayTrans
'Now Check the user permissions
Dim Perm As wis_Permissions

Perm = gCurrUser.UserPermissions

If (Perm And perPigmyAgent) Then
    'He is Pigmy Agent, so test the agent id
    gAgentID = gCurrUser.UserID
Else
    'Now check for the admin previlages
    If Not ((Perm And perBankAdmin) > 0 Or (Perm And perOnlyWaves) > 0) Then
        MsgBox "You do not have permission for the Pigmy operations"
        Call ExitApplication(False, 0)
    End If
    gAgentID = 0
End If

'Now Show the frmmain
frmMain.Show vbModal

On Error Resume Next
Call gDbTrans.CloseDB
End Sub


Public Sub ExitApplication(Confirm As Boolean, Cancel As Integer)

If Confirm Then
    ' Ask for user confirmation.
    Dim nRet As Integer
    'nRet = MsgBox("Do you want to exit this application?", _
            vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
    nRet = MsgBox(LoadResString(gLangOffSet + 750), _
            vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
    If nRet = vbNo Then Cancel = True: Exit Sub
End If

If gWindowHandle Then CloseWindow (gWindowHandle)

On Error Resume Next
If gLangOffSet Then Call NudiResetAllFlags: Call NudiStopKeyboardEngine
Debug.Print IIf(NudiGetLastError = 0, 2, 1)

Unload wisMain
Set wisMain = Nothing
Set gCurrUser = Nothing
'Set wisAppObj = Nothing

If Not gDbTrans Is Nothing Then gDbTrans.CloseDB
Set gDbTrans = Nothing

'1 For Exit
'2 For Reboot
'Call ExitWindowsEx(2, 0)
   
   End


End Sub


Public Sub LoadAgentNames(cmbAgents As ComboBox, Optional AgentID As Integer)

Dim I As Integer
Dim Perms As wis_Permissions
Dim Rst As Recordset

    cmbAgents.Clear
    Dim itemIndex As Integer
    
    Perms = perPigmyAgent
    gDbTrans.SQLStmt = "Select * from UserTab WHERE (DELETED = FALSE or DELETED is NULL) "
    Call gDbTrans.Fetch(Rst, adOpenForwardOnly)
     
    'Dim CustReg As clsCustReg
    'Set CustReg = New clsCustReg
    
    For I = 1 To Rst.RecordCount
        If Val(Rst("Permissions")) And Perms Then
            'CustReg.LoadCustomerInfo (Val(Rst("CustomerID")))
            cmbAgents.AddItem CustomerName(Val(Rst("CustomerId")))
            cmbAgents.ItemData(cmbAgents.NewIndex) = Val(Rst("UserID"))
            If (Val(Rst("UserID")) = AgentID) Then itemIndex = cmbAgents.ListCount
        End If
        Rst.MoveNext
    Next I
    
    cmbAgents.ListIndex = itemIndex - 1
End Sub


Public Function CustomerName(CustomerID As Long)
Dim Rst As ADODB.Recordset
    gDbTrans.SQLStmt = "Select Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName as Name " & _
            " From NameTab Where CustomerId = " & CustomerID
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) <> 1 Then
        CustomerName = " "
    Else
        CustomerName = FormatField(Rst("Name"))
    End If
End Function

Public Function GetRecordSet(AgentID As Integer) As Recordset

Dim SQLStmt As String

SQLStmt = "Select Max(TransId) AS MaxTransID, A.AccID" & _
    " From PDTrans B Inner Join PDMaster A On B.AccId = A.AccId " & _
    " Where A.AgentId = " & AgentID & _
    " AND ClosedDate is NULL GROUP BY A.AccID"
     
gDbTrans.SQLStmt = SQLStmt
If Not gDbTrans.CreateView("QryTemp") Then Exit Function

gDbTrans.SQLStmt = "SELECT CustomerID,FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,IsciName,FullName From NameTab"
gDbTrans.CreateView ("QryUserName")


SQLStmt = "Select  B.Balance, A.AgentId,A.CreateDate, A.AccID,A.AccNum,A.CustomerId, Name,FullName,B.TransID, 'Pigmy' as PigmyType, val(A.AccNum) as AcNum " & _
    " From QryUserName C Inner join (PDMaster A Inner join " & _
    " (PDtrans B Inner join QryTemp D ON B.TransId = D.MaxTransID AND D.AccID = B.AccID )" & _
        " On A.AccID = B.AccId )" & _
    " ON C.CustomerId = A.CustomerId " & _
    " Where A.AgentID = " & AgentID '& _
    " Order by val(A.AccNum)"
    
''here comes the Agent which has Account,but no transaction

gDbTrans.SQLStmt = SQLStmt

SQLStmt = "Select 0 as Balance, A.AgentId, A.CreateDate, A.AccID,A.AccNum, A.CustomerId, Name,FullName, 0 as TransID, 'Pigmy' as PigmyType, val(A.AccNum) as AcNum   " & _
        " From QryUserName C Inner join PDMaster A ON C.CustomerId = A.CustomerId" & _
        " Where A.AgentID = " & AgentID & " And A.AccID not in (select distinct AccID from pdtrans) " & _
        " "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " UNION " & SQLStmt & " Order by AcNum"

Dim Rst As Recordset
Call gDbTrans.Fetch(Rst, adOpenForwardOnly)
Set GetRecordSet = Rst
   
End Function
Public Function InsertPigmyAmount(AccID As Integer, FirstTransDate As Date, Amount() As Double) As Boolean
    Dim TransDate As Date
    
    TransDate = FirstTransDate
    InsertPigmyAmount = False
    
    Dim IsNewTrans As Boolean
    'Get the Last TransID for this account
    gDbTrans.SQLStmt = "Select Balance,AccID,TransID,TransDate from PDTrans Where AccID = " & AccID & _
        " And TransID = (Select max(TransID) from PDTrans where AccID = " & AccID & " )"
    Dim Rst As Recordset
    Dim Balance As Double
    Dim TransID As Long
    Dim Trans As wisTransactionTypes
    Trans = wDeposit
    Call gDbTrans.Fetch(Rst, adOpenDynamic)
    If Rst.RecordCount = 0 Then
        Balance = 0
        TransID = 0
    Else
        Balance = FormatField(Rst("Balance"))
        TransID = FormatField(Rst("TransID"))
    End If
        
    If Not gDbTrans.isInTransaction Then
        IsNewTrans = True
        gDbTrans.BeginTrans
    End If
    
    Dim I As Integer
    For I = 0 To UBound(Amount)
        If Amount(I) > 0 Then
            TransID = TransID + 1
            Balance = Balance + Amount(I)
            gDbTrans.SQLStmt = "Insert into PDTrans (AccID, TransID, " & _
                " TransDate, Amount,Balance, Particulars," & _
                " TransType, VoucherNo) values ( " & _
                AccID & "," & _
                TransID & "," & _
                "#" & TransDate & "#," & _
                Amount(I) & "," & _
                Balance & "," & "'From Device'," & _
                Trans & ",'Pigmy Device')"
            
            If Not gDbTrans.SQLExecute Then
                If IsNewTrans Then gDbTrans.RollBack
                InsertPigmyAmount = False
                Exit Function
            End If
            
        End If
        
        TransDate = DateAdd("d", 1, TransDate)
    Next
    
    If IsNewTrans Then gDbTrans.CommitTrans
    InsertPigmyAmount = True
    Exit Function
ErrLine:
    If IsNewTrans Then gDbTrans.RollBack
End Function
' This function returns the ParentID from the given Headid
' Input is Headid as long
' Returns ParentID long
'
' Pradeep
'
Public Function GetParentID(HeadID As Long) As Long
' Handle Error
On Error GoTo NoParentID:

' Declare Variables
Dim rstParentID As ADODB.Recordset

' Intialiase the Variable
GetParentID = 0

' Check the Input Received if Zero then Exit
If HeadID = 0 Then Exit Function

' set the sqlstmt
gDbTrans.SQLStmt = " SELECT ParentID " & _
                   " FROM Heads " & _
                   " WHERE HeadID=" & HeadID
                   
' Now fetch the record
If gDbTrans.Fetch(rstParentID, adOpenForwardOnly) < 1 Then Exit Function

' Here is the ParentID!
GetParentID = FormatField(rstParentID("ParentID"))

Exit Function

NoParentID:
    
End Function

Public Function ReadSetupValue(Module As String, strKey As String, DefaultValue As String) As String
Dim DBStr As String
Dim Rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset

ReadSetupValue = DefaultValue
gDbTrans.SQLStmt = "Select * from Setup where " & _
                    "ModuleData like '" & Module & "'" & _
                    " And KeyData like '" & strKey & "'"
    If gDbTrans.Fetch(rst1, adOpenDynamic) < 1 Then Exit Function
        
        DBStr = FormatField(rst1("ValueData"))
        ReadSetupValue = ""
        If DBStr <> " " Then ReadSetupValue = DBStr

End Function

Public Function GetDBNameWithPath(ByVal strFinYearFile As String, ByVal YearID As Integer) As String

'Declare the constansts
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"

'Declare the variables
Dim encrKey As String
Dim encrSection As String
Dim strRet As String
Dim strRootPath As String

encrKey = strYear & YearID 'EncryptData(strYear & YearID)
encrSection = strFinYearSection ' EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then Exit Function

strRet = strRet 'DecryptData(strRet)

GetDBNameWithPath = ExportExtractToken(strRet, strDBPath, , ",")
    
strRootPath = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "ServerName")
If Len(strRootPath) Then
    Dim strShareName As String
    strShareName = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "ShareName")
    strRootPath = "\\" & strRootPath & "\" & strShareName
Else
    strRootPath = ReadFromIniFile("DBPath", "Path", App.Path & "\" & constFINYEARFILE)
    If Len(strRootPath) = 0 Then strRootPath = App.Path
End If

GetDBNameWithPath = strRootPath & "\" & GetDBNameWithPath & "\" & constDBName

Exit Function

ErrLine:
    MsgBox "GetDBNameWithPath()" & vbCrLf & Err.Description
    
End Function

' Retrieves the value for a specified token
' in a given source string.
' The source should be of type :
'       name1=value1,name2=value2,...,name(n)=value(n)
'   similar to DSN strings maintained by ODBC manager.
Public Function ExportExtractToken(src As String, TokenName As String, _
        Optional ByVal TokenDelim As String, Optional ByVal SepDelim As String) As String

' If the src is empty, exit.
If Len(src) = 0 Or _
    Len(TokenName) = 0 Then Exit Function

' Search for the token name.
Dim token_pos As Integer
Dim strSearch As String
Dim Delim_pos As Integer

strSearch = Trim$(TokenName & TokenDelim)

' Search for the token_name in the src string.
 token_pos = InStr(1, src, strSearch, vbTextCompare)
Do
    ' The character before the token_name
    ' should be "," or, it should be the first word.
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
    ElseIf Mid$(src, token_pos - 1, 1) = "," Then
        Exit Do
    Else
        'Get next occurance.
        token_pos = InStr(token_pos + 1, src, TokenName, vbTextCompare)
    End If
Loop

token_pos = token_pos + Len(strSearch)

' Search for the delimiter ",", after the token_pos.
Delim_pos = InStr(token_pos, src, SepDelim)

If Delim_pos = 0 Then Delim_pos = Len(src) + 1

'Return the token_value.
ExportExtractToken = Mid$(src, token_pos, Delim_pos - token_pos)

End Function


