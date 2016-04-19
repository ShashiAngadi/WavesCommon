Attribute VB_Name = "CustTransfer"
Option Explicit

Public Function TransferNameTab(OldDBName As String, NewDBName As String) As Boolean

Dim OldTrans As New clsOldUtils
Dim NewTrans As New clsDBUtils

Dim LanguageName As String
Dim CompanyName As String

LanguageName = ""
CompanyName = ""

If Not OldTrans.OpenDB(OldDBName, OldPwd) Then Exit Function
If Not NewTrans.OpenDB(NewDBName, NewPwd) Then
    OldTrans.CloseDB
    Exit Function
End If

Dim SqlStr As String
Dim OldRst As Recordset


OldTrans.SQLStmt = "SELECT * FROM Install WHERE KeyData = 'CompanyName'"
If OldTrans.Fetch(OldRst, adOpenDynamic) > 0 Then _
        CompanyName = FormatField(OldRst("ValueData"))
OldTrans.SQLStmt = "SELECT * FROM Install WHERE KeyData = 'Language'"
If OldTrans.Fetch(OldRst, adOpenDynamic) > 0 Then _
        LanguageName = FormatField(OldRst("ValueData"))
    
SqlStr = "SELECT * FROM NameTab ORDER BY CustomerID"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenDynamic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If


On Error GoTo Err_Line

'Now TransFer the Details of NAme Tab
Dim CustomerId As Long
CustomerId = 0
On Error Resume Next

With frmMain
    .lblProgress = "Transferring Customer details"
    .prg.Max = OldRst.RecordCount + 1
    .Refresh
End With

'We have plan to introduce the CustomerType
'namely Individual,Instintutions,Societies ,etc
'At the same time whle transferring the old customer
'al customers are treated as Individual
'So Insert the Individual customer type
NewTrans.BeginTrans
'Customer Type
NewTrans.SQLStmt = "INSERT INTO CustomerType " & _
        " (CustType,CustTypeName,UIType ) " & _
        " VALUES ( 1, " & _
        AddQuotes(LoadResString(gLangOffSet + 252), True) & "," & _
        " 0 ) "
        
NewTrans.SQLExecute

'We have plan to introduce the Account Group
'namely General,self help group,etc
'At the same time while transferring the old accounts
'all accounts will b treated as genereal
'So Insert the genearal type

'Customer Type
NewTrans.SQLStmt = "INSERT INTO AccountGroup" & _
        " (AccGroupId,GroupName) " & _
        " VALUES ( 1, " & _
        AddQuotes(LoadResString(gLangOffSet + 339), True) & _
        " ) "
        
NewTrans.SQLExecute


''now insert the Company detiails
If Len(CompanyName) > 0 Then
    NewTrans.SQLStmt = " INSERT INTO CompanyCreation " & _
            "(HeadID,CompanyName,CompanyType ) " & _
            " VALUES ( " & _
            1 & "," & _
            AddQuotes(CompanyName, True) & "," & _
            0 & ")"

    NewTrans.SQLExecute
    NewTrans.SQLStmt = " INSERT INTO GodownDet(GodownID,GodownName) " & _
                   " VALUES  ( " & _
                   1 & "," & _
                   AddQuotes(CompanyName, True) & " )"
    NewTrans.SQLExecute
End If

''now insert the Language
If Len(LanguageName) > 0 Then
    NewTrans.SQLStmt = " INSERT INTO Install " & _
            "(KeyData,ValueData) " & _
            " VALUES ( 'Language'," & _
            AddQuotes(LanguageName, True) & ")"
            
    NewTrans.SQLExecute
End If

NewTrans.CommitTrans

SqlStr = "SELECT * FROM NameTab ORDER BY CustomerID"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenDynamic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If

'There are some error in the old Customer Details
'That Some CustomerID entered as 0
'so to carry forward those details
'we are inserting one record with customerid = 0
    CustomerId = 0
    SqlStr = "Insert INTO NameTab (CustomerID,Title," & _
            "FirstName,MIddleNAme,LastName, " & _
            "Gender,DOB,Profession,Caste,Place,CustType," & _
            "MaritalStatus,HomeAddress,OfficeAddress," & _
            "HomePhone,Officephone,eMail," & _
            "Reference,IsciName) " '
    SqlStr = SqlStr & " Values (" & _
            CustomerId & "," & _
            AddQuotes(FormatField(OldRst("Title")), True) & "," & _
            AddQuotes(FormatField(OldRst("FirstName")), True) & "," & _
            AddQuotes(FormatField(OldRst("MiddleName")), True) & "," & _
            AddQuotes(FormatField(OldRst("LastName")), True) & "," & _
            OldRst("Gender") & "," & _
            FormatDateField(OldRst("DOB")) & "," & _
            AddQuotes(FormatField(OldRst("Profession")), True) & "," & _
            AddQuotes(FormatField(OldRst("Caste")), True) & "," & _
            AddQuotes(FormatField(OldRst("Place")), True) & "," & _
            " 0 ," & _
            FormatField(OldRst("MaritalStatus")) & "," & _
            AddQuotes(FormatField(OldRst("HomeAddress")), True) & "," & _
            AddQuotes(FormatField(OldRst("OfficeAddress")), True) & "," & _
            AddQuotes(FormatField(OldRst("HomePhone")), True) & "," & _
            AddQuotes(FormatField(OldRst("OfficePhone")), True) & "," & _
            AddQuotes(FormatField(OldRst("eMail")), True) & "," & _
            FormatField(OldRst("Reference")) & "," & _
            AddQuotes(FormatField(OldRst("IsciName")), True) & _
            " )"
    
    NewTrans.BeginTrans
    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
    Else
        NewTrans.CommitTrans
    End If

''Begin the transaction to for each record
NewTrans.BeginTrans

While Not OldRst.EOF
    If CustomerId = OldRst("CustomerID") Then GoTo NextName
    CustomerId = OldRst("CustomerID")
    SqlStr = "Insert INTO NameTab (CustomerID,Title," & _
            "FirstName,MIddleNAme,LastName, " & _
            "Gender,DOB,Profession,Caste,Place,CustType, " & _
            "MaritalStatus,HomeAddress,OfficeAddress," & _
            "HomePhone,Officephone,eMail,Reference,IsciName) " '
    SqlStr = SqlStr & " Values (" & _
            CustomerId & "," & _
            AddQuotes(OldRst("Title"), True) & "," & _
            AddQuotes(OldRst("FirstName"), True) & "," & _
            AddQuotes(OldRst("MiddleName"), True) & "," & _
            AddQuotes(OldRst("LastName"), True) & "," & _
            OldRst("Gender") + 1 & "," & _
            FormatDateField(OldRst("DOB")) & "," & _
            AddQuotes(OldRst("Profession"), True) & "," & _
            AddQuotes(OldRst("Caste"), True) & "," & _
            AddQuotes(OldRst("Place"), True) & "," & _
            " 0 ," & _
            OldRst("MaritalStatus") & "," & _
            AddQuotes(OldRst("HomeAddress"), True) & "," & _
            AddQuotes(OldRst("OfficeAddress"), True) & "," & _
            AddQuotes(OldRst("HomePhone"), True) & "," & _
            AddQuotes(OldRst("OfficePhone"), True) & "," & _
            AddQuotes(OldRst("eMail"), True) & "," & _
            FormatField(OldRst("Reference")) & "," & _
            AddQuotes(OldRst("IsciName"), True) & _
            " )"
    If Err.Number <> 0 Then
        SqlStr = ""
        Err.Clear
        SqlStr = "Insert INTO NameTab (CustomerID,Title," & _
            "FirstName,MIddleNAme,LastName, " & _
            "Gender,DOB,Profession,Caste,Place,CustType," & _
            "MaritalStatus,HomeAddress,OfficeAddress," & _
            "HomePhone,Officephone,eMail,Reference,IsciName) " '
        SqlStr = SqlStr & " Values (" & _
            CustomerId & "," & _
            AddQuotes(FormatField(OldRst("Title")), True) & "," & _
            AddQuotes(FormatField(OldRst("FirstName")), True) & "," & _
            AddQuotes(FormatField(OldRst("MiddleName")), True) & "," & _
            AddQuotes(FormatField(OldRst("LastName")), True) & "," & _
            Val(FormatField(OldRst("Gender"))) + 1 & "," & _
            FormatDateField(OldRst("DOB")) & "," & _
            AddQuotes(FormatField(OldRst("Profession")), True) & "," & _
            AddQuotes(FormatField(OldRst("Caste")), True) & "," & _
            AddQuotes(FormatField(OldRst("Place")), True) & "," & _
            " 0 ," & _
            FormatField(OldRst("MaritalStatus")) & "," & _
            AddQuotes(Left(FormatField(OldRst("HomeAddress")), 149), True) & "," & _
            AddQuotes(Left(FormatField(OldRst("OfficeAddress")), 99), True) & "," & _
            AddQuotes(FormatField(OldRst("HomePhone")), True) & "," & _
            AddQuotes(FormatField(OldRst("OfficePhone")), True) & "," & _
            AddQuotes(FormatField(OldRst("eMail")), True) & "," & _
            FormatField(OldRst("Reference")) & "," & _
            AddQuotes(FormatField(OldRst("IsciName")), True) & _
            " )"
    End If
    
    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        'Dim x As String
        'x = IIf(OldRst("DOB") = Null, "T", "F")
        Exit Function
    End If
    SqlStr = ""

NextName:
    With frmMain
        .lblProgress = "Transferring customer details"
        .prg.Value = OldRst.AbsolutePosition
    End With
    
    OldRst.MoveNext

Wend

''Commit the transaction
NewTrans.CommitTrans

Set OldRst = Nothing
On Error GoTo 0
SqlStr = "UPDATE NameTab Set DOB = NULL WHERE DOB = #1/1/100# "
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
    Exit Function
End If
OldTrans.CommitTrans


'Now Insert the Caste Details
SqlStr = "SELECT * FROM Castetab"
Dim Count As Integer
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenDynamic) > 0 Then
    Count = 1
    While Not OldRst.EOF
        SqlStr = "Insert INTO CasteTab (Caste,CasteID) " & _
            " Values (" & _
            AddQuotes(FormatField(OldRst("CASTE")), True) & ", " & _
            Count & ")"
        NewTrans.BeginTrans
        NewTrans.SQLStmt = SqlStr
        If Not NewTrans.SQLExecute Then
            NewTrans.RollBack
            Exit Function
        End If
        
        NewTrans.CommitTrans
        Count = Count + 1
NextCaste:
    
    With frmMain
        .lblProgress = "Transferring Customer details"
        .prg.Value = OldRst.AbsolutePosition
    End With
        OldRst.MoveNext

    Wend
End If


'Now Insert the Place Details
SqlStr = "SELECT * FROM PLACETab"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenDynamic) > 0 Then
    Count = 1
'    Debug.Assert Count <> 1
    While Not OldRst.EOF
        If Len(Trim$(FormatField(OldRst("Places")))) Then
            SqlStr = "Insert INTO PLACETab (PLACE,PlaceID) " & _
                " Values (" & _
                AddQuotes(FormatField(OldRst("Places")), True) & ", " & _
                Count & ")"
            NewTrans.BeginTrans
            NewTrans.SQLStmt = SqlStr
            If Not NewTrans.SQLExecute Then
                NewTrans.RollBack
                Exit Function
            End If
            NewTrans.CommitTrans
            Count = Count + 1
        End If
NextPlace:
    With frmMain
        .lblProgress = "Transferring customer details"
    End With
    
        OldRst.MoveNext
    Wend
End If

SqlStr = "Insert INTO UserTab (UserID,CustomerID," & _
        "Permissions,LoginName,LoginPassword,Deleted) " & _
        " Values ( 0 ," & _
        "0 ," & _
        "2043 ," & _
        "'Mahesh' ," & _
        "'Sunil', " & _
        False & ")"
    
NewTrans.BeginTrans
NewTrans.SQLStmt = SqlStr

If Not NewTrans.SQLExecute Then
    NewTrans.RollBack
Else
    NewTrans.CommitTrans
End If
    
'Now Insert the user Details
SqlStr = "SELECT * FROM UserTab"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenDynamic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If

With frmMain
    .lblProgress = "Transferring customer details"
    .prg.Value = 0
    .Refresh
End With

Dim Perm As Long

While Not OldRst.EOF
    If OldRst("CustomerId") = -1 Then GoTo NextUser
    
    Perm = FormatField(OldRst("Permissions"))
    Perm = IIf(Perm = 64, 1, 256) 'perPigmyOperator
    If OldRst.AbsolutePosition = 0 Then Perm = 1 'perFullPermissions
    SqlStr = "Insert INTO UserTab (UserID,CustomerID," & _
        "Permissions,LoginName,LoginPassword,Deleted) " & _
        " Values (" & OldRst("UserId") & "," & _
        OldRst("CustomerId") & "," & _
        Perm & "," & _
        AddQuotes(FormatField(OldRst("LoginName")), True) & "," & _
        AddQuotes(FormatField(OldRst("PassWord")), True) & ", " & _
        False & ")"
    NewTrans.BeginTrans
    NewTrans.SQLStmt = SqlStr
    
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        MsgBox "Unable to transfer the User permissions", vbInformation, wis_MESSAGE_TITLE
        If OldRst.AbsolutePosition > 1 Then GoTo NextUser
        Exit Function
    End If
    NewTrans.CommitTrans

NextUser:
    With frmMain
        .lblProgress = "Transferring customer Details"
        .prg.Value = OldRst.AbsolutePosition
    End With
    OldRst.MoveNext

Wend

'Check for the Transtype
If NewIndexTrans Is Nothing Then Set NewIndexTrans = NewTrans
''Now Check the PArent Heads in the ParentHeads Table
NewTrans.SQLStmt = "Select * From ParentHeads"
If NewTrans.Fetch(OldRst, adOpenDynamic) < 1 Then Call InsertParentHeads

TransferNameTab = True

With frmMain
    .lblProgress = "Transferred the Customer details"
    .prg.Value = 0
    .Refresh
End With

Call CreateDefaultView

Screen.MousePointer = vbDefault

Err_Line:
If Err.Number = 3021 Then Err.Clear: Resume Next
If Err.Number Then
    MsgBox "Error In customer Master"
    Err.Clear
End If

End Function

