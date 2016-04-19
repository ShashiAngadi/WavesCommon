Attribute VB_Name = "basTemp"
Option Explicit
Public strKannadaFont As String
Public strEnglishFont As String
Public dbUtils As clsDBUtils
Dim rst As Recordset

Public Sub InitFonts()
    strEnglishFont = "MS Sans Serif"
    strKannadaFont = "SUCHI-KAN-0850" '"SHREE-KAN-0850"
End Sub

Public Sub ChangeDBFont(dbName As String)
    Dim tableName As String
    If Len(dbName) < 1 Then Exit Sub
    
    Set dbUtils = New clsDBUtils
    'Open the database
    If Not dbUtils.OpenDB(dbName, "WIS!@#") Then
        MsgBox "Unable to open the database", vbOKOnly, "DataBase"
        Exit Sub
    End If
        
    'BEGIN the TRANSACTION
            
    'NameTab
    Dim SQL_Stmt As String
    
    tableName = "NameTab"
    SQL_Stmt = "Select CustomerID,Title,FirstName,MiddleName,LastName," & _
        " Profession,Caste,HomeAddress,OfficeAddress,Place  from NameTab"
    'If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    If Not UpdateNameRecords() Then GoTo EndLine
            
    MsgBox "Customer Names Transfer Done", , "Index Database"
    
    'SBMASTER
    tableName = "SBMaster"
    SQL_Stmt = "Select AccID,JointHolder,NomineeName,NomineeRelation from SBMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'SBTrans
    tableName = "SBTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from SBTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'SBPLTrans
    tableName = "SBPLTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from SBPLTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "Saving Transfer Done", , "Index Database"
    
    'BkCC Master
    tableName = "BKCCMASTER"
    SQL_Stmt = "Select LoanId,Remarks from BKCCMaster" & _
        " Where Loanid Not in (Select LoanId from BKCCMaster Where Remarks = '' OR Remarks IS NULL)"
    
    Call UPdateRecords(SQL_Stmt, tableName)
    'If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'BKCCTrans
    tableName = "BKCCTrans"
    SQL_Stmt = "Select LoanId,TransID, Particulars from BKCCTrans where Particulars <> '' AND NOT Particulars  IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'BKCCIntTrans
    tableName = "BKCCIntTrans"
    SQL_Stmt = "Select LoanId,TransID, Particulars from BKCCIntTrans where Particulars <> '' AND NOT Particulars  IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "BKCC Transfer Done", , "Index Database"
    
    'PlaceTab
    tableName = "PlaceTab"
    SQL_Stmt = "Select PlaceId,Place from PlaceTab"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'CasteTab
    tableName = "CasteTab"
    SQL_Stmt = "Select CasteId,Caste from CasteTab"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    
    'MEMMASTER
    tableName = "MemMaster"
    SQL_Stmt = "Select AccID,NomineeRelation from MemMaster Where NomineeRelation <> '' And Not NomineeRelation is NULL "
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'MEMTrans
    tableName = "MEMTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from MemTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'MEMIntTrans
    tableName = "MEMIntTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from MemIntTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'MemIntPayable
    tableName = "MemIntPayable"
    SQL_Stmt = "Select AccID,TransID, Particulars from MemIntPayable where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "Member Transfer Done", , "Index Database"
    
    'CAMASTER
    tableName = "CAMaster"
    SQL_Stmt = "Select AccID,JointHolder,NomineeName,NomineeRelation from CAMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'CATrans
    tableName = "CATrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from CATrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'CAPLTrans
    tableName = "CAPLTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from CAPLTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'ChequeMaster
    tableName = "ChequeMaster"
    SQL_Stmt = "Select ChequeNo,SeriesNo, Particulars from ChequeMaster where Particulars <> '' AND NOT Particulars IS NULL "
    'If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'ChequeMaster
    tableName = "NoteTab"
    SQL_Stmt = "Select ModuleID,NoteID,AccId, Notes from NoteTab where Notes <> '' AND NOT Notes IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 3) Then GoTo EndLine
    
    'DepositLoanMaster
    tableName = "DepositLoanMaster"
    SQL_Stmt = "Select LoanID,PledgeDescription,Remarks from DepositLoanMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'DepositLoanTrans
    tableName = "DepositLoanTrans"
    SQL_Stmt = "Select LoanID,TransID, Particulars from DepositLoanTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'DepositLoanIntTrans
    tableName = "DepositLoanIntTrans"
    SQL_Stmt = "Select LoanID,TransID, Particulars from DepositLoanIntTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "Deposit Loan Transfer Done", , "Index Database"
    
    'FDMASTER
    tableName = "FDMaster"
    SQL_Stmt = "Select AccID,NomineeName,NomineeRelation from FDMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'FDTrans
    tableName = "FDTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from FDTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'MatFDTrans
    tableName = "MatFDTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from MatFDTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
        
    'FDIntTrans
    tableName = "FDIntTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from FDIntTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'FDIntPayable
    tableName = "FDIntPayable"
    SQL_Stmt = "Select AccID,TransID, Particulars from FDIntPayable where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "Fixed Deposits Transfer Done", , "Index Database"
    
    'LoanScheme
    tableName = "LoanScheme"
    SQL_Stmt = "Select SchemeID,SchemeName,LoanPurpose,Description from LoanScheme"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'LoanMaster
    tableName = "LoanMaster"
    SQL_Stmt = "Select SchemeID,LoanId, PledgeItem,LoanPurpose,OtherDets,Remarks from LoanMaster"
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'LoanTrans
    tableName = "LoanTrans"
    SQL_Stmt = "Select LoanID,TransID, Particulars from LoanTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'LoanIntTrans
    tableName = "LoanIntTrans"
    SQL_Stmt = "Select LoanID,TransID, Particulars from LoanIntTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'LoanIntReceivAble
    tableName = "LoanIntReceivAble"
    SQL_Stmt = "Select LoanID,TransID, Particulars from LoanIntReceivAble where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'LoanPurpose
    tableName = "LoanPurpose"
    SQL_Stmt = "Select PurposeID,Purpose from LoanPurpose where Purpose <> '' AND NOT Purpose IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'LoanAbnEp
    tableName = "LoanAbnEp"
    SQL_Stmt = "Select LoanID,BKCC, ABNDesc,EPDesc from LoanAbnEp "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "Loan Transfer Done", , "Index Database"
    
    'PDMaster
    tableName = "PDMaster"
    SQL_Stmt = "Select AccID,JointHolder,Nominee  from PDMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'AgentTrans
    tableName = "AgentTrans"
    SQL_Stmt = "Select AgentID,TransID, Particulars from AgentTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'PDTrans
    tableName = "PDTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from PDTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'PDIntTrans
    tableName = "PDIntTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from PDIntTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'PDIntPayable
    tableName = "PDIntTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from PDIntPayable where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'RDMaster
    tableName = "RDMaster"
    SQL_Stmt = "Select AccID,NomineeRelation  from RDMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
        
    'RDTrans
    tableName = "RDTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from RDTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'RDIntTrans
    tableName = "RDIntTrans"
    SQL_Stmt = "Select AccID,TransID, Particulars from RDIntTrans where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'RDIntPayable
    tableName = "RDIntPayable"
    SQL_Stmt = "Select AccID,TransID, Particulars from RDIntPayable where Particulars <> '' AND NOT Particulars IS NULL "
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    MsgBox "Pigmy & Recurring Transfer Done", , "Index Database"
    
    'DepositName
    tableName = "DepositName"
    SQL_Stmt = "Select DepositId,DepositName from DepositName"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'SuspAccount
    tableName = "SuspAccount"
    SQL_Stmt = "Select TransID,CustName from SuspAccount"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'DepositName
    tableName = "BankMaster"
    SQL_Stmt = "Select BankID,BankName,Manager,Address from BankMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'CompanyCreation
    tableName = "CompanyCreation"
    SQL_Stmt = "Select HeadID,CompanyName,ContactPerson,Address from CompanyCreation"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'GodownDet
    tableName = "GodownDet"
    SQL_Stmt = "Select GodownID,GodownName,ContactPerson,Address from GodownDet"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'ParentHeads
    tableName = "ParentHeads"
    SQL_Stmt = "Select ParentID,ParentName from ParentHeads"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'Heads
    tableName = "Heads"
    SQL_Stmt = "Select HeadID,HeadName from Heads"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    MsgBox "Head Transfer Done", , "Index Database"
    
    'ProductGroup
    tableName = "ProductGroup"
    SQL_Stmt = "Select GroupID,GroupName from ProductGroup"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'Products
    tableName = "Products"
    SQL_Stmt = "Select ProductID,ProductName from Products"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'TransParticulars
    tableName = "TransParticulars"
    SQL_Stmt = "Select TransID,Particulars from TransParticulars where Particulars <> '' AND Not Particulars is NULL"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'Units
    tableName = "Units"
    SQL_Stmt = "Select UnitID,UnitName  from Units"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
     
    'BankHeadIDs
    tableName = "BankHeadIDs"
    SQL_Stmt = "Select HeadId,HeadName,AliasName  from BankHeadIDs"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
     
    'CustomerType
    tableName = "CustomerType"
    SQL_Stmt = "Select CustType,CustTypeName from CustomerType"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
     
    'AccountGroup
    tableName = "AccountGroup"
    SQL_Stmt = "Select AccGroupID,GroupName from AccountGroup"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
     
    'ClearingTab
    tableName = "ClearingTab"
    SQL_Stmt = "Select ChequeID,BankName,Remarks,Place from ClearingTab"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
     
    'EmpDetails
    tableName = "EmpDetails"
    SQL_Stmt = "Select UserID,BankName from EmpDetails"
    'If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    MsgBox "Place & Caste Transfer Done", , "Index Database"
    
    'ShgMaster
    tableName = "ShgMaster"
    SQL_Stmt = "Select AccID,ContactPerson,MeetingDay,MeetingPlace,Caste,Place,Remarks from ShgMaster"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    'ShgTrans
    tableName = "ShgTrans"
    SQL_Stmt = "Select AccID,TransID, TrainingDetail,Place from ShgTrans"
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
     
    'AssetDetails
    tableName = "AssetDetails"
    SQL_Stmt = "Select CustomerID,AssetID,Place,DryLand,WellLand,CanalLand,RiverLand from AssetDetails"
    If Not UPdateRecords(SQL_Stmt, tableName, 2) Then GoTo EndLine
    
    'AssetDetails
    tableName = "SetUp"
    SQL_Stmt = "Select SetUpID,KeyData,ValueData from SetUp where ValueData <> '' and not VALueData is null"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
    dbUtils.CommitTrans
    
    MsgBox "Nudi to Suchita Transfer done", , "Index Database"
    dbUtils.BeginTrans
    'FarmerTypeTab
    tableName = "FarmerTypeTab"
    SQL_Stmt = "Select FarmerTypeID,TypeName  from FarmerTypeTab"
    If Not UPdateRecords(SQL_Stmt, tableName) Then GoTo EndLine
    
        
    dbUtils.CommitTrans
    MsgBox "ALL Nudi to Suchita Transfer done", , "Index Database"
    dbUtils.CloseDB
    Exit Sub

EndLine:
    MsgBox "Error in " & tableName
    dbUtils.RollBack
    dbUtils.CloseDB
    
End Sub
Private Function GetUpdateSQL(rst As Recordset, keyFieldCount As Integer, Optional maxCount As Integer = 1) As String
    Dim SqlStmt As String
    'Dim maxCount As Integer
    Dim Count As Integer
    Dim fldValue As String
    
    SqlStmt = ""
    
    If maxCount < 2 Then maxCount = rst.Fields.Count
        
    For Count = keyFieldCount To maxCount - 1
        fldValue = FormatField(rst(Count))
        If Len(fldValue) > 0 Then _
            SqlStmt = SqlStmt & rst.Fields(Count).Name & " = " & AddQuotes(Trim$(ConvertNudiToSuchita(fldValue)), True) & ","
    Next Count
    
    If Len(SqlStmt) > 0 Then
        SqlStmt = " SET " & Left(SqlStmt, Len(SqlStmt) - 1) & " WHERE " & rst.Fields(0).Name & "=" & FormatField(rst(0))
        For Count = 1 To keyFieldCount - 1
            SqlStmt = SqlStmt & " AND " & rst.Fields(Count).Name & "=" & FormatField(rst(Count))
        Next
    End If
    
    GetUpdateSQL = SqlStmt
End Function
Private Function UpdateNameRecords() As Boolean
    UpdateNameRecords = False
    Dim SQL_Stmt As String
    SQL_Stmt = "Select CustomerID,Title,FirstName,MiddleName,LastName," & _
        " Profession,Caste,HomeAddress,OfficeAddress,Place,IsciName,FullName from NameTab"
    
    dbUtils.SqlStmt = SQL_Stmt
    'Get the Record Set
    Dim isciName As String
    Dim recCount As Long
    recCount = dbUtils.Fetch(rst, adOpenDynamic)
    If recCount > 0 Then
            While rst.EOF = False
            'Update the Record
            SQL_Stmt = GetUpdateSQL(rst, 1, rst.Fields.Count - 2)
            If Len(SQL_Stmt) > 0 Then
                ''Remove the Where Part
                SQL_Stmt = Left(SQL_Stmt, InStr(1, SQL_Stmt, " Where", vbTextCompare))
            
                'Convert the Isci Name and FullName
                isciName = FormatField(rst("FirstName")) + " " + FormatField(rst("LastName"))
                isciName = ConvertNudiToSuchita(isciName)
                
                ''Add the EnglishName
                If Len(FormatField(rst("FullName"))) < 1 Then _
                    SQL_Stmt = SQL_Stmt & ", FullName = " & AddQuotes(ConvertToEnglish(isciName), True)
                
                SQL_Stmt = SQL_Stmt & ", IsciName = " & AddQuotes(SuchiToIscii(Left(isciName, 20), 7), True)
                
                ''AddBack the Where Caluse
                dbUtils.SqlStmt = "Update NameTab " & SQL_Stmt & " Where CustomerID = " & rst(0)
                If Not dbUtils.SQLExecute Then
                    MsgBox "Error in update"
                    Exit Function
                End If
            Else
                'MsgBox "Check"
            End If
            'Move to next Record
            rst.MoveNext
        Wend

    ElseIf recCount = -10 Then
        UpdateNameRecords = False
        Exit Function
    End If
    
    UpdateNameRecords = True
End Function


Private Function UPdateRecords(SqlStmt As String, tableName As String, Optional keyFieldCount As Integer = 1) As Boolean
    UPdateRecords = False
    dbUtils.SqlStmt = SqlStmt
    'Get the Record Set
    Dim SQL_Stmt As String
    Dim recCount As Long
    recCount = dbUtils.Fetch(rst, adOpenDynamic)
    If recCount > 0 Then
            While rst.EOF = False
            'Update the Record
            SQL_Stmt = GetUpdateSQL(rst, keyFieldCount)
            If Len(SQL_Stmt) > 0 Then
                dbUtils.SqlStmt = "Update " & tableName & SQL_Stmt
                If Not dbUtils.SQLExecute Then
                    MsgBox "Error in update"
                    Exit Function
                End If
            Else
                'MsgBox "Check"
            End If
            'Move to next Record
            rst.MoveNext
        Wend

    ElseIf recCount = -10 Then
        UPdateRecords = False
        Exit Function
    End If
    
    UPdateRecords = True
End Function



