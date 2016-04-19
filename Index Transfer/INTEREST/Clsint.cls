VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_ModuleID As wis_AccountModule
Dim m_SchemeName As String

Dim m_InterestRate() As Single
Dim m_NoOfInterest As Integer
Dim M_Date() As Date
Dim M_NoOfDate As Integer
Dim m_PresentInterest As Integer
Dim m_PresentDate As Integer

Public Function UndoLastInterest(ModuleID As wis_AccountModule, SchemeName As String, InterestRate As Single, _
                                              IndianStartDate As String) As Boolean
On Error GoTo ErrLine
'undo the last transaction
Dim TransID As Long

'get the max transid for the particular Moduleid
'fire SQL
gDbTrans.SqlStmt = "SELECT MAX (TransID) from InterestTab  where ModuleID =" & ModuleID & _
                                " AND SchemeName = " & AddQuotes(SchemeName, True)
If gDbTrans.SQLFetch < 1 Then GoTo ErrLine
TransID = Val(FormatField(gDbTrans.Rst(0)))

gDbTrans.SqlStmt = "Delete * from InterestTab where transid  = " & _
                                TransID & " AND  ModuleID =" & ModuleID & _
                                " AND SchemeName = " & AddQuotes(SchemeName, True)
gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then
      gDbTrans.RollBack
      GoTo ErrLine
 End If
 
 gDbTrans.CommitTrans
 
gDbTrans.SqlStmt = "SELECT MAX (TransID) from InterestTab  where ModuleID =" & ModuleID & _
                                " AND SchemeName = " & AddQuotes(SchemeName, True)
If gDbTrans.SQLFetch < 1 Then GoTo ExitLine

TransID = Val(FormatField(gDbTrans.Rst(0)))

If TransID > 100 Then
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Update InterestTab Set EndDate = " & Null & _
                    " Where TransId = " & TransID & " And ModuleID =" & ModuleID & _
                    " AND SchemeName = " & AddQuotes(SchemeName, True)
    
    If Not gDbTrans.SQLExecute Then
         gDbTrans.RollBack
         GoTo ErrLine
    End If
    
    gDbTrans.CommitTrans
End If

ExitLine:
UndoLastInterest = True
Exit Function

ErrLine:
    MsgBox "Error In Undo Interest Class", vbExclamation, gAppName & " - Error"
    'MsgBox LoadResString(gLangOffSet + 533), vbExclamation, gAppName & " - Error"
End Function





Public Sub ClearInterest()
   m_NoOfInterest = 0
   M_NoOfDate = 0
   m_PresentInterest = 0
   m_PresentDate = 0
   ReDim m_InterestRate(0)
   ReDim M_Date(0)
End Sub


'This Function Will give the Rate of interest between the Period specified
'It Also set the the interest Rate changed Period and Number of Changes
'It Also sets the Next Interest Date

Public Function InterestRate(ModuleID As wis_AccountModule, SchemeName As String, FromDate As Date, _
                                            Optional ToDate As Date) As Single
      Err.Clear
      On Error GoTo ErrLine
      m_ModuleID = ModuleID
      m_SchemeName = SchemeName
      
      If m_ModuleID = 0 Then
            Err.Raise 5002, "Interest Class", "ModuleId not set"
      End If
      If m_SchemeName = "" Then
            Err.Raise 5003, "Interest Class", "SchemeName not set"
      End If
      m_NoOfInterest = 0
      m_PresentInterest = 0
      M_NoOfDate = 0
      If Trim$(ToDate) = "" Then
         ToDate = gStrDate
      End If
      
      Dim Retval As Integer
      ReDim m_InterestRate(0)
      ReDim M_Date(0)
   'First Get The Rate Of Interest As On FromindianDate
   gDbTrans.SqlStmt = "Select Top 1 * from InterestTab Where StartDate <= " & _
            " #" & FromDate & "#" & _
            " And ModuleId = " & m_ModuleID & _
            " And SchemeName = " & AddQuotes(m_SchemeName, True) & _
            " AND BankID = " & gBankID & " Order by TransId Desc "
   
   If gDbTrans.SQLFetch < 1 Then
         gDbTrans.SqlStmt = "Select Top 1 * from InterestTab Where " & _
                " ModuleId = " & m_ModuleID & _
                " And SchemeName = " & AddQuotes(m_SchemeName, True) & _
                " AND BankID = " & gBankID & " Order by StartDate "
        If gDbTrans.SQLFetch < 1 Then GoTo ExitLine
   End If
   
   m_NoOfInterest = 1
   M_NoOfDate = 1
   m_InterestRate(0) = FormatField(gDbTrans.Rst("InterestRate"))
      'Check Whether Interest between Given Period
      gDbTrans.SqlStmt = "Select StartDate,EndDate,InterestRate from InterestTab " & _
            " Where ModuleId = " & m_ModuleID & _
            " And SchemeName = " & AddQuotes(m_SchemeName, True) & _
            " And StartDate > #" & FromDate & "# " & _
            " And StartDate < #" & DateAdd("d", 1, ToDate) & "#" & _
            " AND BankID = " & gBankID
      
      Retval = gDbTrans.SQLFetch
      m_NoOfInterest = m_NoOfInterest + Retval
      M_NoOfDate = M_NoOfDate + Retval
      
      m_PresentInterest = 1
      
      ReDim Preserve m_InterestRate(Retval)
      ReDim Preserve M_Date(Retval)
      
      Dim Rst As Recordset
      Set Rst = gDbTrans.Rst.Clone
      
      Retval = 1
      While Not Rst.EOF   'formatfield(Rst("
            m_InterestRate(Retval) = FormatField(Rst("Interestrate"))
            M_Date(Retval) = FormatField(Rst("StartDate"))
            Retval = Retval + 1
         Rst.MoveNext
      Wend
InterestRate = m_InterestRate(0)

m_PresentDate = LBound(M_Date)
M_NoOfDate = UBound(M_Date)

m_PresentInterest = LBound(m_InterestRate)
m_NoOfInterest = UBound(m_InterestRate)

ExitLine:
      Exit Function
      
ErrLine:
   If Err Then
         Debug.Assert 1 = 1
         MsgBox Err.Number & vbCrLf & "Error : " & Err.Description, vbExclamation, wis_MESSAGE_TITLE
    End If
   'Resume
End Function
Public Function LastInterestRate() As Integer

'Gives The Last Interest Added On Given Two Periods
'use InterestRate function Before Using This


If m_NoOfInterest > 0 Then
   LastInterestRate = m_NoOfInterest
   Exit Function
End If

LastInterestRate = 0
End Function
Public Property Get SchemeName() As Variant
SchemeName = m_SchemeName
End Property
Public Property Get ModuleID() As Variant
ModuleID = m_ModuleID
End Property

Public Property Let ModuleID(ByVal ModuleID As Variant)
m_ModuleID = ModuleID
End Property
Public Property Let SchemeName(ByVal SchemeName As Variant)
m_SchemeName = SchemeName
End Property

'This Function returnts the NextInterestDate in 'mm/dd/yyyy' Format between specified date
'Before Calling This Function You must call InterestRate functin
'which Sets the Array of INterest rates
'And Array of Dates when InterestRate Has been Changed
'If there is  no next date then it returns the NULL String

Public Function NextInterestDate() As Date
      NextInterestDate = ""
      'Check Whether before calling this Function The InterestRate Function Has called or not
      If m_NoOfInterest < 0 Then
         Err.Raise 5002, "Interest Class", "Invalid procedure Call"
      End If
      If m_PresentDate = M_NoOfDate Then GoTo ExitLine
      
   'Returnt the Next Interest Rate from  Interestrate array
      m_PresentDate = m_PresentDate + 1
      NextInterestDate = m_InterestRate(m_PresentDate)
      NextInterestDate = M_Date(m_PresentDate)
ExitLine:


End Function

'This Function returnts the NextInterestRate in Indaina Format between specified date
'Before Calling This Function You must call InterestRate functin which Sets the Array of INterest rates
'If there is  no next interestrate it returns the then it returns 0
Public Function NextInterestRate() As Single
      If m_PresentInterest = m_NoOfInterest Then GoTo ExitLine
   'Returnt the Next Interest Rate from  Interestrate array
      m_PresentInterest = m_PresentInterest + 1
      NextInterestRate = m_InterestRate(m_PresentInterest)
      
ExitLine:
End Function

'This Function will write the interest rate to the Interest Tab
'And Interst Rate will be in % not in fraction , And Fraction will convert to %
Public Function SaveInterest(ModuleID As wis_AccountModule, SchemeName As String, _
                        InterestRate As Single, Optional AsOnDate As Date) As Boolean
      
On Error GoTo ErrLine

If ModuleID = 0 Then GoTo ErrLine
If SchemeName = "" Then GoTo ErrLine
If InterestRate <= 0 Then Exit Function

If IsMissing(AsOnDate) Then
   AsOnDate = gStrDate
ElseIf Trim$(AsOnDate) = "" Then
      AsOnDate = gStrDate
End If
If InterestRate <= 1 Then InterestRate = InterestRate * 100

Dim TransID As Long
Dim LastIntRate As Single
Dim InTrans As Boolean

gDbTrans.SqlStmt = "Select * From InterestTab " & _
      " Where SchemeName = " & AddQuotes(SchemeName, True) & _
      " And ModuleId = " & ModuleID & " AND BankID = " & gBankID
  
If gDbTrans.SQLFetch > 0 Then
   TransID = FormatField(gDbTrans.Rst("Transid"))
   'Check For the Date if he is Entering the Date earlier than
   'the already transacted date
  If DateDiff("d", gDbTrans.Rst("StartDate"), AsOnDate) < 0 Then
      GoTo ExitLine
  End If
End If

gDbTrans.BeginTrans
InTrans = True
If TransID >= 100 Then 'If There Are Transactions Prior To This
      'Check Whether the Last Interest rate & updating Interest rate or same
      'If Last Interest rate & Updating interst rate or same then exit the function
       If Val(FormatField(gDbTrans.Rst("InterestRate"))) = InterestRate Then
          SaveInterest = True
          GoTo ExitLine
      End If
      
      'Check for the Date if he is updating II time on the same day
       'then delete the previous transaction of the day
      If FormatField(gDbTrans.Rst("StartDate")) = AsOnDate Then
               gDbTrans.SqlStmt = "Delete * From InterestTab Where TransID = " & _
                        " (Select Max(TransId) From InterestTab " & _
                        " Where SchemeName = " & AddQuotes(SchemeName, True) & _
                        " And ModuleId = " & ModuleID & " And EndDate = NULL)" & _
                      " AND BankID = " & gBankID
                     
               If Not gDbTrans.SQLExecute Then GoTo ExitLine
               TransID = TransID - 1
      End If
      
      'First UpDate the DataBaseTab For The Interest Date End
       gDbTrans.SqlStmt = "Update InterestTab Set  EndDate = #" & DateAdd("d", -1, AsOnDate) & "# " & _
            " Where SchemeName = " & AddQuotes(SchemeName, True) & _
            " And ModuleId = " & ModuleID & " And TransID = " & TransID & _
            " AND BankID = " & gBankID
                     
      If Not gDbTrans.SQLExecute Then GoTo ExitLine
      TransID = TransID + 1

Else 'first ever transaction
      TransID = 100
End If

'now insert the Present interest rate effective from GivenDate
gDbTrans.SqlStmt = "Insert Into InterestTab(TransID,ModuleId,SchemeName," & _
          " Interestrate,StartDate) Values " & _
          "( " & TransID & ", " & ModuleID & ", " & _
          AddQuotes(SchemeName, True) & "," & _
          InterestRate & ", " & _
          " #" & AsOnDate & "#" & _
          " AND BankID = " & gBankID & ")"
         If Not gDbTrans.SQLExecute Then GoTo ExitLine
            
gDbTrans.CommitTrans
SaveInterest = True
InTrans = False

ExitLine:
If InTrans Then gDbTrans.RollBack
Exit Function

ErrLine:
      MsgBox "Error In SaveInterest - Interest Class", , wis_MESSAGE_TITLE
End Function
