Attribute VB_Name = "basMAin"
Option Explicit

Public gLangOffSet As Integer
Public Const wis_KannadaOffset = 5000

Public Sub Main()
frmMain.Show 1

End Sub


Sub RepairOldDB(DbName As String)

On Error GoTo Err_Line
Dim OldTrans As clsOldUtils
Set OldTrans = New clsOldUtils
Dim OldRst As ADODB.Recordset
Dim TmpRst As ADODB.Recordset

Dim oldCOunt As Integer
Dim newCOunt As Integer

Dim AccID As Long
Dim MaxAccID As Long
Dim PrgVal As Long

If Not OldTrans.OpenDB(DbName, OldPwd) Then
    If MsgBox("Try for new password", vbYesNo, "New PWD") = vbNo Then End
    OldPwd = "WIS!@#"
    If Not OldTrans.OpenDB(DbName, OldPwd) Then
        MsgBox "Invalid DataBasename"
        End
    End If
    'OldPwd = NewPwd
End If

Dim SqlStr As String
Dim SngSpace As String

'Get the Languale Offset
OldTrans.SQLStmt = "SELECT * From Install Where Keydata = 'Language'"
If OldTrans.Fetch(OldRst, adOpenDynamic) > 0 Then
    If UCase(FormatField(OldRst("ValueData"))) = "KANNADA" Then gLangOffSet = 5000
End If

SngSpace = ""

With frmMain
    .lblProgress = "Correcting the old database"
    .prg.Max = 50
    .Refresh
End With

Screen.MousePointer = vbHourglass

'''FIRST ADJUST THE VALUES OF NAME TAB IF ANY WRONG ENTRIES ARE IN THE dATA DABASE
SqlStr = "UPDATE NameTab Set DOB = #1/1/100# WHERE DOB is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

SqlStr = "UPDATE NameTab Set Title = '" & SngSpace & "' WHERE Title is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "Correcting the member details"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'''NOW HECK THE MEMBER TABLE
With frmMain
    .lblProgress = "Correcting the member details"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "UPDATE MMMASTER SET Introduced = 0 WHERE Introduced is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

SqlStr = "UPDATE MMMASTER SET ModifiedDate = #1/1/100# WHERE ModifiedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "Correcting the member details"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Update Closed date
SqlStr = "UPDATE MMMASTER Set ClosedDate = #1/1/100# WHEre ClosedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "Correcting the member details"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Update Nominee
SqlStr = "UPDATE MMMASTER set Nominee = '" & SngSpace & "' WHERE Nominee is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "Correcting the member details"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With



'Delete The transaction Details which are not entered in MAster record
SqlStr = "Delete * FROM MMTrans WHere AccID Not In " & _
    "(Select AccID FROM MMMaster )"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "Correcting the member details"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Delete The Memebr Master Details are entere without any share transaction
SqlStr = "Delete * FROM MMMaster WHere AccID Not In " & _
    "(Select Distinct AccID FROM MMTrans)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
'Delete The Member Transaction Details which are without any MemeDetails
SqlStr = "Delete * FROM MMTrans Where AccID Not In " & _
    "(Select Distinct AccID FROM MMMaster)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

SqlStr = "Delete * FROM MMTrans WHere " & _
    "(TransID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the member Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
End With

'NOW CHECK THE bkCC TABLE
SqlStr = "Select MAx(LoanID) From BKCCTrans"
OldTrans.SQLStmt = SqlStr
Call OldTrans.Fetch(TmpRst, adOpenForwardOnly)
MaxAccID = FormatField(TmpRst(0))
AccID = 0

'For every 100 recs we commit instead of every txn.
'this will speed up things like mad.
'Dim commitInterval As Integer
Const commitInterval = 100
Dim loopCount As Integer
Dim commitPending As Boolean
commitPending = False
loopCount = 0
Do
    If loopCount Mod commitInterval = 0 Then
      If commitPending Then
        OldTrans.CommitTrans
      End If
      OldTrans.BeginTrans
    End If
    loopCount = loopCount + 1
    
    If AccID > MaxAccID Then Exit Do
    'Delete The transaction Details which are not entered in MAster record
    SqlStr = "Delete * FROM BKCCTrans Where " & _
        " LoanID >= " & AccID & " And LoanID < " & AccID + 100 & _
        " ANd LoanID Not In (Select LoanID FROM LoanMaster WHERE " & _
            " LoanID >= " & AccID & " And LoanID < " & AccID + 100 & ")"
    OldTrans.SQLStmt = SqlStr
    
    If Not OldTrans.SQLExecute Then
        OldTrans.RollBack
    End If
    commitPending = True
    
    With frmMain
        .lblProgress = "COrrecting the BKCC Detals"
        PrgVal = PrgVal + 1
        .prg.Value = PrgVal
    End With
    AccID = AccID + 100
Loop
    
'commit any pending records in the last batch
' which could have escaped because of MOD.
If commitPending Then
    OldTrans.CommitTrans
End If

SqlStr = "Delete * FROM BKCCTrans Where " & _
    "(TransID is NULL) OR (LoanID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the BKCC Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
End With


'NOW CHECK THE LOANS TABLE

'Delete The transaction Details which are not exist in Master record
SqlStr = "Select MAx(LoanID) From LoanTrans"
OldTrans.SQLStmt = SqlStr
Call OldTrans.Fetch(TmpRst, adOpenForwardOnly)
MaxAccID = FormatField(TmpRst(0))
AccID = 0

' reset txn monitoring variables
loopCount = 0
commitPending = False

Do
    If loopCount Mod commitInterval = 0 Then
      If commitPending Then
        OldTrans.CommitTrans
      End If
      OldTrans.BeginTrans
    End If
    loopCount = loopCount + 1

    If AccID > MaxAccID Then Exit Do
    SqlStr = "Delete * FROM LoanTrans WHere " & _
        " LoanID >= " & AccID & " And LoanID < " & AccID + 200 & _
        " And LoanID Not In (Select LoanID FROM LoanMaster WHERE " & _
            " LoanID >= " & AccID & " And LoanID < " & AccID + 200 & ")"
    OldTrans.SQLStmt = SqlStr
    
    'OldTrans.BeginTrans
    If Not OldTrans.SQLExecute Then
        OldTrans.RollBack
    End If
    commitPending = True
    
    With frmMain
        .lblProgress = "COrrecting the Loan Detals"
        PrgVal = PrgVal + 1
        .prg.Value = PrgVal
    End With
    AccID = AccID + 200
Loop
'commit for last batch at the end of loop
If commitPending Then
    OldTrans.CommitTrans
End If

SqlStr = "Delete * FROM LoanTrans Where " & _
    "(TransID is NULL) OR (LoanID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "Correcting the Laon Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With


'Set the Instalment mod as 0 where
'installmet mode is weekly,Fortnightly
SqlStr = "UPDate LoanMaster Set InstalmentMode = 0,InstalmentAmt = 0 " & _
    "Where (InstalmentMode is NULL) OR (InstalmentMode =1 )" & _
    " OR (InstalmentMode =2 ) OR (InstalmentMode = 7 )"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If


'NOW CHECK THE SB TABLE

'Delete The transaction Details which are not entered in MAster record
SqlStr = "Delete * FROM SBTrans WHere AccID Not In " & _
    " (Select ACCID FROM SBMaster )"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the SB Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Delete The Master Details whise transaction has not done
SqlStr = "Delete * FROm SBMAster WHere AccID Not In " & _
    " (Select Distinct ACCID FROM SBTrans )"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the SB Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "Delete * FROM SBTrans WHere " & _
    "(TransID is NULL) OR (AccID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
SqlStr = "UPDATE SBMaster Set Nominee = ' ' WHERE Nominee is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "COrrecting the SB Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'NOW CHECK THE CA TABLE

    'Delete The transaction Details which are not entered in MAster record
    SqlStr = "Delete * FROm CATrans WHere AccID Not In " & _
        " (Select ACCID FROM CAMaster )"
    OldTrans.SQLStmt = SqlStr
    OldTrans.BeginTrans
    If Not OldTrans.SQLExecute Then
        OldTrans.RollBack
    Else
        OldTrans.CommitTrans
    End If
With frmMain
    .lblProgress = "Correcting the CA Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

    'Delete The Master Details whise transaction has not done
    SqlStr = "Delete * FROM CAMAster WHere AccID Not In " & _
        " (Select Distinct ACCID FROM CATrans )"
    OldTrans.SQLStmt = SqlStr
    OldTrans.BeginTrans
    If Not OldTrans.SQLExecute Then
        OldTrans.RollBack
    Else
        OldTrans.CommitTrans
    End If
With frmMain
    .lblProgress = "Correcting the CA Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With
SqlStr = "Delete * FROM CATrans WHere " & _
    "(TransID is NULL) OR (AccID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the CA Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "UPDATE CAMASTER SET ModifiedDate = #1/1/100# WHERE ModifiedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

'Update Closed date
SqlStr = "UPDATE CAMASTER Set ClosedDate = #1/1/100# WHEre ClosedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the CA Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Update Nominee
SngSpace = ""
SqlStr = "UPDATE CAMASTER set Nominee = '" & SngSpace & "' WHERE Nominee is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

SqlStr = "UPDATE CAMASTER set JointHolder= '" & SngSpace & "' WHERE JointHolder is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the CA Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "UPDATE CAMASTER set Nominee= '" & SngSpace & "' WHERE Nominee is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "Correcting the CA Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'NOW CHECK THE fD TABLE

'Get the No Of Deposits
SqlStr = "Select Distinct DepositID From FDMAster"
OldTrans.SQLStmt = SqlStr
Set OldRst = Nothing
Call OldTrans.Fetch(OldRst, adOpenDynamic)
If Not OldRst Is Nothing Then
    'Now Check For the No Of Deposits in the FD Trans
    SqlStr = "Select Distinct DepositID From FDTrans"
    OldTrans.SQLStmt = SqlStr
    Set TmpRst = Nothing
    Call OldTrans.Fetch(TmpRst, adOpenDynamic) '> OldRst.RecordCount Then _
        Set OldRst = TmpRst
    If TmpRst.RecordCount > OldRst.RecordCount Then Set OldRst = TmpRst
    
    'RV: 2013-JUN-16
    'Reset txn tracking variables.
    commitPending = False
    loopCount = 0
    While Not OldRst.EOF
        If loopCount Mod commitInterval = 0 Then
            If commitPending Then
                OldTrans.CommitTrans
            End If
            OldTrans.BeginTrans
        End If
    loopCount = loopCount + 1

        AccID = 0
        OldTrans.SQLStmt = "SELECT MAx(AccID) From FDMaster " & _
                "WHERE DepositID = " & OldRst(0)
        Call OldTrans.Fetch(TmpRst, adOpenDynamic)
        MaxAccID = FormatField(TmpRst(0))
        OldTrans.SQLStmt = "SELECT MAx(AccID) From FDTrans " & _
                "WHERE DepositID = " & OldRst(0)
        Call OldTrans.Fetch(TmpRst, adOpenDynamic)
        MaxAccID = IIf(MaxAccID > FormatField(TmpRst(0)), MaxAccID, FormatField(TmpRst(0)))
        'Delete The transaction Details which are not entered in MAster record
        
        'RV: 2013-Jun-16 -- removing commit from inside the loop below.
        ' so that it can speed up the operation.
        OldTrans.BeginTrans
        Do
            SqlStr = "Delete * FROM FDTrans Where DepositID= " & OldRst(0) & _
                " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & _
                " AND AccID Not In (Select ACCID FROM FDMaster WHERE DepositID = " & OldRst(0) & _
                " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & ")"
            OldTrans.SQLStmt = SqlStr
            'OldTrans.BeginTrans
            If Not OldTrans.SQLExecute Then
                OldTrans.RollBack
            'Else
            '    OldTrans.CommitTrans
            End If
            With frmMain
                .lblProgress = "Correcting the FD Detals"
                PrgVal = PrgVal + 1
                .prg.Value = PrgVal
            End With
        
            'Delete The Master Details whise transaction has not done
            SqlStr = "Delete * FROM FDMaster Where DepositID= " & OldRst(0) & _
                " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & _
                " AND AccID Not In (Select Distinct ACCID FROM FDTrans WHERE DepositID = " & OldRst(0) & _
                " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & ")"
            OldTrans.SQLStmt = SqlStr
            'OldTrans.BeginTrans
            If Not OldTrans.SQLExecute Then
                OldTrans.RollBack
            'Else
            '    OldTrans.CommitTrans
            End If
            With frmMain
                .lblProgress = "COrrecting the FD Detals"
                PrgVal = PrgVal + 1
                .prg.Value = PrgVal
            End With
            
            AccID = AccID + 500
            If AccID > MaxAccID Then Exit Do
            
            'commitPending = True
        Loop
        ' RV: commit the txns performed in the above loop.
        OldTrans.CommitTrans
        
        OldRst.MoveNext
    Wend
    
    'final commit for any pending records in the last batch.
    If commitPending Then
        OldTrans.CommitTrans
    End If
End If

SqlStr = "Delete * FROM FDTrans WHere DepositID is NULL " & _
    "OR (TransID is NULL) OR (AccID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "COrrecting the FD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'NOW CHECK THE DL TABLE

    'Get the No Of Deposits
    SqlStr = "Select Distinct DepositID From DLMAster"
    OldTrans.SQLStmt = SqlStr
    Set OldRst = Nothing
    Call OldTrans.Fetch(OldRst, adOpenDynamic)
    
If Not OldRst Is Nothing Then
    'Now Check For the No Of Deposits in the DL Trans
    SqlStr = "Select Distinct DepositID From DLTrans"
    OldTrans.SQLStmt = SqlStr
    Set TmpRst = Nothing
    Call OldTrans.Fetch(TmpRst, adOpenDynamic) '> OldRst.RecordCount Then _
        Set OldRst = TmpRst
    If TmpRst.RecordCount > OldRst.RecordCount Then Set OldRst = TmpRst

    While Not OldRst.EOF
        'RV: 2013-Jun-16
        'reset txn tracking variables
        loopCount = 0
        commitPending = False
            If loopCount Mod commitInterval = 0 Then
                If commitPending Then
                    OldTrans.CommitTrans
                End If
                OldTrans.BeginTrans
        End If
        loopCount = loopCount + 1
        commitPending = True

        AccID = 0
        OldTrans.SQLStmt = "SELECT MAx(AccID) From DLMaster " & _
            " WHERE DepositID = " & OldRst(0)
        Call OldTrans.Fetch(TmpRst, adOpenDynamic)
        MaxAccID = FormatField(TmpRst(0))
        OldTrans.SQLStmt = "SELECT MAx(AccID) From DLTrans " & _
            " WHERE DepositID = " & OldRst(0)
        Call OldTrans.Fetch(TmpRst, adOpenDynamic)
        MaxAccID = IIf(MaxAccID > FormatField(TmpRst(0)), MaxAccID, FormatField(TmpRst(0)))
'        If AccID = 0 Then
            AccID = 0
            'Delete The transaction Details which are not entered in MAster record
            
            'RV: 2013-Jun-16
            ' Begin a txn for the below loop
            OldTrans.BeginTrans
            Do
                
                'Delete The transaction Details which are not entered in MAster record
                SqlStr = "Delete * FROM DLTrans Where DepositID= " & OldRst(0) & _
                    " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & _
                    " AND AccID Not In (Select ACCID FROM DLMaster WHERE DepositID = " & OldRst(0) & _
                        " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & ")"
                OldTrans.SQLStmt = SqlStr

                'OldTrans.BeginTrans
                If Not OldTrans.SQLExecute Then
                    OldTrans.RollBack
                'Else
                '    OldTrans.CommitTrans
                End If
                
                'Delete The Master Details which has no transaction
                SqlStr = "Delete * FROM DLMaster Where DepositID= " & OldRst(0) & _
                    " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & _
                    " AND AccID Not In (Select Distinct ACCID FROM DLTrans WHERE DepositID = " & OldRst(0) & _
                        " AND AccID >= " & AccID & " AND ACCID < " & AccID + 500 & ")"
                OldTrans.SQLStmt = SqlStr
                OldTrans.BeginTrans
                If Not OldTrans.SQLExecute Then
                    OldTrans.RollBack
                'Else
                '    OldTrans.CommitTrans
                End If
                With frmMain
                    .lblProgress = "Correcting the DL Detals"
                    PrgVal = PrgVal + 1
                    .prg.Value = PrgVal
                End With
                
                AccID = AccID + 500
                If AccID > MaxAccID Then Exit Do
            Loop
            'RV: commit the txns done in above loop
            OldTrans.CommitTrans
 '       End If
        
        OldRst.MoveNext
    Wend
    
    If commitPending Then
        OldTrans.CommitTrans
    End If

End If

SqlStr = "Delete * FROM DLTrans WHere DepositID Is NULL " & _
    "Or (TransID is NULL) OR (AccID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
    With frmMain
        .lblProgress = "Correcting the DL Detals"
        PrgVal = PrgVal + 1
        .prg.Value = PrgVal
        .Refresh
    End With

'NOW CHECK THE rd TABLE

SqlStr = "UPDATE RDMASTER SET ModifiedDate = #1/1/100# WHERE ModifiedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If


With frmMain
    .lblProgress = "COrrecting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With


'Update Closed date
SqlStr = "UPDATE RDMASTER Set ClosedDate = #1/1/100# WHEre ClosedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If


With frmMain
    .lblProgress = "COrrecting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With
'Update Nominee
SqlStr = "UPDATE RDMASTER set Nominee = '" & SngSpace & "' WHERE Nominee is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "COrrecting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "UPDATE RDMASTER set JointHolder= '" & SngSpace & "' WHERE JointHolder is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If


With frmMain
    .lblProgress = "COrrecting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Delete The transaction Details which are not entered in MAster record
SqlStr = "Delete * FROM RDTrans WHere AccID Not In " & _
    "(Select ACCID FROM RDMaster)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "Correcting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With


'Delete The Master Details whise transaction has not done
SqlStr = "Delete * From RDMAster WHere AccID Not In " & _
    " (Select Distinct ACCID FROM RDTrans )"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "Correcting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "Delete * FROM RDTrans Where " & _
    "(TransID is NULL) OR (AccID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "Correcting the RD Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'NOW CHECK THE PIGMY TABLE

'Before Fetching Update the Values
'where It can be Null with default value
'Then Fetch the records
With frmMain
    .prg.Value = 1
    .Refresh
End With

If MsgBox("Do you want to repair the Pigmy details?", vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then GoTo EndPigmy
PrgVal = 0

With frmMain
    .lblProgress = "Correcting the details of pigmy agents"
        .prg.Max = 100

    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Delete the any details in the PDMAster Whose transactions has not done
SqlStr = "SELECT Distinct UserID FROM PDMASTER"
OldTrans.SQLStmt = SqlStr
Set OldRst = Nothing
'If OldTrans.Fetch(OldRst, adOpenDynamic) < 1 Then GoTo Exit_Line
If OldTrans.Fetch(OldRst, adOpenDynamic) > 0 Then
    With frmMain
        .lblProgress = "COrrecting the Pigmy Detals"
        PrgVal = PrgVal + 1
        .prg.Value = PrgVal
        .Refresh
    End With

    
    'RV: 2013-Jun-16
    'reset txn tracking variables
    loopCount = 0
    commitPending = False
    
    While Not OldRst.EOF
    
        If loopCount Mod commitInterval = 0 Then
            If commitPending Then
                OldTrans.CommitTrans
            End If
            OldTrans.BeginTrans
        End If
        loopCount = loopCount + 1
        commitPending = True

        SqlStr = "Select Max(AccID) From PDTrans Where UserID = " & OldRst(0)
        OldTrans.SQLStmt = SqlStr
        Call OldTrans.Fetch(TmpRst, adOpenDynamic)
        MaxAccID = FormatField(TmpRst(0))
        If MaxAccID = 0 Then
            SqlStr = "Select Max(AccID) From PDMAster Where UserID = " & OldRst(0)
            OldTrans.SQLStmt = SqlStr
            Call OldTrans.Fetch(TmpRst, adOpenDynamic)
            MaxAccID = FormatField(TmpRst(0))
        End If
        AccID = 0
        'RV: 2013-Jun-16
        ' begin a txn for the operations in the below loop.
        OldTrans.BeginTrans
        Do
            SqlStr = "DELETE * FROM PDTRans WHERE USERID = " & OldRst(0) & _
                " AND AccID > " & AccID & " AND AccID <= " & AccID + 200 & _
                " AND AccId NOT IN (SELECT Distinct AccID From PDMaster " & _
                    " WHERE AccID >= " & AccID & " ANd AccID <= " & AccID + 200 & _
                    " AND USERID = " & OldRst(0) & ")"
            OldTrans.SQLStmt = SqlStr
            'OldTrans.BeginTrans
            If Not OldTrans.SQLExecute Then
                OldTrans.RollBack
            'Else
            '    OldTrans.CommitTrans
            End If
            With frmMain
                .lblProgress = "Correcting the Pigmy Detals"
                PrgVal = PrgVal + 1
                .prg.Value = PrgVal
            End With
            
            SqlStr = "DELETE * FROM PDMaster WHERE USERID = " & OldRst(0) & _
                " AND AccID > " & AccID & " AND AccID <= " & AccID + 200 & _
                " AND AccId NOT IN (SELECT Distinct AccID From PDTrans " & _
                    " WHERE AccID >= " & AccID & " ANd AccID <= " & AccID + 200 & _
                    " AND  USERID = " & OldRst(0) & ")"
            OldTrans.SQLStmt = SqlStr
            OldTrans.BeginTrans
            If Not OldTrans.SQLExecute Then
                OldTrans.RollBack
            'Else
            '    OldTrans.CommitTrans
            End If
            
            With frmMain
                .lblProgress = "Correcting the Pigmy Detals"
                PrgVal = PrgVal + 1
                .prg.Value = PrgVal
                .Refresh
            End With
            AccID = AccID + 200
            If AccID > MaxAccID Then Exit Do
        Loop
        'RV: commit the operations done in above loop.
        OldTrans.CommitTrans
        
        OldRst.MoveNext
    Wend
    
    If commitPending Then
        OldTrans.CommitTrans
    End If
End If

EndPigmy:

SqlStr = "Delete * FROM PDTrans WHere " & _
    "(TransID is NULL) OR (ACCID is NULL) OR (TransType Is NULL)"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

'Update Modify date
SqlStr = "UPDATE PDMASTER SET ModifiedDate = #1/1/100# WHERE ModifiedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "Correcting the Pigmy Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

'Update Closed date
SqlStr = "UPDATE PDMASTER Set ClosedDate = #1/1/100# WHEre ClosedDate is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If
With frmMain
    .lblProgress = "Correcting the Pigmy Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With



'Update Nominee
SqlStr = "UPDATE PDMASTER set Nominee = '" & SngSpace & "'" & _
        " WHERE Nominee is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "COrrecting the Pigmy Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

SqlStr = "UPDATE PDMASTER set JointHolder= '" & SngSpace & "'" & _
        " WHERE JointHolder is NULL"
OldTrans.SQLStmt = SqlStr
OldTrans.BeginTrans
If Not OldTrans.SQLExecute Then
    OldTrans.RollBack
Else
    OldTrans.CommitTrans
End If

With frmMain
    .lblProgress = "COrrecting the Pigmy Detals"
    PrgVal = PrgVal + 1
    .prg.Value = PrgVal
    .Refresh
End With

MsgBox "Database Checked And Found correct"

Exit_line:
Call OldTrans.CloseDB
Screen.MousePointer = vbDefault
Exit Sub

Err_Line:

If Err.Number = 380 Then
    frmMain.prg.Max = PrgVal * 1.5
    Resume Next
ElseIf Err.Number Then
    MsgBox "ERROr In Chcking data Base"
    GoTo Exit_line
    'Resume
End If

End Sub

