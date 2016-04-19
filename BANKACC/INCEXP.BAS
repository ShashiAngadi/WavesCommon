Attribute VB_Name = "basSBAcc"
Option Explicit

Public Function ComputeSBInterest(ByVal Product As Currency, Rate As Double) As Currency
    ComputeSBInterest = (Product * 1 * Rate) / (100 * 12)
End Function

Public Function ComputeSBProducts(AccIDArr() As Long, Mon As Integer, Yr As Integer, ByRef Products() As Currency) As Currency
Dim i As Long
Dim rst_Main As Recordset
Dim rst_PM As Recordset
Dim rst_1_10 As Recordset
Dim rst_11_30 As Recordset
Dim TransDate1 As String
Dim TransDate2 As String

'Validate month and year
    If Mon < 1 Or Mon > 12 Then
        Exit Function
    End If

'First get rec set of all the accounts in SB
    gDBTrans.SQLStmt = "Select AccID, ClosedDate from SBMaster order by AccID"
    If gDBTrans.SQLFetch > 0 Then
        Set rst_Main = gDBTrans.Rst.Clone
    Else
        Set rst_Main = Nothing
    End If

'Get Balance upto Last Day of PrevMonth
'This query gets the maximum transaction performed on or before the
'last day of the previous month specified for every account
    TransDate1 = Mon & "/1/" & Yr
    TransDate1 = DateAdd("d", -1, TransDate1)
    gDBTrans.SQLStmt = "Select * from SBTrans as A where TransID = " & _
                       " (Select MAX(TransID) from SBTrans B where " & _
                        " A.AccID = B.AccID and TransDate <= #" & TransDate1 & _
                        "#) order by AccID"
    If gDBTrans.SQLFetch > 0 Then
        Set rst_PM = gDBTrans.Rst.Clone
    Else
        Set rst_PM = Nothing
    End If

'Get max Balance between 1 to 10
    TransDate1 = Mon & "/1/" & Yr
    TransDate2 = Mon & "/10/" & Yr

'Build the Query for max trans between 1 and 10
    gDBTrans.SQLStmt = "Select * from SBTrans A where TransDate = " & _
                        "(Select MAX (TransDate) from SBTrans B " & _
                        " where A.AccID = B.AccID and TransDate between #" & _
                        TransDate1 & "# and #" & TransDate2 & "#) order by AccID"
    
    If gDBTrans.SQLFetch > 0 Then
        Set rst_1_10 = gDBTrans.Rst.Clone
    Else
        Set rst_1_10 = Nothing
    End If

'Get balance upto LAST DAY of current month
'Set the TransCction dates according to month
    TransDate1 = Mon & "/11/" & Yr
    TransDate2 = IIf(Mon < 12, Mon + 1, 1) & "/1/" & IIf(Mon < 12, Yr, Yr + 1)  '1st of Next month
    TransDate2 = DateAdd("d", -1, TransDate2)                                   'Previous day of (1st of Next Month)

    gDBTrans.SQLStmt = "Select * from SBTrans A where Balance = " & _
                        "(Select MIN(Balance) from SBTrans B " & _
                        " where A.AccID = B.AccID and TransDate between #" & _
                        TransDate1 & "# and #" & TransDate2 & "#) order by AccID"
    
    If gDBTrans.SQLFetch > 0 Then
        Set rst_11_30 = gDBTrans.Rst.Clone
    Else
        Set rst_11_30 = Nothing
    End If

'Get 1 and Last Day of current month
    TransDate1 = Mon & "/1/" & Yr
    TransDate2 = IIf(Mon < 12, Mon + 1, 1) & "/1/" & IIf(Mon < 12, Yr, Yr + 1)  '1st of Next month
    TransDate2 = DateAdd("d", -1, TransDate2)                                   'Previous day of (1st of Next Month)

    Dim TransType As wisTransactionTypes
    Dim rst_Trans As Recordset
    Dim rst_Trans_Count As Recordset
    
    TransType = wWithdraw

'Get transactions for all the accounts made during that month
    gDBTrans.SQLStmt = "Select AccID, TransDate from SBTrans where " & _
                        "TransDate >= #" & TransDate1 & "# and " & _
                        "TransDate <= #" & TransDate2 & "# and " & _
                        "TransType = " & TransType & _
                        " order by AccID, TransDate, TransID"
    If gDBTrans.SQLFetch > 0 Then
        Set rst_Trans = gDBTrans.Rst.Clone
    Else
        Set rst_Trans = Nothing
    End If

                        
'Get Count of transactions for all the accounts made during that month
    gDBTrans.SQLStmt = "Select Count(*) as TotalTrans, AccID from SBTrans where " & _
                        "TransDate >= #" & TransDate1 & "# and " & _
                        "TransDate <= #" & TransDate2 & "# and " & _
                        " TransType = " & TransType & " group by AccID"
    If gDBTrans.SQLFetch > 0 Then
        Set rst_Trans_Count = gDBTrans.Rst.Clone
    Else
        Set rst_Trans_Count = Nothing
    End If
    
    
    

Dim Balance As Currency
Dim AccID As Long
Dim ClosedDate As String
Dim Day7 As String, Day14 As String, Day21 As String, Day30 As String
Day7 = Mon & "/7/" & Yr
Day14 = Mon & "/14/" & Yr
Day21 = Mon & "/21/" & Yr
Day30 = IIf(Mon < 12, Mon + 1, 1) & "/1/" & IIf(Mon < 12, Yr, Yr + 1)
Day30 = DateAdd("d", -1, Day30)                                   'Previous day of (1st of Next Month)
Dim Count7 As Integer, Count14 As Integer, Count21 As Integer, Count30 As Integer

'Loop through all the accounts to calculate the products
    For i = 1 To rst_Main.RecordCount
        Balance = 0
        AccID = FormatField(rst_Main("AccID"))
        
        If Not rst_PM Is Nothing Then
            If Not rst_PM.EOF Then
                If AccID = FormatField(rst_PM("AccID")) Then
                    Balance = FormatField(rst_PM("Balance"))
                    rst_PM.MoveNext
                End If
            End If
        End If
        If Not rst_1_10 Is Nothing Then
            If Not rst_1_10.EOF Then
                If AccID = FormatField(rst_1_10("AccID")) Then
                    Balance = FormatField(rst_1_10("Balance"))
                    rst_1_10.MoveNext
                End If
            End If
        End If
        
        If Not rst_11_30 Is Nothing Then
            If Not rst_11_30.EOF Then
                If AccID = FormatField(rst_11_30("AccID")) Then
                    Balance = IIf(Balance < FormatField(rst_11_30("Balance")), Balance, FormatField(rst_11_30("Balance")))
                    rst_11_30.MoveNext
                End If
            End If
        End If
        
        Dim TotalTrans As Integer
        Dim CheckOut As Boolean
        Dim MaxTrans As Integer
        CheckOut = False
        TotalTrans = 0
        MaxTrans = 2
        Count7 = 0: Count14 = 0: Count21 = 0: Count30 = 0
        If Not rst_Trans_Count Is Nothing Then
            If Not rst_Trans_Count.EOF Then
                If AccID = FormatField(rst_Trans_Count("AccID")) Then
                    TotalTrans = FormatField(rst_Trans_Count("TotalTrans"))
                    Debug.Assert TotalTrans <> 0
                    If TotalTrans > MaxTrans * 4 Then
                        'Set balance = 0
                        Balance = 0
                    ElseIf TotalTrans <= MaxTrans Then
                        'Do nothing
                    Else
                        CheckOut = True
                    End If
                    rst_Trans_Count.MoveNext
                End If
            End If
        End If
        
        If CheckOut Then 'Traverse thro the rec set
            If Not rst_Trans Is Nothing Then
                If Not rst_Trans.EOF Then
                    Do
                        If rst_Trans.EOF Then
                            Exit Do
                        End If
                        If FormatField(rst_Trans("AccID")) <> AccID Then
                            Exit Do
                        End If
                        
                        If Not rst_Trans.EOF Then
                            If WisDateDiff(FormatDate(Day7), FormatField(rst_Trans("TransDate"))) <= 0 Then
                                Count7 = Count7 + 1
                            ElseIf WisDateDiff(FormatDate(Day14), FormatField(rst_Trans("TransDate"))) <= 0 Then
                                Count14 = Count14 + 1
                            ElseIf WisDateDiff(FormatDate(Day21), FormatField(rst_Trans("TransDate"))) <= 0 Then
                                Count21 = Count21 + 1
                            Else
                                Count30 = Count30 + 1
                            End If
                            
                            rst_Trans.MoveNext
                        End If
                    Loop
                End If
            End If
        Else
            'Reach the current account number
            If Not rst_Trans Is Nothing Then
                If Not rst_Trans.EOF Then
                    If AccID = FormatField(rst_Trans("AccID")) Then
                        rst_Trans.Move TotalTrans
                    End If
                End If
            End If
        End If
        
        If Count7 > MaxTrans Or Count14 > MaxTrans Or Count21 > MaxTrans Or Count30 > MaxTrans Then
            Balance = 0
        End If
        
        'Check for closure
        If Not IsNull(rst_Main("ClosedDate")) Then  'Account has been closed
            ClosedDate = rst_Main("ClosedDate")
            If (Yr = Year(ClosedDate) And Mon >= Month(ClosedDate)) Or Yr > Year(ClosedDate) Then
                Balance = -1
            End If
        End If
        
        AccIDArr(UBound(AccIDArr)) = AccID
        Products(UBound(Products)) = Balance
        ReDim Preserve AccIDArr(UBound(AccIDArr) + 1)
        ReDim Preserve Products(UBound(Products) + 1)
     
        rst_Main.MoveNext
    
    Next i

End Function
Public Function ComputeTotalSBLiability(AsOnIndianDate As String) As Currency

Dim Ret As Long
Dim Rst As Recordset

'Build SQL
    gDBTrans.SQLStmt = "Select * from SBTrans A where TransDate = " & _
            "(Select MAX(TransDate) from SBTrans B " & _
            " where A.AccID = B.AccID and TransDate <= # " & _
            FormatDate(AsOnIndianDate) & "# ) " & _
            " order by AccID, TransID desc"
    Ret = gDBTrans.SQLFetch
    If Ret <= 0 Then
        Exit Function
    Else
        Set Rst = gDBTrans.Rst.Clone
    End If
        
'COmpute total Liability
    Dim AccID As Long
    Dim Total As Currency
    While Not Rst.EOF
        'See if you have to consider this record
        If AccID = FormatField(Rst("AccID")) Then
            GoTo NextRecord
        End If
        AccID = FormatField(Rst("AccID"))
        Total = Total + FormatField(Rst("Balance"))
NextRecord:
        Rst.MoveNext
    Wend

ComputeTotalSBLiability = Total
    
End Function






