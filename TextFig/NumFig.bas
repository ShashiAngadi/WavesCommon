Attribute VB_Name = "basNumFig"
 Option Explicit

 Dim m_CentArr() As String
 Dim m_TeenArr() As String
 Dim m_DecaArr() As String
 Dim m_NumArr() As String
 Dim m_DecimalString As String
 Dim m_CurrencyDecimal As String
 Dim m_CurrencyString As String
 Dim m_AndString As String
 Private m_boolInitProp As Boolean
 

Private Sub InitProperties()
    
    m_boolInitProp = True
   m_AndString = "and"
   
   m_DecimalString = "point"
   m_CurrencyDecimal = "paise"
   m_CurrencyString = "Rupees"
    
    ReDim m_CentArr(3)
    m_CentArr(0) = "hundred"
    m_CentArr(1) = "thousand"
    m_CentArr(2) = "lakh"
    m_CentArr(3) = "crore"
    
    ReDim m_TeenArr(8)
    m_TeenArr(0) = "eleven"
    m_TeenArr(1) = "twelwe"
    m_TeenArr(2) = "thirteen"
    m_TeenArr(3) = "fourteen"
    m_TeenArr(4) = "fifteen"
    m_TeenArr(5) = "sixteen"
    m_TeenArr(6) = "seventeen"
    m_TeenArr(7) = "eighteen"
    m_TeenArr(8) = "ninteen"
    
    ReDim m_DecaArr(9)
    m_DecaArr(0) = "ten"
    m_DecaArr(1) = "twenty"
    m_DecaArr(2) = "thirty"
    m_DecaArr(3) = "forty"
    m_DecaArr(4) = "fifty"
    m_DecaArr(5) = "sixty"
    m_DecaArr(6) = "seventy"
    m_DecaArr(7) = "eighty"
    m_DecaArr(8) = "ninty"
    m_DecaArr(9) = "hundred"

    ReDim m_NumArr(9)
    m_NumArr(0) = "zero"
    m_NumArr(1) = "one"
    m_NumArr(2) = "two"
    m_NumArr(3) = "three"
    m_NumArr(4) = "four"
    m_NumArr(5) = "five"
    m_NumArr(6) = "six"
    m_NumArr(7) = "seven"
    m_NumArr(8) = "eight"
    m_NumArr(9) = "nine"

End Sub


Public Function NumberInFigure(strNumber As Double) As String
    
    Dim strText As String
    Dim FigText As String
    Dim retString As String
    Dim DecimalPart As String
    Dim LeftPart As String
    Dim RightPart As String
    
    Dim Pos As Integer
    Dim PrevPos As Integer
    
    FigText = ""
    
'get the values Of the Prop
If Not m_boolInitProp Then InitProperties

    strText = CStr(CCur(strNumber))  '12345678
    If Val(strText) = 0 Then GoTo ExitLine
    
    'Now Devide  saperate the decimal part
    Pos = InStr(1, strText, ".")
    If Pos Then DecimalPart = Mid(strText, Pos + 1)
    If Pos Then strText = Left(strText, Pos - 1)
    DecimalPart = CStr(Val(DecimalPart))
    
    'Now put the delimeters to the
    Dim Ln As Integer
    Dim count As Integer
    Dim NumbArr() As String

    Ln = Len(strText)
    'Now Split the  Given value in parts of
    'Lac,thousand,Hundred ,Tens, Teens and singles
    ReDim NumbArr(0)
    count = 0
    If Ln > 3 Then
        NumbArr(count) = Right(strText, 3)
        strText = Left(strText, Len(strText) - 3)
        count = count + 1
        Do
            If Len(strText) <= 2 Then Exit Do
            ReDim Preserve NumbArr(count)
            NumbArr(count) = Right(strText, 2)
            strText = Left(strText, Len(strText) - 2)
            count = count + 1
        Loop
        ReDim Preserve NumbArr(count)
        NumbArr(count) = Val(strText)
    Else
        NumbArr(0) = Val(strText)
    End If
    
    Dim RevArr() As String
    Dim revCount As Integer
    Dim tmpStr As String
    
    
    ReDim RevArr(0)
    RevArr = NumbArr
    revCount = 0
    Do
        If revCount > UBound(RevArr) Then Exit Do
        NumbArr(revCount) = RevArr(count - revCount)
        revCount = revCount + 1
    Loop
            
    revCount = 0
    Do
        If count < LBound(NumbArr) Then Exit Do
        If count = UBound(NumbArr) Then
            FigText = FigText & GetHundredPart(NumbArr(count))
        ElseIf revCount < 3 Then
            retString = GetNumberString(NumbArr(count))
            retString = IIf(retString = "", "", retString & " " & m_CentArr(revCount) & " ")
            FigText = retString & FigText
        Else
            tmpStr = NumbArr(count) & tmpStr
            If revCount = UBound(NumbArr) Then
                tmpStr = NumberInFigure(Val(tmpStr)) & " " & m_CentArr(3)
                FigText = tmpStr & " " & FigText
            End If
        End If
        count = count - 1: revCount = revCount + 1
    Loop
    
ExitLine:
    If Val(DecimalPart) > 0 Then
        If FigText = "" Then FigText = GetNumberString("0")
        FigText = FigText & " " & m_DecimalString
        Pos = 1
        PrevPos = Len(DecimalPart)
        Do
            strText = Mid(DecimalPart, Pos, 1)
            FigText = FigText & " " & NumberInFigure(Val(strText))
            If Pos = PrevPos Then Exit Do
            Pos = Pos + 1
        Loop
    End If
    
    NumberInFigure = FigText
  
End Function


'Declaration of other public variable

Private Function GetHundredPart(strNumber As String) As String

If Len(strNumber) > 3 Then GoTo ExitLine
Dim tmpStr As String
Dim retString  As String
Dim Retval As Integer
tmpStr = strNumber

If Len(tmpStr) = 3 Then
    retString = GetNumberString(Left(strNumber, 1))
    If retString <> "" Then retString = retString & " " & m_CentArr(0)
    tmpStr = Right(tmpStr, 2)
End If
If Val(tmpStr) > 0 Then
    retString = retString & " " & GetNumberString(Right(strNumber, 2))
    retString = Trim$(retString)
End If

ExitLine:
    GetHundredPart = retString

End Function

Private Function GetNumberString(strNumber As String) As String
    Dim strVal As Integer
    Dim strResult As String
    Dim RetArr() As String
    Dim Retval As String
    strVal = CInt(strNumber)
    
If strVal = 0 Then GoTo ExitLine

If strVal > 0 And strVal < 10 Then
    strResult = m_NumArr(strVal)
ElseIf strVal > 10 And strVal < 20 Then
    strResult = m_TeenArr((strVal - 10) - 1)
Else
    strResult = m_DecaArr(strVal \ 10 - 1)
    If Val(Right(strNumber, 1)) <> 0 Then _
            strResult = strResult & " " & GetNumberString(Right(strNumber, 1))
End If

ExitLine:
GetNumberString = strResult
    
End Function

