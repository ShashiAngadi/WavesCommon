Attribute VB_Name = "basCashIdx"
Option Explicit

Function CashValidateKeyAscii(txt As TextBox, Key As Integer) As Boolean
Dim TextPrev As String
Dim Pos As Integer
Dim COunt As Integer
Dim CHar As String * 1
Dim AscVal As Integer
Dim DecimalFound As Boolean
CashValidateKeyAscii = False

#If JUNK Then
'First check for the valid key set
    If Key < Asc("0") Or Key > Asc("9") Then
        If Key <> Asc(".") And Key <> 8 Then GoTo lastline
    End If
#End If

'Preview the text
    TextPrev = PreviewKeyAscii(txt, Key)
    
'Now check all the characters if there is any invalid character
    For COunt = 1 To Len(TextPrev)
        AscVal = Asc(Right(Left(TextPrev, COunt), 1))
        If AscVal < Asc("0") Or AscVal > Asc("9") Then
            If Not DecimalFound Then
                If AscVal <> Asc(".") Then GoTo lastline Else DecimalFound = True
            Else
                GoTo lastline
            End If
        End If
    Next COunt
    
'Now check if there are more than two decimals digits
    Pos = InStr(1, TextPrev, ".", vbBinaryCompare)
    If Pos <> 0 Then 'There is a dot(.)
        If Len(Right(TextPrev, Len(TextPrev) - Pos)) > 2 Then
            GoTo lastline
        End If
    End If
'Check if the left part of the decimal number is within range of currency
    If Len(Left(TextPrev, Len(TextPrev) - Pos)) > 14 Then 'Gosh there is a lot of money here !!!
        GoTo lastline
    End If
        
        

CashValidateKeyAscii = True
Exit Function
lastline:
Key = 0
CashValidateKeyAscii = False
End Function

'   This function allows only the chars present in the ValidSet passed to it.
'   AllowOtherCase allows the other case also.
'   Eg. If your valid set contains A and you want to allow "a" also,
'   then pass AllowOtherCase as TRUE

Function AllowKeyAscii(txt As Object, ValidSet As String, Key As Integer, Optional AllowOtherCase As Boolean) As Integer
Dim COunt As Integer, I As Integer
Dim Flag As Boolean
Dim TempBuf As String

    ReDim InvalidArr(0)
    
    If Not IsMissing(AllowOtherCase) Then
        If AllowOtherCase Then       'We have to consider the case
            ValidSet$ = UCase(ValidSet$) & LCase(ValidSet)
        End If
    End If

    Flag = 0
    For COunt = 1 To Len(ValidSet)
        If Key = Asc(Mid(ValidSet, COunt, 1)) Then
            Flag = True
        End If
    Next COunt
    

    If Key = 22 Then
        TempBuf = Clipboard.GetText
        For COunt = 1 To Len(TempBuf)
            Flag = False
            For I = 1 To Len(ValidSet)
                If Asc(Mid(TempBuf, COunt, 1)) = Asc(Mid(ValidSet, I, 1)) Then
                    Flag = True
                    Exit For
                End If
            Next I
           If Flag = False Then
                Exit For
           End If

        Next COunt
    End If
    
    If Not Flag Then Key = 0
    
End Function

'This function returns the string that was typed before it is displayed on the object
'Thus one can check the string and validate it accordingly before it will be displayed
Function PreviewKeyAscii(txt As Object, Key As Integer) As String

Dim Start As Integer
Dim Length As Integer
Dim Part1 As String, Part2 As String, Part3 As String


    Start = txt.SelStart
    Length = txt.SelLength
    
    If Start < 1 Then
        Part1 = ""
    Else
        Part1 = Left$(txt.Text, Start)
    End If
    
   If (Len(txt.Text) - Start - Length) > 0 Then
        Part3 = Right(txt.Text, Len(txt.Text) - Start - Length)
   Else
        Part3 = ""
   End If
    
   If Key = 22 Then      'Ctrl - V
    Part2 = Clipboard.GetText
   Else
    Part2 = Chr$(Key)
   End If
   
   If Key = 24 Then  ' Ctrl - X
    Part2 = ""
   End If
   
#If OLD Then
'To take care of the Delete Key
    If KeyCod = 46 Then
        Part2 = ""
    End If
#End If

    
   If Key = 3 Then
     PreviewKeyAscii = txt.Text
     Exit Function
   End If
   
   If Key = 8 Then
      Part2 = ""
      If Len(Part1) > 0 Then
      Part1 = Left(Part1, Len(Part1) - 1)
      End If
   End If

PreviewKeyAscii = Part1 & Part2 & Part3

'MsgBox PreviewKeyAscii
End Function


