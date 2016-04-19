VERSION 5.00
Begin VB.UserControl CurrText 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   Picture         =   "CurrText.ctx":0000
   ScaleHeight     =   540
   ScaleWidth      =   2070
   ToolboxBitmap   =   "CurrText.ctx":0B32
   Begin VB.TextBox txtCurr 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   60
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "00"
      Top             =   60
      Width           =   1905
   End
End
Attribute VB_Name = "CurrText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 Dim m_TextFigure As String
 Dim M_Symbol  As String
 Dim m_NumberString As String
 Dim m_StringDelimeter As String
 Dim m_CurrencyString As String
 Dim m_DecaString As String
 Dim M_TeenString As String
 
 Private m_SymbolExists  As Boolean
 Private m_BackSpace As Boolean
 
 Event Click()
 Event dblClick()
 Event Change()
 Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Event KeyPress(KeyAscii As Integer)
 Event KeyDown(KeyCode As Integer, Shift As Integer)
 Event KeyUp(KeyCode As Integer, Shift As Integer)

'Declarations of enums
Public Enum TextAppearence
   Flat = 0
   ThreeD = 1
End Enum

Public Enum TextBorderStyle
   None = 0
   FixedSingle = 1
End Enum

Public Enum TextAlign
   LeftJustify = 0
   RightJustify = 1
   Center = 2
End Enum
'Declaration of other public variable

Private Function GetHundredPart(strNumber As String) As String

If Len(strNumber) > 3 Then GoTo ExitLine
Dim retString  As String
Dim Retval As Integer
Dim RetArr() As String

Retval = GetStringArray(m_CurrencyString, RetArr, m_StringDelimeter)
If Val(Left(strNumber, 1)) <> 0 Then
    retString = GetNumberString(Left(strNumber, 1)) & _
        " " & RetArr(0)
End If
    retString = retString & " " & GetNumberString(Right(strNumber, 2))

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
    Retval = GetStringArray(m_NumberString, RetArr(), m_StringDelimeter)
    strResult = RetArr(strVal)
ElseIf strVal < 20 Then
    Retval = GetStringArray(M_TeenString, RetArr(), m_StringDelimeter)
    strResult = RetArr((strVal - 10) - 1)
Else
    Retval = GetStringArray(m_DecaString, RetArr(), m_StringDelimeter)
    If Val(Right(strNumber, 1)) = 0 Then
        strResult = RetArr(strVal / 10 + 1)
    Else
        strResult = RetArr(strVal Mod 10 - 1)
        strResult = strResult & " " & GetNumberString(Right(strNumber, 1))
    End If
End If

ExitLine:
GetNumberString = strResult
    
End Function

Private Function NumberInFigure(strText As String) As String
    
    Dim FigText As String
    FigText = ""
    
    If Trim(strText) = "" Then GoTo ExitLine
    
    'Bifurcate the currency symbol
    If M_Symbol <> "" Then strText = Right(strText, Len(strText) - Len(M_Symbol))
    If Trim(strText) = "" Then GoTo ExitLine
    Dim Pos As Integer
    
    Dim NumbArr() As String
    
    Dim UnitArr() As String
    Dim TeenArr() As String
    Dim DecaArr() As String
    
    If GetStringArray(strText, NumbArr, m_StringDelimeter) < 1 Then GoTo ExitLine
    
    'If GetStringArray(m_NumberString, NumbArr, m_StringDelimeter) < 10 Then GoTo ExitLine
    
    If GetStringArray(m_CurrencyString, UnitArr(), m_StringDelimeter) < 5 Then GoTo ExitLine
    If GetStringArray(M_TeenString, TeenArr(), m_StringDelimeter) < 9 Then GoTo ExitLine
    If GetStringArray(m_DecaString, DecaArr(), m_StringDelimeter) < 9 Then GoTo ExitLine
    
    Dim Count As Integer
    Dim revCount As Integer
    
    Count = UBound(NumbArr) '- LBound(NumbArr)
    Do
        If Count < LBound(NumbArr) Then Exit Do
        If Count = UBound(NumbArr) Then
            FigText = FigText & GetHundredPart(NumbArr(Count))
        ElseIf revCount < 4 Then
            FigText = GetNumberString(NumbArr(Count)) & " " & NumbArr(Count) & " " & FigText
        Else
            Dim tmpStr As String
            tmpStr = NumbArr(Count) & m_StringDelimeter & tmpStr
            If revCount = UBound(NumbArr) Then
                FigText = NumberInFigure(tmpStr) & " " & FigText
                
            End If
        End If
        Count = Count - 1: revCount = revCount + 1
    Loop
    
ExitLine:
   
  NumberInFigure = m_TextFigure
End Function

Private Sub txtCurr_Change()

If Trim(txtCurr.Text) = "" Then Exit Sub
   
   Static Entered As Boolean
'if for the same event it is enetered the loop
'Then before chnaging the text value exit sub
'
   If Entered Then Exit Sub
   Entered = True
   Dim CursorPos As Integer
   Dim NoOfCommas As Integer
   Dim StrArr() As String
   ReDim StrArr(0)
   
   Dim strText As String
   CursorPos = txtCurr.SelStart
    strText = txtCurr.Text
    If M_Symbol <> "" Then
         If InStr(1, txtCurr.Text, M_Symbol, vbTextCompare) = 1 Then
               strText = Mid(txtCurr.Text, Len(M_Symbol) + 1)
         Else
               strText = txtCurr.Text
         End If
    End If
   
   NoOfCommas = GetStringArray(strText, StrArr(), ",") - 1
   If NoOfCommas < 0 Then NoOfCommas = 0
   If Trim(strText) = "" Then GoTo LastLine
   
   On Error GoTo ErrLine
   
   Dim Pos As Integer
   Dim DeciText As String
   Dim RightText As String
   Dim LeftText As String
   
   'Text will Have "," so remove such commas
   Pos = 1
    Do
        Pos = InStr(Pos, strText, ",", vbTextCompare)
        If Pos = 0 Then Exit Do
        strText = Left(strText, Pos - 1) & Mid(strText, Pos + 1)
    Loop
    
   'Find the decimal part of the text
   Pos = InStr(1, strText, ".", vbTextCompare)
   If Pos Then
      DeciText = Mid(strText, Pos + 1)
      If DeciText = "" Then DeciText = "."
      strText = Left(strText, Pos - 1)
   End If
   
   Dim Ln As Integer
   Ln = Len(strText)
   If Ln <= 3 Then GoTo LastLine  'Exit Sub
   
   RightText = "," & Right(strText, 3)
   strText = Left(strText, Len(strText) - 3)
   
   Do
      If Len(strText) <= 2 Then GoTo LastLine
      RightText = "," & Right(strText, 2) & RightText
      strText = Left(strText, Len(strText) - 2)
   Loop
   
LastLine:

If DeciText <> "" Then
  If DeciText = "." Then DeciText = ""
  txtCurr.Text = M_Symbol & strText & RightText & "." & DeciText
Else
   txtCurr.Text = M_Symbol & strText & RightText
End If
Dim Commas As Integer
Commas = GetStringArray(txtCurr.Text, StrArr, ",") - 1
'If Commas < 0 Then Commas = 1
'Put the cursor position correct position

CursorPos = IIf(m_BackSpace And CursorPos, CursorPos - 1, CursorPos)
If m_SymbolExists Then
    txtCurr.SelStart = CursorPos + (Commas - NoOfCommas)
Else
    txtCurr.SelStart = CursorPos + (Commas - NoOfCommas) + Len(M_Symbol)
End If

RaiseEvent Change
Entered = False
   
Exit Sub


ErrLine:
If Err Then
   MsgBox Err.Description, , "Currency Box error"
End If

End Sub

Private Sub txtCurr_Click()
   RaiseEvent Click
End Sub

Private Sub txtCurr_DblClick()
   RaiseEvent dblClick
End Sub


Private Sub txtCurr_KeyDown(KeyCode As Integer, Shift As Integer)
    If InStr(1, txtCurr.Text, M_Symbol, vbTextCompare) = 1 Then
        m_SymbolExists = True
    Else
        m_SymbolExists = False
    End If
    
    If KeyCode = 46 Then
        If m_SymbolExists And txtCurr.SelStart < Len(M_Symbol) Then KeyCode = 0
        If Mid(txtCurr.Text, txtCurr.SelStart + 1, 1) = "." Then KeyCode = 0
    End If
    
    
    'Exit Sub

LastLine:
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCurr_KeyPress(KeyAscii As Integer)
   
   If m_SymbolExists And txtCurr.SelStart < Len(M_Symbol) Then GoTo LastLine
   If m_SymbolExists And KeyAscii = 8 And txtCurr.SelStart = Len(M_Symbol) Then GoTo LastLine
         
   m_BackSpace = False
   'Check whether pressed key is numeric or if not numeric
   'then do not consider it
   
   'if key pressed for operation Paste,Cut,copy then
   If KeyAscii = 3 Or KeyAscii = 24 Or KeyAscii = 8 Then GoTo ExeLine
   If KeyAscii = 22 Then
        Dim strTemp As String
        strTemp = Clipboard.GetText(vbCFText)
        If IsNumeric(strTemp) Then GoTo ExeLine
   End If
   
   If KeyAscii = Asc(".") Then 'Check for the exsting "."
      If InStr(Len(M_Symbol) + 1, txtCurr.Text, ".", vbTextCompare) = 0 Then GoTo ExeLine
   End If
   
   If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then GoTo ExeLine
   
LastLine:
'Else neglect it
   KeyAscii = 0
   Exit Sub

ExeLine:
    If KeyAscii = 8 Then m_BackSpace = True
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtCurr_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtCurr_LostFocus()
'usercontrol_lostfocus
End Sub

Private Sub txtCurr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtCurr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtCurr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
   txtCurr.Top = 0
   txtCurr.Left = 0
   m_StringDelimeter = ","
   m_NumberString = "zero,one,two,three,four,five" & _
                    ",Six,Seven,Eight,Nine"
   m_CurrencyString = "hundred,thousand,lakh,crore"
   m_DecaString = "ten,twenty,thiry,forty,fifty,sixty,seventy,eighty,ninty,"
   M_TeenString = "eleven,twelwe,thirteen,fourteen,fifteen," & _
                "sixteen,seventeen,eighteen,ninteen)"
   
   
End Sub
Private Sub UserControl_Resize()
   txtCurr.Width = UserControl.Width
   txtCurr.Height = UserControl.Height
End Sub

Public Property Get Appearance() As TextAppearence
   Appearance = txtCurr.Appearance
End Property

Public Property Let Appearence(ByVal NewValue As TextAppearence)
   'txtCurr.Appearance = NewValue
End Property
Public Property Get BackColor() As OLE_COLOR
   BackColor = txtCurr.BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
   txtCurr.BackColor = NewValue
End Property

Public Property Get BorderStyle() As TextBorderStyle
    BorderStyle = txtCurr.BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As TextBorderStyle)
   UserControl.BorderStyle = NewValue
End Property

Public Property Get CurrencySymbol() As String
   CurrencySymbol = M_Symbol
End Property
Public Property Let CurrencySymbol(ByVal NewValue As String)
    M_Symbol = NewValue
End Property

Public Property Get DecaString() As String
   DecaString = m_DecaString
End Property
Public Property Let DecaString(ByVal NewValue As String)
    m_DecaString = NewValue
End Property

Public Property Get Enabled() As Boolean
   Enabled = txtCurr.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
   txtCurr.Enabled = NewValue
End Property

Public Property Get Font() As StdFont
   Font = txtCurr.Font
End Property

Public Property Let Font(ByVal NewValue As StdFont)
   Set txtCurr.Font = NewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = txtCurr.ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
   txtCurr.ForeColor = NewValue
End Property

Public Property Get NumberString() As String
    NumberString = m_NumberString
End Property
Public Property Let NumberString(NewValue As String)
    m_NumberString = NewValue
End Property

Public Property Get StringDelimeter() As String
    StringDelimeter = m_StringDelimeter
End Property
Public Property Let StringDelimeter(NewValue As String)
    m_StringDelimeter = NewValue
End Property

Public Property Get TeenString() As String
   TeenString = M_TeenString
End Property

Public Property Let TeenString(ByVal NewValue As String)
   M_TeenString = NewValue
End Property

Public Property Get Text() As String
   Text = txtCurr.Text
End Property

Public Property Let Text(ByVal NewValue As String)

    If Not IsNumeric(NewValue) And Trim(NewValue) <> "" Then
        Err.Raise 5001, "CURRENCY BOX", "Only numeric value will assigned"""
        Exit Property
    End If
    txtCurr.Text = NewValue
End Property

Public Property Let TextInFigure(NewValue As String)
   'TextInFigure = m_TextFigure
End Property
Public Property Get TextInFigure() As String
    Call NumberInFigure(txtCurr.Text)
End Property
'   To get an array from a string seperated by a delimiter
'   Date : 24th Nov 1997
'   Dependencies : <None>
Private Function GetStringArray(GivenString As String, strArray() As String, Delim As String) As Integer

Dim Pos As Integer
Dim PrevPos As Integer
Dim tmpStr As String
Dim DelimCount As Integer

ReDim strArray(0)

tmpStr = GivenString
If Trim(tmpStr) = "" Then GoTo ExitLine

'check whether the delimeter is there at the end
If Right(tmpStr, 1) = Delim Then
 tmpStr = Left(tmpStr, Len(tmpStr) - 1)
End If

Pos = 0
PrevPos = 1

Do
    Pos = InStr(1, tmpStr, Delim)
    If Pos = 0 Then
        Exit Do
    End If
    DelimCount = DelimCount + 1
    
    strArray(UBound(strArray)) = Left(tmpStr, Pos - 1)
    'TmpStr = Right(TmpStr, Len(TmpStr) - Pos)
    tmpStr = Mid(tmpStr, Pos + Len(Delim)) 'changed on 27/2/99
    ReDim Preserve strArray(UBound(strArray) + 1)
Loop
    strArray(UBound(strArray)) = tmpStr
    GetStringArray = IIf(DelimCount > 0, DelimCount + 1, 1)
    Exit Function

ExitLine:
'GetStringArray = 0
    
End Function

