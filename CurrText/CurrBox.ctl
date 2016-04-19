VERSION 5.00
Begin VB.UserControl CurrText 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "CurrBox.ctx":0000
   ScaleHeight     =   360
   ScaleWidth      =   2055
   ToolboxBitmap   =   "CurrBox.ctx":0B32
   Begin VB.TextBox txtCurr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   2070
   End
End
Attribute VB_Name = "CurrText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
 
 Dim m_TextFigure As String
 Dim m_Symbol  As String
 
 Dim m_NumberString As String
 Dim m_Delimeter As String
 Dim m_CentString As String
 Dim m_DecaString As String
 Dim m_TeenString As String
 Dim m_AndString As String
 
 'Declarations to Get the numbers in figure
 Dim m_NumericExpanded As Boolean
 Dim m_UnitArr() As String
 Dim m_TeenArr() As String
 Dim m_DecaArr() As String
 Dim m_SingleArr() As String

 Dim m_DecimalString As String
 Dim m_CurrencyString As String
 Dim m_CurrencyDecimal As String
 Dim m_Font As StdFont
 
 Private m_SymbolExists  As Boolean
 Private m_NegativeExists  As Boolean
 'Private m_BackSpace As Boolean
 
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
Public Enum wis_Appearence
   Flat = 0
   ThreeD = 1
End Enum

Public Enum wis_BorderStyle
   None = 0
   FixedSingle = 1
End Enum

Public Enum wis_Align
   LeftJustify = 0
   RightJustify = 1
End Enum
Dim m_Alignment As wis_Align

Private Sub ExpandString()
    m_UnitArr = Split(m_CentString, m_Delimeter)
    m_TeenArr = Split(m_TeenString, m_Delimeter)
    m_DecaArr = Split(m_DecaString, m_Delimeter)
    m_SingleArr = Split(m_NumberString, m_Delimeter)
    
m_NumericExpanded = True
End Sub

'Declaration of other public variable

Private Function GetHundredPart(strNumber As String) As String
If Len(strNumber) > 3 Then GoTo ExitLine
Dim tmpStr As String
Dim retString  As String
Dim Retval As Integer
Dim RetArr() As String
tmpStr = strNumber
Retval = GetStringArray(m_CentString, RetArr, m_Delimeter)

If Len(tmpStr) = 3 Then
    retString = GetNumberString(Left(strNumber, 1))
    If retString <> "" Then retString = retString & " " & RetArr(0)
    tmpStr = Right(tmpStr, 2)
End If
If Val(tmpStr) > 0 Then
    'If Len(retString) And PutAndString Then retString = retString & " " & StringAnd
    retString = retString & " " & GetNumberString(Right(strNumber, 2))
    retString = Trim$(retString)
End If
ExitLine:
    GetHundredPart = retString

End Function

Private Function GetNumberString(ByVal strNumber As String) As String
    
    Dim strVal As Integer
    Dim strResult As String
    strVal = CInt(strNumber)
    
If strVal = 0 Then GoTo ExitLine

If Not m_NumericExpanded Then Call ExpandString

If strVal > 0 And strVal < 10 Then
    strResult = m_SingleArr(strVal)
ElseIf strVal > 10 And strVal < 20 Then
    strResult = m_TeenArr((strVal - 10) - 1)
Else
    If Val(Right(strNumber, 1)) = 0 Then
        strResult = m_DecaArr(strVal \ 10 - 1)
    Else
        strResult = m_DecaArr(strVal \ 10 - 1)
        strResult = strResult & " " & GetNumberString(Right(strNumber, 1))
    End If
End If

ExitLine:
GetNumberString = strResult
    
End Function

Public Property Get Locked() As Boolean
    Locked = txtCurr.Locked
End Property

Public Property Let Locked(NewValue As Boolean)
    txtCurr.Locked = NewValue
    PropertyChanged "Locked"
End Property
Public Function NumberInFigure(ByVal strNumber As Double) As String
    
If Not m_NumericExpanded Then Call ExpandString

If UBound(m_UnitArr) < 3 Then GoTo ExitLine
If UBound(m_TeenArr) < 8 Then GoTo ExitLine
If UBound(m_DecaArr) < 8 Then GoTo ExitLine
    
    Dim strText As String
    Dim FigText As String
    Dim retString As String
    Dim DecimalPart As String
    Dim LeftPart As String
    Dim RightPart As String
    
    Dim Pos As Integer
    Dim PrevPos As Integer

'Static LoopCount As Integer
    
    
    FigText = ""
    
    
    strText = CStr(CCur(strNumber))  '12345678
    If Val(strText) = 0 Then GoTo ExitLine
    
    'Now Devide  saperate the decimal part
    Pos = InStr(1, strText, ".")
    If Pos Then
        DecimalPart = Mid(strText, Pos + 1)
        strText = Left(strText, Pos - 1)
        DecimalPart = CStr(Val(DecimalPart))
    End If
    
    'Now put the delimeters to the
    Dim NumbArr() As String
    Dim Ln As Integer
    Dim MaxCount As Integer
    Dim Count As Integer
    
    Count = 0
    Ln = Len(strText)
    ReDim NumbArr(Count)
    MaxCount = 0
    If Ln > 3 Then
        RightPart = Right(strText, 3)
        strText = Left(strText, Len(strText) - 3)
        
        MaxCount = Len(strText) / 2 + 0.5
        ReDim NumbArr(MaxCount)
        NumbArr(MaxCount) = RightPart
        For Count = MaxCount - 1 To 0 Step -1
            NumbArr(Count) = Right(strText, 2)
            
            If Len(strText) <= 2 Then Exit For
            
            strText = Left(strText, Len(strText) - 2)
        Next
    Else
        NumbArr(0) = Val(strText)
    End If
    
    
MaxCount = 0
Count = UBound(NumbArr) '- LBound(NumbArr)

Do
    If Count < LBound(NumbArr) Then Exit Do
    If Count = UBound(NumbArr) Then
        FigText = FigText & GetHundredPart(NumbArr(Count))
    ElseIf MaxCount < 3 Then
        retString = GetNumberString(NumbArr(Count))
        retString = IIf(retString = "", "", retString & " " & m_UnitArr(MaxCount) & " ")
        FigText = retString & FigText
    Else
        Dim tmpStr As String
        tmpStr = NumbArr(Count) & tmpStr
        If MaxCount = UBound(NumbArr) Then
            tmpStr = NumberInFigure(Val(tmpStr)) & " " & m_UnitArr(3)
            FigText = tmpStr & " " & FigText
        End If
    End If
    Count = Count - 1: MaxCount = MaxCount + 1
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


Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtCurr.SelLength
End Property

Public Property Let SelLength(NewValue As Long)
    txtCurr.SelLength = NewValue
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtCurr.SelStart
End Property

Public Property Let SelStart(NewValue As Long)
    txtCurr.SelStart = NewValue
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
SelText = txtCurr.SelText
End Property
Public Property Let SelText(NewValue As String)
txtCurr.SelText = NewValue
End Property

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
   Dim NegativeExists As Boolean
   Dim LeftCommas As Integer
   
   Dim strarr() As String
   ReDim strarr(0)
   
    Dim strText As String
    
    CursorPos = txtCurr.SelStart
    NoOfCommas = Len(txtCurr.Text)
    
    'Check For negative symbol
    strText = txtCurr.Text
    m_NegativeExists = InStr(1, strText, "-")
    
    strText = IIf(m_Alignment = RightJustify, LTrim(txtCurr.Text), txtCurr.Text)
    On Error Resume Next
    txtCurr.Tag = Mid(txtCurr, CursorPos)
    On Error GoTo ErrLIne
    
    CursorPos = CursorPos - (NoOfCommas - Len(strText))
    If m_Symbol <> "" Then
        If InStr(1, txtCurr.Text, m_Symbol, vbTextCompare) = 1 Then
            strText = Mid(txtCurr.Text, Len(m_Symbol) + 1)
        Else
            strText = txtCurr.Text
        End If
        If Left(strText, 1) = "," Then strText = Mid(strText, 2)
    End If
    If m_NegativeExists Then strText = Mid(strText, 2)
    
    'Get the How may commas were there befor the cursor
    LeftCommas = GetStringArray(Left(txtCurr.Text, CursorPos), strarr(), ",") - 1
    If LeftCommas < 0 Then LeftCommas = 0
    
    'Get the No of Commas in the texts
    NoOfCommas = GetStringArray(strText, strarr(), ",") - 1
    If NoOfCommas < 0 Or CCur(strText) < 1000 Then NoOfCommas = 0
    If Trim(strText) = "" Then GoTo LastLine
   
    On Error GoTo ErrLIne
   
   Dim Pos As Integer
   Dim DeciText As String
   Dim RightText As String
   Dim LeftText As String
   
   'Text will Have commas(","), so remove such commas
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
   strText = CStr(Val(strText))
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
If CursorPos < Len(m_Symbol) Then CursorPos = CursorPos + Len(m_Symbol)
If DeciText <> "" Then
    If DeciText = "." Then DeciText = ""
    strText = m_Symbol & IIf(m_NegativeExists, "-", "") & strText & RightText & "." & DeciText
Else
    strText = m_Symbol & IIf(m_NegativeExists, "-", "") & strText & RightText
End If

'Now assign the new text to the textbox
txtCurr.Text = strText

Dim Commas As Integer
Commas = GetStringArray(txtCurr.Text, strarr, ",") - 1
'Put the cursor position correct position

If m_SymbolExists Then
    CursorPos = CursorPos + (Commas - NoOfCommas)
Else
    CursorPos = CursorPos + (Commas - NoOfCommas) + Len(m_Symbol)
End If

strarr = Split(Left(txtCurr.Text, CursorPos), ",")
If UBound(strarr) - 1 > LeftCommas Then
    CursorPos = CursorPos - 1
End If

txtCurr.SelStart = CursorPos
If Len(txtCurr.Tag) > 1 Then
    'txtCurr.SelStart = CursorPos - 1
End If
If txtCurr.SelStart < Len(m_Symbol) + IIf(m_NegativeExists, 1, 0) Then txtCurr.SelStart = Len(m_Symbol) + IIf(NegativeExists, 1, 0)


RaiseEvent Change
Entered = False
   
Exit Sub

ErrLIne:
If Err Then
   MsgBox Err.Description, , "Currency Box error"
   'Resume
End If

End Sub

Private Sub txtCurr_Click()
   RaiseEvent Click
End Sub

Private Sub txtCurr_DblClick()
   RaiseEvent dblClick
End Sub


Private Sub txtCurr_KeyDown(KeyCode As Integer, Shift As Integer)
''''''Shift=1 Shift
''''''Shift=2 control
''''''KeyCode=36 Shift=0  HOME
''''''KeyCode=35 Shift=0  END
''''''KeyCode=37 Shift=0  Left arrow
''''''KeyCode=39 Shift=0  Right arrow
    
    If InStr(1, txtCurr.Text, m_Symbol, vbTextCompare) = 1 Then
        m_SymbolExists = True
    Else
        m_SymbolExists = False
    End If
     'If he presses "-" then for the xisting of the symbol
   
    If KeyCode = 46 Then
        If txtCurr.SelLength = Len(txtCurr.Text) Then GoTo LastLine
        If m_SymbolExists And txtCurr.SelStart < Len(m_Symbol) Then KeyCode = 0
        If Mid(txtCurr.Text, txtCurr.SelStart + 1, 1) = "." Then KeyCode = 0
    End If
'    Debug.Assert txtCurr.SelLength = 0
    If KeyCode = 37 And m_SymbolExists Then 'Left arrow
        If txtCurr.SelStart = Len(m_Symbol) And txtCurr.SelLength Then GoTo LastLine
        If txtCurr.SelStart = Len(m_Symbol) Then KeyCode = 0
    End If

   
LastLine:
   If KeyCode Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub


'
Private Sub txtCurr_KeyPress(KeyAscii As Integer)
   
    'Check Whether Cursor is in the Text area or not
    'If text is right justified
    'then there is as chance of cusror
    'may be on the Currency symbol
    If m_SymbolExists And txtCurr.SelStart < Len(m_Symbol) And txtCurr.SelLength <> Len(txtCurr.Text) Then GoTo LastLine
    If m_SymbolExists And KeyAscii = 8 And txtCurr.SelStart = Len(m_Symbol) Then GoTo LastLine
         
    If KeyAscii = 45 And Not m_NegativeExists Then  '' Minus charactor '-'
        'Then check for cusror pos
        'If cusror is not at the (zeroth)oth postion then exit
        If txtCurr.SelStart > Len(m_Symbol) + IIf(m_NegativeExists, 1, 0) Then GoTo LastLine
        If InStr(Len(m_Symbol) + 1, txtCurr.Text, ".", vbTextCompare) = 0 Then GoTo ExeLine
    End If
    
    If KeyAscii = 45 And Len(txtCurr.SelText) Then _
        If InStr(1, txtCurr.SelText, "-") Then GoTo ExeLine
    
    'Check whether pressed key is numeric or if not numeric
    'then do not consider it
    'if key pressed for operation Paste,Cut,copy then
    If KeyAscii = 3 Or KeyAscii = 24 Or KeyAscii = 8 Then GoTo ExeLine
    If KeyAscii = 22 Then 'Pressed Ctrl+v
    
          Dim strTemp As String
          strTemp = Clipboard.GetText(vbCFText)
          If IsNumeric(strTemp) Then GoTo ExeLine
     End If
     
     If KeyAscii = Asc(".") Then 'Check for the exsting "."
        If InStr(Len(m_Symbol) + 1, txtCurr.Text, ".", vbTextCompare) = 0 Then GoTo ExeLine
     End If
    
     If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then GoTo ExeLine
   
LastLine:
'Else neglect it
   KeyAscii = 0
   Exit Sub

ExeLine:
    'If KeyAscii = 8 Then m_BackSpace = True
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtCurr_KeyUp(KeyCode As Integer, Shift As Integer)
Dim SelLen As Integer
Dim SelStart As Integer

''''''Shift=1 Shift
''''''Shift=2 control
''''''KeyCode=36 Shift=0  HOME
''''''KeyCode=35 Shift=0  END
''''''KeyCode=37 Shift=0  Left arrow
''''''KeyCode=39 Shift=0  Right arrow
    'Check for the existenace negative symobl
    If Left(txtCurr.Text, 1) = "-" And txtCurr.SelStart = 0 Then txtCurr.SelStart = 1
    Dim NegativeExists As Boolean
    NegativeExists = InStr(txtCurr.Text, "-")
    SelStart = txtCurr.SelStart
    SelLen = txtCurr.SelLength
    If KeyCode = 36 Then  'If key code is HOME
        If m_SymbolExists Then txtCurr.SelStart = Len(m_Symbol) + IIf(NegativeExists, 1, 0)
        If SelLen And m_SymbolExists Then txtCurr.SelLength = SelLen - Len(m_Symbol)
    End If
    If KeyCode = 37 Then  'If key code is LEFT Arrow
        'If txtCurr.SelLength = Len(txtCurr.Text) Then GoTo LastLine
        'If m_SymbolExists And KeyCode = 36 Then txtCurr.SelStart = Len(m_Symbol): KeyCode = 0
        'If Mid(txtCurr.Text, txtCurr.SelStart + 1, 1) = "." Then KeyCode = 0
    End If
        
        
    If KeyCode Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtCurr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtCurr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtCurr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'iF HE PASTED ANY DATA THEN CHECK IT'S VALIDITY
Dim strText As String
Dim Pos As Integer
If Button = vbRightButton Then
    On Error Resume Next
    strText = txtCurr.Text
    If m_Symbol <> "" Then
         If InStr(1, txtCurr.Text, m_Symbol, vbTextCompare) = 1 Then
               strText = Mid(txtCurr.Text, Len(m_Symbol) + 1)
         Else
               strText = txtCurr.Text
         End If
    End If
      
    Pos = 0
    Do
        Pos = Pos + 1
        If Pos > Len(strText) Then Exit Do
        If strText = "" Then Exit Do
        If Asc(Mid(strText, Pos, 1)) < Asc("0") Or Asc(Mid(strText, Pos, 1)) > Asc("9") Then
             strText = Left(strText, Pos - 1) & Mid(strText, Pos + 1)
             Pos = Pos - 1
        End If
    Loop
   txtCurr.Text = strText
End If
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_Initialize()
txtCurr.Top = -1
txtCurr.Left = -1

End Sub

Private Sub UserControl_InitProperties()
   txtCurr.Top = 0
   txtCurr.Left = 0
   txtCurr.Text = "Rs.0"
   Set m_Font = txtCurr.Font
   m_Symbol = "Rs."
   m_Delimeter = ","
   m_NumberString = "zero,one,two,three,four,five" & _
                    ",six,seven,eight,nine"
   m_CentString = "hundred,thousand,lakh,crore"
   m_DecaString = "ten,twenty,thirty,forty,fifty,sixty,seventy,eighty,ninty,hundred"
   m_TeenString = "eleven,twelwe,thirteen,fourteen,fifteen," & _
                "sixteen,seventeen,eighteen,ninteen"
   m_AndString = "and"
   
   m_DecimalString = "point"
   m_CurrencyDecimal = "paise"
   m_CurrencyString = "Rupees"
   
   UserControl.BorderStyle = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim tmpStr As String
m_Delimeter = PropBag.ReadProperty("Delimeter", ",")
m_Symbol = PropBag.ReadProperty("CurrencySymbol", "Rs.")

tmpStr = "eleven,twelwe,thirteen,fourteen,fifteen," & _
                "sixteen,seventeen,eighteen,ninteen,twenty"
m_TeenString = PropBag.ReadProperty("TeenString", tmpStr)

tmpStr = "ten,twenty,thirty,forty,fifty,sixty,seventy,eighty,ninty,hundred"
m_DecaString = PropBag.ReadProperty("DecaString", tmpStr)

tmpStr = "zero,one,two,three,four,five,six,seven,eight,nine"
m_NumberString = PropBag.ReadProperty("NumberString", tmpStr)

tmpStr = "hundred,thousand,lakh,crore"
m_CentString = PropBag.ReadProperty("CentString", tmpStr)
m_AndString = PropBag.ReadProperty("AndString", "and")

m_DecimalString = PropBag.ReadProperty("DecimalString", "point")
m_CurrencyString = PropBag.ReadProperty("CurrencyString", "Rupees")
m_CurrencyDecimal = PropBag.ReadProperty("CurrencyDecimal", "paise")

Value = PropBag.ReadProperty("Value", "0")
UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", "1")

Set m_Font = PropBag.ReadProperty("Font", txtCurr.Font)
m_Font.Name = PropBag.ReadProperty("FontName", "MS Sans Serif")
m_Font.Size = PropBag.ReadProperty("FontSize", "8")
m_Font.Bold = CBool(PropBag.ReadProperty("FontBold", "False"))
m_Font.Italic = CBool(PropBag.ReadProperty("FontItalic", "False"))
m_Font.Strikethrough = CBool(PropBag.ReadProperty("FontStrike", "False"))
m_Font.Underline = CBool(PropBag.ReadProperty("FontUnderLine", "False"))


End Sub


Private Sub UserControl_Resize()
   If UserControl.Width < 150 Then UserControl.Width = 150
   If UserControl.Height < 285 Then UserControl.Height = 285
   txtCurr.Left = 0
   txtCurr.Top = 0
   
   txtCurr.Width = UserControl.Width - 25
   txtCurr.Height = UserControl.Height
End Sub

'Public Property Get Alignment() As wis_Align
'   Alignment = m_Alignment
'End Property
'
'Public Property Let Alignment(ByVal NewValue As wis_Align)
''   m_Alignment = NewValue
'End Property


Public Property Get Appearance() As wis_Appearence
Attribute Appearance.VB_Description = "Returns/sets wheteher or not an object is painted at run time with 3-D effects."
   Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal NewValue As wis_Appearence)
    'UserControl.Ambient.Appearance = NewValue
    PropertyChanged "Appearance"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the bacjground color used to display text and graphics in an object."
   BackColor = txtCurr.BackColor
   
   
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
   txtCurr.BackColor = NewValue
   PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As wis_BorderStyle
Attribute BorderStyle.VB_Description = "Returns/set the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As wis_BorderStyle)
   UserControl.BorderStyle = NewValue
   'txtCurr.BorderStyle = NewValue
   PropertyChanged "BorderStyle"
End Property

Public Property Get CurrencySymbol() As String
Attribute CurrencySymbol.VB_Description = "Returns/set the symol of currency used  in the currency text box."
   CurrencySymbol = m_Symbol
End Property
Public Property Let CurrencySymbol(ByVal NewValue As String)
    m_Symbol = NewValue
    PropertyChanged "CurrencySymbol"
End Property
Public Property Get CurrencyString() As String
   CurrencyString = m_CurrencyString
End Property
Public Property Let CurrencyString(str As String)
   m_CurrencyString = str
   PropertyChanged "CurrencyString"
End Property

Public Property Get CurrencyDecimal() As String
   CurrencyDecimal = m_CurrencyDecimal
End Property
Public Property Let CurrencyDecimal(str As String)
   m_CurrencyDecimal = str
   PropertyChanged "CurrencyDecimal"
End Property

Public Property Get Delimeter() As String
Attribute Delimeter.VB_Description = "Returns/set the delimeter used to saperate the words of string."
    Delimeter = m_Delimeter
End Property
Public Property Let Delimeter(NewDelimeter As String)
    
    If NewDelimeter = "" Then
        Err.Raise 50003, "Cureency Box", "Invalid Delimeter specified"
        Exit Property
    End If
    'Check for The Existing of the delimeter string
    If InStr(1, m_NumberString, NewDelimeter, vbTextCompare) Then GoTo LastLine
    If InStr(1, m_TeenString, NewDelimeter, vbTextCompare) Then GoTo LastLine
    If InStr(1, m_DecaString, NewDelimeter, vbTextCompare) Then GoTo LastLine
    If InStr(1, m_CentString, NewDelimeter, vbTextCompare) Then GoTo LastLine
    
    Dim oldDelimeter As String
    oldDelimeter = m_Delimeter
    'Now Replace the existing delimeter with the new delimeter
    Dim Pos As Integer
    Do
        Pos = InStr(Pos + 1, m_NumberString, oldDelimeter)
        If Pos = 0 Then Exit Do
        m_NumberString = Left(m_NumberString, Pos - 1) & NewDelimeter & Mid(m_NumberString, Pos + 1)
    Loop
    Do
        Pos = InStr(Pos + 1, m_DecaString, oldDelimeter)
        If Pos = 0 Then Exit Do
        m_DecaString = Left(m_DecaString, Pos - 1) & NewDelimeter & Mid(m_DecaString, Pos + 1)
    Loop
    Do
        Pos = InStr(Pos + 1, m_TeenString, oldDelimeter)
        If Pos = 0 Then Exit Do
        m_TeenString = Left(m_TeenString, Pos - 1) & NewDelimeter & Mid(m_TeenString, Pos + 1)
    Loop
    Do
        Pos = InStr(Pos + 1, m_CentString, oldDelimeter)
        If Pos = 0 Then Exit Do
        m_CentString = Left(m_CentString, Pos - 1) & NewDelimeter & Mid(m_CentString, Pos + 1)
    Loop
    
    PropertyChanged "NumberString"
    PropertyChanged "TeenString"
    PropertyChanged "DecaString"
    PropertyChanged "TeenString"
    PropertyChanged "Delimeter"
    PropertyChanged "CentString"
    
    
    m_Delimeter = NewDelimeter
    Exit Property
LastLine:
        Err.Raise 50003, "Cureency Box", "This Delimeter is in the string"
        Exit Property
    
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets  a value that determines wheteher an object can respond to user-genereated events."
   Enabled = txtCurr.Enabled
   PropertyChanged "Enabled"
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
   txtCurr.Enabled = NewValue
   PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/set the foreground color used to display text and graphicsin an object."
   ForeColor = txtCurr.ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
   txtCurr.ForeColor = NewValue
   PropertyChanged "ForeColor"
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets font object."
  Set Font = txtCurr.Font
End Property
Public Property Set Font(fnt As StdFont)
    Set txtCurr.Font = fnt
    PropertyChanged "Font"
End Property
Public Property Get StringAnd() As String
Attribute StringAnd.VB_MemberFlags = "400"
   StringAnd = m_AndString
End Property
Public Property Let StringAnd(str As String)
   m_AndString = str
End Property

Public Property Get StringCent() As String
Attribute StringCent.VB_Description = "Returns/set the string of century and above currency units."
Attribute StringCent.VB_MemberFlags = "400"
   StringCent = m_CentString
End Property
Public Property Let StringCent(ByVal NewValue As String)
 'Now validate the the newvalue to be assigned
    Dim noOfStrings  As Integer
    Dim strarr() As String
    
    noOfStrings = GetStringArray(m_CentString, strarr(), m_Delimeter)
    If noOfStrings <> GetStringArray(NewValue, strarr(), m_Delimeter) Then
        Err.Raise 403, "Currency Box", "Invalid value assigned"
        Exit Property
    End If
    m_CentString = NewValue
End Property

Public Property Get StringDeca() As String
Attribute StringDeca.VB_Description = "Returns/set the string of words of ten to hundred."
Attribute StringDeca.VB_MemberFlags = "400"
   StringDeca = m_DecaString
End Property
Public Property Let StringDeca(ByVal NewValue As String)
'Now validate the the newvalue to be assigned
    Dim noOfStrings  As Integer
    Dim strarr() As String
    noOfStrings = GetStringArray(m_DecaString, strarr(), m_Delimeter)
    
    If noOfStrings <> GetStringArray(NewValue, strarr(), m_Delimeter) Then
        Err.Raise 403, "Currency Box", "Invalid value assigned"
        Exit Property
    End If
    m_DecaString = NewValue
End Property
Public Property Get StringDecimal() As String
Attribute StringDecimal.VB_MemberFlags = "400"
   StringDecimal = m_DecimalString
End Property
Public Property Let StringDecimal(str As String)
   m_DecimalString = str
End Property
Public Property Get StringNumber() As String
Attribute StringNumber.VB_Description = "Returns/set the string of words of zero to ten."
Attribute StringNumber.VB_MemberFlags = "400"
    StringNumber = m_NumberString
End Property
Public Property Let StringNumber(NewValue As String)
'Now validate the the newvalue to be assigned
    Dim noOfStrings  As Integer
    Dim strarr() As String
    noOfStrings = GetStringArray(m_NumberString, strarr(), m_Delimeter)
    
    If noOfStrings <> GetStringArray(NewValue, strarr(), m_Delimeter) Then
        Err.Raise 403, "Currency Box", "Invalid value assigned"
        Exit Property
    End If
    m_NumberString = NewValue
End Property
Public Property Get StringTeen() As String
Attribute StringTeen.VB_Description = "Returns/set the string of words of eleven to twenty."
Attribute StringTeen.VB_MemberFlags = "400"
   StringTeen = m_TeenString
End Property

Public Property Let StringTeen(ByVal NewValue As String)
'Now validate the the newvalue to be assigned
    Dim noOfStrings  As Integer
    Dim strarr() As String
    noOfStrings = GetStringArray(m_TeenString, strarr(), m_Delimeter)
    
    If noOfStrings <> GetStringArray(NewValue, strarr(), m_Delimeter) Then
        Err.Raise 403, "Currency Box", "Invalid value assigned"
        Exit Property
    End If
   m_TeenString = NewValue
End Property


Public Property Get Value() As Currency
Attribute Value.VB_Description = "Sets/Returns the numeric value contained in the control"
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "4"
    Value = 0
    Dim strText As String
    strText = txtCurr.Text
    If Left(strText, Len(m_Symbol)) = m_Symbol Then strText = Mid(strText, Len(m_Symbol) + 1)
    If IsNumeric(strText) Then
        Value = strText
    End If
End Property

Public Property Let Value(ByVal NewValue As Currency)
    If Not IsNumeric(NewValue) And Trim(NewValue) <> "" Then
        Err.Raise 5001, "CURRENCY BOX", "Only numeric value will assigned"""
        Exit Property
    End If
    txtCurr.Text = NewValue
End Property
Public Property Get TextInFigure() As String
Dim strText As String
Dim Pos As Integer
Dim DecimalPart As String
    
    strText = txtCurr.Text
    If m_Symbol <> "" Then strText = Right(strText, Len(strText) - Len(m_Symbol))
    
    'Saperate the decimal part
    Pos = InStr(1, strText, ".")
    
    If Pos Then DecimalPart = Mid(strText, Pos + 1)
    If Pos Then strText = Left(strText, Pos - 1)
    DecimalPart = Format(Val(DecimalPart) * 100, "00")
    DecimalPart = Left(DecimalPart, 2)
    DecimalPart = IIf(Val(DecimalPart) > 0, NumberInFigure(Val(DecimalPart)), "")
    If strText = "" Then strText = "0" 'strText
    strText = NumberInFigure(CCur(strText))
    If strText <> "" Then strText = m_CurrencyString & " " & strText
    If strText <> "" And DecimalPart <> "" Then strText = strText & " " & m_AndString
    If DecimalPart <> "" Then strText = strText & " " & DecimalPart & " " & m_CurrencyDecimal
    
    TextInFigure = strText
    
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Sets/Returns the string value of the control"
    Text = txtCurr.Text
End Property
Public Property Let Text(strText As String)
    'Check for the numeric
    On Error GoTo LastLine
    On Error Resume Next
    If InStr(1, strText, m_Symbol) = 1 Then strText = Mid(strText, Len(strText) + 1)
    If Not IsNumeric(strText) And Trim(strText) <> "" Then GoTo LastLine
    Value = Val(strText)
    Exit Property
    
LastLine:
        Err.Raise 5001, "CURRENCY BOX", "Only numeric value will assigned"""
        Exit Property
    
End Property


Private Function GetStringArray(GivenString As String, strArray() As String, Delim As String) As Integer


ReDim strArray(0)
GetStringArray = 0
strArray() = Split(GivenString, Delim)
On Error GoTo ErrLIne
If UBound(strArray) = 0 And strArray(0) = "" Then
    GetStringArray = 0
Else
    GetStringArray = UBound(strArray) + 1
End If

Exit Function

ErrLIne:
    GetStringArray = 0
    Exit Function

'The Below code for Vb 5
Dim Pos As Integer
Dim PrevPos As Integer
Dim tmpStr As String
Dim DelimCount As Integer

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
    If tmpStr <> "" Then
        strArray(UBound(strArray)) = tmpStr
        DelimCount = DelimCount + 1
    Else
        ReDim Preserve strArray(UBound(strArray) - 1)
    End If
    
    GetStringArray = IIf(DelimCount > 0, DelimCount, IIf(tmpStr <> "", 1, 0))
    
    Exit Function

ExitLine:
'GetStringArray = 0
    
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim tmpStr As String

Call PropBag.WriteProperty("CurrencySymbol", m_Symbol, "Rs.")

tmpStr = "eleven,twelwe,thirteen,fourteen,fifteen," & _
                "sixteen,seventeen,eighteen,ninteen,twenty"
Call PropBag.WriteProperty("TeenString", m_TeenString, tmpStr)

tmpStr = "ten,twenty,thirty,forty,fifty,sixty,seventy,eighty,ninty,hundred"
Call PropBag.WriteProperty("DecaString", m_DecaString, tmpStr)

tmpStr = "one,two,three,four,five,Six,Seven,Eight,Nine,zero"
Call PropBag.WriteProperty("NumberString", m_NumberString, tmpStr)

Call PropBag.WriteProperty("Delimeter", m_Delimeter, ",")

tmpStr = "hundred,thousand,lakh,crore"
Call PropBag.WriteProperty("CentString", m_CentString, tmpStr)

Call PropBag.WriteProperty("AndString", m_AndString, "and")
Call PropBag.WriteProperty("DecimalString", m_DecimalString, "point")
Call PropBag.WriteProperty("CurrencyString", m_CurrencyString, "Rupees")
Call PropBag.WriteProperty("CurrencyDecimal", m_CurrencyDecimal, "paise")

Call PropBag.WriteProperty("Value", Value, "0")
Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, "1")
Call PropBag.WriteProperty("FontName", m_Font.Name, "MS Sans Serif")
Call PropBag.WriteProperty("FontSize", m_Font.Size, "8")
Call PropBag.WriteProperty("FontBold", CStr(m_Font.Bold), "False")
Call PropBag.WriteProperty("FontItalic", CStr(m_Font.Italic), "False")
Call PropBag.WriteProperty("FontStrike", CStr(m_Font.Strikethrough), "False")
Call PropBag.WriteProperty("FontUnderLine", CStr(m_Font.Underline), "False")

End Sub


