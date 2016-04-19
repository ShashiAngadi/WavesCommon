VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_Name As String
Public value As String
'Public FontName As String
'Public FontSize As Byte
'Public FontBold As Boolean
'Public FontItalic As Boolean
'Public FontUnderline As Boolean
'Public FontStrikethru As Boolean

'Private m_Align As Integer
Private m_Align As AlignmentConstants
Private m_ForeColor As Long
Private m_Font As StdFont

Public Rectangle As Boolean
Public RectMargin As Single
Private Const FIELD_MARGIN = 150

Private m_sum_field As Boolean      ' Indicates whether or not to total this field.
Private m_sum_val As Variant       ' Stores the total of this field, if above is true.

' Applies in case of title and heading sections...
Private m_Height As Single
Private m_Width As Single
Private PrintTop As Single
Public DatePrint As Boolean        ' Sets the flag, whether or not to print date in report title
Public PageNums As Boolean         '   "                   "                               page numbers.
Public Property Let Font(NewValue As StdFont)
Set m_Font = NewValue
End Property

Public Property Get Font() As StdFont
'it returns the font proerty of the field
    Set Font = m_Font
End Property
Public Property Get RectHeight(Obj As Object) As Single
    RectHeight = Obj.TextHeight("A")
End Property

Public Sub SetAttrib(Obj As Object)
' Set the font and color
SetFont Obj
setColor Obj
End Sub
Private Sub setColor(Obj As Object)
On Error Resume Next
Obj.ForeColor = Me.Color
Err.Clear

End Sub

Private Sub SetFont(Obj As Object)
On Error GoTo Err_Line
With Obj
    'If m_Font.Name = "" Then FontName = gFont.Name
    .FontName = m_Font.Name
    .FontSize = m_Font.Size
    .FontBold = m_Font.Bold
    .FontItalic = m_Font.Italic
    .FontUnderline = m_Font.Underline
    .FONTSTRIKETHROUGH = m_Font.Strikethrough
    
    ' if the height of the font exceeds, the height
    ' of the title section, increase the section height.
    If Obj.TextHeight("A") > Height Then
        Height = Obj.TextHeight("A") * 2
    End If

End With

Exit_Line:
    Exit Sub

Err_Line:
    If Err.Number = 438 Then Resume Next
    If Err Then
        MsgBox "SetFont: " & vbCrLf _
            & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Sub

Private Sub Class_Initialize()

' Initilize to various members to default values...
m_Name = ""
value = ""
Set m_Font = New StdFont
m_Font.Name = "MS Sans Serif"
m_Font.Size = 12
m_Font.Bold = True
m_Font.Italic = "False"
m_Font.Underline = False
m_Font.Strikethrough = False

'm_BackColor = vbWhite
Color = vbBlack
RectMargin = 50
' Set the default height for this field.
'Height = 700
End Sub
Public Property Get Name() As Variant
Name = m_Name
End Property

Public Property Let Name(ByVal vNewValue As Variant)
m_Name = vNewValue
End Property

Public Property Get Height() As Variant
' Return height as it is...
Height = m_Height
End Property

Public Property Let Height(ByVal vNewValue As Variant)
m_Height = vNewValue
End Property
Public Property Get Width() As Variant
Width = Printer.TextWidth(Me.Name)
End Property

Public Property Get Align() As AlignmentConstants
'Select Case m_Align
'    Case 0
'        Align = "LEFT"
'    Case 1
'        Align = "CENTER"
'    Case 2
'        Align = "RIGHT"
'End Select
Align = m_Align
End Property
Public Property Let Align(ByVal vNewValue As AlignmentConstants)
'Select Case LCase(vNewValue)
'    Case "left"
'        m_Align = 0
'    Case "center"
'        m_Align = 1
'    Case "right"
'        m_Align = 2
'End Select
m_Align = vNewValue
End Property

Public Property Get Color() As String
Color = m_ForeColor
End Property
Public Property Let Color(ByVal vNewValue As String)
On Error Resume Next
' Breakup the given string into RGB components.
Dim rgbArray() As String
'GetStringArray vNewValue, rgbArray(), ","
rgbArray = Split(vNewValue, ",")

m_ForeColor = RGB(rgbArray(0), rgbArray(1), rgbArray(2))
End Property


Public Property Get RectWidth(Obj As Object) As Single
RectWidth = Obj.ScaleWidth - 2 * FIELD_MARGIN
End Property


Public Property Get SumField() As Boolean
SumField = m_sum_field
End Property

Public Property Let SumField(ByVal vNewValue As Boolean)
m_sum_field = vNewValue
End Property

Public Property Get SumVal() As Variant
SumVal = m_sum_val
End Property

Public Property Let SumVal(ByVal vNewValue As Variant)
m_sum_val = vNewValue
End Property

