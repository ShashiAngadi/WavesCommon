VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ColStruct
    ColLength As Integer
    ColText As String
    ColHeader As String
End Type

Private Col() As ColStruct
Private m_Cols As Integer
Private m_Header As String
Private NewPageStarted As Boolean



Public Property Let ColHeader(Index As Integer, Hdr As String)
    Col(Index).ColHeader = Hdr
End Property

Public Property Get ColHeader(Index As Integer) As String
    ColHeader = Col(Index).ColHeader
End Property

Public Property Get Cols() As Integer
    Cols = m_Cols
End Property
Public Property Let Cols(NewValue As Integer)
    If NewValue < 0 Then Err.Raise 10001, "Invalid Cols value specified", "Print Class"
    m_Cols = NewValue
    ReDim Preserve Col(m_Cols)
End Property

Public Property Get ColText(Index As Integer) As String
    ColText = Col(Index).ColText
End Property
Public Property Let ColText(Index As Integer, ColText As String)
    On Error GoTo Errline
    Col(Index).ColText = ColText
    Exit Property
Errline:
    Err.Raise 10001, "Invalid property assignment", "Print Class"
End Property

Public Property Get ColWidth(Index As Integer) As Integer
    ColWidth = Col(Index).ColLength
End Property

Public Property Let ColWidth(Index As Integer, Width As Integer)
    On Error GoTo Errline
    Col(Index).ColLength = Width
    Exit Property
Errline:
    Err.Raise 10001, "Invalid property assignment", "Print Class"
End Property
Public Property Let Header(NewValue As String)
m_Header = NewValue
End Property

Public Property Get Header() As String
Header = m_Header
End Property

Public Sub NewPage()
    Printer.NewPage
    NewPageStarted = True
End Sub

Public Function PrintText()
Dim TmpArr() As String
Dim StrToPrint As String
Dim Flag As Boolean
Dim Count As Integer
Dim TmpStr As String
Printer.Font = "COurier"
If NewPageStarted Then

    If m_Header <> "" Then
        Printer.Print m_Header & vbCrLf & vbCrLf
    End If
        
'Temporary Code
    ReDim TmpArr(m_Cols)
    For Count = 0 To m_Cols
        TmpArr(Count) = Col(Count).ColHeader
    Next Count
    Flag = True
    While Flag
        Flag = False
        StrToPrint = ""
        For Count = 0 To m_Cols
            TmpStr = TmpArr(Count) & Space(Col(Count).ColLength)
            StrToPrint = StrToPrint & Left(TmpStr, Col(Count).ColLength) & " | "
            TmpArr(Count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(Count).ColLength))
            If Len(TmpArr(Count)) > 0 Then
                Flag = True
            End If
        Next Count
        
        Debug.Print StrToPrint
        Printer.Print StrToPrint
        'frmMain.txt.Text = frmMain.txt.Text & StrToPrint
    Wend

'Junk Code
    
    
    NewPageStarted = False
    
End If


'You have a structure called col(count)
    ReDim TmpArr(m_Cols)
    For Count = 0 To m_Cols
        TmpArr(Count) = Col(Count).ColText
    Next Count
    Flag = True
    While Flag
        Flag = False
        StrToPrint = ""
        For Count = 0 To m_Cols
            TmpStr = TmpArr(Count) & Space(Col(Count).ColLength)
            StrToPrint = StrToPrint & Left(TmpStr, Col(Count).ColLength) & " | "
            TmpArr(Count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(Count).ColLength))
            If Len(TmpArr(Count)) > 0 Then
                Flag = True
            End If
        Next Count

        Debug.Print StrToPrint
        Printer.Print StrToPrint
        'frmMain.txt.Text = frmMain.txt.Text & StrToPrint
    Wend
End Function

Private Sub Class_Initialize()
    NewPageStarted = True
End Sub


