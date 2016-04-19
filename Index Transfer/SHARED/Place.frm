VERSION 5.00
Begin VB.Form frmPlaceCaste 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   2520
   ClientTop       =   4200
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   4050
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3150
      TabIndex        =   4
      Top             =   1185
      Width           =   780
   End
   Begin VB.ComboBox cmbPlace 
      Height          =   315
      Left            =   1215
      TabIndex        =   1
      Top             =   435
      Width           =   2715
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   3150
      TabIndex        =   2
      Top             =   780
      Width           =   810
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   2220
      TabIndex        =   3
      Top             =   795
      Width           =   810
   End
   Begin VB.TextBox txtPlace 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   2715
   End
   Begin VB.Label lblPlaceList 
      Caption         =   "Label1"
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   450
      Width           =   1815
   End
   Begin VB.Label lblPlace 
      Caption         =   "Label1"
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   90
      Width           =   1800
   End
End
Attribute VB_Name = "frmPlaceCaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event AddClick(strName As String)
Public Event RemoveClick(strName As String)
Public Event CancelClick(Cancel As Boolean)
Private m_TableName As String
Private PlaceBool As Boolean 'true=place false=caste





Public Sub LoadCombobox(TableName As String)
gDbTrans.SqlStmt = "Select * from " & TableName
cmbPlace.Clear
cmbPlace.AddItem ""
If gDbTrans.SQLFetch > 0 Then
    Dim Rst As Recordset
    Set Rst = gDbTrans.Rst.Clone
    While Not Rst.EOF
        cmbPlace.AddItem FormatField(Rst(0))
        Rst.MoveNext
    Wend
End If
End Sub

Private Sub SetKannadaCaption()
On Error Resume Next
Dim ctrl As Control
    For Each ctrl In Me
        ctrl.Font.Name = gFontName
        If Not TypeOf ctrl Is ComboBox Then
            ctrl.Font.Size = gFontSize
        End If
    Next
    cmdRemove.Caption = LoadResString(gLangOffSet + 12)
    cmdAdd.Caption = LoadResString(gLangOffSet + 10)
    cmdCancel.Caption = LoadResString(gLangOffSet + 2)
    
    ' Labels has to load from the Calling Function
End Sub


Private Sub cmbPlace_Click()
If cmbPlace.Text <> "" Then
cmdRemove.Enabled = True
cmdAdd.Enabled = False
Me.txtPlace.Text = cmbPlace.Text
Else
cmdRemove.Enabled = False
cmdAdd.Enabled = True
End If
End Sub

Private Sub cmdAdd_Click()
Dim Id As Integer
Dim Count As Integer
Dim strText As String
If Trim$(txtPlace.Text) = "" Then Exit Sub
strText = Trim$(txtPlace.Text)

Id = cmbPlace.ListCount - 1
For Count = 0 To Id
    If StrComp(cmbPlace.List(Count), strText, vbTextCompare) = 0 Then
        MsgBox "Already exists", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
Next
Id = 1
'Get the New Id From Table
Dim SqlStr As String
If m_TableName = "PlaceTab" Then
    SqlStr = "SELECT MAX(PlaceID) From PlaceTab"
Else
    SqlStr = "SELECT MAX(CasteID) From CasteTab"
End If
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then Id = FormatField(gDbTrans.Rst(0)) + 1

SqlStr = "INSERT INTO " & m_TableName & " VALUES (" & Id & "," & _
        AddQuotes(strText, True) & ")"

gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If
gDbTrans.CommitTrans
    Unload Me
End Sub


Private Sub cmdCancel_Click()
RaiseEvent CancelClick(True)
Unload Me
End Sub

Private Sub cmdRemove_Click()
Dim Id As Integer

If Trim$(txtPlace.Text) = "" Then Exit Sub
If cmbPlace.ListIndex < 1 Then Exit Sub
Id = cmbPlace.ItemData(cmbPlace.ListIndex)
If Id = 0 Then Exit Sub
If MsgBox("Are you sure you want to remove " & txtPlace, _
    vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Sub

Dim SqlStr As String
If m_TableName = "PlaceTab" Then
    SqlStr = "Delete * FROM PlaceTab WHERE PlaceID = " & Id
Else
    SqlStr = "Delete * FROM CasteTab WHERE CasteID = " & Id
End If
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If
gDbTrans.CommitTrans

   ' RaiseEvent RemoveClick(txtPlace.Text)
    Unload Me

End Sub


Private Sub Form_Activate()
Dim Rst As Recordset
Dim StrPlace As String
Dim StrCaste As String
StrPlace = "Place"
StrCaste = "Caste"

Me.cmbPlace.Clear
cmbPlace.AddItem ""

If InStr(1, lblPlace.Caption, StrPlace) > 0 Then
    PlaceBool = True
    m_TableName = "PlaceTab"
ElseIf InStr(1, lblPlace.Caption, StrCaste) > 0 Then
    PlaceBool = False
    m_TableName = "CasteTab"
Else
    Exit Sub
End If
gDbTrans.SqlStmt = "Select * From " & m_TableName
Me.cmbPlace.Clear
cmbPlace.AddItem ""
'toggle the command Buttons
If gDbTrans.SQLFetch < 1 Then
Me.cmdAdd.Enabled = True
Me.cmdRemove.Enabled = False
GoTo ErrLine
End If
Set Rst = gDbTrans.Rst.Clone
Rst.MoveFirst

cmbPlace.Clear
cmbPlace.AddItem ""
    cmbPlace.ItemData(cmbPlace.NewIndex) = 0
Do While Rst.EOF = False
    cmbPlace.AddItem CStr(FormatField(Rst(1)))
    cmbPlace.ItemData(cmbPlace.NewIndex) = FormatField(Rst(0))
    Rst.MoveNext
Loop
ErrLine:
End Sub

Private Sub Form_Load()
'set icon for the form caption


End Sub


