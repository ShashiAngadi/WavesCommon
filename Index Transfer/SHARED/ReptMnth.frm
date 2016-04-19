VERSION 5.00
Begin VB.Form frmReptMonth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select month"
   ClientHeight    =   1635
   ClientLeft      =   9045
   ClientTop       =   3825
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3045
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1350
      TabIndex        =   7
      Top             =   810
      Width           =   1605
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      Top             =   90
      Width           =   1605
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   1260
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   285
      Left            =   1530
      TabIndex        =   2
      Top             =   1260
      Width           =   675
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Top             =   450
      Width           =   1605
   End
   Begin VB.Label lblDate 
      Caption         =   "Report Date :"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   870
      Width           =   1035
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   195
      Left            =   630
      TabIndex        =   5
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblCurMonth 
      Caption         =   "Current month :"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   510
      Width           =   1125
   End
End
Attribute VB_Name = "frmReptMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event OKClicked(strIndainDate As String)
Public Event CancelClicked()

Private Sub cmbMonth_Click()
Dim TransDate As Date
Dim Mn As Integer
    Mn = cmbMonth.ListIndex + 1
    If cmbYear.Text = "" Then cmbYear.Text = Year(Now)
    If Me.ActiveControl.Name <> cmbMonth.Name Then Exit Sub
    txtDate.Locked = True
    If cmbMonth.ListIndex = 5 Then txtDate.Locked = False
    If Mn = 3 Then
        TransDate = "31/3/" & cmbYear.Text
        txtDate = FormatDate(CStr(TransDate))
        Exit Sub
    End If
    
    'Now get the date on last friday of the selected month
    TransDate = Format(Mn + 1 & "/1/" & cmbYear.Text, "mm/dd/yyyy", vbFriday)
    Do
        TransDate = DateAdd("d", -1, TransDate)
        If Format(TransDate, "dddd") = "Friday" Then Exit Do
    Loop
    txtDate = FormatDate(CStr(TransDate))
    
End Sub


Private Sub cmdCancel_Click()
    RaiseEvent CancelClicked
    Unload Me
End Sub


Private Sub cmdOk_Click()
    If DateValidate(txtDate.Text, "/", True) Then
        RaiseEvent OKClicked(txtDate)
        Unload Me
    End If
    
End Sub


Private Sub Form_Load()

Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
'Now Load All the Months to the combobox
    cmbMonth.Clear
    cmbMonth.AddItem "January"
    cmbMonth.AddItem "February"
    cmbMonth.AddItem "March"
    cmbMonth.AddItem "April"
    cmbMonth.AddItem "May"
    cmbMonth.AddItem "June"
    cmbMonth.AddItem "July"
    cmbMonth.AddItem "August"
    cmbMonth.AddItem "September"
    cmbMonth.AddItem "October"
    cmbMonth.AddItem "November"
    cmbMonth.AddItem "December"

'Now Load The Year
    Dim Count As Integer
    cmbYear.Clear
    For Count = 0 To 5
        cmbYear.AddItem CStr(Year(Now) + Count - 4)
    Next Count

End Sub


