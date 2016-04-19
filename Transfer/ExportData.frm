VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportData 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   1860
   ClientTop       =   1815
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   495
      Left            =   1890
      TabIndex        =   2
      Top             =   330
      Width           =   1275
   End
   Begin VB.CommandButton cmdImportData 
      Caption         =   "&ImportData"
      Height          =   555
      Left            =   2580
      TabIndex        =   1
      Top             =   1020
      Width           =   2265
   End
   Begin VB.CommandButton cmdExportMaster 
      Caption         =   "&ExportData"
      Height          =   555
      Left            =   300
      TabIndex        =   0
      Top             =   1020
      Width           =   2205
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   420
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents m_FrmLookUp As frmLookUp
Attribute m_FrmLookUp.VB_VarHelpID = -1
Private Sub cmdExportMaster_Click()
Dim Rst() As Recordset

If Not ExportData() Then
    MsgBox "Failed to transfer the data", vbInformation
End If
End Sub


Private Sub cmdImportData_Click()
If Not ImportData Then
    MsgBox "Unable to import the data"
End If
End Sub


Private Sub cmdView_Click()
Dim Rst As Recordset
gdbtrans.SQLStmt = "SELECT BankID, BankName From BankDet WHERE BankID Mod 100 = 0"
If gdbtrans.SQLFetch < 0 Then Exit Sub

Set m_FrmLookUp = New frmLookUp
Call FillView(m_FrmLookUp.LvwReport, gdbtrans.Rst, True)
m_FrmLookUp.Show vbModal
If gBankId <= 0 Then Exit Sub
gdbtrans.Rst.FindFirst "BankId = " & gBankId
gBankName = FormatField(gdbtrans.Rst("BankName"))
End Sub


