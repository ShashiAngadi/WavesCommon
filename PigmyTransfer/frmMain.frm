VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Index 2000 Pigmy Amount Transfer"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbAgent 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdPig2PC 
      Caption         =   "Pigmy Machine to Computer"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdPC2Pig 
      Caption         =   "Computer to Pigmy Machine"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblAgent 
      Caption         =   "Pigmy AgentName"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAgent_Change()

cmdPC2Pig.Enabled = IIf(cmbAgent.ListIndex >= 0, True, False)
cmdPig2PC.Enabled = IIf(cmbAgent.ListIndex >= 0, True, False)
gAgentID = cmbAgent.ItemData(cmbAgent.ListIndex)

End Sub

Private Sub cmbAgent_Click()
Call cmbAgent_Change
End Sub

Private Sub cmdPC2Pig_Click()
If (gAgentID > 0) Then frmPC2Pig.Show vbModal
    
End Sub

Private Sub cmdPig2PC_Click()
    frmPig2PC.Show vbModal
End Sub

Private Sub Form_Load()

SetKannadaCaption
Dim Ret As Long
Dim strRet As String

Call LoadAgentNames(cmbAgent, gAgentID)
cmbAgent.Enabled = Not (gAgentID > 0)

On Error Resume Next
Dim appid
appid = Shell("Prati-Nidhi66.exe 2", vbNormalFocus)
AppActivate appid, True

Dim Filename As String
Dim agentFileName As String

Filename = "\Pig_2_PC.Dat"
agentFileName = agentFileName

If gDEVICE = "BALAJI" Then
    Filename = "\PCRX.TXT"
    If gAgentID > 0 Then agentFileName = Format(gAgentID, "0000") + "-pcrx.dat.txt"
    If Dir(App.Path & agentFileName) <> "" Then
        Filename = "\" & agentFileName
    Else
        If gAgentID > 0 Then agentFileName = Format(gAgentID, "0000") + "-pcrx.dat"
        If Dir(App.Path & agentFileName) <> "" Then Filename = "\" & agentFileName
    End If
End If

If Dir(App.Path & Filename) = "" Then
    cmdPig2PC.Enabled = False
Else
    cmdPig2PC.Enabled = True
End If
'Now Load the Agent Names as per the password logged in

End Sub

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

'lblAgent.Caption = LoadResString(gLangOffSet + 330)

End Sub

