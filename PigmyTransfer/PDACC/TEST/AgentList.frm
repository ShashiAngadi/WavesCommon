VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAgentList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AgentCustomerList"
   ClientHeight    =   5685
   ClientLeft      =   1665
   ClientTop       =   1935
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView lstList 
      Height          =   5025
      Left            =   60
      TabIndex        =   0
      Top             =   570
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   8864
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agent Total Customer List "
      Height          =   195
      Left            =   1170
      TabIndex        =   1
      Top             =   270
      Width           =   1860
   End
End
Attribute VB_Name = "frmAgentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

