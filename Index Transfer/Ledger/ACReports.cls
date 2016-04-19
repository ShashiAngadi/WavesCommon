VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsACReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum m_WisReportList

ReportNothing = 0
BalancesAsON = 1
TotalTransActionsMade = 2
ProfitandLossTrans = 3
AccountsClosed = 4
GeneralLedger = 5
AccountLedger = 6
DayBook = 7

End Enum

Public erer
