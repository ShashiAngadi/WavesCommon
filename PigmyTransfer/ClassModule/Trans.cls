VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_HeadID As Long 'local copy
Private m_ParentID As Long 'local copy
Private m_DrAmount As Currency 'local copy
Private m_CrAmount As Currency 'local copy
Private m_Balance As Currency  'Local copy 'sda

Public Property Get ObjectName()
    ObjectName = "TransClass"
End Property


Public Property Let DrAmount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    m_DrAmount = vData
End Property

Public Property Get DrAmount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
    DrAmount = m_DrAmount
End Property
Public Property Let CrAmount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    m_CrAmount = vData
End Property
Public Property Get CrAmount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    CrAmount = m_CrAmount
End Property
Public Property Let HeadID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadID = 5
    m_HeadID = vData
End Property
Public Property Get HeadID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadID
    HeadID = m_HeadID
End Property
Public Property Let ParentID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadID = 5
    m_ParentID = vData
End Property

Public Property Get ParentID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadID
    ParentID = m_ParentID
End Property




