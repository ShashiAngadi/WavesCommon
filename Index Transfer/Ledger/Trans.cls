VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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
Private m_Balance As Currency  'Loacal copy 'sda
Public Property Get Balance() As Currency
'This returns the BAlance of the Head after this transaction
    Balance = m_Balance
End Property

Public Property Let Balance(ByVal Bal As Currency)
'This Will Stroe the BAlance of the Head after this transaction
    m_Balance = Bal
End Property

Public Property Let DrAmount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    m_DrAmount = vData
End Property

Public Property Get DrAmount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
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




