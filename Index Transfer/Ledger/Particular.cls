VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_TransID As Long
Private m_Particulars As String

Public Property Get TransID() As Long
    TransID = m_TransID
End Property
Public Property Let TransID(ByVal NewValue As Long)
    m_TransID = NewValue
End Property
Public Property Get Particulars() As String
    Particulars = m_Particulars
End Property
Public Property Let Particulars(ByVal NewValue As String)
    m_Particulars = NewValue
End Property
