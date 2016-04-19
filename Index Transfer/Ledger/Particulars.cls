VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticulars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsParticular"
Attribute VB_Ext_KEY = "Member0" ,"clsParticular"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private m_col As Dictionary

Public Function Add(Particulars As String, TransID As Long) As clsParticular

    'create a new object
    Dim objNewMember As clsParticular
    Dim skey As String
    
    Set objNewMember = New clsParticular


    'set the properties passed into the method
    objNewMember.Particulars = Particulars
    objNewMember.TransID = TransID
    skey = "P" & TransID
    Call m_col.Add(skey, objNewMember)
    

    'return the object created
    Set Add = objNewMember
    
    Set objNewMember = Nothing


End Function

' This function checks if the given key exists in the collection
' if exists returns true
Public Function Exists(KeyToSearch As Long) As Boolean

On Error GoTo Hell:

Exists = False

Dim KeyString As String

KeyString = "P" & KeyToSearch

Exists = m_col.Exists(KeyString)
  
Exit Function

Hell:

End Function

Public Function GetKeys() As Variant

    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
  GetKeys = m_col.Keys
  
  
End Function
Public Property Get Item(IndexKey As Variant) As clsParticular
Attribute Item.VB_UserMemId = 0
  Set Item = m_col(IndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = m_col.Count
End Property


Public Sub Remove(vntIndexKey As Variant)

    m_col.Remove vntIndexKey
End Sub




Private Sub Class_Initialize()
    'creates the collection when this class is created
   Set m_col = New Dictionary
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set m_col = Nothing
End Sub

