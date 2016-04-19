VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AllTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CrHead"
Attribute VB_Ext_KEY = "Member0" ,"CrHead"
Option Explicit

'local variable to hold collection
Private m_col As Dictionary
Private m_ObjArray() As String

Public Function Update(ByVal HeadID As Long, _
                    ByVal DrAmount As Currency, _
                    ByVal CrAmount As Currency, _
                    ByVal ParentID As Long, _
                    ByVal PrevHeadID As Long, _
                    Optional sKey As String) As Trans
                    
On Error GoTo Hell

    'Create a New Object
    Dim objNewMember As Trans
    Dim CurrentKey As String
    
    CurrentKey = "Tr" & PrevHeadID
    Set objNewMember = New Trans

    'set the properties passed into the method
    objNewMember.HeadID = HeadID
    objNewMember.DrAmount = DrAmount
    objNewMember.CrAmount = CrAmount
    objNewMember.ParentID = ParentID
    
    m_col(CurrentKey).ParentID = ParentID
    m_col(CurrentKey).HeadID = HeadID
    m_col(CurrentKey).DrAmount = DrAmount
    m_col(CurrentKey).CrAmount = CrAmount
    m_col(CurrentKey).Balance = CrAmount
    m_col.Key(CurrentKey) = sKey
    
    'return the object created
    Set Update = objNewMember
    Set objNewMember = Nothing
    
    Exit Function
    
Hell:
    
    Debug.Assert ParentID = 0
'    Resume
    MsgBox "Error In AllTransUpdate"

End Function
Public Function Add(ByVal HeadID As Long, _
                    ByVal DrAmount As Currency, _
                    ByVal CrAmount As Currency, _
                    ByVal ParentID As Long, _
                    Optional sKey As String) As Trans
                    
    'Create a New Object
    Dim objNewMember As Trans
    Set objNewMember = New Trans

    'set the properties passed into the method
    objNewMember.HeadID = HeadID
    objNewMember.DrAmount = DrAmount
    objNewMember.CrAmount = CrAmount
    objNewMember.ParentID = ParentID
    'objNewMember.Balance = Balance
    
    'Add the Trans Class object to the Collection
    Call m_col.Add(sKey, objNewMember)
            
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
  
End Function

'this function checks if the given key exists in the collection
' if exists returns true
Public Function Exists(KeyToSearch As String) As Boolean

On Error GoTo Hell:

Exists = False

Exists = m_col.Exists(KeyToSearch)
  
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

Public Property Get Item(vntIndexKey As Variant) As Trans
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = m_col(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = m_col.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    m_col.Remove vntIndexKey
End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set m_col = New Dictionary
    ReDim m_ObjArray(0)
    
End Sub


Private Sub Class_Terminate()

On Error Resume Next
    'destroys collection when this class is terminated
    Set m_col = Nothing
End Sub

