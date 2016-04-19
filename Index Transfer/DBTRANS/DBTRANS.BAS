Attribute VB_Name = "basDBTrans"
Option Explicit


Public Declare Function MakeArchive Lib "Archive" Alias "_MakeArchive@12" (ByVal Fn As Long, ByVal FfileName As String, ByVal Password As String) As Integer


Public Function Archive_CALLBACK(ByRef Path As String, ByRef Size As Long)

Path = String(100, Chr(0))
Path = "Girish"
'Size = 1234
Archive_CALLBACK = 1
End Function


