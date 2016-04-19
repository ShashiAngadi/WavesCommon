VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDosPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub PrintText(strText As String)
Dim txtStream As TextStream
Dim FileObj As New FileSystemObject

'Now 'Create a File
Set txtStream = FileObj.OpenTextFile(App.Path & "\InvPrint.txt", ForWriting, True, TristateUseDefault)

'Write the initialisation
txtStream.Write Chr$(27) & "@"
'Now set the Page Height
txtStream.Write Chr$(27) & Chr$(67) & Chr$(0) & Chr$(3)

''Now Set the Left Margin
txtStream.Write Chr$(27) & Chr$(108) & Chr(4)
''Now Set the RIGHT Margin
txtStream.Write Chr$(27) & Chr$(81) & Chr(80)

'now print the left  text
txtStream.Write Chr$(27) & Chr$(97) & "0" & "Left"
txtStream.WriteBlankLines 1

'now print the left text
txtStream.Write Chr$(27) & Chr$(97) & Chr(1) & "Centre"
txtStream.WriteBlankLines 1


'now print the right text
txtStream.Write Chr$(27) & Chr$(97) & "2" & "Right"
txtStream.WriteBlankLines 1


'End of PAGE
txtStream.Write Chr$(12) ' & "@"
txtStream.Write Chr$(27) & Chr$(97) & "0" & Chr$(27) & Chr$(71) & "NextPage"
txtStream.WriteBlankLines 1


'now print the left  text
txtStream.Write Chr$(27) & Chr$(97) & "0" & "Left"
txtStream.WriteBlankLines 1

'now print the left text
txtStream.Write Chr$(27) & Chr$(97) & Chr(1) & "Centre"
txtStream.WriteBlankLines 1


'now print the right text
txtStream.Write Chr$(27) & Chr$(97) & "2" & "Right"
txtStream.WriteBlankLines 1

'End of PAGE
txtStream.Write Chr$(12) ' & "@"
txtStream.Write Chr$(27) & Chr$(97) & "0" & "ThirdPage"
txtStream.WriteBlankLines 1


End Sub

