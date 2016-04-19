Attribute VB_Name = "basKannada"
Option Explicit
Declare Function WINTOISCII Lib "win2isc.dll" (ByVal inputStr As String) As String
Declare Function NudiStartKeyboardEngine Lib "Kannada-Nudi.dll" _
    Alias "_NudiStartKeyboardEngineVB@12" (ByVal isGlobal As Boolean, _
    ByVal isMonoLingual As Boolean, _
    ByVal needTrayIcon As Boolean) As Integer
Declare Sub NudiTurnOnScrollLock Lib "Kannada-Nudi.dll" Alias "_NudiTurnOnScrollLockVB@0" ()
Declare Function NudiStopKeyboardEngine Lib "Kannada-Nudi.dll" Alias "_NudiStopKeyboardEngineVB@0" () As Integer
Declare Sub NudiResetAllFlags Lib "Kannada-Nudi.dll" Alias "_NudiResetAllFlagsVB@0" ()
Declare Function NudiGetLastError Lib "Kannada-Nudi.dll" Alias "_NudiGetLastErrorVB@0" () As Integer


Public gFontName As String
Public gFontSize As Single
Public gLangOffSet As Integer
'Public Const wis_KannadaOffset = 5000
Public Const wis_KannadaOffset = 2000
Public Const wis_KannadaSamhitaOffset = 4000
Public Const wis_NoLangOffset = 0
Public gLangShree As Boolean

Private Const NUDI_ERR_ALREADY_RUNNING = -1


Public Function ConvertToIscii(AsciStr As String) As String
Dim StrLen As Integer
Dim IsciWord  As String
Dim I As Integer
Dim SingleChar As String * 1
StrLen = Len(AsciStr)
IsciWord = ""
SingleChar = ""
For I = 1 To StrLen
    SingleChar = Hex(Int(Asc(WINTOISCII(Mid(AsciStr, I, 1)))))
    IsciWord = IsciWord & SingleChar
Next I
ConvertToIscii = IsciWord


End Function

Public Sub KannadaInitialize()
gFontName = "MS Sans Serif"
gFontSize = 8
gLangOffSet = 0
Dim langTool As String
Dim strRet As String
Dim lngRetVal As Long
Dim Rst As ADODB.Recordset
'Include  ..\Shared\wisReg.bas File to the project
'First Get The Lanuage Constant From the Registry

'Get the Language information From Database
If gDbTrans.CommandObject Is Nothing Then
    strRet = ReadFromIniFile("Language", "Language", App.Path & "\" & constFINYEARFILE)
Else
    gDbTrans.SQLStmt = "select * From Install " & _
                    " Where KeyData = 'Language'"
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
                    strRet = FormatField(Rst("ValueData"))
    Set Rst = Nothing
    
    If Len(strRet) = 0 Then _
        strRet = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "Language")
    If Len(strRet) = 0 Then _
        strRet = ReadFromIniFile("Language", "Language", App.Path & "\" & constFINYEARFILE)
End If

If UCase(strRet) = "KANNADA" Then
    
    gLangOffSet = wis_KannadaOffset
    langTool = ReadFromIniFile("Language", "LanguageTool", App.Path & "\" & constFINYEARFILE)
    gFontName = ReadFromIniFile("Language", "FontName", App.Path & "\" & constFINYEARFILE)
End If

If gLangOffSet = wis_KannadaOffset Then
    'First Get the  Windows Path &
    'Akruti Installed Path Form The Win.ini
    Dim WinPath As String
'    If NudiStartKeyboardEngine(False, False, True) = 0 Then
'        MsgBox "Cannot Start Nudi", vbInformation
'        Exit Sub
'    End If
    gFontName = "Nudi B-Akshar"
    gFontSize = 11
    
    If Len(langTool) = 0 Or UCase(langTool) = "NUDI" Then
        ''Default is NUDI
        lngRetVal = NudiStartKeyboardEngine(False, False, False)
        gFontName = "Nudi B-Akshar"
        gLangOffSet = wis_KannadaOffset
    Else
        'Stop the NUdi in case its running
        Call NudiResetAllFlags
        lngRetVal = NudiStopKeyboardEngine()
        gLangShree = True
        gFontName = ReadFromIniFile("Language", "FontName", App.Path & "\" & constFINYEARFILE)
        'Start the Samhita
        Call InitializeSamhita
        gFontSize = 13
        gLangOffSet = wis_KannadaSamhitaOffset
    End If
End If

End Sub
