Attribute VB_Name = "basKannadaConv"
Option Explicit
Declare Function LOADTRANSLITERATION Lib "TRANS32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long) As Long
'Procedure to be called before the transliteration process starts.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function CONVERTENGTOLANG Lib "TRANS32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal AString As String, ByVal BString As String, ByVal LSCRPT As Long) As Long
'Procedure to be called to transliterate an English string to language string.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function CONVERTLANGTOENG Lib "TRANS32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal AString As String, ByVal BString As String, ByVal LSCRPT As Long) As Long
'Procedure to be called to transliterate a language string to English.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Sub UNLOADTRANSLITERATION Lib "TRANS32.DLL" ()
'Procedure to be called to close the transliteration process.

Declare Sub SETOPTIONS Lib "TRANS32.DLL" ()

Private inputStr, outputStr As String
Public glScriptCode As Long

'Procedure to be called for performing setup for the transliteration process.
Public Function ConvertToEnglish(ByVal KannadaString As String) As String
  
  On Error GoTo ErrorHandler
  inputStr = adhTrimNull(KannadaString)
  outputStr = Space$(255)
  If inputStr <> "" Then
    CONVERTLANGTOENG Pass1, Pass2, inputStr, outputStr, lScript
  End If
  ConvertToEnglish = adhTrimNull(outputStr)
  Exit Function
ErrorHandler:
  MsgBox " Unable to transliterate from Kannada to English"
End Function

Public Function ConvertToKannada(ByVal EnglishString As String) As String
   On Error GoTo ErrorHandler
 
  inputStr = adhTrimNull(EnglishString)
  outputStr = Space$(255)
  'lScript = 1
  If inputStr <> "" Then
    CONVERTENGTOLANG Pass1, Pass2, inputStr, outputStr, lScript
  End If
  ConvertToKannada = adhTrimNull(outputStr)
  Exit Function
ErrorHandler:
  MsgBox "Unable to transliterate from English to Kannada"

End Function


Public Function GetSortingValue(strKannada As String) As String
    Dim strInput As String
    Dim strOutput As String
    strOutput = Space$(200)
    strInput = Space$(200)
    
    strInput = Trim$(strKannada)
    
    strOutput = SAMHITA.SUCHI2000_SORT32(Pass1, Pass2, strInput, strOutput, glScriptCode)
    
    GetSortingValue = adhTrimNull(strOutput)
    
End Function

Public Function GetNumberToString(value As Double) As String

  Dim str1 As String
  str1 = Space$(200)
  'API call to convert number to words
  Call SHREE2000_NUM_TO_WORDS(Pass1, Pass2, value, str1, glScriptCode, 1, 0)
  GetNumberToString = str1


End Function

Public Function ConvertNudiToSuchita(strNudi As String) As String
              
 Dim strInput As String
 Dim strOutput As String
 strInput = Space$(500)
 strOutput = Space$(500)
 
 strInput = strNudi
 
'Call CONVERTDATA(Pass1, Pass2, str_Renamed1, str2, 7, 18, 61)      ' SUCHI to NUDI Conversion

Call CONVERTDATA(Pass1, Pass2, strInput, strOutput, glScriptCode, 61, 18)      ' NUDI to SUCHI Conversion

ConvertNudiToSuchita = adhTrimNull(strOutput)
End Function

Public Function adhTrimNull(strval As String) As String
    ' adhTrimnull the end of a string, stopping at the first
    ' null character.
    Dim intpos As Integer
    intpos = InStr(strval, vbNullChar)
    If intpos > 0 Then
        strval = Left$(strval, intpos - 1)
    End If
    adhTrimNull = Trim$(strval)
End Function
