Attribute VB_Name = "SAMHITA"
'SHREE DLL
Const FONTNAMELEN = 32
Const KEYBOARDNAMELEN = 12

Type TSHREEDATA
    SHREE_ERROR As Byte
    SHREE_ACTIVE As Byte
    CURSCR As Byte
    FontName(FONTNAMELEN) As Byte
    FONTSIZE As Long
    FONTATTR As Byte
    FONTLAYOUT As Byte
    ACTIVATIONKEY As Byte
    HYPHENATION_ON As Byte
    KEYBOARDNAME(KEYBOARDNAMELEN) As Byte
End Type

'The script code values used throughout the Shree-Lipi Soft API calls are as
'follows:
'0:    English
'1:    Devnagri (Marathi, Hindi)
'2:    Gujarati
'3:    Punjabi (Gurumukhi)
'4:    Bengali
'5:    Oriya
'6:    Tamil
'7:    Kannada
'8:    Telugu
'9:    Malayalam
'10:   Sanskrit
'11:   Diacritical
'13:   Arabic
'18:   Assamese

Public Const ENG = 0
Public Const DEV = 1
Public Const GUJ = 2
Public Const PUN = 3
Public Const BAN = 4
Public Const ORI = 5
Public Const TAM = 6
Public Const KAN = 7
Public Const TEL = 8
Public Const MAL = 9
Public Const SAN = 10
Public Const DIA = 11
Public Const ARA = 13
Public Const ASS = 18

'Following fontlayouts are supported
 Public Const MS = 0                        ' Shree-Lipi 2,3
 Public Const SUCHI = 1                     ' Suchika for Shree-Lipi 2,3
 Public Const MS2000 = 15                   ' Shree-Lipi 4,5,6
 Public Const SUCHI2000 = 18                ' Suchika for Shree-Lipi 4,5,6
 Public Const ISCII = 22                    ' ISCII : useful in conversion
 Public Const PCISCII = 23                  ' PCISCII : useful in conversion
 Public Const EAISCII = 24                  ' EAISCII (7 bit) : useful in conversion
 Public Const SORT32 = 25                   ' Sort32 : useful in conversion
 Public Const UNICODELAY = 45               ' Unicode
 Public Const TAMIL99 = 12                  ' Tamil 99 Monolingual
 Public Const BTAMIL99 = 13                 ' Tamil 99 Bilingual

Public glScriptCode As Long
Public SamhitaInitialized As Boolean

Public SHREEREC As TSHREEDATA

Public Pass1 As Long
Public Pass2 As Long

Public bTutorOn As Boolean
Public lScript As Long
Private lFontType As Long
Public gsScript As String * 4
Public gsFontName As String * 32
'Public Const HKEY_LOCAL_MACHINE = &H80000002
Private AStr As String, BStr As String
Public Const MB_OK = &H0&
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&

Declare Function LOADTRANSLITERATION Lib "TRANS32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long) As Long
'Procedure to be called before the transliteration process starts.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function START_SHREE Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal UserProcPtr As Long, ByRef LSHREEDATA As TSHREEDATA) As Long

'  Procedure call to start Shree-Samhita. Unless this procedure is called first,
'  no other procedure should be called.
'  UserProcPtr : Address of the callback procedure (e.g. SHREECHANGEPROC below).
'                This procedure will be called by Shree-Samhita whenever any of
'                its settings change. You can insert your code in this procedure
'                if you want any action to be taken in your application depending on
'                these changes.
'  LShreeData : Pointer to the 'SHREEDATA' data structure used for exchange of data}
'
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.
'  e.g. START_SHREE(PASS1,PASS2,@SHREECHANGEPROC,@SHREEDATA);

Declare Function START_SHREE2 Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long) As Long

'  Procedure call to start Shree-Samhita. Unless this procedure is called first,
'  no other procedure should be called.
'  This start procedure is provided for applications like Visual Foxpro in which
'  it is difficult to pass records as parameters and there is no mechanism to get
'  feedback from SHREE dll. To read status of Shree-Samhita, these applications have
'  to call SHREE_GET_STATUS .
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.
'  e.g. START_SHREE2(PASS1,PASS2);

Declare Function SET_REGISTRY_ROOT_KEY Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal PCREG As String, ByVal HROOT As Long) As Long

'  HKEY_CLASSES_ROOT     = &H80000000;
'  HKEY_CURRENT_USER     = &H80000001;
'  HKEY_LOCAL_MACHINE    = &H80000002;
'  HKEY_USERS            = &H80000003;
'  HKEY_PERFORMANCE_DATA = &H80000004;
'  HKEY_CURRENT_CONFIG   = &H80000005;
'  HKEY_DYN_DATA         = &H80000006;

'Routine to set the name of the registry key under which application specific
'  settings will be stored.
'
'  Parameters
'  PCREG : Pointer to the name of the registry key
'  HROOT : Predefined reserved values which can be one of the following
'        HKEY_CLASSES_ROOT
'        HKEY_CURRENT_USER
'        HKEY_LOCAL_MACHINE
'        HKEY_USERS
'
'  Returns 0 if successful, non zero if any error occurred
'
'  for eg. if you want to store settings under the key
'  'HKEY_LOCAL_MACHINE\SOFTWARE\Modular InfoTech\Shree-Lipi Samhita' then
'  call
'   SET_REGISTRY_ROOT_KEY('SOFTWARE\Modular InfoTech\Shree-Lipi Samhita',HKEY_LOCAL_MACHINE);


Declare Function SHREE_RESTORE_REGISTRY Lib "Shree.dll" () As Long
'Routine to set the RESTORE the default application specific registry settings
'  Returns 0 if successful, non zero if any error occurred

Declare Sub CLOSE_SHREE Lib "Shree.dll" ()
'Procedure call to close Shree-Samhita

Declare Sub SHREE_SETUP Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long)
'Procedure call to invoke the 'Full Setup' dialog of Shree-Samhita.
'  Call this procedure when you want your users to change the Shree-Samhita setup.
'  This procedure need never be called if the setup commands are given by program only
'  and no control is extended to the user.

Declare Sub SHREE_KBD_SETUP Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long)
'Procedure call to invoke the 'Concise Setup' dialog of Shree-Samhita.
'  This dialog contains only the keyboard related setup.
'  Call this procedure when you want your users to change the Shree-Samhita keyboard related setup.
'  This procedure need never be called if the setup commands are given by program only
'  and no control is extended to the user.

Declare Function SHREE_SETSCRIPT Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long) As Long
'Procedure call to change the current script.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Sub SHREE_TUTOR_ON Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal HCallerWnd As Long)
'Procedure to invoke the Shree-Samhita keyboard tutor on screen.
'  HCALLERWINDOW : Handle of the caller window. This handle is passed to the tutor
'                  so that if any characters are entered via mouse clicks on the tutor
'                  keyboard, the same appear in the caller window.


Declare Sub SHREE_TUTOR_ON1 Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal CALLERTITLE As String)
'Procedure to invoke the Shree-Samhita keyboard tutor on screen.
'  CALLERTITLE : Title of the caller window. This title is passed to the tutor
'                  so that if any characters are entered via mouse clicks on the tutor
'                  keyboard, the same appear in the window having the CALLERTITLE.

Declare Sub SHREE_TUTOR_OFF Lib "Shree.dll" ()
'Procedure to remove the Shree-Samhita keyboard tutor from the screen.

Declare Function SHREE_SET_KEYBOARD Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long, ByVal KbdName As String) As Long
'Procedure to set the desired keyboard layout of the specified script
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  KbdName : should point to the string specifying the keyboard filename e.g. ENG.DEV
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.
            
Declare Function SHREE_FONTNAME_TO_SCRIPT Lib "Shree.dll" (ByVal FontName As String) As Long
'  Function to be called to find out the script to which a particular font belongs
'  FontName : Pointer to the Fontname String
'  Returns The Script constant. The Values are interpreted as documented above.

Declare Function SHREE_SET_APPLICATION_TYPE Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal APPTYPE As Long) As Long
'Function call to set the application type of the caller. This call is required
'  because different applications / windows components behave differently as far as
'  handling of Indian scripts are concerned.
'  APPTYPE : This value indicates the application type. The permissible values are
'            0 or 1. Check for your application which value is suitable. The default
'            setting is 0. Most of the edit controls require this value to be 0, but the
'            richedit components require the value to be 1.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function SHREE_SET_APPLICATION_NMAE Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal APPNAME As String) As Long
'Function call to set the application name of the caller. This call is required
'  because sometimes if Shree.Dll has to take some special care for that particular
'  application.
'  APPNAME : should point to the string specifying the application name.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Sub SHREE_FIRSTFONT_FOR_SCRIPT Lib "Shree.dll" (ByVal LSCR As Long, ByVal FNAME As String)
' Procedure to be called to find out the first installed Shree-Lipi font for a given script.
'  This function is useful to assign default font to components for a given script
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  FNAME : Pointer to the output string giving the fontname. Sufficient memory must be
'          allocated by the caller before calling this function.

Declare Sub SUCHI_FIRSTFONT_FOR_SCRIPT Lib "Shree.dll" (ByVal LSCR As Long, ByVal FNAME As String)
'  Procedure to be called to find out the first installed Suchika font for a given script.
'  This function is useful to assign default font to components for a given script
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  FNAME : Pointer to the output string giving the fontname. Sufficient memory must be
'          allocated by the caller before calling this function.

Declare Sub SHREE_SCRIPT_TO_STR Lib "Shree.dll" (ByVal LSCR As Long, ByVal OpStr As String)
'Procedure to be called to get full name of a given script. For example, full name
'  for script value of 1 is "Devnagri"
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  OPSTR : Pointer to the output string giving the script name. Sufficient memory must be
'          allocated by the caller before calling this function.

Declare Sub SHREE_SCRIPT_TO_SHORTSTR Lib "Shree.dll" (ByVal LSCR As Long, ByVal OpStr As String)
'Procedure to be called to get short name of a given script. For example, short name
'  for script value of 1 is "DEV"
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  OPSTR : Pointer to the output string giving the script name. Sufficient memory must be
'          allocated by the caller before calling this function. The ouput string will
'          be 3 character long

Declare Sub SHREE_SETFONTTYPE Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long, ByVal LF As Long)
'Function call to change the font layout of the current font.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  LF = 0 means the Shree-Lipi fonts (old layout)
'  LF = 1 means the Suchika bilingual fonts(old layout)
'  LF = 12 means the TAMIL99 monolingual fonts
'  LF = 13 means the TAMIL99 bilingual fonts
'  LF = 15 means the Shree-lipi 2000 Fonts
'  LF = 18 means the Suchi 2000 Fonts
'  LF = 35 means the Shree Deccan Herald Fonts
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function GET_SHREE_INSTALLED_SCRIPTS Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal SCRLIST As String, ByRef SCRLISTSIZE As Long) As Long
'This procedure can be called to get a list of installed scripts. The script
'  codes are returned in SCRLIST string and are separated by comma. If this
'  procedure is called with a null SCRLIST string, the size of the
'  required string is returned in SCRLISTSIZE. The caller should acquire enough
'  memory to hold the returned string. If the size of SCRLIST is less than the
'  required size, the returned string will be truncated to that size. The size should
'  be indicated by SCRLISTSIZE.
'  SCRLIST : Pointer to the string for the returned list of installed scripts.
'  SCRLISTSIZE : Size of the SCRLIST string.

Declare Function SHREE_GET_DEFAULTFONT_FOR_SCRIPT Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long, ByVal FNAME As String) As Long
'  Procedure to be called to find out the default font for a given script.
'  This function is useful to assign default font to components for a given script
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  FNAME : Pointer to the output string giving the fontname. Sufficient memory must be
'          allocated by the caller before calling this function.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function GET_SHREE_STATUS Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, LSHREEDATA As TSHREEDATA) As Long
'Function to be called to get the Shree Samhita status data structure values.
'  LSHREEDATA : Pointer to the SHREEDATA data structure.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function SHREE_GET_FONTNAMES Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long, ByVal FONTLIST As String, ByRef FONTLISTSIZE As Long) As Long
'This procedure can be called to get a list of available Shree-Lipi fonts of a given script.
'  The font names are returned in FONTLIST string and are separated by
'  comma. If this procedure is called with a null FONTLIST string, the size of the
'  required string is returned in FONTLISTSIZE. The caller should acquire enough
'  memory to hold the returned string. If the size of FONTLIST is less than the
'  required size, the returned string will be truncated to that size. The size should
'  be indicated by FONTLISTSIZE.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  FONTLIST : Pointer to the string for the returned list of Shree-Lipi fonts.
'  FONTLISTSIZE : Size of the FONTNAMES string.

Declare Function SUCHI_GET_FONTNAMES Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long, ByVal FONTLIST As String, ByRef FONTLISTSIZE As Long) As Long
'This procedure can be called to get a list of available Suchika fonts of a given script.
'  The font names are returned in FONTLIST string and are separated by
'  comma. If this procedure is called with a null FONTLIST string, the size of the
'  required string is returned in FONTLISTSIZE. The caller should acquire enough
'  memory to hold the returned string. If the size of FONTNAMES is less than the
'  required size, the returned string will be truncated to that size. The size should
'  be indicated by FONTLISTSIZE.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  FONTLIST : Pointer to the string for the returned list of Suchika fonts.
'  FONTLISTSIZE : Size of the FONTNAMES string.

Declare Function SHREE_GET_ILSCRIPT Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long) As Long
'Returns the current Indian Language script. If the user toggles the activation
'  key, Shree Samhita toggles between English and this script.
'  The Values are interpreted as documented above.

Declare Function SHREE_GET_KEYBOARDS Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal KBDNAMES As String, ByVal KBDNAMESIZE As Long) As Long
'  This procedure can be called to get a list of available keyboards of a given script.
'  The keyboard names are returned in KBDNAMES string and are separated by
'  comma. If this procedure is called with a null KBDNAMES string, the size of the
'  required string is returned in KBDNAMESIZE. The caller should acquire enough
'  memory to hold the returned string. If the size of KBDNAMES is less than the
'  required size, the returned string will be truncated to that size. The size should
'  be indicated by KBDNAMESIZE.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  KBDNAMES : Pointer to the string for the returned list of keyboard layouts.
'  KBDNAMESIZE : Size of the KBDNAMES string.

Declare Function SHREE_GET_FONTTYPE Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long) As Long
'Function call to get the font layout of the current font.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns 0 means Shree-Lipi fonts (old layout)
'  Returns 1 means Suchika bilingual fonts (old layout)
'  Returns 12 means TAMIL99 monolingual fonts
'  Returns 13 means TAMIL99 bilingual fonts
'  Returns 15 means the Shree-lipi 2000 Fonts
'  Returns 18 means the Shree-lipi 2000 Fonts
'  Returns : -1 if unsuccessful.
  
Declare Function SHREE_STR_TO_SCRIPT Lib "Shree.dll" (ByVal OpStr As String) As Long
'Function to be called to get script constant for the full name of a given script
'  IPSTR : Pointer to the full name string for a script. The Full name string should be as
'          given by the SHREE_SCRIPT_TO_STR function
'  Returns The Script constant. The Values are interpreted as documented above.


Declare Function SHREE_SHORTSTR_TO_SCRIPT Lib "Shree.dll" (ByVal IpStr As String) As Long
'  Function to be called to get script constant for the short name of a given script
'  IPSTR : Pointer to the short name string for a script. The short name string should be as
'          given by the SHREE_SCRIPT_TO_SHORTSTR function
'  Returns The Script constant. The Values are interpreted as documented above.
Declare Sub SHREE_SET_ACTIVATION_KEY Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal K1 As Long)
'  Function call to set the activation key.
'  K1 : The windows standard virtual key code for the activation keys. The activation key
'       can be one of the following
'       VK_Scroll : the scroll lock key
'       VK_Num : the num lock key
'       VK_Capital : the caps lock key

Declare Function SHREE_SET_ESCAPEMENT Lib "Shree.dll" (ByVal BYCTRL As Byte, ByVal BYSHIFT As Byte, BYKEYCODE As Byte) As Long
'PROCEDURE TO SET THE ESCAPEMENT KEY
'PARAMETERS:
'    BYCTRL : WHETHER CONTROL KEY FORMS THE ESCAPEMENT COMBINATION
'             POSSIBLE Values
'                0 : CONTROL KEY IS NOT
'                1 : CONTROL KEY IS ONE OF THE KEYS
'    BYSHIFT : WHETHER SHIFT KEY FORMS THE ESCAPEMENT COMBINATION
'             POSSIBLE Values
'                0 : SHIFT KEY IS NOT
'                1 : SHIFT KEY IS ONE OF THE KEYS
'    BYKEYCODE : VIRTUAL KEY CODE OF THE KEY. THIS CANNOT BE 0
'
'    FOR eg.
'
'    1. IF YOU WANT TO SET ESCAPEMENT KEY AS CTRL + SHIFT + 1
'       THEN THE FUNCTION SHOULD BE CALLED IN FOLLOWING WAY
'           SHREE_SET_ESCAPEMENT(1,1,48);
'    2. IF YOU WANT TO SET ESCAPEMENT KEY AS CTRL + 1
'       THEN THE FUNCTION SHOULD BE CALLED IN FOLLOWING WAY
'           SHREE_SET_ESCAPEMENT(1,0,48);
'    3. IF YOU WANT TO SET ESCAPEMENT KEY AS F1
'       THEN THE FUNCTION SHOULD BE CALLED IN FOLLOWING WAY
'           SHREE_SET_ESCAPEMENT(0,0,48);
'
'    RETURNS 0 IF SUCCESSFUL ELSE NEGATIVE VALUE WHICH SIGNIFIES THAT
'    IMPROPER KEY COMBINATION

Declare Sub TAMIL99_FIRSTFONT Lib "Shree.dll" (ByVal BILINGUAL As Long, ByVal FNAME As String)
'  Procedure to be called to find out the first installed TAMIL99 font.
'  This function is useful to assign default font to components when using TAMIL99 layout.
'  BILINGUAL : Parameter to indicate whether monolingual or bilingual font is required.
'              values to be passed : 0 if monolingual, 1 if bilingual.
'  FNAME : Pointer to the output string giving the fontname. Sufficient memory must be
'          allocated by the caller before calling this function.

Declare Function TAMIL99_GET_FONTNAMES Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal BILINGUAL As Long, ByVal FONTLIST As String, ByVal FONTLISTSIZE As Long) As Long
'This procedure can be called to get a list of available fonts with Tamil99 layout.
'  The font names are returned in FONTLIST string and are separated by
'  comma. If this procedure is called with a null FONTLIST string, the size of the
'  required string is returned in FONTLISTSIZE. The caller should acquire enough
'  memory to hold the returned string. If the size of FONTLIST is less than the
'  required size, the returned string will be truncated to that size. The size should
'  be indicated by FONTLISTSIZE.
'  BILINGUAL : Parameter to indicate whether monolingual or bilingual fonts are required.
'              values to be passed : 0 if monolingual, 1 if bilingual.
'  FONTLIST : Pointer to the string for the returned list of Suchika fonts.
'  FONTLISTSIZE : Size of the FONTLIST string.

Declare Function SHREE_FIRSTFONT_FORSCRIPT_EX Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Long, ByVal lFontType As Long, ByVal pcFontName As String) As Long

'  Purpose           : retrieves the first font for Script and font layout
'  Input             : PASS1 : Password 1
'                      PASS2:  Password 2
'                      lScr  : The Script constant. The Values are interpreted
'                              as documented above.
'                      lFontType :  the values are interpreted as shown below
'                      0  : Shree-Lipi fonts
'                      1  : Suchika bilingual fonts
'                      2  : ITR Swadesh 2.0
'                      12 : TAMIL99 monolingual fonts
'                      13 : TAMIL99 bilingual fonts
'                      15 : Shree-lipi 2000 Fonts
'                      18:                        Suchika 2000
'                      20:                        Prakashak Bengali
'                      pcFontName : pointer to the string for the returned
'                                   First Font
'  Result            : returns 0 if successfull and -1 if fails

Declare Function SHREE_SET_TUTOR_POSITION Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal iPos As Integer) As Integer
'  Purpose           : sets the position of keyboard tutor in the Screen
'  Input             : PASS1 : Password 1
'                      PASS1 : Password 2
'                      iPos  : the values are interpreted as shown below
'0:                         Top Left
'1:                         Top Right
'2:                         Bottom Left
'3:                         Bottom Right
'4:                         Screen Center
'  Result            : returns 0 if successfull and -1 if fails

Declare Function SHREE_SET_DEFAULT_FONT Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal LSCR As Integer, ByVal lFontType As Long, ByVal pcFontName As String, ByVal iFontSize As Integer, ByVal iFontStyle As Integer) As Integer
'  Purpose           : set the default font for Script and font layout
'  Input             : PASS1 : Password 1
'                      PASS2:  Password 2
'                      lScr  : The Script constant. The Values are interpreted
'                              as documented above.
'                      lFontType :  the values are interpreted as shown below
'                      0  : Shree-Lipi fonts
'                      1  : Suchika bilingual fonts
'                      2  : ITR Swadesh 2.0
'                      12 : TAMIL99 monolingual fonts
'                      13 : TAMIL99 bilingual fonts
'                      15 : Shree-lipi 2000 Fonts
'                      18 :  Suchika 2000
'                      20 :  Prakashak Bengali
'                      pcFontName : pointer to the name of the default font
'                      iFontSize  : size of the default font
'                      iFontStyle : style of the default font. The values are
'                                   interpreted As below
'                      0 : Normal
'                      1 : Bold
'                      2 : Italic
'                      3 : Bold , Italic
'  Result            : returns 0 if successfull and -1 if fails

Declare Function SHREE_ENABLE_TOOLTIP Lib "Shree.dll" (ByVal iShow As Integer) As Integer
'  Purpose           : Show Key contents on tool tip is to be enabled/disabled
'  Input             : iShow  : the values are interpreted as shown below
'                      0  : not to be shown
'                      1  : to be shown
'  Result            : returns 0 if successfull and -1 if fails

Declare Function SHREE_ENABLE_IN_DIALOGS Lib "Shree.dll" (ByVal byEnable As Byte) As Integer
'  Purpose           : To disable typing in Indian language on Edit Control
'  Input             : byEnable  : the values are interpreted as shown below
'                      0  : not allow typing on Edit controls
'                      1  : allow typing on Edit controls
'  Result            : returns 0 if successfull and -1 if fails

Declare Function SHREE_GET_MAINSCRIPT Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long) As Long
'  Returns the main Indian Language script for the package.
'  The Values are interpreted as documented above.

Declare Function SHREE_SETUPEX Lib "Shree.dll" (ByVal Pass1 As Long, ByVal Pass2 As Long) As Long
'  Procedure call to invoke the 'Full Setup' dialog of Shree-Samhita.
'  Call this procedure when you want your users to change the Shree-Samhita setup.
'  This procedure need never be called if the setup commands are given by program only
'  and no control is extended to the user.
'  Returns: 0 if cancel button on the setup form was pressed.
'           1 if OK button on the setup form was pressed.

'************************************************

'CNVAPI32.DLL
'=================================================================================

'=================================================================================

' Shree-lipi Samhita version 1.1 supports two font layouts viz Shree-lipi and
' Shree-lipi 2000 Font format which is Windows 2000 Compatible. If you have older
' version then you should the Font Layout Code of 0 else you should use the
' Font layout code of 15. All functions in this module have been updates
' according to the new transformations


Declare Function INIT_CONVERT Lib "CNVAPI32.DLL" () As Long
'Procedure to be called once before any conversion operation starts. Initializes
'  data structures used by conversion procedures.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function CONVERTDATA Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, _
    ByVal LangName As Long, ByVal AFontType As Long, ByVal BFontType As Long) As String

'  Function to convert string in one font codes to a string in another font code
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  LangName : The Script constant. The Values are interpreted as documented above.
'  AFontType : The font layout of the Input string. The Values are interpreted as documented above.
'  BFontType : The required font layout of the Output string. The Values are interpreted as documented above.


Declare Function SHREE_NUM_TO_WORDS Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal NUM As Double, ByVal OpStr As String, ByVal LSCR As Long, ByVal SCRIPT_LANG As Long, ByVal SplitNum As Long) As Long
'Procedure to convert a number in figures to number in words in a given language
'  NUM : Number to be converted
'  OPSTR : Pointer to Output String in Shree-Lipi font codes.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : A value indicating the language within the given script. This value
'                has to be specified when a script supports multiple languages.
'                when script = Devnagri, SCRIPT_LANG = 0 means MArathi,
'                SCRIPT_LANG = 1 means Hindi.
'  SPLIT_NUM : This variable is passed to express number in thousands, lakhs, or crores
'              0 : Return the number without any split
'              1 : Return the number in thousands
'              2 : Return the number in lakhs
'              3 : Return the number in crores
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function SUCHI_NUM_TO_WORDS Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal NUM As Double, ByVal OpStr As String, ByVal LSCR As Long, ByVal SCRIPT_LANG As Long, ByVal SplitNum As Long) As Long
'Procedure to convert a number in figures to number in words in a given language
'  NUM : Number to be converted
'  OPSTR : Pointer to Output String in Suchika font codes.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : A value indicating the language within the given script. This value
'                has to be specified when a script supports multiple languages.
'                when script = Devnagri, SCRIPT_LANG = 0 means MArathi,
'                SCRIPT_LANG = 1 means Hindi.
'  SPLIT_NUM : This variable is passed to express number in thousands, lakhs, or crores
'              0 : Return the number without any split
'              1 : Return the number in thousands
'              2 : Return the number in lakhs
'              3 : Return the number in crores
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function SHREE2000_NUM_TO_WORDS Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal NUM As Double, ByVal OpStr As String, ByVal LSCR As Long, ByVal SCRIPT_LANG As Long, ByVal SplitNum As Long) As Long
'Procedure to convert a number in figures to number in words in a given language
'  NUM : Number to be converted
'  OPSTR : Pointer to Output String in Shree-Lipi font codes.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : A value indicating the language within the given script. This value
'                has to be specified when a script supports multiple languages.
'                when script = Devnagri, SCRIPT_LANG = 0 means MArathi,
'                SCRIPT_LANG = 1 means Hindi.
'  SPLIT_NUM : This variable is passed to express number in thousands, lakhs, or crores
'              0 : Return the number without any split
'              1 : Return the number in thousands
'              2 : Return the number in lakhs
'              3 : Return the number in crores
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function SUCHI2000_NUM_TO_WORDS Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal NUM As Double, ByVal OpStr As String, ByVal LSCR As Long, ByVal SCRIPT_LANG As Long, ByVal SplitNum As Long) As Long
'Procedure to convert a number in figures to number in words in a given language
'  NUM : Number to be converted
'  OPSTR : Pointer to Output String in Suchika font codes.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : A value indicating the language within the given script. This value
'                has to be specified when a script supports multiple languages.
'                when script = Devnagri, SCRIPT_LANG = 0 means MArathi,
'                SCRIPT_LANG = 1 means Hindi.
'  SPLIT_NUM : This variable is passed to express number in thousands, lakhs, or crores
'              0 : Return the number without any split
'              1 : Return the number in thousands
'              2 : Return the number in lakhs
'              3 : Return the number in crores
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function SUCHI_ISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function ISCII_SUCHI Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI_ISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )}

Declare Function ISCII_SUCHIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII  to Suchika bi-lingual font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SHREE_ISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  This function is provided only for backward compatiblity and is not recommended for
'  newer applications. New applications should make use of the function SHREE2000_ISCII

Declare Function ISCII_SHREE Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII code to Shree-Lipi font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  This function is provided only for backward compatiblity and is not recommended for
'  newer applications. New applications should make use of the function ISCII_SHREE2000

Declare Function SHREE_ISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'  Function to convert string in Shree-Lipi font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  (opstr )

Declare Function ISCII_SHREEEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
' Function to convert string in standard ISCII  to Shree-Lipi font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SHREE_SORT32 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  Note : This function has been provided for backward compatibility and for newer
'  applications it is recommended that you make use of the SHREE2000_SORT32 call.

Declare Function SORT32_SHREE Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in propritory sort code to Shree-Lipi font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  Note : This function has been provided for backward compatibility and for newer
'  applications it is recommended that you make use of the SORT32_SHREE2000 call.

Declare Function SHREE_SORT32EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to Hexadecimal string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SORT32_SHREEEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Hexadecimal to Shree-Lipi font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SUCHI_SORT32 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SORT32_SUCHI Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in propritory sort code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI_SORT32EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to HEx string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SHREE_PCISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  Note : This function has been provided for backward compatibility and for
'  newer applications you should use the SHREE2000_PCISCII function

Declare Function PCISCII_SHREE Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in PC ISCII code to Shree-Lipi font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  Note : This function has been provided for backward compatibility and for
'  newer applications you should use the PCISCII_SHREE2000 function

Declare Function SHREE_PCISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to standard PCISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )}

Declare Function PCISCII_SHREEEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard PCISCII  to Shree-Lipi font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SUCHI_PCISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function PCISCII_SUCHI Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in PC ISCII code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI_PCISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to standard PCISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function PCISCII_SUCHIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard PCISCII  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SHREE_EAISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to extended ISCII code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  Note : This function has been provided for backward compatibility and for newer
'  applications it is recommended that you make use of SHREE2000_EAISCII

Declare Function EAISCII_SHREE Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in extended ISCII code to Shree-Lipi font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
'  Note : This function has been provided for backward compatibility and for newer
'  applications it is recommended that you make use of EAISCII_SHREE2000

Declare Function SHREE_EAISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi font codes to standard EAISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function EAISCII_SHREEEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard EAISCII  to Shree-Lipi font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )
  
Declare Function SUCHI_EAISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to extended ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function EAISCII_SUCHI Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in extended ISCII code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI_EAISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to standard EAISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function EAISCII_SUCHIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard EAISCII  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )


' Shree 2000 Format Compatible functions
Declare Function SHREE2000_ISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi 2000 font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function ISCII_SHREE2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII code to Shree-Lipi 2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (OpStr)

Declare Function SHREE2000_ISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi 2000 font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function ISCII_SHREE2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII code to Shree-Lipi 2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (OpStr)

Declare Function SHREE2000_SORT32 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi2000 font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SORT32_SHREE2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in propritory sort code to Shree-Lipi 2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
  
Declare Function SHREE2000_SORT32EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi2000 font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SORT32_SHREE2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
' Function to convert string in propritory sort code to Shree-Lipi 2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
Declare Function SHREE2000_PCISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi2000 font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function PCISCII_SHREE2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in PC ISCII code to Shree-Lipi2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SHREE2000_PCISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi2000 font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function PCISCII_SHREE2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in PC ISCII code to Shree-Lipi2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
  
Declare Function SHREE2000_EAISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi2000 font codes to extended ISCII code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function EAISCII_SHREE2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in extended ISCII code to Shree-Lipi2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SHREE2000_EAISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Shree-Lipi2000 font codes to extended ISCII code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)


Declare Function EAISCII_SHREE2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in extended ISCII code to Shree-Lipi2000 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI2000_ISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function ISCII_SUCHI2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI2000_ISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function ISCII_SUCHI2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard ISCII  to Suchika bi-lingual font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SUCHI2000_SORT32 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SORT32_SUCHI2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in propritory sort code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)
Declare Function SUCHI2000_SORT32EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to HEx string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SORT32_SUCHI2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Hex  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function SUCHI2000_PCISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function PCISCII_SUCHI2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in PC ISCII code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI2000_PCISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard PCISCII  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function PCISCII_SUCHI2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard PCISCII  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )
  
Declare Function SUCHI2000_EAISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in Suchi bilingual font codes to extended ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function EAISCII_SUCHI2000 Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Integer) As String
'Function to convert string in extended ISCII code to Suchi bilingual font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SUCHI2000_EAISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in Suchika font codes to standard EAISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function EAISCII_SUCHI2000EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal IpStr As String, ByVal OpStr As String, ByRef lSize As Long, ByVal LSCR As Integer) As String
'Function to convert string in standard EAISCII  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

'Date and Time formats

'Format No 00 - dd/mm/yyyy               e.g. 02/02/1999
'Format No 01 - dd/mm/yyyy hh:mi:se a.m. e.g. 02/02/1999 15:4:30 P.M.
'Format No 02 - hh:mi:se a.m.            e.g. 15:4:30 P.M.
'Format No 03 - dd, Month, yyyy          e.g. 2, February, 1999
'Format No 04 - Month dd, yyyy           e.g. February 2, 1999
'Format No 05 - dd-Mon-yy                e.g. 2-Feb-99
'Format No 06 - Month, yy                e.g. February, 99
'Format No 07 - Mon-yy                   e.g. Feb-99
'Format No 08 - dd/mm/yy hh:mi           e.g. 2/2/99 15:4
'Format No 09 - hh:mi                    e.g. 15:4
'Format No 10 - hh:mi:se                 e.g. 15:4:30
'Format No 11 - hh:mi a.m.               e.g. 3:4 P.M.
'Format No 12 - Day, Mon dd, yyyy        e.g. Tuesday, Feb 2, 1999


Declare Sub SHREE_DATETIMETOSTR Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal DateTimeString As String, ByVal OpStr As String, ByRef OUTSIZE As Integer, ByVal AFORMAT As Integer, ByVal LSCR As Integer, ByVal SCR_LANG As Integer)
'Procedure to be called to get different types of date and time format in English
'  and Indian language in Shree Lipi format.
'
'  InDateTimeStr : Pointer to input string containt dat and time converted to
'                  string.
'  OPSTR : Pointer to Output String in Shree-Lipi font codes.
'  OUTSIZE : Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  AFormat : Output format in which you want returned date and time. Different formats
'            are given as per above list.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCR_LANG : A value indicating the language within the given script. This value
'             has to be specified when a script supports multiple languages.
'             when script = Devnagri, SCR_LANG = 0 means Marathi,
'             SCR_LANG = 1 means Hindi.

Declare Sub SHREE2000_DATETIMETOSTR Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal DateTimeString As String, ByVal OpStr As String, ByRef OUTSIZE As Integer, ByVal AFORMAT As Integer, ByVal LSCR As Integer, ByVal SCR_LANG As Integer)
'Procedure to be called to get different types of date and time format in English
'  and Indian language in Shree Lipi format.
'
'  InDateTimeStr : Pointer to input string containt dat and time converted to
'                  string.
'  OPSTR : Pointer to Output String in Shree-Lipi font codes.
'  OUTSIZE : Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  AFormat : Output format in which you want returned date and time. Different formats
'            are given as per above list.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCR_LANG : A value indicating the language within the given script. This value
'             has to be specified when a script supports multiple languages.
'             when script = Devnagri, SCR_LANG = 0 means Marathi,
'             SCR_LANG = 1 means Hindi.

Declare Sub SUCHI_DATETIMETOSTR Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal DateTimeString As String, ByVal OpStr As String, ByRef OUTSIZE As Integer, ByVal AFORMAT As Integer, ByVal LSCR As Integer, ByVal SCR_LANG As Integer)
'Procedure to be called to get different types of date and time format in English
'  and Indian language in Suchika format.
'
'  InDateTimeStr : Pointer to input string containt dat and time converted to
'                  string.
'  OPSTR : Pointer to Output String in Shree-Lipi font codes.
'  OUTSIZE : Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  AFormat : Output format in which you want returned date and time. Different formats
'            are given as per above list.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCR_LANG : A value indicating the language within the given script. This value
'             has to be specified when a script supports multiple languages.
'             when script = Devnagri, SCR_LANG = 0 means Marathi,
'             SCR_LANG = 1 means Hindi.

Declare Sub SUCHI2000_DATETIMETOSTR Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal DateTimeString As String, ByVal OpStr As String, ByRef OUTSIZE As Integer, ByVal AFORMAT As Integer, ByVal LSCR As Integer, ByVal SCR_LANG As Integer)
'Procedure to be called to get different types of date and time format in English
'  and Indian language in Suchi format.
'
'  InDateTimeStr : Pointer to input string containt date and time converted to
'                  string.
'  OPSTR : Pointer to Output String in Suchi font codes.
'  OUTSIZE : Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  AFormat : Output format in which you want returned date and time. Different formats
'            are given as per above list.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCR_LANG : A value indicating the language within the given script. This value
'             has to be specified when a script supports multiple languages.
'             when script = Devnagri, SCR_LANG = 0 means Marathi,
'             SCR_LANG = 1 means Hindi.

Declare Function INIT_CUSTSORT Lib "CNVAPI32.DLL" (ByVal LSCR As Long) As Long
'  Procedure call to initiliaze custom sorting.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.
'            151 : if the script value in not between 1 to 9,18
'            152 : if the custom sorting entries cannot be written in registry
'            153 : if the custom sorting file does not exists
'            154 : if the custom sorting file cannot be opened

Declare Function ISCII_TO_CUST Lib "CNVAPI32.DLL" (ByVal pcInput As String, ByVal pcOutput As String, ByVal LSCR As Long) As String
'  Function to convert ISCII string to Custom sort format suitable for customised
'  sorting
'  pcInput  : Pointer to Input String
'  pcOutput : Pointer to Output String. Sufficient memory must be allocated by the
'             caller before calling this function.
'  lScr     : The Script constant. The Values are interpreted as documented above.
'
'  Returns the pointer to output string (pcOutput)

Declare Function CUST_TO_ISCII Lib "CNVAPI32.DLL" (ByVal pcInput As String, ByVal pcOutput As String, ByVal LSCR As Long) As String
'  Function to convert string in Custom sort format to ISCII
'  pcInput  : Pointer to Input String
'  pcOutput : Pointer to Output String. Sufficient memory must be allocated by the
'             caller before calling this function.
'  lScr     : The Script constant. The Values are interpreted as documented above.
'
'  Returns the pointer to output string (pcOutput)

Declare Function SORT32_TO_CUST Lib "CNVAPI32.DLL" (ByVal pcInput As String, ByVal pcOutput As String, ByVal LSCR As Long) As String
'  Function to convert string in SORT32 format to string in Custom sort format
'  pcInput  : Pointer to Input String
'  pcOutput : Pointer to Output String. Sufficient memory must be allocated by the
'             caller before calling this function.
'  lScr     : The Script constant. The Values are interpreted as documented above.
'
'  Returns the pointer to output string (pcOutput)
  
Declare Function CUST_TO_SORT32 Lib "CNVAPI32.DLL" (ByVal pcInput As String, ByVal pcOutput As String, ByVal LSCR As Long) As String
'  Function to convert string in Custom sort format to string in Sort32 format
'  pcInput  : Pointer to Input String
'  pcOutput : Pointer to Output String. Sufficient memory must be allocated by the
'             caller before calling this function.
'  lScr     : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (pcOutput)
  
Declare Function SET_CUSTFILE Lib "CNVAPI32.DLL" (ByVal pcFileName As String) As Long
'  Function to specify the file containing custom sorting rules.
'  pcFileName  : Pointer to name of the file containing custom sorting rules.
'
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.
'            155 : if the custom sorting entries cannot be written in registry or
'                  if the custom sorting file does not exists
'            156 : if the script set before is not in range 1 to 9, 18
'            157 : if the name of the file cannot be written in the registry

Declare Function SELECT_CUSTFILE Lib "CNVAPI32.DLL" (ByVal LSCR As Long, ByVal pcFileName As String) As Long
' Function to pop up the dialog for selecting the custom sorting file
'  lScr     : The Script constant. The Values are interpreted as documented above.
'  pcFileName  : Pointer to name of the last selected custom sorting file
'
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.
'            158 : if the script set before is not in range 1 to 9, 18
'            159 : if the custom sorting entries cannot be written in registry
'            160 : if the name of the selected file cannot be written in the registry

Declare Function BTAMIL99_DATETIMETOSTR Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal InDateTimeStr As String, ByVal OutStr As String, ByVal OUTSIZE As Long, AFORMAT As Long, ByVal LSCR As Long, ByVal SCR_LANG As Long) As Long
'  Procedure to be called to get different types of date and time format in English
'  and Indian language in Bilingual Tamil 99.
'
'  InDateTimeStr : Pointer to input string containt dat and time converted to
'                  string.
'  OPSTR : Pointer to Output String in Bilingual Tamil 99 font codes.
'  OUTSIZE : Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  AFormat : Output format in which you want returned date and time. Different formats
'            are given as per above list.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCR_LANG : Not Used and must be 0.

Declare Function TAMIL99_DATETIMETOSTR Lib "CNVAPI32.DLL" (ByVal Pass1 As Long, ByVal Pass2 As Long, ByVal InDateTimeStr As String, ByVal OutStr As String, ByVal OUTSIZE As Long, AFORMAT As Long, ByVal LSCR As Long, ByVal SCR_LANG As Long) As Long
'  Procedure to be called to get different types of date and time format in English
'  and Indian language in monolingual Tamil 99.
'
'  InDateTimeStr : Pointer to input string containt dat and time converted to
'                  string.
'  OPSTR : Pointer to Output String in monolingual Tamil 99 font codes.
'  OUTSIZE : Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'  AFormat : Output format in which you want returned date and time. Different formats
'            are given as per above list.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCR_LANG : Not Used and must be 0.

Declare Function BTAMIL99_NUM_TO_WORDS Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal NUM As Double, ByVal OpStr As String, ByVal LSCR As Integer, ByVal SCRIPT_LANG As Integer, ByVal SPLIT_NUM As Integer) As Integer
'  Procedure to convert a number in figures to number in words in a given language
'  NUM : Number to be converted
'  OPSTR : Pointer to Output String in Bilingual Tamil99 font codes.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : Not Used and must be 0.
'  SPLIT_NUM : This variable is passed to express number in thousands, lakhs, or crores
'              0 : Return the number without any split
'              1 : Return the number in thousands
'              2 : Return the number in lakhs
'              3 : Return the number in crores
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function TAMIL99_NUM_TO_WORDS Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal NUM As Double, ByVal OpStr As String, ByVal LSCR As Integer, ByVal SCRIPT_LANG As Integer, ByVal SPLIT_NUM As Integer) As Integer
'  Procedure to convert a number in figures to number in words in a given language
'  NUM : Number to be converted
'  OPSTR : Pointer to Output String in monolingual Tamil99 font codes.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : Not Used and must be 0.
'  SPLIT_NUM : This variable is passed to express number in thousands, lakhs, or crores
'              0 : Return the number without any split
'              1 : Return the number in thousands
'              2 : Return the number in lakhs
'              3 : Return the number in crores
'  Returns : 0 if successful, non zero if any error occurred. The return value is the
'            error code.

Declare Function BTAMIL99_SORT32EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Bilingual tamil 99 font codes to Hexadecimal string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )
'
Declare Function SORT32_BTAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Hexadecimal to Bilingual tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function BTAMIL99_EAISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
' Function to convert string in Bilingual Tamil 99 font codes to standard EAISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function EAISCII_BTAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in standard EAISCII  to Bilingual Tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function BTAMIL99_PCISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Bilingual Tamil99 font codes to standard PCISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function PCISCII_BTAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
' Function to convert string in standard PCISCII  to Bilingual Tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )}

Declare Function BTAMIL99_ISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'Function to convert string in Bilingual Tamil 99 font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  (opstr )}

Declare Function ISCII_BTAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in standard ISCII  to bilingual Tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )}

Declare Function BTAMIL99_SORT32 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
' Function to convert string in Bilingual Tamil 99 font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SORT32_BTAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in propritory sort code to Bilingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function BTAMIL99_EAISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Bilingual Tamil 99 font codes to extended ISCII code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function EAISCII_BTAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in extended ISCII code to Bilingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function BTAMIL99_PCISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Bilingual Tamil 99 font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function PCISCII_BTAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in PC ISCII code to Bilingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function BTAMIL99_ISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Bilingual Tamil 99 font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function ISCII_BTAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in standard ISCII code to Bilingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function TAMIL99_SORT32EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual tamil 99 font codes to Hexadecimal string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )}

Declare Function SORT32_TAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
' Function to convert string in Hexadecimal to Monolingual tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function TAMIL99_EAISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual Tamil 99 font codes to standard EAISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function EAISCII_TAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in standard EAISCII  to Monolingual Tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function TAMIL99_PCISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual Tamil99 font codes to standard PCISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )


Declare Function PCISCII_TAMIL99EX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in standard PCISCII  to Monolingual Tamil 99 font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )

Declare Function TAMIL99_ISCIIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual Tamil 99 font codes to a properietory code suitable
'  for sorting in some 32 bit applications like PowerBuilder 5,6 etc.
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function SORT32_TAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in propritory sort code to Monolingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function TAMIL99_EAISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual Tamil 99 font codes to extended ISCII code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function EAISCII_TAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in extended ISCII code to Monolingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function TAMIL99_PCISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual Tamil 99 font codes to PC ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function PCISCII_TAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in PC ISCII code to Monolingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function TAMIL99_ISCII Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in Monolingual Tamil 99 font codes to standard ISCII string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)

Declare Function ISCII_TAMIL99 Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal LSCR As Long) As String
'  Function to convert string in extended ISCII code to Monolingual Tamil 99 font code string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string (opstr)


Declare Function SORT32_SUCHIEX Lib "CNVAPI32.DLL" (ByVal Pass1 As Integer, ByVal Pass2 As Integer, ByVal IpStr As String, ByVal OpStr As String, ByVal OUTSIZE As Long, ByVal LSCR As Long) As String
'  Function to convert string in Hex  to Suchika font codes string
'  IpStr : Pointer to Input String
'  OpStr : Pointer to Output String. Sufficient memory must be allocated by the
'          caller before calling this function.
'  OutSize :  Size of pointer to output string.
'            The size of the required string is returned in OutSize. The caller
'            should acquire enough memory to hold the returned string. If the size
'            of OutStr is less than the required size, the returned string will be
'            truncated to that size. The size should be indicated by OutSize.
'
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  Returns the pointer to output string  ( opstr )


Public Declare Function INITMSGBOXDLL Lib "MsgBox32" (ByVal LSCR As Long, ByVal SCRIPT_LANG As Long, ByVal pFontName As String, ByVal iFontSize As Long, ByVal TitleFont As Long) As Long
'  This function Initializes the Message Box dll
'  Lscr : The Script constant. The Values are interpreted as documented above.
'  SCRIPT_LANG : A value indicating the language within the given script. This value
'                has to be specified when a script supports multiple languages.
'                when script = Devnagri, SCRIPT_LANG = 0 means MArathi,
'                SCRIPT_LANG = 1 means Hindi. when script = Tamil,
'                SCRIPT_LANG = 0 means Monolingual Tamil 99 layout.
'                SCRIPT_LANG = 1 means Shree Layout.
'                This Parameter is used to retrive the button captions that are
'                displayed on MessageBox. for eg (OK, Cancel , Yes No , etc)
'  pFontName : Font Name that is to be applied to the Message Box Caption.
'  iFontSize : Font Size for the Message Box Caption.
'  TitleFont : 0 if the Message Box Title is to be displayed in English
'            : 1 if the Message Box Title is to be displayed in Indian Language

Public Declare Function SMESSAGEBOX Lib "MsgBox32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'This function is used to display indian language messages.
'  Here the parameters are exactly similar to the Win32 MessageBox API.

Public Declare Sub SETCAPTIONFONTLAYOUT Lib "MsgBox32" (ByVal lfonlayout As Integer)




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
Option Explicit


Public Sub DoOnStart(AShreeChange As Integer)
  If AShreeChange And &H2 <> 0 Then
    'change in the current script
  End If
  If AShreeChange And &H4 <> 0 Then
     'change in the default fontname
  End If
  If AShreeChange And &H8 <> 0 Then
   'change in the default point size
  End If
  If AShreeChange And &H10 <> 0 Then
    'change in the default font attributes
  End If
  If AShreeChange And &H20 <> 0 Then
    'change in the current keyboard layout
  End If
  If AShreeChange And &H80 <> 0 Then
    'change in the assignment of the activation key
  End If
  If AShreeChange And &H100 <> 0 Then
    'change in the font layout change (Shree-Lipi / Suchika)
    'Call REFRESH_TRANS_SETUP(Pass1, Pass2)
  End If
  If AShreeChange And &H200 <> 0 Then
    ' Keyboard tutor is turned on or off
  '  If bTutorOn Then
   '   bTutorOn = False
    '  frmComposing.chktutor.Value = 0
    'End If
  End If
End Sub

Public Function GetDefFontName(FontArr() As Byte) As String
    Dim BString As String * 40
    Dim I As Integer
    I = LBound(FontArr)
    While ((I <= UBound(FontArr)) And (FontArr(I) <> 0))
        BString = BString + Chr(FontArr(I))
        I = I + 1
    Wend
    GetDefFontName = adhTrimNull(BString)
End Function

Public Function GetKeyboardName(KBDArr() As Byte) As String
    Dim BString As String * 15
    Dim I As Integer
    I = LBound(KBDArr)
    While ((I <= UBound(KBDArr)) And (KBDArr(I) <> 0))
        BString = BString + Chr(KBDArr(I))
        I = I + 1
    Wend
    GetKeyboardName = adhTrimNull(BString)
End Function

'Public Function GetScriptFromName(ByVal FntName As String) As String
'  Dim BString As String
'  Dim Retval As Long
' Retval = SHREE_FONTNAME_TO_SCRIPT(FntName)
' Call SHREE_SCRIPT_TO_STR(Retval, BString)
' GetScriptFromName = adhTrimNull(BString)
'End Function

Public Function GetKannadaNumber(ByVal InNumbr As Double) As String
  Dim BString As String * 200
  If lFontType = 1 Then
    Call SUCHI_NUM_TO_WORDS(Pass1, Pass2, InNumbr, BString, glScriptCode, 0, 0)
  ElseIf lFontType = 15 Then
    Call SHREE2000_NUM_TO_WORDS(Pass1, Pass2, InNumbr, BString, glScriptCode, 0, 0)
  ElseIf lFontType = 18 Then
    Call SUCHI2000_NUM_TO_WORDS(Pass1, Pass2, InNumbr, BString, glScriptCode, 0, 0)
  End If
  
  GetKannadaNumber = adhTrimNull(BString)
End Function

Public Function IsciiToSuchi(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 1 Then
    BStr = ISCII_SUCHI(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = ISCII_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  IsciiToSuchi = adhTrimNull(BStr)
End Function

Public Function SuchiToIscii(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 1 Then
    BStr = SUCHI_ISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SUCHI2000_ISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  SuchiToIscii = adhTrimNull(BStr)
End Function

Public Function IsciiToShree(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 0 Then
    BStr = ISCII_SHREE(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = ISCII_SHREE2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  IsciiToShree = adhTrimNull(BStr)
End Function

Public Function ShreeToIscii(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(255)
  If lFontType = 0 Then
    BStr = SHREE_ISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = SHREE2000_ISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SUCHI2000_ISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  ShreeToIscii = adhTrimNull(BStr)
End Function

Public Function Sort32ToSuchi(ByRef AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(255)
  If lFontType = 1 Then
    BStr = SORT32_SUCHI(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SORT32_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  Sort32ToSuchi = adhTrimNull(BStr)
End Function

Public Function SuchiToSort32(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(255)
  If lFontType = 1 Then
    BStr = SUCHI_SORT32(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SUCHI2000_SORT32(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  SuchiToSort32 = adhTrimNull(BStr)
End Function

Public Function Sort32ToShree(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(255)
  If lFontType = 0 Then
    BStr = SORT32_SHREE(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = SORT32_SHREE2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SORT32_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  Sort32ToShree = adhTrimNull(BStr)
End Function

Public Function ShreeToSort32(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(255)
  If lFontType = 0 Then
    BStr = SHREE_SORT32(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = SHREE2000_SORT32(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SUCHI2000_SORT32(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  ShreeToSort32 = adhTrimNull(BStr)
End Function

Public Function PcIsciiToSuchi(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 1 Then
    BStr = PCISCII_SUCHI(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = PCISCII_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  PcIsciiToSuchi = adhTrimNull(BStr)
End Function
Public Function SuchiToPcIscii(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 1 Then
    BStr = SUCHI_PCISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SUCHI2000_PCISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  SuchiToPcIscii = adhTrimNull(BStr)
End Function

Public Function PcIsciiToShree(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 0 Then
    BStr = PCISCII_SHREE(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = PCISCII_SHREE2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  PcIsciiToShree = adhTrimNull(BStr)
End Function

Public Function ShreeToPcIscii(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 0 Then
    BStr = SHREE_PCISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = SHREE2000_PCISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = SUCHI2000_PCISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  ShreeToPcIscii = adhTrimNull(BStr)
End Function

Public Function EAIsciiToSuchi(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 1 Then
    BStr = EAISCII_SUCHI(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = EAISCII_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  EAIsciiToSuchi = adhTrimNull(BStr)
End Function

Public Function SuchiToEAIscii(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 1 Then
    BStr = EAISCII_SUCHI(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = EAISCII_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  SuchiToEAIscii = adhTrimNull(BStr)
End Function

Public Function EAIsciiToShree(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 0 Then
    BStr = EAISCII_SHREE(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = EAISCII_SHREE2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  EAIsciiToShree = adhTrimNull(BStr)
End Function
Public Function ShreeToEAIscii(ByVal AString As String) As String
  Dim AStr As String, BStr As String
  AStr = adhTrimNull(AString)
  BStr = Space$(256)
  If lFontType = 0 Then
    BStr = SHREE_EAISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 15 Then
    BStr = SHREE2000_EAISCII(Pass1, Pass2, AStr, BStr, glScriptCode)
  ElseIf lFontType = 18 Then
    BStr = EAISCII_SUCHI2000(Pass1, Pass2, AStr, BStr, glScriptCode)
  End If
  ShreeToEAIscii = adhTrimNull(BStr)
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

Public Function adhTrimReserveChar(strval As String) As String
Dim intpos As Integer
Dim strval2 As String
Dim strval1 As String
strval2 = Space$(100)
strval2 = strval
intpos = 1
'intpos = Asc("#")
intpos = InStr(intpos, strval2, (Chr(35)))
If intpos > 0 Then
  
 strval1 = adhTrimPipe(adhTrimHash(strval2))
Else
  strval1 = adhTrimPipe(strval2)
End If

  adhTrimReserveChar = Trim$(strval1)
End Function

Public Function adhTrimHash(strval As String) As String
    Dim intpos As Integer, Charcnt As Integer
    Dim strArr() As String
    Dim strval1 As String
    Dim Cnt As Integer
    
    strval1 = Space$(100)
    strval1 = strval
    Charcnt = Len(strval)
    intpos = 1
    Cnt = 0
    ReDim strArr(Charcnt * 2)
    
    While (intpos <= Charcnt) And (Charcnt <> 0)
        Charcnt = Len(Trim(strval1))
        intpos = InStr(intpos, strval1, (Chr(35)))
        If intpos > 0 Then
          strArr(Cnt) = Left$(strval1, intpos - 1)
          Cnt = Cnt + 1
          strArr(Cnt) = "(#)"  'Mid(StrVal1, intPos, 1)
          If strval1 <> Empty Then
            strval1 = Mid(strval1, intpos + 1, Charcnt)
            Cnt = Cnt + 1
            intpos = 1
          End If
        Else
          If strval1 <> Empty Then
            strArr(Cnt) = Mid(strval1, intpos + 1, Charcnt)
            strval1 = ""
            Cnt = Cnt + 1
            intpos = 1
            Charcnt = Len(Trim(strval1))
          End If
        End If

    Wend
    
      For Charcnt = 0 To Cnt - 1
          If Charcnt >= 1 Then
            strval1 = strval1 + strArr(Charcnt)
          Else
            strval1 = strArr(Charcnt)
          End If
      Next
    
       adhTrimHash = Trim$(strval1)
End Function


Public Function adhTrimPipe(strval As String) As String
    ' As Character "|" is a reserved word in Access database, the function is call if there
    ' is pipeline character in string will writing in database
    Dim intpos As Integer, Hashpos As Integer
    Dim Charcnt As Integer
    Dim Cnt As Integer, Posicnt As Integer
    Dim strval1 As String
    Dim strArr() As String
    Dim HashArr() As String
    Dim Searchstr As String
    Charcnt = Len(Trim(strval))
    Cnt = 0
    intpos = 1
    strval1 = Space$(100)
    strval1 = strval
    ReDim strArr(Charcnt * 2)
     
    While (intpos <= Charcnt) And (Charcnt <> 0)
      
      Charcnt = Len(Trim(strval1))
      intpos = InStr(intpos, strval1, Chr(124))
      If intpos > 0 Then
       strArr(Cnt) = "'" + Left$(strval1, intpos - 1) + "'"
       Cnt = Cnt + 1
       strArr(Cnt) = " chr(124) "  'Mid(StrVal1, intPos, 1)
        If strval1 <> Empty Then
          strval1 = Mid(strval1, intpos + 1, Charcnt)
          
          Cnt = Cnt + 1
          intpos = 1
        
        End If
      Else
        If strval1 <> Empty Then
          strArr(Cnt) = "'" + Mid(strval1, intpos + 1, Charcnt) + "'"
          strval1 = ""
          Cnt = Cnt + 1
          intpos = 1
          Charcnt = Len(Trim(strval1))
        End If
     End If

    Wend
    'ReDim StrArr(Charcnt)
      For Charcnt = 0 To Cnt - 1
        If Charcnt >= 1 Then
          If Cnt - 1 = Charcnt Then
            strval1 = strval1 + strArr(Charcnt) + " & " + "'*'"
          Else
            strval1 = strval1 + strArr(Charcnt) + " & "
          End If
        Else
          If Charcnt = Cnt - 1 Then
            strval1 = strArr(Charcnt)
          Else
            strval1 = strArr(Charcnt) + " & "
          End If
        End If
      Next
      adhTrimPipe = Trim$(strval1)
      
End Function

'**************************************************

Public Sub AddBilingualFont(cmbFont As ComboBox)
Dim suchifn As String
Dim shortscr As String
Dim count As Integer

shortscr = Space$(500)
'Get the Shortscript of active script
Call SHREE_SCRIPT_TO_SHORTSTR(glScriptCode, shortscr)
shortscr = Trim(shortscr)
cmbFont.Clear

For count = 0 To Screen.FontCount - 1
    suchifn = InStr(1, Screen.Fonts(count), "SUCHI-" + Mid(shortscr, 1, Len(shortscr) - 1), vbTextCompare)
    If suchifn <> "0" Then
        cmbFont.AddItem Screen.Fonts(count) ' Put each font into list box.
    End If
Next count

If cmbFont.ListCount = 0 Then
  MsgBox ("No Suchi Fonts are available, install the fonts for selected language")
Else
  cmbFont.ListIndex = 0
  
End If
End Sub

Public Sub AddMonoLingualFont(cmbFont As ComboBox)
    Dim suchifn As String
    Dim shortscr As String
    Dim count As Integer
    
    shortscr = Space$(500)
    'Get the Shortscript of active script
    Call SHREE_SCRIPT_TO_SHORTSTR(glScriptCode, shortscr)
    shortscr = Trim(shortscr)
    cmbFont.Clear
    
    For count = 0 To Screen.FontCount - 1
        suchifn = InStr(1, Screen.Fonts(count), "SHREE-" + Mid(shortscr, 1, Len(shortscr) - 1), vbTextCompare)
        If suchifn <> "0" Then
            cmbFont.AddItem Screen.Fonts(count) ' Put each font into list box.
        End If
    Next count
    
    If cmbFont.ListCount = 0 Then
      MsgBox ("No Shree Fonts are available, install the fonts for selected language")
    Else
      cmbFont.ListIndex = 0
    End If

End Sub

Public Sub InitializeSamhita()

'Set Kannada script code
lScript = 7
glScriptCode = 7
lFontType = 0
   
  Pass1 = 73412761
  Pass2 = 651917425
  
If Not SamhitaInitialized Then
      Dim lRetVal As Integer
          'API to Activate Shree Dll
      lRetVal = START_SHREE2(Pass1, Pass2)
      If lRetVal <> 0 Then
         MsgBox "Error Initialising Shree-Samhita : " + CStr(lRetVal)
      End If
        
      'Loading Transliteration Dll
      lRetVal = LOADTRANSLITERATION(Pass1, Pass2)
      If lRetVal <> 0 Then
         MsgBox "Error Initialising transliteration : " + CStr(lRetVal)
      End If
    
    ' Variable Shree Font Layout
     lFontType = 15
    ' Variable Suchi2000 Font Layout
     lFontType = 18
        
     'set the FonttType to suchi2000
     Call SHREE_SETFONTTYPE(Pass1, Pass2, glScriptCode, SUCHI2000)
        
     'set the Keyboard to ENG passing Language Script :Dev, 2: Guj,3:Pun...
     Call SHREE_SET_KEYBOARD(Pass1, Pass2, glScriptCode, "KGP")
        
     'set application type to 0
     Call SHREE_SET_APPLICATION_TYPE(Pass1, Pass2, 0)
          
    Call INIT_CONVERT
     'Call SHREE_FIRSTFONT_FOR_SCRIPT(lScriptCode, strfontname)
     'txtConvToLang.Font.Name = strfontname
     'txtLang.Font.Name = strfontname
     'Call fnSetLabelFont
     SamhitaInitialized = True
     gFontName = "SUCHI-KAN-0850"

Else
 
 gFontName = "SUCHI-KAN-0850"
 'Dim SetUp As New clsSetup
 'gFontName = SetUp.ReadSetupValue("General", "FontName", gFontName)
 'Set SetUp = Nothing

End If
 
 
End Sub
Public Sub Translate(txtLanguage As TextBox, txtEnglish As TextBox)
    Call ToggleWindowsKey(winScrlLock, False)
    If Len(Trim$(txtEnglish.Text)) < 1 And gLangOffSet = wis_KannadaSamhitaOffset Then
        ''Translate teh first name and last name
        txtEnglish.Text = Trim$(ConvertToEnglish(txtLanguage.Text))
    End If

End Sub

Public Function ConvertToEnglish(ByVal KannadaString As String) As String

  ConvertToEnglish = ""
  If Not gLangShree Then Exit Function
  
  If Len(Trim$(KannadaString)) < 1 Then Exit Function
  Dim inputStr As String
  Dim outputStr As String
  On Error GoTo ErrorHandler
  inputStr = adhTrimNull(KannadaString)
  outputStr = Space$(255)
  If inputStr <> "" Then CONVERTLANGTOENG Pass1, Pass2, inputStr, outputStr, glScriptCode
  
  ConvertToEnglish = adhTrimNull(outputStr)
  Exit Function
ErrorHandler:
  MsgBox " Unable to transliterate from Kannada to English"
End Function

Public Function ConvertToKannada(ByVal EnglishString As String) As String
   On Error GoTo ErrorHandler
  ConvertToKannada = ""
  If Len(Trim$(EnglishString)) < 1 Then Exit Function

  Dim inputStr As String
  Dim outputStr As String
  
  inputStr = adhTrimNull(EnglishString)
  outputStr = Space$(255)
  'lScript = 1
  If inputStr <> "" Then CONVERTENGTOLANG Pass1, Pass2, inputStr, outputStr, glScriptCode
  
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

Public Function GetNumberToString(Value As Double) As String

  Dim str1 As String
  str1 = Space$(200)
  'API call to convert number to words
  If lScript = 15 Then
    Call SHREE2000_NUM_TO_WORDS(Pass1, Pass2, Value, str1, glScriptCode, 1, 0)
  Else
    Call SUCHI2000_NUM_TO_WORDS(Pass1, Pass2, Value, str1, glScriptCode, 1, 0)
    End If
  GetNumberToString = str1


End Function

