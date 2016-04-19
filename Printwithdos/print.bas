Attribute VB_Name = "basPrint"
Option Explicit
Public wis_MESSAGE_TITLE As String
Public gCompanyName As String
Public Const wisGray = &H80000000
Public Const wisWhite = &H80000005
'Shashi 4/12/2000
Public Const vbWhite = &H80000005    '&H80000005&
' Status variable constants...
Public Const wis_CANCEL = 0
Public Const wis_FAILURE = 0
Public Const wis_OK = 1
Public Const wis_SUCCESS = 2
Public Const wis_COMPLETE = 3
Public Const wis_EVENT_SUCCESS = 4
Public Const wis_SHOW_FIRST = 5
Public Const wis_SHOW_PREVIOUS = 6
Public Const wis_SHOW_NEXT = 7
Public Const wis_SHOW_LAST = 8
Public Const wis_PRINT_CURRENT = 9
Public Const wis_PRINT_ALL = 10
Public Const wis_PRINT_ALL_PAUSE = 11
Public Const wis_PRINT_CURRENT_PAUSE = 12
Public Const wis_Print_Excel = 13
Public Const wis_SHOW_PAGE = 14
Public Const wis_SHOW_REFRESH = 15


Public gFont As StdFont



Public Function PrintToExplorer(grd As MSFlexGrid)

End Function


