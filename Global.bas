Attribute VB_Name = "Global"
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Public giComOpenOK(21 To 26) As Integer
Public giAGVOpenOK As Integer

Public gsSQLCon As String
Public gsSQLDSN As String
Public gsSQLUserID As String
Public gsSQLPassword As String
Public gsSQLTable As String
Public gsSQLDbPvd As String

Public gADOCon As ADODB.Connection
Public gADOCmd As ADODB.Command
Public gADORecord As New ADODB.Recordset

Public goTrace As Object
Public mtpyOSVerInfo As OSVERSIONINFO

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const HOST = "1"
Public Const AGV = "2"
Public Const DUMMY = "0"
Public Const READY = "1"
Public Const GET_EMPTY = "1"
Public Const PUT_EMPTY = "2"
Public Const GET_FULL = "3"
Public Const PUT_FULL = "4"
Public Const D001 = "C1   "
Public Const D002 = "C2   "
Public Const D003 = "C3   "
Public Const D004 = "C4   "

Public PLC_LEN(21 To 25) As Integer

Public ST21_Temp(1 To 3) As String    '原紗PLC STATUS TEMP BUFFER
Public ST21_StNo(1 To 3) As String
Public ST21_StByte(1 To 3) As Integer

Public ST22_Temp(1 To 7) As String    '經軸PLC STATUS TEMP BUFFER
Public ST22_StNo(1 To 7) As String
Public ST22_StByte(1 To 7) As Integer

Public ST23_Temp(1 To 6) As String    '織軸PLC STATUS TEMP BUFFER
Public ST23_StNo(1 To 6) As String
Public ST23_StByte(1 To 6) As Integer

Public ST24_Temp(1 To 5) As String    '胚布PLC STATUS TEMP BUFFER
Public ST24_StNo(1 To 5) As String
Public ST24_StByte(1 To 5) As Integer

Public ST25_Temp(1 To 11) As String    '成品PLC STATUS TEMP BUFFER
Public ST25_StNo(1 To 11) As String
Public ST25_StByte(1 To 11) As Integer

Public AGV_InBuf As String * 100
Public Send_AGV_Str() As Byte
Public AGV_TempStr As String
Public BCC As Byte
Public AGV_Port As Integer

' Non-registry API structures
Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
