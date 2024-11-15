Attribute VB_Name = "mdlWindowsAPI"
 Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
 Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
 Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
 Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
 Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Public Const SWP_NOMOVE = &H2      '不更動目前視窗位置
Public Const SWP_NOSIZE = &H1       '不更動目前視窗大小
Public Const HWND_TOPMOST = -1     '設定為最上層
Public Const HWND_NOTOPMOST = -2   '取消最上層設定
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SWP_HIDEWINDOW = &H80       '隱藏視窗
Public Const SWP_SHOWWINDOW = &H40    '顯示視窗

Public Declare Function SetTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long, _
   ByVal uElapse As Long, _
   ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long) As Long
'=============================================================================================================================
' 全域定義
' ----------------------------
' 全域應用程式設計介面(API)宣告
' ----------------------------
'get current time by millisecond
Public Declare Function timeGetTime Lib "winmm.dll" () As Long


'=============================================================================================================================
' 區域定義
' ----------------------------
' 區域應用程式設計介面(API)宣告
' ----------------------------
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean








