Attribute VB_Name = "mdlWindowsAPI"
 Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
 Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
 Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
 Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
 Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Public Const SWP_NOMOVE = &H2      '����ʥثe������m
Public Const SWP_NOSIZE = &H1       '����ʥثe�����j�p
Public Const HWND_TOPMOST = -1     '�]�w���̤W�h
Public Const HWND_NOTOPMOST = -2   '�����̤W�h�]�w
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SWP_HIDEWINDOW = &H80       '���õ���
Public Const SWP_SHOWWINDOW = &H40    '��ܵ���

Public Declare Function SetTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long, _
   ByVal uElapse As Long, _
   ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long) As Long
'=============================================================================================================================
' ����w�q
' ----------------------------
' �������ε{���]�p����(API)�ŧi
' ----------------------------
'get current time by millisecond
Public Declare Function timeGetTime Lib "winmm.dll" () As Long


'=============================================================================================================================
' �ϰ�w�q
' ----------------------------
' �ϰ����ε{���]�p����(API)�ŧi
' ----------------------------
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean








