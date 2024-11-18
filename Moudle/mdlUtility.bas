Attribute VB_Name = "mdlUtility"

Option Explicit

Public TimeStart(5) As Single
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    X As Long
    Y As Long
End Type

'Declare Function ClipCursor Lib "user32" (ByVal lpRect As RECT) As Long
'Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Declare Function GetWindowRect Lib "user32" (ByVal hwnd As IntPtr, ByVal lpRect As RECT) As Long

Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public RetValue As Long
Public ClipMode As Boolean

Public Const GB_PI = 3.14159265358979

Public Function CheckStringIsNumber(strNumber As String) As Boolean
    Dim dblValue As Double
    On Error GoTo ERR_NUMBER
    
    dblValue = Val(strNumber)
    CheckStringIsNumber = True
    Exit Function
ERR_NUMBER:
    CheckStringIsNumber = False
End Function

Public Function CheckValueScale(strNumber As String) As Long
    Dim iDotPos As Integer
    Dim iLen As Integer
    Dim strTemp As String
    
    strTemp = CStr(Val(strNumber))
    If InStr(1, strTemp, ".") <> 0 Then
        strTemp = Mid(strTemp, 1, InStr(1, strTemp, ".") - 1)
        If Val(strTemp) < 1 Then
            strTemp = Mid(strNumber, InStr(1, strNumber, ".") + 1)
            CheckValueScale = 10 ^ (Len(strTemp) + 2)
            Exit Function
        End If
    Else
        strTemp = Mid(strTemp, 1)
    End If
    
    iLen = Len(strTemp)
    If (iLen > 3) Then iLen = 3
    CheckValueScale = 10 ^ (3 - iLen)
End Function

Public Function SetTime(Index As Integer) As Boolean
    TimeStart(Index) = timeGetTime

End Function

Public Function GetTime(Index As Integer) As Single
    
    GetTime = timeGetTime - TimeStart(Index)
End Function

Public Function DelayTime(Delay As Long) As Boolean
   Dim lngFlag As Long
   
   lngFlag = Timer
   
   While (Timer - lngFlag) < Delay
      DoEvents
   Wend
End Function

Public Function FileExists(TheFile As String) As Boolean
    Dim Results As String
    
On Error GoTo ERR_NUMBER
    Results = dir$(TheFile)
    FileExists = IIf(Results = "", False, True)
    Exit Function
ERR_NUMBER:
    FileExists = False
End Function




