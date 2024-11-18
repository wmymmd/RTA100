Attribute VB_Name = "Module2"
Option Explicit

' API ÉùÃ÷
 Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
 Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const INFINITE = &HFFFFFFFF

Public Sub MyThreadProc(ByVal lParam As Long)
     frmPlotProcessLog.Show
            frmPlotProcessLog.ZOrder
            frmPlotProcessLog.fraProcessHistory.Visible = True
            frmPlotProcessLog.fraProcessHistory.ZOrder
End Sub
