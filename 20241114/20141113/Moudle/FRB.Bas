Attribute VB_Name = "FRB"
'*********************************************************************************
'      The Declare of FRB.DLL for FRB-100/200 Card
'*********************************************************************************

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*********************************************************/

Global Const FRB_NoError = 0
Global Const FRB_DriverOpenError = 1
Global Const FRB_DriverNoOpen = 2
Global Const FRB_GetDriverVersionError = 3
Global Const FRB_FindBoardError = 4
Global Const FRB_ExceedBoardNumber = 5
Global Const FRB_InputParameterError = 6
Global Const FRB_DelayTimeOut = 7


'//AO Config Code
Global Const FRB_AO_20MA = &H30        '// 0 ~ 20 mA
Global Const FRB_AO_4_20MA = &H31      '// 4 ~ 20 mA
Global Const FRB_AO_UNI_10V = &H32     '// 0 ~ 10 V
Global Const FRB_AO_BI_10V = &H33      '// +/- 10 V
Global Const FRB_AO_UNI_5V = &H34      '// 0 ~  5 V
Global Const FRB_AO_BI_5V = &H35       '// +/-  5 V

'//AI Config Code
Global Const FRB_AI_20MA = &H1A        '// 0 ~ 20 mA
Global Const FRB_AI_4_20MA = &H7       '// 4 ~ 20 mA
Global Const FRB_AI_BI_10V = &H8       '// +/- 10 V
Global Const FRB_AI_BI_5V = &H9        '// +/-  5 V
Global Const FRB_AI_BI_1V = &HA        '// +/-  1 V
Global Const FRB_AI_BI_0V5 = &HB       '// +/-500 mA
Global Const FRB_AI_BI_0V15 = &HC      '// +/-150 mA
Global Const FRB_AI_BI_0V02 = &HD      '// +/- 20 mA

Declare Function FRB_DriverInit Lib "FRB.dll" (wTotalBoards As Integer) As Integer
Declare Sub FRB_DriverClose Lib "FRB.dll" ()

Declare Function FRB_ActiveBoard Lib "FRB.dll" (ByVal wBoardNo As Integer) As Integer
Declare Function FRB_GetDllVersion Lib "FRB.dll" () As Integer
Declare Function FRB_GetDriverVersion Lib "FRB.dll" (wDriverVersion As Integer) As Integer
Declare Function FRB_GetConfigAddressSpace Lib "FRB.dll" (ByVal wBoardNo As Integer, wAddrBase As Long, wIrqNo As Integer, wSubVendor As Integer, wSubDevice As Integer, wSubAux As Integer, wSlotBus As Integer, wSlotDevice As Integer) As Integer

Declare Function FRB_Write_DO Lib "FRB.dll" (ByVal wPort As Integer, ByVal SAn As Integer, ByVal wOutputData As Integer) As Integer
Declare Function FRB_Read_DI Lib "FRB.dll" (ByVal wPort As Integer, ByVal RAn As Integer, ByRef wInputData As Integer) As Integer
Declare Function FRB_DIO_Write_DO Lib "FRB.dll" (ByVal wPort As Integer, ByVal SAn As Integer, ByVal wOutputData As Integer) As Integer
Declare Function FRB_DIO_Read_DI Lib "FRB.dll" (ByVal wPort As Integer, ByVal RAn As Integer, ByRef wInputData As Integer) As Integer


Declare Function FRB_WriteAOHex Lib "FRB.dll" (ByVal wPort As Integer, ByVal SAn As Integer, ByVal wChannel As Integer, ByVal wResolution As Integer, ByVal wValue As Integer, ByVal wDelay As Integer) As Integer
Declare Function FRB_WriteAOFloat Lib "FRB.dll" (ByVal wPort As Integer, ByVal SAn As Integer, ByVal wChannel As Integer, ByVal wResolution As Integer, ByVal wConfigCode As Integer, ByVal fValue As Single, ByVal wDelay As Integer) As Integer

Declare Function FRB_ReadAI_CH Lib "FRB.dll" (ByVal wPort As Integer, ByVal RAn As Integer, ByVal wResolution As Integer, ByVal wConfigCode As Integer, ByVal wChannel As Integer, ByRef fValue As Single) As Integer
Declare Function FRB_ReadAI_CHScan Lib "FRB.dll" (ByVal wPort As Integer, ByVal RAn As Integer, ByVal wResolution As Integer, ByVal wInputType As Integer, ByRef wConfigList As Integer, ByRef fValue As Single) As Integer

Declare Function FRB_SendSA Lib "FRB.dll" (ByVal wPort As Integer, ByVal SAn As Integer, ByVal wOutputData As Integer) As Integer
Declare Function FRB_ReceiveRA Lib "FRB.dll" (ByVal wPort As Integer, ByVal RAn As Integer, wInputData As Integer) As Integer
Declare Function FRB_ReadRAStatus Lib "FRB.dll" (ByVal wPort As Integer, bRAStatus As Byte) As Integer

