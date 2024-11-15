Attribute VB_Name = "ICPDAS_USBIO_VB_Lib"
Option Explicit

Global Const USBIO_MAX_SUPPORT_TYPE = 50
Global Const USBIO_DI_MAX_CHANNEL = 16
Global Const USBIO_DO_MAX_CHANNEL = 16
Global Const USBIO_AI_MAX_CHANNEL = 16
Global Const USBIO_AO_MAX_CHANNEL = 16
Global Const USBIO_PI_MAX_CHANNEL = 16
Global Const USBIO_NICKNAME_LENGTH = 32
Global Const USBIO_SN_LENGTH = 32

'--------------------------------------------------
' Base Error Code Offset
'--------------------------------------------------
Global Const DEV_RETURN_ERR_CODE_BASE = 0   'No Error
Global Const DEVLIB_ERR_CODE_BASE = &H10000 'The Module Name Error
Global Const IOLIB_ERR_CODE_BASE = &H11000  'The Module doesn't exist in this Port

'--------------------------------------------------
' Error Code
'--------------------------------------------------
Global Const ERR_NO_ERR = 0                                         'No Error

Global Const ERR_USBDEV_INVALID_DEV = DEVLIB_ERR_CODE_BASE + 0      'The device is invalid
Global Const ERR_USBDEV_DEV_OPENED = DEVLIB_ERR_CODE_BASE + 1       'The device has already opened
Global Const ERR_USBDEV_DEVNOTEXISTS = DEVLIB_ERR_CODE_BASE + 2     'The device does not exists
Global Const ERR_USBDEV_GETDEVINFO = DEVLIB_ERR_CODE_BASE + 3       'An error occurs while getting device information
Global Const ERR_USBDEV_ERROR_PKTSIZE = DEVLIB_ERR_CODE_BASE + 4    'The packet size is invalid
Global Const ERR_USBDEV_ERROR_WRITEFILE = DEVLIB_ERR_CODE_BASE + 5  'An error occurs while writing packet to module

Global Const ERR_USBIO_COMM_TIMEOUT = IOLIB_ERR_CODE_BASE + 0       'Communication timeout
Global Const ERR_USBIO_DEV_OPENED = IOLIB_ERR_CODE_BASE + 1         'The device has already opened
Global Const ERR_USBIO_DEV_NOTOPEN = IOLIB_ERR_CODE_BASE + 2        'The device has not opened
Global Const ERR_USBIO_INVALID_RESP = IOLIB_ERR_CODE_BASE + 3       'The returning command is invalid
Global Const ERR_USBIO_IO_NOTSUPPORT = IOLIB_ERR_CODE_BASE + 4      'The function for this device is not supported
Global Const ERR_USBIO_PARA_ERROR = IOLIB_ERR_CODE_BASE + 5         'The parameter error
Global Const ERR_USBIO_BULKVALUE_ERR = IOLIB_ERR_CODE_BASE + 6      'An error occurs while getting bulk value
Global Const ERR_USBIO_GETDEVINFO = IOLIB_ERR_CODE_BASE + 7         'An error occurs while getting device information

Global Const USBIO_MINPID = &H400
Global Const USB2019 = USBIO_MINPID + 19
Global Const USB2026 = USBIO_MINPID + 26
Global Const USB2045 = USBIO_MINPID + 45
Global Const USB2051 = USBIO_MINPID + 51
Global Const USB2055 = USBIO_MINPID + 55
Global Const USB2060 = USBIO_MINPID + 60
Global Const USB2064 = USBIO_MINPID + 64
Global Const USB2084 = USBIO_MINPID + 84

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                                    ByRef Source As Any, ByVal Length As Long)

'System
Declare Function OpenDevice Lib "ICPDAS_USBIO_VB.dll" (ByVal i_wUSB_DID As Integer, _
                                                       ByVal i_byUSB_BID As Byte) As Long
Declare Function CloseDevice Lib "ICPDAS_USBIO_VB.dll" () As Long

'Get device information
Declare Function SetCommTimeout Lib "ICPDAS_USBIO_VB.dll" (ByVal i_dwCommTimeout As Long) As Long
Declare Function GetCommTimeout Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwCommTimeout As Long) As Long
Declare Function SetAutoResetWDT Lib "ICPDAS_USBIO_VB.dll" (ByVal i_bEnable As Boolean) As Long
Declare Function RefreshDeviceInfo Lib "ICPDAS_USBIO_VB.dll" () As Long
Declare Function GetSoftWDTTimeout Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwSoftWDTTimeout As Long) As Long
Declare Function GetDeviceID Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwDeviceID As Long) As Long
Declare Function GetFwVer Lib "ICPDAS_USBIO_VB.dll" (ByRef o_wFwVer As Integer) As Long
Declare Function GetDeviceNickName Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byDeviceNickName As Byte) As Long
Declare Function GetDeviceSN Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byDeviceSN As Byte) As Long
Declare Function GetSupportIOMask Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySupportIOMask As Byte) As Long
Declare Function GetDITotal Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byDITotal As Byte) As Long
Declare Function GetDOTotal Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byDOTotal As Byte) As Long
Declare Function GetAITotal Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byAITotal As Byte) As Long
Declare Function GetAOTotal Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byAOTotal As Byte) As Long
Declare Function GetPITotal Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byPITotal As Byte) As Long
Declare Function GetPOTotal Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byPOTotal As Byte) As Long

'Set device information
Declare Function SetUserDefinedBoardID Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byBID As Byte) As Long
Declare Function SetDeviceNickName Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byDeviceNickName As String) As Long
Declare Function SetSoftWDTTimeout Lib "ICPDAS_USBIO_VB.dll" (ByVal i_dwSoftWDTTimeout As Long) As Long
Declare Function LoadDefault Lib "ICPDAS_USBIO_VB.dll" () As Long

'Callback registration
Declare Function RegisterEmergencyPktEventHandle Lib "ICPDAS_USBIO_VB.dll" (ByVal i_evtHandle As Long) As Long

'Get DO configuration and data
Declare Function DO_GetPowerOnEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byPowerOnEnable As Byte) As Long
Declare Function DO_GetSafetyEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySafetyEnable As Byte) As Long
Declare Function DO_GetSafetyValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySafetyValue As Byte) As Long
Declare Function DO_GetDigitalOutputInverse Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwInverse As Long) As Long
Declare Function DO_ReadValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byDOValue As Byte) As Long
'Set DO configuration and data
Declare Function DO_SetPowerOnEnable Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                                ByVal i_byPowerOnEnable As Byte) As Long
Declare Function DO_SetPowerOnEnables Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byPowerOnEnables As Byte) As Long
Declare Function DO_SetSafetyEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef i_bySafetyEnable As Byte) As Long
Declare Function DO_SetSafetyValue Lib "ICPDAS_USBIO_VB.dll" (ByRef i_bySafetyValue As Byte) As Long
Declare Function DO_SetDigitalOutputInverse Lib "ICPDAS_USBIO_VB.dll" (ByVal i_dwInverse As Long) As Long
Declare Function DO_WriteValue Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byDOValue As Byte) As Long
Declare Function DO_WriteChannelValue Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChannel As Byte, ByVal i_byValue As Byte) As Long
'Get DI configuration and data
Declare Function DI_GetDigitalFilterWidth Lib "ICPDAS_USBIO_VB.dll" (ByRef o_wFilterWidth As Integer) As Long
Declare Function DI_GetDigitalValueInverse Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwInverse As Long) As Long
Declare Function DI_GetCntEdgeTrigger Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwEdgeTrig As Long) As Long
Declare Function DI_ReadValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byDIValue As Byte) As Long
Declare Function DI_ReadCounterValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwDICntValue As Long) As Long
'Set DI configuration and data
Declare Function DI_SetDigitalFilterWidth Lib "ICPDAS_USBIO_VB.dll" (ByVal i_wFilterWidth As Integer) As Long
Declare Function DI_SetDigitalValueInverse Lib "ICPDAS_USBIO_VB.dll" (ByVal i_dwInverse As Long) As Long
Declare Function DI_SetCntEdgeTrigger Lib "ICPDAS_USBIO_VB.dll" (ByVal i_dwEdgeTrig As Long) As Long
Declare Function DI_WriteClearCounter Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToClr As Byte) As Long
Declare Function DI_WriteClearCounters Lib "ICPDAS_USBIO_VB.dll" (ByVal i_dwCntClrMask As Long) As Long
'Get AI configuration and data
Declare Function AI_GetTotalSupportType Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTotalSupportType As Byte) As Long
Declare Function AI_GetSupportTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySupportTypeCode As Byte) As Long
Declare Function AI_GetTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTypeCode As Byte) As Long
Declare Function AI_GetChCJCOffset Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fChCJCOffset As Single) As Long
Declare Function AI_GetChEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byChEnable As Byte) As Long
Declare Function AI_GetFilterRejection Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byFilterRejection As Byte) As Long
Declare Function AI_GetCJCOffset Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fCJCOffset As Single) As Long
Declare Function AI_GetCJCEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byCJCEnable As Byte) As Long
Declare Function AI_GetWireDetectEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byWireDetectEnable As Byte) As Long
Declare Function AI_GetResolution Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byResolution As Byte) As Long
Declare Function AI_ReadValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwAIValue As Long) As Long
Declare Function AI_ReadValueDigitalWithChSta Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwAIValue As Long, _
                                                                         ByRef o_byAIChStatus() As Byte) As Long
Declare Function AI_ReadValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fAIValue As Single) As Long
Declare Function AI_ReadValueAnalogWithChSta Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fAIValue As Single, _
                                                                        ByRef o_byAIChStatus() As Byte) As Long
Declare Function AI_ReadBulkValue Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byStartCh As Byte, _
                                                             ByVal i_byChTotal As Byte, _
                                                             ByVal i_dwSampleWidth As Long, _
                                                             ByVal i_fSampleRate As Single, _
                                                             ByVal i_dwBufferWidth As Long, _
                                                             ByRef o_dwDataBuffer As Long, _
                                                             ByVal i_CBFunc As Long) As Long
Declare Function AI_ReadCJCValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fCJCValue As Single) As Long

'Set AI configuration
Declare Function AI_SetTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                           ByVal i_byTypeCode As Byte) As Long
Declare Function AI_SetTypeCodes Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byTypeCodes As Byte) As Long
Declare Function AI_SetChCJCOffset Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                              ByVal i_fChCJCOffset As Single) As Long
Declare Function AI_SetChCJCOffsets Lib "ICPDAS_USBIO_VB.dll" (ByRef i_fChCJCOffsets As Single) As Long
Declare Function AI_SetChEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byChEnable As Byte) As Long
Declare Function AI_SetFilterRejection Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byFilterRejection As Byte) As Long
Declare Function AI_SetCJCOffset Lib "ICPDAS_USBIO_VB.dll" (ByVal i_fCJCOffset As Single) As Long
Declare Function AI_SetCJCEnable Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byCJCEnable As Byte) As Long
Declare Function AI_SetWireDetectEnable Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byWireDetectEnable As Byte) As Long

'Get AO configuration and data
Declare Function AO_GetTotalSupportType Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTotalSupportType As Byte) As Long
Declare Function AO_GetSupportTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySupportTypeCode As Byte) As Long
Declare Function AO_GetTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTypeCode As Byte) As Long
Declare Function AO_GetChEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byChEnable As Byte) As Long
Declare Function AO_GetResolution Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byResolution As Byte) As Long
Declare Function AO_ReadExpValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwAOExpValue As Long) As Long
Declare Function AO_ReadExpValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fAOExpValue As Single) As Long
Declare Function AO_ReadCurValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwAOCurValue As Long) As Long
Declare Function AO_ReadCurValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fAOCurValue As Single) As Long
Declare Function AO_GetPowerOnEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byPowerOnEnable As Byte) As Long
Declare Function AO_GetSafetyEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySafetyEnable As Byte) As Long
Declare Function AO_GetPowerOnValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwPwrOnValue As Long) As Long
Declare Function AO_GetPowerOnValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fPwrOnValue As Single) As Long
Declare Function AO_GetSafetyValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwSafetyValue As Long) As Long
Declare Function AO_GetSafetyValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fSafetyValue As Single) As Long
Declare Function AO_GetSlewRate Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySlewRate As Byte) As Long
'Set AO configuration and data
Declare Function AO_SetTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                           ByVal i_byTypeCode As Byte) As Long
Declare Function AO_SetTypeCodes Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byTypeCodes As Byte) As Long
Declare Function AO_SetChEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byChEnable As Byte) As Long
Declare Function AO_WriteChannelValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                          ByVal i_dwAOVal As Long) As Long
Declare Function AO_WriteValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef i_dwAOValue As Long) As Long
Declare Function AO_WriteChannelValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                          ByVal i_fAOExpValue As Single) As Long
Declare Function AO_WriteValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef i_fAOExpValue As Single) As Long
Declare Function AO_SetPowerOnEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byPowerOnEnable As Byte) As Long
Declare Function AO_SetSafetyEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef i_bySafetyEnable As Byte) As Long
Declare Function AO_SetPowerOnValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef i_dwPwrOnValue As Long) As Long
Declare Function AO_SetPowerOnChannelValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                               ByVal i_dwPwrOnValue As Long) As Long
Declare Function AO_SetPowerOnValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef i_fPwrOnValue As Single) As Long
Declare Function AO_SetPowerOnChannelValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                               ByVal i_fPwrOnValue As Single) As Long
Declare Function AO_SetSafetyValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByRef i_dwSafetyValue As Long) As Long
Declare Function AO_SetSafetyChannelValueDigital Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                              ByVal i_dwSafetyValue As Long) As Long
Declare Function AO_SetSafetyValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByRef i_fSafetyValue As Single) As Long
Declare Function AO_SetSafetyChannelValueAnalog Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                              ByVal i_fSafetyValue As Single) As Long
Declare Function AO_SetSlewRate Lib "ICPDAS_USBIO_VB.dll" (ByRef i_bySlewRate As Byte) As Long

'Get PI configuration and data
Declare Function PI_GetTotalSupportType Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTotalSupportType As Byte) As Long
Declare Function PI_GetSupportTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_bySupportTypeCode As Byte) As Long
Declare Function PI_GetTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTypeCode As Byte) As Long
Declare Function PI_GetTriggerMode Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byTriggerMode As Byte) As Long
Declare Function PI_GetChIsolatedFlag Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byChIsolatedFlag As Byte) As Long
Declare Function PI_GetLPFilterEnable Lib "ICPDAS_USBIO_VB.dll" (ByRef o_byLPFilterEnable As Byte) As Long
Declare Function PI_GetLPFilterWidth Lib "ICPDAS_USBIO_VB.dll" (ByRef o_wLPFilterWidth As Integer) As Long
Declare Function PI_ReadValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwPIValue As Long, _
                                                         ByRef o_byChStatus As Byte) As Long
Declare Function PI_ReadCntValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_dwCnyValue As Long, _
                                                            ByRef o_byChStatus As Byte) As Long
Declare Function PI_ReadFreqValue Lib "ICPDAS_USBIO_VB.dll" (ByRef o_fFreqValue As Single, _
                                                             ByRef o_byChStatus As Byte) As Long
Declare Function PI_ReadBulkValue Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byStartCh As Byte, _
                                                             ByVal i_byChTotal As Byte, _
                                                             ByVal i_dwSampleWidth As Long, _
                                                             ByVal i_fSampleRate As Single, _
                                                             ByVal i_dwBufferWidth As Long, _
                                                             ByRef o_dwDataBuffer As Long, _
                                                             ByVal i_CBFunc As Long) As Long
'Set PI configuration
Declare Function PI_SetTypeCode Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                           ByVal i_byTypeCode As Byte) As Long
Declare Function PI_SetTypeCodes Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byTypeCodes As Byte) As Long
Declare Function PI_ClearSingleChCount Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToClr As Byte) As Long
Declare Function PI_ClearChCount Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byClrMask As Byte) As Long
Declare Function PI_ClearSingleChStatus Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToClr As Byte) As Long
Declare Function PI_ClearChStatus Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byClrMask As Byte) As Long
Declare Function PI_SetTriggerMode Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                              ByVal i_byTriggerMode As Byte) As Long
Declare Function PI_SetTriggerModes Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byTriggerModes As Byte) As Long
Declare Function PI_SetChIsolatedFlag Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                                 ByVal i_bChIsolatedFlag As Boolean) As Long
Declare Function PI_SetChIsolatedFlags Lib "ICPDAS_USBIO_VB.dll" (ByRef i_bChIsolatedFlags As Byte) As Long
Declare Function PI_SetLPFilterEnable Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                                 ByVal i_bLPFilterEnable As Boolean) As Long
Declare Function PI_SetLPFilterEnables Lib "ICPDAS_USBIO_VB.dll" (ByRef i_byLPFilterEnables As Byte) As Long
Declare Function PI_SetLPFilterWidth Lib "ICPDAS_USBIO_VB.dll" (ByVal i_byChToSet As Byte, _
                                                                ByVal i_wLPFilterWidth As Integer) As Long
Declare Function PI_SetLPFilterWidths Lib "ICPDAS_USBIO_VB.dll" (ByRef i_wLPFilterWidths As Integer) As Long

Public Sub StrToByte(ByVal StrData As String, ByRef bytData As Byte)

    On Error GoTo ErrHandle:
   
    Dim i As Long, StrDataLen As Integer
    Dim TmpByt As Byte, TmpBytData(1024) As Byte
   
    StrDataLen = Len(StrData)
    If StrDataLen > 1024 Then
        StrDataLen = 1024
    End If
   
    For i = 0 To (StrDataLen - 1)
        TmpByt = Asc(Mid(StrData, i + 1, 1))
            TmpBytData(i) = TmpByt
    Next
       
    Call CopyMemory(bytData, TmpBytData(0), StrDataLen)
    Exit Sub
       
       
ErrHandle:
    MsgBox Err.Description
   
End Sub

Public Function AscStrToLong(ByVal StrAscData As String) As Long

    On Error GoTo ErrHandle:
    
    Dim i As Long, StrDataLen As Integer
    Dim TmpByt As Byte
    Dim dwRetVal As Long
   
    StrDataLen = Len(StrAscData)
    If StrDataLen > 1024 Then
        StrDataLen = 1024
    End If
   
    dwRetVal = 0
    For i = 0 To (StrDataLen - 1)
        TmpByt = Asc(Mid(StrAscData, i + 1, 1))
        
        If TmpByt >= Asc("A") And TmpByt <= Asc("F") Then
            TmpByt = 10 + TmpByt - Asc("A")
        ElseIf TmpByt >= Asc("a") And TmpByt <= Asc("f") Then
            TmpByt = 10 + TmpByt - Asc("a")
        ElseIf TmpByt >= Asc("0") And TmpByt <= Asc("9") Then
            TmpByt = TmpByt - Asc("0")
        Else
            AscStrToLong = 0
            Exit Function
        End If
        
        If i = (StrDataLen - 1) Then
            dwRetVal = dwRetVal + TmpByt
        Else
            dwRetVal = dwRetVal + TmpByt * (16 * (StrDataLen - 1) - i)
        End If
    Next
    
    AscStrToLong = dwRetVal
    Exit Function
    
ErrHandle:
    MsgBox Err.Description

End Function

