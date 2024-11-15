VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdvThermo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DeviceNum As Long
Public DeviceHandle As Long
Public ErrCde As Long
Public lpSubDeviceList As Long
Public dwDeviceNum As Long
Public collectionTCVoltageRange As New Collection



Private ptTCMuxRead As PT_TCMuxRead
Private ptDevGetFeatures As PT_DeviceGetFeatures
Private lpDevFeatures As DEVFEATURES
Private ptAIRead As PT_AIVoltageIn
Private ptAOWrite As PT_AOVoltageOut
Private ptDioReadPort As PT_DioReadPortByte
Private lpAIConfig As PT_AIConfig


'for filter
Private TC_Filter(16, 100) As Single
Public TC_FilterCount As Integer
Public TC_FilterCountMax As Integer

Public IsActive As Boolean

Public TC_CVT1 As Single
Public TC_CVT2 As Single
Public DO1 As Byte
Public DO2 As Byte






Public Function InitialCard() As Long
    Dim i As Integer
    Dim ii As Integer
    Dim tempStr As String
    Dim lpDeviceList As Long
    
    ' retrieve the information of all installed devices
    lpDeviceList = DRV_GetAddress(devicelist(0))
    ErrCde = DRV_DeviceGetList(lpDeviceList, 255, 1)
    If (ErrCde <> 0) Then
        DRV_GetErrorMessage ErrCde, szErrMsg
        'Call MsgBox(szErrMsg, vbOKOnly, "Error!!")
        InitialCard = ErrCde
        Exit Function
    End If
    DeviceNum = devicelist(0).dwDeviceNum
    ErrCde = DRV_DeviceOpen(DeviceNum, DeviceHandle)
    If ErrCde <> 0 Then GoTo RETURNFAIL
    
    ptDevGetFeatures.buffer = DRV_GetAddress(lpDevFeatures)
    ErrCde = DRV_DeviceGetFeatures(DeviceHandle, ptDevGetFeatures)
    If ErrCde <> 0 Then GoTo RETURNFAIL
      If (lpDevFeatures.usNumGain > 0) Then
        For i = 0 To collectionTCVoltageRange.Count - 1
            collectionTCVoltageRange.Remove (i)
        Next i
        For i = 0 To (lpDevFeatures.usNumGain - 1)
          tempStr = ""
          For ii = 0 To 15
              tempStr = tempStr + Chr(lpDevFeatures.glGainList(i).szGainStr(ii))
          Next ii
          collectionTCVoltageRange.Add (tempStr)
        Next i
    End If
    With ptTCMuxRead
        .DasChan = 2
        .DasGain = lpDevFeatures.glGainList(2).usGainCde '11 ' for R type '2 for kType
        '.DasGain = lpDevFeatures.glGainList(11).usGainCde '11 ' for R type '2 for kType
'        If gbintTCType = 5 Then
'            .DasGain = lpDevFeatures.glGainList(11).usGainCde '11 ' for R type '2 for kType
'        Else
'            .DasGain = lpDevFeatures.glGainList(2).usGainCde '11 ' for R type '2 for kType
'        End If
        '.DasGain = lpDevFeatures.glGainList(gbintTCVoltageRange).usGainCde '11 ' for R type '2 for kType
        .ExpChan = 0
        .TCType = gbintTCType '5 '5:Rtype,1:KType
        .TempScale = 0
    End With
    IsActive = True
    gbblnActiveAD = True
    For i = 0 To 15
        lpAIConfig.DasChan = i
        lpAIConfig.DasGain = 0
        ErrCde = DRV_AIConfig(DeviceHandle, lpAIConfig)
    Next i
    
    TC_CVT1 = gbsngTCCVT1
    TC_CVT2 = gbsngTCCVT2
    Exit Function
RETURNFAIL:
    gbblnActiveAD = False
    gbstrAlarmHint = " Initial Advan TC error"
    ShowAlarmFlash 1
    InitialCard = ErrCde
End Function

Private Sub Class_Initialize()
    IsActive = False
    TC_FilterCount = 0
    TC_FilterCountMax = 30
    TC_CVT1 = 0
    TC_CVT2 = 0
End Sub

Private Sub Class_Terminate()
    DRV_ClearOverrun (DeviceHandle)
    DRV_DeviceClose (DeviceHandle)
    
End Sub

Public Function ReadTemperature(iChannel As Integer, fTemp As Single) As Long
    ptTCMuxRead.DasChan = iChannel
    ptTCMuxRead.temp = DRV_GetAddress(fTemp)
    ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
    
    ReadTemperature = ErrCde
End Function

Public Function ReadTemperatureAllChannel(fTemp() As Single) As Long
    Dim i As Integer
    Dim idx As Integer
    
    idx = 0
    For i = 2 To 14 Step 2
        ptTCMuxRead.DasChan = i
        ptTCMuxRead.temp = DRV_GetAddress(fTemp(idx))
        ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
        idx = idx + 1
        If (ErrCde <> 0) Then Exit For
    Next i
    ReadTemperatureAllChannel = ErrCde
End Function

Public Function ReadTemperatureAllChannelFT(fTemp() As Single) As Long
    Dim i As Integer
    Dim idx As Integer
    Dim st(10) As Single
    
    On Error GoTo ERRLINE
    
    idx = 0
    Select Case Para.RtaType
    Case 1
        For i = 2 To 14 Step 2
            ptTCMuxRead.DasChan = i
            ptTCMuxRead.temp = DRV_GetAddress(fTemp(idx))
            ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
            idx = idx + 1
            If (ErrCde <> 0) Then Exit For
        Next i
        If (ErrCde = 0) Then
            Call GetFilterTemperature(6, fTemp, st)
            For i = 0 To 6
                fTemp(i) = st(i)
            Next i
        End If
        
    Case 2
        For i = 2 To 2 Step 2
            ptTCMuxRead.DasChan = i
            ptTCMuxRead.temp = DRV_GetAddress(fTemp(idx))
            ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
            idx = idx + 1
            If (ErrCde <> 0) Then Exit For
        Next i
        If (ErrCde = 0) Then
            Call GetFilterTemperature(1, fTemp, st)
            fTemp(0) = st(0)
        End If
    
    Case 6
        For i = 2 To 14 Step 2
            ptTCMuxRead.DasChan = i
            ptTCMuxRead.temp = DRV_GetAddress(fTemp(idx))
            ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
            idx = idx + 1
            If (ErrCde <> 0) Then Exit For
        Next i
        If (ErrCde = 0) Then
            Call GetFilterTemperature(6, fTemp, st)
            For i = 0 To 6
                fTemp(i) = st(i)
            Next i
        End If
            
    End Select
    
    
'    If Para.RtaType = 1 Or Para.RtaType = 3 Then
'        For i = 2 To 14 Step 2
'            ptTCMuxRead.DasChan = i
'            ptTCMuxRead.temp = DRV_GetAddress(fTemp(idx))
'            ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
'            idx = idx + 1
'            If (ErrCde <> 0) Then Exit For
'        Next i
'        If (ErrCde = 0) Then
'            Call GetFilterTemperature(6, fTemp, st)
'            For i = 0 To 6
'                fTemp(i) = st(i)
'            Next i
'        End If
'    Else
'
'        For i = 2 To 2 Step 2
'            ptTCMuxRead.DasChan = i
'            ptTCMuxRead.temp = DRV_GetAddress(fTemp(idx))
'            ErrCde = DRV_TCMuxRead(DeviceHandle, ptTCMuxRead)
'            idx = idx + 1
'            If (ErrCde <> 0) Then Exit For
'        Next i
'        If (ErrCde = 0) Then
'            Call GetFilterTemperature(1, fTemp, st)
'            fTemp(0) = st(0)
'        End If
'
'
'    End If
    ReadTemperatureAllChannelFT = ErrCde
    
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Read Advan TC Filter error"
    ShowAlarmFlash 1
End Function

Public Sub GetFilterTemperature(TC_Num As Integer, InputTemp() As Single, OutputTemp() As Single)
    Dim i As Integer
    Dim j As Integer
    Dim sngInputValue As Single
    
    On Error GoTo ERRLINE
    For i = 0 To TC_Num
        sngInputValue = (InputTemp(i) + TC_CVT2) * TC_CVT1
        TC_Filter(i, TC_FilterCount) = sngInputValue
    
        If TC_Filter(i, TC_FilterCountMax - 1) = 0 Then
            OutputTemp(i) = sngInputValue
        Else
            TC_Filter(i, TC_FilterCountMax) = 0
            For j = 0 To TC_FilterCountMax - 1
                TC_Filter(i, TC_FilterCountMax) = TC_Filter(i, TC_FilterCountMax) + TC_Filter(i, j)
            Next j
    
            TC_Filter(i, TC_FilterCountMax) = TC_Filter(i, TC_FilterCountMax) / TC_FilterCountMax
            OutputTemp(i) = TC_Filter(i, TC_FilterCountMax)
            For j = 0 To TC_FilterCountMax - 1
                TC_Filter(i, j) = TC_Filter(i, j + 1)
            Next j
        End If
    Next i
    If (TC_FilterCount < TC_FilterCountMax - 1) Then
        TC_FilterCount = TC_FilterCount + 1
    End If
    Exit Sub
ERRLINE:
    MsgBox "Error In Filter Temperature!!"
End Sub
Public Function ReadAI(fTemp() As Single) As Long
    Dim i As Integer
    Dim idx As Integer
    
    
    idx = 0
    
    
    For i = 4 To 8 Step 2
    
        lpAIConfig.DasChan = i
        lpAIConfig.DasGain = 4 '+- 10 V
        ErrCde = DRV_AIConfig(DeviceHandle, lpAIConfig)
        
        ptAIRead.chan = i
        ptAIRead.gain = 4   '+- 10 V
        ptAIRead.TrigMode = 0
        ptAIRead.voltage = DRV_GetAddress(fTemp(idx))

        ErrCde = DRV_AIVoltageIn(DeviceHandle, ptAIRead)

        idx = idx + 1
'        If (ErrCde <> 0) Then Exit For
    Next i
    ReadAI = ErrCde
    
    
    
End Function

Public Function ReadDI() As Long
    Dim value As Integer
    Dim i As Integer
    Dim Shift As Integer
    Dim j As Integer

    Shift = 1
    lpDioReadPort.Port = 0
    lpDioReadPort.value = DRV_GetAddress(value)
    ErrCde = DRV_DioReadPortByte(DeviceHandle, lpDioReadPort)
    For i = 0 To 7
        
        If (Shift And value) = 2 ^ i Then
            SysDI.value(CLng(i)) = 0
            AvDaq17101.lngDI(CLng(i)) = 0
        Else
            SysDI.value(CLng(i)) = 1
            AvDaq17101.lngDI(CLng(i)) = 1
        End If
        Shift = Shift * 2
    Next i
    lpDioReadPort.Port = 1
    lpDioReadPort.value = DRV_GetAddress(value)
    ErrCde = DRV_DioReadPortByte(DeviceHandle, lpDioReadPort)
    Shift = 1
    For i = 0 To 7
        If (Shift And value) = 2 ^ i Then
            
            SysDI.value(CLng(i + 8)) = 0
            AvDaq17101.lngDI(CLng(i + 8)) = 0
        Else
            SysDI.value(CLng(i + 8)) = 1
            AvDaq17101.lngDI(CLng(i + 8)) = 1
        End If
        Shift = Shift * 2
    Next i
    

End Function

Public Function WriteDO(Port As Long, bol As Boolean) As Long
    
    Dim value As Byte
    Dim i As Integer

    
     
    
    
    
    value = 0
    If Port < 8 Then
        For i = 0 To 7
            'value = value + (2 ^ i) * AvDaq17101.lngDO(i)
            value = value + (2 ^ i) * SysDO.value(i)
        Next i
        
        
        
         
        
        lpDioWritePort.Port = 0
        lpDioWritePort.Mask = 255
        lpDioWritePort.state = value
        ErrCde = DRV_DioWritePortByte(DeviceHandle, lpDioWritePort)
        WriteDO = ErrCde

    Else
        For i = 0 To 7
            'value = value + (2 ^ i) * AvDaq17101.lngDO(i + 8)
            value = value + (2 ^ i) * SysDO.value(i + 8)
        Next i
         
        
        lpDioWritePort.Port = 1
        lpDioWritePort.Mask = 255
        lpDioWritePort.state = value
        ErrCde = DRV_DioWritePortByte(DeviceHandle, lpDioWritePort)
        WriteDO = ErrCde
    End If
    
    

    
End Function

Public Function WriteAO(Port As Integer, value As Single) As Long
    If Port < 2 Then
        ptAOWrite.chan = Port
        ptAOWrite.OutputValue = value
        ErrCde = DRV_AOVoltageOut(DeviceHandle, ptAOWrite)
        WriteAO = ErrCde
    Else
        WriteAO = -1
    End If
End Function


