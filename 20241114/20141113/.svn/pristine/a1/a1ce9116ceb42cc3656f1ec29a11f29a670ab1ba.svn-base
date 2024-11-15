Attribute VB_Name = "mdlUSB4718"
Global Const MaxEntries = 255
Global Const MaxAIChl = 8
Global DeviceHandle As Long
Global ptDevGetFeatures As PT_DeviceGetFeatures
Global lpDevFeatures As DEVFEATURES
Global devicelist(0 To MaxEntries) As PT_DEVLIST
Global SubDevicelist(0 To MaxEntries) As PT_DEVLIST
Global ErrCde As Long
Global szErrMsg As String * 80
Global bRun As Boolean

Global ptAIGetConfig As PT_AIGetConfig
Global ptTCMuxRead As PT_TCMuxRead
Global BurnOutSupport As Boolean

Global lpAIGetConfig As DEVCONFIG_AI

Global BurnOutProperty(0 To MaxAIChl - 1) As Long


Public Function InitialUSB4718() As Long
    Dim gnNumOfDevices As Integer
    Dim nOutEntries As Integer
    Dim i, j As Integer
    Dim lpDeviceList As Long
    Dim ii As Integer
    Dim tempNum As Integer
    Dim TestRes As Boolean
    Dim lpSubDeviceList As Long
    Dim dwDeviceNum As Long
    Dim Length As Long
    Dim Response
    Dim BoardID As Integer
    
    ' initialize the device open/close flag
    bRun = False
    
    ErrCde = DRV_DeviceGetNumOfList(gnNumOfDevices)
    If (ErrCde <> 0) Then
      DRV_GetErrorMessage ErrCde, szErrMsg
      Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
      InitialUSB4718 = Response
      Exit Function
    End If
    
    If (gnNumOfDevices > MaxDev) Then
      gnNumOfDevices = MaxDev
    End If
    
    ' retrieve the information of all installed devices
    lpDeviceList = DRV_GetAddress(devicelist(0))
    ErrCde = DRV_DeviceGetList(lpDeviceList, MaxEntries, nOutEntries)
    If (ErrCde <> 0) Then
      DRV_GetErrorMessage ErrCde, szErrMsg
      Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
      InitialUSB4718 = Response
      Exit Function
    End If
    
    If (gnNumOfDevices > MaxDev) Then
      gnNumOfDevices = MaxDev
    End If
    
    ' Check if there is any device attatched on this COM port or CAN
    gnNumOfSubdevices = devicelist(0).nNumOfSubdevices
    If (gnNumOfSubdevices > MaxDev) Then
        gnNumOfSubdevices = MaxDev
    End If
    
    ' retrieve the information of all installed devices
    If (gnNumOfSubdevices <> 0) Then
        dwDeviceNum = devicelist(0).dwDeviceNum
        lpSubDeviceList = DRV_GetAddress(SubDevicelist(0))
        ErrCde = DRV_DeviceGetSubList(dwDeviceNum, lpSubDeviceList, gnNumOfSubdevices, nOutEntries)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            InitialUSB4718 = Response
            Exit Function
        End If
    End If
    
    ' Open PCL card device
    If (gnNumOfSubdevices = 0) Then
        dwDeviceNum = devicelist(0).dwDeviceNum
        ErrCde = DRV_DeviceOpen(dwDeviceNum, DeviceHandle)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            InitialUSB4718 = Response
            Exit Function
        Else
            bRun = True
        End If
    
        ptDevGetFeatures.buffer = DRV_GetAddress(lpDevFeatures)
        ErrCde = DRV_DeviceGetFeatures(DeviceHandle, ptDevGetFeatures)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            InitialUSB4718 = Response
            Exit Function
        End If
    
        ptAIGetConfig.buffer = DRV_GetAddress(lpAIGetConfig)
        ErrCde = DRV_AIGetConfig(DeviceHandle, ptAIGetConfig)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            InitialUSB4718 = Response
            Exit Function
        End If
    
    
          ' Add thermo channel item
          If (lpDevFeatures.usMaxAISiglChl > lpDevFeatures.usMaxAIDiffChl) Then
             tempNum = lpDevFeatures.usMaxAISiglChl
          Else
             tempNum = lpDevFeatures.usMaxAIDiffChl
          End If
    
          ' Support Burn Test?
    
          ' Please refer the software manual for the return value when burn out occurs.
          ' Here is for USB-4718, 0: Disable, 1: 888888, 2: -888888, 3: Maximum value, 4: Minimum value
    '        cmbBTEnable.Clear
    '        cmbBTEnable.AddItem ("Disable")
    '        cmbBTEnable.AddItem ("888888")
    '        cmbBTEnable.AddItem ("-888888")
    '        cmbBTEnable.AddItem ("Maximum value")
    '        cmbBTEnable.AddItem ("Minimum value")
    
          Length = 4 * MaxAIChl
          ErrCde = DRV_DeviceGetProperty(DeviceHandle, CFG_BURNTEST, BurnOutProperty(0), Length)
          If ((ErrCde And &HFFF) = InvalidInputParam) Then
             ErrCde = DRV_DeviceGetProperty(DeviceHandle, CFG_BURNTEST, BurnOutProperty(0), Length)
          End If
    
          If ((ErrCde And &HFFF) = ParamNameNotSupported) Then
              BurnOutSupport = False
          ElseIf ErrCde = 0 Then
              BurnOutSupport = True
    
          Else
              DRV_GetErrorMessage ErrCde, szErrMsg
              Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
              InitialUSB4718 = Response
                Exit Function
          End If
      End If
    
    ptTCMuxRead.DasChan = 0
    
     If (BurnOutSupport = True) Then
        BurnOutProperty(ptTCMuxRead.DasChan) = 0
        Length = 4 * MaxAIChl
        ErrCde = DRV_DeviceSetProperty(DeviceHandle, CFG_BURNTEST, BurnOutProperty(0), Length)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            InitialUSB4718 = Response
            Exit Function
        End If
     End If
    ' Gain code no use for ADAM series
    BoardID = lpAIGetConfig.Daughter(0).dwBoardID
    If gnNumOfSubdevices = 0 Then
         If BoardID > 0 And BoardID <> BD_PCLD8710 Then
            ptTCMuxRead.DasGain = 0
            ptTCMuxRead.ExpChan = 0
         Else
            tempNum = 0
            ptTCMuxRead.DasGain = lpDevFeatures.glGainList(tempNum).usGainCde
            ptTCMuxRead.ExpChan = 0
         End If
    End If
    
    ptTCMuxRead.TCType = 1

    InitialUSB4718 = 0
End Function

