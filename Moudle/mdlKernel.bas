Attribute VB_Name = "mdlKernel"
Public Para As ParaType
Public SysDI As DIType
Public SysDO As DOType
Public SysAI As AIType
Public SysAO As AOType
Public Kernel As KernelType
Public MultiLoop As MultiLoopType
Public Az1 As AzbilType
Public Az2 As AzbilType

Public CurrProc As ProcType
Public iniPara As New cInifile
Public strServerFileName As String
Public IsDebugMode As Boolean
Public mUSBIO_1 As ICPDAS_USBIO
Public mUSBIO_2 As ICPDAS_USBIO
Public Write_Result As Boolean
Public TCM1_WriteErrDetail As String
Public TCM2_WriteErrDetail As String
Public WriteErr_Count As Integer
Public IsInitAzbil As Boolean
Public Az1_ConNum As Integer
Public Az2_ConNum As Integer
Public Az1_ConRes As Boolean
Public Az2_ConRes As Boolean




Option Explicit


Public Function ShowAlarm(sMessage As String) As Long
    frmShowAlarm.lblMessage.Caption = sMessage
    frmShowAlarm.Show
End Function

Public Function ShowMessageOK(sMessage As String) As Long
    If gbblnNoModalForm = False Then
        frmShowAlarm.lblMessage.Caption = sMessage
        frmShowAlarm.Show
    End If
End Function

Public Function ShowMessageYN(sMessage As String) As Long
    gbblnNoModalForm = True
    frmStartProcess.lblMessage.Caption = sMessage
    frmStartProcess.Show vbModal
    gbblnNoModalForm = False
End Function
'
'Public Function ShowAlarmFlash(iId As Integer) As Long
'    Dim S As String
'
'    If Para.AlarmActive(iId) = 0 Then
'        If Para.AlarmDo(iId) <> 0 Then
'            frmPlotProcess.lbAlarmID.Caption = CStr(4000 + iId)
'            S = Para.AlarmName(iId) & gbstrAlarmHint
'            frmPlotProcess.lbAlarmName.Caption = S
'            frmPlotProcess.lbAlarmTime.Caption = CStr(Date) & " " & Format(Time, "hh:mm:ss")
'        End If
'        Select Case Para.AlarmDo(iId)
'
'            Case 1  'Only show
'                frmPlotProcess.tmrAlarmFlash.Enabled = True
'                frmPlotProcess.lbAlarmDo.Caption = "請通知設備管理員"
'                frmHistory.AppendLogAlert 1, "Alarm", 4000 + iId, S, 1
'                Kernel.IsAlarm = iId
'            Case 2  'Only show+Beep
'                frmPlotProcess.tmrAlarmFlash.Enabled = True
'                frmPlotProcess.lbAlarmDo.Caption = "請通知設備管理員"
'                frmHistory.AppendLogAlert 1, "Alarm", 4000 + iId, S, 1
'                Kernel.IsAlarm = iId
'                SetTower 1, True
'            Case 4  'Show + Beep + Stop
'                frmPlotProcess.tmrAlarmFlash.Enabled = True
'                frmPlotProcess.lbAlarmDo.Caption = "立即停止!請通知設備管理員"
'                frmHistory.AppendLogAlert 1, "Alarm", 4000 + iId, S, 1
'                Kernel.IsAlarm = iId
'                SetTower 1, True
'                StopProc 1
'        End Select
'
'        Para.AlarmActive(iId) = 1
'
'    End If
'End Function

Public Function ShowAlarmFlash(iId As Integer) As Long
    Dim S As String
    Dim s1 As String
    
    If Para.AlarmActive(iId) = 0 Then
        If Para.AlarmDo(iId) <> 0 Then
            frmPlotProcess.lbAlarmID.Caption = CStr(4000 + iId)
            S = Para.AlarmName(iId) & gbstrAlarmHint
            If iId = 23 Then
'                    S = Para.AlarmName(iId) & gbstrAlarmHint & Kernel.allZerosColumns
                s1 = Para.AlarmName(iId) & Kernel.allZerosColumns
                frmPlotProcess.lbAlarmName.Caption = s1
            ElseIf iId = 29 Or iId = 30 Then
            frmPlotProcess.lbAlarmName.Caption = Para.AlarmName(iId) + "---" + mdifrmRTP.AlarmDetail
            
            Else
                frmPlotProcess.lbAlarmName.Caption = S
            End If
'            frmPlotProcess.lbAlarmName.Caption = S
            frmPlotProcess.lbAlarmTime.Caption = CStr(Date) & " " & Format(Time, "hh:mm:ss")
        End If
        Select Case Para.AlarmDo(iId)
            
            Case 1  'Only show
                frmPlotProcess.tmrAlarmFlash.Enabled = True
                frmPlotProcess.lbAlarmDo.Caption = "請通知設備管理員"
                frmHistory.AppendLogAlert 1, "Alarm", 4000 + iId, S, 1
                Kernel.IsAlarm = iId
            Case 2  'Only show+Beep
                frmPlotProcess.tmrAlarmFlash.Enabled = True
                frmPlotProcess.lbAlarmDo.Caption = "請通知設備管理員"
                frmHistory.AppendLogAlert 1, "Alarm", 4000 + iId, S, 1
                Kernel.IsAlarm = iId
                SetTower 1, True
            Case 4  'Show + Beep + Stop
                frmPlotProcess.tmrAlarmFlash.Enabled = True
                frmPlotProcess.lbAlarmDo.Caption = "立即停止!請通知設備管理員"
                frmHistory.AppendLogAlert 1, "Alarm", 4000 + iId, S, 1
                Kernel.IsAlarm = iId
                SetTower 1, True
                StopProc 1
        End Select
        
        Para.AlarmActive(iId) = 1
                
    End If
End Function

Public Function AlarmFlashClose() As Long
    Dim i As Integer
    
    Kernel.IsAlarm = 0
    frmPlotProcess.fraProcessChart.BackColor = IIf(Kernel.IsRun = 0, &H8000000F, &HFF00&)
    frmPlotProcess.tmrAlarmFlash.Enabled = False
    frmPlotProcess.fraAlarm.Visible = False
    frmHistory.AppendLogAlert 1, "Manual", 1030, "警報重置", 1
    gbstrAlarmHint = ""
    If Kernel.IsRun = 0 Then
        SetTower 2, True
    Else
        SetTower 3, True
    End If
    For i = 0 To 30
        Para.AlarmActive(i) = 0
    Next i
    For i = 0 To 35
        Kernel.intOverCT(i) = 0
        Kernel.intUnderCT(i) = 0
        frmPlotProcess.lbCT(i).BackColor = &H8000000F
    Next i
    SetDO gblngDO_ARM_FRONT, False
    StopProc 2
    
    
    If Para.UseCIM = 1 Then
        frmCIM.Send "$SPR=7,"
    End If
End Function

Public Function InitHG() As Boolean
    Dim iRet As Integer
    Dim i As Integer
    Dim wBoardIndex As Integer
    
On Error GoTo INITERR
    
    gbintTotalBoard = 0
    iRet = Ixud_DriverInit(gbintTotalBoard)
    If (iRet) Then
        gbstrAlarmHint = " Initial HG error"
        ShowAlarmFlash 1
        InitHG = False
        Exit Function
    End If
    
    gbintPCI822A = 0
    gbintPIODA8A = 1
    For wBoardIndex = 0 To gbintTotalBoard - 1
        iRet = Ixud_GetCardInfo(wBoardIndex, gbstDevInfo(wBoardIndex), gbstCardInfo(wBoardIndex), gbstrModelName)
        If (iRet) Then
            gbstrAlarmHint = " Initial HG error"
            ShowAlarmFlash 1
            InitHG = False
            Exit Function
        End If
        iRet = InStr(1, gbstrModelName, "PCI-822")
        If iRet >= 1 Then
            gbintPCI822A = wBoardIndex
            gbblnActiveHGD = True
        End If
        iRet = InStr(1, gbstrModelName, "PIO-DA4/8/16")
        If iRet >= 1 Then
            gbintPIODA8A = wBoardIndex
            gbblnActiveHGA = True
        End If
    Next
    
    For i = 0 To 15
        iRet = Ixud_ConfigAO(gbintPIODA8A, i, 3) '0~10V
    Next i
    
    InitHG = True
    Exit Function
INITERR:
    InitHG = False
    gbblnActiveHGD = False
    gbblnActiveHGA = False
End Function

Public Function InitHGU() As Boolean
    Dim iRet As Integer
    Dim i As Integer
    Dim wBoardIndex As Integer
    
On Error GoTo INITERR
    
    Set mUSBIO_1 = New ICPDAS_USBIO
        
    'mUSBIO.GetDeviceID
    iRet = mUSBIO_1.OpenDevice(&H413, 1)
    If iRet = ERR_NO_ERR Then
        Kernel.IsActiveTC1 = 1
        If Para.UseMTC = 1 Then
            Set mUSBIO_2 = New ICPDAS_USBIO
            iRet = mUSBIO_2.OpenDevice(&H413, 2)
            If iRet = ERR_NO_ERR Then
                Kernel.IsActiveTC2 = 1
            Else
                GoTo INITERR
            End If
        End If
    Else
        GoTo INITERR
    End If
    
    InitHGU = True
    gbblnActiveHGU = True
    Exit Function
INITERR:
    
    InitHGU = False
    gbblnActiveHGU = False
    gbbnlPC_STOP = True
    gbstrAlarmHint = " Initial HGU error=" & CStr(iRet)
    ShowAlarmFlash 1
End Function

Public Function InitHGA() As Boolean
    Dim iRet As Integer
    Dim i As Integer
    Dim wBoardIndex As Integer
' On Error GoTo ErrorHandler
On Error GoTo INITERR
    
    Set mUSBIO_1 = New ICPDAS_USBIO
        
    iRet = mUSBIO_1.OpenDevice(&H413, 1)
    If iRet = ERR_NO_ERR Then
        Kernel.IsActiveTC1 = 1
    Else
        GoTo INITERR
    End If
    
    InitHGA = True
    gbblnActiveHGU = True
    Exit Function
'ErrorHandler:
'    Dim errorMessage As String
'    errorMessage = "在函數MyFunction中發生錯誤,錯誤代碼: " & Err.Number & ", 錯誤描述: " & Err.Description & ", 錯誤行號: " & Erl & vbNewLine
'    WriteLog (errorMessage)
INITERR:
    
    InitHGA = False
    gbblnActiveHGU = False
    gbbnlPC_STOP = True
    gbstrAlarmHint = " Initial HGA error=" & CStr(iRet)
    ShowAlarmFlash 1
End Function

Public Function InitHGB() As Boolean
    Dim iRet As Integer
    Dim i As Integer
    Dim wBoardIndex As Integer
    
On Error GoTo INITERR
    
    Set mUSBIO_2 = New ICPDAS_USBIO
        
    iRet = mUSBIO_2.OpenDevice(&H413, 2)
    If iRet = ERR_NO_ERR Then
        Kernel.IsActiveTC2 = 1
       
    Else
        GoTo INITERR
    End If
    
    InitHGB = True
    gbblnActiveHGU = True
    Exit Function
INITERR:
    
    InitHGB = False
    gbblnActiveHGU = False
    gbbnlPC_STOP = True
    gbstrAlarmHint = " Initial HGB error=" & CStr(iRet)
    ShowAlarmFlash 1
End Function


Public Function InitFR() As Boolean
    Dim iRet As Integer
    
    
    
On Error GoTo INITERR
    iRet = FRB_DriverInit(0)
    If iRet <> FRB_NoError Then
        InitFR = False
        Exit Function
    End If
    
    
    
    Kernel.IsActiveIO = 1
    gbblnActiveHGF = True
    InitFR = True
    
    Exit Function
INITERR:
    
    InitFR = False
    gbblnActiveHGF = False
    gbbnlPC_STOP = True
    gbstrAlarmHint = " Initial FR error"
    ShowAlarmFlash 1
End Function

Public Function InitPM() As Boolean
    Dim i As Integer
On Error GoTo ERRCOM:
        
    Net_ID = 1
    frmConfiguration.MSComm1.InputMode = comInputModeBinary
    InitCRCTable
'    i = 0                           '==> Modbus function(0x03): read Holding Registers
'        CmdData(i, 0) = Net_ID      'device ID
'        CmdData(i, 1) = &H3         'function call
'        CmdData(i, 2) = &H10        'Start Address (Hi byte)
'        CmdData(i, 3) = &H0         'Start Address (Lo byte)
'        CmdData(i, 4) = &H0         'No. of registers (Hi byte)
'        CmdData(i, 5) = &H5         'No. of registers (Lo byte)
'        CmdData(i, 6) = &H81        'CRC Check (Lo Byte)
'        CmdData(i, 7) = &H9         'CRC Check (Hi Byte)
'        frmConfiguration.MSComm1.RThreshold = 15

    i = 1                           '==> Modbus function(0x04): read Input Registers
        CmdData(i, 0) = Net_ID      'device ID
        CmdData(i, 1) = &H4         'function call
        CmdData(i, 2) = &H11        'Start Address (Hi byte)
        CmdData(i, 3) = &H0         'Start Address (Lo byte)
        CmdData(i, 4) = &H0         'No. of registers (Hi byte)
        CmdData(i, 5) = &H48        'No. of registers (Lo byte)
        CmdData(i, 6) = &HF5        'CRC Check (Lo Byte)
        CmdData(i, 7) = &H0         'CRC Check (Hi Byte)
            
    Kernel.IsPM = 0
    gbintPMcmdID = 0
    gbblnReceivedPM = False
    frmConfiguration.MSComm1.RThreshold = 149
    frmConfiguration.MSComm1.CommPort = Para.intComCT
    frmConfiguration.MSComm1.PortOpen = True
    frmConfiguration.tmrSendCT.Enabled = True
    
    'Call SendCmdPM(0)
    'frmConfiguration.MSComm1.RThreshold = 15
    'frmConfiguration.MSComm1.Output = gbbtOutByte
    Kernel.IsActivePM = 1
    InitPM = True
    Exit Function
ERRCOM:
    ShowAlarmFlash 20
End Function

Public Function InitSys() As Boolean
    Set advThermo = New clsAdvThermo
    If Para.RtaType = 1 Or Para.RtaType = 2 Then
        advThermo.InitialCard
        InitHG
        frmConfiguration.tmrAIO.Enabled = advThermo.IsActive
        frmConfiguration.tmrDIO.Enabled = advThermo.IsActive
        
    End If
    If Para.RtaType = 3 Then
        InitHG
        InitHGU
        frmConfiguration.tmrAIO.Enabled = gbblnActiveHGD
        frmConfiguration.tmrDIO.Enabled = gbblnActiveHGD
    End If
    If Para.UseAutoMode = 1 Then frmUDP.OpenUDP
    If Para.RtaType = 5 Then
        If InitFR = True Then
            InitHGU
            frmConfiguration.tmrAIO.Enabled = gbblnActiveHGF
            frmConfiguration.tmrDIO.Enabled = gbblnActiveHGF
        End If
    End If
    If Para.RtaType = 6 Then
        If InitFR = True Then
            advThermo.InitialCard
            frmConfiguration.tmrAIO.Enabled = gbblnActiveHGF
            frmConfiguration.tmrDIO.Enabled = gbblnActiveHGF
        End If
    End If
    If Para.RtaType = 7 Then
        frmTCP.OpenTCP
        InitHG
        frmConfiguration.tmrAIO.Enabled = gbblnActiveHGD
        frmConfiguration.tmrDIO.Enabled = gbblnActiveHGD
    End If
    If Para.RtaType = 8 Then
        If InitFR = True Then
            frmTCP.OpenTCP
            frmConfiguration.tmrAIO.Enabled = gbblnActiveHGF
            frmConfiguration.tmrDIO.Enabled = gbblnActiveHGF
        End If
        frmTCP.OpenTCP
        frmConfiguration.tmrTest.Enabled = True
    End If
    
    If Para.RtaType = 9 Then
         If InitFR = True Then
            If Para.UseMTC = 1 Then InitHGA
            If Para.UseMTCB = 1 Then InitHGB
            frmConfiguration.tmrAIO.Enabled = gbblnActiveHGF
            frmConfiguration.tmrDIO.Enabled = gbblnActiveHGF
        End If
        If Para.UseAz1 = 1 Then
'            Call frmAz1.OpenTCP(Para.strAzIP1)
            For Az1_ConNum = 0 To 2
            Az1_ConRes = frmAz1.OpenTCP(Para.strAzIP1)
            If Az1_ConRes = True Then Exit For
            Next Az1_ConNum
        End If
        If Para.UseAz2 = 1 Then
           For Az2_ConNum = 0 To 2
            Az2_ConRes = frmAz2.OpenTCP(Para.strAzIP2)
            If Az2_ConRes = True Then Exit For
            Next Az2_ConNum
'          Call frmAz2.OpenTCP(Para.strAzIP2)
        End If
    End If
    
    If Para.UseCT = 1 Then Call InitPM
    
    If Para.RtaType = 9 And IsUsedSCR = 1 Then frmModBusRtu.InitModBusRtu ("SCR")
    
    If Para.UseCIM = 1 Then
        frmDCR.OpenUDP
        frmCIM.OpenUDP
        If Para.intCIMPort = 1 Then
            frmCIM.Send "$GRS=1,"
            
        End If
    End If
              
    Kernel.intCurrCycleRun = 0
    Kernel.intCurrMonitorRun = 0
    ResetDO
    ResetAO
'    SetTower 2, True

    If Para.UseCover = 1 Then
        InitCover
    End If
    
        
    IsDebugMode = True

End Function

Public Function InitAzbil() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim iMaxStep As Integer
    Dim b1, b2 As Boolean
    Dim ProcData(9) As Integer
    Dim ParaData(0 To 19) As Integer
    Dim iTemp, iTime As Integer
    Dim iStep As Integer
    Dim iAdd As Long
    Dim Data1(3) As Integer
    Dim Data2(3) As Integer
    Dim SysProName As String
    Dim HoldTimes As Integer
    Dim WriteErr_Count As Integer
    
    

    On Error GoTo ERRLINE:
    
    IsInitAzbil = True
    
    If Para.UseAz1 Or Para.UseAz2 Then
        iMaxStep = 32
        
               
        HoldTimes = 0
        iAdd = 48100
        For i = 1 To GB_MAX_STEP_PROCESS
            SysProName = Readini(gbProcessRecipeStep(i).strAction)
            If SysProName <> "" Then
            gbProcessRecipeStep(i).strAction = SysProName
            End If
            If gbProcessRecipeStep(i).strAction = GB_ACTION_IDLE Then
                iTemp = 0
                iTime = gbProcessRecipeStep(i).sngTime
            ElseIf gbProcessRecipeStep(i).strAction = GB_ACTION_RAMPUP Or gbProcessRecipeStep(i).strAction = GB_ACTION_HOLD Then
                iTemp = gbProcessRecipeStep(i).sngTemperature
                iTime = gbProcessRecipeStep(i).sngTime
                If gbProcessRecipeStep(i).strAction = GB_ACTION_HOLD Then
                 HoldTimes = HoldTimes + 1
                If HoldTimes = Val(CommnonReadini("Special_Setting", "Hold_Times", App.Path + ProcDict_Path)) Then
                     iTime = iTime + Val(CommnonReadini("Special_Setting", "Hold_Offset", App.Path + ProcDict_Path))
                End If
                End If
            ElseIf gbProcessRecipeStep(i).strAction = GB_ACTION_STOP Then
                Exit For
            End If
            If gbProcessRecipeStep(i).strAction <> GB_ACTION_IOCONTROL Then
                ProcData(0) = iTemp
                ProcData(1) = iTime
                ProcData(2) = 1
                ProcData(3) = 0
                ProcData(4) = 1
                If Az1.blnUseAzbil = True Then
                For WriteErr_Count = 0 To 2
                Write_Result = frmAz1.WriteParas(iAdd, ProcData, True)
                If Write_Result = True Then Exit For
                Next WriteErr_Count
                If Write_Result = False Then TCM1_WriteErrDetail = "參數1"
                
                End If
                If Az2.blnUseAzbil = True Then
                For WriteErr_Count = 0 To 2
                Write_Result = frmAz2.WriteParas(iAdd, ProcData, True)
                If Write_Result = True Then Exit For
                Next WriteErr_Count
                If Write_Result = False Then TCM2_WriteErrDetail = "參數1"
'                Call frmAz2.WriteParas(iAdd, ProcData, True)
                End If
                iAdd = iAdd + 10
            End If
        Next i
        For i = i To iMaxStep
            
            ProcData(0) = 0
            ProcData(1) = 0
            ProcData(2) = 1
            ProcData(3) = 0
            ProcData(4) = 1
            If Az1.blnUseAzbil = True Then
'            Call frmAz1.WriteParas(iAdd, ProcData, True)
                For WriteErr_Count = 0 To 2
                Write_Result = frmAz1.WriteParas(iAdd, ProcData, True)
                If Write_Result = True Then Exit For
                Next WriteErr_Count
                If Write_Result = False Then TCM1_WriteErrDetail = "參數2"
            End If
            If Az2.blnUseAzbil = True Then
'            Call frmAz2.WriteParas(iAdd, ProcData, True)
                For WriteErr_Count = 0 To 2
                Write_Result = frmAz2.WriteParas(iAdd, ProcData, True)
                If Write_Result = True Then Exit For
                Next WriteErr_Count
                If Write_Result = False Then TCM2_WriteErrDetail = "參數2"
            End If
            iAdd = iAdd + 10
        Next i



        'Last Step must be 0
'        ProcData(0) = 0
'        ProcData(1) = 0
'        If Az1.blnUseAzbil = True Then Call frmAz1.WriteParas(iAdd, ProcData, True)
'        If Az2.blnUseAzbil = True Then Call frmAz2.WriteParas(iAdd, ProcData, True)

        
        If Az1.blnUseAzbil = True Then
            For i = 0 To 4
                For j = 0 To 3
                    If i = 0 Then ParaData(i * 4 + j) = Az1.sngPN(j) * 100
                    If i = 1 Then ParaData(i * 4 + j) = Az1.sngIN(j) * 100
                    If i = 2 Then ParaData(i * 4 + j) = Az1.sngDN(j) * 100
                    If i = 3 Then ParaData(i * 4 + j) = Az1.sngRT(j) * 10000
                    If i = 4 Then ParaData(i * 4 + j) = Az1.sngST(j) * 10
                Next j
                
            Next i
'            Call frmAz1.WriteParas(201, ParaData(), True)
                For WriteErr_Count = 0 To 2
                Write_Result = frmAz1.WriteParas(201, ParaData(), True)
                If Write_Result = True Then Exit For
                Next WriteErr_Count
                If Write_Result = False Then TCM1_WriteErrDetail = "參數3"
        End If
        If Az2.blnUseAzbil = True Then
            For i = 0 To 4
                For j = 0 To 3
                    If i = 0 Then ParaData(i * 4 + j) = Az2.sngPN(j) * 100
                    If i = 1 Then ParaData(i * 4 + j) = Az2.sngIN(j) * 100
                    If i = 2 Then ParaData(i * 4 + j) = Az2.sngDN(j) * 100
                    If i = 3 Then ParaData(i * 4 + j) = Az2.sngRT(j) * 10000
                    If i = 4 Then ParaData(i * 4 + j) = Az2.sngST(j) * 10
                Next j
            Next i
'            Call frmAz2.WriteParas(201, ParaData(), True)
                For WriteErr_Count = 0 To 2
                Write_Result = frmAz2.WriteParas(201, ParaData(), True)
                If Write_Result = True Then Exit For
                Next WriteErr_Count
                If Write_Result = False Then TCM2_WriteErrDetail = "參數3"
        End If
        
        If Az1.blnUseAzbil Then
            gbintAz1ProcNo = 1
        End If
        If Az2.blnUseAzbil Then
            gbintAz2ProcNo = 1
        End If
    End If
    
    Exit Function
ERRLINE:
    gbstrAlarmHint = " StartProc(Server) error"
    ShowAlarmFlash 17
End Function




Public Function InitCover() As Boolean
    Dim Timeout As Long
    Dim tmrStamp As Single
    Dim tmrStampNow As Single
    
    SetDO gblngDO_COVER_ARESET, True
    SetDO gblngDO_COVER_SERVO, True
    Sleep 500
    SetDO gblngDO_COVER_ORGIN, True
    gbintCoverOrigCount = 0
    frmDiagnosis.tmrCoverOrig.Enabled = True

    Exit Function
ERRLINE:
    gbstrAlarmHint = " Cover Servo error"
    ShowAlarmFlash 28
End Function


Private Sub HideTcmModule()
On Error GoTo LockTcmModule:
frmRecipeEdit.cmdWriteAz1.Visible = False
frmRecipeEdit.cmdWriteOther.Visible = False
frmRecipeEdit.cmdReadAz1.Visible = False
Exit Sub
LockTcmModule:
WriteLog ("隱藏TCM模塊按鈕失敗!!!!")
End Sub

Private Sub ShowTcmModule()
On Error GoTo ShowTcmModule:
frmRecipeEdit.cmdWriteAz1.Visible = True
frmRecipeEdit.cmdWriteOther.Visible = True
frmRecipeEdit.cmdReadAz1.Visible = True
Exit Sub
ShowTcmModule:
WriteLog ("顯示TCM模塊按鈕失敗!!!!")
End Sub

Public Function StartProc() As Boolean
    Dim strTemp As String
    On Error GoTo ERRLINE:
    
    frmPlotProcess.InitPlotChart
    strTemp = "Start " & frmRecipeEdit.lbRecipeName.Caption
    frmHistory.AppendLogAlert 1, "Process", 1000, strTemp, 1
    Kernel.strCurrReportTime = Format(Time, "hh:mm:ss")
    SetTower 3, True
    SetLampCooling True
    
    Kernel.IsRun = 1
    Call HideTcmModule
    If Para.UseBarcodeServer = 1 Then
        strTemp = Kernel.strServerPath & Kernel.strBarcodeID & "_" & Kernel.strServerRecipe & "_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_Start.txt"
        Open strTemp For Random As #1
        Close #1
    End If
    
    If gbblnRecipeStartAutoClose = True Then
        If SetDoor(0) = False Then
            gbstrAlarmHint = " StartProc(Door Close) error"
            ShowAlarmFlash 11
            Exit Function
        Else
            Call frmHistory.AppendLogAlert(1, "Process", 1016, "自動關門", 1)
        End If
    End If
    
    If Para.UseCover = 1 Then
        If gbblnRecipeStartCloseCover = True Then
            If SetCover(True) = True Then
                Call frmHistory.AppendLogAlert(1, "Process", 1017, "自動蓋板", 1)
            End If
        End If
    End If
        
    If Para.RtaType = 9 Then
        InitAzbil
    End If
'    If Para.RtaType = 9 And IsUsedSCR = 1 Then
'        frmModBusRtu.WriteRamupSCR
'    End If
    If MultiLoop.blnUseMultiLoop = True And Para.RtaType = 7 Or Para.RtaType = 8 Or Para.RtaType = 9 Then
        MultiLoop.blnUseMultiLoop = False
    End If
    
'    If Kernel.IsAlarm = 0 And Kernel.IsRun = 1 Then InitProcessStep
     If Kernel.IsRun = 1 Then InitProcessStep
     
     While Kernel.isEnd = False
        DoEvents
    Wend
    
    Exit Function
ERRLINE:
    gbstrAlarmHint = " StartProc(Server) error"
    ShowAlarmFlash 17
End Function

Public Function StopProc(iStopMode As Integer) As Boolean
    Dim strTemp As String
    Dim sngGas(GB_GAS_MAX) As Single
    Dim i As Integer
      Dim j As Integer
    Dim data(4) As Integer
    
    On Error GoTo ERRLINE:
    
'    If Para.UseBarcodeServer = 1 And Kernel.IsRun = 1 Then
'        If iStopMode = 0 Then
'            strServerFileName = Kernel.strServerPath & Kernel.strBarcodeID & "_" & Kernel.strServerRecipe & "_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_End.txt"
'        Else
'            strServerFileName = Kernel.strServerPath & Kernel.strBarcodeID & "_" & Kernel.strServerRecipe & "_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_Stop.txt"
'        End If
'        FileCopy Kernel.strCurrLogFile, strServerFileName
'        Open strServerFileName For Random As #1
'        Close #1
'    End If
    
    'iStopMode =0 Auto End
    'iStopMode =1 Auto Stop when alarms had occured
    'iStopMode =2 Manual Stop
    Select Case iStopMode
        Case 0
                Call frmHistory.AppendLogAlert(1, "Process", 1015, "自動完成", 1)
'                Call frmPlotProcessLog.SaveProcessLog
                Para.strLastRecipe = Kernel.strCurrRecipe
                iniPara.Section = "Utility"
                iniPara.Key = "LastLoadRecipe"
                iniPara.value = Para.strLastRecipe
                
                frmHistory.AppendReport Kernel.strCurrReportTime, "OK", Kernel.strCurrRecipe, Kernel.strBarcodeID, frmPlotProcess.txtBN, frmPlotProcess.txtPN
                If Para.UseBarcodeServer = 1 Then
                    strServerFileName = Kernel.strServerPath & Kernel.strBarcodeID & "_" & Kernel.strServerRecipe & "_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_End.txt"
                    FileCopy Kernel.strCurrLogFile, strServerFileName
                End If
                
                
                If gbintFinishedLight > 0 Then
                    frmConfiguration.tmrFinishedLight.Enabled = True
                Else
                    If gbintFinishedBeep > 0 Then
                        frmConfiguration.tmrFinishedBeep.Interval = gbintFinishedBeep * 1000
                        frmConfiguration.tmrFinishedBeep.Enabled = True
                        SetTower 1, True
                    Else
                        SetTower 2, True
                    End If
                End If
                
                If gbblnAutoCloseValve Then
                    
                    For i = 0 To 5
                        Kernel.sngCurrOutMFC(i) = 0
                    Next i
                    Kernel.sngIntensity = 0
                    Call ResetAO
                    Call frmHistory.AppendLogAlert(1, "Process", 1018, "自動關MFC", 1)
                End If
                
                If Para.UseCover = 1 And gbblnRecipeEndOpenCover = True Then
                    If SetCover(False) = True Then
                        Call frmHistory.AppendLogAlert(1, "Process", 1019, "自動開蓋板", 1)
                    End If
                End If
                
                If gbblnRecipeEndAutoOpen = True Then
                    
                    If SetDoor(1) = False Then
                        Kernel.IsRun = 0
                        ShowAlarmFlash 11
                        Exit Function
                    Else
                        Call frmHistory.AppendLogAlert(1, "Process", 1016, "自動開門", 1)
                    End If
                End If
                
                Kernel.IsRun = 0
                Kernel.intCurrMonitorRun = Kernel.intCurrMonitorRun + 1
                If Para.UseAutoMode = 1 Then
'                    WriteLog ("執行加工完成---#PR=2")
                    frmUDP.wsServer.sendData "#PR=2,"
                End If
                                
                If Para.UseCIM = 1 Then
                    frmCIM.Send "$SPR=0,"
                End If
                If Para.intCycleRuns > 0 And Kernel.intCurrCycleRun > 0 And Kernel.intCurrCycleRun < Para.intCycleRuns Then
'                    mdifrmRTP.tbrRTP_ButtonClick mdifrmRTP.tbrRTP.Buttons("iRun")
                Else
                    If gbblnRecipeFinishedClear = True Then
                        frmPlotProcess.PlotProcessChartClean
                    End If
                End If
                
                If StopTCM = 1 Then
                     For j = 0 To 3
                        data(j) = 1
                    Next j
                    Call frmAz1.WriteParas(109, data, False)
                    Call frmAz2.WriteParas(109, data, False)
                End If
                
        Case 1
            If Kernel.IsRun = 1 Then
                Call frmHistory.AppendLogAlert(1, "Process", 1017, "自動停止", 1)
'                Call frmPlotProcessLog.SaveProcessLog
                If Para.UseBarcodeServer = 1 Then
                    strServerFileName = Kernel.strServerPath & Kernel.strBarcodeID & "_" & Kernel.strServerRecipe & "_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_Stop.txt"
                    FileCopy Kernel.strCurrLogFile, strServerFileName
                End If
                frmHistory.AppendReport Kernel.strCurrReportTime, "Broken", Kernel.strCurrRecipe, Kernel.strBarcodeID, frmPlotProcess.txtBN, frmPlotProcess.txtPN
                
                For i = 0 To 5
                    Kernel.sngCurrOutMFC(i) = 0
                Next i
                Kernel.sngIntensity = 0
                Call ResetAO
                
                If Para.UseAutoMode = 1 Then
                    frmUDP.wsServer.sendData "#PR=0,"
                End If
                If Para.UseCIM = 1 Then
                    frmCIM.Send "$SPR=3,"
                End If
                
                If Para.UseCover = 1 And gbblnRecipeEndOpenCover = True Then
                    If SetCover(False) = True Then
                        Call frmHistory.AppendLogAlert(1, "Process", 1019, "自動開蓋板", 1)
                    End If
                End If
            End If
            Kernel.intCurrCycleRun = 0
               If Para.intCycleRuns > 1 Then ManualStop = True
        Case 2
            If Kernel.IsRun = 1 Then
                Call frmHistory.AppendLogAlert(1, "Manual", 1014, "手動停止", 1)
'                Call frmPlotProcessLog.SaveProcessLog
                If Para.UseBarcodeServer = 1 Then
                    strServerFileName = Kernel.strServerPath & Kernel.strBarcodeID & "_" & Kernel.strServerRecipe & "_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_Stop.txt"
                    FileCopy Kernel.strCurrLogFile, strServerFileName
                End If
                frmHistory.AppendReport Kernel.strCurrReportTime, "Broken", frmRecipeEdit.lbRecipeName.Caption, Kernel.strBarcodeID, frmPlotProcess.txtBN, frmPlotProcess.txtPN
                If Para.UseCIM = 1 Then
                    frmCIM.Send "$SPR=2,"
                End If
            Else
                If Para.UseCIM = 1 Then
                    frmCIM.Send "$SPR=5,"
                End If
            End If
            SetDoor 2   'unlock the door
            SetTower 2, True
            SetAngle False
            
            If Para.useTPump = 0 Then SetPump False
            frmConfiguration.tmrFinishedBeep.Enabled = False
            frmConfiguration.tmrFinishedLight.Enabled = False
            Kernel.intCurrCycleRun = 0
            If Para.UseAutoMode = 1 Then
                frmUDP.wsServer.sendData "#PR=0,"
            End If
            
            For i = 0 To 5
                Kernel.sngCurrOutMFC(i) = 0
            Next i
            Kernel.sngIntensity = 0
            Call ResetAO
            
            If Para.UseCover = 1 And gbblnRecipeEndOpenCover = True Then
                If SetCover(False) = True Then
                    Call frmHistory.AppendLogAlert(1, "Process", 1019, "自動開蓋板", 1)
                End If
            End If
            If Para.intCycleRuns > 1 Then ManualStop = True
            Call ShowTcmModule
              
    End Select
    
    
    
    If Para.UseCIM = 1 Then
        Kernel.strServerRecipe = ""
        CurrProc.strWaferID(0) = ""
    End If
             
             
    If Kernel.intCurrCycleRun = 0 Then
        Kernel.IsRun = 0
        SetAngle False
                
        frmDiagnosis.tmrPumpON.Enabled = False
        If SysDO.IsAngle Then
            SetAngle False
        End If
        
        
        
'        If SysDO.IsPumping Then
'            frmDiagnosis.tmrPumpOFF.Enabled = True
'        End If
        Call SetLampCooling(False)
        Call mdifrmRTP.ShowTitleBar(True)
    End If
    
    'Call ResetAO
    'Kernel.sngIntensity = 0
    Kernel.strCurrStep = ""
    Kernel.IsPurge = 0
    Kernel.intCurrStep = 0
    Kernel.lngCurrStepCount = 0
    gbblnPlayFakeBall = False
    gbintPlayFakeBall = 0
            
    If Az1.blnUseAzbil Then
        gbintAz1ProcNo = 0
    End If
    If Az2.blnUseAzbil Then
        gbintAz2ProcNo = 0
    End If
    
    
    frmProcess.tmrProcessStep.Enabled = False
    SetDO gblngDO_ARM_FRONT, False
    SetDO gblngDO_APCGaugeAngle, False
    gbblnPumpDownTimeout = False
    CurrProc.blnOxygenTimeout = False
    frmDiagnosis.tmrPurge.Enabled = False
    frmConfiguration.StartWatchDog
    
    Kernel.isEnd = True
    Exit Function
ERRLINE:
    ShowAlarmFlash 17
End Function

Public Function LoadPara() As Boolean
    
    Dim i As Integer
    
    On Error GoTo ERR_PARA_OPEN
    
       
    iniPara.Path = gbSystemPath & "\System\system.cfg"
    
    
    iniPara.Section = "PARAMETER"
    iniPara.Default = 0
    iniPara.Key = "UseAutoMode"
    Para.UseAutoMode = CInt(iniPara.value)
    iniPara.Key = "UseCT"
    Para.UseCT = CInt(iniPara.value)
    iniPara.Key = "UseMTC"
    Para.UseMTC = CInt(iniPara.value)
    iniPara.Key = "UseMTCB"
    Para.UseMTCB = CInt(iniPara.value)
    iniPara.Key = "UseCIM"
    Para.UseCIM = CInt(iniPara.value)
    iniPara.Key = "UseAz1"
    Para.UseAz1 = CInt(iniPara.value)
    iniPara.Key = "UseAz2"
    Para.UseAz2 = CInt(iniPara.value)
    iniPara.Key = "UseTPump"
    Para.useTPump = CInt(iniPara.value)
    iniPara.Key = "UseCover"
    Para.UseCover = CInt(iniPara.value)
    
    
    iniPara.Key = "UseBarcodeServer"
    Para.UseBarcodeServer = CInt(iniPara.value)
    iniPara.Default = "Y:\"
    iniPara.Key = "ServerPath"
    Para.strServerPath = iniPara.value
    iniPara.Key = "CycleRun"
    Para.intCycleRuns = CInt(iniPara.value)
    iniPara.Default = 10
    iniPara.Key = "MonitorRuns"
    Para.intMonitorRuns = CInt(iniPara.value)
    iniPara.Default = 3
    iniPara.Key = "ComCT"
    Para.intComCT = CInt(iniPara.value)
    iniPara.Default = ""
    iniPara.Key = "TestRunKey"
    Para.strTestRunKey = iniPara.value
    iniPara.Default = 0
    iniPara.Key = "IsHoldSafety"
    Para.IsHoldSafety = iniPara.value
    iniPara.Key = "IsCali"
    Para.IsCali = iniPara.value
    iniPara.Key = "GaugeAngle"
    Para.sngGaugeAngle = iniPara.value
    iniPara.Key = "OnlyRecipe"
    Para.intOnlyRecipe = iniPara.value
    iniPara.Key = "O2Gate"
    Para.sngO2Gate = iniPara.value
    iniPara.Key = "OpenDoorTime"
    Para.intOpenDoorTime = iniPara.value
    iniPara.Default = 2
    iniPara.Key = "MonitorIndex"
    Para.intMonitorIndex = iniPara.value
    iniPara.Default = 20
    iniPara.Key = "LampAlarmTime"
    Para.intLampAlarmTime = iniPara.value
    iniPara.Default = 168
    iniPara.Key = "AutoPort"
    Para.intAutoPort = iniPara.value
    iniPara.Default = 0
    iniPara.Key = "CIMPort"
    Para.intCIMPort = iniPara.value
    iniPara.Default = 1
    iniPara.Key = "PumpDelay"
    Para.intPumpDelay = iniPara.value
    
    iniPara.Default = 1
    iniPara.Key = "PMbig"
    Para.intPMbig = iniPara.value
    
    iniPara.Default = 3
    iniPara.Key = "PMsmall"
    Para.intPMsmall = iniPara.value
    
    iniPara.Default = "192.168.0.11"
    iniPara.Key = "AzIP1"
    Para.strAzIP1 = iniPara.value
    iniPara.Default = "192.168.0.12"
    iniPara.Key = "AzIP2"
    Para.strAzIP2 = iniPara.value
    
    iniPara.Default = "255.255.255.255"
    iniPara.Key = "RobotIP"
    Para.strRobotIP = iniPara.value
    
    iniPara.Default = 0
    iniPara.Key = "GaugeD"
    Para.sngGaugeD = iniPara.value
    iniPara.Key = "GaugeVN"
    Para.sngGaugeVN = iniPara.value
    iniPara.Default = 10
    iniPara.Key = "GaugeVP"
    Para.sngGaugeVP = iniPara.value
    
        
    iniPara.Section = "Utility"
    iniPara.Key = "RtaType"
    Para.RtaType = iniPara.value
    iniPara.Key = "LastLoadRecipe"
    Para.strLastRecipe = iniPara.value
    
    iniPara.Default = 0
    iniPara.Section = "Custom"
    iniPara.Key = "UseCustom"
    Para.IsUseCustom = iniPara.value
    iniPara.Default = 1
    iniPara.Key = "RatioCUP"
    Para.sngRatioCUP = CSng(iniPara.value)
    For i = 0 To 9
        iniPara.Key = "RatioCUT" & CStr(i)
        Para.sngRatioCUT(i) = CSng(iniPara.value)
        iniPara.Key = "RatioCUM" & CStr(i)
        Para.sngRatioCUM(i) = CSng(iniPara.value)
    Next i
    
        
    iniPara.Section = "Alarm"
    iniPara.Default = 4
    For i = 1 To 33
                
        iniPara.Key = "AlarmDo_" & CStr(4000 + i)
        Para.AlarmDo(i) = CInt(iniPara.value)
        
    Next i
    
      
    Exit Function
ERR_PARA_OPEN:
    ShowMessageOK "Parameter檔案開啟失敗"
    
End Function

Public Function SavePara() As Boolean
    
    Dim i As Integer
    
    On Error GoTo ERR_PARA_SAVE
    
    iniPara.Path = gbSystemPath & "\System\system.cfg"
    iniPara.Section = "PARAMETER"
    iniPara.Key = "UseAutoMode"
    iniPara.value = Para.UseAutoMode
    iniPara.Key = "UseCT"
    iniPara.value = Para.UseCT
    iniPara.Key = "UseMTC"
    iniPara.value = Para.UseMTC
    iniPara.Key = "UseMTCB"
    iniPara.value = Para.UseMTCB
    iniPara.Key = "UseCIM"
    iniPara.value = Para.UseCIM
    iniPara.Key = "UseBarcodeServer"
    iniPara.value = Para.UseBarcodeServer
    iniPara.Key = "ServerPath"
    iniPara.value = Para.strServerPath
    iniPara.Key = "ComCT"
    iniPara.value = Para.intComCT
    iniPara.Key = "CycleRun"
    iniPara.value = Para.intCycleRuns
    iniPara.Key = "MonitorRuns"
    iniPara.value = Para.intMonitorRuns
    iniPara.Key = "TestRunKey"
    iniPara.value = Para.strTestRunKey
    iniPara.Key = "IsHoldSafety"
    iniPara.value = Para.IsHoldSafety
    iniPara.Key = "IsCali"
    iniPara.value = Para.IsCali
    
    iniPara.Key = "GaugeAngle"
    iniPara.value = Para.sngGaugeAngle
    iniPara.Key = "MonitorIndex"
    iniPara.value = Para.intMonitorIndex
    iniPara.Key = "O2Gate"
    iniPara.value = Para.sngO2Gate
    iniPara.Key = "OpenDoorTime"
    iniPara.value = Para.intOpenDoorTime
    iniPara.Key = "LampAlarmTime"
    iniPara.value = Para.intLampAlarmTime
    iniPara.Key = "AutoPort"
    iniPara.value = Para.intAutoPort
    iniPara.Key = "CIMPort"
    iniPara.value = Para.intCIMPort
    iniPara.Key = "PumpDelay"
    iniPara.value = Para.intPumpDelay
    
    iniPara.Key = "OnlyRecipe"
    iniPara.value = Para.intOnlyRecipe
    iniPara.Key = "PMbig"
    iniPara.value = Para.intPMbig
    iniPara.Key = "PMsmall"
    iniPara.value = Para.intPMsmall
    
    iniPara.Key = "UseAz1"
    iniPara.value = Para.UseAz1
    iniPara.Key = "UseAz2"
    iniPara.value = Para.UseAz2
    iniPara.Key = "UseTPump"
    iniPara.value = Para.useTPump
    
    iniPara.Key = "AzIP1"
    iniPara.value = Para.strAzIP1
    iniPara.Key = "AzIP2"
    iniPara.value = Para.strAzIP2
    iniPara.Key = "RobotIP"
    iniPara.value = Para.strRobotIP
    
    iniPara.Key = "GaugeD"
    iniPara.value = Para.sngGaugeD
    iniPara.Key = "GaugeVP"
    iniPara.value = Para.sngGaugeVP
    iniPara.Key = "GaugeVN"
    iniPara.value = Para.sngGaugeVN
    
    iniPara.Key = "UseCover"
    iniPara.value = Para.UseCover
           
    iniPara.Section = "Utility"
    iniPara.Key = "RtaType"
    iniPara.value = Para.RtaType
    iniPara.Key = "LastLoadRecipe"
    iniPara.value = Para.strLastRecipe
    
    iniPara.Section = "Custom"
    iniPara.Key = "UseCustom"
    iniPara.value = CStr(Para.IsUseCustom)
    iniPara.Key = "RatioCUP"
    iniPara.value = CStr(Para.sngRatioCUP)
    For i = 0 To 9
        iniPara.Key = "RatioCUT" & CStr(i)
        iniPara.value = CStr(Para.sngRatioCUT(i))
        iniPara.Key = "RatioCUM" & CStr(i)
        iniPara.value = CStr(Para.sngRatioCUM(i))
    Next i
    
        
        
        
    iniPara.Section = "Alarm"
    iniPara.Default = 4
    For i = 1 To 33
        iniPara.Key = "AlarmDo_" & CStr(4000 + i)
        iniPara.value = Para.AlarmDo(i)
    Next i
    
    
    Exit Function
ERR_PARA_SAVE:
    ShowMessageOK "Parameter檔案開啟失敗"
    
End Function

Public Function SaveDebugLog(StepName As String, StepNo As Integer) As Boolean
    
    If IsDebugMode = True Then
        iniPara.Section = "Debug"
        iniPara.Key = StepName
        iniPara.value = StepNo
    End If
    
End Function


Public Function CalMultiRT(Index As Integer) As Boolean
   
    Dim i As Integer
    Dim lngRet                As Long
    
    'If gbintLoginRight = 1 And MultiLoop.blnUseMultiLoop = True Then
    If MultiLoop.blnUseMultiLoop = True Then
        If MultiLoop.blnUseLoop(Index) = True Then
            Dim tc As Double
            Dim mtc As Double
            Dim sum As Double
                        
            tc = Kernel.sngTC(MultiLoop.intLoopTC(Index))
            sum = 0
            If MultiLoop.intLoopMA(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMA(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMB(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMB(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMC(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMC(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMD(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMD(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopME(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopME(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMF(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMF(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMG(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMG(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMH(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMH(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMJ(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMJ(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
'            If MultiLoop.intLoopMK(Index) > 0 Then
'                mtc = Kernel.sngTC(MultiLoop.intLoopMK(Index))
'                sum = sum + mtc / tc
'                i = i + 1
'            End If
            If i > 0 Then
                sum = sum / i
                MultiLoop.sngLoopRT(Index) = MultiLoop.sngLoopRT(Index) * sum
                
                lngRet = WritePrivateProfileString("MultiLoop", "RT" & CStr(Index), CStr(MultiLoop.sngLoopRT(Index)), Kernel.strCurrRecipeFile)
                
                CalMultiRT = True
                
            End If
           
        End If
        
    End If
    CalMultiRT = False
    
End Function

Public Function CalAzbilRT(Index As Integer) As Boolean

    Dim i As Integer
    Dim ii As Integer
    Dim lngRet                As Long
    
    ii = Index
    If Index > 3 And Index < 8 Then ii = Index - 4
    
    If Kernel.IsRun = 1 Then
        Dim tc As Double
        Dim mtc As Double
        Dim sum As Double
                    
        tc = Kernel.sngTC(Index)
        sum = 0
        If MultiLoop.intLoopMA(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMA(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMB(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMB(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMC(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMC(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMD(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMD(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopME(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopME(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMF(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMF(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMG(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMG(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMH(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMH(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        If MultiLoop.intLoopMJ(Index) > 0 Then
            mtc = Kernel.sngTC(MultiLoop.intLoopMJ(Index) - 1)
            sum = sum + mtc / tc
            i = i + 1
        End If
        
        If i <= 0 Then CalAzbilRT = False
        
        If Index >= 0 And Index < 4 Then
         If i <> 0 Then
           sum = sum / i
            Az1.sngRT(Index) = Az1.sngRT(Index) * sum
            If Az1.sngRT(Index) = 0 Then Az1.sngRT(Index) = 1
            lngRet = WritePrivateProfileString("Azbil", "Az1RT" & CStr(Index), CStr(Az1.sngRT(Index)), Kernel.strCurrRecipeFile)
         End If
        Else
        If i <> 0 Then
            sum = sum / i
            Az2.sngRT(Index - 4) = Az2.sngRT(Index - 4) * sum
            If Az2.sngRT(Index - 4) = 0 Then Az2.sngRT(Index - 4) = 1
            lngRet = WritePrivateProfileString("Azbil", "Az2RT" & CStr(Index - 4), CStr(Az2.sngRT(Index - 4)), Kernel.strCurrRecipeFile)
        End If
        
        End If
                     
    End If
    
End Function







