Attribute VB_Name = "mdlProcess"
Option Explicit

Public Const GB_PROC_ACTION_NONE = -1
Public Const GB_PROC_ACTION_IDLE = 0
Public Const GB_PROC_ACTION_PREHEAT = 1
Public Const GB_PROC_ACTION_RAMPUP = 2
Public Const GB_PROC_ACTION_HOLD = 3
Public Const GB_PROC_ACTION_RAMPDOWN = 4
Public Const GB_PROC_ACTION_STOP = 5
Public Const GB_PROC_ACTION_PUMPDOWN = 6
Public Const GB_PROC_ACTION_PUMPDOWNKEEP = 7
Public Const GB_PROC_ACTION_PUMPMANUAL = 8
Public Const GB_PROC_ACTION_COOLING = 9
Public Const GB_MAX_MSEC = 65535

Public pblngScanTime    As Long     'mSec
Public gbintCurrProcessStep  As Integer
Public gbsngPump                         As Single

'============================================================================================================================
'Define Process Timer
'============================================================================================================================
Public gbdblTotalProcessTime As Double
Public gbdblProcessTimeStamp(255) As Double
Public gbdblStartProcessFlag As Double
Public gbdblStartPumpDownProcessFlag As Double
Public gbdblStartHeatingProcessFlag As Double
Public gbdblStartRampupProcessFlag As Double

Public gbdblProcessRampTime As Double
Public gbdblProcessHoldTime As Double
Public gbdblProcessPumpDownTimerout As Double
Public gbdblProcessPreheatTimerout As Double
Public gbdblProcessTimeFlag As Double
Public gbdblProcessTimerout As Double
Public gbdblProcessKeepTime As Double
Public gbdblLampLifeTime As Double

Public gbsngProcessSlope As Single
Public gbblnStartHeatingProcess As Boolean
Public gbblnPumpDownTimeout As Boolean


'============================================================================================================================
'Define Run Process Timer Event ID
'============================================================================================================================
Public Const TIMER_EVENT_ID = 4999 'Run Process Timer Event ID
Public Const TIMER_EVENT_PROCESS_ID = 5000 'Run Process Timer Event ID
Public Const TIMER_EVENT_CONTROL_ID = 5001  'Run PID Timer Event ID
Public Const TIMER_EVENT_IDLE_ID = 5002  'Run PID Timer Event ID
Public Const TIMER_EVENT_PUMPDOWN_ID = 5003  'Run PID Timer Event ID
Public Const TIMER_EVENT_PREHEAT_ID = 5004  'Run PID Timer Event ID
Public Const TIMER_EVENT_RAMPUP_ID = 5005  'Run PID Timer Event ID
Public Const TIMER_EVENT_HOLD_ID = 5006  'Run PID Timer Event ID
Public Const TIMER_EVENT_VENT_ID = 5007  'Run PID Timer Event ID
Public Const TIMER_EVENT_PURGE_ID = 5008  'Run PID Timer Event ID
Public Const TIMER_EVENT_STOP_ID = 5009  'Run PID Timer Event ID
Public Const TIMER_EVENT_RAMPDOWN_ID = 5010  'Run PID Timer Event ID

'============================================================================================================================
'Define Process recipe for processing
'============================================================================================================================
    
Public Type Process_Recipe
    lngAction                   As Long
    sngTime                    As Single
    sngPump                  As Single
    sngTemperature      As Single
    sngGas(10)                As Single
End Type
    
Public gbProcessRecipe()            As Process_Recipe
Public gbProcessSecondCount         As Long 'Record proceess second
Public gbProcessRecordCount         As Long 'Record proceess second
Public gbProcessDataCount         As Long 'Record proceess second

'--------------------------------------------------------------------
'Define Process recipe step for processing
'--------------------------------------------------------------------
Public Type Process_Recipe_Step
    strAction                       As String
    sngTime                         As Single
    sngPump                         As Single
    sngTemperature                  As Single
    sngGas(10)                       As Single
End Type

Public gbProcessRecipeStep()            As Process_Recipe_Step

'============================================================================================================================
'Define the ramp smooth variable
'============================================================================================================================
Public gbsngRampSmoothDist As Single
Public gbsngRampsmoothTheta As Single
Public gbsngRampSmoothCx As Single
Public gbsngRampSmoothCy As Single
Public gbsngRampSmoothR As Single
Public gbsngRampSmoothDistX As Single
Public gbsngRampSmoothDistY As Single
Public gbsngRampSmoothStart As Single
Public gbsngRampSmoothMid As Single
Public gbsngRampSmoothEnd As Single

'============================================================================================================================
'Define the Uniformity variable Rev4.1.6
'============================================================================================================================
Public gbsngUniformityDistA As Single
Public gbsngUniformityDistB1 As Single
Public gbsngUniformityDistB2 As Single
Public gbsngUniformityDistX As Single
Public gbsngUniformitySlopeA1 As Single
Public gbsngUniformitySlopeA2 As Single

'============================================================================================================================
Public m_blnStartPIDLoop               As Boolean
Public gblngHoldStartTime As Long

Private m_intProcessStep               As Integer
Public bExecuted1 As Boolean
Public bExecuted2 As Boolean
Private m_sngOutPut                         As Single
Private m_lngControlInterval               As Long
Private m_lngPreTime                        As Long
Private m_lngCurrTime                       As Long
Public m_sngSetTemperature                    As Single

Private blnResetVI As Boolean
Private blnCheckOverCT As Boolean
Private blnCheckUnderCT As Boolean
'Public gbsngProcessRecorder(65535, GB_MAX_DRAW_COL + 1) As Single
Public gbsngProcessRecorder(GB_MAX_DRAW_COL + 1) As Single
Public gbsngDrawData(GB_MAX_DRAW_COL + 1) As Single
Public gbintRampHoldCount As Integer
Public gbintRampCount As Integer
Public sngCurrTemp(10)                    As Single

Public Function InitProcessStep() As Boolean
    Dim i As Integer
    Dim lngGetTime As Long
    HoldCount = 0
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    If lngGetTime <= 0 Then GoTo ERRLINE
    
    gbblnStartHeatingProcess = False
    gbdblStartProcessFlag = 0
    gbdblStartHeatingProcessFlag = 0
    gbdblProcessTimeFlag = 0
    gbdblProcessKeepTime = 0
    gbdblLampLifeTime = 0
    gblngPumpDownTime = 0
    gbProcessRecordCount = 0
    gbProcessDataCount = 0
    gbintRampHoldCount = 0
    gbintRampCount = 0
    
    blnResetVI = True
    
    gbdblStartProcessFlag = lngGetTime
    gbdblStartRampupProcessFlag = lngGetTime
    m_lngPreTime = lngGetTime
    gbdblProcessTimeStamp(0) = gbdblStartProcessFlag
    
    For i = 1 To UBound(gbProcessRecipeStep)
        gbdblProcessTimeStamp(i) = 0
    Next i
    For i = 0 To gbintNumOfBanks - 1
        Kernel.sngCurrOutSCR(i) = 0
    Next i
    For i = 0 To GB_MAX_LOOPS - 1
        MultiLoop.sngLoopOut(i) = 0
        MultiLoop.blnLoopReset(i) = True
    Next i
    frmPlotProcessLog.SaveCurProcessLog (False)
    RecordCount = 0
    CurrProc.lngScanTime = 15 'mSec
    CurrProc.intStep = 1
    CurrProc.blnDoStep = True
    frmProcess.tmrProcessStep.Enabled = True
    Exit Function
'    InitProcessStep = True
'
'    m_intProcessStep = 1
'    Call RunProcessStep(m_intProcessStep)
ERRLINE:
    gbstrAlarmHint = " InitProcessStep error"
    ShowAlarmFlash 1

End Function

Public Function RunProcessStep(intStep As Integer) As Boolean
    Dim strAction                           As String
    Dim sngTime                           As Single
    Dim sngTemperature             As Single
    Dim lngCurrentTime              As Long
    Dim sngRampSmoothResult(5) As Single
    Dim sngGas(6) As Single
    Dim i                               As Integer
    Dim lngRet          As Long
    Dim StrFileName     As String
    Dim strTemp     As String
    Dim tmrStamp As Single
    Dim bol As Boolean
    Dim SysAction As String
    
        
    On Error Resume Next
    
    StrFileName = gbSystemPath & "\System\system.cfg"
    strAction = gbProcessRecipeStep(intStep).strAction
    sngTime = gbProcessRecipeStep(intStep).sngTime
    sngTemperature = gbProcessRecipeStep(intStep).sngTemperature
    m_sngOutPut = 0
    gbsngPump = 0
    SaveDebugLog "Start", 1
    SysAction = Readini(strAction)
    If SysAction <> "" Then
    strAction = SysAction
    End If
    If strAction <> GB_ACTION_STOP Then
        For i = 0 To 5
            sngGas(i) = gbProcessRecipeStep(m_intProcessStep).sngGas(i)
            If gbstrGasAlias(i) = "Pump" Then
                If gbProcessRecipeStep(m_intProcessStep).sngGas(i) > 0 Or gbProcessRecipeStep(m_intProcessStep).sngGas(i) < 0 Then
                    gbsngPump = gbProcessRecipeStep(m_intProcessStep).sngGas(i)
                    
                    If SetAngle(True) > 0 Then
                    
                        If SysDI.IsChamberGaugeL = 1 Then frmDiagnosis.tmrPumpON.Enabled = True
                    End If
                Else
                    SetAngle False
                    'frmDiagnosis.tmrPumpOFF.Enabled = True
                    SetPump False
                    If gbintReleaseOpenDelay > 0 Then frmDiagnosis.tmrSetReleaseOFF.Enabled = True
                    SaveDebugLog "Start", 4
                End If
                sngGas(i) = 0
            End If
        Next i
        SetAO_MFC sngGas
        SaveDebugLog "Start", 5
    End If
    SetDO gblngDO_ARM_FRONT, False
    If Para.IsHoldSafety = 1 Then
        frmDiagnosis.tmrHoldSafeON.Enabled = True
    End If
    blnCheckOverCT = True
    blnCheckUnderCT = False
    lngCurrentTime = timeGetTime
    SaveDebugLog "Start", 6
    Select Case strAction
        Case GB_ACTION_IDLE
            pblngScanTime = 1000 'mSec
            gbintCurrProcessStep = GB_ACTION_INDEX_IDLE
            gbdblProcessTimeStamp(m_intProcessStep) _
                = lngCurrentTime + gbProcessRecipeStep(m_intProcessStep).sngTime * 1000
            SetAO_SCR 0, gbsngRecipeIntensityWeightDynamic
            SaveDebugLog "Start", 7
            Call frmHistory.AppendLogAlert(1, "Process", 1001, "Idle process", 1)
            SaveDebugLog "Start", 8
            Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID, pblngScanTime, AddressOf Process_Idle)
            SaveDebugLog "Start", 9
            Kernel.strCurrStep = "(" & CStr(intStep) & ")-IDLE"
            Kernel.intCurrStep = intStep
            
            
        Case GB_ACTION_RAMPUP
            pblngScanTime = 1 'mSec
            gbintCurrProcessStep = GB_ACTION_INDEX_RAMPUP
            gbintRampHoldCount = gbintRampHoldCount + 1
            m_blnStartPIDLoop = False
            If gbblnResetInteral = True Then Call InitPIDParameter
            SaveDebugLog "Start", 10
            gbdblProcessTimeStamp(m_intProcessStep) _
                = lngCurrentTime + gbProcessRecipeStep(m_intProcessStep).sngTime * 1000
            '130109 Josh
            gbdblLampLifeTime = gbdblLampLifeTime + gbProcessRecipeStep(m_intProcessStep).sngTime
            SaveDebugLog "Start", 11
            If gbProcessRecipeStep(intStep + 1).strAction = GB_ACTION_HOLD And gbsngSmoothTime > 0 Then
                gbsngRampSmoothDist = gbsngSmoothTime * 1000
                Call CalSmoothCurve(gbsngRampSmoothDist, CSng(lngCurrentTime), gbProcessRecipeStep(intStep - 1).sngTemperature, CSng(gbdblProcessTimeStamp(m_intProcessStep)), gbProcessRecipeStep(intStep).sngTemperature, sngRampSmoothResult)
                'Rev4.1.6
                Call CalUniformityCompensation
            End If
            SaveDebugLog "Start", 12
            Call frmHistory.AppendLogAlert(1, "Process", 1002, "RampUp process", 1)
            SaveDebugLog "Start", 13
            Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID, pblngScanTime, AddressOf Process_RampUp)
            SaveDebugLog "Start", 14
            Kernel.strCurrStep = "(" & CStr(intStep) & ")-Ramp Up"
            Kernel.intCurrStep = intStep
            
        Case GB_ACTION_HOLD 'Hold
            pblngScanTime = 1 'mSec
            gbintCurrProcessStep = GB_ACTION_INDEX_HOLD
            m_blnStartPIDLoop = False
            gbdblLampLifeTime = gbdblLampLifeTime + gbProcessRecipeStep(m_intProcessStep).sngTime
            
            gbdblProcessTimeStamp(m_intProcessStep) _
                = lngCurrentTime + gbProcessRecipeStep(m_intProcessStep).sngTime * 1000
            SaveDebugLog "Start", 16
            gblngHoldStartTime = timeGetTime
            SaveDebugLog "Start", 17
            Call frmHistory.AppendLogAlert(1, "Process", 1003, "Hold process", 1)
            SaveDebugLog "Start", 18
            Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID, pblngScanTime, AddressOf Process_Hold)
            SaveDebugLog "Start", 19
            Kernel.strCurrStep = "(" & CStr(intStep) & ")-Hold"
            Kernel.intCurrStep = intStep
            
        Case GB_ACTION_STOP 'Stop
            gbintCurrProcessStep = GB_ACTION_INDEX_STOP
            gbsngUsedLamp = gbsngUsedLamp + gbdblLampLifeTime
            lngRet = WritePrivateProfileString("PARAMETER", "UsedLamp", CStr(gbsngUsedLamp), StrFileName)
            gbblnAutoCloseValve = gbblnRecipeAutoCloseValve1
            Kernel.strCurrStep = "(" & CStr(intStep) & ")-Stop"
            
            SaveDebugLog "Start", 20
            Call frmHistory.AppendLogAlert(1, "Process", 1004, "Stop process", 1)
            SaveDebugLog "Start", 21
            Call Process_Stop
            SaveDebugLog "Start", 22
        Case GB_ACTION_RAMPDOWN 'Ramp Down
            gbintCurrProcessStep = GB_ACTION_INDEX_RAMPDOWN
            gbdblProcessTimeStamp(m_intProcessStep) _
                = lngCurrentTime + gbProcessRecipeStep(m_intProcessStep).sngTime * 1000
            Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPDOWN_ID, pblngScanTime, AddressOf Process_RampDown)
            SaveDebugLog "Start", 23
            Call frmHistory.AppendLogAlert(1, "Process", 1005, "RampDown process", 1)
            SaveDebugLog "Start", 24
            Kernel.strCurrStep = "(" & CStr(intStep) & ")-RampDown"
            Kernel.intCurrStep = intStep
    End Select
    
End Function

Public Function Process_Idle() As Boolean
    Dim i As Integer
    Dim bRet As Boolean
    Dim sngTemp(10) As Single
    Dim strTemp As String
    Dim lngGetTime As Long
    
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "Idle", 1
    
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) And Kernel.IsRun = 1 Then
        If CurrProc.sngPump > 0 And CurrProc.sngPump < 100 Then
            If Kernel.sngPressure > CurrProc.sngPump Then
                If gbdblProcessPumpDownTimerout > 0 Then
                    If gbblnPumpDownTimeout = False Then
                        gbdblStartPumpDownProcessFlag = lngGetTime
                        strTemp = "Pressure Check (" & Format(Kernel.sngPressure, "0.000") & ")Torr"
                        Call frmHistory.AppendLogAlert(1, "Check", 3053, strTemp, 1)
                        strTemp = "Va=" & Format(CurrProc.sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        If Para.RtaType = 9 Then
                            If Az1.blnUseAzbil Then
                                gbintAz1ProcNo = 3
                            End If
                            If Az2.blnUseAzbil Then
                                gbintAz2ProcNo = 3
                            End If
                        End If
                        gbblnPumpDownTimeout = True
    
                    Else
                        If (lngGetTime - gbdblStartPumpDownProcessFlag) > gbdblProcessPumpDownTimerout Then
                            gbblnPumpDownTimeout = False
                            ShowAlarmFlash 6
                            strTemp = "Va=" & Format(CurrProc.sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000") & Format(gbdblProcessPumpDownTimerout, "0")
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            Call frmHistory.AppendLogAlert(1, "Alarm", 3054, strTemp, 1)
                            
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    ShowAlarmFlash 6
                    strTemp = "Va=" & Format(CurrProc.sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000") & Format(gbdblProcessPumpDownTimerout, "0")
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    Call frmHistory.AppendLogAlert(1, "Alarm", 3054, strTemp, 1)
                    Exit Function
                End If
            Else
                
                If gbblnPumpDownTimeout = True Then
                    gbblnPumpDownTimeout = False
                    gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - gbdblStartPumpDownProcessFlag)
                End If
                strTemp = "Pressure Check (" & Format(Kernel.sngPressure, "0.000") & ")Torr"
                Call frmHistory.AppendLogAlert(1, "Check", 3053, strTemp, 1)
                gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                strTemp = "Va=" & Format(CurrProc.sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000") & "/" & Format(gblngPumpDownTime, "0")
                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
                gbdblStartRampupProcessFlag = lngGetTime
                If Para.RtaType = 9 Then
                    If Az1.blnUseAzbil Then
                        gbintAz1ProcNo = 2
                    End If
                    If Az2.blnUseAzbil Then
                        gbintAz2ProcNo = 2
                    End If
                End If
                CurrProc.intStep = CurrProc.intStep + 1
                CurrProc.blnDoStep = True
                Exit Function
            End If
        ElseIf gbblnActivePrepare = True Then
            
            If gbsngRecipePrepareTimeout > 0 Then
                If Kernel.sngOxygen > gbsngRecipePrepareGaugeO2 Then
                    If CurrProc.blnOxygenTimeout = False Then
                        CurrProc.dblStartO2Flag = lngGetTime
                        strTemp = "O2 Check (" & Format(Kernel.sngOxygen, "0.00") & ")ppm"
                        Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                        strTemp = "O2=" & Format(gbsngRecipePrepareGaugeO2, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        CurrProc.blnOxygenTimeout = True
                    Else
                        If (lngGetTime - CurrProc.dblStartO2Flag) > gbsngRecipePrepareTimeout Then
                               
                            CurrProc.blnOxygenTimeout = False
                            strTemp = "O2=" & Format(gbsngRecipePrepareGaugeO2, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                            gbstrAlarmHint = strTemp
                            ShowAlarmFlash 24
                            
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    strTemp = "O2 Check (" & Format(Kernel.sngOxygen, "0.00") & ")ppm"
                    Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                    If CurrProc.blnOxygenTimeout = True Then
                        CurrProc.blnOxygenTimeout = False
                        gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - CurrProc.dblStartO2Flag)
                    End If
                    gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                    strTemp = "O2=" & Format(gbsngRecipePrepareGaugeO2, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00") & "/" & Format(gblngPumpDownTime, "0")
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    
                    Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
                    gbdblStartRampupProcessFlag = lngGetTime
                    CurrProc.intStep = CurrProc.intStep + 1
                    CurrProc.blnDoStep = True
    
                    Exit Function
                End If
            Else
                gbstrAlarmHint = strTemp
                strTemp = "O2=" & Format(Para.sngO2Gate, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                ShowAlarmFlash 24
                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                Exit Function
            End If
        ElseIf gbblnActiveTempDown = True Then
            
            If gbsngRecipeTempDownTimeout > 0 Then
                If Kernel.sngTC(0) > gbProcessRecipeStep(CurrProc.intStep).sngTemperature Then
                    
                    If CurrProc.blnTempDownTimeout = False Then
                        CurrProc.dblStartTempDownFlag = lngGetTime
                        strTemp = "Temp Check (" & Format(Kernel.sngTC(0), "0.0") & ")C"
                        Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                        strTemp = "Temp=" & Format(gbProcessRecipeStep(CurrProc.intStep).sngTemperature, "0.0") & "/" & Format(Kernel.sngTC(0), "0.0")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        CurrProc.blnTempDownTimeout = True
                    Else
                        If (lngGetTime - CurrProc.dblStartTempDownFlag) > gbsngRecipeTempDownTimeout * 1000 Then
                               
                            CurrProc.blnTempDownTimeout = False
                            strTemp = "Temp=" & Format(gbProcessRecipeStep(CurrProc.intStep).sngTemperature, "0.0") & "/" & Format(Kernel.sngTC(0), "0.0")
                            gbstrAlarmHint = strTemp
                            ShowAlarmFlash 9
                            
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
'                        Else
'                            Exit Function
                        End If
                    End If
                Else
                    strTemp = "Temp Check (" & Format(Kernel.sngTC(0), "0.0") & ")C"
                    Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                    If CurrProc.blnTempDownTimeout = True Then
                        CurrProc.blnTempDownTimeout = False
                        gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - CurrProc.dblStartTempDownFlag)
                    End If
                    gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                    strTemp = "Temp=" & Format(gbProcessRecipeStep(CurrProc.intStep).sngTemperature, "0.0") & "/" & Format(Kernel.sngTC(0), "0.0") & "/" & Format(gblngPumpDownTime, "0")
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    
                    Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
                    gbdblStartRampupProcessFlag = lngGetTime
                    CurrProc.intStep = CurrProc.intStep + 1
                    CurrProc.blnDoStep = True
    
                    Exit Function
                End If
            Else
                gbstrAlarmHint = strTemp
                strTemp = "Temp=" & Format(gbProcessRecipeStep(CurrProc.intStep).sngTemperature, "0.0") & "/" & Format(Kernel.sngTC(0), "0.0")
                ShowAlarmFlash 9
                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                Exit Function
            End If
        
        Else
            Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
            gbdblStartRampupProcessFlag = lngGetTime
            m_lngPreTime = lngGetTime
            CurrProc.intStep = CurrProc.intStep + 1
            CurrProc.blnDoStep = True
            
'            m_intProcessStep = m_intProcessStep + 1
'            Call RunProcessStep(m_intProcessStep)
            Exit Function
       
        End If
 
    End If
    
    SaveDebugLog "Idle", 2
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    gbsngProcessSlope = 0
            
    If CurrProc.sngPump < 0 Then
        If Kernel.sngPressure < Abs(CurrProc.sngPump) And blnResetVI = True Then
            blnResetVI = False
            CalVacuumPID True, 0, 0, 0, 0
            SetTime 0
        End If
        If GetTime(0) > gbsngAPCInterval Then
           SetTime 0
           
           sngTemp(0) = CalVacuumPID(False, gbsngAPC_P, gbsngAPC_I, Abs(CurrProc.sngPump), Kernel.sngPressure)
           frmPlotProcess.fraVacFunc.Caption = "溃北-" & CStr(sngTemp(0))
           sngTemp(0) = Percent2Volt(sngTemp(0), 0, 5)
           Call VarControlMFC(gbintAPC_MFC_Port, sngTemp(0))
        End If
    End If
    
    SaveDebugLog "Idle", 3
    
    Call ReadTC
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True Then
            Kernel.sngTC(MultiLoop.intLoopTC(i)) = Kernel.sngTC(MultiLoop.intLoopTC(i)) * MultiLoop.sngLoopRT(i)
        End If
    Next i
    
    CheckAlarm
    
    For i = 0 To 7
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
    
    If gbProcessRecipeStep(CurrProc.intStep).sngTemperature > 0 Then
        If sngCurrTemp(0) < gbProcessRecipeStep(CurrProc.intStep).sngTemperature Then
            Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
            Call frmHistory.AppendLogAlert(1, "Process", 1018, "笷秨放", 1)
            Process_Stop
        End If
    End If
    
     
    
    SaveDebugLog "Idle", 4
    Call frmPlotProcess.ShowStatus
    SaveDebugLog "Idle", 5
    Call RecordProcessData
    SaveDebugLog "Idle", 6
    frmPlotProcessLog.SaveCurProcessLog (True)
    Call frmPlotProcess.DrawCurve
    SaveDebugLog "Idle", 7
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Idle error"
    ShowAlarmFlash 1
        
End Function

Public Function Process_RampUp() As Boolean
    Dim i As Integer
    Dim intSA As Integer
    Dim bRet As Boolean
    Dim sngTemp(GB_SCR_MAX) As Single
    Dim strTemp     As String
    Dim lngGetTime As Long
    Dim intLoopInterval As Integer
    Dim intLamps As Integer
    
    
    On Error GoTo ERRLINE
        
    '糶SCRΩ
'    Dim bExecuted As Boolean
'    If Not bExecuted1 Then
'        If Para.RtaType = 9 And IsUsedSCR = 1 Then
'            frmModBusRtu.WriteRamupSCR
'        End If
'        bExecuted1 = True
'    End If
    
    lngGetTime = timeGetTime
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "RampUp", 1
    
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    If gbblnStartHeatingProcess = False Then
        gbblnStartHeatingProcess = True
        gbdblStartHeatingProcessFlag = lngGetTime
        gbdblStartHeatingProcessFlag = gbdblStartHeatingProcessFlag
    End If
    gbsngProcessSlope = 0
        
    SaveDebugLog "RampUp", 2
    'Waiting the time up, next step
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
        gbdblStartRampupProcessFlag = lngGetTime
        
        CurrProc.intStep = CurrProc.intStep + 1
        CurrProc.blnDoStep = True
'        m_intProcessStep = m_intProcessStep + 1
'        Call RunProcessStep(m_intProcessStep)
        Exit Function
    End If
            
    SaveDebugLog "RampUp", 3
    m_lngCurrTime = lngGetTime
    m_lngControlInterval = m_lngCurrTime - m_lngPreTime
    m_lngPreTime = m_lngCurrTime
    If m_lngControlInterval = 0 Then m_lngControlInterval = 1
    
    If gbsngSmoothTime > 0 Then
        If (m_lngCurrTime > gbsngRampSmoothStart) And gbsngRampSmoothR > 0 Then
            m_sngSetTemperature = gbsngRampSmoothCy + ((gbsngRampSmoothR ^ 2 - (gbsngRampSmoothEnd - m_lngCurrTime) ^ 2) ^ 0.5)
        Else
            If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
                 gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                                     / gbProcessRecipeStep(CurrProc.intStep).sngTime
            Else
                gbsngProcessSlope = 0
            End If
            m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
                                     + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
        End If
        
        gbsngUniformitySubWeightA = gbsngUniformitySlopeA1 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD1
        gbsngUniformitySubWeightB = gbsngUniformitySlopeA2 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD2
    Else
   
        If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
             gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                                 / gbProcessRecipeStep(CurrProc.intStep).sngTime
        Else
            gbsngProcessSlope = 0
        End If
        m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
                                 + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
    End If
    
    SaveDebugLog "RampUp", 4
    
    Call ReadTC
    For i = 0 To 7
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
    
    If Para.RtaType <> 9 Then
        SaveDebugLog "RampUp", 5
        If m_sngSetTemperature >= Kernel.sngTC(0) Then m_blnStartPIDLoop = True
        SaveDebugLog "RampUp", 6
        If m_blnStartPIDLoop = True Then
            
            If gbintRampHoldCount = 1 Then
                SetPIDParameter m_sngSetTemperature, _
                                        frmRecipeEdit.sngRecipeProportional, _
                                        frmRecipeEdit.sngRecipeIntegral, _
                                        frmRecipeEdit.sngRecipeDerivational, _
                                        frmRecipeEdit.sngRecipePredit, _
                                        frmRecipeEdit.sngRecipeFeedForward
            Else
                If frmRecipeEdit.sngRecipeProportional2 > 0 Then
                    SetPIDParameter m_sngSetTemperature, _
                                        frmRecipeEdit.sngRecipeProportional2, _
                                        frmRecipeEdit.sngRecipeIntegral2, _
                                        frmRecipeEdit.sngRecipeDerivational, _
                                        frmRecipeEdit.sngRecipePredit, _
                                        frmRecipeEdit.sngRecipeFeedForward
                Else
                    SetPIDParameter m_sngSetTemperature, _
                                        frmRecipeEdit.sngRecipeProportional, _
                                        frmRecipeEdit.sngRecipeIntegral2, _
                                        frmRecipeEdit.sngRecipeDerivational, _
                                        frmRecipeEdit.sngRecipePredit, _
                                        frmRecipeEdit.sngRecipeFeedForward
                End If
            End If
            CurrProc.sngOutput = PID_Loop(Kernel.sngTC(0), m_lngControlInterval) '* sngOutputScale    'out max 200 degree/sec
        Else
            CurrProc.sngOutput = gbsngIntensityKeep
        End If
        SaveDebugLog "RampUp", 7
        If CurrProc.sngOutput < gbsngIntensityKeep Then CurrProc.sngOutput = gbsngIntensityKeep
        Kernel.sngIntensity = CurrProc.sngOutput
        SetAO_SCR CurrProc.sngOutput, gbsngRecipeIntensityWeightDynamic
        SaveDebugLog "RampUp", 8
    
    Else
        CurrProc.sngOutput = Az1.sngMV(0)
        Kernel.sngIntensity = CurrProc.sngOutput / 10
        
    End If
    
    If CurrProc.sngPump < 0 Then
        If Kernel.sngPressure < Abs(CurrProc.sngPump) And blnResetVI = True Then
            blnResetVI = False
            CalVacuumPID True, 0, 0, 0, 0
            SetTime 0
        End If
        If GetTime(0) > gbsngAPCInterval Then
           SetTime 0
           
           sngTemp(0) = CalVacuumPID(False, gbsngAPC_P, gbsngAPC_I, Abs(CurrProc.sngPump), Kernel.sngPressure)
           frmPlotProcess.fraVacFunc.Caption = "溃北-" & CStr(sngTemp(0))
           sngTemp(0) = Percent2Volt(sngTemp(0), 0, 5)
           Call VarControlMFC(gbintAPC_MFC_Port, sngTemp(0))
 
        End If
    End If
    
    CheckAlarm
    
    
    intLamps = 35
    If Para.RtaType = 9 Then intLamps = 54
    
    
    If CurrProc.blnCheckOverCT = True And gbblnRecipeUseCT = True Then
        For i = 0 To intLamps
            If Lamp2Scr(CLng(i), intSA) = True Then
                If Kernel.dblCT(i) > gbsngRecipeCT(intSA) And Kernel.intOverCT(i) = 0 Then
                    If CurrProc.lngOverTime(i) = 0 Then
                        CurrProc.lngOverTime(i) = lngGetTime
                    
                    ElseIf (lngGetTime - CurrProc.lngOverTime(i)) > Para.intLampAlarmTime * 1000 Then
                        strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngOverTime(i))
                        gbstrAlarmHint = strTemp
                        Kernel.intOverCT(i) = 1
                        ShowAlarmFlash 22
                    End If
                Else
                    Kernel.intOverCT(i) = 0
                    CurrProc.lngOverTime(i) = 0
                End If
            End If
        Next
    End If
    
    intLoopInterval = 1
    If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 3 Then intLoopInterval = 20
    CurrProc.intRecordLoop = CurrProc.intRecordLoop + 1
    If CurrProc.intRecordLoop > intLoopInterval Then
        CurrProc.intRecordLoop = 0
        Call RecordProcessData
    End If
    
    SaveDebugLog "RampUp", 9
    frmPlotProcessLog.SaveCurProcessLog (True)
    Call frmPlotProcess.DrawCurve
    SaveDebugLog "RampUp", 10
    Call frmPlotProcess.ShowStatus
    SaveDebugLog "RampUp", 11
    
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Ramp up Error"
    ShowAlarmFlash 1
    
End Function

Public Function Process_Hold() As Boolean
    Dim i, j As Integer
    Dim sngMiddleError As Single
    Dim sngEdgeError(5) As Single
    Dim bRet As Boolean
    Dim temp As Single
    Dim strTemp As String
    Dim lnTemp As Long
    Dim sngTemp(3) As Single
    Dim sngPump As Single
    Dim intSA As Integer
    Dim lngGetTime As Long
    Dim intLoopInterval As Integer
    Dim intLamps As Integer
    
    On Error GoTo ERRLINE
 
    lngGetTime = timeGetTime
    If lngGetTime <= 0 Then GoTo ERRLINE
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    '糶SCRΩ
   
'    If Not bExecuted2 Then
'        If Para.RtaType = 9 And IsUsedSCR = 1 Then
'            frmModBusRtu.WriteHoldSCR
'        End If
'        bExecuted2 = True
'    End If
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "Hold", 1
    'Waiting the time up, next step
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) Then
        If CurrProc.sngPump > 0 And CurrProc.sngPump <> 100 Or (CurrProc.sngPump = 100 And gbintRampHoldCount = 1 And gbsngPumpDownGate > 0) Then
            sngPump = CurrProc.sngPump
            If CurrProc.sngPump = 100 Then
                sngPump = gbsngPumpDownGate / 1000
            End If
            If Kernel.sngPressure > sngPump And Kernel.sngPressure > 0.001 Then
                If gbdblProcessPumpDownTimerout > 0 Then
                    If gbblnPumpDownTimeout = False Then
                        gbdblStartPumpDownProcessFlag = lngGetTime
                        strTemp = "Pressure Check (" & Format(Kernel.sngPressure, "0.000") & ")Torr"
                        Call frmHistory.AppendLogAlert(1, "Check", 3053, strTemp, 1)
                        strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        If Para.RtaType = 9 Then
                            If Az1.blnUseAzbil Then
                                gbintAz1ProcNo = 3
                            End If
                            If Az2.blnUseAzbil Then
                                gbintAz2ProcNo = 3
                            End If
                        End If
                        gbblnPumpDownTimeout = True
                        
                    Else
                        If (lngGetTime - gbdblStartPumpDownProcessFlag) > gbdblProcessPumpDownTimerout Then
                               
                            gbblnPumpDownTimeout = False
                            strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                            gbstrAlarmHint = strTemp
                            ShowAlarmFlash 6
                            
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            Call frmHistory.AppendLogAlert(1, "Alarm", 3054, strTemp, 1)
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    gbstrAlarmHint = strTemp
                    strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                    ShowAlarmFlash 6
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    Call frmHistory.AppendLogAlert(1, "Alarm", 3054, strTemp, 1)
                    Exit Function
                End If
            Else
                strTemp = "Pressure Check (" & Format(Kernel.sngPressure, "0.000") & ")Torr"
                Call frmHistory.AppendLogAlert(1, "Check", 3053, strTemp, 1)
                If gbblnPumpDownTimeout = True Then
                    gbblnPumpDownTimeout = False
                    gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - gbdblStartPumpDownProcessFlag)
                End If
                gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000") & "/" & Format(gblngPumpDownTime, "0")
                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                
                Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
                gbdblStartRampupProcessFlag = lngGetTime
                If Para.RtaType = 9 Then
                    If Az1.blnUseAzbil Then
                        gbintAz1ProcNo = 2
                    End If
                    If Az2.blnUseAzbil Then
                        gbintAz2ProcNo = 2
                    End If
                End If
                CurrProc.intStep = CurrProc.intStep + 1
                CurrProc.blnDoStep = True
                
                Exit Function
                
                
            End If
        Else
            If Para.sngO2Gate > 0 Then
                If gbsngRecipePrepareTimeout > 0 Then
                    If Kernel.sngOxygen > Para.sngO2Gate Then
                        If CurrProc.blnOxygenTimeout = False Then
                            CurrProc.dblStartO2Flag = lngGetTime
                            strTemp = "O2 Check (" & Format(Kernel.sngOxygen, "0.00") & ")ppm"
                            Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                            strTemp = "O2=" & Format(gbsngRecipePrepareGaugeO2, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            CurrProc.blnOxygenTimeout = True
                        Else
                            If (lngGetTime - CurrProc.dblStartO2Flag) > gbsngRecipePrepareTimeout Then
                                   
                                CurrProc.blnOxygenTimeout = False
                                strTemp = "O2=" & Format(gbsngRecipePrepareGaugeO2, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                                gbstrAlarmHint = strTemp
                                ShowAlarmFlash 24
                                
                                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                                Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                            Else
                                Exit Function
                            End If
                        End If
                    Else
                        strTemp = "O2 Check (" & Format(Kernel.sngOxygen, "0.00") & ")ppm"
                        Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                        If CurrProc.blnOxygenTimeout = True Then
                            CurrProc.blnOxygenTimeout = False
                            gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - CurrProc.dblStartO2Flag)
                        End If
                        gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                        strTemp = "O2=" & Format(gbsngRecipePrepareGaugeO2, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00") & "/" & Format(gblngPumpDownTime, "0")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        
                        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
                        gbdblStartRampupProcessFlag = lngGetTime
                        CurrProc.intStep = CurrProc.intStep + 1
                        CurrProc.blnDoStep = True
        
                        Exit Function
                    End If
                Else
                    gbstrAlarmHint = strTemp
                    strTemp = "O2=" & Format(Para.sngO2Gate, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                    ShowAlarmFlash 24
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                    Exit Function
                End If
            Else
            
            
            
                Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
                gbdblStartRampupProcessFlag = lngGetTime
                CurrProc.intStep = CurrProc.intStep + 1
                CurrProc.blnDoStep = True
                If HoldCount = TempOffset(5) Then GbHoldState = False
                Exit Function
            End If
        End If
    End If
    
    SaveDebugLog "Hold", 2
    
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    
    If gbblnStartHeatingProcess = False Then
        gbblnStartHeatingProcess = True
        gbdblStartHeatingProcessFlag = lngGetTime
    End If
    SaveDebugLog "Hold", 3
    gbsngProcessSlope = 0
    m_lngCurrTime = lngGetTime
    m_lngControlInterval = m_lngCurrTime - m_lngPreTime
    m_lngPreTime = m_lngCurrTime
    If m_lngControlInterval = 0 Then m_lngControlInterval = 1
    SaveDebugLog "Hold", 4
    If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
        gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                            / gbProcessRecipeStep(CurrProc.intStep).sngTime
    Else
        gbsngProcessSlope = 0
    End If
    SaveDebugLog "Hold", 5
    Call ReadTC
    
    For i = 0 To 7
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
    
'    SaveDebugLog "Hold", 6
'    If (Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) Then
'        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
'        Call Process_Abort
'        gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
'        ShowAlarmFlash 9
'        Exit Function
'    End If
    
    SaveDebugLog "Hold", 7
    If gbsngSmoothTime > 0 Then
        If (m_lngCurrTime < gbsngRampSmoothEnd) And gbsngRampSmoothR > 0 Then
            temp = gbsngRampSmoothR ^ 2 - (gbsngRampSmoothEnd - m_lngCurrTime) ^ 2
            If temp > 0 Then
                m_sngSetTemperature = gbsngRampSmoothCy + temp ^ 0.5
            Else
                m_sngSetTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
            End If
        Else
            m_sngSetTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
        End If
        gbsngUniformitySubWeightA = gbsngUniformitySlopeA1 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD1
        gbsngUniformitySubWeightB = gbsngUniformitySlopeA2 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD2
    Else
        m_sngSetTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
    End If
    SaveDebugLog "Hold", 8
           
    If Para.RtaType <> 9 Then
        If gbintRampHoldCount = 1 Then
            SetPIDParameter m_sngSetTemperature, _
                            frmRecipeEdit.sngRecipeProportional, _
                            frmRecipeEdit.sngRecipeIntegral, _
                            frmRecipeEdit.sngRecipeDerivational, _
                            frmRecipeEdit.sngRecipePredit, _
                            frmRecipeEdit.sngRecipeFeedForward
        Else
            If frmRecipeEdit.sngRecipeProportional2 > 0 Then
                SetPIDParameter m_sngSetTemperature, _
                                    frmRecipeEdit.sngRecipeProportional2, _
                                    frmRecipeEdit.sngRecipeIntegral2, _
                                    frmRecipeEdit.sngRecipeDerivational, _
                                    frmRecipeEdit.sngRecipePredit, _
                                    frmRecipeEdit.sngRecipeFeedForward
            Else
                SetPIDParameter m_sngSetTemperature, _
                                    frmRecipeEdit.sngRecipeProportional, _
                                    frmRecipeEdit.sngRecipeIntegral2, _
                                    frmRecipeEdit.sngRecipeDerivational, _
                                    frmRecipeEdit.sngRecipePredit, _
                                    frmRecipeEdit.sngRecipeFeedForward
            End If
        End If
        SaveDebugLog "Hold", 9
        
        CurrProc.sngOutput = PID_Loop(Kernel.sngTC(0), m_lngControlInterval)
        If CurrProc.sngOutput < gbsngIntensityKeep Then CurrProc.sngOutput = gbsngIntensityKeep
        Kernel.sngIntensity = CurrProc.sngOutput
        SaveDebugLog "Hold", 10
        SetAO_SCR CurrProc.sngOutput, gbsngRecipeIntensityWeightSteady
        
    Else
        Dim blnCal As Boolean
        
        CurrProc.sngOutput = Az1.sngMV(0)
        Kernel.sngIntensity = CurrProc.sngOutput / 10
        blnCal = False
        If gbintRampHoldCount >= Para.intMonitorIndex And Para.IsCali = 1 Then
            If Az1.blnUseAzbil = True Then
                For i = 0 To 3
                    If Az1.blnUseLoop(i) = True Then
                        For j = 1 To MultiLoop.intLoopMK(i)
                            If lngGetTime > MultiLoop.lnLoopRTFlag(i, j) And MultiLoop.blnLoopRTActive(i, j) = False Then
                                MultiLoop.blnLoopRTActive(i, j) = True
                                CalAzbilRT (i)
                                blnCal = True
                            End If
                        Next j
                    End If
                Next i
                If blnCal = True Then
                    gbintAz1ProcNo = 6
                End If
            End If
            
            blnCal = False
            If Az2.blnUseAzbil = True Then
                For i = 0 To 3
                    If Az2.blnUseLoop(i) = True Then
                        For j = 1 To MultiLoop.intLoopMK(i + 4)
                            If lngGetTime > MultiLoop.lnLoopRTFlag(i + 4, j) And MultiLoop.blnLoopRTActive(i + 4, j) = False Then
                                MultiLoop.blnLoopRTActive(i + 4, j) = True
                                CalAzbilRT (i + 4)
                                blnCal = True
                            End If
                        Next j
                    End If
                Next i
                If blnCal = True Then
                    gbintAz2ProcNo = 6
                End If
            End If
        End If
    
    End If
    
    
    
    
    SaveDebugLog "Hold", 11
    If CurrProc.sngPump < 0 Then
        If Kernel.sngPressure < Abs(CurrProc.sngPump) And blnResetVI = True Then
            blnResetVI = False
            CalVacuumPID True, 0, 0, 0, 0
            SetTime 0
        End If
        If GetTime(0) > gbsngAPCInterval Then
           SetTime 0
           
           sngTemp(0) = CalVacuumPID(False, gbsngAPC_P, gbsngAPC_I, Abs(CurrProc.sngPump), Kernel.sngPressure)
           frmPlotProcess.fraVacFunc.Caption = "溃北-" & CStr(sngTemp(0))
           sngTemp(0) = Percent2Volt(sngTemp(0), 0, 5)
           Call VarControlMFC(gbintAPC_MFC_Port, sngTemp(0))
        End If
        If frmRecipeEdit.sngRecipeOverPressure > 0 And gbintRampHoldCount > 1 Then
            If Kernel.sngPressure > (Abs(CurrProc.sngPump) + frmRecipeEdit.sngRecipeOverPressure) Then
                ShowAlarmFlash 6
                strTemp = "APC溃北禬絛瞅(" & Format(Kernel.sngPressure, "0.00") & ">" & Format(Abs(CurrProc.sngPump) + frmRecipeEdit.sngRecipeOverPressure, "0.00") & ")"
                Call frmHistory.AppendLogAlert(1, "Alarm", 3074, strTemp, 1)
                Exit Function
            End If
        End If
    End If
        
    If gbintRampHoldCount >= Para.intMonitorIndex Then
        CurrProc.blnCheckUnderCT = True
        If gbsngMaxMonitorError > 0 Then
            lnTemp = lngGetTime
            If (lnTemp - gblngHoldStartTime) > (gbsngMaxMonitorTime * 1000) Then
                temp = Abs(Kernel.sngTC(0) - Kernel.sngTC(1))
                If temp > gbsngMaxMonitorError Then
                    gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
                    ShowAlarmFlash 16
                End If
            End If
        End If
        If frmRecipeEdit.sngRecipeUnderTemp <> 0 Then
            lnTemp = lngGetTime
            If (lnTemp - gblngHoldStartTime) > (gbsngMaxMonitorTime * 1000) Then
                temp = Abs(Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature)
                If temp > frmRecipeEdit.sngRecipeUnderTemp Then
                    gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
                    ShowAlarmFlash 8
                End If
            End If
        End If
        
        intLamps = 35
        If Para.RtaType = 9 Then intLamps = 54
        
        If CurrProc.blnCheckUnderCT = True And gbblnRecipeUseCT = True Then
            lnTemp = lngGetTime
            If (lnTemp - gblngHoldStartTime) > (gbsngMaxMonitorTime * 1000) Then
                For i = 0 To intLamps
                    If Lamp2Scr(CLng(i), intSA) = True Then
                        If Kernel.dblCT(i) < gbsngRecipeCD(intSA) And Kernel.intUnderCT(i) = 0 Then
                                    
                            If CurrProc.lngUnderTime(i) = 0 Then
                                CurrProc.lngUnderTime(i) = lngGetTime
                            
                            ElseIf (lngGetTime - CurrProc.lngUnderTime(i)) > Para.intLampAlarmTime * 1000 Then
                                Kernel.intUnderCT(i) = 1
                                strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngUnderTime(i))
                                gbstrAlarmHint = strTemp
                                ShowAlarmFlash 23
                            End If
                        Else
                            Kernel.intUnderCT(i) = 0
                            CurrProc.lngUnderTime(i) = 0
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    If CurrProc.blnCheckOverCT = True And gbblnRecipeUseCT = True Then
        For i = 0 To intLamps
            If Lamp2Scr(CLng(i), intSA) = True Then
                If Kernel.dblCT(i) > gbsngRecipeCT(intSA) And Kernel.intOverCT(i) = 0 Then
                    If CurrProc.lngOverTime(i) = 0 Then
                        CurrProc.lngOverTime(i) = lngGetTime
                    
                    ElseIf (lngGetTime - CurrProc.lngOverTime(i)) > Para.intLampAlarmTime * 1000 Then
                        Kernel.intOverCT(i) = 1
                        strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngOverTime(i))
                        gbstrAlarmHint = strTemp
                        ShowAlarmFlash 22
                    End If
                Else
                    Kernel.intOverCT(i) = 0
                    CurrProc.lngOverTime(i) = 0
                End If
            End If
        Next
    End If
    
    CheckAlarm
    
    SaveDebugLog "Hold", 12
    intLoopInterval = 1
    If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 3 Then intLoopInterval = 20
    CurrProc.intRecordLoop = CurrProc.intRecordLoop + 1
    If CurrProc.intRecordLoop > intLoopInterval Then
        CurrProc.intRecordLoop = 0
        Call RecordProcessData
    End If
    SaveDebugLog "Hold", 13
    frmPlotProcessLog.SaveCurProcessLog (True)
    Call frmPlotProcess.DrawCurve
    SaveDebugLog "Hold", 14
    Call frmPlotProcess.ShowStatus
    SaveDebugLog "Hold", 15
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Hold error"
    ShowAlarmFlash 1
    
End Function

Public Function Process_RampDown() As Boolean
    
    Dim i As Integer
    Dim lngGetTime As Long
    
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    If lngGetTime <= 0 Then GoTo ERRLINE
    
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    gbsngProcessSlope = 0
    Kernel.sngIntensity = gbsngRecipeRampDownPower
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPDOWN_ID)
        Call Process_Abort
        Exit Function
    End If
    
    Call ControlSCR(gbsngRecipeRampDownPower, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1, _
                    1)
       
    Call RecordProcessData
               
    Call ReadTC
    For i = 0 To 6
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
    Call frmPlotProcess.ShowStatus
    
    
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) And Kernel.IsRun = 1 Then
        
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPDOWN_ID)
        m_lngPreTime = lngGetTime
        gbdblStartRampupProcessFlag = lngGetTime
        CurrProc.intStep = CurrProc.intStep + 1
        CurrProc.blnDoStep = True
'        m_intProcessStep = m_intProcessStep + 1
'        Call RunProcessStep(m_intProcessStep)
        Exit Function
    End If
    
    Call frmPlotProcess.DrawProcessChartData(gbsngRecipeRampDownPower, 0, CLng(gbdblProcessTimeFlag))
    
    Exit Function
ERRLINE:
    gbstrAlarmHint = " RampDown error"
    ShowAlarmFlash 1
    
End Function

Public Function Process_Stop() As Boolean
    On Error Resume Next
    
    frmProcess.tmrProcessStep.Enabled = False
    gbintCurrProcessStep = GB_ACTION_INDEX_STOP
    'Call ControlSCR(0, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)
    'SetAO_SCR 0, gbsngRecipeIntensityWeightDynamic
    Call StopProc(0)
     
End Function

Public Function Process_Abort() As Boolean
    'Call ControlSCR(0, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)
    frmProcess.tmrProcessStep.Enabled = False
    SetAO_SCR 0, gbsngRecipeIntensityWeightDynamic
End Function


Public Sub RecordProcessData()
    Dim i As Integer
    '0: time '1: Power '2: TC1 '3: PM '4: Gas1 '5: Gas2 '6: Gas3 '7: Gas4 '8: Vacuum
    
    
'    gbsngDrawData(0) = Format(CLng(gbdblProcessTimeFlag), "0")
'    For i = 0 To 23
'        gbsngDrawData(i + 2) = Format(Kernel.sngTC(i), "0.00")
'    Next i
'
'    For i = 0 To 5
'        gbsngDrawData(i + 26) = Format(SysAI.sngMFC(i), "0.00")
'    Next i
'
'    For i = 0 To 8
'        gbsngDrawData(i + 26) = Format(SysAI.sngMFC(i), "0.00")
'    Next i
    If gbdblProcessTimeFlag > (gbdblTotalProcessTime + 3) * 1000 Then
    Exit Sub
    End If
    gbsngDrawData(0) = Format(CLng(gbdblProcessTimeFlag), "0")
    gbsngDrawData(1) = Format(Kernel.sngIntensity, "0.00")
    gbsngDrawData(2) = Format(Kernel.sngTC(0), "0.00")
    gbsngDrawData(3) = Format(gbsngPM, "0.00")
    gbsngDrawData(4) = Format(SysAI.sngMFC(0), "0.000")
    gbsngDrawData(5) = Format(SysAI.sngMFC(1), "0.000")
    gbsngDrawData(6) = Format(SysAI.sngMFC(2), "0.000")
    gbsngDrawData(7) = Format(SysAI.sngMFC(3), "0.000")
    gbsngDrawData(8) = Format(SysAI.sngMFC(4), "0.000")
    gbsngDrawData(9) = Format(Kernel.sngPressure, "0.000")
    gbsngDrawData(10) = Format(SysAI.sngMFC(5), "0.000")
    gbsngDrawData(11) = Format(Kernel.sngTC(1), "0.000")
    gbsngDrawData(12) = Format(Kernel.sngTC(2), "0.000")
    gbsngDrawData(13) = Format(Kernel.sngTC(3), "0.000")
    gbsngDrawData(14) = Format(Kernel.sngTC(4), "0.000")
    gbsngDrawData(15) = Format(Kernel.sngTC(5), "0.000")
    gbsngDrawData(16) = Format(gbsngPower(0), "0.000")
    gbsngDrawData(17) = Format(gbsngPower(1), "0.000")
    gbsngDrawData(18) = Format(gbsngPower(2), "0.000")
    gbsngDrawData(19) = Format(gbsngPower(3), "0.000")
    gbsngDrawData(20) = Format(gbsngPower(4), "0.000")
    gbsngDrawData(21) = Format(gbsngOxygenPPM, "0.000")
    gbsngDrawData(22) = Format(Kernel.sngTC(6), "0.000")
    gbsngDrawData(23) = Format(Kernel.sngTC(7), "0.000")
    If GbTestMode_Switch = 1 And gblngAI_Vacuum_Gauge2 >= 0 Then gbsngDrawData(24) = Format(Kernel.sngPressure2, "0.000")
    If Para.UseMTC = 1 Then
        For i = 24 To 31
            gbsngDrawData(i) = Format(Kernel.sngTC(i - 16), "0.000")
        Next i
    End If
    If Para.UseMTCB = 1 Then
        For i = 32 To 39
            gbsngDrawData(i) = Format(Kernel.sngTC(i - 16), "0.000")
        Next i
    End If

    For i = 0 To GB_MAX_DRAW_COL - 1
'        gbsngProcessRecorder(gbProcessRecordCount, i) = gbsngDrawData(i)
         gbsngProcessRecorder(i) = gbsngDrawData(i)
    Next i
    
    
    
    
'    gbProcessRecordCount = gbProcessRecordCount + 1
'    If gbProcessRecordCount > 65000 Then gbProcessRecordCount = 65000
    
End Sub

'Rev4.1.6
Public Sub CalUniformityCompensation()
    gbsngUniformityDistA = gbsngRampSmoothDist + gbsngRampSmoothDistX
    gbsngUniformityDistB1 = gbsngUniformitySubWeight1 - gbsngUniformitySubWeightD1
    gbsngUniformityDistB2 = gbsngUniformitySubWeight2 - gbsngUniformitySubWeightD2
    gbsngUniformitySlopeA1 = 1
    gbsngUniformitySlopeA2 = 1
    If gbsngUniformityDistA <> 0 Then
        gbsngUniformitySlopeA1 = gbsngUniformityDistB1 / gbsngUniformityDistA
        gbsngUniformitySlopeA2 = gbsngUniformityDistB2 / gbsngUniformityDistA
    End If
End Sub


Public Function Process_MultiRampUp() As Boolean
    Dim i As Integer
    Dim intSA As Integer
    Dim bRet As Boolean
    Dim sngTemp(GB_SCR_MAX) As Single
    Dim strTemp     As String
    Dim lngGetTime As Long
    Dim intLoopInterval As Integer
    Dim intLamps As Integer
    
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "RampUp", 1
    
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    If gbblnStartHeatingProcess = False Then
        gbblnStartHeatingProcess = True
        gbdblStartHeatingProcessFlag = lngGetTime
        gbdblStartHeatingProcessFlag = gbdblStartHeatingProcessFlag
    End If
    gbsngProcessSlope = 0
        
    SaveDebugLog "RampUp", 2
    'Waiting the time up, next step
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
        gbdblStartRampupProcessFlag = lngGetTime
        
        CurrProc.intStep = CurrProc.intStep + 1
        CurrProc.blnDoStep = True
'        m_intProcessStep = m_intProcessStep + 1
'        Call RunProcessStep(m_intProcessStep)
        Exit Function
    End If
            
    SaveDebugLog "RampUp", 3
    m_lngCurrTime = lngGetTime
    m_lngControlInterval = m_lngCurrTime - m_lngPreTime
    m_lngPreTime = m_lngCurrTime
    If m_lngControlInterval = 0 Then m_lngControlInterval = 1
    
    If gbsngSmoothTime > 0 Then
        If (m_lngCurrTime > gbsngRampSmoothStart) And gbsngRampSmoothR > 0 Then
            m_sngSetTemperature = gbsngRampSmoothCy + ((gbsngRampSmoothR ^ 2 - (gbsngRampSmoothEnd - m_lngCurrTime) ^ 2) ^ 0.5)
        Else
            If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
                 gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                                     / gbProcessRecipeStep(CurrProc.intStep).sngTime
            Else
                gbsngProcessSlope = 0
            End If
            m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
                                     + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
        End If
        
        gbsngUniformitySubWeightA = gbsngUniformitySlopeA1 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD1
        gbsngUniformitySubWeightB = gbsngUniformitySlopeA2 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD2
    Else
   
        If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
             gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                                 / gbProcessRecipeStep(CurrProc.intStep).sngTime
        Else
            gbsngProcessSlope = 0
        End If
        m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
                                 + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
    End If
    
    SaveDebugLog "RampUp", 4
    
    Call ReadTC
    For i = 0 To 7
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
        
    SaveDebugLog "RampUp", 5
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True Then
            Kernel.sngTC(MultiLoop.intLoopTC(i)) = Kernel.sngTC(MultiLoop.intLoopTC(i)) * MultiLoop.sngLoopRT(i)
            sngTemp(0) = m_sngSetTemperature
            If MultiLoop.sngLoopFT(i) > 40 Then
                sngTemp(0) = MultiLoop.sngLoopFT(i)
            Else
                If MultiLoop.sngLoopFT(i) >= 20 And MultiLoop.sngLoopFT(i) < 40 Then
                    sngTemp(0) = Kernel.sngTC(MultiLoop.sngLoopFT(i) - 20)
                End If
            End If
            
            If sngTemp(0) >= Kernel.sngTC(MultiLoop.intLoopTC(i)) And MultiLoop.blnLoopReset(i) = True Then
                MultiLoop.blnLoopReset(i) = False
                CalMultiPID i, True, 0, 0, 0, 0, 0, 0
            End If
            
            If MultiLoop.blnLoopReset(i) = False Then
                MultiLoop.sngLoopOut(i) = CalMultiPID(i, False, _
                                            MultiLoop.sngLoopPN(i), _
                                            MultiLoop.sngLoopIN(i), _
                                            MultiLoop.sngLoopDN(i), _
                                            sngTemp(0), Kernel.sngTC(MultiLoop.intLoopTC(i)), m_lngControlInterval)
            Else
                MultiLoop.sngLoopOut(i) = gbsngIntensityKeep
            End If
        End If
    Next i
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True And MultiLoop.sngLoopCV(i) > 0 Then
             If Abs(MultiLoop.sngLoopOut(i) - MultiLoop.sngLoopOut(MultiLoop.intLoopCN(i))) > MultiLoop.sngLoopCV(i) Then
                gbstrAlarmHint = ",int" & CStr(i) & "=" & CStr(MultiLoop.sngLoopOut(i)) & ",int" & CStr(MultiLoop.intLoopCN(i)) & "=" & CStr(MultiLoop.sngLoopOut(MultiLoop.intLoopCN(i)))
                ShowAlarmFlash 4
             End If
        End If
    Next i

    SaveDebugLog "RampUp", 7
    
    
    
    SetAO_SCR_Multi MultiLoop.sngLoopOut, MultiLoop.sngWeight
    
    SaveDebugLog "RampUp", 8
'    If m_sngSetTemperature >= Kernel.sngTC(0) Then m_blnStartPIDLoop = True
'    If m_blnStartPIDLoop = True Then
'
'        If gbintRampHoldCount = 1 Then
'            SetPIDParameter m_sngSetTemperature, _
'                                    frmRecipeEdit.sngRecipeProportional, _
'                                    frmRecipeEdit.sngRecipeIntegral, _
'                                    frmRecipeEdit.sngRecipeDerivational, _
'                                    frmRecipeEdit.sngRecipePredit, _
'                                    frmRecipeEdit.sngRecipeFeedForward
'        Else
'            If frmRecipeEdit.sngRecipeProportional2 > 0 Then
'                SetPIDParameter m_sngSetTemperature, _
'                                    frmRecipeEdit.sngRecipeProportional2, _
'                                    frmRecipeEdit.sngRecipeIntegral2, _
'                                    frmRecipeEdit.sngRecipeDerivational, _
'                                    frmRecipeEdit.sngRecipePredit, _
'                                    frmRecipeEdit.sngRecipeFeedForward
'            Else
'                SetPIDParameter m_sngSetTemperature, _
'                                    frmRecipeEdit.sngRecipeProportional, _
'                                    frmRecipeEdit.sngRecipeIntegral2, _
'                                    frmRecipeEdit.sngRecipeDerivational, _
'                                    frmRecipeEdit.sngRecipePredit, _
'                                    frmRecipeEdit.sngRecipeFeedForward
'            End If
'        End If
'        CurrProc.sngOutput = PID_Loop(Kernel.sngTC(0), m_lngControlInterval) '* sngOutputScale    'out max 200 degree/sec
'    Else
'        CurrProc.sngOutput = gbsngIntensityKeep
'    End If
'    SaveDebugLog "RampUp", 7
'
'    If CurrProc.sngOutput < gbsngIntensityKeep Then CurrProc.sngOutput = gbsngIntensityKeep
'    Kernel.sngIntensity = CurrProc.sngOutput
'    SetAO_SCR CurrProc.sngOutput, gbsngRecipeIntensityWeightDynamic
'    SaveDebugLog "RampUp", 8
    
    
    
    If CurrProc.sngPump < 0 Then
        If Kernel.sngPressure < Abs(CurrProc.sngPump) And blnResetVI = True Then
            blnResetVI = False
            CalVacuumPID True, 0, 0, 0, 0
            SetTime 0
        End If
        If GetTime(0) > gbsngAPCInterval Then
           SetTime 0
           
           sngTemp(0) = CalVacuumPID(False, gbsngAPC_P, gbsngAPC_I, Abs(CurrProc.sngPump), Kernel.sngPressure)
           frmPlotProcess.fraVacFunc.Caption = "溃北-" & CStr(sngTemp(0))
           sngTemp(0) = Percent2Volt(sngTemp(0), 0, 5)
           Call VarControlMFC(gbintAPC_MFC_Port, sngTemp(0))
 
        End If
    End If
    
    intLamps = 35
    If Para.RtaType = 9 Then intLamps = 54
    
    If CurrProc.blnCheckOverCT = True And gbblnRecipeUseCT = True Then
        For i = 0 To intLamps
            If Lamp2Scr(CLng(i), intSA) = True Then
                If Kernel.dblCT(i) > gbsngRecipeCT(intSA) And Kernel.intOverCT(i) = 0 Then
                    If CurrProc.lngOverTime(i) = 0 Then
                        CurrProc.lngOverTime(i) = lngGetTime
                    
                    ElseIf (lngGetTime - CurrProc.lngOverTime(i)) > Para.intLampAlarmTime * 1000 Then
                        strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngOverTime(i))
                        gbstrAlarmHint = strTemp
                        Kernel.intOverCT(i) = 1
                        ShowAlarmFlash 22
                    End If
                Else
                    Kernel.intOverCT(i) = 0
                    CurrProc.lngOverTime(i) = 0
                End If
            End If
        Next
    End If
    
    intLoopInterval = 1
    If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 3 Then intLoopInterval = 20
    CurrProc.intRecordLoop = CurrProc.intRecordLoop + 1
    If CurrProc.intRecordLoop > intLoopInterval Then
        CurrProc.intRecordLoop = 0
        Call RecordProcessData
    End If
    
    
    SaveDebugLog "RampUp", 9
    frmPlotProcessLog.SaveCurProcessLog (True)
    Call frmPlotProcess.DrawCurve
    SaveDebugLog "RampUp", 10
    Call frmPlotProcess.ShowStatus
    SaveDebugLog "RampUp", 11
    
    If CurrProc.intStep > 0 And gbProcessRecipeStep(CurrProc.intStep).sngTemperature > gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature Then
        If (Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) Then
            'Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
            'Call Process_Abort
            gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
            ShowAlarmFlash 9
        End If
    End If
    SaveDebugLog "RampUp", 12
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Ramp up Error"
    ShowAlarmFlash 1
    
End Function

Public Function Process_MultiHold() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim sngMiddleError As Single
    Dim sngEdgeError(5) As Single
    Dim bRet As Boolean
    Dim temp As Single
    Dim strTemp As String
    Dim lnTemp As Long
    Dim sngTemp(3) As Single
    Dim sngPump As Single
    Dim intSA As Integer
    Dim lngGetTime As Long
    Dim intLoopInterval As Integer
    Dim intLamps As Integer
    
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    If lngGetTime <= 0 Then GoTo ERRLINE
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "Hold", 1
    'Waiting the time up, next step
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) Then
        If CurrProc.sngPump > 0 And CurrProc.sngPump <> 100 Or (CurrProc.sngPump = 100 And gbintRampHoldCount = 1 And gbsngPumpDownGate > 0) Then
            sngPump = CurrProc.sngPump
            If CurrProc.sngPump = 100 Then
                sngPump = gbsngPumpDownGate / 1000
            End If
            If Kernel.sngPressure > sngPump And Kernel.sngPressure > 0.001 Then
                If gbdblProcessPumpDownTimerout > 0 Then
                    If gbblnPumpDownTimeout = False Then
                        gbdblStartPumpDownProcessFlag = lngGetTime
                        strTemp = "Pressure Check (" & Format(Kernel.sngPressure, "0.000") & ")Torr"
                        Call frmHistory.AppendLogAlert(1, "Check", 3053, strTemp, 1)
                        strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        gbblnPumpDownTimeout = True
                        
                    Else
                        If (lngGetTime - gbdblStartPumpDownProcessFlag) > gbdblProcessPumpDownTimerout Then
                               
                            gbblnPumpDownTimeout = False
                            strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                            gbstrAlarmHint = strTemp
                            ShowAlarmFlash 6
                            
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            Call frmHistory.AppendLogAlert(1, "Alarm", 3054, strTemp, 1)
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    gbstrAlarmHint = strTemp
                    strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000")
                    ShowAlarmFlash 6
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    Call frmHistory.AppendLogAlert(1, "Alarm", 3054, strTemp, 1)
                    Exit Function
                End If
            Else
                strTemp = "Pressure Check (" & Format(Kernel.sngPressure, "0.000") & ")Torr"
                Call frmHistory.AppendLogAlert(1, "Check", 3053, strTemp, 1)
                If gbblnPumpDownTimeout = True Then
                    gbblnPumpDownTimeout = False
                    gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - gbdblStartPumpDownProcessFlag)
                End If
                gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                strTemp = "Va=" & Format(sngPump, "0.000") & "/" & Format(Kernel.sngPressure, "0.000") & "/" & Format(gblngPumpDownTime, "0")
                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                
                Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
                gbdblStartRampupProcessFlag = lngGetTime
                CurrProc.intStep = CurrProc.intStep + 1
                CurrProc.blnDoStep = True

                Exit Function
                
                
            End If
        Else
            If Para.sngO2Gate > 0 Then
                If gbsngRecipePrepareTimeout > 0 Then
                    If Kernel.sngOxygen > Para.sngO2Gate Then
                        If CurrProc.blnOxygenTimeout = False Then
                            CurrProc.dblStartO2Flag = lngGetTime
                            strTemp = "O2 Check (" & Format(Kernel.sngOxygen, "0.00") & ")ppm"
                            Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                            strTemp = "O2=" & Format(Para.sngO2Gate, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                            mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                            CurrProc.blnOxygenTimeout = True
                        Else
                            If (lngGetTime - CurrProc.dblStartO2Flag) > gbsngRecipePrepareTimeout Then
                                   
                                CurrProc.blnOxygenTimeout = False
                                strTemp = "O2=" & Format(Para.sngO2Gate, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                                gbstrAlarmHint = strTemp
                                ShowAlarmFlash 24
                                
                                mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                                Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                            Else
                                Exit Function
                            End If
                        End If
                    Else
                        strTemp = "O2 Check (" & Format(Kernel.sngOxygen, "0.00") & ")ppm"
                        Call frmHistory.AppendLogAlert(1, "Check", 4053, strTemp, 1)
                        If CurrProc.blnOxygenTimeout = True Then
                            CurrProc.blnOxygenTimeout = False
                            gbdblProcessKeepTime = gbdblProcessKeepTime + (lngGetTime - CurrProc.dblStartO2Flag)
                        End If
                        gblngPumpDownTime = CLng(gbdblProcessKeepTime) / 1000
                        strTemp = "O2=" & Format(Para.sngO2Gate, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00") & "/" & Format(gblngPumpDownTime, "0")
                        mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                        
                        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
                        gbdblStartRampupProcessFlag = lngGetTime
                        CurrProc.intStep = CurrProc.intStep + 1
                        CurrProc.blnDoStep = True
        
                        Exit Function
                    End If
                Else
                    gbstrAlarmHint = strTemp
                    strTemp = "O2=" & Format(Para.sngO2Gate, "0.00") & "/" & Format(Kernel.sngOxygen, "0.00")
                    ShowAlarmFlash 24
                    mdifrmRTP.stabarRTP.Panels(7).text = strTemp
                    Call frmHistory.AppendLogAlert(1, "Alarm", 4054, strTemp, 1)
                    Exit Function
                End If
            Else
            
            
            
                Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
                gbdblStartRampupProcessFlag = lngGetTime
                CurrProc.intStep = CurrProc.intStep + 1
                CurrProc.blnDoStep = True
'                If HoldCount = TempOffset(5) Then GbHoldState = False
                Exit Function
            End If
        End If
    End If
    
    SaveDebugLog "Hold", 2
    
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    
    If gbblnStartHeatingProcess = False Then
        gbblnStartHeatingProcess = True
        gbdblStartHeatingProcessFlag = lngGetTime
    End If
    SaveDebugLog "Hold", 3
    gbsngProcessSlope = 0
    m_lngCurrTime = lngGetTime
    m_lngControlInterval = m_lngCurrTime - m_lngPreTime
    m_lngPreTime = m_lngCurrTime
    If m_lngControlInterval = 0 Then m_lngControlInterval = 1
    SaveDebugLog "Hold", 4
    If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
        gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                            / gbProcessRecipeStep(CurrProc.intStep).sngTime
    Else
        gbsngProcessSlope = 0
    End If
    SaveDebugLog "Hold", 5
    Call ReadTC
    For i = 0 To 7
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
    SaveDebugLog "Hold", 6
    If (Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID)
        Call Process_Abort
        gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
        ShowAlarmFlash 9
        Exit Function
    End If
    SaveDebugLog "Hold", 7
    If gbsngSmoothTime > 0 Then
        If (m_lngCurrTime < gbsngRampSmoothEnd) And gbsngRampSmoothR > 0 Then
            temp = gbsngRampSmoothR ^ 2 - (gbsngRampSmoothEnd - m_lngCurrTime) ^ 2
            If temp > 0 Then
                m_sngSetTemperature = gbsngRampSmoothCy + temp ^ 0.5
            Else
                m_sngSetTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
            End If
        Else
            m_sngSetTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
        End If
        gbsngUniformitySubWeightA = gbsngUniformitySlopeA1 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD1
        gbsngUniformitySubWeightB = gbsngUniformitySlopeA2 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD2
    Else
        m_sngSetTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
    End If
    SaveDebugLog "Hold", 8
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True Then
            Kernel.sngTC(MultiLoop.intLoopTC(i)) = Kernel.sngTC(MultiLoop.intLoopTC(i)) * MultiLoop.sngLoopRT(i)
            sngTemp(0) = m_sngSetTemperature
            If MultiLoop.sngLoopFT(i) > 40 Then
                sngTemp(0) = MultiLoop.sngLoopFT(i)
            Else
                If MultiLoop.sngLoopFT(i) >= 20 And MultiLoop.sngLoopFT(i) < 40 Then
                    sngTemp(0) = Kernel.sngTC(MultiLoop.sngLoopFT(i) - 20)
                End If
            End If
            
            If sngTemp(0) >= Kernel.sngTC(MultiLoop.intLoopTC(i)) And MultiLoop.blnLoopReset(i) = True Then
                MultiLoop.blnLoopReset(i) = False
                CalMultiPID i, True, 0, 0, 0, 0, 0, 0
            End If
            
            If MultiLoop.blnLoopReset(i) = False Then
                MultiLoop.sngLoopOut(i) = CalMultiPID(i, False, _
                                            MultiLoop.sngLoopPN(i), _
                                            MultiLoop.sngLoopIN(i), _
                                            MultiLoop.sngLoopDN(i), _
                                            sngTemp(0), Kernel.sngTC(MultiLoop.intLoopTC(i)), m_lngControlInterval)
            Else
                MultiLoop.sngLoopOut(i) = gbsngIntensityKeep
            End If
            
            If gbintRampHoldCount >= Para.intMonitorIndex And Para.IsCali = 1 Then
                For j = 1 To MultiLoop.intLoopMK(i)
                    If lngGetTime > MultiLoop.lnLoopRTFlag(i, j) And MultiLoop.blnLoopRTActive(i, j) = False Then
                        MultiLoop.blnLoopRTActive(i, j) = True
                        CalMultiRT (i)
                        'Call frmPlotProcess.lbIntensityLabel_DblClick(i)
                        
                    End If
                Next j
            End If
        End If
    Next i
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True And MultiLoop.sngLoopCV(i) > 0 Then
             If Abs(MultiLoop.sngLoopOut(i) - MultiLoop.sngLoopOut(MultiLoop.intLoopCN(i))) > MultiLoop.sngLoopCV(i) Then
                gbstrAlarmHint = ",int" & CStr(i) & "=" & CStr(MultiLoop.sngLoopOut(i)) & ",int" & CStr(MultiLoop.intLoopCN(i)) & "=" & CStr(MultiLoop.sngLoopOut(MultiLoop.intLoopCN(i)))
                ShowAlarmFlash 4
             End If
        End If
    Next i
    
    SaveDebugLog "Hold", 9
    
    SetAO_SCR_Multi MultiLoop.sngLoopOut, MultiLoop.sngWeight
            
'    If gbintRampHoldCount = 1 Then
'        SetPIDParameter m_sngSetTemperature, _
'                        frmRecipeEdit.sngRecipeProportional, _
'                        frmRecipeEdit.sngRecipeIntegral, _
'                        frmRecipeEdit.sngRecipeDerivational, _
'                        frmRecipeEdit.sngRecipePredit, _
'                        frmRecipeEdit.sngRecipeFeedForward
'    Else
'        If frmRecipeEdit.sngRecipeProportional2 > 0 Then
'            SetPIDParameter m_sngSetTemperature, _
'                                frmRecipeEdit.sngRecipeProportional2, _
'                                frmRecipeEdit.sngRecipeIntegral2, _
'                                frmRecipeEdit.sngRecipeDerivational, _
'                                frmRecipeEdit.sngRecipePredit, _
'                                frmRecipeEdit.sngRecipeFeedForward
'        Else
'            SetPIDParameter m_sngSetTemperature, _
'                                frmRecipeEdit.sngRecipeProportional, _
'                                frmRecipeEdit.sngRecipeIntegral2, _
'                                frmRecipeEdit.sngRecipeDerivational, _
'                                frmRecipeEdit.sngRecipePredit, _
'                                frmRecipeEdit.sngRecipeFeedForward
'        End If
'    End If
'    SaveDebugLog "Hold", 9
'    CurrProc.sngOutput = PID_Loop(Kernel.sngTC(0), m_lngControlInterval)
'    If CurrProc.sngOutput < gbsngIntensityKeep Then CurrProc.sngOutput = gbsngIntensityKeep
'    Kernel.sngIntensity = CurrProc.sngOutput
'    SaveDebugLog "Hold", 10
'    SetAO_SCR CurrProc.sngOutput, gbsngRecipeIntensityWeightSteady
    
    SaveDebugLog "Hold", 11
    If CurrProc.sngPump < 0 Then
        If Kernel.sngPressure < Abs(CurrProc.sngPump) And blnResetVI = True Then
            blnResetVI = False
            CalVacuumPID True, 0, 0, 0, 0
            SetTime 0
        End If
        If GetTime(0) > gbsngAPCInterval Then
           SetTime 0
           
           sngTemp(0) = CalVacuumPID(False, gbsngAPC_P, gbsngAPC_I, Abs(CurrProc.sngPump), Kernel.sngPressure)
           frmPlotProcess.fraVacFunc.Caption = "溃北-" & CStr(sngTemp(0))
           sngTemp(0) = Percent2Volt(sngTemp(0), 0, 5)
           Call VarControlMFC(gbintAPC_MFC_Port, sngTemp(0))
        End If
        If frmRecipeEdit.sngRecipeOverPressure > 0 And gbintRampHoldCount > 1 Then
            If Kernel.sngPressure > (Abs(CurrProc.sngPump) + frmRecipeEdit.sngRecipeOverPressure) Then
                ShowAlarmFlash 6
                strTemp = "APC溃北禬絛瞅(" & Format(Kernel.sngPressure, "0.00") & ">" & Format(Abs(CurrProc.sngPump) + frmRecipeEdit.sngRecipeOverPressure, "0.00") & ")"
                Call frmHistory.AppendLogAlert(1, "Alarm", 3074, strTemp, 1)
                Exit Function
            End If
        End If
    End If
        
    If gbintRampHoldCount >= Para.intMonitorIndex Then
        CurrProc.blnCheckUnderCT = True
        If gbsngMaxMonitorError > 0 Then
            lnTemp = lngGetTime
            If (lnTemp - gblngHoldStartTime) > (gbsngMaxMonitorTime * 1000) Then
                temp = Abs(Kernel.sngTC(0) - Kernel.sngTC(1))
                If temp > gbsngMaxMonitorError Then
                    gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
                    ShowAlarmFlash 16
                End If
            End If
        End If
        If frmRecipeEdit.sngRecipeUnderTemp <> 0 Then
            lnTemp = lngGetTime
            If (lnTemp - gblngHoldStartTime) > (gbsngMaxMonitorTime * 1000) Then
                temp = Abs(Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature)
                If temp > frmRecipeEdit.sngRecipeUnderTemp Then
                    gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
                    ShowAlarmFlash 8
                End If
            End If
        End If
        
        intLamps = 35
        If Para.RtaType = 9 Then intLamps = 54
        
        If CurrProc.blnCheckUnderCT = True And gbblnRecipeUseCT = True Then
            lnTemp = lngGetTime
            If (lnTemp - gblngHoldStartTime) > (gbsngMaxMonitorTime * 1000) Then
                For i = 0 To intLamps
                    If Lamp2Scr(CLng(i), intSA) = True Then
                        If Kernel.dblCT(i) < gbsngRecipeCD(intSA) And Kernel.intUnderCT(i) = 0 Then
                                    
                            If CurrProc.lngUnderTime(i) = 0 Then
                                CurrProc.lngUnderTime(i) = lngGetTime
                            
                            ElseIf (lngGetTime - CurrProc.lngUnderTime(i)) > Para.intLampAlarmTime * 1000 Then
                                Kernel.intUnderCT(i) = 1
                                strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngUnderTime(i))
                                gbstrAlarmHint = strTemp
                                ShowAlarmFlash 23
                            End If
                        Else
                            Kernel.intUnderCT(i) = 0
                            CurrProc.lngUnderTime(i) = 0
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    If CurrProc.blnCheckOverCT = True And gbblnRecipeUseCT = True Then
        For i = 0 To intLamps
            If Lamp2Scr(CLng(i), intSA) = True Then
                If Kernel.dblCT(i) > gbsngRecipeCT(intSA) And Kernel.intOverCT(i) = 0 Then
                    If CurrProc.lngOverTime(i) = 0 Then
                        CurrProc.lngOverTime(i) = lngGetTime
                    
                    ElseIf (lngGetTime - CurrProc.lngOverTime(i)) > Para.intLampAlarmTime * 1000 Then
                        Kernel.intOverCT(i) = 1
                        strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngOverTime(i))
                        gbstrAlarmHint = strTemp
                        ShowAlarmFlash 22
                    End If
                Else
                    Kernel.intOverCT(i) = 0
                    CurrProc.lngOverTime(i) = 0
                End If
            End If
        Next
    End If
    
    
    SaveDebugLog "Hold", 12
    intLoopInterval = 1
    If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 3 Then intLoopInterval = 20
    CurrProc.intRecordLoop = CurrProc.intRecordLoop + 1
    If CurrProc.intRecordLoop > intLoopInterval Then
        CurrProc.intRecordLoop = 0
        Call RecordProcessData
    End If
    SaveDebugLog "Hold", 13
    frmPlotProcessLog.SaveCurProcessLog (True)
    Call frmPlotProcess.DrawCurve
    SaveDebugLog "Hold", 14
    Call frmPlotProcess.ShowStatus
    SaveDebugLog "Hold", 15
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Hold error"
    ShowAlarmFlash 1
    
End Function

Public Function Process_MultiRampDown() As Boolean
    Dim i As Integer
    Dim intSA As Integer
    Dim bRet As Boolean
    Dim sngTemp(GB_SCR_MAX) As Single
    Dim strTemp     As String
    Dim lngGetTime As Long
    Dim intLoopInterval As Integer
    Dim intLamps As Integer
    
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "RampDown", 1
    
    gbdblProcessTimeFlag = lngGetTime - gbdblStartProcessFlag - gbdblProcessKeepTime
    If gbblnStartHeatingProcess = False Then
        gbblnStartHeatingProcess = True
        gbdblStartHeatingProcessFlag = lngGetTime
        gbdblStartHeatingProcessFlag = gbdblStartHeatingProcessFlag
    End If
    gbsngProcessSlope = 0
        
    SaveDebugLog "RampDown", 2
    'Waiting the time up, next step
    If lngGetTime > gbdblProcessTimeStamp(CurrProc.intStep) Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID)
        gbdblStartRampupProcessFlag = lngGetTime
        
        CurrProc.intStep = CurrProc.intStep + 1
        CurrProc.blnDoStep = True
'        m_intProcessStep = m_intProcessStep + 1
'        Call RunProcessStep(m_intProcessStep)
        Exit Function
    End If
            
    SaveDebugLog "RampDown", 3
    m_lngCurrTime = lngGetTime
    m_lngControlInterval = m_lngCurrTime - m_lngPreTime
    m_lngPreTime = m_lngCurrTime
    If m_lngControlInterval = 0 Then m_lngControlInterval = 1
    
    If gbsngSmoothTime > 0 Then
        If (m_lngCurrTime > gbsngRampSmoothStart) And gbsngRampSmoothR > 0 Then
            m_sngSetTemperature = gbsngRampSmoothCy + ((gbsngRampSmoothR ^ 2 - (gbsngRampSmoothEnd - m_lngCurrTime) ^ 2) ^ 0.5)
        Else
            If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
                 gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                                     / gbProcessRecipeStep(CurrProc.intStep).sngTime
            Else
                gbsngProcessSlope = 0
            End If
            m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
                                     + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
        End If

        gbsngUniformitySubWeightA = gbsngUniformitySlopeA1 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD1
        gbsngUniformitySubWeightB = gbsngUniformitySlopeA2 * (m_lngCurrTime - gbsngRampSmoothStart) + gbsngUniformitySubWeightD2
    Else

        If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
             gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
                                                 / gbProcessRecipeStep(CurrProc.intStep).sngTime
        Else
            gbsngProcessSlope = 0
        End If
        m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
                                 + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
    End If
    
'    If gbProcessRecipeStep(CurrProc.intStep).sngTime <> 0 Then
'         gbsngProcessSlope = (gbProcessRecipeStep(CurrProc.intStep).sngTemperature - gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature) _
'                                             / gbProcessRecipeStep(CurrProc.intStep).sngTime
'    Else
'        gbsngProcessSlope = 0
'    End If
'    m_sngSetTemperature = ((m_lngCurrTime - gbdblStartRampupProcessFlag) / 1000 * gbsngProcessSlope) _
'                             + gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature
    
    SaveDebugLog "RampDown", 4
    
    Call ReadTC
    For i = 0 To 7
        sngCurrTemp(i) = Kernel.sngTC(i)
    Next i
        
    SaveDebugLog "RampDown", 5
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True Then
            Kernel.sngTC(MultiLoop.intLoopTC(i)) = Kernel.sngTC(MultiLoop.intLoopTC(i)) * MultiLoop.sngLoopRT(i)
            
            sngTemp(0) = m_sngSetTemperature
            
            If MultiLoop.sngLoopFT(i) > 20 Then
                sngTemp(0) = MultiLoop.sngLoopFT(i)
            Else
                If MultiLoop.sngLoopFT(i) >= 0 And MultiLoop.sngLoopFT(i) < 20 Then
                    sngTemp(0) = Kernel.sngTC(MultiLoop.sngLoopFT(i))
                ElseIf MultiLoop.sngLoopFT(i) >= 20 And MultiLoop.sngLoopFT(i) < 40 Then
                    sngTemp(0) = Kernel.sngTC(MultiLoop.sngLoopFT(i) - 20)
                End If
            End If
            
            If sngTemp(0) >= Kernel.sngTC(MultiLoop.intLoopTC(i)) And MultiLoop.blnLoopReset(i) = True Then
                MultiLoop.blnLoopReset(i) = False
                CalMultiPID i, True, 0, 0, 0, 0, 0, 0
            End If
            
            If MultiLoop.blnLoopReset(i) = False Then
                MultiLoop.sngLoopOut(i) = CalMultiPID(i, False, _
                                            MultiLoop.sngLoopPN(i), _
                                            MultiLoop.sngLoopIN(i), _
                                            MultiLoop.sngLoopDN(i), _
                                            sngTemp(0), Kernel.sngTC(MultiLoop.intLoopTC(i)), m_lngControlInterval)
            Else
                MultiLoop.sngLoopOut(i) = gbsngIntensityKeep
            End If
        End If
    Next i
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True And MultiLoop.sngLoopCV(i) > 0 Then
             If Abs(MultiLoop.sngLoopOut(i) - MultiLoop.sngLoopOut(MultiLoop.intLoopCN(i))) > MultiLoop.sngLoopCV(i) Then
                gbstrAlarmHint = ",int" & CStr(i) & "=" & CStr(MultiLoop.sngLoopOut(i)) & ",int" & CStr(MultiLoop.intLoopCN(i)) & "=" & CStr(MultiLoop.sngLoopOut(MultiLoop.intLoopCN(i)))
                ShowAlarmFlash 4
             End If
        End If
    Next i

    SaveDebugLog "RampDown", 7
    
    
    
    SetAO_SCR_Multi MultiLoop.sngLoopOut, MultiLoop.sngWeight
    
    SaveDebugLog "RampDown", 8
'    If m_sngSetTemperature >= Kernel.sngTC(0) Then m_blnStartPIDLoop = True
'    If m_blnStartPIDLoop = True Then
'
'        If gbintRampHoldCount = 1 Then
'            SetPIDParameter m_sngSetTemperature, _
'                                    frmRecipeEdit.sngRecipeProportional, _
'                                    frmRecipeEdit.sngRecipeIntegral, _
'                                    frmRecipeEdit.sngRecipeDerivational, _
'                                    frmRecipeEdit.sngRecipePredit, _
'                                    frmRecipeEdit.sngRecipeFeedForward
'        Else
'            If frmRecipeEdit.sngRecipeProportional2 > 0 Then
'                SetPIDParameter m_sngSetTemperature, _
'                                    frmRecipeEdit.sngRecipeProportional2, _
'                                    frmRecipeEdit.sngRecipeIntegral2, _
'                                    frmRecipeEdit.sngRecipeDerivational, _
'                                    frmRecipeEdit.sngRecipePredit, _
'                                    frmRecipeEdit.sngRecipeFeedForward
'            Else
'                SetPIDParameter m_sngSetTemperature, _
'                                    frmRecipeEdit.sngRecipeProportional, _
'                                    frmRecipeEdit.sngRecipeIntegral2, _
'                                    frmRecipeEdit.sngRecipeDerivational, _
'                                    frmRecipeEdit.sngRecipePredit, _
'                                    frmRecipeEdit.sngRecipeFeedForward
'            End If
'        End If
'        CurrProc.sngOutput = PID_Loop(Kernel.sngTC(0), m_lngControlInterval) '* sngOutputScale    'out max 200 degree/sec
'    Else
'        CurrProc.sngOutput = gbsngIntensityKeep
'    End If
'    SaveDebugLog "RampDown", 7
'
'    If CurrProc.sngOutput < gbsngIntensityKeep Then CurrProc.sngOutput = gbsngIntensityKeep
'    Kernel.sngIntensity = CurrProc.sngOutput
'    SetAO_SCR CurrProc.sngOutput, gbsngRecipeIntensityWeightDynamic
'    SaveDebugLog "RampDown", 8
    
    
    
    If CurrProc.sngPump < 0 Then
        If Kernel.sngPressure < Abs(CurrProc.sngPump) And blnResetVI = True Then
            blnResetVI = False
            CalVacuumPID True, 0, 0, 0, 0
            SetTime 0
        End If
        If GetTime(0) > gbsngAPCInterval Then
           SetTime 0
           
           sngTemp(0) = CalVacuumPID(False, gbsngAPC_P, gbsngAPC_I, Abs(CurrProc.sngPump), Kernel.sngPressure)
           frmPlotProcess.fraVacFunc.Caption = "溃北-" & CStr(sngTemp(0))
           sngTemp(0) = Percent2Volt(sngTemp(0), 0, 5)
           Call VarControlMFC(gbintAPC_MFC_Port, sngTemp(0))
 
        End If
    End If
    
    intLamps = 35
    If Para.RtaType = 9 Then intLamps = 54
    
    If CurrProc.blnCheckOverCT = True And gbblnRecipeUseCT = True Then
        For i = 0 To intLamps
            If Lamp2Scr(CLng(i), intSA) = True Then
                If Kernel.dblCT(i) > gbsngRecipeCT(intSA) And Kernel.intOverCT(i) = 0 Then
                    If CurrProc.lngOverTime(i) = 0 Then
                        CurrProc.lngOverTime(i) = lngGetTime
                    
                    ElseIf (lngGetTime - CurrProc.lngOverTime(i)) > Para.intLampAlarmTime * 1000 Then
                        strTemp = "(No)=" & i & ",A=" & Format(Kernel.dblCT(i), "0.0") & ",Sec=" & CStr(lngGetTime - CurrProc.lngOverTime(i))
                        gbstrAlarmHint = strTemp
                        Kernel.intOverCT(i) = 1
                        ShowAlarmFlash 22
                    End If
                Else
                    Kernel.intOverCT(i) = 0
                    CurrProc.lngOverTime(i) = 0
                End If
            End If
        Next
    End If
    
    intLoopInterval = 1
    If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 3 Then intLoopInterval = 20
    CurrProc.intRecordLoop = CurrProc.intRecordLoop + 1
    If CurrProc.intRecordLoop > intLoopInterval Then
        CurrProc.intRecordLoop = 0
        Call RecordProcessData
    End If
    
    
    SaveDebugLog "RampDown", 9
     frmPlotProcessLog.SaveCurProcessLog (True)
    Call frmPlotProcess.DrawCurve
    SaveDebugLog "RampDown", 10
    Call frmPlotProcess.ShowStatus
    SaveDebugLog "RampDown", 11
    
    If CurrProc.intStep > 0 And gbProcessRecipeStep(CurrProc.intStep).sngTemperature > gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature Then
        If (Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) Then
            gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
            ShowAlarmFlash 9
        End If
    End If
    SaveDebugLog "RampDown", 12
    Exit Function
ERRLINE:
    gbstrAlarmHint = " RampDown Error"
    ShowAlarmFlash 1
    
End Function

Public Function CheckAlarm()
Dim i As Integer

    If Para.RtaType = 9 Then
        If Az1.blnUseAzbil = True Then
            For i = 0 To 3
                If Az1.blnUseLoop(i) = True Then
                    If Kernel.sngTC(i) < gbsngMinTemperature Then
                        gbstrAlarmHint = "TC(" & CStr(i) & ")= " & Format(Kernel.sngTC(i), "0.0")
                        ShowAlarmFlash 8
                    End If
                    If Kernel.sngTC(i) > gbsngMaxTemperature Then
                        gbstrAlarmHint = "TC(" & CStr(i) & ")= " & Format(Kernel.sngTC(i), "0.0")
                        ShowAlarmFlash 9
                    End If
                    If gbProcessRecipeStep(CurrProc.intStep).strAction = GB_ACTION_RAMPUP Or gbProcessRecipeStep(CurrProc.intStep).strAction = GB_ACTION_HOLD Then
                        If (Kernel.sngTC(i) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) Then
                            gbstrAlarmHint = ",TC(" & CStr(i) & ")=" & Format(Kernel.sngTC(i), "0.0")
                            ShowAlarmFlash 9
                        End If
                    End If
                    If MultiLoop.sngLoopCV(i) > 0 Then
                        If MultiLoop.intLoopCN(i) < 4 Then
                            If Abs(Az1.sngMV(i) - Az1.sngMV(MultiLoop.intLoopCN(i))) > MultiLoop.sngLoopCV(i) Then
                               gbstrAlarmHint = ",int" & CStr(i) & "=" & CStr(Az1.sngMV(i)) & ",int" & CStr(MultiLoop.intLoopCN(i)) & "=" & CStr(Az1.sngMV(MultiLoop.intLoopCN(i)))
                               ShowAlarmFlash 4
                            End If
                        Else
                            If Abs(Az1.sngMV(i) - Az2.sngMV(MultiLoop.intLoopCN(i) - 4)) > MultiLoop.sngLoopCV(i) Then
                               gbstrAlarmHint = ",int" & CStr(i) & "=" & CStr(Az1.sngMV(i)) & ",int" & CStr(MultiLoop.intLoopCN(i)) & "=" & CStr(Az2.sngMV(MultiLoop.intLoopCN(i) - 4))
                               ShowAlarmFlash 4
                            End If
                        End If
                    End If
                    
                End If
            Next i
        End If
        If Az2.blnUseAzbil = True Then
            For i = 0 To 3
'                i = 0
                If Az2.blnUseLoop(i) = True Then
                    If Kernel.sngTC(i + 4) < gbsngMinTemperature And InStr(1, gbstrNameTC(i + 4), "TC") > 0 Then
                        gbstrAlarmHint = "TC(" & CStr(i + 4) & ")= " & Format(Kernel.sngTC(i + 4), "0.0")
                        ShowAlarmFlash 8
                    End If
                    If Kernel.sngTC(i + 4) > gbsngMaxTemperature And InStr(1, gbstrNameTC(i + 4), "TC") > 0 Then
                        gbstrAlarmHint = "TC(" & CStr(i + 4) & ")= " & Format(Kernel.sngTC(i + 4), "0.0")
                        ShowAlarmFlash 9
                    End If
                    If gbProcessRecipeStep(CurrProc.intStep).strAction = GB_ACTION_RAMPUP Or gbProcessRecipeStep(CurrProc.intStep).strAction = GB_ACTION_HOLD Then
                        If (Kernel.sngTC(i + 4) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) And InStr(1, gbstrNameTC(i + 4), "TC") > 0 Then
                            gbstrAlarmHint = ",TC(" & CStr(i + 4) & ")=" & Format(Kernel.sngTC(i + 4), "0.0")
                            ShowAlarmFlash 9
                        End If
                    End If
                    
                    If MultiLoop.sngLoopCV(i + 4) > 0 Then
                        If MultiLoop.intLoopCN(i) < 4 Then
                            If Abs(Az2.sngMV(i) - Az1.sngMV(MultiLoop.intLoopCN(i))) > MultiLoop.sngLoopCV(i + 4) Then
                               gbstrAlarmHint = ",int" & CStr(i + 4) & "=" & CStr(Az2.sngMV(i)) & ",int" & CStr(MultiLoop.intLoopCN(i + 4)) & "=" & CStr(Az1.sngMV(MultiLoop.intLoopCN(i)))
                               ShowAlarmFlash 4
                            End If
                        Else
                            If Abs(Az2.sngMV(i) - Az2.sngMV(MultiLoop.intLoopCN(i) - 4)) > MultiLoop.sngLoopCV(i + 4) Then
                               gbstrAlarmHint = ",int" & CStr(i + 4) & "=" & CStr(Az2.sngMV(i)) & ",int" & CStr(MultiLoop.intLoopCN(i + 4)) & "=" & CStr(Az2.sngMV(MultiLoop.intLoopCN(i) - 4))
                               ShowAlarmFlash 4
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        
        If CurrProc.sngOutput > frmRecipeEdit.sngRecipeIntLimit Then
            gbstrAlarmHint = " Recipe_LIMIT=" & CStr(CurrProc.sngOutput)
            ShowAlarmFlash 4
        End If
            
    Else
        If gbProcessRecipeStep(CurrProc.intStep).strAction = GB_ACTION_RAMPUP Or gbProcessRecipeStep(CurrProc.intStep).strAction = GB_ACTION_HOLD Then
            If (Kernel.sngTC(0) - gbProcessRecipeStep(CurrProc.intStep).sngTemperature) > CSng(frmRecipeEdit.sngRecipeOverTemp) Then
                gbstrAlarmHint = ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0")
                ShowAlarmFlash 9
            End If
        End If
    End If
    
    If gbsngRecipeGatePS2 > 0 Then
        For i = 0 To 23
            If gbstrNameTC(i) = "PS" Then
                If Kernel.sngTC(i) < 1000 And Kernel.sngTC(i) >= gbsngRecipeGatePS1 And Kernel.sngTC(i) > gbsngRecipeGatePS2 Then
                    gbstrAlarmHint = "PS=" & Format(Kernel.sngTC(i), "0.0") & ">" & Format(gbsngRecipeGatePS2, "0.0")
                    ShowAlarmFlash 7
                End If
            End If
        Next i
    End If
End Function

Public Function Process_IO() As Boolean
    Dim i As Integer
    Dim bRet As Boolean
    Dim sngTemp(10) As Single
    Dim strTemp As String
    Dim lngGetTime As Long
    
    On Error GoTo ERRLINE
        
    lngGetTime = timeGetTime
    CurrProc.lngPrevTime = CurrProc.lngCurrentTime
    CurrProc.lngCurrentTime = lngGetTime
    Kernel.lngCurrStepCount = (gbdblProcessTimeStamp(CurrProc.intStep) - lngGetTime) / 1000
    
    If Kernel.IsRun = 0 Then
        Call KillTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID)
        Call Process_Abort
        Exit Function
    End If
    SaveDebugLog "IO Control", 1
    
    If Para.UseCover = 1 And gbProcessRecipeStep(CurrProc.intStep).sngTemperature = 1 Then
        If gbProcessRecipeStep(CurrProc.intStep).sngTime = 1 Then
            SetCover True
            Call frmHistory.AppendLogAlert(1, "Process", 1051, "Cover Down", 1)
        Else
            SetCover False
            Call frmHistory.AppendLogAlert(1, "Process", 1052, "Cover Up", 1)
        End If
    End If
    
    CurrProc.intStep = CurrProc.intStep + 1
    CurrProc.blnDoStep = True
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Process_IO error"
    ShowAlarmFlash 1
        
End Function
    
