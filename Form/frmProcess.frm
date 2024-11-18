VERSION 5.00
Begin VB.Form frmProcess 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrOpenDoor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   360
   End
   Begin VB.Timer tmrProcessStep 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   360
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrOpenDoor_Timer()
    If Kernel.intOpenDoorCount <> 9999 Then
        Kernel.intOpenDoorCount = Kernel.intOpenDoorCount + 1
        If Kernel.intOpenDoorCount > Para.intOpenDoorTime Then
            tmrOpenDoor.Enabled = False
            Kernel.intOpenDoorCount = 9999
            ShowMessageOK "腔體開門時間超出設定值!"
            Call frmHistory.AppendLogAlert(1, "Alarm", 4055, "腔體開門時間超出設定值!", 1)
        End If
    End If
    
End Sub

Private Sub ReWritesngGas()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Index As New Collection
Dim NewSngGas As New Collection
For i = 0 To 5
If gbstrGasAlias(i) = "NA" Then
Index.Add i
End If
Next i
For j = 0 To UBound(CurrProc.sngGas) - Index.Count
If IsExist(j, Index) = False Then
NewSngGas.Add CurrProc.sngGas(j)
Else
NewSngGas.Add 0
NewSngGas.Add CurrProc.sngGas(j)
End If
Next j
For k = 1 To NewSngGas.Count
CurrProc.sngGas(k - 1) = NewSngGas(k)
Next k
End Sub


Private Function IsExist(Index As Integer, IndexList As Collection) As Boolean
Dim result As Boolean
Dim i As Integer
result = False
For i = 1 To IndexList.Count
If Index = IndexList(i) Then
result = True
Exit For
End If
Next i
IsExist = result
End Function


Private Sub tmrProcessStep_Timer()
    Dim i As Integer
    Dim j As Integer
    Dim tn As Long
    Dim sngRampSmoothResult(5) As Single
    Dim StrFileName     As String
    Dim data(4) As Integer
    Dim SysProName As String

    
    
    If Kernel.IsRun = 0 Then Exit Sub
    StrFileName = gbSystemPath & "\System\system.cfg"
    If CurrProc.blnDoStep Then
        CurrProc.blnDoStep = False
        CurrProc.strAction = gbProcessRecipeStep(CurrProc.intStep).strAction
        SysProName = Readini(CurrProc.strAction)
        If SysProName <> "" Then
        CurrProc.strAction = SysProName
        End If
        If CurrProc.strAction <> GB_ACTION_STOP Then
             
            CurrProc.sngTime = gbProcessRecipeStep(CurrProc.intStep).sngTime
            CurrProc.sngTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
            CurrProc.sngPump = 0
            SaveDebugLog "Start", 1
            
'            For i = 0 To 5
             For i = 1 To frmRecipeEdit.GasNames.Count
                CurrProc.sngGas(i - 1) = gbProcessRecipeStep(CurrProc.intStep).sngGas(i - 1)
                
'                If gbstrGasAlias(i) = "Pump" Then
                   If frmRecipeEdit.GasNames(i) = "Pump" Then
                    If CurrProc.sngGas(i - 1) > 0 Or CurrProc.sngGas(i - 1) < 0 Then
                        CurrProc.sngPump = CurrProc.sngGas(i - 1)
                        
                        If SetAngle(True) > 0 Then
                            If Para.useTPump = 0 Then
                                If SysDI.IsChamberGaugeL = 1 Then frmDiagnosis.tmrPumpON.Enabled = True
                            End If
                        End If
                    Else
                        SetAngle False
                        frmDiagnosis.tmrPumpON.Enabled = False
                        If Para.useTPump = 0 Then
                            If Para.intPumpDelay > 0 Then
                                If Para.intPumpDelay < 30 Then
                                    frmDiagnosis.tmrPumpOFF.Interval = Para.intPumpDelay * 1000
                                    frmDiagnosis.tmrPumpOFF.Enabled = True
                                End If
                            Else
                                SetPump False
                                If gbintReleaseOpenDelay > 0 Then
                                    frmDiagnosis.tmrSetReleaseOFF.Interval = gbintReleaseOpenDelay
                                    frmDiagnosis.tmrSetReleaseOFF.Enabled = True
                                End If
                            End If
                        End If
                        'SetPump False
                        SaveDebugLog "Start", 4
                    End If
                    CurrProc.sngGas(i - 1) = 0
                End If
            Next i
            ReWritesngGas
            SetAO_MFC CurrProc.sngGas
            SaveDebugLog "Start", 5
            
            SetDO gblngDO_ARM_FRONT, False
            If Para.IsHoldSafety = 1 Then
                frmDiagnosis.tmrHoldSafeON.Enabled = True
            End If
        End If
                
        
        CurrProc.blnCheckOverCT = True
        CurrProc.blnCheckUnderCT = False
        CurrProc.lngCurrentTime = timeGetTime
        CurrProc.lngPrevTime = CurrProc.lngCurrentTime
        CurrProc.sngOutput = 0
        For i = 0 To 35
            CurrProc.lngOverTime(i) = 0
            CurrProc.lngUnderTime(i) = 0
        Next i
        
        CurrProc.intRecordLoop = 0
        Kernel.lngCurrStepCount = 0
           
        If MultiLoop.blnUseMultiLoop = True Then
            For i = 0 To GB_MAX_LOOPS - 1
                MultiLoop.sngLoopOut(i) = 0
            Next i
        End If
        
        gbblnAutoCloseValve = gbblnRecipeAutoCloseValve1
           
        SaveDebugLog "Start", 6
        Select Case CurrProc.strAction
            Case GB_ACTION_IDLE
                CurrProc.lngScanTime = 1000 'mSec
                CurrProc.intAction = GB_ACTION_INDEX_IDLE
                gbdblProcessTimeStamp(CurrProc.intStep) = CurrProc.lngCurrentTime + gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000
                CurrProc.lngCurrStepTime = gbProcessRecipeStep(CurrProc.intStep).sngTime
                'SetAO_SCR 0, gbsngRecipeIntensityWeightDynamic
                CurrProc.sngOutput = 0
                Kernel.sngIntensity = 0
                If gbProcessRecipeStep(CurrProc.intStep).sngTemperature > 0 And gbProcessRecipeStep(CurrProc.intStep).sngTemperature < 80 Then
                    CurrProc.sngOutput = gbProcessRecipeStep(CurrProc.intStep).sngTemperature / 10
                    Kernel.sngIntensity = CurrProc.sngOutput
                    If Para.RtaType > 4 Then CurrProc.sngOutput = gbProcessRecipeStep(CurrProc.intStep).sngTemperature / 20 '0~5v
                End If
                
'                If Az1.blnUseAzbil Then
'                    gbintAz1ProcNo = Int(gbProcessRecipeStep(CurrProc.intStep).sngTemperature)
'                End If
'                If Az2.blnUseAzbil Then
'
'                    gbintAz2ProcNo = Int(gbProcessRecipeStep(CurrProc.intStep).sngTemperature)
'                End If
                
                gbblnActivePrepare = False
                If gbsngRecipePrepareGaugeO2 > 0 Then
                    If gbProcessRecipeStep(CurrProc.intStep + 1).strAction = GB_ACTION_RAMPUP Then
                        gbblnActivePrepare = True
                    End If
                End If
                
                gbblnActiveTempDown = False
                If gbProcessRecipeStep(CurrProc.intStep).sngTemperature > 0 And gbsngRecipeTempDownTimeout > 0 Then
                    gbblnActiveTempDown = True
                End If
                
                SetAO_SCR CurrProc.sngOutput, gbsngRecipeIntensityWeightDynamic
                SaveDebugLog "Start", 7
                Call frmHistory.AppendLogAlert(1, "Process", 1001, "Idle process", 1)
                SaveDebugLog "Start", 8
                Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID, CurrProc.lngScanTime, AddressOf Process_Idle)
                SaveDebugLog "Start", 9
                Kernel.strCurrStep = "(" & CStr(CurrProc.intStep) & ")-IDLE"
                Kernel.intCurrStep = CurrProc.intStep
                                
            Case GB_ACTION_RAMPUP
                CurrProc.lngScanTime = 15 'mSec
                If Para.RtaType = 9 Then CurrProc.lngScanTime = 200 'mSec
                CurrProc.intAction = GB_ACTION_INDEX_RAMPUP
                gbintRampHoldCount = gbintRampHoldCount + 1
                m_blnStartPIDLoop = False
                If gbblnResetInteral = True Then Call InitPIDParameter
                SaveDebugLog "Start", 10
                gbdblProcessTimeStamp(CurrProc.intStep) = CurrProc.lngCurrentTime + gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000
                CurrProc.lngCurrStepTime = gbProcessRecipeStep(CurrProc.intStep).sngTime
                CurrProc.sngTemperature = gbProcessRecipeStep(CurrProc.intStep).sngTemperature
                
                '130109 Josh
                gbdblLampLifeTime = gbdblLampLifeTime + gbProcessRecipeStep(CurrProc.intStep).sngTime
                SaveDebugLog "Start", 11
                If gbProcessRecipeStep(CurrProc.intStep + 1).strAction = GB_ACTION_HOLD And gbsngSmoothTime > 0 Then
                    gbsngRampSmoothDist = gbsngSmoothTime * 1000
                    Call CalSmoothCurve(gbsngRampSmoothDist, CSng(CurrProc.lngCurrentTime), gbProcessRecipeStep(CurrProc.intStep - 1).sngTemperature, CSng(gbdblProcessTimeStamp(CurrProc.intStep)), gbProcessRecipeStep(CurrProc.intStep).sngTemperature, sngRampSmoothResult)
                    'Rev4.1.6
                    Call CalUniformityCompensation
                End If
                SaveDebugLog "Start", 12
                Call frmHistory.AppendLogAlert(1, "Process", 1002, "RampUp process", 1)
                SaveDebugLog "Start", 13
                
'                If Az1.blnUseAzbil Then
'                    Az1.intTemp1 = CurrProc.sngTemperature
'                    Az1.intTime1 = CurrProc.lngCurrStepTime
'                    Az1.intTemp2 = gbProcessRecipeStep(CurrProc.intStep + 1).sngTemperature
'                    Az1.intTime2 = gbProcessRecipeStep(CurrProc.intStep + 1).sngTime
'                    gbintAz1ProcNo = 1
'                End If
'                If Az2.blnUseAzbil Then
'                    Az2.intTemp1 = CurrProc.sngTemperature
'                    Az2.intTime1 = CurrProc.lngCurrStepTime
'                    Az2.intTemp2 = gbProcessRecipeStep(CurrProc.intStep + 1).sngTemperature
'                    Az2.intTime2 = gbProcessRecipeStep(CurrProc.intStep + 1).sngTime
'                    gbintAz2ProcNo = 1
'                End If
                If Para.RtaType = 9 And IsUsedSCR = 1 Then
                    DoEvents
                    frmModBusRtu.WriteRamupSCR
                End If
                If MultiLoop.blnUseMultiLoop = True Then
                    If gbintRampHoldCount = 1 Then
                        For i = 0 To GB_MAX_LOOPS - 1
                            MultiLoop.blnLoopReset(i) = True
                        Next i
                    End If
                    For i = 0 To GB_SCR_MAX - 1
                        MultiLoop.sngWeight(i) = gbsngRecipeIntensityWeightDynamic(i)
                    Next i
                    Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID, CurrProc.lngScanTime, AddressOf Process_MultiRampUp)
                Else
                    bExecuted1 = False
                    Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID, CurrProc.lngScanTime, AddressOf Process_RampUp)
                End If
                
                SaveDebugLog "Start", 14
                Kernel.strCurrStep = "(" & CStr(CurrProc.intStep) & ")-Ramp Up"
                Kernel.intCurrStep = CurrProc.intStep
                
            Case GB_ACTION_HOLD 'Hold
                CurrProc.lngScanTime = 15 'mSec
                CurrProc.intAction = GB_ACTION_INDEX_HOLD
                m_blnStartPIDLoop = False
                gbdblLampLifeTime = gbdblLampLifeTime + gbProcessRecipeStep(CurrProc.intStep).sngTime
                gbProcessRecipeStep(CurrProc.intStep).sngTime = gbProcessRecipeStep(CurrProc.intStep).sngTime
                CurrProc.lngCurrStepTime = gbProcessRecipeStep(CurrProc.intStep).sngTime
                gbdblProcessTimeStamp(CurrProc.intStep) = CurrProc.lngCurrentTime + gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000
                SaveDebugLog "Start", 16
                gblngHoldStartTime = timeGetTime
                SaveDebugLog "Start", 17
                Call frmHistory.AppendLogAlert(1, "Process", 1003, "Hold process", 1)
                SaveDebugLog "Start", 18
                If Para.RtaType = 9 And IsUsedSCR = 1 Then
                    DoEvents
                    frmModBusRtu.WriteHoldSCR
                End If
                If MultiLoop.blnUseMultiLoop = True Then
                    For i = 0 To GB_MAX_LOOPS - 1
                        If gbintRampHoldCount >= Para.intMonitorIndex And MultiLoop.intLoopMK(i) > 0 And Para.IsCali = 1 Then
                            'tn = gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000 / MultiLoop.intLoopMK(i) - 1000
                            tn = gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000 / (MultiLoop.intLoopMK(i) + 1)
                            For j = 1 To MultiLoop.intLoopMK(i)
                                MultiLoop.lnLoopRTFlag(i, j) = CurrProc.lngCurrentTime + tn * j
                                MultiLoop.blnLoopRTActive(i, j) = False
                            Next j
                        End If
                    Next i
                    For i = 0 To GB_SCR_MAX - 1
                        MultiLoop.sngWeight(i) = gbsngRecipeIntensityWeightSteady(i)
                    Next i
                                        
                    Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID, CurrProc.lngScanTime, AddressOf Process_MultiHold)
                Else
                    If Az1.blnUseAzbil = True Then
                        For i = 0 To 3
                            If gbintRampHoldCount >= Para.intMonitorIndex And MultiLoop.intLoopMK(i) > 0 And Para.IsCali = 1 Then
                                'tn = gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000 / MultiLoop.intLoopMK(i) - 1000
                                tn = gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000 / (MultiLoop.intLoopMK(i) + 1)
                                For j = 1 To MultiLoop.intLoopMK(i)
                                    MultiLoop.lnLoopRTFlag(i, j) = CurrProc.lngCurrentTime + tn * j
                                    MultiLoop.blnLoopRTActive(i, j) = False
                                Next j
                            End If
                        Next i
                    End If
                    If Az2.blnUseAzbil = True Then
                        For i = 4 To 7
                            If gbintRampHoldCount >= Para.intMonitorIndex And MultiLoop.intLoopMK(i) > 0 And Para.IsCali = 1 Then
                                'tn = gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000 / MultiLoop.intLoopMK(i) - 1000
                                tn = gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000 / (MultiLoop.intLoopMK(i) + 1)
                                For j = 1 To MultiLoop.intLoopMK(i)
                                    MultiLoop.lnLoopRTFlag(i, j) = CurrProc.lngCurrentTime + tn * j
                                    MultiLoop.blnLoopRTActive(i, j) = False
                                Next j
                            End If
                        Next i
                    End If
                    bExecuted2 = False
                    If GbTcoffset_Switch = 1 Then
                        HoldCount = HoldCount + 1
                        If HoldCount = TempOffset(5) Then GbHoldState = True
                    End If
                    Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_HOLD_ID, CurrProc.lngScanTime, AddressOf Process_Hold)
                End If
                SaveDebugLog "Start", 19
                Kernel.strCurrStep = "(" & CStr(CurrProc.intStep) & ")-Hold"
                Kernel.intCurrStep = CurrProc.intStep
                
            Case GB_ACTION_STOP 'Stop
                tmrProcessStep.Enabled = False
                gbintCurrProcessStep = GB_ACTION_INDEX_STOP
                gbsngUsedLamp = gbsngUsedLamp + gbdblLampLifeTime
                Call WritePrivateProfileString("PARAMETER", "UsedLamp", CStr(gbsngUsedLamp), StrFileName)
                gbblnAutoCloseValve = gbblnRecipeAutoCloseValve1
                Kernel.strCurrStep = "(" & CStr(CurrProc.intStep) & ")-Stop"
                Kernel.intCurrStep = CurrProc.intStep
                SaveDebugLog "Start", 20
                Call frmHistory.AppendLogAlert(1, "Process", 1004, "Stop process", 1)
                SaveDebugLog "Start", 21
                Call Process_Stop
                SaveDebugLog "Start", 22
                
            Case GB_ACTION_RAMPDOWN 'Ramp Down
                CurrProc.lngScanTime = 15 'mSec
                CurrProc.intAction = GB_ACTION_INDEX_RAMPDOWN
                gbintRampHoldCount = gbintRampHoldCount + 1
                m_blnStartPIDLoop = False
                If gbblnResetInteral = True Then Call InitPIDParameter
                SaveDebugLog "Start", 10
                gbdblProcessTimeStamp(CurrProc.intStep) = CurrProc.lngCurrentTime + gbProcessRecipeStep(CurrProc.intStep).sngTime * 1000
                CurrProc.lngCurrStepTime = gbProcessRecipeStep(CurrProc.intStep).sngTime
                '130109 Josh
                gbdblLampLifeTime = gbdblLampLifeTime + gbProcessRecipeStep(CurrProc.intStep).sngTime
                SaveDebugLog "Start", 11
                Call frmHistory.AppendLogAlert(1, "Process", 1005, "RampDown process", 1)
                If MultiLoop.blnUseMultiLoop = True Then
                    'If gbintRampHoldCount = 1 Then
                        For i = 0 To GB_MAX_LOOPS - 1
                            MultiLoop.blnLoopReset(i) = True
                        Next i
                    'End If
                    For i = 0 To GB_SCR_MAX - 1
                        MultiLoop.sngWeight(i) = gbsngRecipeIntensityWeightDynamic(i)
                    Next i
                    Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID, CurrProc.lngScanTime, AddressOf Process_MultiRampDown)
                Else
                    Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_RAMPUP_ID, CurrProc.lngScanTime, AddressOf Process_RampUp)
                End If
                SaveDebugLog "Start", 14
                Kernel.strCurrStep = "(" & CStr(CurrProc.intStep) & ")-Ramp Down"
                Kernel.intCurrStep = CurrProc.intStep
                
            Case GB_ACTION_IOCONTROL 'IO Control
                Call SetTimer(mdifrmRTP.hwnd, TIMER_EVENT_IDLE_ID, CurrProc.lngScanTime, AddressOf Process_IO)
                Kernel.strCurrStep = "(" & CStr(CurrProc.intStep) & ")-IO"
                Kernel.intCurrStep = CurrProc.intStep
                
                
        End Select
    End If
    
End Sub





