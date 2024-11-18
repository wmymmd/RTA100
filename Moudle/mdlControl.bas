Attribute VB_Name = "mdlControl"
Option Explicit
Const VACUUM_TORR_CNT_CT = 6.304
'Const VACUUM_TORR_CNT_K = 1.286
Const VACUUM_TORR_CNT_K = 1.282

'<<<<<<<<<<<< Advantech PCI-1710-HGU DAQ >>>>>>>>>>>>
Public advThermo As clsAdvThermo
Public PumpState As Boolean
Public BreakVacuum As Boolean

'<<<<<<<<<<<< ADLINK 9114-HG DAQ >>>>>>>>>>>>
'DI x 16
'DO x 16
'AI x 32

'========================================================================================================
'<<<<<<<<<<<< ADLINK 6216V Analog Output >>>>>>>>>>>>
'AO x 16
'DI x 4 (TTL)
'DO x 4 (TTL)
'========================================================================================================


'Define the Name of IO
Public Const DAQ_9114HG1 = 1
Public Const DAQ_9114HG2 = 2
Public Const DAQ_6208V1 = 0
'Public Const DAQ_6208V2 = 2
'========================================================================================================
'<<<<<<<<<<<< ICPDAQ PCI-822LU MultiFunction >>>>>>>>>>>>
'DI x 16 (TTL)
'DO x 16 (TTL)
'AI x 32 (singal end) / AI x16 (Diff)

'========================================================================================================
'<<<<<<<<<<<< ICPDAQ PIO-DA8 Analog Output >>>>>>>>>>>>
'AO x 16
'DI x 4 (TTL)
'DO x 4 (TTL)
'========================================================================================================
Public gbintTotalBoard As Integer

Public gbstrModelName As String * 20

Public gbstDevInfo(0 To MAX_BOARD_NUMBER - 1) As IXUD_DEVICE_INFO
Public gbstCardInfo(0 To MAX_BOARD_NUMBER - 1) As IXUD_CARD_INFO
Public gbstSelDevInfo As IXUD_DEVICE_INFO
Public gbstSelCardInfo As IXUD_CARD_INFO
Public gbintPCI822A As Integer
Public gbintPIODA8A As Integer


'Declare port no of DI
Public gblngDI_SystemAlarm          As Long
Public gblngDI_SystemReady          As Long
'Public lngDI_PowerInput220         As Long
Public gblngDI_EMS_Alarm            As Long
Public gblngDI_MC_Input             As Long
Public gblngDI_ChamberOverheat      As Long
Public gblngDI_AirAlarm             As Long
Public gblngDI_WaterAlarm           As Long
Public gblngDI_VAcuumGaugeSwitch1   As Long
Public gblngDI_VAcuumGaugeSwitch2   As Long
Public gblngDI_DoorOpen             As Long
Public gblngDI_DoorClose            As Long
Public gblngDI_DoorOpenSensor       As Long
Public gblngDI_DoorCloseSensor      As Long
Public gblngDI_DoorClamp            As Long
Public gblngDI_VacuumGateR          As Long
Public gblngDI_VacuumGateL          As Long
Public gblngDI_ChamberGateR         As Long
Public gblngDI_ChamberGateL         As Long
Public gblngDI_CT1                  As Long
Public gblngDI_CT2                  As Long
Public gblngDI_CT3                  As Long
Public gblngDI_CT4                  As Long
Public gblngDI_CT5                  As Long
Public gblngDI_CT6                  As Long
Public gblngDI_ARM_FRONT            As Long
Public gblngDI_ARM_REAR             As Long
Public gblngDI_CoverAlarm1           As Long
Public gblngDI_CoverServoRdy1        As Long
Public gblngDI_CoverOrigRdy1         As Long
Public gblngDI_CoverIsMoving1         As Long
Public gblngDI_CoverUpInpos1         As Long
Public gblngDI_CoverDownInpos1       As Long
Public gblngDI_CoverAlarm2           As Long
Public gblngDI_CoverServoRdy2        As Long
Public gblngDI_CoverOrigRdy2         As Long
Public gblngDI_CoverIsMoving2         As Long
Public gblngDI_CoverUpInpos2         As Long
Public gblngDI_CoverDownInpos2       As Long
Public gblngDI_PumpAlarm             As Long

'========================================================================================================

'Declare port status of DI
Dim blnDI_STA_PowerInput As Boolean
Dim blnDI_STA_MC_Input As Boolean
Dim blnDI_STA_EMS_Alarm As Boolean
Dim blnDI_STA_ChamberOverheat   As Boolean
Dim blnDI_STA_AirAlarm          As Boolean
Dim blnDI_STA_WaterAlarm As Boolean
Dim blnDI_STA_VacuumAlarm As Boolean
Dim blnDI_STA_DoorInterlock As Boolean
Dim blnDI_STA_DoorOpen As Boolean
Dim blnDI_STA_DoorClose As Boolean

'========================================================================================================

'Declare port no of DO
'System
'Public lngDO_WaterAirStop                   As Long
Public gblngDO_PC_Check1                     As Long
Public gblngDO_PC_Check2                      As Long
Public gblngDO_SystemAlarm      As Long
Public gblngDO_BuzzerStop       As Long
'Gas
Public gblngDO_MFC_ValveServo1        As Long
Public gblngDO_MFC_ValveServo2        As Long
Public gblngDO_MFC_ValveServo3        As Long
Public gblngDO_MFC_ValveServo4        As Long
Public gblngDO_MFC_ValveServo5        As Long
Public gblngDO_MFC_ValveServo6        As Long
Public gblngDO_GasValve1        As Long
Public gblngDO_GasValve2        As Long
Public gblngDO_GasValve3        As Long
Public gblngDO_GasValve4        As Long
Public gblngDO_GasValve5        As Long
Public gblngDO_GasValve6        As Long
'CDA
Public gblngDO_ValveCDA         As Long
'Exhaust
Public gblngDO_Exhaust          As Long
'Vacuum
Public gblngDO_PumpPower        As Long
Public gblngDO_AngleValve       As Long
Public gblngDO_ReleaseValve     As Long
Public gblngDO_APCGaugeValve    As Long
Public gblngDO_APCGaugeAngle    As Long
'Public lngDO_PumpingSlow      As Long
'Public lngDO_PumpingFast      As Long
Public lngDO_DoorOpenValve As Long
Public lngDO_DoorCloseValve As Long
Public lngDO_DoorClamp As Long
Public gblngDO_AlarmRed As Long
Public gblngDO_AlarmYellow As Long
Public gblngDO_AlarmBlue As Long
Public gblngDO_AlarmGreen As Long
Public gblngDO_ARM_FRONT            As Long
Public gblngDO_ARM_REAR            As Long
Public gblngDO_COVER_ARESET            As Long
Public gblngDO_COVER_SERVO            As Long
Public gblngDO_COVER_ORGIN            As Long
Public gblngDO_COVER_MOVE            As Long
Public gblngDO_COVER_POS_01            As Long
'========================================================================================================

'Declare port status of DI
Dim blnDO_STA_WaterAirStop      As Boolean
Dim blnDO_STA_PC_Check1      As Boolean
Dim blnDO_STA_PC_Check2      As Boolean
Dim blnDO_STA_MFC_ValveServo1      As Boolean
Dim blnDO_STA_MFC_ValveServo2      As Boolean
Dim blnDO_STA_MFC_ValveServo3      As Boolean
Dim blnDO_STA_MFC_ValveServo4      As Boolean
Dim blnDO_STA_MFC_ValveServo5      As Boolean
Dim blnDO_STA_MFC_ValveServo6      As Boolean
Dim blnDO_STA_GasInletSlow      As Boolean
Dim blnDO_STA_DoorOpen      As Boolean
Dim blnDO_STA_DoorClose      As Boolean
'========================================================================================================

'Declare port no of AI
Public gblngAI_CT_01          As Long
Public gblngAI_CT_02          As Long
Public gblngAI_CT_03          As Long
Public gblngAI_CT_04          As Long
Public gblngAI_CT_05          As Long
Public gblngAI_CT_06          As Long
Public gblngAI_CT_07          As Long
Public gblngAI_MFC_Read(GB_GAS_MAX)     As Long
'Public gblngAI_MFC_Read1     As Long
'Public gblngAI_MFC_Read2     As Long
'Public gblngAI_MFC_Read3     As Long
'Public gblngAI_MFC_Read4     As Long
Public gblngAI_Vacuum_Gauge  As Long
Public gblngAI_Vacuum_Gauge2  As Long
Public gblngAI_Vacuum_Gauge_APC  As Long
Public gblngAI_TC_Cvt1       As Long
Public gblngAI_TC_Cvt2       As Long
Public gblngAI_Pyrometer     As Long
'Rev4.1.4 Add the TC wafer read port address
Public gblngAI_TCWafer1     As Long
Public gblngAI_TCWafer2     As Long
Public gblngAI_TCWafer3     As Long
Public gblngAI_TCWafer4     As Long
Public gblngAI_TCWafer5     As Long
Public gblngAI_TCWafer6     As Long
Public gblngAI_TCWafer7     As Long
Public gblngAI_TCWafer8     As Long
Public gblngAI_TCWafer9     As Long

Public gblngAI_Oxygen_Gauge  As Long
'========================================================================================================
'Declare port no of AI of advantech pci-1710hgu
Public gblngAI_TC_01          As Long
Public gblngAI_TC_02          As Long
Public gblngAI_TC_03          As Long
Public gblngAI_TC_04          As Long
Public gblngAI_TC_05          As Long
Public gblngAI_TC_06          As Long
Public gblngAI_TC_07          As Long
Public gblngAI_TC_08          As Long
Public gblngAI_TC_09          As Long
Public gblngAI_TC_10          As Long
Public gblngAI_TC_11          As Long
Public gblngAI_TC_12          As Long
Public gblngAI_TC_13          As Long
Public gblngAI_TC_14          As Long
Public gblngAI_TC_15          As Long
Public gblngAI_TC_16          As Long
Public gblngAI_TC_17          As Long
Public gblngAI_TC_18          As Long
Public gblngAI_TC_19          As Long
Public gblngAI_TC_20          As Long
Public gblngAI_TC_21          As Long
Public gblngAI_TC_22          As Long
Public gblngAI_TC_23          As Long
Public gblngAI_TC_24          As Long
'========================================================================================================

'Declare port no of AO
Public gblngAO_SCR_TBC    As Long
Public gblngAO_SCR_TR    As Long
Public gblngAO_SCR_TL    As Long
Public gblngAO_SCR_BF    As Long
Public gblngAO_SCR_BR    As Long
'120713 Josh
Public gblngAO_SCR_6    As Long
Public gblngAO_SCR_7    As Long
Public gblngAO_SCR_8    As Long
Public gblngAO_SCR_9    As Long
Public gblngAO_SCR_10    As Long
Public gblngAO_SCR_11    As Long
Public gblngAO_SCR_12    As Long
Public gblngAO_SCR_13    As Long
Public gblngAO_SCR_14    As Long
Public gblngAO_SCR_15    As Long
Public gblngAO_SCR_16    As Long
Public gblngAO_SCR_17    As Long
Public gblngAO_MFC1     As Long
Public gblngAO_MFC2     As Long
Public gblngAO_MFC3     As Long
Public gblngAO_MFC4     As Long
Public gblngAO_MFC5     As Long
Public gblngAO_MFC6     As Long
Public gbintAO_Type(32)     As Integer

Public gbintDO_Status(32)     As Integer
Public gbsngCurrMFC_Output(GB_GAS_MAX)     As Single


Private Type ADVDAQ1710IO
    lngDI(16) As Long
    lngDO(16) As Long
    lngDOValue As Long
    sngAI(16) As Single
    sngAO(8) As Single
End Type
'========================================================================================================
'Define the DAQ 9114 IO
Private Type DAQ9114IO
    lngDI(16) As Long
    lngDO(16) As Long
    lngDOValue As Long
    sngAI(16) As Single
End Type

Public AvDaq17101 As ADVDAQ1710IO
Public DAQ91141 As DAQ9114IO


Public Const TIMEREVT_ANGLEVALVE_OPEN = 1001
Public Const TIMEREVT_THROTTLEALVE_OPEN = 1002
Public Const TIMEREVT_RELEASEVALVE_CLOSE = 1003
Public Const TIMEREVT_MFCVALVE_OPEN = 1011

Public gbbnlPC_STOP As Boolean
Public sngSumAI(32) As Single
Public intCountAI(32) As Integer
Public gbintPMcmdID As Integer
Public sngTemp_1(10) As Single
Public sngTemp_2(10) As Single
Public intErrorCount_1 As Integer
Public intErrorCount_2 As Integer


Public Sub ReadDI()
    If Para.RtaType = 1 Or Para.RtaType = 3 Or Para.RtaType = 7 Then ReadDI_HG
    If Para.RtaType = 2 Then ReadDI_AD
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 8 Or Para.RtaType = 9 Then ReadDI_FR
End Sub

Public Sub ReadAI()
    If Para.RtaType = 1 Or Para.RtaType = 3 Or Para.RtaType = 7 Then ReadAI_HG
    If Para.RtaType = 2 Then ReadAI_AD
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 8 Or Para.RtaType = 9 Then ReadAI_FR
End Sub

Public Sub ReadTC()

    If Para.RtaType = 9 Then
        If Para.UseAz1 Then ReadAz1
        If Para.UseAz2 Then ReadAz2
        
        If Para.UseMTC = 1 And Kernel.IsActiveTC1 = 1 Then ReadTC_HGU_A
        If Para.UseMTCB = 1 And Kernel.IsActiveTC2 = 1 Then ReadTC_HGU_B
        
    Else
        If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 6 Then ReadTC_AD
        If Para.RtaType = 3 Or Para.RtaType = 5 Then ReadTC_HGU_1
        If Para.UseMTC = 1 And Kernel.IsActiveTC2 = 1 Then ReadTC_HGU_2
        If Para.RtaType = 7 Or Para.RtaType = 8 Then ReadTC_HGT_1
        
    End If
        
    Call CheckTC
End Sub

Public Sub ReadDI_FR()
    Dim i As Integer
    Dim wDiVal As Integer
    
On Error GoTo ERRLINE:
    
    If gbblnActiveHGF = True Then
        Call FRB_ReceiveRA(0, 8, wDiVal)
        
        For i = 0 To 15
            SysDI.value(i) = IIf(wDiVal And 2 ^ i, 1, 0)
        Next i
            
        If gblngDI_MC_Input >= 0 Then
            If SysDI.value(gblngDI_MC_Input) = 1 Then
                SysDI.IsReady = 0
            Else
                SysDI.IsReady = 1
            End If
        End If
        If gblngDI_ChamberGateL >= 0 Then SysDI.IsChamberGaugeL = SysDI.value(gblngDI_ChamberGateL)
        If gblngDI_ChamberGateR >= 0 Then SysDI.IsChamberGaugeH = SysDI.value(gblngDI_ChamberGateR)
        If gblngDI_EMS_Alarm >= 0 Then SysDI.IsEMO = SysDI.value(gblngDI_EMS_Alarm)
        If gblngDI_DoorCloseSensor >= 0 Then SysDI.IsDoorClose = SysDI.value(gblngDI_DoorCloseSensor)
        If gblngDI_DoorOpenSensor >= 0 Then SysDI.IsDoorOpen = SysDI.value(gblngDI_DoorOpenSensor)
        If gblngDI_DoorClamp >= 0 Then SysDI.IsDoorClamp = SysDI.value(gblngDI_DoorClamp)
        If gblngDI_ChamberOverheat >= 0 Then SysDI.IsOverHeat = SysDI.value(gblngDI_ChamberOverheat)
        If gblngDI_WaterAlarm >= 0 Then SysDI.IsWater = SysDI.value(gblngDI_WaterAlarm)
        If gblngDI_AirAlarm >= 0 Then SysDI.IsCDA = SysDI.value(gblngDI_AirAlarm)
        If gblngDI_CT1 >= 0 Then SysDI.IsLampError(1) = SysDI.value(gblngDI_CT1)
        If gblngDI_CT2 >= 0 Then SysDI.IsLampError(2) = SysDI.value(gblngDI_CT2)
        If gblngDI_CT3 >= 0 Then SysDI.IsLampError(3) = SysDI.value(gblngDI_CT3)
        If gblngDI_CT4 >= 0 Then SysDI.IsLampError(4) = SysDI.value(gblngDI_CT4)
        If SysDI.IsDoorClose = 0 And Para.intOpenDoorTime > 0 Then
            If frmProcess.tmrOpenDoor.Enabled = False And Kernel.intOpenDoorCount <> 9999 Then
                Kernel.intOpenDoorCount = 0
                frmProcess.tmrOpenDoor.Enabled = True
            End If
        Else
            frmProcess.tmrOpenDoor.Enabled = False
        End If
        
        Call FRB_ReceiveRA(0, 9, wDiVal)
        
        For i = 0 To 15
            SysDI.value(i + 16) = IIf(wDiVal And 2 ^ i, 1, 0)
        Next i
        If gblngDI_CoverAlarm1 >= 0 And gblngDI_CoverAlarm2 >= 0 Then SysDI.IsCoverAlarm = SysDI.value(gblngDI_CoverAlarm1) Or SysDI.value(gblngDI_CoverAlarm2)
        If gblngDI_CoverServoRdy1 >= 0 And gblngDI_CoverServoRdy2 >= 0 Then SysDI.IsCoverServoRdy = SysDI.value(gblngDI_CoverServoRdy1) And SysDI.value(gblngDI_CoverServoRdy2)
        If gblngDI_CoverOrigRdy1 >= 0 And gblngDI_CoverOrigRdy2 >= 0 Then SysDI.IsCoverOrigRdy = SysDI.value(gblngDI_CoverOrigRdy1) And SysDI.value(gblngDI_CoverOrigRdy2)
        If gblngDI_CoverIsMoving1 >= 0 And gblngDI_CoverIsMoving2 >= 0 Then SysDI.IsCoverMoving = SysDI.value(gblngDI_CoverIsMoving1) Or SysDI.value(gblngDI_CoverIsMoving2)
        If gblngDI_CoverUpInpos1 >= 0 And gblngDI_CoverUpInpos2 >= 0 Then SysDI.IsCoverUp = SysDI.value(gblngDI_CoverUpInpos1) And SysDI.value(gblngDI_CoverUpInpos2)
        If gblngDI_CoverDownInpos1 >= 0 And gblngDI_CoverDownInpos2 >= 0 Then SysDI.IsCoverDown = SysDI.value(gblngDI_CoverDownInpos1) And SysDI.value(gblngDI_CoverDownInpos2)

    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadDI_FR error"
    ShowAlarmFlash 1
End Sub

Public Sub ReadDI_HG()
    Dim i As Integer
    Dim lngValue As Long
    
On Error GoTo ERRLINE:
    
    If gbblnActiveHGD = True Then
    
        Call Ixud_ReadDI(gbintPCI822A, 1, lngValue)
        For i = 0 To 15
            SysDI.value(i) = lngValue Mod 2
            lngValue = CInt(Int(lngValue / 2))
        Next i
        
        If gblngDI_MC_Input >= 0 Then
            If SysDI.value(gblngDI_MC_Input) = 1 Then
                SysDI.IsReady = 0
            Else
                SysDI.IsReady = 1
            End If
        End If
        If gblngDI_ChamberGateL >= 0 Then SysDI.IsChamberGaugeL = SysDI.value(gblngDI_ChamberGateL)
        If gblngDI_EMS_Alarm >= 0 Then SysDI.IsEMO = SysDI.value(gblngDI_EMS_Alarm)
        If gblngDI_DoorCloseSensor >= 0 Then SysDI.IsDoorClose = SysDI.value(gblngDI_DoorCloseSensor)
        If gblngDI_DoorOpenSensor >= 0 Then SysDI.IsDoorOpen = SysDI.value(gblngDI_DoorOpenSensor)
        If gblngDI_ChamberOverheat >= 0 Then SysDI.IsOverHeat = SysDI.value(gblngDI_ChamberOverheat)
        If gblngDI_WaterAlarm >= 0 Then SysDI.IsWater = SysDI.value(gblngDI_WaterAlarm)
        If gblngDI_AirAlarm >= 0 Then SysDI.IsCDA = SysDI.value(gblngDI_AirAlarm)
        If gblngDI_CT1 >= 0 Then SysDI.IsLampError(1) = SysDI.value(gblngDI_CT1)
        If gblngDI_CT2 >= 0 Then SysDI.IsLampError(2) = SysDI.value(gblngDI_CT2)
        If gblngDI_CT3 >= 0 Then SysDI.IsLampError(3) = SysDI.value(gblngDI_CT3)
        If gblngDI_CT4 >= 0 Then SysDI.IsLampError(4) = SysDI.value(gblngDI_CT4)
        
        If SysDI.IsDoorClose = 0 And Para.intOpenDoorTime > 0 Then
            If frmProcess.tmrOpenDoor.Enabled = False And Kernel.intOpenDoorCount <> 9999 Then
                Kernel.intOpenDoorCount = 0
                frmProcess.tmrOpenDoor.Enabled = True
            End If
        Else
            frmProcess.tmrOpenDoor.Enabled = False
        End If
    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadDI_HG error"
    ShowAlarmFlash 1
End Sub

Public Sub ReadDI_AD()
    Dim i As Integer
    
On Error GoTo ERRLINE:
    Call advThermo.ReadDI
    For i = 0 To 15
        
        SysDI.value(i) = AvDaq17101.lngDI(i)
    Next i
    If gblngDI_MC_Input >= 0 Then
        If SysDI.value(gblngDI_MC_Input) = 1 Then
            SysDI.IsReady = 0
        Else
            SysDI.IsReady = 1
        End If
    End If
    If gblngDI_ChamberGateL >= 0 Then SysDI.IsChamberGaugeL = SysDI.value(gblngDI_ChamberGateL)
    If gblngDI_EMS_Alarm >= 0 Then SysDI.IsEMO = SysDI.value(gblngDI_EMS_Alarm)
    If gblngDI_DoorCloseSensor >= 0 Then SysDI.IsDoorClose = SysDI.value(gblngDI_DoorCloseSensor)
    If gblngDI_DoorOpenSensor >= 0 Then SysDI.IsDoorOpen = SysDI.value(gblngDI_DoorOpenSensor)
    If gblngDI_ChamberOverheat >= 0 Then SysDI.IsOverHeat = SysDI.value(gblngDI_ChamberOverheat)
    If gblngDI_WaterAlarm >= 0 Then SysDI.IsWater = SysDI.value(gblngDI_WaterAlarm)
    If gblngDI_AirAlarm >= 0 Then SysDI.IsCDA = SysDI.value(gblngDI_AirAlarm)
    If gblngDI_CT1 >= 0 Then SysDI.IsLampError(1) = SysDI.value(gblngDI_CT1)
    If gblngDI_CT2 >= 0 Then SysDI.IsLampError(2) = SysDI.value(gblngDI_CT2)
    If gblngDI_CT3 >= 0 Then SysDI.IsLampError(3) = SysDI.value(gblngDI_CT3)
    If gblngDI_CT4 >= 0 Then SysDI.IsLampError(4) = SysDI.value(gblngDI_CT4)
        
    
'    If gblngDI_WaterAlarm >= 0 Then gbblnWaterGate = IIf(AvDaq17101.lngDI(gblngDI_WaterAlarm) <> 0, True, False)
'    If gblngDI_AirAlarm >= 0 Then gbblnAirGate = IIf(AvDaq17101.lngDI(gblngDI_AirAlarm) <> 0, True, False)
'    If gblngDI_DoorOpenSensor >= 0 Then gbblnDoorOpen = IIf(AvDaq17101.lngDI(gblngDI_DoorOpenSensor) <> 0, True, False)
'    If gblngDI_DoorCloseSensor >= 0 Then gbblnDoorClose = IIf(AvDaq17101.lngDI(gblngDI_DoorCloseSensor) <> 0, True, False)
'    If gblngDI_DoorOpen >= 0 Then gbblnDoorOpenSwitch = IIf(AvDaq17101.lngDI(gblngDI_DoorOpen) <> 0, False, True)
'    If gblngDI_DoorClose >= 0 Then gbblnDoorCloseSwitch = IIf(AvDaq17101.lngDI(gblngDI_DoorClose) <> 0, False, True)
'    If gblngDI_ChamberOverheat >= 0 Then gbblnChamberOverheat = IIf(AvDaq17101.lngDI(gblngDI_ChamberOverheat) <> 0, True, False)
'    If gblngDI_SystemReady >= 0 Then gbblnPowerInput24 = IIf(AvDaq17101.lngDI(gblngDI_SystemReady) <> 0, False, True)
'    If gblngDI_EMS_Alarm >= 0 Then gbblnEMS_Alarm = IIf(AvDaq17101.lngDI(gblngDI_EMS_Alarm) <> 0, False, True)
'    If gblngDI_MC_Input >= 0 Then gbblnMC_Input = IIf(AvDaq17101.lngDI(gblngDI_MC_Input) <> 0, True, False)
'    If gblngDI_SystemAlarm >= 0 Then gbblnSystemAlarm = IIf(AvDaq17101.lngDI(gblngDI_SystemAlarm) <> 0, True, False)
'    If gblngDI_VacuumGateR >= 0 Then gbblnGaugeGateR = IIf(AvDaq17101.lngDI(gblngDI_VacuumGateR) <> 0, True, False)
'    If gblngDI_VacuumGateL >= 0 Then gbblnGaugeGateL = IIf(AvDaq17101.lngDI(gblngDI_VacuumGateL) <> 0, True, False)
'    If gblngDI_ChamberGateR >= 0 Then gbblnChamberGaugeGateR = IIf(AvDaq17101.lngDI(gblngDI_ChamberGateR) <> 0, True, False)
'    If gblngDI_ChamberGateL >= 0 Then gbblnChamberGaugeGateL = IIf(AvDaq17101.lngDI(gblngDI_ChamberGateL) <> 0, True, False)
'    If gblngDI_CT1 >= 0 Then gbIsCT1 = IIf(AvDaq17101.lngDI(gblngDI_CT1) <> 0, True, False)
'    If gblngDI_CT2 >= 0 Then gbIsCT2 = IIf(AvDaq17101.lngDI(gblngDI_CT2) <> 0, True, False)
'
'    If gblngDI_ARM_FRONT >= 0 Then gbblnArmInFront = IIf(AvDaq17101.lngDI(gblngDI_ARM_FRONT) <> 0, True, False)
'    If gblngDI_ARM_REAR >= 0 Then gbblnArmInRear = IIf(AvDaq17101.lngDI(gblngDI_ARM_REAR) <> 0, True, False)
'
'    If gblngDI_DoorClose >= 0 And Not gbblnDoorCloseSwitch Then
'        gbblnDoorClose = True
'    End If
'
'    If gbblnDoorOpen = True And gbintDO_Status(gblngDO_AlarmGreen) > 0 Then
'        SetDO gblngDO_ReleaseValve, True
'    Else
'        SetDO gblngDO_ReleaseValve, False
'    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadDI_AD error"
    ShowAlarmFlash 1
End Sub

Public Sub ReadAI_FR()
    Dim i As Integer
    Dim wRtn As Integer
    Dim fValue(0 To 15) As Single
    Dim wConfigCodeList(0 To 15) As Integer
    
    Dim Index As Integer
    Dim iRet As Integer
    Dim bRet As Boolean
    Dim iValue As Integer
    Dim lngValue As Long
    Dim sngValueAI As Single
    Dim sngValue As Single
        
On Error GoTo ERRLINE:
       
    For Index = 0 To 15
        wConfigCodeList(Index) = &H8 '// Set the range to +/-10 V for first channel.
    Next
    'wRtn = FRB_ReadAI_CHScan(wPort, wRAn, wResolution, cbxInputType.ListIndex, wConfigCodeList(0), fValue(0))
    wRtn = FRB_ReadAI_CHScan(0, 10, 12, 0, wConfigCodeList(0), fValue(0))
    For i = 0 To 15
        SysAI.value(i) = fValue(i)
        AverageValue i, SysAI.value(i)
    Next
    
    If gblngAI_Vacuum_Gauge >= 0 Then
        Kernel.sngPressure = AI2Vacuum(SysAI.AvgValue(gblngAI_Vacuum_Gauge))
        If SysAI.AvgValue(gblngAI_Vacuum_Gauge) < 0 Then
        ShowAlarmFlash 31
        End If
    Else
        Kernel.sngPressure = 760
    End If
    
    If gblngAI_Vacuum_Gauge2 >= 0 Then
        Kernel.sngPressure2 = AI2Vacuum(SysAI.AvgValue(gblngAI_Vacuum_Gauge2))
        If SysAI.AvgValue(gblngAI_Vacuum_Gauge2) < 0 Then
        ShowAlarmFlash 31
        End If
    Else
        Kernel.sngPressure2 = 760
    End If
    
    
    If gbintTCO2 >= 0 Then
        Kernel.sngOxygen = Kernel.sngTC(gbintTCO2)
    Else
        If gblngAI_Oxygen_Gauge >= 0 Then
            Kernel.sngOxygen = AI2Oxygen(SysAI.AvgValue(gblngAI_Oxygen_Gauge))
        Else
            Kernel.sngOxygen = 0
        End If
    End If
    For i = 0 To 5
        If gblngAI_MFC_Read(i) >= 0 Then
            SysAI.sngMFC(i) = AI2MFC(SysAI.AvgValue(gblngAI_MFC_Read(i)), gbsngMaxGasSLMP(i), SysAI.ErrorV(gblngAI_MFC_Read(i)))
        End If
    Next i
            
    If Para.IsUseCustom Then
        Kernel.sngPressure = Kernel.sngPressure * Para.sngRatioCUP
        For i = 0 To 5
            SysAI.sngMFC(i) = SysAI.sngMFC(i) * Para.sngRatioCUM(i)
        Next i
    End If
        
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadAI_FR AI error"
    ShowAlarmFlash 1

End Sub

Public Sub ReadAI_HG()
    Dim i As Integer
    
    Dim iRet As Integer
    Dim bRet As Boolean
    Dim iValue As Integer
    Dim lngValue As Long
    Dim sngValueAI As Single
    Dim sngValue As Single
        
On Error GoTo ERRLINE:
    If gbblnActiveHGD = True Then
    
        For i = 0 To 15
            Call Ixud_ReadAI(gbintPCI822A, i, IXUD_BI_10V, SysAI.value(i))
            AverageValue i, SysAI.value(i)
        Next i
        
        If gblngAI_Vacuum_Gauge >= 0 Then
            Kernel.sngPressure = AI2Vacuum(SysAI.AvgValue(gblngAI_Vacuum_Gauge))
        Else
            Kernel.sngPressure = 760
        End If
        If gblngAI_Oxygen_Gauge >= 0 Then
            Kernel.sngOxygen = AI2Oxygen(SysAI.AvgValue(gblngAI_Oxygen_Gauge))
        Else
            Kernel.sngOxygen = 0
        End If
        
        For i = 0 To 4
            If gblngAI_MFC_Read(i) >= 0 Then
                SysAI.sngMFC(i) = AI2MFC(SysAI.AvgValue(gblngAI_MFC_Read(i)), gbsngMaxGasSLMP(i), gbsngGasBias(i))
            End If
        Next i
    End If
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadAI_HG error"
    ShowAlarmFlash 1

End Sub

Public Sub ReadAI_AD()
    Dim i As Integer
    Dim j As Long
    Dim iRet As Integer
    Dim bRet As Boolean
    Dim iValue As Integer
    Dim lngValue As Long
    Dim sngValueAI As Single
        
        
On Error GoTo ERRLINE:
          
    Call advThermo.ReadAI(SysAI.value)
    For i = 0 To 4
        AverageValue i, SysAI.value(i)
        If gblngAI_MFC_Read(i) >= 0 Then
            SysAI.sngMFC(i) = AI2MFC(SysAI.AvgValue(gblngAI_MFC_Read(i)), gbsngMaxGasSLMP(i), gbsngGasBias(i))
        End If
    Next i
    If gblngAI_Vacuum_Gauge >= 0 Then
            Kernel.sngPressure = AI2Vacuum(SysAI.AvgValue(gblngAI_Vacuum_Gauge))
        Else
            Kernel.sngPressure = 760
        End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadAI_AD error"
    ShowAlarmFlash 1
End Sub

'Public Sub ReadTC()
'    Dim i As Integer
'    Dim offset As Integer
'    Dim rand As Single
'    Dim fTemp As Single
'    Dim sngTemp(8) As Single
'    Dim sngTemp1(8) As Single
'    Dim iErrCode As Long
'
'
'On Error GoTo ERRLINE
'
'    If Para.RtaType = 1 Or Para.RtaType = 2 Then Call advThermo.ReadTemperatureAllChannelFT(sngTemp)
'    If Para.RtaType = 3 Or Para.RtaType = 5 Then
'
'
'            iErrCode = mUSBIO.AI_ReadValueAnalog(sngTemp(0))
'            If iErrCode <> ERR_NO_ERR Then
'                intErrorCount = intErrorCount + 1
'                For i = 0 To 7
'                    If Kernel.sngTC(i) > 0 And Kernel.sngTC(i) < 999 Then
'                        sngTemp(i) = Kernel.sngTC(i)
'                    End If
'                Next i
'                If intErrorCount > 66 Then GoTo ERRLINE
'            Else
'                intErrorCount = 0
'            End If
'    End If
'    For i = 0 To 7
'        Kernel.sngTC(i) = sngTemp(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
'    Next i
'
'    If Para.UseMTC = 1 And Kernel.IsActiveTC2 = 1 Then
'        iErrCode = mUSBIO_1.AI_ReadValueAnalog(sngTemp1(0))
'        If iErrCode <> ERR_NO_ERR Then GoTo ERRLINE
'        For i = 8 To 15
'            Kernel.sngTC(i) = sngTemp1(i - 8) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
'        Next i
'    End If
'
'    If gbblnPlayFakeBall = True And gbintCurrProcessStep = GB_ACTION_INDEX_HOLD Then
'        offset = Kernel.sngTC(0) / 100 - 1
'        For i = 1 To 6
'            If (Kernel.sngTC(i) + PlayFakeBall(i)) < (Kernel.sngTC(0) - offset) Then
'                PlayFakeBall(i) = PlayFakeBall(i) + 0.1
'
'            ElseIf (Kernel.sngTC(i) + PlayFakeBall(i)) > (Kernel.sngTC(0) + offset) Then
'                PlayFakeBall(i) = PlayFakeBall(i) - 0.1
'
'            Else
'                Randomize
'                fTemp = Rnd * 0.3 - Rnd * 0.3
'                PlayFakeBall(i) = PlayFakeBall(i) + fTemp
'
'            End If
'            Kernel.sngTC(i) = Kernel.sngTC(i) + PlayFakeBall(i)
'        Next i
'    Else
'        For i = 0 To 6
'            PlayFakeBall(i) = 0
'        Next i
'        gbblnPlayFakeBall = False
'    End If
'
'    If Kernel.IsRun = 1 And Kernel.sngTC(0) < gbsngMinTemperature Then
'        gbstrAlarmHint = ",Err=" & CStr(iErrCode) & ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0") & ",Count=" & CStr(intErrorCount)
'        ShowAlarmFlash 8
'    End If
'
'    If Kernel.IsRun = 1 And Kernel.sngTC(0) > gbsngMaxTemperature Then
'        gbstrAlarmHint = ",Err=" & CStr(iErrCode) & ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0") & ",Count=" & CStr(intErrorCount)
'        ShowAlarmFlash 9
'    End If
'
'    Exit Sub
'ERRLINE:
'    gbstrAlarmHint = " Read TC error"
'    ShowAlarmFlash 1
'End Sub

Public Sub ReadTC_HGU_1()
    Dim i As Integer
    Dim iErrCode As Long
    
On Error GoTo ERRLINE
       
    iErrCode = mUSBIO_1.AI_ReadValueAnalog(sngTemp_1(0))
    If iErrCode <> ERR_NO_ERR Then
        intErrorCount_1 = intErrorCount_1 + 1
        For i = 0 To 7
            If Kernel.sngTC(i) > gbsngMinTemperature And Kernel.sngTC(i) < gbsngMaxTemperature Then
                sngTemp_1(i) = Kernel.sngTC(i)
            End If
        Next i
        If intErrorCount_1 > 66 Then GoTo ERRLINE
    Else
        intErrorCount_1 = 0
    End If
    
    If Para.intOnlyRecipe = 1 Then
        For i = 0 To 7
            If Kernel.IsRun = 1 Then
                Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                                gbsngPowerTC(i) * sngTemp_1(i) ^ 2 + gbsngPower3C(i) * sngTemp_1(i) ^ 3 + _
                                gbsngPower4C(i) * sngTemp_1(i) ^ 4 + gbsngPower5C(i) * sngTemp_1(i) ^ 5
            Else
                Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
            End If
            Kernel.sngOrigTC(i) = sngTemp_1(i)
        Next i
    Else
        For i = 0 To 7
            Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                               gbsngPowerTC(i) * sngTemp_1(i) ^ 2 + gbsngPower3C(i) * sngTemp_1(i) ^ 3 + _
                               gbsngPower4C(i) * sngTemp_1(i) ^ 4 + gbsngPower5C(i) * sngTemp_1(i) ^ 5
            Kernel.sngOrigTC(i) = sngTemp_1(i)
        Next i
    End If
    
    If Para.IsUseCustom Then
        For i = 0 To 7
            Kernel.sngTC(i) = Kernel.sngTC(i) * Para.sngRatioCUT(i)
        Next i
    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadTC_HGU_1 error,Count=" & CStr(intErrorCount_1)
    ShowAlarmFlash 1
End Sub

Public Sub ReadTC_HGU_2()
    Dim i As Integer
    Dim iErrCode As Long
    
On Error GoTo ERRLINE
       
    iErrCode = mUSBIO_2.AI_ReadValueAnalog(sngTemp_2(0))
    If iErrCode <> ERR_NO_ERR Then
        intErrorCount_2 = intErrorCount_2 + 1
        For i = 8 To 15
            If Kernel.sngTC(i) > gbsngMinTemperature And Kernel.sngTC(i) < gbsngMaxTemperature Then
                sngTemp_2(i - 8) = Kernel.sngTC(i)
            End If
        Next i
        If intErrorCount_2 > 66 Then GoTo ERRLINE
    Else
        intErrorCount_2 = 0
    End If
    
    If Para.intOnlyRecipe = 1 Then
        For i = 8 To 15
            If Kernel.IsRun = 1 Then
                Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                                gbsngPowerTC(i) * sngTemp_2(i - 8) ^ 2 + gbsngPower3C(i) * sngTemp_2(i - 8) ^ 3 + _
                                gbsngPower4C(i) * sngTemp_2(i - 8) ^ 4 + gbsngPower5C(i) * sngTemp_2(i - 8) ^ 5
            Else
                Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioEX(i) + gbsngErrorTC(i)
            End If
            
            Kernel.sngOrigTC(i) = sngTemp_2(i - 8)
        Next i
    Else
        For i = 8 To 15
            Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                            gbsngPowerTC(i) * sngTemp_2(i - 8) ^ 2 + gbsngPower3C(i) * sngTemp_2(i - 8) ^ 3 + _
                            gbsngPower4C(i) * sngTemp_2(i - 8) ^ 4 + gbsngPower5C(i) * sngTemp_2(i - 8) ^ 5
            Kernel.sngOrigTC(i) = sngTemp_2(i - 8)
        Next i
    End If
    
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadTC_HGU_2 error,Count=" & CStr(intErrorCount_2)
    ShowAlarmFlash 1
End Sub

Public Sub ReadTC_HGU_A()
    Dim i As Integer
    Dim iErrCode As Long
    
On Error GoTo ERRLINE
       
    iErrCode = mUSBIO_1.AI_ReadValueAnalog(sngTemp_2(0))
    If iErrCode <> ERR_NO_ERR Then
        intErrorCount_2 = intErrorCount_2 + 1
        For i = 8 To 15
            If Kernel.sngTC(i) > gbsngMinTemperature And Kernel.sngTC(i) < gbsngMaxTemperature Then
                sngTemp_2(i - 8) = Kernel.sngTC(i)
            End If
        Next i
        If intErrorCount_2 > 66 Then GoTo ERRLINE
    Else
        intErrorCount_2 = 0
    End If
    
    If Para.intOnlyRecipe = 1 Then
        For i = 8 To 15
            If Kernel.IsRun = 1 Then
                Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                                gbsngPowerTC(i) * sngTemp_2(i - 8) ^ 2 + gbsngPower3C(i) * sngTemp_2(i - 8) ^ 3 + _
                                gbsngPower4C(i) * sngTemp_2(i - 8) ^ 4 + gbsngPower5C(i) * sngTemp_2(i - 8) ^ 5
'                Kernel.sngTC(i) = sngTemp_2(i - 8)
            Else
                Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioEX(i) + gbsngErrorTC(i)
'                 Kernel.sngTC(i) = sngTemp_2(i - 8)
            End If
            
            Kernel.sngOrigTC(i) = sngTemp_2(i - 8)
        Next i
    Else
        For i = 8 To 15
            Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                            gbsngPowerTC(i) * sngTemp_2(i - 8) ^ 2 + gbsngPower3C(i) * sngTemp_2(i - 8) ^ 3 + _
                            gbsngPower4C(i) * sngTemp_2(i - 8) ^ 4 + gbsngPower5C(i) * sngTemp_2(i - 8) ^ 5
'            Kernel.sngTC(i) = sngTemp_2(i - 8)
            Kernel.sngOrigTC(i) = sngTemp_2(i - 8)
        Next i
    End If
    
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadTC_HGU_A error,Count=" & CStr(intErrorCount_2)
    ShowAlarmFlash 1
End Sub

Public Sub ReadTC_HGU_B()
    Dim i As Integer
    Dim Count As Integer
    Dim iErrCode As Long
    Dim sngTemp(10) As Single
    
    
On Error GoTo ERRLINE
       
    iErrCode = mUSBIO_2.AI_ReadValueAnalog(sngTemp(0))
    Count = 0
    If iErrCode <> ERR_NO_ERR Then
        Count = Count + 1
        For i = 16 To 23
            If Kernel.sngTC(i) > gbsngMinTemperature And Kernel.sngTC(i) < gbsngMaxTemperature Then
                sngTemp(i - 16) = Kernel.sngTC(i)
            End If
        Next i
        If Count > 66 Then GoTo ERRLINE
    Else
        Count = 0
    End If
    
    If Para.intOnlyRecipe = 1 Then
        For i = 16 To 23
            If Kernel.IsRun = 1 Then
                Kernel.sngTC(i) = sngTemp(i - 16) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                                gbsngPowerTC(i) * sngTemp(i - 16) ^ 2 + gbsngPower3C(i) * sngTemp(i - 16) ^ 3 + _
                                gbsngPower4C(i) * sngTemp(i - 16) ^ 4 + gbsngPower5C(i) * sngTemp(i - 16) ^ 5
'                 Kernel.sngTC(i) = sngTemp(i - 16)
            Else
                Kernel.sngTC(i) = sngTemp(i - 16) * gbsngRatioEX(i) + gbsngErrorTC(i)
'                 Kernel.sngTC(i) = sngTemp(i - 16)
            End If
            
            Kernel.sngOrigTC(i) = sngTemp(i - 16)
        Next i
    Else
        For i = 16 To 23
            Kernel.sngTC(i) = sngTemp(i - 16) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                            gbsngPowerTC(i) * sngTemp(i - 16) ^ 2 + gbsngPower3C(i) * sngTemp(i - 16) ^ 3 + _
                            gbsngPower4C(i) * sngTemp(i - 16) ^ 4 + gbsngPower5C(i) * sngTemp(i - 16) ^ 5
'            Kernel.sngTC(i) = sngTemp(i - 16)
            Kernel.sngOrigTC(i) = sngTemp(i - 16)
        Next i
    End If
    
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadTC_HGU_B error,Count=" & CStr(Count)
    ShowAlarmFlash 1
End Sub

Public Sub ReadTC_HGT_1()
    
On Error GoTo ERRLINE
       
    Call frmTCP.ReadTC
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " ReadTC_HGT_1 error"
    ShowAlarmFlash 1
End Sub

Public Sub ReadTC_AD()
    Dim i As Integer
On Error GoTo ERRLINE
       
    Call advThermo.ReadTemperatureAllChannelFT(sngTemp_1)
    For i = 0 To 7
        Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngPowerTC(i) * sngTemp_1(i) * sngTemp_1(i) + gbsngErrorTC(i)
    Next i
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " Read Advan TC error"
    ShowAlarmFlash 1
End Sub

Public Sub CheckTC()
    Dim i As Integer
    Dim offset As Integer
    Dim rand As Single
    Dim fTemp As Single
    Dim iErrCode As Long
    
    
On Error GoTo ERRLINE
       
'    For i = 0 To 7
'        Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
'    Next i
'    For i = 8 To 15
'        Kernel.sngTC(i) = sngTemp_2(i - 8) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
'    Next i
        
    If gbblnPlayFakeBall = True And gbintCurrProcessStep = GB_ACTION_INDEX_HOLD Then
        offset = Kernel.sngTC(0) / 100 - 1
        For i = 1 To 6
            If (Kernel.sngTC(i) + PlayFakeBall(i)) < (Kernel.sngTC(0) - offset) Then
                PlayFakeBall(i) = PlayFakeBall(i) + 0.1
                
            ElseIf (Kernel.sngTC(i) + PlayFakeBall(i)) > (Kernel.sngTC(0) + offset) Then
                PlayFakeBall(i) = PlayFakeBall(i) - 0.1
                
            Else
                Randomize
                fTemp = Rnd * 0.3 - Rnd * 0.3
                PlayFakeBall(i) = PlayFakeBall(i) + fTemp
                
            End If
            Kernel.sngTC(i) = Kernel.sngTC(i) + PlayFakeBall(i)
        Next i
    Else
        For i = 0 To 6
            PlayFakeBall(i) = 0
        Next i
        gbblnPlayFakeBall = False
    End If
    
    If Kernel.IsRun = 1 And (Kernel.sngTC(0) < gbsngMinTemperature Or Kernel.sngTC(1) < gbsngMinTemperature) Then
        gbstrAlarmHint = ",Err=" & CStr(iErrCode) & ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0") & ",Count=" & CStr(intErrorCount_1)
        ShowAlarmFlash 8
    End If
    
    If Kernel.IsRun = 1 And (Kernel.sngTC(0) > gbsngMaxTemperature Or Kernel.sngTC(1) > gbsngMaxTemperature) Then
        gbstrAlarmHint = ",Err=" & CStr(iErrCode) & ",TC=" & Format(Kernel.sngTC(0), "0.0") & ",MTC=" & Format(Kernel.sngTC(1), "0.0") & ",Count=" & CStr(intErrorCount_1)
        ShowAlarmFlash 9
    End If
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " CheckTC error"
    ShowAlarmFlash 1
End Sub


Public Sub ResetDO()
    Dim i As Long
    
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To 15
            SetDO i, False
        Next i
        SysDO.OnceValue = 0
        Call Ixud_WriteDO(gbintPCI822A, 0, 0)
    End If
    If Para.RtaType = 2 Then
        For i = 0 To 15
            Call advThermo.WriteDO(i, False)
        Next i
        SysDO.OnceValue = 0
        Call Ixud_WriteDO(gbintPCI822A, 0, 0)
    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        For i = 0 To 31
            SetDO i, False
        Next i


    End If
    
End Sub

Public Sub ResetAO()
    Dim i As Long
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To 15
            Call Ixud_WriteAOVoltage(gbintPIODA8A, i, 0)
        Next i
        For i = 0 To 2
            Kernel.sngCurrOutMFC(i) = 0
        Next i
        
    End If
    If Para.RtaType = 2 Then
        Call advThermo.WriteAO(0, 0)
        Call advThermo.WriteAO(1, 0)
        For i = 0 To 2
            Kernel.sngCurrOutMFC(i) = 0
        Next i
    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        If Kernel.IsActiveIO = 0 Then
            Exit Sub
        End If
        For i = 0 To 19
            SetAO i, 0
        Next i
    End If
    For i = 0 To GB_SCR_MAX - 1
        Kernel.sngCurrOutSCR(i) = 0
    Next i
    For i = 0 To 19
        SetAO i, 0
    Next i
End Sub

Public Function SetDO(lngIndex As Long, blnIsOn As Boolean) As Boolean
    Dim i As Integer
    Dim iPort As Integer
    Dim lngValue As Long
            
    On Error GoTo ERRLINE

    If lngIndex >= 0 Then
        
        If Para.RtaType = 1 Or Para.RtaType = 3 Then
            lngValue = 2 ^ lngIndex
            If blnIsOn = True Then
                If SysDO.value(lngIndex) = 0 Then
                    SysDO.OnceValue = SysDO.OnceValue + lngValue
                    SysDO.value(lngIndex) = 1
                    Call Ixud_WriteDO(gbintPCI822A, 0, SysDO.OnceValue)
                End If
            Else
                If SysDO.value(lngIndex) = 1 Then
                    SysDO.OnceValue = SysDO.OnceValue - lngValue
                    SysDO.value(lngIndex) = 0
                    Call Ixud_WriteDO(gbintPCI822A, 0, SysDO.OnceValue)
                End If
            End If
            SetDO = True
        End If
        If Para.RtaType = 2 Then
            If blnIsOn Then
                SysDO.value(lngIndex) = 1
            Else
                SysDO.value(lngIndex) = 0
            End If
            Call advThermo.WriteDO(lngIndex, blnIsOn)
            SetDO = True
            gbintDO_Status(lngIndex) = IIf(blnIsOn = True, 1, 0)
        End If
        If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
            If Kernel.IsActiveIO = 0 Then
                Exit Function
            End If
            SysDO.value(lngIndex) = IIf(blnIsOn = True, 1, 0)
            lngValue = 0
            If lngIndex < 16 Then
                iPort = 0
                For i = 0 To 15
                    If SysDO.value(i) = 1 Then
                        lngValue = lngValue + 2 ^ i
                    End If
                Next i
                i = FRB_SendSA(0, 0, CInt("&H" & Hex(lngValue)))
                If i <> 0 Then Call FRB_SendSA(0, 0, CInt("&H" & Hex(lngValue)))
                
                
            Else
                iPort = 1
                For i = 16 To 31
                    If SysDO.value(i) = 1 Then
                        lngValue = lngValue + 2 ^ (i - 16)
                    End If
                Next i
                i = FRB_SendSA(0, iPort, CInt("&H" & Hex(lngValue)))
                If i <> 0 Then Call FRB_SendSA(0, iPort, CInt("&H" & Hex(lngValue)))
            End If
            SetDO = True
        End If
    End If
    SetDO = True
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Set DO error"
    ShowAlarmFlash 1
End Function

Public Function SetAO(lngIndex As Long, sngValue As Single) As Boolean
    Dim intSA As Integer
    Dim intCH As Integer
    Dim sngVolt As Single
    Dim intRet As Integer
    Dim C As Integer
    Dim VType As Integer
    On Error GoTo ERRLINE
       
    
    
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        gbstrAlarmHint = " SetAO AD error"
        sngVolt = sngValue / 10
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngIndex, sngVolt)
        
    End If
    If Para.RtaType = 2 Then
        gbstrAlarmHint = " SetAO HG error"
        sngVolt = sngValue / 10
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngIndex, sngVolt)
    End If
    
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        gbstrAlarmHint = " SetAO FR error"
        sngVolt = sngValue / 20
        If CH2FR(lngIndex, intSA, intCH) = True Then
            'intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, sngVolt, 2)   '0~5V
            VType = &H34
            If gbintAO_Type(lngIndex) = 1 Then VType = &H32
            For C = 1 To 3
                intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, VType, sngVolt, C)   '0~5V
                If intRet = 0 Then Exit For
                DelayTime (1)
            Next C
        End If
    End If
    'SysAO.value(lngIndex) = sngVolt
    Exit Function
ERRLINE:
    
    ShowAlarmFlash 1
End Function

    
Public Function SetAO_SCR(sngValue As Single, sngInputWeight() As Single) As Boolean
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    Dim i As Integer
    Dim C As Integer
    Dim VType As Integer
    Dim intSA As Integer
    Dim intCH As Integer
    Dim maxOut As Single
    Dim intRet As Integer
    
    On Error GoTo ERRLINE
    
    If Kernel.IsRun = 1 Then
        If sngValue > (frmRecipeEdit.sngRecipeIntLimit / 10) Then
            gbstrAlarmHint = " SetAO_SCR_LIMIT=" & CStr(sngValue)
            ShowAlarmFlash 4
            Exit Function
        End If
    End If
    
    If sngValue > 10 Then sngValue = 10 'Max AO output
    If sngValue < 0 Then sngValue = 0   'Min AO output
    
    lngSCR_AOChannel(0) = gblngAO_SCR_TBC
    lngSCR_AOChannel(1) = gblngAO_SCR_TR
    lngSCR_AOChannel(2) = gblngAO_SCR_TL
    lngSCR_AOChannel(3) = gblngAO_SCR_BF
    lngSCR_AOChannel(4) = gblngAO_SCR_BR
    '120713 Josh
    lngSCR_AOChannel(5) = gblngAO_SCR_6
    lngSCR_AOChannel(6) = gblngAO_SCR_7
    lngSCR_AOChannel(7) = gblngAO_SCR_8
    lngSCR_AOChannel(8) = gblngAO_SCR_9
    lngSCR_AOChannel(9) = gblngAO_SCR_10
    lngSCR_AOChannel(10) = gblngAO_SCR_11
    lngSCR_AOChannel(11) = gblngAO_SCR_12
    lngSCR_AOChannel(12) = gblngAO_SCR_13
    lngSCR_AOChannel(13) = gblngAO_SCR_14
    lngSCR_AOChannel(14) = gblngAO_SCR_15
    lngSCR_AOChannel(15) = gblngAO_SCR_16
    lngSCR_AOChannel(16) = gblngAO_SCR_17
        
    maxOut = 10
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then maxOut = 5
    
    For i = 0 To gbintNumOfBanks - 1
        Kernel.sngCurrOutSCR(i) = sngValue * sngInputWeight(i) / 100
        If Kernel.sngCurrOutSCR(i) > maxOut Then
            Kernel.sngCurrOutSCR(i) = maxOut
        End If
    Next i
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To gbintNumOfBanks - 1
            Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(i), Kernel.sngCurrOutSCR(i))
        Next i
    End If
    If Para.RtaType = 2 Then
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(0), Kernel.sngCurrOutSCR(0))
    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        
        For i = 0 To gbintNumOfBanks - 1
            If CH2FR(lngSCR_AOChannel(i), intSA, intCH) = True Then
                VType = &H34
                If gbintAO_Type(lngSCR_AOChannel(i)) = 1 Then VType = &H32
                For C = 1 To 3
                    intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, VType, Kernel.sngCurrOutSCR(i), C)    '0~5V
                    If intRet = 0 Then Exit For
                    DelayTime (C)
                Next C
                'Call FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, Kernel.sngCurrOutSCR(i), 2)    '0~5V
            End If
        Next i
        
    End If
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Set AO SCR error"
    ShowAlarmFlash 1
End Function

Public Function SetAO_SCR_Multi(sngValue() As Single, sngInputWeight() As Single) As Boolean
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    Dim i As Integer
    Dim no As Integer
    Dim VType As Integer
    
    Dim intSA As Integer
    Dim intCH As Integer
    Dim maxOut As Single
    Dim intRet As Integer
    Dim C As Integer
    
    On Error GoTo ERRLINE
    
    lngSCR_AOChannel(0) = gblngAO_SCR_TBC
    lngSCR_AOChannel(1) = gblngAO_SCR_TR
    lngSCR_AOChannel(2) = gblngAO_SCR_TL
    lngSCR_AOChannel(3) = gblngAO_SCR_BF
    lngSCR_AOChannel(4) = gblngAO_SCR_BR
    '120713 Josh
    lngSCR_AOChannel(5) = gblngAO_SCR_6
    lngSCR_AOChannel(6) = gblngAO_SCR_7
    lngSCR_AOChannel(7) = gblngAO_SCR_8
    lngSCR_AOChannel(8) = gblngAO_SCR_9
    lngSCR_AOChannel(9) = gblngAO_SCR_10
    lngSCR_AOChannel(10) = gblngAO_SCR_11
    lngSCR_AOChannel(11) = gblngAO_SCR_12
    '190731 Josh
    lngSCR_AOChannel(12) = gblngAO_SCR_13
    lngSCR_AOChannel(13) = gblngAO_SCR_14
    lngSCR_AOChannel(14) = gblngAO_SCR_15
    lngSCR_AOChannel(15) = gblngAO_SCR_16
    lngSCR_AOChannel(16) = gblngAO_SCR_17
    maxOut = 10
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then maxOut = 5
    
    For i = 0 To GB_MAX_LOOPS - 1
        If MultiLoop.blnUseLoop(i) = True Then
            If Kernel.IsRun = 1 Then
                If sngValue(i) > (frmRecipeEdit.sngRecipeIntLimit / 10) Then
                    gbstrAlarmHint = " SetAO_SCR_LIMIT=" & CStr(i) & "," & CStr(sngValue(i))
                    ShowAlarmFlash 4
                    Exit Function
                End If
            End If
            no = MultiLoop.intLoopA(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopB(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopC(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopD(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopE(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopF(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopG(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopH(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopJ(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
            no = MultiLoop.intLoopK(i)
            If no > 0 Then
                no = no - 1
                Kernel.sngCurrOutSCR(no) = sngValue(i) * sngInputWeight(no) / 100
            End If
        End If
    Next i
        
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To gbintNumOfBanks - 1
            Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(i), Kernel.sngCurrOutSCR(i))
        Next i
    End If
    If Para.RtaType = 2 Then
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(0), Kernel.sngCurrOutSCR(0))
    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        For i = 0 To gbintNumOfBanks - 1
            If CH2FR(lngSCR_AOChannel(i), intSA, intCH) = True Then
                VType = &H34
                If gbintAO_Type(lngSCR_AOChannel(i)) = 1 Then VType = &H32
                If Kernel.sngCurrOutSCR(i) > maxOut Then
                    Kernel.sngCurrOutSCR(i) = maxOut
                End If
                For C = 1 To 3
                    intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, VType, Kernel.sngCurrOutSCR(i), C)    '0~5V
                    If intRet = 0 Then Exit For
                    DelayTime (C)
                Next C
                
                'Call FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, Kernel.sngCurrOutSCR(i), 2)    '0~5V
            End If
        Next i
    End If
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Set AO SCR error"
    ShowAlarmFlash 1
End Function

Public Function SetAO_MFC(sngGasValue() As Single) As Boolean
    Dim sngMFCValue(10) As Single
    Dim lngMFCAOChannel(10) As Long
    Dim lngGASDOChannel(10) As Long
    Dim i As Integer
    Dim j As Integer
    Dim VType As Integer
    Dim intSA As Integer
    Dim intCH As Integer
    Dim intRet As Integer
    Dim C As Integer
    
    lngMFCAOChannel(0) = gblngAO_MFC1
    lngMFCAOChannel(1) = gblngAO_MFC2
    lngMFCAOChannel(2) = gblngAO_MFC3
    lngMFCAOChannel(3) = gblngAO_MFC4
    lngMFCAOChannel(4) = gblngAO_MFC5
    lngMFCAOChannel(5) = gblngAO_MFC6
    lngGASDOChannel(0) = gblngDO_GasValve1
    lngGASDOChannel(1) = gblngDO_GasValve2
    lngGASDOChannel(2) = gblngDO_GasValve3
    lngGASDOChannel(3) = gblngDO_GasValve4
    lngGASDOChannel(4) = gblngDO_GasValve5
    lngGASDOChannel(5) = gblngDO_GasValve6
    
'    Select Case cbxRange.ListIndex
'
'    Case 0        '// 0 ~ 20 mA
'     wAOConfig = &H30
'     lblType.Caption = "(mA)"
'    Case 1       '// 4 ~ 20 mA
'     wAOConfig = &H31
'     lblType.Caption = "(mA)"
'    Case 2      '// 0 ~ 10 V
'     wAOConfig = &H32
'     lblType.Caption = "(V)"
'    Case 3       '// +/- 10 V
'     wAOConfig = &H33
'     lblType.Caption = "(V)"
'    Case 4       '// 0 ~  5 V
'     wAOConfig = &H34
'     lblType.Caption = "(V)"
'    Case 5        '// +/-  5 V
'     wAOConfig = &H35
'     lblType.Caption = "(V)"
'    End Select
    
    For i = 0 To 5
        If gbintGasEnable(i) > 0 And gbsngMaxGasSLMP(i) > 0 Then sngMFCValue(i) = 5 * sngGasValue(i) / gbsngMaxGasSLMP(i)
        If sngMFCValue(i) > 5 Then sngMFCValue(i) = 5
        Kernel.sngCurrOutMFC(i) = sngGasValue(i)
        
        If sngGasValue(i) > 0 Then
            If gbsngGasError(i) <> 0 And frmDiagnosis.tmrCheckMFC(i).Enabled = False Then
                gbsngGasErrorC(i) = 0
                frmDiagnosis.tmrCheckMFC(i).Enabled = True
            End If
        
        End If
    Next i
        
    For i = 0 To 5
        If lngGASDOChannel(i) >= 0 Then
            If sngMFCValue(i) > 0 And SysDO.value(lngGASDOChannel(i)) = 0 Then
                Call SetDO(lngGASDOChannel(i), True)
            ElseIf sngMFCValue(i) = 0 And _
               SysDO.value(lngGASDOChannel(i)) = 1 Then
                Call SetDO(lngGASDOChannel(i), False)
            End If
        End If
    Next i
          
    
    
    For i = 0 To 5
        If sngMFCValue(i) > 0 Then
            If lngMFCAOChannel(i) < 20 Then
                If Para.RtaType = 1 Or Para.RtaType = 2 Or Para.RtaType = 3 Then
                    Call Ixud_WriteAOVoltage(gbintPIODA8A, lngMFCAOChannel(i), sngMFCValue(i)) '0~5V
                End If
                If (Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9) And Kernel.IsActiveIO = 1 Then
                    If CH2FR(lngMFCAOChannel(i), intSA, intCH) = True Then
                        VType = &H34
                        If gbintAO_Type(lngMFCAOChannel(i)) = 1 Then VType = &H32
                        
                        
                        For C = 1 To 3
                            intRet = intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, VType, sngMFCValue(i), C)   '0~5V
                            If intRet = 0 Then Exit For
                            DelayTime (C)
                        Next C
                        
                        If intRet <> 0 Then
                            gbstrAlarmHint = " SetAO_MFC error"
                            ShowAlarmFlash 1
                        End If
                        
'                        intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, sngMFCValue(i), 2)   '0~5V
'                        intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, sngMFCValue(i), 2)   '0~5V
'                        If intRet <> 0 Then
'                            For j = 0 To 10
'                                intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, sngMFCValue(i), 2) '0~5V
'                                If intRet = 0 Then
'                                    Exit Function
'                                End If
'                            Next j
'                            gbstrAlarmHint = " SetAO_MFC error"
'                            ShowAlarmFlash 1
'                        End If
                    
                    End If
                End If
            End If
        ElseIf sngMFCValue(i) = 0 Then
            If lngMFCAOChannel(i) < 20 Then
                If Para.RtaType = 1 Or Para.RtaType = 3 Then
                    Call Ixud_WriteAOVoltage(gbintPIODA8A, lngMFCAOChannel(i), 0) '0~5V
                End If
                If (Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9) And Kernel.IsActiveIO = 1 Then
                    If CH2FR(lngMFCAOChannel(i), intSA, intCH) = True Then
                        VType = &H34
                        If gbintAO_Type(lngMFCAOChannel(i)) = 1 Then VType = &H32
                        
                        For C = 1 To 3
                            intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, VType, 0, C)
                            If intRet = 0 Then Exit For
                            DelayTime (C)
                        Next C
                        If intRet <> 0 Then
                            gbstrAlarmHint = " SetAO_MFC error"
                            ShowAlarmFlash 1
                        End If
                        
'                        intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, 0, 2)    '0~5V
'                        intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, 0, 2)    '0~5V
'                        If intRet <> 0 Then
'                            For j = 0 To 10
'                                intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, 0, 2) '0~5V
'                                If intRet = 0 Then
'                                    Exit Function
'                                End If
'                            Next j
'                            gbstrAlarmHint = " SetAO_MFC error"
'                            ShowAlarmFlash 1
'                        End If
                    
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function CH2FR(lngIndex As Long, ByRef SA As Integer, ByRef CH As Integer) As Boolean
    If lngIndex >= 0 And lngIndex < 4 Then
        SA = 2
        CH = lngIndex
    ElseIf lngIndex > 3 And lngIndex < 8 Then
        SA = 3
        CH = lngIndex - 4
    ElseIf lngIndex > 7 And lngIndex < 12 Then
        SA = 4
        CH = lngIndex - 8
    ElseIf lngIndex > 11 And lngIndex < 16 Then
        SA = 5
        CH = lngIndex - 12
    ElseIf lngIndex > 15 And lngIndex < 20 Then
        SA = 6
        CH = lngIndex - 16
    Else
        CH2FR = False
        Exit Function
    End If
    CH2FR = True
    
End Function
Public Function Lamp2Scr(lngIndex As Long, ByRef SA As Integer) As Boolean
    If lngIndex >= 0 And lngIndex <= 2 Then
        SA = 0
    ElseIf lngIndex >= 3 And lngIndex <= 5 Then
        SA = 1
    ElseIf lngIndex >= 6 And lngIndex <= 7 Then
        SA = 2
    ElseIf lngIndex >= 8 And lngIndex <= 9 Then
        SA = 3
    ElseIf lngIndex >= 10 And lngIndex <= 11 Then
        SA = 4
    ElseIf lngIndex >= 12 And lngIndex <= 14 Then
        SA = 5
    ElseIf lngIndex >= 15 And lngIndex <= 17 Then
        SA = 6
    ElseIf lngIndex >= 18 And lngIndex <= 20 Then
        SA = 7
    ElseIf lngIndex >= 21 And lngIndex <= 24 Then
        SA = 8
    ElseIf lngIndex >= 25 And lngIndex <= 28 Then
        SA = 9
    ElseIf lngIndex >= 29 And lngIndex <= 32 Then
        SA = 10
    ElseIf lngIndex >= 33 And lngIndex <= 35 Then
        SA = 11
    Else
        Lamp2Scr = False
        Exit Function
    End If
    Lamp2Scr = True
    
End Function

Public Function SetTower(intIndex As Integer, blnIsOn As Boolean) As Boolean
    Select Case intIndex
        Case 0 'close signal
            Call SetDO(gblngDO_AlarmRed, False)
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, False)
            If GbShowDebugButton = 1 Then Call SetDO(gblngDO_AlarmBlue, False)
        Case 1 'red
            Call SetDO(gblngDO_AlarmRed, True)
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, False)
            If GbShowDebugButton = 1 Then Call SetDO(gblngDO_AlarmBlue, False)
        Case 2 'Yellow
            Call SetDO(gblngDO_AlarmRed, False)
            Call SetDO(gblngDO_AlarmYellow, True)
            Call SetDO(gblngDO_AlarmGreen, False)
            If GbShowDebugButton = 1 Then Call SetDO(gblngDO_AlarmBlue, False)
        Case 3 'Green
            Call SetDO(gblngDO_AlarmRed, False)
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, True)
            If GbShowDebugButton = 1 Then Call SetDO(gblngDO_AlarmBlue, False)
        Case 4 'Blue
            Call SetDO(gblngDO_AlarmBlue, blnIsOn)
            Call SetDO(gblngDO_AlarmRed, False)
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, False)
        Case 5 'red--Only Beep
            Call SetDO(gblngDO_AlarmRed, blnIsOn)
            
    End Select
    gbintTowerIndex = intIndex
End Function

Public Sub ControlSCR(sngValue As Single, sngRatioTBC As Single, sngRatioTR As Single, sngRatioTL As Single, sngRatioBF As Single, sngRatioBR As Single, sngRatio6 As Single, sngRatio7 As Single, sngRatio8 As Single, sngRatio9 As Single, sngRatio10 As Single, sngRatio11 As Single, sngRatio12 As Single, sngRatio13 As Single, sngRatio14 As Single, sngRatio15 As Single, sngRatio16 As Single, sngRatio17 As Single)
    Dim iRet As Integer
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    Dim sngOutValue(GB_SCR_MAX) As Single
    Dim sngOutputWeight(GB_SCR_MAX) As Single
    Dim i As Integer
    Dim intRet As Integer
    Dim VType As Integer
    Dim C As Integer
    
    'Rev 12.0.0.2 add intensity limit
    If Kernel.IsRun = 1 Then
        If sngValue > (frmRecipeEdit.sngRecipeIntLimit / 10) Then
            ShowAlarmFlash 4
            Exit Sub
        End If
    End If
    If sngValue > 10 Then sngValue = 10 'Max AO output
    If sngValue < 0 Then sngValue = 0   'Min AO output
    
    lngSCR_AOChannel(0) = gblngAO_SCR_TBC
    lngSCR_AOChannel(1) = gblngAO_SCR_TR
    lngSCR_AOChannel(2) = gblngAO_SCR_TL
    lngSCR_AOChannel(3) = gblngAO_SCR_BF
    lngSCR_AOChannel(4) = gblngAO_SCR_BR
    '120713 Josh
    lngSCR_AOChannel(5) = gblngAO_SCR_6
    lngSCR_AOChannel(6) = gblngAO_SCR_7
    lngSCR_AOChannel(7) = gblngAO_SCR_8
    lngSCR_AOChannel(8) = gblngAO_SCR_9
    lngSCR_AOChannel(9) = gblngAO_SCR_10
    lngSCR_AOChannel(10) = gblngAO_SCR_11
    lngSCR_AOChannel(11) = gblngAO_SCR_12
    lngSCR_AOChannel(12) = gblngAO_SCR_13
    lngSCR_AOChannel(13) = gblngAO_SCR_14
    lngSCR_AOChannel(14) = gblngAO_SCR_15
    lngSCR_AOChannel(15) = gblngAO_SCR_16
    lngSCR_AOChannel(16) = gblngAO_SCR_17
    
    sngOutputWeight(0) = sngRatioTBC
    sngOutputWeight(1) = sngRatioTR
    sngOutputWeight(2) = sngRatioTL
    sngOutputWeight(3) = sngRatioBF
    sngOutputWeight(4) = sngRatioBR
    
    sngOutputWeight(5) = sngRatio6
    sngOutputWeight(6) = sngRatio7
    sngOutputWeight(7) = sngRatio8
    sngOutputWeight(8) = sngRatio9
    sngOutputWeight(9) = sngRatio10
    sngOutputWeight(10) = sngRatio11
    sngOutputWeight(11) = sngRatio12
    sngOutputWeight(12) = sngRatio13
    sngOutputWeight(13) = sngRatio14
    sngOutputWeight(14) = sngRatio15
    sngOutputWeight(15) = sngRatio16
    sngOutputWeight(16) = sngRatio17
    For i = 0 To GB_SCR_MAX - 1
        gbsngCurrIntensity(i) = sngValue * sngOutputWeight(i) / 100
        sngOutValue(i) = sngValue * sngOutputWeight(i) / 100
        If sngOutValue(i) > 10 Then
            sngOutValue(i) = 10
        End If
    Next i
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To GB_SCR_MAX - 1
            Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(i), sngOutValue(i))
        Next i
    End If
    If Para.RtaType = 2 Then
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(0), sngOutValue(0))
'        Call advThermo.WriteAO(0, sngOutValue(0))
'        Call advThermo.WriteAO(1, sngOutValue(1))
    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        For i = 0 To GB_SCR_MAX - 1
            VType = &H34
            If gbintAO_Type(lngSCR_AOChannel(i)) = 1 Then VType = &H32
            
            For C = 1 To 3
                intRet = FRB_WriteAOFloat(0, 0, lngSCR_AOChannel(i), 12, VType, Kernel.sngCurrOutSCR(i), C)    '0~5V
                If intRet = 0 Then Exit For
                DelayTime (1)
            Next C
            'Call FRB_WriteAOFloat(0, 0, lngSCR_AOChannel(i), 12, &H34, Kernel.sngCurrOutSCR(i), 2)    '0~5V
        Next i
    End If
End Sub

Public Sub ResetSCR()
    Dim i As Integer
    Dim iRet As Integer
    Dim VType As Integer
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    
    lngSCR_AOChannel(0) = gblngAO_SCR_TBC
    lngSCR_AOChannel(1) = gblngAO_SCR_TR
    lngSCR_AOChannel(2) = gblngAO_SCR_TL
    lngSCR_AOChannel(3) = gblngAO_SCR_BF
    lngSCR_AOChannel(4) = gblngAO_SCR_BR
    '120713 Josh
    lngSCR_AOChannel(5) = gblngAO_SCR_6
    lngSCR_AOChannel(6) = gblngAO_SCR_7
    lngSCR_AOChannel(7) = gblngAO_SCR_8
    lngSCR_AOChannel(8) = gblngAO_SCR_9
    lngSCR_AOChannel(9) = gblngAO_SCR_10
    lngSCR_AOChannel(10) = gblngAO_SCR_11
    lngSCR_AOChannel(11) = gblngAO_SCR_12
    lngSCR_AOChannel(12) = gblngAO_SCR_13
    lngSCR_AOChannel(13) = gblngAO_SCR_14
    lngSCR_AOChannel(14) = gblngAO_SCR_15
    lngSCR_AOChannel(15) = gblngAO_SCR_16
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To GB_SCR_MAX - 1
            Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(i), 0)
        Next i
    End If
    If Para.RtaType = 2 Then
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngSCR_AOChannel(0), 0)

    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        For i = 0 To GB_SCR_MAX - 1
            VType = &H34
            If gbintAO_Type(lngSCR_AOChannel(i)) = 1 Then VType = &H32
            Call FRB_WriteAOFloat(0, 0, lngSCR_AOChannel(i), 12, VType, 0, 2)    '0~5V
        Next i
    End If
End Sub

Public Function SetAlarmStatus(intStatus As Integer)
    
    
    Select Case intStatus
        Case 0 'nothing
            If frmConfiguration.tmrFinishedBeep.Enabled = False Then
                Call SetDO(gblngDO_AlarmRed, False)
            End If
            
            Call SetDO(gblngDO_AlarmYellow, False)
            SetAlarmStatus = SetDO(gblngDO_AlarmBlue, False)
        Case 1 'Red
            Call SetDO(gblngDO_AlarmRed, True)
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, False)
            SetAlarmStatus = SetDO(gblngDO_AlarmBlue, False)
        Case 2 'Yellow
            If frmConfiguration.tmrFinishedBeep.Enabled = False Then
                Call SetDO(gblngDO_AlarmRed, False)
            End If
            If frmConfiguration.tmrFinishedLight.Enabled = False Then
                Call SetDO(gblngDO_AlarmYellow, True)
                Call SetDO(gblngDO_AlarmGreen, False)
                SetAlarmStatus = SetDO(gblngDO_AlarmBlue, False)
            End If
        Case 3 'Green
            If frmConfiguration.tmrFinishedBeep.Enabled = False Then
                Call SetDO(gblngDO_AlarmRed, False)
            End If
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, True)
            SetAlarmStatus = SetDO(gblngDO_AlarmBlue, False)
        Case 4 'Blue
            If frmConfiguration.tmrFinishedBeep.Enabled = False Then
                Call SetDO(gblngDO_AlarmRed, False)
            End If
            Call SetDO(gblngDO_AlarmYellow, False)
            Call SetDO(gblngDO_AlarmGreen, False)
            SetAlarmStatus = SetDO(gblngDO_AlarmBlue, True)
    End Select

End Function

Public Sub ControlMFC(sngGasValue() As Single)
    Dim sngMFCValue(6) As Single
    Dim lngMFCAOChannel(6) As Long
    Dim lngMFCDOChannel(6) As Long
    Dim lngGASDOChannel(6) As Long
    Dim i As Integer
    Dim j As Integer
    Dim lngOutputValue As Long
    Dim tmrStamp As Single
    Dim iRtn As Integer
    
    
    lngMFCAOChannel(0) = gblngAO_MFC1
    lngMFCAOChannel(1) = gblngAO_MFC2
    lngMFCAOChannel(2) = gblngAO_MFC3
    lngMFCAOChannel(3) = gblngAO_MFC4
    lngMFCAOChannel(4) = gblngAO_MFC5
    lngMFCAOChannel(5) = gblngAO_MFC6
    lngMFCDOChannel(0) = gblngDO_MFC_ValveServo1
    lngMFCDOChannel(1) = gblngDO_MFC_ValveServo2
    lngMFCDOChannel(2) = gblngDO_MFC_ValveServo3
    lngMFCDOChannel(3) = gblngDO_MFC_ValveServo4
    lngMFCDOChannel(4) = gblngDO_MFC_ValveServo5
    lngMFCDOChannel(5) = gblngDO_MFC_ValveServo6
    lngGASDOChannel(0) = gblngDO_GasValve1
    lngGASDOChannel(1) = gblngDO_GasValve2
    lngGASDOChannel(2) = gblngDO_GasValve3
    lngGASDOChannel(3) = gblngDO_GasValve4
    lngGASDOChannel(4) = gblngDO_GasValve5
    lngGASDOChannel(5) = gblngDO_GasValve6
    
    For i = 0 To gbintMaxGasEnable
        If gbintGasEnable(i) > 0 And gbsngMaxGasSLMP(i) > 0 Then sngMFCValue(i) = 5 * sngGasValue(i) / gbsngMaxGasSLMP(i)
        If sngMFCValue(i) > 5 Then sngMFCValue(i) = 5
        Kernel.sngCurrOutMFC(i) = sngGasValue(i)
        
        If sngGasValue(i) > 0 Then
            If gbsngGasError(i) <> 0 And frmDiagnosis.tmrCheckMFC(i).Enabled = False Then
                gbsngGasErrorC(i) = 0
                frmDiagnosis.tmrCheckMFC(i).Enabled = True
            End If
        
        End If
    Next i
    
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        For i = 0 To gbintMaxGasEnable
            lngOutputValue = 0
            If lngGASDOChannel(i) >= 0 Then
                If sngMFCValue(i) > 0 And _
                   DAQ91141.lngDO(lngGASDOChannel(i)) = 0 Then
                    Call SetDO(lngGASDOChannel(i), True)
                ElseIf sngMFCValue(i) = 0 And _
                   DAQ91141.lngDO(lngGASDOChannel(i)) = 1 Then
                    Call SetDO(lngGASDOChannel(i), False)
                End If
            End If
        Next i
    Else
        For i = 0 To gbintMaxGasEnable
            lngOutputValue = 0
            If lngGASDOChannel(i) >= 0 Then
                If sngMFCValue(i) > 0 And _
                   AvDaq17101.lngDO(lngGASDOChannel(i)) = 0 Then
                   Call SetDO(lngGASDOChannel(i), True)
                ElseIf sngMFCValue(i) = 0 And _
                   AvDaq17101.lngDO(lngGASDOChannel(i)) = 1 Then
                    Call SetDO(lngGASDOChannel(i), False)
                End If
            End If
        Next i
    End If
      
    For i = 0 To gbintMaxGasEnable
        If sngMFCValue(i) > 0 Then

            If lngMFCAOChannel(i) < 16 Then
                Call Ixud_WriteAOVoltage(gbintPIODA8A, lngMFCAOChannel(i), sngMFCValue(i)) '0~5V
            
            End If
        ElseIf sngMFCValue(i) = 0 Then
            If lngMFCAOChannel(i) < 16 Then
                
                iRtn = Ixud_WriteAOVoltage(gbintPIODA8A, lngMFCAOChannel(i), 0) '0~5V
                        
            End If
        End If
    Next i

End Sub

Public Sub VarControlMFC(Port As Integer, sngGasVolt As Single)
    Dim sngMFCValue As Single
    Dim lngMFCAOChannel(6) As Long
    Dim lngMFCDOChannel(6) As Long
    Dim lngGASDOChannel(6) As Long
    Dim i As Integer
    Dim j As Integer
    Dim lngOutputValue As Long
    Dim tmrStamp As Single
    Dim iRtn As Integer
    Dim intSA As Integer
    Dim intCH As Integer
    Dim intRet As Integer
    Dim C As Integer
    Dim VType As Integer
    
    lngMFCAOChannel(0) = gblngAO_MFC1
    lngMFCAOChannel(1) = gblngAO_MFC2
    lngMFCAOChannel(2) = gblngAO_MFC3
    lngMFCAOChannel(3) = gblngAO_MFC4
    lngMFCAOChannel(4) = gblngAO_MFC5
    lngMFCAOChannel(5) = gblngAO_MFC6
    lngMFCDOChannel(0) = gblngDO_MFC_ValveServo1
    lngMFCDOChannel(1) = gblngDO_MFC_ValveServo2
    lngMFCDOChannel(2) = gblngDO_MFC_ValveServo3
    lngMFCDOChannel(3) = gblngDO_MFC_ValveServo4
    lngMFCDOChannel(4) = gblngDO_MFC_ValveServo5
    lngMFCDOChannel(5) = gblngDO_MFC_ValveServo6
    lngGASDOChannel(0) = gblngDO_GasValve1
    lngGASDOChannel(1) = gblngDO_GasValve2
    lngGASDOChannel(2) = gblngDO_GasValve3
    lngGASDOChannel(3) = gblngDO_GasValve4
    lngGASDOChannel(5) = gblngDO_GasValve5
    lngGASDOChannel(6) = gblngDO_GasValve6

    Call SetDO(lngGASDOChannel(Port - 1), True)
    If Para.RtaType = 1 Or Para.RtaType = 3 Then
        Call Ixud_WriteAOVoltage(gbintPIODA8A, lngMFCAOChannel(Port - 1), sngGasVolt) '0~5V
    End If
    If Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 9 Then
        If CH2FR(lngMFCAOChannel(Port - 1), intSA, intCH) = True Then
            VType = &H34
            If gbintAO_Type(lngMFCAOChannel(Port - 1)) = 1 Then VType = &H32
            
            For C = 1 To 3
                intRet = FRB_WriteAOFloat(0, intSA, intCH, 12, VType, sngGasVolt, C)    '0~5V
                If intRet = 0 Then Exit For
                DelayTime (1)
            Next C
            'Call FRB_WriteAOFloat(0, intSA, intCH, 12, &H34, sngGasVolt, 2)    '0~5V
        End If
    End If
End Sub

Public Sub ResetMFC()
    Dim sngMFCValue(6) As Single
    Dim lngMFCAOChannel(6) As Long
    Dim lngMFCDOChannel(6) As Long
    Dim lngGASDOChannel(6) As Long
    Dim i As Integer
    
    lngMFCAOChannel(0) = gblngAO_MFC1
    lngMFCAOChannel(1) = gblngAO_MFC2
    lngMFCAOChannel(2) = gblngAO_MFC3
    lngMFCAOChannel(3) = gblngAO_MFC4
    lngMFCAOChannel(4) = gblngAO_MFC5
    lngMFCAOChannel(5) = gblngAO_MFC6
    lngMFCDOChannel(0) = gblngDO_MFC_ValveServo1
    lngMFCDOChannel(1) = gblngDO_MFC_ValveServo2
    lngMFCDOChannel(2) = gblngDO_MFC_ValveServo3
    lngMFCDOChannel(3) = gblngDO_MFC_ValveServo4
    lngMFCDOChannel(4) = gblngDO_MFC_ValveServo5
    lngMFCDOChannel(5) = gblngDO_MFC_ValveServo6
    lngGASDOChannel(0) = gblngDO_GasValve1
    lngGASDOChannel(1) = gblngDO_GasValve2
    lngGASDOChannel(2) = gblngDO_GasValve3
    lngGASDOChannel(3) = gblngDO_GasValve4
    lngGASDOChannel(4) = gblngDO_GasValve5
    lngGASDOChannel(5) = gblngDO_GasValve6
    
    For i = 0 To 5
        If lngGASDOChannel(i) >= 0 Then
            If DAQ91141.lngDO(lngGASDOChannel(i)) = 1 Then
                Call SetDO(lngGASDOChannel(i), False)
            End If
        End If
        If lngMFCDOChannel(i) >= 0 Then
            If DAQ91141.lngDO(lngMFCDOChannel(i)) = 1 Then
                Call SetDO(lngMFCDOChannel(i), False)
            End If
        End If
        If lngMFCAOChannel(i) < 16 Then
            Call Ixud_WriteAOVoltage(gbintPIODA8A, lngMFCAOChannel(i), 0)
        End If
    Next i
    
    frmDiagnosis.chkGasN2.value = 0
    
End Sub

Public Function SetLampCooling(blnIsOn As Boolean) As Boolean
   Dim bRet As Boolean

   bRet = SetDO(gblngDO_ValveCDA, blnIsOn)
   SetLampCooling = bRet
   
End Function

'120713 Josh
Public Function SetDoor(ActNo As Integer) As Boolean
    Dim Timeout As Long
    Dim tmrStamp As Single
    Dim tmrStampNow As Single
    Dim i As Long
    Timeout = 30
    tmrStamp = Timer
    tmrStampNow = Timer
    
    Select Case ActNo
        Case 0
            SetDO lngDO_DoorOpenValve, False
            SetDO lngDO_DoorCloseValve, True
            Kernel.IsDoorMoving = 1
            For i = 0 To 10000
            Next i
            While tmrStampNow - tmrStamp < Timeout And tmrStampNow - tmrStamp >= 0
                'ReadDI
                DoEvents
                If SysDI.IsDoorClose = 1 Then
                    Kernel.IsDoorMoving = 0
                    SetDO lngDO_DoorClamp, True
                    SetDoor = True
                    Exit Function
                End If
                tmrStampNow = Timer
            Wend
            SetDoor = False
            Exit Function
        Case 1
            If SysDI.IsChamberGaugeL = 0 Then
                ShowAlarmFlash 5
                SetDoor = False
                Exit Function
            Else
                SetDO lngDO_DoorClamp, False
                If lngDO_DoorClamp >= 0 Then
                    If gblngDI_DoorClamp >= 0 Then
                        While tmrStampNow - tmrStamp < Timeout And tmrStampNow - tmrStamp >= 0
                            DoEvents
                            If SysDI.IsDoorClamp = 1 Then
                                Timeout = 0
                            End If
                            tmrStampNow = Timer
                        Wend
                        If Timeout > 0 Then
                            SetDoor = False
                            Exit Function
                        End If
                    Else
                        For i = 0 To 10000
                        Next i
                    End If
                End If
                Sleep (3000)
                tmrStamp = Timer
                tmrStampNow = Timer
                Timeout = 30
                SetDO lngDO_DoorOpenValve, True
                SetDO lngDO_DoorCloseValve, False
                Kernel.IsDoorMoving = 1
                While tmrStampNow - tmrStamp < Timeout And tmrStampNow - tmrStamp >= 0
                    'ReadDI
                    DoEvents
                    If SysDI.IsDoorOpen = 1 And SysDI.IsDoorClose = 0 Then

                        For i = 0 To 10000
                        Next i
                        ReadDI
                        If SysDI.IsDoorOpen = 1 And SysDI.IsDoorClose = 0 Then
                            Kernel.IsDoorMoving = 0
                            SetDoor = True
                            Exit Function
                        Else
                            Kernel.IsDoorMoving = 0
                            SetDoor = False
                            Exit Function
                        End If
                    End If
                    tmrStampNow = Timer
                    
                Wend
                SetDoor = False
                Exit Function
            End If
        Case 2
            Kernel.IsDoorMoving = 0
            SetDO lngDO_DoorOpenValve, False
            SetDO lngDO_DoorCloseValve, False
            Exit Function
    End Select
End Function



Public Sub SetPump(IsON As Boolean)
    Dim bRet As Boolean
    bRet = SetDO(gblngDO_PumpPower, IsON)
    
    SysDO.IsPumping = IIf(IsON = True, 1, 0)
End Sub

Public Function SetAngle(IsON As Boolean) As Integer
    If IsON Then
        If SysDI.IsChamberGaugeL = 1 And Kernel.sngPressure > 700 Then
            SetDO gblngDO_AngleValve, IsON
        Else
            If SysDO.IsAngle = 0 Then
                SetAngle = -1
                ShowAlarmFlash 5
                Exit Function
            End If
        End If
    Else
        SetDO gblngDO_AngleValve, False
    End If
    
    SetAngle = 1
    SysDO.IsAngle = IIf(IsON = True, 1, 0)
End Function

Public Sub SetRelease(IsON As Boolean)
    Dim bRet As Boolean
    
    If gblngDO_ReleaseValve >= 0 Then
        bRet = SetDO(gblngDO_ReleaseValve, IsON)
    End If
End Sub

Public Sub AverageValue(intIndex As Integer, sngValue As Single)
    
    If intIndex >= 0 And intIndex < 32 Then
        If intCountAI(intIndex) >= 10 Then
            SysAI.AvgValue(intIndex) = sngSumAI(intIndex) / 10
            SysAI.AvgValue(intIndex) = SysAI.AvgValue(intIndex) + SysAI.ErrorV(intIndex)
            intCountAI(intIndex) = 0
            sngSumAI(intIndex) = 0
        Else
           intCountAI(intIndex) = intCountAI(intIndex) + 1
           sngSumAI(intIndex) = sngSumAI(intIndex) + sngValue
        End If
    End If
   
End Sub

Public Function AI2Vacuum(sngValue As Single) As Single
    Dim Vac As Single
    Dim vacuum As Single
    If Para.sngGaugeD = 0 Then
        If sngValue > 10 Then
'           Vac = 760
           Vac = 10 ^ ((sngValue - VACUUM_TORR_CNT_CT) / VACUUM_TORR_CNT_K)
        Else
            Vac = 10 ^ ((sngValue - VACUUM_TORR_CNT_CT) / VACUUM_TORR_CNT_K)
        End If
    Else
        If sngValue > 10 Then
            Vac = 760
        Else
            Vac = 10 ^ (1.667 * sngValue - Para.sngGaugeD)
        End If
    End If
    
    If Vac < 0 Then
        Vac = 0
    ElseIf Vac > 760 Then
        Vac = 760
    End If
    If gbsngVacuumGaugeCompensation <> 0 Then
        Vac = Vac + gbsngVacuumGaugeCompensation
        If Vac < 0 Then
        Vac = Vac - gbsngVacuumGaugeCompensation
        End If
    End If
    AI2Vacuum = Vac
End Function

Public Function AI2MFC(sngValue As Single, sngMax As Single, sngOffset As Single) As Single
    Dim mfc As Single
        
    mfc = (sngValue - sngOffset) * sngMax / 5
    If mfc < 0 Then
        mfc = 0
    ElseIf mfc > sngMax Then
        mfc = sngMax
    End If
    AI2MFC = mfc
        
End Function
Public Function AI2Oxygen(sngValue As Single) As Single
    Dim mfc As Single
        
    mfc = (sngValue) * 100
    If mfc < 0 Then
        mfc = 0
    ElseIf mfc > 100 Then
        mfc = 100
    End If
    AI2Oxygen = mfc
        
End Function

'Public Function Delay(sngPause As Single)
'    Dim tmrStamp As Single
'
'    tmrStamp = Timer
'    While Timer - tmrStamp < sngPause And Timer - tmrStamp >= 0
'        'DoEvents
'    Wend
'End Function


Public Sub SendCmdPM(id As Integer)
    Dim i As Integer
    Dim i2 As Integer
    Dim iStart As Integer
    Dim OutByte() As Byte, tmpByte(0) As Byte

    iStart = id
    
'    If Para.RtaType = 9 Then        '2 Big Box + 3 small Box
'        If id >= 0 And id <= 7 Then
'            CmdData(1, 0) = 1
'        ElseIf id >= 8 And id <= 15 Then
'            CmdData(1, 0) = 2
'            iStart = iStart - 8
'        Else
'            If id = 16 Then CmdData(1, 0) = 3
'            If id = 17 Then CmdData(1, 0) = 4
'            If id = 18 Then CmdData(1, 0) = 5
'            iStart = 0
'        End If
'    Else                            '1 Big Box + 3 small Box
'        If id >= 0 And id <= 7 Then
'        CmdData(1, 0) = 1
'        Else
'            If id = 8 Then CmdData(1, 0) = 2
'            If id = 9 Then CmdData(1, 0) = 3
'            If id = 10 Then CmdData(1, 0) = 4
'            iStart = 0
'        End If
'    End If
    
    If Para.intPMbig = 1 Then
        If id >= 0 And id <= 7 Then
            CmdData(1, 0) = 1
        Else
            CmdData(1, 0) = id - 6
            iStart = 0
        End If
    
    Else
        If id >= 0 And id <= 7 Then
            CmdData(1, 0) = 1
        ElseIf id >= 8 And id <= 15 Then
            CmdData(1, 0) = 2
            iStart = iStart - 8
        Else
            CmdData(1, 0) = id - 13
            iStart = 0
        End If
    End If
    
    
    
    i2 = 8
    CmdData(1, 2) = &H11 + iStart
    For i = 0 To i2 - 1
        SendBuffer(i) = CmdData(1, i)
    Next
    
    Call Modbus_CRC16(i2 - 2)
    SendBuffer(i2 - 1) = CRC_Low
    SendBuffer(i2 - 2) = CRC_High
    
    For i = 0 To i2 - 1
        tmpByte(0) = SendBuffer(i)

        OutByte = tmpByte
        frmConfiguration.MSComm1.Output = OutByte
    Next
    
End Sub

Public Function ReadAz1() As Boolean
    Dim sngData(0 To 99) As Single
    Dim i As Integer
        
On Error GoTo ERRLINE
    
    For i = 0 To 3

        Az1.sngPV(i) = gbsngAz1Data(i) / (10 ^ gbintPrecisionDigit(i))
        Az1.sngMV(i) = gbsngAz1Data(i + 4) / 10
        gbsngPower(i) = Az1.sngMV(i) / 10
        Az1.blnStart(i) = IIf(gbsngAz1Data(i + 8) > 0, False, True)
        Az1.intMode(i) = gbsngAz1Data(i + 12)
        sngTemp_1(i) = Az1.sngPV(i)
    Next i

    If Para.intOnlyRecipe = 1 Then
        For i = 0 To 3
            If Kernel.IsRun = 1 Then
                Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                                gbsngPowerTC(i) * sngTemp_1(i) ^ 2 + gbsngPower3C(i) * sngTemp_1(i) ^ 3 + _
                                gbsngPower4C(i) * sngTemp_1(i) ^ 4 + gbsngPower5C(i) * sngTemp_1(i) ^ 5
'                 Kernel.sngTC(i) = sngTemp_1(i)

            Else
                Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
'                Kernel.sngTC(i) = sngTemp_1(i)
            End If

            Kernel.sngOrigTC(i) = sngTemp_1(i)
        Next i
    Else
        For i = 0 To 3
            Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                            gbsngPowerTC(i) * sngTemp_1(i) ^ 2 + gbsngPower3C(i) * sngTemp_1(i) ^ 3 + _
                            gbsngPower4C(i) * sngTemp_1(i) ^ 4 + gbsngPower5C(i) * sngTemp_1(i) ^ 5
'            Kernel.sngTC(i) = sngTemp_1(i)
            Kernel.sngOrigTC(i) = sngTemp_1(i)
        Next i
    End If

    
    ReadAz1 = True
    
    Exit Function
ERRLINE:
    gbstrAlarmHint = " ReadAz1 error"
    ShowAlarmFlash 1
End Function

Public Function ReadAz2() As Boolean
    Dim sngData(0 To 99) As Single
    Dim i As Integer
        
On Error GoTo ERRLINE
    
    
    For i = 0 To 3
'        Az2.sngPV(i) = gbsngAz2Data(i) / 10
        Az2.sngPV(i) = gbsngAz2Data(i) / (10 ^ gbintPrecisionDigit(i + 4))
        Az2.sngMV(i) = gbsngAz2Data(i + 4) / 10
        Az2.blnStart(i) = IIf(gbsngAz2Data(i + 8) > 0, False, True)
        Az2.intMode(i) = gbsngAz2Data(i + 12)
        sngTemp_1(i + 4) = Az2.sngPV(i)
    Next i
     gbsngPower(4) = Az2.sngMV(0) / 10
    If Para.intOnlyRecipe = 1 Then
        For i = 4 To 7
            If Kernel.IsRun = 1 Then
                Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                                gbsngPowerTC(i) * sngTemp_1(i) ^ 2 + gbsngPower3C(i) * sngTemp_1(i) ^ 3 + _
                                gbsngPower4C(i) * sngTemp_1(i) ^ 4 + gbsngPower5C(i) * sngTemp_1(i) ^ 5
'                Kernel.sngTC(i) = sngTemp_1(i)
            Else
                Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioEX(i) + gbsngErrorTC(i)
'                Kernel.sngTC(i) = sngTemp_1(i)
            End If

            Kernel.sngOrigTC(i) = sngTemp_1(i)
        Next i
    Else
        For i = 4 To 7
            Kernel.sngTC(i) = sngTemp_1(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngErrorTC(i) + _
                            gbsngPowerTC(i) * sngTemp_1(i) ^ 2 + gbsngPower3C(i) * sngTemp_1(i) ^ 3 + _
                            gbsngPower4C(i) * sngTemp_1(i) ^ 4 + gbsngPower5C(i) * sngTemp_1(i) ^ 5
'            Kernel.sngTC(i) = sngTemp_1(i)
            Kernel.sngOrigTC(i) = sngTemp_1(i)
        Next i
    End If
    Exit Function
ERRLINE:
    gbstrAlarmHint = " ReadAz2 error"
    ShowAlarmFlash 1
End Function


Public Function SetCover(IsON As Boolean) As Boolean
    Dim bRet As Boolean
    
    If SysDI.IsCoverServoRdy = 1 And SysDI.IsCoverOrigRdy = 1 Then
        bRet = SetDO(gblngDO_COVER_POS_01, IsON) 'Down=1 up=0
        bRet = SetDO(gblngDO_COVER_MOVE, False)
        Sleep 100
        bRet = SetDO(gblngDO_COVER_MOVE, True)
        SetCover = True
    Else
        gbstrAlarmHint = " Cover Servo error"
        ShowAlarmFlash 28
        SetCover = False
    End If
End Function

Public Sub TogglePumping(chkPumping As Object, useTPump As Boolean, tmrSetReleaseOFF As Object)
    If chkPumping.Enabled = False Then Exit Sub
    chkPumping.Enabled = False
    
    If Para.useTPump = False Then
        If chkPumping.value = 1 Then
            If SysDI.IsChamberGaugeL = 0 Then
                chkPumping.value = 0
                ShowMessageOK "腔體處於真空狀態,請先釋放壓力!"
                chkPumping.Enabled = True
                Exit Sub
            End If
                
            SetAngle True
            frmDiagnosis.tmrPumpON.Enabled = True
            Call frmHistory.AppendLogAlert(1, "Manual", 1012, "手動開啟泵浦", 1)
            PumpState = True
        Else
            SetAngle False
            frmDiagnosis.tmrPumpON.Enabled = False
            frmDiagnosis.tmrPumpOFF.Enabled = True
            Call frmHistory.AppendLogAlert(1, "Manual", 1013, "手動關閉泵浦", 1)
            PumpState = False
        End If
    Else
        If chkPumping.value = 1 Then
            If SysDI.IsChamberGaugeL = 0 Then
                chkPumping.value = 0
                ShowMessageOK "腔體處於真空狀態,請先釋放壓力!"
                chkPumping.Enabled = True
                Exit Sub
            End If
                
            SetPump True
            Call frmHistory.AppendLogAlert(1, "Manual", 1012, "手動開啟泵浦", 1)
            PumpState = True
        Else
            SetPump False
            Call frmHistory.AppendLogAlert(1, "Manual", 1013, "手動關閉泵浦", 1)
            If gbintReleaseOpenDelay > 0 Then tmrSetReleaseOFF.Enabled = True
            PumpState = False
        End If
    End If
    chkPumping.Enabled = True
End Sub

