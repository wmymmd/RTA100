Attribute VB_Name = "mdlDeclare"
Option Explicit

'--------------------------------------------------------------------
'Define color
'--------------------------------------------------------------------
Public Const GB_ColorMode = 0   'mode 0->Blue 1->Red
Public Const GB_ColorBlack = &H80000008
Public Const GB_ColorLightRed = &H8080FF
Public Const GB_ColorRed = &HFF&
Public Const GB_ColorHeavyRed = &HC0&
Public Const GB_ColorHeavyBlue = &H800000
Public Const GB_ColorMuchBlue = &HA56E32
Public Const GB_ColorFewBlue = &HE0C674
Public Const GB_ColorLightBlue = &HFFC0C0
Public Const GB_ColorFewLightBlue = &HFFFF00
Public Const GB_ColorBlue = &HC00000
Public Const GB_ColorHeavyPurple = &H800080
Public Const GB_ColorPurple = &HFF00FF
Public Const GB_ColorLightPurple = &HFF80FF
Public Const GB_ColorLightGray = &HE0E0E0
Public Const GB_ColorGray = &H8000000F
Public Const GB_ColorHeavyGray = &HC0C0C0
Public Const GB_ColorPlusGray = &H808080
Public Const GB_ColorPinkPurple = &HFFC0FF
Public Const GB_ColorGreen = &HC000&
Public Const GB_ColorLightGreen = &HFF00&
Public Const GB_ColorVeryLightGreen = &H80FF80
Public Const GB_ColorHeavyGreen = &H8000&
Public Const GB_ColorVeryHeavyGreen = &H4000&
Public Const GB_ColorCyan = &HC0C000
Public Const GB_ColorLightCyan = &HFFFF00
Public Const GB_ColorVeryLightCyan = &HFFFF80
Public Const GB_ColorHeaveyCyan = &H808000
Public Const GB_ColorVeryHeaveyCyan = &H404000
Public Const GB_ColorBrown = &H4080&
Public Const GB_ColorOrange = &H80FF&

'--------------------------------------------------------------------
'Define recipe action type
'--------------------------------------------------------------------
Public Const GB_ACTION_IDLE = "Idle"
Public Const GB_ACTION_PREHEAT = "PreHeat"
Public Const GB_ACTION_RAMPUP = "Ramp up"
Public Const GB_ACTION_RAMPDOWN = "Ramp Down"
Public Const GB_ACTION_HOLD = "Hold"
Public Const GB_ACTION_VENT = "Vent"
Public Const GB_ACTION_STOP = "Stop"
Public Const GB_ACTION_PURGE = "Purge"
Public Const GB_ACTION_COOLING = "Cooling"
Public Const GB_ACTION_PUMPDOWN = "Pump Down"
Public Const GB_ACTION_PUMPDOWNKEEP = "Pump Keep"
Public Const GB_ACTION_MANUALPUMP = "Manual Pressure"
Public Const GB_ACTION_IOCONTROL = "IO Control"

Public Const GB_ACTION_INDEX_IDLE = 1
Public Const GB_ACTION_INDEX_PREHEAT = 2
Public Const GB_ACTION_INDEX_RAMPUP = 3
Public Const GB_ACTION_INDEX_RAMPDOWN = 4
Public Const GB_ACTION_INDEX_HOLD = 5
Public Const GB_ACTION_INDEX_VENT = 6
Public Const GB_ACTION_INDEX_STOP = 0
Public Const GB_ACTION_INDEX_PURGE = 8
Public Const GB_ACTION_INDEX_COOLING = 9
Public Const GB_ACTION_INDEX_PUMPDOWN = 10
Public Const GB_ACTION_INDEX_PUMPDOWNKEEP = 11
Public Const GB_ACTION_INDEX_MANUALPUMP = 12
'--------------------------------------------------------------------
' Recipe condition constant
'--------------------------------------------------------------------
Public Const GB_MAX_RAMPUP_COUNT = 50
Public Const GB_MAX_ACTION_TYPE = 9
Public Const GB_MAX_STEP_PROCESS = 50
Public Const GB_MAX_GAS_NUMBER = 3
Public Const GB_MAX_LOOPS = 6
Public Const GB_MAX_DRAW_COL = 50
'--------------------------------------------------------------------
' Recipe process type index
'--------------------------------------------------------------------
Public Const GB_PROCESS_STEP = 0
Public Const GB_PROCESS_ACTION = 1
Public Const GB_PROCESS_TEMP = 2
Public Const GB_PROCESS_TIME = 3
Public Const GB_PROCESS_GAS1 = 4
Public Const GB_PROCESS_GAS2 = 5
Public Const GB_PROCESS_GAS3 = 6
Public Const GB_PROCESS_GAS4 = 7
Public Const GB_PROCESS_GAS5 = 8
Public Const GB_PROCESS_GAS6 = 9
'Public Const GB_PROCESS_N2 = 4
'Public Const GB_PROCESS_N2H2 = 5
'Public Const GB_PROCESS_O2 = 6


'--------------------------------------------------------------------
'Define DaqCard
'--------------------------------------------------------------------
Public Const GB_AO_RANGE = 10
Public Const GB_AI_RANGE = 10
Public Const GB_DAQ_BASE = 32727
Public Const GB_AO_RATIO = GB_DAQ_BASE / GB_AO_RANGE
Public Const GB_AI_RATIO = GB_DAQ_BASE / GB_AI_RANGE
'--------------------------------------------------------------------
'Define Element Convert Scale(Ratio)
'--------------------------------------------------------------------
Public Const GB_TC_CVT_SCALE = 100
Public Const GB_PM_CVT_SCALE = 1389 / 5 - 18
Public Const GB_VACUUM_TORR_CNT_CT = 6.304
'Public Const GB_VACUUM_TORR_CNT_K = 1.286
Public Const GB_VACUUM_TORR_CNT_K = 1.282

'--------------------------------------------------------------------
' Pumping sequence status
'--------------------------------------------------------------------
Public Const GB_PUMPING_SLOWTIME = 0
Public Const GB_PUMPING_FAST = 1
Public Const GB_PUMPING_SLOW = 2
Public Const GB_PUMPING_STOP = 3

'--------------------------------------------------------------------
'Status bar status
'--------------------------------------------------------------------
Public Const GB_STATUSBAR_MODE_MANUAL = 1
Public Const GB_STATUSBAR_MODE_AUTO = 2
Public Const GB_STATUSBAR_NORMAL = 11
Public Const GB_STATUSBAR_WARNING = 12
Public Const GB_STATUSBAR_PROCESS = 13
Public Const GB_STATUSBAR_ALARM = 14

Public Const GB_STATUSBAR_IO_DOOR_OPEN = 101
Public Const GB_STATUSBAR_IO_DOOR_CLOSE = 102
Public Const GB_STATUSBAR_READY = 11
Public Const GB_STATUSBAR_NOTREADY = 12

Public Const GB_STATUS_READY = 1
Public Const GB_STATUS_ALARM = 2
Public Const GB_STATUS_DR = 3
Public Const GB_STATUS_OH = 4
Public Const GB_STATUS_VAC = 5


Public Const GB_SCR_MAX = 17
Public Const GB_GAS_MAX = 6

Global gbintCurrStat As Integer
Public Const GB_SYS_NRDY = 0
Public Const GB_SYS_RDY = 1
Public Const GB_SYS_ALM = 2

'--------------------------------------------------------------------
'Define EQ Status
'--------------------------------------------------------------------
'***System status***

Public gbblnChamberOverheat As Boolean
Public gbblnCT(GB_SCR_MAX) As Boolean
Public gbblnDoorOpen As Boolean
Public gbblnDoorClose As Boolean
Public gbblnDoorOpenSwitch As Boolean
Public gbblnDoorCloseSwitch As Boolean
Public gbblnPowerInput24 As Boolean
'Public gbblnPowerInput220 As Boolean
Public gbblnEMS_Alarm As Boolean
Public gbblnMC_Input  As Boolean
Public gbblnSystemAlarm As Boolean
Public gbblnMaxTemperatureAlarm As Boolean

Public gbblnCTActive(GB_SCR_MAX) As Boolean
Public gbblnDoorOpenActive As Boolean
Public gbblnDoorCloseActive As Boolean
Public gbblnChamberOverheatActive As Boolean
Public gbblnPowerInput24Active As Boolean
'Public gbblnPowerInput220Active As Boolean
Public gbblnEMS_AlarmActive As Boolean
Public gbblnMC_InputActive  As Boolean
Public gbblnSystemAlarmActive As Boolean
Public gbblnMaxTemperatureAlarmActive As Boolean

'120713 Josh
Public gbblnArmInFront  As Boolean
Public gbblnArmInRear  As Boolean
Public gbblnPlayFakeBall As Boolean
Public gbintPlayFakeBall As Integer
'***The Heating Module***
Public gbsngIntensity  As Single
'120713 Josh
Public gbsngCurrIntensity(GB_SCR_MAX)  As Single

Public gbsngPM  As Single
Public gbsngCTValue(7) As Single

Public gbIsCT1 As Boolean
Public gbIsCT2 As Boolean
Public gbIsCT3 As Boolean
Public gbIsCT4 As Boolean
Public gbIsCT5 As Boolean
Public gbIsCT6 As Boolean

Public gbintTCO2 As Integer

Public gbsngPower(GB_SCR_MAX)  As Single
Public gbsngTemperature(GB_SCR_MAX)  As Single
Public gbsngTempInput(GB_SCR_MAX)  As Single
Public gbsngInputTCMap(GB_SCR_MAX)  As Single
Public gbstrNameTC(30)  As String
Public gbsngPowerTC(30)  As Single
Public gbsngErrorTC(30)  As Single
Public gbsngRatioTC(30)  As Single
Public gbsngRatioEX(30)  As Single
Public gbsngPower3C(30)  As Single
Public gbsngPower4C(30)  As Single
Public gbsngPower5C(30)  As Single
Public gbintLoopNo(30)  As Integer
Public gbintPrecisionDigit(30)  As Integer

Public gbChamberNo As String

'***The Gas Module***
Public gbsngGas(GB_GAS_MAX) As Single

'***The Cooling Module***
Public gbblnWaterGate As Boolean
Public gbblnAirGate As Boolean
Public gbblnWaterGateActive As Boolean
Public gbblnAirGateActive As Boolean

Public gbblnGaugeGate As Boolean
Public gbblnGaugeGateR As Boolean
Public gbblnGaugeGateL As Boolean
Public gbblnChamberGaugeGateR As Boolean
Public gbblnChamberGaugeGateL As Boolean

Public gbblnSwitchGate As Boolean
Public gbblnGaugeGateActive As Boolean
Public gbblnGaugeGateRActive As Boolean
Public gbblnGaugeGateLActive As Boolean
Public gbblnSwitchGateActive As Boolean
Public gbblnChamberGaugeGateActive As Boolean
Public gbblnChamberGaugeGateRActive As Boolean
Public gbblnChamberGaugeGateLActive As Boolean

Public gbsngOxygenPPM As Single

'--------------------------------------------------------------------
'Define EQ Configuration
'--------------------------------------------------------------------
'Heating
Public gbsngSCRAddress(GB_SCR_MAX) As Byte
Public gbsngIntensityWeight(GB_SCR_MAX) As Single
Public gbsngIntensityWeightS(GB_SCR_MAX) As Single
Public gbsngPropertyCoefficient(GB_SCR_MAX) As Single
Public gbsngIntensityKeep As Single
Public gbintPreheatIntensity As Single
Public gbsngMaxTemperature As Single
Public gbsngMinTemperature As Single
Public gbsngOpenTemperature As Single
Public gbsngChamberOverheat As Single
Public gbsngValidPMTempature As Single
Public gbsngPMCVT1 As Single
Public gbsngPMCVT2 As Single
Public gbsngTCCVT1 As Single
Public gbsngTCCVT2 As Single
Public gbsngTCDifferentialRange As Single
Public gbblnResetInteral As Boolean
Public gbintRampSmooth As Integer
Public gbintSmoothDisplay As Integer
Public gbsngSmoothTime As Single
Public gbsngLifeLamp As Single
Public gbsngUsedLamp As Single
Public gbsngMaxMonitorError As Single
Public gbsngMaxMonitorTime As Single
Public gbintNumOfBanks As Integer
Public gbstrLogFilePath As String
Public gbstrRecipeFilePath As String
Public gbblnMessageShowed As Boolean
Public gblngPreStatus As Integer
Public gbsngGaugeZoomIn As Single
Public gbsngAPCInterval As Single
Public gbsngAPC_P As Single
Public gbsngAPC_I As Single
Public gbsngKeepPurge As Single
Public gbintAPC_MFC_Port As Integer
Public gbintMFC_Ratio As Integer
Public gbintRtaType As Integer
Public gbintAutoDeleteRecipe As Integer
Public gbintRobotPort As Integer
Public gbintRobotSpeed As Integer
Public gbsngPickH As Single
Public gbsngPlaceH As Single
Public gbsngTeach(100, 5) As Single

    
'Rev 12.0.0.2
Public gbsngIntensityLimit As Single
Public gbintFinishedBeep As Integer
Public gbintFinishedLight As Integer
Public gbintCycleRun As Integer
Public gbsngGaugeValue As Single

'120822 Josh
Public gbintLoginRight As Integer
Public gbstrAdminPW As String
Public gbstrEngineerPW As String
Public gbstrOperatorPW As String


Public gbsngUniformityStartPointHold As Single
Public gbsngUniformitySubWeight1 As Single
Public gbsngUniformitySubWeight2 As Single
Public gbsngUniformitySubWeightD1 As Single
Public gbsngUniformitySubWeightD2 As Single

Public gbsngUniformitySubWeightA As Single 'Rev4.1.6
Public gbsngUniformitySubWeightB As Single 'Rev4.1.6

'--------------------------------------------------------------------
' CT Macro
'--------------------------------------------------------------------
Public Const GB_CT_MAX = 6

Public gbsngIntensityRef(2) As Single
Public gbsngCTGate1(10) As Single
Public gbsngCTGate2(10) As Single
Public gbsngCTGateSlope(10) As Single
Public gbsngCTGateWeight As Single

'Vacuum
Public gbintPumpTimeout As Integer
Public gbsngPumpDownGate As Single
Public gbsngPumpingDelay As Single
Public gbsngVacuumGaugeCompensation As Single
Public gbintAngleOpenDelay As Integer
Public gbintReleaseOpenDelay As Integer
Public gbintThrottleFullOpenDelay As Integer
Public gbintThrottleInitialPos  As Integer
Public gbsngAPCGaugeValveLimit  As Single
Public gbsngAPCGaugePressureValue  As Single
Public gbsngAPC_FullScale As Single
Public gbsngVentGate As Single 'for open angle valve

Public gbIsPumpInStep As Boolean

'--------------------------------------------------------------------
' Gas Macro
'--------------------------------------------------------------------

Public gbstrGasPrecision(GB_GAS_MAX) As String

'Gas
Public gbintMaxGasEnable As Integer
Public gbintGasEnable(GB_GAS_MAX) As Integer
Public gbstrGasAlias(GB_GAS_MAX) As String
Public gbstrGasUnit(GB_GAS_MAX) As String
Public gbsngMaxGasSLMP(GB_GAS_MAX) As Single
Public gbsngGasBias(GB_GAS_MAX) As Single
Public gbsngGasError(GB_GAS_MAX) As Single
Public gbsngGasErrorN(GB_GAS_MAX) As Single
Public gbsngGasErrorC(GB_GAS_MAX) As Single
Public gbsngGasFlowScale(GB_GAS_MAX) As Single

'Barcode id
Public gbstrPN As String
Public gbstrBN As String
Public gbstrID1 As String
Public gbstrID2 As String

'System
Public gbintLampMonitor As Integer
Public gbintAlarmBuzzer As Integer
Public gbintMonitorTC As Integer

Public gbintMonitorTCActive(10) As Integer
Public gbintErrorTCActive(10) As Integer
Public gbintActivePage(10) As Integer

'--------------------------------------------------------------------
'Define Alarm Activity Enable/Disable
'--------------------------------------------------------------------
'***System status***
Public gbintActiveAlarm_System As Integer
Public gbintActiveAlarm_Ready As Integer
Public gbintActiveAlarm_DC As Integer
Public gbintActiveAlarm_EMS As Integer
Public gbintActiveAlarm_Buzzer As Integer

'***The Heating Module***
Public gbintActiveAlarm_Overheat As Integer
Public gbintActiveAlarm_CT As Integer
Public gbintActiveAlarm_TC As Integer

'***The Gas Module***
'***The Cooling Module***
Public gbintActiveAlarm_Water As Integer
Public gbintActiveAlarm_Air As Integer
'***The Vacuum Module***
Public gbintActiveAlarm_VacuumGauge As Integer
Public gbintActiveAlarm_VacuumSwitch As Integer
Public gbintActiveAlarm_VacuumGateR As Integer
Public gbintActiveAlarm_VacuumGateL As Integer
Public gbintActiveAlarm_OxygenGauge As Integer
'Misc
Public gbintActiveAlarm_Door As Integer
Public gbintActiveAlarm_APC As Integer
Public gbintActiveAlarm_TcWafer As Integer

Public gbintActiveAlarm_RValue As Integer
'--------------------------------------------------------------------
'Define Module Activity Enable/Disable
'--------------------------------------------------------------------
Public gbintActiveModule_Heating As Integer
Public gbintActiveModule_Cooling As Integer
Public gbintActiveModule_Gas As Integer
Public gbintActiveModule_Vacuum As Integer
Public gbintActiveModule_Facility As Integer
Public gbintActiveModule_Oxygen As Integer
Public gbintActiveModule_Barcode As Integer
Public gbintActiveModule_Door As Integer
Public gbintActiveModule_PNRecipe As Integer
Public gbintActiveModule_APC As Integer
Public gbintActiveModule_Auto As Integer
Public gbintActiveModule_Database As Integer
Public gbintActiveModule_CIM As Integer
Public gbintActiveModule_MLoop As Integer
'120702 Josh
Public gbintActiveMotion As Integer

Public gbintCTCheck As Integer
Public gbRValRange As Double


Public gbintActiveAlarm_ChamberGate As Integer
Public gbintActiveAlarm_ChamberGateR As Integer
Public gbintActiveAlarm_ChamberGateL As Integer

'--------------------------------------------------------------------
'Define System Path
'--------------------------------------------------------------------
Public gbSystemPath As String
Public gbSystemFile As String
'--------------------------------------------------------------------
'Define Thermal Filter
'--------------------------------------------------------------------
Public gbsngTCFilter(100)  As Single
Public gbintTCCount  As Integer
Public gbintReadTCCountMax  As Integer

Public gbdbTCFilterDegreeMap(100)  As Double
Public gbdbTCFilterMap(10, 100) As Double
Public gbdbTCFilterCount As Long



'--------------------------------------------------------------------
'Define Pyrometer Compensation Parameter
'--------------------------------------------------------------------
Public gbblnCompensationPM As Boolean
Public gbdblPMCoffForWafer() As Single
Public gbdblPMCoffForSPT() As Single
Public gbsngPMFilter(10)  As Single
Public gbintPMCount  As Single
Public gbintPMDetectObject As Integer
Public gbintCompensationCountWAF  As Integer
Public gbintCompensationCountSPT  As Integer

Public devicelist(0 To 255) As PT_DEVLIST
Public SubDevicelist(0 To 255) As PT_DEVLIST
Global ErrCde As Long
Global szErrMsg As String * 80



'-------------------------------------------------------------------
'Rev10.0.0.5
'-------------------------------------------------------------------
'120713 Josh
Public gbsngRecipeIntensityWeightSteady(GB_SCR_MAX) As Single
Public gbsngRecipeIntensityWeightDynamic(GB_SCR_MAX) As Single
Public gbsngRecipeCT(GB_SCR_MAX) As Single
Public gbsngRecipeCD(GB_SCR_MAX) As Single

Public gbsngRecipePinHeight As Single
Public gbblnRecipeStartAutoClose As Boolean
Public gbblnRecipeEndAutoOpen As Boolean
Public gbblnRecipeStartCloseCover As Boolean
Public gbblnRecipeEndOpenCover As Boolean
Public gbblnRecipeAutoCloseValve1 As Boolean
Public gbblnRecipeAutoCloseValve2 As Boolean
Public gbblnRecipeFinishedClear As Boolean
Public gbblnRecipeUseCT As Boolean
Public gbblnRecipeSaveLogCT As Boolean

Public gbblnAutoCloseValve As Boolean
Public gbintTCType  As Integer
Public gbintTCVoltageRange  As Integer
Public gbsngRecipeRampDownPower        As Single

Public cIni As New cInifile
Public cReportIni As New cInifile
Public gbintCurrReportIndex  As Integer

Public gbblnPNLoad As Boolean
Public gbstrPNRecipeFile  As String

Public lpDioWritePort As PT_DioWritePortByte
Public lpDioReadPort As PT_DioReadPortByte

Public gbsngIdleCount  As Single
Public gbsngIdleWarning  As Single

Public gbintValidDays  As Integer
Public gbintCurrDays  As Integer
Public gbblnNoModalForm As Boolean
Public gbblnShowHint As Boolean
Public gblngPumpDownTime  As Long
Public gblngCheckPC As Long
Public gbintTowerIndex As Integer
Public gbblnReceivedPM As Boolean
Public gbstrAlarmHint  As String

Public gbblnActiveHGD As Boolean
Public gbblnActiveHGA As Boolean
Public gbblnActiveHGU As Boolean
Public gbblnActiveHGF As Boolean
Public gbblnActiveAD As Boolean

Public gbblnGetDCR As Boolean
Public gbblnGetRecipe As Boolean
Public gbstrSend As String

Public gbblnSendRecipe As Boolean

Public DefineprocStep As Integer

Public OffsetWriteToTcm As Integer
Public StopTCM As Integer

Public CTNumbers As Integer
Public ForcePreheat As Integer
Public CTDisplay As Integer
Public CTNumber1 As Integer
Public CTNumber2 As Integer
Public CTNumber3 As Integer
Public CTNumber4 As Integer

Public CTName1 As String
Public CTName2 As String
Public CTName3 As String
Public CTName4 As String

Public CTOrder1 As String
Public CTOrder2 As String
Public CTOrder3 As String
Public CTOrder4 As String

Public IsUsedSCR As Integer
Public PortSCR As String

Public gbsngAz1Data(99) As Single
Public gbsngAz2Data(99) As Single
Public gbsngAz1Para(99) As Single
Public gbsngAz2Para(99) As Single
Public gbintAz1ProcNo As Integer
Public gbintAz2ProcNo As Integer

Public gbintCoverOrigCount As Integer

Public gbintRecipePrepareIndex            As Integer
Public gbsngRecipePrepareTimeout            As Long
Public gbsngRecipePrepareGaugeO2            As Single
Public gbsngRecipeTempDownTimeout            As Long
Public gbblnActivePrepare As Boolean
Public gbblnActiveTempDown As Boolean
Public gbintPreheatPower            As Integer
Public gbintPreheatTime            As Integer
Public gbsngRecipeGatePS1 As Single
Public gbsngRecipeGatePS2 As Single

Public Type ProcType
    strAction As String
    sngTime As Single
    sngTemperature As Single
    sngOutput As Single
    intStep As Integer
    intAction As Integer
    sngPump As Single
    sngGas(10) As Single
    lngCurrentTime As Long
    lngCurrStepTime As Long
    lngPrevTime As Long
    lngScanTime As Long
    lngOverTime(60) As Long
    lngUnderTime(60) As Long
    blnCheckOverCT As Boolean
    blnCheckUnderCT As Boolean
    blnDoStep As Boolean
    intRecordLoop As Integer
    blnOxygenTimeout As Boolean
    dblStartO2Flag As Double
    blnTempDownTimeout As Boolean
    dblStartTempDownFlag As Double
    strLogFilePath As String
    strLogFileName As String
    strUserID As String
    strCaseID As String
    strWaferID(25) As String
End Type

Public Type RectType
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Type DIType
    value(64) As Integer
    IsReady As Integer
    IsEMO As Integer
    IsDoorOpen As Integer
    IsDoorClose As Integer
    IsDoorClamp As Integer
    IsOverHeat As Integer
    IsCDA As Integer
    IsWater As Integer
    IsLampError(10) As Integer
    IsChamberGaugeH As Integer
    IsChamberGaugeL As Integer
    IsPumpGaugeH As Integer
    IsPumpGaugeL As Integer
    OverHeatCount As Integer
    LampErrorCount As Integer
    
    IsCoverAlarm As Integer
    IsCoverServoRdy As Integer
    IsCoverOrigRdy As Integer
    IsCoverMoving As Integer
    IsCoverUp As Integer
    IsCoverDown As Integer
End Type

Public Type DOType
    OnceValue As Long
    value(64) As Integer
    IsPumping As Integer
    IsAngle As Integer
    
End Type

Public Type AIType
    value(32) As Single
    AvgValue(32) As Single
    ErrorV(32) As Single
    sngMFC(10) As Single
    
End Type

Public Type AOType
    value(32) As Single
    sngMFC(20) As Single
    sngSCR(20) As Single
End Type

Public Type KernelType
    IsRun As Integer
    IsAlarm As Integer
    IsDoorMoving As Integer
    IsPurge As Integer
    IsNeedTestRun As Integer
    IsRemoteStart As Integer
    IsRemoteStop As Integer
    IsRemoteConnect As Integer
    IsTcpTempConnect As Integer
    IsPM As Integer
    IsPreHeat As Integer
    strCurrReportTime As String
    intCurrCycleRun As Integer
    intCurrMonitorRun As Integer
    strServerRecipe As String
    strBarcodeID As String
    strServerPath As String
    sngOrigTC(30)  As Single
    sngTC(30)  As Single
    sngIntensity  As Single
    intCurrActSCR(20) As Integer
    sngCurrOutSCR(20) As Single
    intCurrActMFC(20) As Integer
    sngCurrOutMFC(20) As Single
    strCurrStep As String
    intCurrStep As Integer
    lngCurrStepCount As Long
    intOpenDoorCount As Integer
    
    strCurrRecipe As String
    strCurrRecipeFile As String
    strCurrLogFile As String
    sngPressure As Single
    sngOxygen As Single
    intOverCT(60) As Integer
    intUnderCT(60) As Integer
    dblCT(60) As Double
    IsActiveIO As Integer
    IsActiveTC1 As Integer
    IsActiveTC2 As Integer
    IsActivePM As Integer
    
    allZerosColumns As String
    isEnd As Boolean
End Type

Public Type ParaType
    AlarmDo(100) As Integer
    AlarmName(100)  As String
    AlarmActive(100) As Integer
    AlarmBypass(100) As Integer
    UseBarcodeServer As Integer
    UseAutoMode As Integer
    UseCT As Integer
    UseMTC As Integer
    UseMTCB As Integer
    UseCIM As Integer
    UseTempMeter As Integer
    UseAz1 As Integer
    UseAz2 As Integer
    useTPump As Integer
    UseCover As Integer
    
    sngGaugeAngle As Single
    intMonitorIndex As Integer
    RtaType As Integer
    intComCT As Integer
    intCycleRuns As Integer
    intMonitorRuns As Integer
    strTestRunKey As String
    strServerPath As String
    strLastRecipe As String
    strAzIP1 As String
    strAzIP2 As String
    strRobotIP As String
    IsHoldSafety As Integer
    IsCali As Integer
    sngO2Gate As Single
    intLampAlarmTime As Integer
    intOpenDoorTime As Integer
    intOnlyRecipe As Integer
    
    intAutoPort As Integer
    intCIMPort As Integer
    intPumpDelay As Integer
    intPMbig As Integer
    intPMsmall As Integer
    
    IsUseGas(10) As Integer
    strGasAlias(10) As String
    strGasUnit(10) As String
    sngGasBias(10) As Single
    sngGasError(10) As Single
    sngGasErrorN(10) As Single
    
    IsUseCustom As Integer
    sngRatioCUT(10) As Single
    sngRatioCUM(10) As Single
    sngRatioCUP As Single
    
    sngGaugeD As Single
    sngGaugeVP As Single
    sngGaugeVN As Single
        
End Type

Public Type MultiLoopType
    blnUseMultiLoop  As Boolean
    blnUseLoop(GB_MAX_LOOPS + 2) As Boolean
    blnLoopReset(GB_MAX_LOOPS + 2) As Boolean
    sngLoopOut(GB_MAX_LOOPS + 2) As Single
    sngLoopPN(GB_MAX_LOOPS + 2) As Single
    sngLoopIN(GB_MAX_LOOPS + 2) As Single
    sngLoopDN(GB_MAX_LOOPS + 2) As Single
    sngLoopRT(GB_MAX_LOOPS + 2) As Single
    sngLoopFT(GB_MAX_LOOPS + 2) As Single
    sngLoopCV(GB_MAX_LOOPS + 2) As Single
    intLoopCN(GB_MAX_LOOPS + 2) As Integer
    intLoopTC(GB_MAX_LOOPS + 2) As Integer
    intLoopA(GB_MAX_LOOPS + 2) As Integer
    intLoopB(GB_MAX_LOOPS + 2) As Integer
    intLoopC(GB_MAX_LOOPS + 2) As Integer
    intLoopD(GB_MAX_LOOPS + 2) As Integer
    intLoopE(GB_MAX_LOOPS + 2) As Integer
    intLoopF(GB_MAX_LOOPS + 2) As Integer
    intLoopG(GB_MAX_LOOPS + 2) As Integer
    intLoopH(GB_MAX_LOOPS + 2) As Integer
    intLoopJ(GB_MAX_LOOPS + 2) As Integer
    intLoopK(GB_MAX_LOOPS + 2) As Integer
    
    intLoopMA(GB_MAX_LOOPS + 2) As Integer
    intLoopMB(GB_MAX_LOOPS + 2) As Integer
    intLoopMC(GB_MAX_LOOPS + 2) As Integer
    intLoopMD(GB_MAX_LOOPS + 2) As Integer
    intLoopME(GB_MAX_LOOPS + 2) As Integer
    intLoopMF(GB_MAX_LOOPS + 2) As Integer
    intLoopMG(GB_MAX_LOOPS + 2) As Integer
    intLoopMH(GB_MAX_LOOPS + 2) As Integer
    intLoopMJ(GB_MAX_LOOPS + 2) As Integer
    intLoopMK(GB_MAX_LOOPS + 2) As Integer
    blnLoopRTActive(GB_MAX_LOOPS + 2, 10) As Boolean
    lnLoopRTFlag(GB_MAX_LOOPS + 2, 10) As Long
    
    sngIntergalSigma(GB_MAX_LOOPS + 2) As Single
    sngWeight(GB_SCR_MAX + 2) As Single
    
End Type

Public Type AzbilType
    blnUseAzbil  As Boolean
    blnAutoTuning  As Boolean
    blnUseLoop(GB_MAX_LOOPS)  As Boolean
    blnReset(GB_MAX_LOOPS)  As Boolean
    blnStart(GB_MAX_LOOPS)  As Boolean
    intMode(GB_MAX_LOOPS) As Integer
    intINP(GB_MAX_LOOPS) As Integer
    
    intOut(GB_MAX_LOOPS) As Integer
    
    sngPV(GB_MAX_LOOPS) As Single
    sngMV(GB_MAX_LOOPS) As Single
    sngPN(GB_MAX_LOOPS) As Single
    sngIN(GB_MAX_LOOPS) As Single
    sngDN(GB_MAX_LOOPS) As Single
    sngRT(GB_MAX_LOOPS) As Single
    sngST(GB_MAX_LOOPS) As Single
    sngOffset(GB_MAX_LOOPS) As Single
    intTemp1 As Integer
    intTime1 As Integer
    intTemp2 As Integer
    intTime2 As Integer
    
End Type

Public Type RecipeType
    
    sngP1 As Single
    sngP2 As Single
    sngI1 As Single
    sngI2 As Single
    sngD1 As Single
    sngD2 As Single
End Type

'Public Type TempOffsetType
'    Tc1Offset As Single
'    Tc2Offset As Single
'    Tc3Offset As Single
'    Tc4Offset As Single
'    Tc5Offset As Single
'End Type

Public ManualStop As Boolean
Public RecordCount As Integer
Public GbTcoffset_Switch As Integer
Public GbChamberNo_Switch As Integer
Public GbRcpName As String
Public GbHoldState As Boolean
Public GbLogRcdCount As Long
Public TempOffset() As Single
Public GbShowDebugButton As Integer
Public HoldCount As Integer



