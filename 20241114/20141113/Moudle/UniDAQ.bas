Attribute VB_Name = "UniDAQ"
Option Explicit

Global Const MAX_BOARD_NUMBER = 32
Global Const MAX_EVENT_NUMBER = 20
Global Const MAX_AO_CHANNEL = 32

'ModelNumber
Global Const PIOD56 = &H800140
Global Const PIOD48 = &H800130
Global Const PIOD64 = &H800120
Global Const PIOD96 = &H800110
Global Const PIOD144 = &H800100
Global Const PIOD168 = &H800150
Global Const PIODA = &H800400
Global Const PIO821 = &H800310
Global Const PIO827 = &HFF0000

Global Const PISOP16R16U = &H1800FF
Global Const PISOC64 = &H800800
Global Const PISOP64 = &H800810
Global Const PISOA64 = &H800850
Global Const PISOP32C32 = &H800820
Global Const PISOP32A32 = &H800870
Global Const PISOP8R8 = &H800830
Global Const PISO730 = &H800840
Global Const PISO730A = &H800880
Global Const PISO725 = &H8008FF
Global Const PISODA2 = &H800B00

Global Const PISO813 = &H800A00

Global Const PCITMC12 = &HDF2962
Global Const PCIM512 = &HDE9562
Global Const PCIM256 = &HDE92A6
Global Const PCIM128 = &HDE9178
Global Const PCIFC16 = &HB13017
Global Const PCID64 = &HDE3513

Global Const PCI822 = &HDE3823
Global Const PCI826 = &HDE3827
Global Const PCI827 = &HDE3828
Global Const PCI100x = &H341002
Global Const PCI1202 = &H345672
Global Const PCI1602 = &H345676
Global Const PCI180x = &H345678
Global Const PCIP8R8 = &HD6102B
Global Const PCIP16R16 = &HD61E39


Global Const MAX_CFGCODE_NUMBER = 23 '使用者定義ConfigCode的Total數量

'User Config Code
Global Const IXUD_BI_10V = 0      'Bipolar +/- 10V
Global Const IXUD_BI_5V = 1       'Bipolar +/-  5V
Global Const IXUD_BI_2V5 = 2      'Bipolar +/-  2.5V
Global Const IXUD_BI_1V25 = 3     'Bipolar +/-  1.25V
Global Const IXUD_BI_0V625 = 4    'Bipolar +/-  0.625V
Global Const IXUD_BI_0V3125 = 5   'Bipolar +/-  0.3125V
Global Const IXUD_BI_0V5 = 6      'Bipolar +/-  0.5V
Global Const IXUD_BI_0V05 = 7     'Bipolar +/-  0.05V
Global Const IXUD_BI_0V005 = 8    'Bipolar +/-  0.005V
Global Const IXUD_BI_1V = 9       'Bipolar +/-  1V
Global Const IXUD_BI_0V1 = 10     'Bipolar +/-  0.1V
Global Const IXUD_BI_0V01 = 11    'Bipolar +/-  0.01V
Global Const IXUD_BI_0V001 = 12   'Bipolar +/-  0.001V
Global Const IXUD_UNI_20V = 13    'Unipolar 0 ~ 20V
Global Const IXUD_UNI_10V = 14    'Unipolar 0 ~ 10V
Global Const IXUD_UNI_5V = 15     'Unipolar 0 ~  5V
Global Const IXUD_UNI_2V5 = 16    'Unipolar 0 ~  2.5V
Global Const IXUD_UNI_1V25 = 17   'Unipolar 0 ~  1.25V
Global Const IXUD_UNI_0V625 = 18  'Unipolar 0 ~  0.625V
Global Const IXUD_UNI_1V = 19     'Unipolar 0 ~  1V
Global Const IXUD_UNI_0V1 = 20    'Unipolar 0 ~  0.1V
Global Const IXUD_UNI_0V01 = 21   'Unipolar 0 ~  0.01V
Global Const IXUD_UNI_0V001 = 22  'Unipolar 0 ~  0.001V

'//User AO Config Code for Voltage
Global Const IXUD_AO_UNI_5V = 0       'Unipolar 0  ~  5V
Global Const IXUD_AO_BI_5V = 1        'Bipolar  +/-   5V
Global Const IXUD_AO_UNI_10V = 2      'Unipolar 0  ~ 10V
Global Const IXUD_AO_BI_10V = 3       'Bipolar +/-  10V
Global Const IXUD_AO_UNI_20V = 4      'Unipolar 0  ~ 20V
Global Const IXUD_AO_BI_20V = 5       'Bipolar +/-  20V

'//User AO Config Code for Current
Global Const IXUD_AO_I_0_20_MA = 16     '0 ~ 20mA
Global Const IXUD_AO_I_4_20_MA = 17     '4 ~ 20mA

' Return code
Global Const Ixud_NoErr = 0
Global Const Ixud_OpenDriverErr = 1
Global Const Ixud_PnPDriverErr = 2
Global Const Ixud_DriverNoOpen = 3
Global Const Ixud_GetDriverVersionErr = 4
Global Const Ixud_ExceedBoardNumber = 5
Global Const Ixud_FindBoardErr = 6
Global Const Ixud_BoardMappingErr = 7
Global Const Ixud_DIOModesErr = 8
Global Const Ixud_InvalidAddress = 9
Global Const Ixud_InvalidSize = 10
Global Const Ixud_InvalidPortNumber = 11
Global Const Ixud_ISetDio24AddrErr = 12
Global Const Ixud_ISetDio16AddrErr = 13
Global Const Ixud_UnSupportedModel = 14

Global Const Ixud_UnSupportedFun = 16
Global Const Ixud_InvalidChannelNumber = 17
Global Const Ixud_InvalidValue = 18
Global Const Ixud_InvalidMode = 19

Global Const Ixud_EAITimeOut = 22
Global Const Ixud_ITimeOutErr = 24
Global Const Ixud_EAIChNumErr = 25              '沒有此AI Channel或無AI功能
Global Const Ixud_EAIModelErr = 26              '沒有此AI模組
Global Const Ixud_ECfgCodeMapErr = 27           '尋找不到適用的ConfigCode Record
Global Const Ixud_IADCtrllerTimeoutErr = 28     'AD Controller Time Out
Global Const Ixud_IPCIRecordMapErr = 29         '尋找不到適用的UniDAQ_PCI
Global Const Ixud_ESetCardTypeErr = 30          '沒有此CardType
Global Const Ixud_EAllocateMemErr = 31          '分配記憶體空間失敗
Global Const Ixud_EDisableCounterErr = 32       '關閉counter失敗
Global Const Ixud_EInstallEventErr = 33         '安裝中斷事件失敗
Global Const Ixud_EInstallIrqErr = 34           '安裝中斷IRQ失敗
Global Const Ixud_ClearIntCountErr = 35         '清除中斷計數量失敗
Global Const Ixud_EGetSysBufferErr = 36         '取得系統buffer失敗
Global Const Ixud_ERemoveIrqErr = 37            '移除IRQ失敗
Global Const Ixud_ECreateEventErr = 38          'CreateEvent 失敗
Global Const Ixud_EWaitEventTimeOutErr = 39     '等待事件失敗
Global Const Ixud_EAIResolutionErr = 40         '無此AI解析度
Global Const Ixud_ECreateThreadErr = 41         '建立執行緒失敗
Global Const Ixud_EThreadTimeOutErr = 42        '
Global Const Ixud_EFIFOOverFlowErr = 43         'FIFO OverFlow
Global Const Ixud_EFIFOTimeOutErr = 44
Global Const Ixud_EGetIntInstallStatusErr = 45
Global Const Ixud_EBufCountErr = 46             '指定錯誤的Count數(不能為零或超過Start AI的設定值)
Global Const Ixud_ESetBufCountErr = 47          '設定Ioctrl buf錯誤
Global Const Ixud_EGetBufStatusErr = 48         '取得sys內的buf狀態錯誤
Global Const Ixud_EGetBoardNoErr = 49           '取得版卡號碼錯誤(找不到對應的Card ID)
Global Const Ixud_IEventThreadErr = 50          'EventThread Err
Global Const Ixud_IAutoCreateEventErr = 51      '
Global Const Ixud_RegThreadErr = 52
Global Const Ixud_SearchEventErr = 53
Global Const Ixud_TimeOutErr = 54
Global Const Ixud_FifoOverflow = 55             'FIFO Overflow
Global Const Ixud_InvalidBlock = 56             'EEP Block 設定錯誤
Global Const Ixud_InvalidAddr = 56              'EEP Address 設定錯誤


Public Type IXUD_DEVICE_INFO
    AdwSize As Long
    
    wVendorID As Integer
    wDeviceID As Integer
    
    wSubVendorID As Integer
    wSubDeviceID As Integer
    dwBar(0 To 5)  As Long
   
    BusNo As Byte
    DevNo As Byte
    IRQ As Byte
    Aux As Byte
    
    dwReserved1(0 To 5) As Long  'Reserver
   
End Type



Public Type IXUD_CARD_INFO
    dwSize As Long        'Structure size
    dwModelNo As Long         'Model Number
    
   'CardID is update when calling the function each time.
    CardID As Byte            'for new cards, =&hFF=N/A
    wSingleEnded As Byte   'for new cards,1:S.E 2:D.I.F,=&hFF=N/A
    wReserved As Integer      'Reserver

    wAIChannels As Integer    'Number of AI channels(AD)
    wAOChannels As Integer        'Number of AO channels(DA)

    wDIPorts As Integer           'Number of DI ports
    wDOPorts As Integer           'Number of DO ports

    wDIOPorts As Integer          'Number of DIO ports
    wDIOPortWidth As Integer      'The width is 8/16/32 bit.

    wCounterChannels As Integer   'Number of Timers/Counters
    wMemorySize As Integer        'PCI-M512==>512, Units in KB.

    dwReserved1(0 To 5) As Long  'Reserver
  
    
End Type




Declare Function Ixud_GetDllVersion Lib "UniDAQ.dll" (ByRef wDLLVer As Long) As Integer

'Driver functions
Declare Function Ixud_DriverInit Lib "UniDAQ.dll" (ByRef wTotalBoard As Integer) As Integer
Declare Function Ixud_DriverClose Lib "UniDAQ.dll" () As Integer
Declare Function Ixud_GetBoardNoByCardID Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwModelNumber As Long, ByVal wCardID As Integer) As Integer
Declare Function Ixud_SearchCard Lib "UniDAQ.dll" (ByRef wTotalBoard As Integer, ByVal dwModelNo As Long) As Integer
Declare Function Ixud_GetCardInfo Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByRef sDevInfo As IXUD_DEVICE_INFO, ByRef sCardInfo As IXUD_CARD_INFO, ByVal szModelNmae As String) As Integer


'Declare Function    Ixud_GetDeviceInfo
Declare Function Ixud_ReadPort Lib "UniDAQ.dll" (ByVal dwAddress As Long, ByVal wSize As Integer, ByRef dwVal As Long) As Integer
Declare Function Ixud_WritePort Lib "UniDAQ.dll" (ByVal dwAddress As Long, ByVal wSize As Integer, ByVal dwVal As Long) As Integer
            
Declare Function Ixud_SetDIOModes32 Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwDioMode As Long) As Integer
Declare Function Ixud_SetDIOMode Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wPortNo As Integer, ByVal wDioMode As Integer) As Integer
Declare Function Ixud_ReadDI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wPortNo As Integer, ByRef dwDIVal As Long) As Integer
Declare Function Ixud_WriteDO Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wPortNo As Integer, ByVal dwDOVal As Long) As Integer
'Declare Function    Ixud_ReadDI2
'Declare Function    Ixud_WriteDO2
Declare Function Ixud_SoftwareReadbackDO Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wPortNo As Integer, ByRef dwDOVal As Long) As Integer
Declare Function Ixud_SetPWMOutput Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wMode As Integer, ByVal fFrequency As Single, ByVal fDutyRate As Single, ByVal fDelayUS As Single) As Integer

Declare Function Ixud_SetEventCallback Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wInterruptSource As Integer, ByRef hEvent As Long, ByVal CallbackFun As Long, ByVal dwCallBackParameter As Long) As Integer
Declare Function Ixud_RemoveEventCallback Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wInterruptSource As Integer) As Integer

Declare Function Ixud_InstallIrq Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwInterruptMask As Long) As Integer
Declare Function Ixud_RemoveIrq Lib "UniDAQ.dll" (ByVal wBoardNo As Integer) As Integer


Declare Function Ixud_ReadCounter Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByRef dwValue As Long) As Integer
Declare Function Ixud_SetCounter Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wMode As Integer, ByVal dwValue As Long) As Integer
Declare Function Ixud_DisableCounter Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer) As Integer
Declare Function Ixud_SetFCChannelMode Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wMode As Integer, ByVal wDelayMs As Integer) As Integer
Declare Function Ixud_ReadFrequency Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByRef fFrequency As Single, ByVal dwTimeOutMs As Long, ByRef wStatus As Integer) As Integer

Declare Function Ixud_ReadMemory Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwOffset As Long, ByVal size As Integer, ByRef dwValue As Long) As Integer
Declare Function Ixud_WriteMemory Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwOffset As Long, ByVal size As Integer, ByRef dwValue As Long) As Integer

Declare Function Ixud_ReadAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByRef fValue As Single) As Integer
Declare Function Ixud_ReadAIH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByRef dwValue As Long) As Integer
Declare Function Ixud_PollingAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByVal dwDataCount As Long, ByRef fValue As Single) As Integer
Declare Function Ixud_PollingAIH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByVal dwDataCount As Long, ByRef dwValue As Long) As Integer
Declare Function Ixud_PollingAIScan Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannels As Integer, ByRef wChannelList As Integer, ByRef wConfigList As Integer, ByVal dwDataCountPerChannel As Long, ByVal fValue As Single) As Integer
Declare Function Ixud_PollingAIScanH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannels As Integer, ByRef wChannelList As Integer, ByRef wConfigList As Integer, ByVal dwDataCountPerChannel As Long, ByVal dwValue As Long) As Integer
Declare Function Ixud_PacerAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByVal fSamplingRate As Single, ByVal dwDataCount As Long, ByRef fValue As Single) As Integer
Declare Function Ixud_PacerAIH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByVal fSamplingRate As Single, ByVal dwDataCount As Long, ByRef dwValue As Long) As Integer
Declare Function Ixud_PacerAIScan Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannels As Integer, ByRef wChannelList As Integer, ByRef wConfigList As Integer, ByVal fSamplingRate As Single, ByVal dwDataCountPerChannel As Long, ByRef fValue As Single) As Integer
Declare Function Ixud_PacerAIScanH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannels As Integer, ByRef wChannelList As Integer, ByRef wConfigList As Integer, ByVal fSamplingRate As Single, ByVal dwDataCountPerChannel As Long, ByRef hValue As Long) As Integer

Declare Function Ixud_ConfigAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wFIFOSizeKB As Integer, ByVal BufferSizeKB As Long, ByVal wCardType As Integer, ByVal wDelaySettlingTime As Integer) As Integer
Declare Function Ixud_ClearAIBuffer Lib "UniDAQ.dll" (ByVal wBoardNo As Integer) As Integer
Declare Function Ixud_GetBufferStatus Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByRef wBufferStatus As Integer, ByRef dwDataCount As Long) As Integer

Declare Function Ixud_StartAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByVal fSamplingRate As Single, ByVal dwDataCount As Long) As Integer
Declare Function Ixud_StartAIScan Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannels As Integer, ByRef wChannelList As Integer, ByRef wConfigList As Integer, ByVal fSamplingRate As Single, ByVal dwDataCountPerChannel As Long) As Integer
Declare Function Ixud_StartExtAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wActive As Integer, ByVal wChannel As Integer, ByVal wConfig As Integer, ByVal fSamplingRate As Single, ByVal dwPostDataCount As Long, ByVal dwPreDataCount As Long) As Integer
Declare Function Ixud_StartExtAIScan Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannels As Integer, ByVal wActive As Integer, ByRef wChannelList As Integer, ByRef wConfigList As Integer, ByVal fSamplingRate As Single, ByVal dwPostDataCountPerChannel As Long, ByVal dwPreDataCountPerChannel As Long) As Integer
Declare Function Ixud_GetAIBuffer Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwDataCount As Long, ByRef fValue As Single) As Integer
Declare Function Ixud_GetAIBufferH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal dwDataCount As Long, ByRef hValue As Long) As Integer
Declare Function Ixud_StopAI Lib "UniDAQ.dll" (ByVal wBoardNo As Integer) As Integer

Declare Function Ixud_ConfigAO Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal wCfgCode As Integer) As Integer
Declare Function Ixud_WriteAOVoltage Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal fValue As Single) As Integer
Declare Function Ixud_WriteAOVoltageH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal hValue As Long) As Integer
Declare Function Ixud_WriteAOCurrent Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal fValue As Single) As Integer
Declare Function Ixud_WriteAOCurrentH Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal hValue As Long) As Integer

Declare Function Ixud_WriteEEP Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByRef dwWriteVal As Long) As Integer
Declare Function Ixud_ReadEEP Lib "UniDAQ.dll" (ByVal wBoardNo As Integer, ByRef dwReadVal As Long) As Integer
