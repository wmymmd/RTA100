Attribute VB_Name = "PIODA"

Global Const PIODA_NoError = 0
Global Const PIODA_DriverOpenError = 1
Global Const PIODA_DriverNoOpen = 2
Global Const PIODA_GetDriverVersionError = 3
Global Const PIODA_InstallIrqError = 4
Global Const PIODA_ClearIntCountError = 5
Global Const PIODA_GetIntCountError = 6
Global Const PIODA_RegisterApcError = 7
Global Const PIODA_RemoveIrqError = 8
Global Const PIODA_FindBoardError = 9
Global Const PIODA_ExceedBoardNumber = 10
Global Const PIODA_ResetError = 11

Global Const PIODA_EEPROMDataError = 12
Global Const PIODA_EEPROMWriteError = 13


' to trigger a interrupt when high -> low
Global Const PIODA_ActiveLow = 0
' to trigger a interrupt when low -> high
Global Const PIODA_ActiveHigh = 1


' ID
Global Const PIO_DA = &H800400         ' PIO-DA16/DA8/DA4


' The Test functions
Declare Function PIODA_ShortSub Lib "PIODA.dll" _
        (ByVal a As Integer, ByVal b As Integer) As Integer
Declare Function PIODA_FloatSub Lib "PIODA.dll" _
        (ByVal a As Single, ByVal b As Single) As Single
Declare Function PIODA_GetDllVersion Lib "PIODA.dll" () As Integer

' The Driver functions
Declare Function PIODA_DriverInit Lib "PIODA.dll" () As Integer
Declare Sub PIODA_DriverClose Lib "PIODA.dll" ()
Declare Function PIODA_SearchCard Lib "PIODA.dll" _
        (wBoards As Integer, ByVal dwPIOPISOCardID As Long) As Integer
Declare Function PIODA_GetDriverVersion Lib "PIODA.dll" _
        (wDriverVersion As Integer) As Integer
Declare Function PIODA_GetConfigAddressSpace Lib "PIODA.dll" ( _
        ByVal wBoardNo As Integer, wAddrBase As Long, wIrqNo As Integer, _
        wSubVendor As Integer, wSubDevice As Integer, wSubAux As Integer, _
        wSlotBus As Integer, wSlotDevice As Integer) As Integer

Declare Function PIODA_ActiveBoard Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer) As Integer
Declare Function PIODA_WhichBoardActive Lib "PIODA.dll" () As Integer
Declare Function PIODA_SetCounter Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wWhichCounter As Integer, _
         ByVal bConfig As Integer, ByVal wValue As Long) As Long
Declare Function PIODA_GetBaseAddress Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer) As Long

' EEPROM functions
Declare Function PIODA_EEP_READ Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wOffset As Integer, _
         bHi As Integer, bLo As Integer) As Integer
Declare Function PIODA_EEP_WR_EN Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer) As Integer
Declare Function PIODA_EEP_WR_DIS Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer) As Integer
Declare Function PIODA_EEP_WRITE Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wOffset As Integer, _
         ByVal HI As Integer, ByVal LO As Integer) As Integer

' DA functions
Declare Function PIODA_Voltage Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wChannel As Integer, _
         ByVal fValue As Single) As Integer
Declare Function PIODA_Current Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wChannel As Integer, _
         ByVal fValue As Single) As Integer
Declare Function PIODA_CalVoltage Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wChannel As Integer, _
         ByVal fValue As Single) As Integer
Declare Function PIODA_CalCurrent Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wChannel As Integer, _
         ByVal fValue As Single) As Integer

' DIO functions
Declare Sub PIODA_OutputByte Lib "PIODA.dll" _
        (ByVal wBaseAddress As Long, ByVal dataout As Integer)
Declare Sub PIODA_OutputWord Lib "PIODA.dll" _
        (ByVal wBaseAddress As Long, ByVal dataout As Long)
Declare Function PIODA_InputByte Lib "PIODA.dll" _
        (ByVal wBaseAddress As Long) As Integer
Declare Function PIODA_InputWord Lib "PIODA.dll" _
        (ByVal wBaseAddress As Long) As Long
Declare Function PIODA_DI Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, wVal As Long) As Integer
Declare Function PIODA_DO Lib "PIODA.dll" _
        (ByVal wBoardNo As Integer, ByVal wDO As Long) As Integer

' Interrupt functions
Declare Function PIODA_IntInstall Lib "PIODA.dll" _
        (ByVal wBoard As Integer, hEvent As Long, _
        ByVal wInterruptSource As Integer, _
        ByVal wActiveMode As Integer) As Integer
Declare Function PIODA_IntRemove Lib "PIODA.dll" () As Integer
Declare Function PIODA_IntResetCount Lib "PIODA.dll" () As Integer
Declare Function PIODA_IntGetCount Lib "PIODA.dll" _
        (dwIntCount As Long) As Integer

