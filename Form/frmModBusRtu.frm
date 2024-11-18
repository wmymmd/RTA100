VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmModBusRtu 
   Caption         =   "SerialCommunication"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   12360
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame5 
      Caption         =   "GiveBytes"
      Height          =   9855
      Left            =   6240
      TabIndex        =   9
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         Height          =   540
         Left            =   4560
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   8535
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   5895
      End
      Begin VB.CommandButton cmdSendBytes 
         Caption         =   "Send"
         Height          =   540
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test"
      Height          =   6855
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   6015
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cdbSetting2 
         Caption         =   "SetBaudRate"
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Text            =   "19200"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cdbSetting 
         Caption         =   "SetAddress"
         Height          =   375
         Left            =   2640
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Text            =   "2"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cdbReadO2 
         Caption         =   "ReadO2"
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cdbSetMaxPower 
         Caption         =   "SetMaxPower"
         Height          =   495
         Left            =   2640
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Bank2:"
         Height          =   255
         Left            =   720
         TabIndex        =   36
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Bank1:"
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "SCRBaudRate:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "SCRAddress:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "O2:"
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.ComboBox cmbDeviceName 
      Height          =   300
      ItemData        =   "frmModBusRtu.frx":0000
      Left            =   1680
      List            =   "frmModBusRtu.frx":000D
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbPortNumber 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Communication"
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin MSCommLib.MSComm MSCommSCR 
         Left            =   5400
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         BaudRate        =   19200
      End
      Begin MSCommLib.MSComm MSCommO2SenSor 
         Left            =   4800
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         BaudRate        =   19200
      End
      Begin MSCommLib.MSComm MSCommLamp 
         Left            =   4200
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.TextBox txtTimeOut 
         Height          =   270
         Left            =   1680
         TabIndex        =   15
         Text            =   "500"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox ckeIsUsed 
         Caption         =   "Enable"
         Height          =   180
         Left            =   1680
         TabIndex        =   14
         Top             =   2520
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmdDisConnect 
         Caption         =   "DisConnect"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtStopBits 
         Height          =   270
         Left            =   1680
         TabIndex        =   5
         Text            =   "1"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtDataBits 
         Height          =   270
         Left            =   1680
         TabIndex        =   4
         Text            =   "8"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtBaudRate 
         Height          =   270
         Left            =   1680
         TabIndex        =   3
         Text            =   "9600"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbResult 
         Height          =   375
         Left            =   4320
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "TimeOut:"
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "StopBits:"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "DataBits:"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "BaudRate:"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "ComPort:"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "DeviceName:"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmModBusRtu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" ( _
    ByVal lpDeviceName As String, _
    ByVal lpTargetPath As String, _
    ByVal ucchMax As Long) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" _
    (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Dim deviceInfo As Object
Dim openedPorts As New Collection
Dim Response() As Byte
Dim DataReceived As Boolean


Private Sub cdbReadO2_Click()
    Dim ReadCommand() As Byte
    ReadCommand = GenerateReadCommand(&H1, 100, 1, &H4)
    Call sendData(ReadCommand, "O2 Sensor")
    If UBound(Response) >= 0 Then
        If UBound(Response) >= LBound(Response) Then
            Text5.text = CStr(Response(0) * 256 + Response(1))
        End If
    End If
    
End Sub

Private Sub cdbSetMaxPower_Click()
    On Error GoTo ERRLINE:
    Dim writeCommand() As Byte
    Dim commands(11) As Byte
    Dim values(11) As Integer
    Dim i As Integer
    i = 0
    commands(0) = &H1
    commands(1) = &H2
'    commands(2) = &H3
'    commands(3) = &H4
'    commands(4) = &H5
'    commands(5) = &H6
'    commands(6) = &H7
'    commands(7) = &H8
'    commands(8) = &H9
'    commands(9) = &HA
'    commands(10) = &HB
'    commands(11) = &HC
    values(0) = CInt(Text11.text)
    values(1) = CInt(Text14.text)
'    values(2) = 60
'    values(3) = 60
'    values(4) = 60
'    values(5) = 60
'    values(6) = 60
'    values(7) = 60
'    values(8) = 60
'    values(9) = 60
'    values(10) = 60
'    values(11) = 60
'    If i > UBound(commands) Then Exit Sub
    
    For i = 0 To 1
    writeCommand = GenerateWriteCommand(commands(i), 131, values(i))
    Call sendData(writeCommand, "SCR")
    
    If UBound(Response) >= 0 Then
        If Response(0) = &H0 Then
            MsgBox "Write MaxPower Succeeded!"
        End If
    End If
    Sleep (3)
    Next i
    Exit Sub
ERRLINE:
    WriteLog ("SetMaxPowererError")
    
End Sub

Private Sub cdbSetting_Click()
    Dim writeCommand() As Byte
    writeCommand = GenerateWriteCommand(&H1, 256, CInt(Text8.text))
    Call sendData(writeCommand, "SCR")
    If UBound(Response) >= 0 Then
        If Response(0) = &H0 Then
            MsgBox "Write Address Succeeded!"
        End If
    End If
    
End Sub

Private Sub cdbSetting2_Click()
    Dim writeCommand() As Byte
    writeCommand = GenerateWriteCommand(&H1, 257, CInt(Text16.text))
    Call sendData(writeCommand, "SCR")
    If UBound(Response) >= 0 Then
        If Response(0) = &H0 Then
            MsgBox "Write BaudRate Succeeded!"
        End If
    End If
End Sub

Private Sub cmbPortNumber_DropDown()
    RefreshSerialPorts
End Sub

Private Sub cmdConnect_Click()
       Dim result As Boolean
    
'    InitModBusRtu
   Call WriteConfig(cmbDeviceName.text, ckeIsUsed.value, cmbPortNumber.text, txtBaudRate.text, txtDataBits.text, txtStopBits.text, txtTimeOut.text)
    ReadConfig
    If Not IsEmpty(deviceInfo) And deviceInfo.Count > 0 Then
'        OpenMSComm(DeviceName)
        result = OpenMSComm(cmbDeviceName.text)
        If result Then
            lbResult.Caption = "Connect Succeed!"
            lbResult.ForeColor = &HFF00&
            cmdConnect.Enabled = False
        Else
            lbResult.Caption = "Connect Fail!"
            lbResult.ForeColor = &HFF&
            cmdConnect.Enabled = True
        End If
    Else
         MsgBox "no Config", vbExclamation
    End If
End Sub

Private Sub cmdDisConnect_Click()
    If CloseMSComm(cmbDeviceName.text) Then
        lbResult.Caption = ""
        cmdConnect.Enabled = True
    End If
End Sub

Private Sub cmdSendBytes_Click()
     Dim data() As Byte
     data = StringToByte(Text10.text)
    Call sendData(data, cmbDeviceName.text)
End Sub

Private Sub Command6_Click()
 Text12.text = ""
End Sub

Private Sub Form_Load()
    ReadConfig
    DataReceived = True
End Sub

Private Function ProcessResponse(recivedData() As Byte) As Byte()
    Const MINIMUM_FRAME_LENGTH As Integer = 4
    
    Dim address As Byte
    Dim functionCode As Byte
    Dim dataBytes() As Byte
    Dim checksum As Byte
    Dim byteCount As Integer
    Dim registerCount As Integer
    Dim registerData() As Byte
    Dim i As Integer
    
     If UBound(recivedData) < MINIMUM_FRAME_LENGTH - 1 Then
        MsgBox "length is error"
        Exit Function
     End If
     
'    dataBytes = recivedData
    Dim j As Integer, k As Integer
    ReDim dataBytes(UBound(recivedData) \ 2)
    k = 0
    For j = LBound(recivedData) To UBound(recivedData) Step 2
        dataBytes(k) = CByte("&H" & Format(Hex(recivedData(j)), "00"))
        k = k + 1
    Next j
    
    address = dataBytes(0)
    functionCode = dataBytes(1)
    
    Select Case functionCode
        Case &H3, &H4
            byteCount = dataBytes(2)
            Dim expectedLength As Integer
            expectedLength = 3 + byteCount + 1
            registerCount = byteCount \ 2
            ReDim registerData(registerCount * 2 - 1)
            For i = 0 To registerCount - 1
                registerData(i * 2) = dataBytes(3 + i * 2)
                registerData(i * 2 + 1) = dataBytes(4 + i * 2)
            Next i
            ProcessResponse = registerData
        Case &H6
             Dim registerAddressHigh As Byte
             Dim registerAddressLow As Byte
             Dim registerValueHigh As Byte
             Dim registerValueLow As Byte
             
             registerAddressHigh = dataBytes(2)
             registerAddressLow = dataBytes(3)
             registerValueHigh = dataBytes(4)
             registerValueLow = dataBytes(5)
             ReDim registerData(0 To 0)
             registerData(0) = &H0
             ProcessResponse = registerData
             
        Case Else
             ProcessResponse = ""
    End Select
    
End Function

Private Function BuildCommand(ByVal deviceID As Byte, ByVal functionCode As Byte, parameters() As Byte) As Byte()
    Dim command(5) As Byte
    command(0) = deviceID
    command(1) = functionCode
    Dim i As Integer
    For i = 0 To UBound(parameters)
        command(2 + i) = parameters(i)
    Next

    BuildCommand = command
End Function

Private Function GenerateReadCommand(ByVal deviceID As Byte, ByVal startAddress As Integer, ByVal quantity As Integer, ByVal functionCode As Byte) As Byte()
    Dim parameters(3) As Byte
    parameters(0) = startAddress \ 256
    parameters(1) = startAddress Mod 256
    parameters(2) = quantity \ 256
    parameters(3) = quantity Mod 256

    GenerateReadCommand = BuildCommand(deviceID, functionCode, parameters)
End Function

Private Function GenerateWriteCommand(ByVal deviceID As Byte, ByVal startAddress As Integer, ByVal value As Integer) As Byte()
    Dim parameters(3) As Byte
    parameters(0) = startAddress \ 256
    parameters(1) = startAddress Mod 256
    parameters(2) = value \ 256
    parameters(3) = value Mod 256

    GenerateWriteCommand = BuildCommand(deviceID, &H6, parameters)
End Function

Public Function StringToByte(SendCmd As String) As Byte()
'    Dim hexString As String
    Dim byteCount As Integer
    Dim i As Integer
    Dim bytes() As Byte
    
'    hexString = Text10.text
    If Len(SendCmd) < 12 Then
        MsgBox "cmd length is error"
        Exit Function
    End If

    SendCmd = Replace(SendCmd, " ", "")
    byteCount = Len(SendCmd) \ 2
    ReDim bytes(0 To byteCount - 1)

    For i = 0 To byteCount - 1
        bytes(i) = Val("&H" & Mid(SendCmd, i * 2 + 1, 2))
    Next i
    StringToByte = bytes
End Function

Public Sub sendData(data() As Byte, portName As String)
    Dim crc() As Byte
    Dim sendData() As Byte
    Dim i As Integer
    
        crc = Modbus_CRC(data)
    
         ReDim sendData(LBound(data) To UBound(data) + UBound(crc))
         For i = LBound(data) To UBound(data)
            sendData(i) = data(i)
         Next i
         For i = LBound(crc) To UBound(crc)
            sendData(UBound(data) + i) = crc(i)
        Next i
        Select Case portName
            Case "SCR"
                 If MSCommSCR.PortOpen Then
                    MSCommSCR.Output = sendData
                Else
                    MsgBox "Serial not open!"
                End If
            Case "O2 Sensor"
                 If MSCommO2SenSor.PortOpen Then
                    MSCommO2SenSor.Output = sendData
                Else
                    MsgBox "Serial not open!"
                End If
            Case "Lamp"
                 If MSCommLamp.PortOpen Then
                    MSCommLamp.Output = sendData
                Else
                    MsgBox "Serial not open!"
                End If
            Case Else
                MsgBox "Invalid Serial Port!"
        End Select
        
'    DataReceived = False
'    Dim startTime As Single
'    startTime = Timer
'
'     Do While Not DataReceived
'        DoEvents
'        If Timer - startTime > 500 Then
'            MsgBox "Response timeout. Please try again."
'            Exit Sub
'        End If
'    Loop
End Sub


Private Sub MSCommSCR_OnComm()
    HandleCommEvent MSCommSCR
End Sub
Private Sub MSCommO2SenSor_OnComm()
    HandleCommEvent MSCommO2SenSor
End Sub
Private Sub MSCommLamp_OnComm()
    HandleCommEvent MSCommLamp
End Sub
Public Sub HandleCommEvent(comm As MSComm)
    Dim receivedData() As Byte
    Dim receivedString As String
    Dim i As Integer
    Dim bytesToReceive As Integer
    
    If comm.CommEvent = comEvReceive Then
'        Dim startTime As Single
'        startTime = Timer
'        Do While Timer < startTime + 0.05
'            DoEvents
'        Loop
        bytesToReceive = 7
        Do While comm.InBufferCount > 0 And comm.InBufferCount < bytesToReceive
            DoEvents
        Loop
        
        If comm.InBufferCount > 0 Then
               
                receivedData = comm.Input
            
        '        receivedString = StrConv(receivedData, vbUnicode)
                For i = LBound(receivedData) To UBound(receivedData)
                        If i Mod 2 = 0 Then
                            receivedString = receivedString & Format(Hex(receivedData(i)), "00") & " "
                        End If
                Next i
                
'                Text12.text = "SendData:" & Text10 & "ReceiveData:" & receivedString
                
                Response = ProcessResponse(receivedData)
                DataReceived = True
                
        End If
    End If
End Sub

Private Sub RefreshSerialPorts()
    Dim portName As String * 255
    Dim result As Long
    Dim i As Integer
    Dim portList As String
    
    cmbPortNumber.Clear
    
      For i = 1 To 16
        portName = Space(255)
        result = QueryDosDevice("COM" & CStr(i), portName, 255)
        If result <> 0 Then
            portList = "COM" & CStr(i)
            cmbPortNumber.AddItem portList
        End If
    Next i

    If cmbPortNumber.ListCount = 0 Then
        MsgBox "No Serial port was found"
    End If
End Sub


Sub InitializeDeviceInfo()
    Set deviceInfo = CreateObject("Scripting.Dictionary")
    deviceInfo.RemoveAll
End Sub

Public Sub ReadConfig()
    Dim configFileName As String
    Dim returnValue As Long
    Dim sectionNames As String
    Dim sectionList() As String
    Dim deviceIndex As Integer
    Dim sectionNamesBuffer As String * 255
    Dim deviceSection As String
    Dim DeviceName As String
    Dim IsUsed As Boolean
    Dim Port As String
    Dim BaudRate As Long
    Dim DataBits As Long
    Dim StopBits As Long
    Dim Timeout As Long
    
    InitializeDeviceInfo
    configFileName = gbSystemPath & "\Config\ModbusRtu.ini"
    If dir(configFileName) = "" Then
        GoTo ERR_GETCONFIG
    End If
    
    returnValue = GetPrivateProfileSectionNames(sectionNamesBuffer, Len(sectionNamesBuffer), configFileName)
    sectionNames = Left$(sectionNamesBuffer, returnValue)
    sectionList = Split(sectionNames, Chr$(0))
    
    For deviceIndex = LBound(sectionList) To UBound(sectionList)
        deviceSection = sectionList(deviceIndex)
    
    If Left$(deviceSection, 6) = "Device" Then
            DeviceName = ReadIniValue(configFileName, deviceSection, "DeviceName", "")
            IsUsed = (ReadIniValue(configFileName, deviceSection, "IsUsed", "0") = "1")
            Port = ReadIniValue(configFileName, deviceSection, "Port", "")
            BaudRate = Val(ReadIniValue(configFileName, deviceSection, "BaudRate", "0"))
            DataBits = Val(ReadIniValue(configFileName, deviceSection, "DataBits", "0"))
            StopBits = Val(ReadIniValue(configFileName, deviceSection, "StopBits", "0"))
            Timeout = Val(ReadIniValue(configFileName, deviceSection, "Timeout", "0"))
    
            Dim deviceDetails As Object
            Set deviceDetails = CreateObject("Scripting.Dictionary")
            deviceDetails.Add "DeviceName", DeviceName
            deviceDetails.Add "IsUsed", IsUsed
            deviceDetails.Add "Port", Port
            deviceDetails.Add "BaudRate", BaudRate
            deviceDetails.Add "DataBits", DataBits
            deviceDetails.Add "StopBits", StopBits
            deviceDetails.Add "Timeout", Timeout
    
            deviceInfo.Add CStr(deviceIndex + 1), deviceDetails
    End If
    Next deviceIndex
    
'    Dim key As Variant
'    For Each key In deviceInfo.Keys
'        Debug.Print "Device " & key & " Name: " & deviceInfo(key)("DeviceName") & " Port: " & deviceInfo(key)("Port")
'    Next key
    Exit Sub
ERR_GETCONFIG:
    Call AlertShow("打開 ModubusRtu 配置文件失敗!!", ERRORTYPE)
End Sub

Private Function ReadIniValue(ByVal fileName As String, ByVal Section As String, _
    ByVal Key As String, ByVal defaultValue As String) As String
    Dim buffer As String * 255
    Dim Length As Long
    
    Length = GetPrivateProfileString(Section, Key, defaultValue, buffer, Len(buffer), fileName)
    ReadIniValue = Left$(buffer, Length)
End Function

Function OpenMSComm(DeviceName As String) As Boolean

    Dim MSComm As MSComm
    
    
    Dim portNumber As String
    Dim deviceIndex As Variant
    Dim foundDevice As Boolean
    foundDevice = False
    
    Select Case DeviceName
        Case "SCR"
            Set MSComm = MSCommSCR
        Case "O2 Sensor"
            Set MSComm = MSCommO2SenSor
        Case "Lamp"
            Set MSComm = MSCommLamp
        Case Else
            MsgBox "Device not found: " & DeviceName, vbExclamation
            Exit Function
    End Select
    
    If MSComm.PortOpen Then
        Debug.Print "Serial port already open for device: " & DeviceName
        OpenMSComm = True
        Exit Function
    End If
    
    
    With MSComm
        .CommPort = 1
        .RThreshold = 1
        .SThreshold = 0
    End With
     On Error GoTo ERRLINE:
    
    For Each deviceIndex In deviceInfo.Keys
        If deviceInfo(deviceIndex)("DeviceName") = DeviceName And deviceInfo(deviceIndex)("IsUsed") = True Then
            portNumber = deviceInfo(deviceIndex)("Port")
            MSComm.CommPort = CInt(Mid(portNumber, 4))
            MSComm.Settings = deviceInfo(deviceIndex)("BaudRate") & _
                              ",N," & deviceInfo(deviceIndex)("DataBits") & _
                              "," & deviceInfo(deviceIndex)("StopBits")
            MSComm.PortOpen = True
            If MSComm.PortOpen Then
'                Debug.Print "Serial port opened for device: " & DeviceName
                openedPorts.Add portNumber
                foundDevice = True
                Exit For
            Else
                MsgBox "Failed to open serial port for device: " & DeviceName, vbExclamation
            End If
        End If
    Next deviceIndex
    
    If Not foundDevice Then
        MsgBox "Device not found or is not marked as used: " & DeviceName, vbExclamation
    End If
    
    OpenMSComm = foundDevice
    Exit Function
ERRLINE:
'    MsgBox "Modbus Connect Fail!"
    ShowAlarmFlash 32
End Function
Function IsPortOpened(openedPorts As Collection, portNumber As Variant) As Boolean
    On Error Resume Next
'    IsPortOpened = (openedPorts.item(portNumber) <> "")
    Dim item As Variant
    For Each item In openedPorts
        If item = portNumber Then
            IsPortOpened = True
            Exit Function
        End If
    Next
    IsPortOpened = False
End Function

Function CloseMSComm(DeviceName As String) As Boolean

     Dim MSComm As MSComm
     Dim deviceIndex As Variant
     Dim foundDevice As Boolean
     Dim portNumber As String
     foundDevice = False
     
     Select Case DeviceName
        Case "SCR"
            Set MSComm = MSCommSCR
        Case "O2 Sensor"
            Set MSComm = MSCommO2SenSor
        Case "Lamp"
            Set MSComm = MSCommLamp
        Case Else
            MsgBox "Device not found: " & DeviceName, vbExclamation
            Exit Function
    End Select
 On Error GoTo ERRLINE:
    For Each deviceIndex In deviceInfo.Keys
''        If deviceInfo(deviceIndex)("DeviceName") = DeviceName And deviceInfo(deviceIndex)("IsUsed") = True Then
         If deviceInfo(deviceIndex)("DeviceName") = DeviceName Then
            portNumber = deviceInfo(deviceIndex)("Port")
            If IsPortOpened(openedPorts, portNumber) Then
'               msComm.CommPort = CInt(Mid(portNumber, 4))
               MSComm.PortOpen = False
                   If Not MSComm.PortOpen Then
                        openedPorts.Remove 1
                        foundDevice = True
                   Exit For
                   Else
                       MsgBox "Seial is not open" & portNumber, vbExclamation
                   End If
            End If
         End If
    Next deviceIndex
    
    CloseMSComm = foundDevice
    Exit Function
ERRLINE:
'    MsgBox "Modbus DisConnect Fail!"
End Function

Public Sub WriteConfig(DeviceName As String, IsUsed As Integer, Port As String, BaudRate As String, DataBits As String, StopBits As String, Timeout As String)
    Dim configFileName As String
    Dim fileContent As String
    Dim lines() As String
    Dim deviceSection As String
    Dim i As Integer
    Dim deviceFound As Boolean
    Dim deviceIndex As Integer
    
    
    configFileName = gbSystemPath & "\Config\ModbusRtu.ini"
    
    Open configFileName For Input As #1
    fileContent = Input$(LOF(1), #1)
    Close #1
    
    lines = Split(fileContent, vbCrLf)
    
    deviceFound = False
    
    For i = 0 To UBound(lines)
        
        If Left(Trim(lines(i)), Len("[Device")) = "[Device" Then
            
            deviceIndex = Val(Mid(Trim(lines(i)), Len("[Device") + 1, InStr(1, lines(i), "]") - (Len("[Device") + 1)))
            
            
            deviceSection = "[" & Trim(lines(i)) & "]"
            
            
'            If deviceSection = "[" & "Device" & CStr(deviceIndex) & "]" Then
                Dim currentDeviceName As String
                currentDeviceName = Trim(Mid(lines(i + 1), Len("DeviceName=") + 1))
                If currentDeviceName = DeviceName Then
                lines(i + 1) = "DeviceName=" & DeviceName
                lines(i + 2) = "IsUsed=" & IsUsed
                lines(i + 3) = "Port=" & Port
                lines(i + 4) = "BaudRate=" & BaudRate
                lines(i + 5) = "DataBits=" & DataBits
                lines(i + 6) = "StopBits=" & StopBits
                lines(i + 7) = "Timeout=" & Timeout
                deviceFound = True
                Exit For
                End If
            End If
'        End If
    Next i
    
     If Not deviceFound Then
       
        deviceIndex = 1
        Do While InStr(fileContent, "[Device" & CStr(deviceIndex) & "]") > 0
            deviceIndex = deviceIndex + 1
        Loop
        
        
        ReDim Preserve lines(UBound(lines) + 7)
        lines(UBound(lines) - 7) = "[Device" & deviceIndex & "]"
        lines(UBound(lines) - 6) = "DeviceName=" & DeviceName
        lines(UBound(lines) - 5) = "IsUsed=" & IsUsed
        lines(UBound(lines) - 4) = "Port=" & Port
        lines(UBound(lines) - 3) = "BaudRate=" & BaudRate
        lines(UBound(lines) - 2) = "DataBits=" & DataBits
        lines(UBound(lines) - 1) = "StopBits=" & StopBits
        lines(UBound(lines)) = "Timeout=" & Timeout
    End If
    
    
    Open configFileName For Output As #1
    Print #1, Join(lines, vbCrLf)
    Close #1
End Sub

Public Sub InitModBusRtu(DeviceName As String)
    ReadConfig
    If Not IsEmpty(deviceInfo) And deviceInfo.Count > 0 Then
        OpenMSComm (DeviceName)
    Else
         MsgBox "no Config", vbExclamation
    End If
End Sub

'Dim ReadingO2Sensor As Boolean
'Private Sub Timer1_Timer()
'    If ReadingO2Sensor Then
'        Dim ReadCommand() As Byte
'        ReadCommand = GenerateReadCommand(&H1, 100, 1, &H4)
'        Call sendData(ReadCommand, "O2 Sensor")
'
'        If UBound(Response) >= 0 Then
'            If UBound(Response) >= LBound(Response) Then
'                Text5.text = CStr(Response(0) * 256 + Response(1))
'            End If
'        End If
'    End If
'End Sub


Public Sub WriteRamupSCR()
     Dim writeCommand() As Byte
     Dim i As Integer
     
    Dim upperLimit As Integer
    upperLimit = IIf(gbintNumOfBanks > 17, 17, gbintNumOfBanks)
    Call SCRAddress
'    upperLimit = 5
    For i = 0 To upperLimit - 1
    writeCommand = GenerateWriteCommand(gbsngSCRAddress(i), 131, gbsngRecipeIntensityWeightDynamic(i))
    Call sendData(writeCommand, "SCR")
'    Sleep (2)
'    If UBound(Response) >= 0 Then
'        If Response(0) = &H0 Then
''            MsgBox "Write MaxPower Succeeded!"
''            WriteLog ("WriteRamupSCR MaxPower" & writeCommand(i) & " Succeeded!")
'        Else
'            MsgBox Response(0)
'        End If
'    Else
'        MsgBox "Write MaxPower failed!"
'    End If
     DoEvents
    Next i
End Sub

Public Sub WriteHoldSCR()
    Dim writeCommand() As Byte
     Dim i As Integer
     
    Dim upperLimit As Integer
    upperLimit = IIf(gbintNumOfBanks > 17, 17, gbintNumOfBanks)
    Call SCRAddress
'     upperLimit = 5
    For i = 0 To upperLimit - 1
    writeCommand = GenerateWriteCommand(gbsngSCRAddress(i), 131, gbsngRecipeIntensityWeightSteady(i))
    Call sendData(writeCommand, "SCR")

'    If UBound(Response) >= 0 Then
'        If Response(0) = &H0 Then
''            MsgBox "Write MaxPower Succeeded!"
'            WriteLog ("WriteHoldSCR MaxPower" & i & " Succeeded!")
'        End If
'    End If
'    Sleep (2)
   DoEvents
    Next i
End Sub
Private Sub SCRAddress()
    gbsngSCRAddress(0) = &H1
    gbsngSCRAddress(1) = &H2
    gbsngSCRAddress(2) = &H3
    gbsngSCRAddress(3) = &H4
    gbsngSCRAddress(4) = &H5
    gbsngSCRAddress(5) = &H6
    gbsngSCRAddress(6) = &H7
    gbsngSCRAddress(7) = &H8
    gbsngSCRAddress(8) = &H9
    gbsngSCRAddress(9) = &HA
    gbsngSCRAddress(10) = &HB
    gbsngSCRAddress(11) = &HC
'    gbsngSCRAddress(12) = &HD
'    gbsngSCRAddress(13) = &HE
'    gbsngSCRAddress(14) = &HF
'    gbsngSCRAddress(15) = &H16
'    gbsngSCRAddress(16) = &H17
    
End Sub
