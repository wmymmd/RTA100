VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAz1 
   Caption         =   "Azbil"
   ClientHeight    =   7455
   ClientLeft      =   -4485
   ClientTop       =   555
   ClientWidth     =   16530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   16530
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   0
      Width           =   1815
   End
   Begin VB.Timer tmrWriteProc 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer tmrProc 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   0
   End
   Begin VB.CommandButton cmdReadProcNo 
      Caption         =   "Read"
      Height          =   495
      Left            =   13440
      TabIndex        =   17
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdWriteProcNo 
      Caption         =   "Write"
      Height          =   495
      Left            =   14640
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdWriteProc 
      Caption         =   "Write"
      Height          =   495
      Left            =   14640
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdReadRAM 
      Caption         =   "Read"
      Height          =   495
      Left            =   7920
      TabIndex        =   13
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdReadProc 
      Caption         =   "Read"
      Height          =   495
      Left            =   13440
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   495
      Left            =   10200
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmbClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtData 
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Text            =   "0"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CheckBox chkShowPolling 
      Caption         =   "Show Polling"
      Height          =   270
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Timer tmrSendStatus 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   0
   End
   Begin VB.TextBox txtReceive 
      Height          =   6135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1080
      Width           =   6375
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtAdd 
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Text            =   "101"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Text            =   "502"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "192.168.0.11"
      Top             =   480
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wsTemp 
      Left            =   3240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.0.171"
      RemotePort      =   502
      LocalPort       =   17000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAzData 
      Height          =   6135
      Left            =   6600
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   17
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAzPara 
      Height          =   6135
      Left            =   9960
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   17
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSWinsockLib.Winsock wsReadPara 
      Left            =   3840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.0.171"
      RemotePort      =   502
      LocalPort       =   17001
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAzProc 
      Height          =   2775
      Left            =   13200
      TabIndex        =   11
      Top             =   4560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   17
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAzProcNo 
      Height          =   2775
      Left            =   13200
      TabIndex        =   15
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   17
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmAz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnStopRead As Boolean

Dim sngParaData(0 To 100) As Single
Dim sngAzibilData(0 To 100) As Single
Dim sngRamData(0 To 100) As Single
Dim sngProcData(0 To 100) As Single
Dim sngProcNoData(0 To 100) As Single
Dim intTimeoutCount As Integer
Dim intReceiveCount As Integer
Dim intProcNo As Integer
Dim lnCurrPos As Long
Dim blnShowRespons As Boolean
Dim blnBusy As Boolean
Dim blnGetPara As Boolean
Dim bFirstRun As Boolean


Private Sub cmdWriteProc_Click()
    
    CurrProc.sngTemperature = 101
    CurrProc.lngCurrStepTime = 11

    'WriteProc CurrProc.sngTemperature, CurrProc.lngCurrStepTime, 202, 22
End Sub


Public Sub RunProc(iType As Integer)
    Dim i As Integer
    Dim Data(11) As Integer
    Dim ProcData(20) As Integer
    Dim HoldData(3) As Integer
    
    
    Select Case iType
        Case 0
            For i = 0 To 3
                Data(i) = 1
            Next i
            For i = 4 To 7
                Data(i) = 0
            Next i
            For i = 8 To 11
                Data(i) = 0
            Next i
            Call WriteParas(109, Data, False)
        Case 1
'            ProcData(0) = Az1.intTemp1
'            ProcData(1) = Az1.intTime1
'            ProcData(2) = 1
'            ProcData(3) = 0
'            ProcData(4) = 1
'
'            ProcData(10) = Az1.intTemp2
'            ProcData(11) = Az1.intTime2
'            ProcData(12) = 1
'            ProcData(13) = 0
'            ProcData(14) = 1
'            Call WriteParas(48100, ProcData, True)
            
            If Kernel.IsRun = 1 Then
                For i = 0 To 3
                    If Az1.blnUseLoop(i) = True Then
                        Data(i) = 0
                        Data(i + 4) = 0
                    Else
                        Data(i) = 1
                        Data(i + 4) = 1
                    End If
                Next i
                For i = 8 To 11
                    Data(i) = 0
                Next i
                Call WriteParas(109, Data, False)
                
            End If
        Case 2  'hold stop
            For i = 0 To 3
                HoldData(i) = 0
            Next i
            Call WriteParas(121, HoldData, False)
        Case 3  'hold start
            For i = 0 To 3
                If Az1.blnUseLoop(i) = True Then
                    HoldData(i) = 1
                End If
            Next i
            Call WriteParas(121, HoldData, False)
        Case 4  'AT stop
            For i = 0 To 3
                HoldData(i) = 0
            Next i
            Call WriteParas(125, HoldData, False)
        Case 5  'AT start
            For i = 0 To 3
                If Az1.blnUseLoop(i) = True Then
                    HoldData(i) = 1
                End If
            Next i
            Call WriteParas(125, HoldData, False)
        Case 6  'Write Ratio
            For i = 0 To 3
                HoldData(i) = Az1.sngRT(i) * 10000
            Next i
            Call WriteParas(213, HoldData, False)
        Case Else
            For i = 0 To 7
                Data(i) = 1
            Next i
            For i = 8 To 11
                Data(i) = 0
            Next i
            Call WriteParas(109, Data, True)
            
            For i = 0 To 3
                If Az1.blnUseLoop(i) = True Then
                    Data(i + 8) = iType
                Else
                    Data(i + 8) = 0
                End If
            Next i
            Call frmAz1.WriteParas(109, Data, False)
    End Select
    
End Sub


Public Function WriteProc(Temp1 As Integer, Time1 As Integer, Temp2 As Integer, Time2 As Integer) As Boolean
    Dim i As Integer
    Dim Data(0 To 20) As Integer

    Data(0) = Temp1
    Data(1) = Time1
    Data(2) = 1
    Data(3) = 0
    Data(4) = 1

    Data(10) = Temp2
    Data(11) = Time2
    Data(12) = 1
    Data(13) = 0
    Data(14) = 1

    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i

    If WriteParas(48100, Data, True) Then
        For i = 0 To 14
            fgAzProc.TextMatrix(i + 1, 1) = CStr(Data(i))
        Next i
    Else
        gbstrAlarmHint = " Az1 WriteProc Error"
        ShowAlarmFlash 1
        blnBusy = False
    End If
    
End Function

Private Sub cmbClear_Click()
    txtReceive.text = ""
    intReceiveCount = 0
End Sub


Private Sub cmdRead_Click()
    Dim i As Integer
    Dim sngData(0 To 100) As Single
    
    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i
    
    If ReadParas(201, sngData(), True) Then
            
        For i = 0 To 99
            fgAzPara.TextMatrix(i + 1, 1) = CStr(sngData(i))
        Next i
        Exit Sub
    End If
    
ERRLINE:
    
    gbstrAlarmHint = " Azbil ReadParas Error"
    ShowAlarmFlash 1
    blnBusy = False
    
End Sub

Private Sub cmdReadProcNo_Click()
    Dim i As Integer
    Dim sngData(0 To 100) As Single
    
    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i
    
    If ReadParas(48000, sngData(), True) Then
        For i = 0 To 99
            fgAzProcNo.TextMatrix(i + 1, 1) = CStr(sngData(i))
        Next i
        Exit Sub
    End If
    
ERRLINE:
    gbstrAlarmHint = " Azbil ReadParas Error"
    ShowAlarmFlash 1
    blnBusy = False
End Sub


Private Sub cmdReadProc_Click()
    Dim i As Integer
    Dim sngData(0 To 100) As Single
    
    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i
    
    If ReadParas(48100, sngData(), True) Then
        For i = 0 To 99
            fgAzProc.TextMatrix(i + 1, 1) = CStr(sngData(i))
        Next i
        Exit Sub
    End If
    
ERRLINE:
    gbstrAlarmHint = " Azbil ReadParas Error"
    ShowAlarmFlash 1
    blnBusy = False
        
End Sub

Private Sub cmdReadRAM_Click()
    Dim i As Integer
    Dim sngData(0 To 100) As Single
    
    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i
    
    If ReadParas(101, sngData(), True) Then
        For i = 0 To 99
            fgAzData.TextMatrix(i + 1, 1) = CStr(sngData(i))
        Next i
        Exit Sub
    End If
    
ERRLINE:
    gbstrAlarmHint = " Azbil ReadParas Error"
    ShowAlarmFlash 1
    blnBusy = False
        
End Sub

Private Sub cmdWrite_Click()
'    Call WritePara(109, CInt(txtData.Text))
    Call WritePara(CLng(txtAdd.text), CInt(txtData.text))
'    Dim ii(0 To 2) As Integer
'
'    ii(0) = 6
'    ii(1) = 6
'    ii(2) = 6

'    Call WriteParas(CInt(txtAdd.Text), ii)
End Sub



Private Sub Command1_Click()
    RunProc (3)
End Sub

Private Sub Form_Initialize()
    '
End Sub

Private Sub Form_Load()
    Dim BB(0 To 1) As Byte
    Dim i As Integer
    
    
    With fgAzData
        .Cols = 2
        .Rows = 101
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .TextMatrix(0, 0) = "Address"
        .TextMatrix(0, 1) = "Value"
        
        For i = 1 To 100
            .TextMatrix(i, 0) = CStr(i + 100) & "     "
            .TextMatrix(i, 1) = "0"
        Next i
        'SngToByte (1.02)
        'Call SngToAzbil(280, 1, bb)
        'Call WriteData(116, 1)
    End With
    
    With fgAzPara
        .Cols = 2
        .Rows = 101
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .TextMatrix(0, 0) = "Address"
        .TextMatrix(0, 1) = "Value"
        
        For i = 1 To 100
            .TextMatrix(i, 0) = CStr(i + 200) & "     "
            .TextMatrix(i, 1) = "0"
        Next i
    End With
    
    With fgAzProc
        .Cols = 2
        .Rows = 101
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .TextMatrix(0, 0) = "Address"
        .TextMatrix(0, 1) = "Value"
        
        For i = 1 To 100
            .TextMatrix(i, 0) = CStr(i + 48099)
            .TextMatrix(i, 1) = "0"
        Next i
        
    End With
    
    With fgAzProcNo
        .Cols = 2
        .Rows = 101
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .TextMatrix(0, 0) = "Address"
        .TextMatrix(0, 1) = "Value"
        
        For i = 1 To 100
            .TextMatrix(i, 0) = CStr(i + 47999)
            .TextMatrix(i, 1) = "0"
        Next i
        
    End With
    
    
    blnShowRespons = False
    
End Sub

Public Function OpenTCP(sIP As String) As Boolean

On Error GoTo ERRLINE:
    If Kernel.IsTcpTempConnect = 0 Then
        If sIP = "" Then sIP = txtRemoteIP.text
        wsReadPara.RemoteHost = sIP
        wsReadPara.RemotePort = 502
        wsReadPara.Connect
        
        bFirstRun = True
        gbintAz1ProcNo = 0
        tmrSendStatus.Enabled = True
        intTimeoutCount = 0
    End If
    OpenTCP = True
    Exit Function
ERRLINE:
    OpenTCP = False
    If mdlKernel.Az1_ConNum = 2 Then
    gbstrAlarmHint = " Az1 OpenTCP error"
    ShowAlarmFlash 1
    End If
    
End Function

Public Sub CloseTCP()

On Error GoTo ERRLINE:
    tmrSendStatus.Enabled = False
    wsReadPara.Close
    Kernel.IsTcpTempConnect = 0
    intTimeoutCount = 0
    Exit Sub
ERRLINE:

    gbstrAlarmHint = " CloseTCP error"
    ShowAlarmFlash 1
    
End Sub

Sub SngToByte(ByVal D As Single)
    Dim Bytes(LenB(D) - 1) As Byte
    Dim i As Integer
    Dim S As String

    CopyMemory Bytes(0), D, LenB(D)

    For i = 0 To UBound(Bytes)
        S = S & CStr(Bytes(i)) & " "
    Next
    
End Sub



Sub SngToAzbil(ByVal sng As Single, ByVal Floats As Integer, b() As Byte)

    Dim value As Long
    'Dim b(0 To 1) As Byte
    
    value = sng * Floats
    b(0) = CByte(value Mod 256)
    b(1) = CByte((value / 256) And &HFF)
    
    'SngToAzbil = b
End Sub


Public Function ReadProc(Data() As Single) As Boolean
    Dim i As Integer
    Dim sngData(0 To 100) As Single
    
    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i
    
    If ReadParas(48100, sngData(), True) Then
        For i = 0 To 99
            fgAzProc.TextMatrix(i + 1, 1) = CStr(sngData(i))
        Next i
        ReadProc = True
        Exit Function
    End If
    
    ReadProc = False
ERRLINE:
    gbstrAlarmHint = " Azbil ReadParas Error"
    ShowAlarmFlash 1
    blnBusy = False
End Function

Public Function ReadPID(Data() As Single) As Boolean
        
    Dim i As Integer
        
    For i = 0 To 20000
        If blnBusy = False Then
            Exit For
        End If
        DoEvents
    Next i
    
    If ReadParas(201, Data(), True) Then
        For i = 0 To 99
            fgAzPara.TextMatrix(i + 1, 1) = CStr(Data(i))
        Next i
        ReadPID = True
        Exit Function
    End If
    
    ReadPID = False
ERRLINE:
    gbstrAlarmHint = " Azbil ReadPara Error"
    ShowAlarmFlash 1
    blnBusy = False
End Function

Public Function ReadParas(lnStartPos As Long, Data() As Single, IsWait As Boolean) As Boolean
    Dim bytBinary(0 To 11) As Byte
    Dim bytPos(0 To 1) As Byte
    Dim S As String
    Dim s1 As String
    Dim i As Integer
    Dim j As Integer
    
    
    If blnBusy = True Then
        ReadParas = False
        Exit Function
    End If
    
    blnBusy = True
    
    lnCurrPos = lnStartPos
'    tmrSendStatus.Enabled = False
    blnShowRespons = True
    blnGetPara = False
On Error GoTo ERRLINE:
    Call CopyMemory(bytPos(0), lnStartPos, 2)
    
    bytBinary(0) = 1
    bytBinary(1) = 2
    bytBinary(2) = 0
    bytBinary(3) = 0
    bytBinary(4) = 0
    bytBinary(5) = 6
    bytBinary(6) = 0
    bytBinary(7) = 3    'Function Code 3 to read Holding registers
    bytBinary(8) = bytPos(1)
    bytBinary(9) = bytPos(0)  'From RegNo=100
    bytBinary(10) = 0
    bytBinary(11) = 30  'Number of Regs to read
    
    If wsReadPara.state = sckConnected Then
        wsReadPara.SendData bytBinary
        'tmrTimeout.Enabled = True
    Else
        GoTo ERRLINE:
    End If
        
    If chkShowPolling.value Then
        intReceiveCount = intReceiveCount + 1
        If intReceiveCount > 50 Then
            txtReceive.text = ""
            intReceiveCount = 0
        End If

        S = "S>"
        For i = 0 To 11
            S = S & Format(Hex(bytBinary(i)), "00") & " "
        Next i
        S = S + vbCr + vbLf
        txtReceive.text = txtReceive.text + S
    End If

    If IsWait = True Then
    
        For i = 0 To 30000
            If blnGetPara = True Then Exit For
            DoEvents
        Next i
        If blnGetPara = False Then
            For j = 0 To 1
                s1 = "Retry " & CStr(j) & ">" & S
                txtReceive.text = txtReceive.text + s1
                wsReadPara.SendData bytBinary
    
                For i = 0 To 20000
                    If blnGetPara = True Then Exit For
                    DoEvents
                Next i
                If blnGetPara = True Then Exit For
            Next j
    
            If blnGetPara = False Then GoTo ERRLINE:
    
        End If
    
        If lnStartPos = 101 Then
            For i = 0 To UBound(Data)
                Data(i) = sngRamData(i)
            Next
        ElseIf lnStartPos = 201 Then
            For i = 0 To UBound(Data)
                Data(i) = sngParaData(i)
            Next
        ElseIf lnStartPos = 48000 Then
            For i = 0 To UBound(Data)
                Data(i) = sngProcNoData(i)
            Next
        Else
            For i = 0 To UBound(Data)
                Data(i) = sngProcData(i)
            Next
        End If
    End If
    
    ReadParas = True
    Exit Function
ERRLINE:
    gbstrAlarmHint = " Az1 ReadParas error"
    ShowAlarmFlash 1
    ReadParas = False
End Function


Public Function WritePara(address As Long, Data As Integer) As Boolean

    Dim bytSend(0 To 14) As Byte
    Dim bytAdd(0 To 1) As Byte
    Dim bytData(0 To 1) As Byte
    Dim S As String
    Dim s1 As String
    Dim i As Integer
    Dim j As Integer
    
On Error GoTo ERRLINE:
    blnShowRespons = True
    blnBusy = True
    
    Call CopyMemory(bytAdd(0), address, 2)
    Call CopyMemory(bytData(0), Data, 2)
    
    bytSend(0) = 1
    bytSend(1) = 2
    bytSend(2) = 0
    bytSend(3) = 0
    bytSend(4) = 0
    bytSend(5) = 9
    bytSend(6) = 0
    bytSend(7) = 16    'Function Code 6 to write Holding registers
    bytSend(8) = bytAdd(1)
    bytSend(9) = bytAdd(0)  'From RegNo=100
    bytSend(10) = 0
    bytSend(11) = 1 'Regs
    bytSend(12) = 2 'bytes=Regs *2
    bytSend(13) = bytData(1)
    bytSend(14) = bytData(0)  'Data value

    If wsReadPara.state = sckConnected Then
        wsReadPara.SendData bytSend
        blnGetPara = False
    Else

        GoTo ERRLINE:
    End If
        
    If chkShowPolling.value Or blnShowRespons = True Then
        intReceiveCount = intReceiveCount + 1
        If intReceiveCount > 50 Then
            txtReceive.text = ""
            intReceiveCount = 0
        End If
        
        
        S = "S>"
        For i = 0 To 14
            S = S & Format(Hex(bytSend(i)), "00") & " "
        Next i
        S = S + vbCr + vbLf
        txtReceive.text = txtReceive.text + S
    End If
    
    For i = 0 To 30000
        If blnGetPara = True Then Exit For
        DoEvents
    Next i
    
    If blnGetPara = False Then
        For j = 0 To 1
            s1 = "Retry " & CStr(j) & ">" & S
            txtReceive.text = txtReceive.text + s1
            wsReadPara.SendData bytSend
            
            For i = 0 To 10000
                If blnGetPara = True Then Exit For
                DoEvents
            Next i
            If blnGetPara = True Then Exit For
        Next j
        
        If blnGetPara = False Then GoTo ERRLINE:
     
    End If
    WritePara = True
    Exit Function
ERRLINE:
    WritePara = False
    gbstrAlarmHint = " Az1 WritePara error"
    ShowAlarmFlash 1
    
End Function


Public Function WriteParas(address As Long, Data() As Integer, IsWait As Boolean) As Boolean

    Dim bytSend() As Byte
    Dim bytData() As Byte
    Dim bytAdd(0 To 1) As Byte
    Dim bytLen(0 To 1) As Byte
    Dim DLen As Integer
    Dim i As Integer
    Dim j As Integer
    Dim S As String
    Dim s1 As String
                
On Error GoTo ERRLINE:

    blnShowRespons = True
    blnBusy = True
    tmrSendStatus.Enabled = False
    'Sleep 100
    
    DLen = UBound(Data) + 1
    ReDim bytSend(13 + DLen * 2 - 1)
    ReDim bytData(DLen * 2 - 1)
    
    Call CopyMemory(bytAdd(0), address, 2)
    Call CopyMemory(bytLen(0), DLen, 2)
    Call CopyMemory(bytData(0), Data(0), DLen * 2)

    bytSend(0) = 1
    bytSend(1) = 2
    bytSend(2) = 0
    bytSend(3) = 0
    bytSend(4) = 0
    bytSend(5) = 9
    bytSend(6) = 0
    bytSend(7) = 16    'Function Code 16 to write Holding registers
    bytSend(8) = bytAdd(1)
    bytSend(9) = bytAdd(0)  'From RegNo=100
    bytSend(10) = bytLen(1)
    bytSend(11) = bytLen(0) 'Regs
    bytSend(12) = DLen * 2 'bytes=Regs *2
    
    'Call CopyMemory(bytSend(13), bytData(0), DLen * 2)
    'Call CopyMemory(bytSend(13), Data(0), DLen * 2)
            
    For i = 0 To DLen * 2 - 2 Step 2
        bytSend(13 + i) = bytData(i + 1)
        bytSend(14 + i) = bytData(i)
    Next i
    
    
    If wsReadPara.state = sckConnected Then
        wsReadPara.SendData bytSend
        blnGetPara = False
    Else
        GoTo ERRLINE:
    End If
        
    If chkShowPolling.value Or blnShowRespons = True Then
        S = "S>"
        For i = 0 To 12 + DLen * 2
            S = S & Format(Hex(bytSend(i)), "00") & " "
        Next i
        S = S + vbCr + vbLf
        txtReceive.text = txtReceive.text + S
    End If
    
    If IsWait = True Then
        For i = 0 To 30000
            If blnGetPara = True Then Exit For
            DoEvents
        Next i
    
        If blnGetPara = False Then
            For j = 0 To 1
                s1 = "Retry " & CStr(j) & ">" & S
                txtReceive.text = txtReceive.text + s1
                wsReadPara.SendData bytSend
    
                For i = 0 To 30000
                    If blnGetPara = True Then Exit For
                    DoEvents
                Next i
                If blnGetPara = True Then Exit For
            Next j
    
            If blnGetPara = False Then GoTo ERRLINE:
    
        End If
    End If

    tmrSendStatus.Enabled = True
    WriteParas = True
    Exit Function
ERRLINE:
    WriteParas = False
    If IsInitAzbil = True Then
     If WriteErr_Count = 2 And Write_Result = False Then
        gbstrAlarmHint = " Az1 WriteParas error" + TCM1_WriteErrDetail
        WriteLog ("Az1報錯:" & gbstrAlarmHint)
        ShowAlarmFlash 1
    End If
    Else
        gbstrAlarmHint = " Az1 WriteParas error"
        ShowAlarmFlash 1
    End If
End Function
Private Sub tmrWriteProc_Timer()
    tmrWriteProc.Enabled = False
    
    Call WriteProc(Az1.intTemp1, Az1.intTime1, Az1.intTemp2, Az1.intTime2)
End Sub


Private Sub tmrSendStatus_Timer()
    Dim i As Integer
    Dim Data(99) As Single
    
    If gbintAz1ProcNo = -1 Then
        Call ReadParas(101, Data(), False)
    Else
        Call frmAz1.RunProc(gbintAz1ProcNo)
        gbintAz1ProcNo = -1
    End If
'    Select Case gbintAz1ProcNo
'    Case -1
'        Call ReadParas(101, Data(), False)
'    Case 0
'        Call frmAz1.RunProc(0)
'        gbintAz1ProcNo = -1
'    Case 1
'        Call frmAz1.RunProc(1)
'        gbintAz1ProcNo = -1
'    Case 2
'        Call frmAz1.RunProc(2)
'        gbintAz1ProcNo = -1
'    Case 3
'        Call frmAz1.RunProc(3)
'        gbintAz1ProcNo = -1
'    Case 4
'        Call frmAz1.RunProc(4)
'        gbintAz1ProcNo = -1
'    Case 5
'        Call frmAz1.RunProc(5)
'        gbintAz1ProcNo = -1
'    Case Else
'        Call frmAz1.RunProc(gbintAz1ProcNo)
'        gbintAz1ProcNo = -1
'    End Select
    
End Sub

Private Sub tmrTimeout_Timer()
    
    If intTimeoutCount > 10 Then
        tmrSendStatus.Enabled = False
        tmrTimeout.Enabled = False
        intTimeoutCount = 0
        gbstrAlarmHint = " TCP TC Timeout"
        ShowAlarmFlash 1
        Exit Sub
    End If
    intTimeoutCount = intTimeoutCount + 1
    

End Sub


Private Sub wsReadPara_DataArrival(ByVal bytesTotal As Long)


    Dim bytBuf() As Byte
    Dim i As Integer
    Dim Counter As Single
    Dim RegNum As Integer
    Dim S As String
    
    ReDim bytBuf(bytesTotal - 1)
    
    
    tmrTimeout.Enabled = False
    intTimeoutCount = 0
    wsReadPara.GetData bytBuf, vbByte + vbArray
'     If bytBuf(7) <> 3 Then
'        WriteLog ("TCM RetrunCode:" + CStr(bytBuf(7)))
'     End If

    Select Case bytBuf(7) 'Function Call
        Case 3 'read DI status
             RegNum = bytBuf(8) / 2 - 1     '20 bytes = 10 Regs --> for  0~9
             
             For i = 0 To RegNum
'                sngParaData(i) = CSng((bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10)))
             
                If lnCurrPos = 101 Then
                    sngRamData(i) = CSng(bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10))
                    gbsngAz1Data(i) = sngRamData(i)
                ElseIf lnCurrPos = 201 Then
                    sngParaData(i) = CSng(bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10))
                    gbsngAz1Para(i) = sngParaData(i)
                ElseIf lnCurrPos = 48000 Then
                    sngProcNoData(i) = CSng(bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10))

                ElseIf lnCurrPos = 48100 Then
                    sngProcData(i) = CSng(bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10))

                End If
                
             Next
             
             
        Case 4 'read input register
'            For i = 0 To 9
'                Kernel.sngTC(i) = CSng((bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10))) / 10
'            Next i
'            frmPlotProcess.ShowStatus
        Case 144
            'gbstrAlarmHint = " WriteParas (invalid number)"
            'ShowAlarmFlash 1
'            MsgBox "資料格式錯誤!", vbOK
    End Select
        
    
    If chkShowPolling.value Then
        blnShowRespons = False
        S = "R>"
        For i = 0 To bytesTotal - 1
            S = S & Format(Hex(bytBuf(i)), "00") & " "
        Next i
        S = S + vbCr + vbLf + vbCr + vbLf
        txtReceive.text = txtReceive.text + S
        
        intReceiveCount = intReceiveCount + 1
        If intReceiveCount > 50 Then
            txtReceive.text = ""
            intReceiveCount = 0
        End If
    End If
    blnGetPara = True
    blnBusy = False
End Sub
