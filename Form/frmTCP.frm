VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTCP 
   Caption         =   "TCP"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
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
   ScaleHeight     =   4935
   ScaleWidth      =   7185
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox chkReceive 
      Caption         =   "Check1"
      Height          =   270
      Left            =   5640
      TabIndex        =   5
      Top             =   2040
      Width           =   255
   End
   Begin VB.Timer tmrSendStatus 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6240
      Top             =   3240
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   3240
   End
   Begin VB.TextBox txtReceive 
      Height          =   3015
      Left            =   600
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   4
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Text            =   "SPW1"
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Text            =   "502"
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "192.168.0.171"
      Top             =   480
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wsTemp 
      Left            =   5880
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.0.171"
      RemotePort      =   502
      LocalPort       =   171
   End
End
Attribute VB_Name = "frmTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytBinary(0 To 11) As Byte
Dim intTimeoutCount As Integer

Private Sub cmdSend_Click()
    wsServer.SendData txtSend.Text
    
End Sub

Private Sub Form_Initialize()
    '
End Sub

Private Sub Form_Load()
    '
End Sub

Public Sub OpenTCP()
On Error GoTo ERRLINE:
    If Kernel.IsTcpTempConnect = 0 Then
        wsTemp.RemoteHost = txtRemoteIP.Text
        wsTemp.RemotePort = 502
        wsTemp.Connect
        
        'tmrSendStatus.Enabled = True
        Kernel.IsTcpTempConnect = 1
        intTimeoutCount = 0
    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " OpenTCP error"
    ShowAlarmFlash 1
    
End Sub

Public Sub CloseTCP()

On Error GoTo ERRLINE:
    tmrSendStatus.Enabled = False
    wsTemp.Close
    Kernel.IsTcpTempConnect = 0
    intTimeoutCount = 0
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " CloseTCP error"
    ShowAlarmFlash 1
    
End Sub

Public Sub ReadTC()

On Error GoTo ERRLINE:
    bytBinary(0) = 1
    bytBinary(1) = 2
    bytBinary(2) = 0
    bytBinary(3) = 0
    bytBinary(4) = 0
    bytBinary(5) = 6
    bytBinary(6) = 1
    bytBinary(7) = 4    'Function Code 4 to read AI registers
    bytBinary(8) = 0
    bytBinary(9) = 0
    bytBinary(10) = 0
    bytBinary(11) = 10  '1 counter occupies 2 AI registers,
                        '8 counters occupy 16 AI registers.
    
    If wsTemp.state = sckConnected Then
        wsTemp.SendData bytBinary
        tmrTimeout.Enabled = True
        chkReceive.value = 0
    End If
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " TCP ReadTC error"
    ShowAlarmFlash 1
    
End Sub


Private Sub tmrSendStatus_Timer()
    
On Error GoTo ERRLINE:
    bytBinary(0) = 1
    bytBinary(1) = 2
    bytBinary(2) = 0
    bytBinary(3) = 0
    bytBinary(4) = 0
    bytBinary(5) = 6
    bytBinary(6) = 1
    bytBinary(7) = 4    'Function Code 4 to read AI registers
    bytBinary(8) = 0
    bytBinary(9) = 0
    bytBinary(10) = 0
    bytBinary(11) = 10  '1 counter occupies 2 AI registers,
                        '8 counters occupy 16 AI registers.
    
    If wsTemp.state = sckConnected Then
        wsTemp.SendData bytBinary
        tmrTimeout.Enabled = True
    End If
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " TCP Send error"
    ShowAlarmFlash 1
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
    ReadTC

End Sub

Private Sub wsTemp_DataArrival(ByVal bytesTotal As Long)
    Dim bytBuf() As Byte
    Dim i As Integer
    Dim Counter As Single         'add on Dec 20,2013
    
    ReDim bytBuf(bytesTotal - 1)
    
    tmrTimeout.Enabled = False
    intTimeoutCount = 0
    wsTemp.GetData bytBuf, vbByte + vbArray
    
    Select Case bytBuf(7) 'Function Call
'        Case 1 'read DO status
'             For i = 0 To 7
'                If bytBuf(9) And Bitmask(i) Then
'                    shpDO(i).BackColor = vbGreen
'                Else
'                    shpDO(i).BackColor = vbWhite
'                End If
'             Next
'        Case 2 'read DI status
'             For i = 0 To 7
'                If bytBuf(9) And Bitmask(i) Then
'                     shapDI(i).BackColor = vbRed
'                Else
'                    shapDI(i).BackColor = vbWhite
'                End If
'             Next
        Case 4 'read input register
            For i = 0 To 9
                Kernel.sngTC(i) = CSng((bytBuf(i * 2 + 9)) * 256 + CSng(bytBuf(i * 2 + 10))) / 10
            Next i
            frmPlotProcess.ShowStatus
            
'            txtCounter(i).Text = Counter
'            For i = 0 To 7
'                Counter = CCur(bytBuf(i * 4 + 9)) * 256 ^ 3 + CCur(bytBuf(i * 4 + 1 + 9)) * 256 ^ 2 + CCur(bytBuf(i * 4 + 2 + 9)) * 256 + CCur(bytBuf(i * 4 + 3 + 9))
'                                'add on Dec 20,2013
'                txtCounter(i).Text = Counter
'            Next i
    End Select
End Sub


