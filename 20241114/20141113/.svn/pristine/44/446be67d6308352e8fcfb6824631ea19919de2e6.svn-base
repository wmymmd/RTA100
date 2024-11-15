VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDCR 
   Caption         =   "DCR"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7245
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdScanID 
      Caption         =   "Scan ID"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Timer tmrTimuout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock wsDCR 
      Left            =   5640
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Text            =   "10.5.6.100"
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Text            =   "9600"
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "MEASURE"
      Top             =   720
      Width           =   4695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtReceive 
      Height          =   3015
      Left            =   240
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      Caption         =   "60"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label lblID 
      Alignment       =   2  '置中對齊
      Caption         =   "000000"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   4560
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Get ID:"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "frmDCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public gbintTimeout As Integer

Private Sub cmdScanID_Click()
    Scan True
End Sub

Private Sub cmdStop_Click()
    Scan False
End Sub

Private Sub Form_Load()
    'DCR.PortOpen = True
    
    
End Sub

Private Sub Form_Terminate()
    'DCR.PortOpen = False
End Sub

Public Sub OpenUDP()
On Error GoTo ERRLINE:
    Dim s1 As String
    
    wsDCR.LocalPort = 9601
    wsDCR.RemotePort = CInt(txtRemotePort.Text)
'    s1 = "10.5.6.100"
    wsDCR.RemoteHost = txtRemoteIP.Text
    
    wsDCR.Connect
    'wsDCR.Bind
    'wsDCR.Bind 9601, "10.5.6.101"
    
'    If Kernel.IsRemoteConnect = 0 Then
'        wsServer.RemoteHost = "192.168.0.168"
'        If Para.intAutoPort = 0 Then
'            wsServer.RemotePort = 168
'        Else
'            wsServer.LocalPort = Para.intAutoPort
'            wsServer.RemotePort = Para.intAutoPort
'        End If
'
'        wsServer.Connect
'        tmrTimeout.Enabled = True
'        Kernel.IsRemoteConnect = 1
'    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " UDP Open error"
    ShowAlarmFlash 21
    
End Sub

Private Sub cmdSend_Click()
    'gbstrSend = txtSend.Text & vbCr
    wsDCR.SendData txtSend.Text
End Sub

Private Sub tmrTimuout_Timer()
    gbintTimeout = gbintTimeout + 1
    lblCount.Caption = gbintTimeout
    If gbintTimeout > 60 Then
        tmrTimuout.Enabled = False
        Send "MEASURE /E"
        lblID.Caption = "Timeout,Rescan it"
    End If
End Sub

Private Sub wsDCR_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim S As String
    'Dim bytBuf() As Byte
    Dim i As Integer
    Dim Counter As Single
    
    'ReDim bytBuf(bytesTotal - 1)
    
    If bytesTotal < 100 Then
        wsDCR.GetData sData
        'wsDCR.GetData bytBuf, vbByte + vbArray
        
        'sData = CStr(bytBuf(0))
        txtReceive.Text = sData
        S = Trim(sData)
        If Mid(S, 1, 2) <> "OK" And Mid(S, 1, 2) <> "ER" Then
            lblID.Caption = S
            gbblnGetDCR = True
            tmrTimuout.Enabled = False
            Send "MEASURE /E"
        End If
    End If
End Sub

Public Sub Send(sSend As String)

On Error GoTo ERRLINE:
    
    wsDCR.SendData sSend
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " UDP Send error"
    ShowAlarmFlash 21

End Sub

Public Sub Scan(IsStart As Boolean)
    If IsStart Then
    
        gbblnGetDCR = False
        gbintTimeout = 0
        Send "MEASURE /C"
        lblID.Caption = "Scanning..."
        tmrTimuout.Enabled = True
    Else
        Send "MEASURE /E"
        lblID.Caption = "Stop scan"
        tmrTimuout.Enabled = False
    End If

End Sub



