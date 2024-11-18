VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUDP 
   Caption         =   "UDP"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.Timer tmrSendStatus 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6240
      Top             =   3240
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5640
      Top             =   3240
   End
   Begin MSWinsockLib.Winsock wsServer 
      Left            =   5880
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "192.168.1.168"
      RemotePort      =   168
      LocalPort       =   169
   End
   Begin VB.TextBox txtReceive 
      Height          =   3015
      Left            =   600
      ScrollBars      =   2  'Vertical
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
      Text            =   "168"
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "192.168.1.168"
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmUDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    
    wsServer.sendData txtSend.text
    
End Sub

Private Sub Form_Initialize()
    '
End Sub

Private Sub Form_Load()
    '
End Sub


Public Sub OpenUDP()
On Error GoTo ERRLINE:
    'wsServer.Close
    'wsServer.Bind (169)
    If Kernel.IsRemoteConnect = 0 Then
         wsServer.RemoteHost = Para.strRobotIP
        'wsServer.RemoteHost = "192.168.0.168"
        'wsServer.RemoteHost = "255.255.255.255"
        wsServer.LocalPort = 169
        
        If Para.intAutoPort = 0 Then
            wsServer.RemotePort = 168
        Else
            wsServer.LocalPort = Para.intAutoPort
            wsServer.RemotePort = Para.intAutoPort
'            wsServer.RemotePort = 1008
        End If
        
        wsServer.Bind
        'wsServer.Connect
        'wsServer.SendData "#SC=1,"
        tmrTimeout.Enabled = True
        Kernel.IsRemoteConnect = 1
    End If
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " UDP Open error"
    WriteLog (gbstrAlarmHint)
    ShowAlarmFlash 21
    
End Sub


Private Sub tmrSendStatus_Timer()
    Dim S As String
    Dim i As Integer
    
    If Para.UseAutoMode = 1 Then
        For i = 0 To 6
            S = S & "#TC" & CStr(i) & "=" & Format(Kernel.sngTC(i), "0.0") & ","
        Next i
        'wsServer.SendData s
        's = ""
        For i = 0 To 2
            S = S & "#MF" & CStr(i) & "=" & Format(SysAI.sngMFC(i), "0.0") & ","
        Next i
        'wsServer.SendData s

        S = S & "#VA=" & Format(Kernel.sngPressure, "0.0") & "," _
          & "#IN=" & Format(Kernel.sngIntensity, "0.0") & ","
        'wsServer.SendData s

        S = S & "#IR=" & CStr(Kernel.IsRun) & "," _
          & "#VS=" & CStr(SysDI.IsChamberGaugeL) & "," _
          & "#PW=" & CStr(SysDI.IsReady) & "," _
          & "#OH=" & CStr(SysDI.IsOverHeat) & "," _
          & "#EM=" & CStr(SysDI.IsEMO) & "," _
          & "#WA=" & CStr(SysDI.IsWater) & "," _
          & "#CD=" & CStr(SysDI.IsCDA) & "," _
          & "#PP=" & CStr(SysDO.IsPumping) & "," _
          & "#AG=" & CStr(SysDO.IsAngle) & "," _
          & "#AL=" & CStr(Kernel.IsAlarm) & "," _
          & "#TW=" & CStr(gbintTowerIndex) & "," _
          & "#CS=" & CStr(Kernel.intCurrStep) & "," _
          & "#WD=" & CStr(gblngCheckPC) & ","

    '& "#PT=" & Format(sngRecipeTotalSec, "0") & ","

        wsServer.sendData S
        txtReceive.text = S
    End If
End Sub

Private Sub tmrTimeout_Timer()
    On Error GoTo ERRLINE:
    
    wsServer.sendData "#SC=1,"
    
ERRLINE:
    tmrTimeout.Enabled = False
    gbstrAlarmHint = " UDP Timeout"
    WriteLog (gbstrAlarmHint)
    ShowAlarmFlash 21
End Sub

Private Sub wsServer_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim S As String
    Dim s1 As String
    Dim s2 As String
    Dim ss() As String
    Dim ps() As String
    Dim i As Integer
    Dim j As Integer
    Dim ix As Integer
 On Error GoTo ERRLINE:
 
    tmrTimeout.Enabled = False
    wsServer.GetData sData
    txtReceive.text = sData
'    WriteLog ("接收到RTA300數據為:" & sData)
    ss = Split(sData, ",")
    For i = 0 To UBound(ss)
        s1 = ss(i)
        If Mid(s1, 1, 1) = "$" Then
            S = Mid(s1, 2, 2)
            Select Case S
                Case "SC"
                    If Mid(s1, 5, 1) = "1" Then
                        wsServer.sendData "#SC=1,"
                    End If
                Case "PR"
                    If Para.UseAutoMode = 1 Then
                        If Mid(s1, 5, 1) = "1" Then
                            Kernel.IsRemoteStart = 1
                            Kernel.IsRemoteStop = 0
                            mdifrmRTP.tbrRTP_ButtonClick mdifrmRTP.tbrRTP.Buttons("iRun")
                            If Kernel.IsAlarm = 0 Then
                                wsServer.sendData "#PR=1,"
                            Else
                                wsServer.sendData "#PR=0,"
                                WriteLog ("11111")
                            End If
                        Else
                            Kernel.IsRemoteStop = 1
                            StopProc (2)
                            
                        End If
                    End If
                Case "LR"
                    S = "C:\Program Files\eRTA100\Recipe\" & Mid(s1, 8, Len(s1) - 7)
                    If FileExists(S) = True Then
                        gbstrPNRecipeFile = S
                        wsServer.sendData "#LR=1,"
                    Else
                        gbstrPNRecipeFile = ""
                        wsServer.sendData "#LR=0,"
                    End If
                Case "RS"
                    AlarmFlashClose
                Case "PU"
                    If Mid(s1, 5, 1) = "1" Then
                        frmPlotProcess.chkPurge.value = 1
                        wsServer.sendData "#PU=1,"
                    Else
                        frmPlotProcess.chkPurge.value = 0
                        wsServer.sendData "#PU=0,"
                    End If
                Case "PP"
                    If Mid(s1, 5, 1) = "1" Then
                        frmPlotProcess.chkPumping.value = 1
                        wsServer.sendData "#PP=1,"
                    Else
                        frmPlotProcess.chkPumping.value = 0
                        wsServer.sendData "#PP=0,"
                    End If
                Case "PH"
                    s2 = Mid(s1, 5, Len(s1) - 4)
                    If s2 <> "" Then
                        gbstrPNRecipeFile = gbstrRecipeFilePath & s2
                        gbblnPNLoad = True
                        frmRecipeEdit.cmdRecipeOpen_Click
                        gbblnPNLoad = False
                        Call mdifrmRTP.ShowTitleBar(frmLogin.LoginSucceeded)
                        StartProc
                        wsServer.sendData "#PH=1,"
                    Else
                        If Kernel.IsRun = 1 Then
                            wsServer.sendData "#PH=0,"
                        End If
                    End If
'                    ps = Split(s2, "_")
'                    If UBound(ps) > 0 Then
'                        gbintPreheatPower = CInt(ps(0))
'                        gbintPreheatTime = CInt(ps(1))
'                        Call frmDiagnosis.RunPreheat(gbintPreheatPower)
'                        'wsServer.SendData "#PH=1,"
'                    Else
'                        'wsServer.SendData "#PH=0,"
'                    End If
                    
                Case "GS"
                    S = ""
                    For j = 0 To 6
                        S = S & "#TC" & CStr(j) & "=" & Format(Kernel.sngTC(j), "0.0") & ","
                    Next j
                    
                    For j = 0 To 2
                        S = S & "#MF" & CStr(j) & "=" & Format(SysAI.sngMFC(j), "0.0") & ","
                    Next j
                    For j = 0 To 5
                        S = S & "#MN" & CStr(j) & "=" & gbstrGasAlias(j) & ","
                    Next j
                    
                    S = S & "#CY="
                    For j = 0 To 59
                        S = S & Format(Kernel.dblCT(j), "0")
                    Next j
                    S = S & ","
                                
                    S = S & "#VA=" & Format(Kernel.sngPressure, "0.000000") & "," _
                      & "#IN=" & Format(Kernel.sngIntensity, "0.00") & ","
                    SysDI.IsCoverDown = 0
                    S = S & "#IR=" & CStr(Kernel.IsRun) & "," _
                      & "#VS=" & CStr(SysDI.IsChamberGaugeL) & "," _
                      & "#PW=" & CStr(SysDI.IsReady) & "," _
                      & "#OH=" & CStr(SysDI.IsOverHeat) & "," _
                      & "#EM=" & CStr(SysDI.IsEMO) & "," _
                      & "#WA=" & CStr(SysDI.IsWater) & "," _
                      & "#CD=" & CStr(SysDI.IsCDA) & "," _
                      & "#PP=" & CStr(SysDO.IsPumping) & "," _
                      & "#AG=" & CStr(SysDO.IsAngle) & "," _
                      & "#AL=" & CStr(Kernel.IsAlarm) & "," _
                      & "#PH=" & CStr(Kernel.IsPreHeat) & "," _
                      & "#TW=" & CStr(gbintTowerIndex) & "," _
                      & "#CS=" & CStr(Kernel.intCurrStep) & "," _
                      & "#CT=" & CStr(Kernel.lngCurrStepCount) & "," _
                      & "#CV=" & CStr(SysDI.IsCoverDown) & "," _
                      & "#MD=" & CStr(Para.UseAutoMode) & ","
                      
            
                     '& "#PT=" & Format(sngRecipeTotalSec, "0") & ","
            
                    wsServer.sendData S
                    
                    If Mid(s1, 5, 1) = "1" Then
                        gbblnSendRecipe = True
                    End If
                    If gbblnSendRecipe = True Then
                        gbblnSendRecipe = False
                        
                        If Kernel.strCurrRecipe <> "" Then
                        
                            S = "#RB=1,"
                            s1 = ""
                            ix = 0
                            
                            For j = 1 To GB_MAX_STEP_PROCESS
                                s2 = Trim(gbProcessRecipeStep(j).strAction)
                                If s2 = "Stop" Or s2 = "" Then
                                    Exit For
                                Else
                                    S = S & "#RX" & CStr(ix) & "=" & Format(gbProcessRecipeStep(j).sngTime, "0")
                                    s1 = s1 & "#RY" & CStr(ix) & "=" & Format(gbProcessRecipeStep(j).sngTemperature, "0")
                                    ix = ix + 1
                                End If
                                
                            Next j
                            S = S & s1 & "#RB=0,"
                            wsServer.sendData S
                        End If
                    End If
                    
                    txtReceive.text = S
            End Select

        End If
    Next i
Exit Sub
ERRLINE:
    gbstrAlarmHint = " UDP Timeout"
    WriteLog (gbstrAlarmHint)
    ShowAlarmFlash 21
End Sub


