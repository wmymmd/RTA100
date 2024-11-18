VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCIM 
   Caption         =   "CIM"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
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
   ScaleHeight     =   6240
   ScaleWidth      =   8760
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer tmrPolling 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7080
      Top             =   3840
   End
   Begin VB.TextBox txtReceive 
      Height          =   3015
      Left            =   360
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   4
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Text            =   "$GRS=1"
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "8888"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Timer tmrTimuout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   3840
   End
   Begin MSWinsockLib.Winsock wsCIM 
      Left            =   5760
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   8888
      LocalPort       =   7777
   End
End
Attribute VB_Name = "frmCIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub OpenUDP()
On Error GoTo ERRLINE:
    
    wsCIM.Close
    wsCIM.RemotePort = 8888
    wsCIM.RemoteHost = "127.0.0.1"
    'wsCIM.RemoteHost = "192.168.0.234"
    wsCIM.LocalPort = 7777
    
    
    'wsCIM.Connect
    wsCIM.Bind
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
    tmrPolling.Enabled = True
    Exit Sub
ERRLINE:
    'gbstrAlarmHint = " UDP Open error"
    'ShowAlarmFlash 21
    
End Sub



Public Sub Send(sSend As String)

On Error GoTo ERRLINE:
    
    wsCIM.SendData sSend & vbCr
    
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " UDP Send error"
    'ShowAlarmFlash 21

End Sub


Private Sub cmdSend_Click()
    Send txtSend.Text
    
End Sub


Private Sub tmrPolling_Timer()
    Send "$GRS=1,"
End Sub

'SysDI.IsDoorClose = 0 ==open 1==close
'Aries Add Door Close State DRS for online remote
Private Sub wsCIM_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim sOutput As String
    Dim S As String
    Dim s1 As String
    Dim ss() As String
    Dim i As Integer
    Dim no As Integer
    On Error GoTo ERRLINE
    
    wsCIM.GetData sData
    txtReceive.Text = sData
    If Mid(sData, 1, 1) = "$" Then
        S = Mid(sData, 2, 3)
        'Aries Change the determine of the command where $xxx before #xxx
        If S = "SPR" Then
            i = CInt(Mid(sData, 6, Len(sData) - 5))
            Select Case i
                Case 1 'Process Run
                    'gbstrAlarmHint = " Remote Cycle Stop"
                    If SysDI.IsDoorClose = 1 Then 'door close
                        If Kernel.strServerRecipe <> "" Then
                            gbblnGetRecipe = True
                        End If
                    Else
                        gbstrAlarmHint = " Chamber Door Alert"
                        ShowAlarmFlash 11
                    End If
                Case 2
                    gbstrAlarmHint = " Remote Cycle Stop"
                    ShowAlarmFlash 26
                    StopProc (2)
                Case 4
                    gbstrAlarmHint = " Remote Abort"
                    ShowAlarmFlash 25
            
            End Select
        ElseIf S = "RFN" Then 'Add for host select recipe
            Dim sRcp As String
            sRcp = Mid(sData, 6, Len(sData) - 7)
            
            Kernel.strServerRecipe = sRcp
            frmInputDCR.tmrGetRecipe.Enabled = True
        ElseIf S = "PPT" Then
            gbintPumpTimeout = Mid(sData, 6, Len(sData) - 5)
        ElseIf S = "CMD" Then
            Para.intCIMPort = CInt(Mid(sData, 6, Len(sData) - 5))
        ElseIf S = "MSN" Then
            i = CInt(Mid(sData, 6, Len(sData) - 5))
            SetTower 5, IIf(i = 1, True, False)
        ElseIf S = "GRS" Then
            i = Mid(sData, 6, Len(sData) - 5)
            If i = 1 Then
            
                For i = 0 To 7
                    sOutput = sOutput & "#TC" & CStr(i) & "=" & Format(Kernel.sngTC(i), "0.0") & ","
                Next i
                For i = 0 To 4
                    sOutput = sOutput & "#MF" & CStr(i) & "=" & Format(SysAI.sngMFC(i), "0.0") & ","
                Next i
                For i = 0 To 12
                    sOutput = sOutput & "#BK" & CStr(i) & "=" & Format(Kernel.sngCurrOutSCR(i), "0.0") & ","
                Next i
                
                
                s1 = IIf(Kernel.IsAlarm = 0, "0", CStr(Kernel.IsAlarm + 4000))
                sOutput = sOutput & "#VAC=" & Format(Kernel.sngPressure, "0.000") & "," _
                                  & "#INT=" & Format(Kernel.sngIntensity, "0.00") & "," _
                                  & "#PRS=" & CStr(Kernel.IsRun) & "," _
                                  & "#VAS=" & CStr(SysDI.IsChamberGaugeL) & "," _
                                  & "#PWS=" & CStr(SysDI.IsReady) & "," _
                                  & "#OHS=" & CStr(SysDI.IsOverHeat) & "," _
                                  & "#EMS=" & CStr(SysDI.IsEMO) & "," _
                                  & "#WAS=" & CStr(SysDI.IsWater) & "," _
                                  & "#CDS=" & CStr(SysDI.IsCDA) & "," _
                                  & "#DRS=" & CStr(SysDI.IsDoorClose) & "," _
                                  & "#PPS=" & CStr(SysDO.IsPumping) & "," _
                                  & "#ALC=" & s1 & "," _
                            & "#RPT=" & CStr(CurrProc.lngCurrStepTime) & "," _
                            & "#RPS=" & CStr(Kernel.intCurrStep) & "," _
                            & "#RPA=" & CStr(CurrProc.intAction) & "," _
                            & "#RCT=" & CStr(CurrProc.lngCurrentTime) & "," _
                            & "#RFN=" & Kernel.strServerRecipe & "," _
                            & "#RFP=" & gbstrRecipeFilePath & "," _
                            & "#RLN=" & CurrProc.strLogFileName & "," _
                            & "#RLP=" & CurrProc.strLogFilePath & "," _
                            & "#CMD=" & CStr(Para.intCIMPort) & "," _
                            & "#UID=" & CurrProc.strUserID & "," _
                            & "#CID=" & CurrProc.strCaseID & "," _
                            & "#WD0=" & CurrProc.strWaferID(0) & ","
                        
                     
                Send sOutput
            End If
            

        End If
    ElseIf Mid(sData, 1, 1) = "#" Then
        ss = Split(sData, ",")
        For i = 0 To UBound(ss)
            s1 = ss(i)
            If Mid(s1, 1, 1) = "#" Then
            
                S = Mid(s1, 2, 3)
                If S = "CMD" Then
                    If Len(s1) >= 6 Then
                        no = CInt(Mid(s1, 6, Len(s1) - 5))
                        'Para.intCIMPort = no
                        SetTower 4, IIf(no = 1, True, False)
                    End If
                ElseIf S = "RFN" Then
                'Aries remoark for host need to 2 step active immedicate
'                    Kernel.strServerRecipe = Mid(s1, 6, Len(s1) - 5)
'                    If Kernel.strServerRecipe <> "" Then
'                        gbblnGetRecipe = True
'                    End If
                ElseIf S = "ALC" Then
                    no = CInt(Mid(s1, 6, Len(s1) - 5))
                    If no > 0 Then
                        ShowAlarmFlash no
                    End If
                ElseIf S = "ABS" Then
                    no = CInt(Mid(s1, 6, Len(s1) - 5))
                    gbstrAlarmHint = ""
                    If no > 0 Then
                        Select Case no
                        Case 2
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "S2F22 <>0,2"
                        Case 4
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "S7F72=0,Host cancel"
                        Case 5
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "PPID Not Defined"
                        Case 6
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "Host Send cancel SF7F65"
                        Case 8
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "Host be cancelled by host"
                        Case 10
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "Lot be abort by host in process"
                        Case 13
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "Lot be cycle-stop by host in process"
                        Case 15
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "Glass ID not in cassette"
                        Case 22
                            gbstrAlarmHint = " Abnormal(" & CStr(no) & ")" & "Download recipe timeout"
                        End Select
                                                
                    End If
                End If
            End If
        Next i
    End If
    
    Exit Sub
ERRLINE:
    'gbstrAlarmHint = " UDP Open error"
    'ShowAlarmFlash 21
    OpenUDP
End Sub



