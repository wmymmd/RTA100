VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdifrmRTP 
   BackColor       =   &H8000000C&
   Caption         =   "ProRTP"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14295
   Icon            =   "mdifrmRTP.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '���f�ʬ�
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   12240
      Top             =   7680
   End
   Begin MSComctlLib.ImageList imgToolBarList 
      Left            =   4920
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":0ECA
            Key             =   "iRun"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":1DA4
            Key             =   "iStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":2C7E
            Key             =   "iProcess"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":3B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":4A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":590C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":67E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":76C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":859A
            Key             =   "iLogin"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":9474
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":978E
            Key             =   "iBuzzerOn"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":A068
            Key             =   "iBuzzerOff"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":A942
            Key             =   "iAsk"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":AD94
            Key             =   "iWarning"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":B1E6
            Key             =   "iAlarm"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":B638
            Key             =   "iNormal"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":BA8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":E70D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":103E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":120C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":1367B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":15355
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":17787
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":19BB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":25C0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmRTP.frx":31C5D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrRTP 
      Align           =   1  'Align Top
      Height          =   1170
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   2064
      ButtonWidth     =   2302
      ButtonHeight    =   1905
      ToolTips        =   0   'False
      Appearance      =   1
      ImageList       =   "imgToolBarList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����[�u"
            Key             =   "iRun"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "iStop"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�u�����u"
            Key             =   "iProcess"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�}�ҵ{��"
            Key             =   "iOpen"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�{�ǽs��"
            Key             =   "iRecipe"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���x���A"
            Key             =   "iDiagnosis"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���x�]�m"
            Key             =   "iConfiguration"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���v���u��"
            Key             =   "iChart Log"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�ާ@�O��"
            Key             =   "iHistory"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�n�X"
            Key             =   "iLogout"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ѱ�ĵ��"
            Key             =   "iAlarmReset"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����             "
            Key             =   "iAbout"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�������}"
            Key             =   "iExit"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����s��:"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�ո�/PM TEST"
            Key             =   "Debug"
            ImageIndex      =   16
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin MSComctlLib.StatusBar stabarRTP 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   8340
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   2734
            MinWidth        =   2734
            Key             =   "iSys"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "mdifrmRTP.frx":32537
            Text            =   "Mode"
            TextSave        =   "Mode"
            Key             =   "iMode"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2893
            MinWidth        =   2893
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3883
            MinWidth        =   3883
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Picture         =   "mdifrmRTP.frx":32E11
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3316
            MinWidth        =   3316
            Text            =   "Version 11.5"
            TextSave        =   "Version 11.5"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "mdifrmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
    Dim historyData(100, 1000) As Double
    Dim historyIndex As Integer
    Dim lastCaptureTime As Date
    Dim historyTime(1000) As Double
    Dim ConfigName As String
    Dim MTCSetting As New Collection
    Public AlarmDetail As String
'    Dim thread1 As New Thread.Threadhelp
     Dim Btn15ClickNum As Integer

Private Sub RefreshCurrDays()
    Dim StrFileName         As String
    Dim StrData    As String * 30
    Dim lngRet                As Long
    Dim iTemp As Integer
    Dim strTemp As String
    Dim strID As String
    
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    Dim lDayCheck As Long
    Dim bCheckValid As Boolean
    Dim dtStart As Date
    Dim Ctgate15 As String
    
    
    StrFileName = gbSystemPath & "\System\system.cfg"
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate16", "2015", StrData, 20, StrFileName)
    yyyy = Mid(StrData, 1, 4)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate17", "01", StrData, 20, StrFileName)
    mm = Mid(StrData, 1, 2)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate18", "01", StrData, 20, StrFileName)
'    Ctgate15 = CommnonReadini("PARAMETER", "CTGate15", strFileName)
    dd = Mid(StrData, 1, 2)
    yyyy = CStr(Val(yyyy))
    mm = CStr(Val(mm))
    dd = CStr(Val(dd))
    If yyyy <> "9999" Then
     strTemp = yyyy & "/" & mm & "/" & dd
    dtStart = CDate(strTemp)
    iTemp = Date - dtStart
    If iTemp > 0 And iTemp > gbintCurrDays Then
        gbintCurrDays = iTemp
        strTemp = CStr(iTemp)
        
'        strTemp = EncryptDecrypt(Ctgate15 & "|" & strTemp, 123)
        lngRet = WritePrivateProfileString("PARAMETER", "PropertyCoefficient5", strTemp, StrFileName)
    End If
    End If
    
End Sub
Private Sub MDIForm_Activate()
'    frmPlotProcessLog.Label1.Caption = frmPlotProcessLog.CurrRatio
'    frmPlotProcessLog.CurrRatio = frmPlotProcessLog.CurrRatio + 1
'    frmPlotProcessLog.Refresh
'    ShowTitleBar (frmLoginUser.LoginSucceeded)
End Sub

Private Sub MDIForm_Load()
    Dim strFilePath As String
    Dim strDir              As String
    Dim iTemp As Integer
    
    
     
    ConfigName = gbSystemPath & "\Config\CTSetting.ini"
    historyIndex = 0
    lastCaptureTime = Now
    
    
    If App.PrevInstance Then  '�˵��e�@����
        'MsgBox "���{���w�g�b���椤�I", 48
        Unload mdifrmRTP
        Exit Sub
    End If
    If CommnonReadini("FuncSwitch", "ShowToolBar14", App.Path + Function_Path) = 1 Then
    mdifrmRTP.tbrRTP.Buttons(14).Visible = True
    Else
    mdifrmRTP.tbrRTP.Buttons(14).Visible = False
    End If
   
    gbSystemPath = App.Path
    gbSystemFile = gbSystemPath & "\System\system.cfg"
    
    frmLogin.LoginSucceeded = False
    frmLogin.Visible = False
    
    
    Load frmConfiguration
    Load frmRecipeEdit
    Load frmPlotProcess
    Load frmPlotProcessLog
    Load frmHistory
    Load frmDiagnosis
    Load frmDCR
    Load frmAz1
    Load frmAz2
    InitSys
    
    
     frmRecipeEdit.RefreshRecipeGridTitle
'    stabarRTP.Panels(9).text = "���� " & App.Major & "." & App.Minor & "." & App.Revision
     stabarRTP.Panels(9).text = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    
    gblngPreStatus = GB_STATUSBAR_ALARM
    cIni.Path = gbSystemPath & "\System\system.cfg"
     
    StopProc (2)
    ManualStop = False
    GbHoldState = False
    HideFile (App.Path + "\Config\Active.txt")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Kernel.IsRun = 1 Then
        ShowMessageOK "�t�Τ��b�[�u���q�A�Х������[�u!"
    Else
        QuestionAns = vbNo
        ShowMessageYN "�T�w�n�������}?"
        If QuestionAns = vbNo Then
           Cancel = 1
        Else
           If Kernel.IsActiveTC1 = 1 And (Para.RtaType = 3 Or Para.RtaType = 5 Or Para.RtaType = 9) Then Call mUSBIO_1.CloseDevice
           If Kernel.IsActiveTC2 = 1 Then Call mUSBIO_2.CloseDevice
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Unload frmLogin
    Unload frmRecipeEdit
    Unload frmPlotProcess
    Unload frmDiagnosis
    Unload frmConfiguration
       
    End
End Sub

Private Sub stabarRTP_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 1 Then
        frmLogin.Show
        frmLogin.ZOrder
    End If
End Sub



' �B��e�ˬd
Private Function PreStart_Check() As Boolean
Dim result As Boolean
result = True
If PreStart_MtcCheck <> "" Then
  result = False
End If
PreStart_Check = result
End Function

' ���MTC�]�w
Private Function IsExistCol(MTCNo As Integer) As Boolean
Dim i As Integer
Dim result As Boolean
result = False
If MTCSetting.Count > 0 Then
For i = 1 To MTCSetting.Count
If MTCNo = MTCSetting(i) Then
result = True
Exit For
End If
Next i
End If
IsExistCol = result
End Function


' ���MTC�]�w
Private Sub GetMTCSetting()
Dim i As Integer
For i = 0 To 5
If IsExistCol(MultiLoop.intLoopA(i)) = False And MultiLoop.intLoopA(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopA(i)
End If
If IsExistCol(MultiLoop.intLoopB(i)) = False And MultiLoop.intLoopB(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopB(i)
End If
If IsExistCol(MultiLoop.intLoopC(i)) = False And MultiLoop.intLoopC(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopC(i)
End If
If IsExistCol(MultiLoop.intLoopD(i)) = False And MultiLoop.intLoopD(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopD(i)
End If
If IsExistCol(MultiLoop.intLoopE(i)) = False And MultiLoop.intLoopE(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopE(i)
End If
If IsExistCol(MultiLoop.intLoopF(i)) = False And MultiLoop.intLoopF(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopF(i)
End If
If IsExistCol(MultiLoop.intLoopG(i)) = False And MultiLoop.intLoopG(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopG(i)
End If
If IsExistCol(MultiLoop.intLoopH(i)) = False And MultiLoop.intLoopH(i) > 0 Then
MTCSetting.Add MultiLoop.intLoopH(i)
End If
Next i
End Sub


' R���ˬd
Private Function PreStart_RCheck() As String
Dim i As Integer
Dim TCM1_Err As String
Dim TCM2_Err As String
Dim result As String
 If Az1.blnUseAzbil = True Then
  For i = 0 To 3
   If Az1.blnUseLoop(i) = True Then
    If InStr(1, gbstrNameTC(i), "TC") > 0 Then
     If Az1.sngRT(i) > 1 + gbRValRange Or Az1.sngRT(i) < 1 - gbRValRange Then
       TCM1_Err = TCM1_Err + "Loop" + CStr(i + 1) + ","
     End If
    End If
   End If
  Next i
 End If
 
  If Az2.blnUseAzbil = True Then
  For i = 0 To 3
   If Az2.blnUseLoop(i) = True Then
     If InStr(1, gbstrNameTC(i + 4), "TC") > 0 Then
      If Az2.sngRT(i) > 1 + gbRValRange Or Az2.sngRT(i) < 1 - gbRValRange Then
       TCM2_Err = TCM2_Err + "Loop" + CStr(i + 1) + ","
      End If
     End If
   End If
  Next i
 End If

 If TCM1_Err <> "" Then TCM1_Err = "TCM1:" + Left(TCM1_Err, Len(TCM1_Err) - 1)
 If TCM2_Err <> "" Then TCM2_Err = "TCM2:" + Left(TCM2_Err, Len(TCM2_Err) - 1)
 result = TCM1_Err + TCM2_Err
 PreStart_RCheck = result
End Function

' MTC�ˬd
Private Function PreStart_MtcCheck() As String
Dim Mtc1_Tips As String
Dim Mtc2_Tips As String
GetMTCSetting
If Para.UseMTC = 1 Then
Mtc1_Tips = MTC1_Check
If Mtc1_Tips <> "" Then
Mtc1_Tips = "MTC1:" + Mtc1_Tips
End If
End If
If Para.UseMTCB = 1 Then
Mtc2_Tips = MTC2_Check
If Mtc2_Tips <> "" Then
Mtc2_Tips = "MTC2:" + Mtc2_Tips
End If
End If
PreStart_MtcCheck = Mtc1_Tips + Mtc2_Tips
End Function

Private Function MTC1_Check() As String
Dim ReturnStr As String
Dim i As Integer
For i = 8 To 15
 If IsExistCol(i + 1) = True And Kernel.sngTC(i) >= 1372 Then
 ReturnStr = ReturnStr + CStr(i + 1) + ";"
 End If
Next i
If ReturnStr <> "" Then
 ReturnStr = TrimCustom(ReturnStr, ";")
End If
MTC1_Check = ReturnStr
End Function


Private Function MTC2_Check() As String
Dim ReturnStr As String
Dim i As Integer
For i = 16 To 23
 If IsExistCol(i + 1) = True And Kernel.sngTC(i) >= 1372 Then
 ReturnStr = ReturnStr + CStr(i + 1) + ";"
 End If
Next i
If ReturnStr <> "" Then
ReturnStr = TrimCustom(ReturnStr, ";")
End If
MTC2_Check = ReturnStr
End Function

Function TrimCustom(str As String, charToTrim As String) As String
    Dim startIndex As Integer
    Dim endIndex As Integer

    startIndex = 1
    While Left$(str, 1) = charToTrim
        str = Mid$(str, 2)
        If Len(str) = 0 Then Exit Function
    Wend

    endIndex = Len(str)
    While Right$(str, 1) = charToTrim
        str = Left$(str, Len(str) - 1)
        If Len(str) = 0 Then Exit Function
    Wend

    TrimCustom = str
End Function


Public Sub tbrRTP_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim iRet As Integer
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim strTemp1 As String
    Dim k As Integer
    If Btn15ClickNum > 2 Then Btn15ClickNum = 0
    Select Case Button.Index
        Case 1
        If Para.intCycleRuns = 0 Then
            Para.intCycleRuns = 1
        End If
        For k = 1 To Para.intCycleRuns
            If ManualStop = True Then
               ManualStop = False
               Exit For
            End If
            If Para.UseAutoMode = 0 Then
                Kernel.isEnd = False
                RefreshCurrDays
                If gbintValidDays < 9999 And gbintCurrDays > gbintValidDays Then
                    ShowMessageOK "�t�γs������,���p����tGTC!"
                    Exit Sub
                End If
                
                If frmRecipeEdit.blnSave = True Then
                    QuestionAns = vbNo
                    ShowMessageYN "�t��{�Ǥw�ק�,�O�_�x�s?"
                    If QuestionAns = vbYes Then
                        frmRecipeEdit.cmdRecipeSave_Click
                    End If
                    frmRecipeEdit.blnSave = False
                End If
                
'                If SysDI.IsReady = 0 Then
'                    ShowMessageOK "�Х��}�ҹq��!"
'                    Exit Sub
'                End If
                          
                If Para.intMonitorRuns > 0 And Kernel.intCurrMonitorRun >= Para.intMonitorRuns Then
                    Kernel.intCurrMonitorRun = 0
                    gbblnShowHint = True
                    frmStartProcess.lblHint = "�w�W�L�ʱ��]�w����,�аO�o��m�ʱ���!"
                    frmInputBarcode.lblHint = "�w�W�L�ʱ��]�w����,�аO�o��m�ʱ���!"
                End If
                
                If Para.UseBarcodeServer = 0 Then
                    If Para.UseCIM = 0 Then
                    
                        If frmRecipeEdit.lbRecipeName.Caption = " " Or frmRecipeEdit.lbRecipeName.Caption = "Unknown" Then
                            ShowMessageOK "�Х����J�t��{���ɮ�!"
                            Exit Sub
                        End If
                                 
                        frmPlotProcess.Show
                        frmPlotProcess.ZOrder
                        
                        If Para.intCycleRuns > 1 Then
                            QuestionAns = vbYes
                            Kernel.intCurrCycleRun = Kernel.intCurrCycleRun + 1
                            Kernel.IsRemoteStart = 0
                            Kernel.IsAlarm = 0
                        Else
                            QuestionAns = vbNo
                            ShowMessageYN "�}�l�[�u?"
                        End If
        
        '                    If frmPlotProcess.txtPN <> "" And gbintActiveModule_PNRecipe = 1 Then
        '                        If frmPNRecipe.LoadPNRecipe = False Then
        '                            Exit Sub
        '                        End If
        '                    End If
                        If QuestionAns = vbYes Then
                            If gbsngLifeLamp > 0 And gbsngUsedLamp >= gbsngLifeLamp Then
                                ShowAlarmFlash 15
                            End If
                           ' TC Wafer�ű��˴�
                           If gbintActiveAlarm_TcWafer = 1 Then
                            If Para.IsCali = 1 Then
                               AlarmDetail = PreStart_MtcCheck
                            If AlarmDetail <> "" Then
                                ShowAlarmFlash 29
                                If Para.AlarmDo(29) = 4 Then
                                       Exit Sub
                                   End If
                            End If
                           End If
                           End If
                            ' R���˴�
                           If gbintActiveAlarm_RValue = 1 Then
                              AlarmDetail = PreStart_RCheck
                              If AlarmDetail <> "" Then
                                ShowAlarmFlash 30
                                If Para.AlarmDo(30) = 4 Then
                                       Exit Sub
                                   End If
                            End If
                           End If
                          ' �O���ˬd
                            If CommnonReadini("PARAMETER", "ForcePreheat", ConfigName) = 1 Then
                                CaptureAndPrintRecentHistory
                                If CaptureAndCheckRecentHistory Then
                                    ShowAlarmFlash 23
                                    If Para.AlarmDo(23) = 4 Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                            
                            Call ShowTitleBar(frmLogin.LoginSucceeded)
                            StartProc
                        Else
                            Exit Sub
                        End If
                    Else
                        Kernel.strServerRecipe = ""
                        If Para.intCIMPort = 0 Or Para.intCIMPort = 2 Then
                            If frmRecipeEdit.lbRecipeName.Caption = " " Or frmRecipeEdit.lbRecipeName.Caption = "Unknown" Then
                                ShowMessageOK "�ʱ��Ҧ��ݥ����J�t��{��!"
                                Exit Sub
                            End If
                        End If
                                                
                        If Para.intCIMPort = 1 Then
                            'If SysDI.IsDoorClose = 1 Then
                            '    Call StartProcessOnlineRemote
                            'End If
                        ElseIf Para.intCIMPort = 2 Then
                        
                            QuestionAns = vbNo
                            frmInputDCR.Show vbModal
                            'frmInputDCR.ZOrder 0
                            If QuestionAns = vbYes Then
                                Call StartProcessOnlineLocal
                            End If
                        End If
                        
                    End If
                    
                Else
                    QuestionAns = vbNo
                    frmInputBarcode.Show vbModal
                    If QuestionAns = vbYes Then
                        If Para.strTestRunKey <> "" Then
                            If Kernel.IsNeedTestRun = 1 And Para.strLastRecipe <> Kernel.strServerRecipe & "-" & Para.strTestRunKey Then
                                ShowMessageOK "���m�ɶ��L��,�Х�����Ŷ]�{��!"
                                Exit Sub
                            End If
                            
                            If InStr(Kernel.strServerRecipe, Para.strTestRunKey) = 0 Then
                                strTemp = Kernel.strServerRecipe & ".rcp"
                                strTemp1 = Kernel.strServerRecipe & " " & Para.strTestRunKey & ".rcp"
                                If Para.strLastRecipe <> strTemp And Para.strLastRecipe <> strTemp1 Then
                                    If strTemp <> Para.strLastRecipe Then
                                        ShowMessageOK "�����{�ǫe,�Х�����Ŷ]�{��!"
                                        Exit Sub
                                    End If
                                End If
                                
                            End If
                        End If
                        
                        gbstrPNRecipeFile = gbstrRecipeFilePath & Kernel.strServerRecipe & ".rcp"
                        gbblnPNLoad = True
                        frmRecipeEdit.cmdRecipeOpen_Click
                        gbblnPNLoad = False
                        Call ShowTitleBar(frmLogin.LoginSucceeded)
                        StartProc
                    End If
                End If
            Else
                If Kernel.IsRemoteStart = 1 Then
                    Kernel.IsRemoteStart = 0
                    gbblnPNLoad = True
                    frmRecipeEdit.cmdRecipeOpen_Click
                    gbblnPNLoad = False
                    Call ShowTitleBar(frmLogin.LoginSucceeded)
                    StartProc
                End If
            End If
Next k
        Case 2
            StopProc (2)
            
        Case 3
            frmPlotProcess.Show
            frmPlotProcess.ZOrder
        Case 4
            frmRecipeEdit.cmdRecipeOpen_Click
        Case 5
            frmRecipeEdit.Show
            frmRecipeEdit.ZOrder
        Case 6
            frmDiagnosis.Show
            frmDiagnosis.ZOrder
        Case 7
            frmConfiguration.Show
            frmConfiguration.ZOrder
        Case 8
'            thread1.OpenForm ("frmPlotProcessLog")
            
            frmPlotProcessLog.Show
            frmPlotProcessLog.ZOrder
            frmPlotProcessLog.fraProcessHistory.Visible = True
            frmPlotProcessLog.fraProcessHistory.ZOrder
        Case 9
'            If Kernel.IsRun = 1 Then
'            Dim exePath As String
'            exePath = "D:\Debug\iHistory.exe"
'            Shell exePath, vbNormalFocus
'            Else
            frmHistory.Show
            frmHistory.ZOrder
'            End If

        Case 10
            frmLoginUser.IsLogout = True
            frmLoginUser.Show
            frmLoginUser.ZOrder
        Case 11
            AlarmFlashClose
        Case 12
            frmAbout.Show
            frmAbout.ZOrder
            frmLogin.LoginSucceeded = False
        Case 13
            QuestionAns = vbNo
            ShowMessageYN "�T�w�n�������}?"
            If QuestionAns = vbYes Then
                gbbnlPC_STOP = True
                frmConfiguration.tmrAIO.Enabled = False
                frmConfiguration.tmrDIO.Enabled = False
                ResetDO
                ResetAO
                SetDO gblngDO_PC_Check1, True
                SetDO gblngDO_PC_Check2, True
                  
                If (Para.RtaType = 1 Or Para.RtaType = 3) Then
                    Ixud_DriverClose
                End If
                If (Para.RtaType = 3 Or Para.RtaType = 5) And Kernel.IsActiveTC1 = 1 Then
                    mUSBIO_1.CloseDevice
                    If Para.UseMTC = 1 And Kernel.IsActiveTC2 = 1 Then
                        mUSBIO_2.CloseDevice
                        
                    End If
                End If
                
                If Para.RtaType = 9 Then
                    If Kernel.IsActiveTC1 = 1 Then mUSBIO_1.CloseDevice
                    If Kernel.IsActiveTC2 = 1 Then mUSBIO_2.CloseDevice
                    
                End If
                If (Para.RtaType = 5 Or Para.RtaType = 6 Or Para.RtaType = 8 Or Para.RtaType = 9) And Kernel.IsActiveIO = 1 Then
                    FRB_DriverClose
                End If
                
                If Para.RtaType = 7 Or Para.RtaType = 8 Then
                    frmTCP.CloseTCP
                End If
    
                If Para.RtaType = 9 Then
                    frmAz1.CloseTCP
                    frmAz2.CloseTCP
                    
                End If
                
                 If Para.RtaType = 9 And IsUsedSCR = 1 Then
                    frmModBusRtu.CloseMSComm ("SCR")
                End If
                
                On Error GoTo ERRCOM:
                If Para.UseCT = 1 Then
                    frmConfiguration.MSComm1.PortOpen = False
                End If
                
                
                End
            End If
    Case 15
        If Kernel.IsRun = 0 And Kernel.IsAlarm = 0 Then
          Btn15ClickNum = Btn15ClickNum + 1
          If Btn15ClickNum = 1 Then
             SetTower 4, True
          Else
             SetTower 4, False
          End If
        End If
    End Select
    Exit Sub
ERRCOM:
    End
End Sub

'Aries Add two function for online remote/local
Public Sub StartProcessOnlineRemote()
    frmPlotProcess.txtPN.text = CurrProc.strCaseID
    frmPlotProcess.txtID2.text = CurrProc.strWaferID(0)
    Kernel.strBarcodeID = CurrProc.strWaferID(0)
    gbstrPNRecipeFile = gbstrRecipeFilePath & Kernel.strServerRecipe
    gbblnPNLoad = True
    frmRecipeEdit.cmdRecipeOpen_Click
    gbblnPNLoad = False
    Call ShowTitleBar(frmLogin.LoginSucceeded)
    StartProc
End Sub

Public Sub StartProcessOnlineLocal()
    frmPlotProcess.txtPN.text = CurrProc.strCaseID
    frmPlotProcess.txtID2.text = CurrProc.strWaferID(0)
    Kernel.strBarcodeID = CurrProc.strWaferID(0)
    gbstrPNRecipeFile = gbstrRecipeFilePath & Kernel.strServerRecipe
    gbblnPNLoad = True
    frmRecipeEdit.cmdRecipeOpen_Click
    gbblnPNLoad = False
    Call ShowTitleBar(frmLogin.LoginSucceeded)
    StartProc
End Sub

Public Sub ShowStatus()
    'Aries Modify 3 mode
    If Para.UseAutoMode = 0 Then
        'stabarRTP.Panels(2).Text = IIf(Para.UseBarcodeServer = 0 And Para.intCIMPort = 0, "���u�Ҧ�", "�s�u�Ҧ�")
        If (Para.UseBarcodeServer = 0 And Para.intCIMPort = 0) Then stabarRTP.Panels(2).text = "���u�Ҧ�"
        If (Para.UseBarcodeServer = 0 And Para.intCIMPort = 1) Then stabarRTP.Panels(2).text = "���ݼҦ�"
        If (Para.UseBarcodeServer = 0 And Para.intCIMPort = 2) Then stabarRTP.Panels(2).text = "�ʵ��Ҧ�"
    Else
        stabarRTP.Panels(2).text = "�۰ʼҦ�"
    End If
    
    If SysDI.IsReady = 0 Then
        stabarRTP.Panels(3).text = "���Ƨ�-����"
    Else
        stabarRTP.Panels(3).text = IIf(Kernel.IsRun = 0, "�Ƨ�-����", "�Ƨ�-�B�त")
    End If
    stabarRTP.Panels(4).text = "�n�J��=" & frmLoginUser.txtUserName
    stabarRTP.Panels(5).text = IIf(SysDI.IsDoorClose = 0, "�}�����A", "�������A")
    stabarRTP.Panels(6).text = Kernel.strCurrStep
    
    stabarRTP.Panels(10).text = IIf(Para.intCycleRuns > 1, CStr(Kernel.intCurrCycleRun) & "/" & CStr(Para.intCycleRuns), "")
        
    If Para.UseCover Then
        stabarRTP.Panels(7).text = IIf(SysDI.IsCoverDown = 1, "�\��-�U��", "�\��-�W��")
    End If
    
    
    If Kernel.IsRun = 1 Or Kernel.IsAlarm > 0 Then
        tbrRTP.Buttons("iRun").Enabled = False
    Else
        If Kernel.IsPurge = 0 Then
            tbrRTP.Buttons("iRun").Enabled = True
        End If
    End If
       
    
End Sub

Public Sub ShowTitleBar(blnIsTrue As Boolean)
    tbrRTP.Buttons("iRun").Enabled = blnIsTrue
    tbrRTP.Buttons("iConfiguration").Enabled = blnIsTrue
    tbrRTP.Buttons("iDiagnosis").Enabled = blnIsTrue
    tbrRTP.Buttons("iChart Log").Enabled = blnIsTrue
    tbrRTP.Buttons("iHistory").Enabled = blnIsTrue
    tbrRTP.Buttons("iRecipe").Enabled = blnIsTrue
    tbrRTP.Buttons("iOpen").Enabled = blnIsTrue
    tbrRTP.Buttons("iLogout").Enabled = blnIsTrue
    tbrRTP.Buttons("iAbout").Enabled = blnIsTrue
    tbrRTP.Buttons("iExit").Enabled = blnIsTrue
    tbrRTP.Buttons("iAlarmReset").Enabled = blnIsTrue
    'frmRecipeEdit.cmdRecipeSave.Enabled = blnIsTrue
    Select Case gbintLoginRight
        Case 1
        tbrRTP.Buttons("iConfiguration").Enabled = True
        tbrRTP.Buttons("iDiagnosis").Enabled = True
        tbrRTP.Buttons("iRecipe").Enabled = True
'        tbrRTP.Buttons("iChart Log").Enabled = True
        tbrRTP.Buttons("iHistory").Enabled = True
        tbrRTP.Buttons("iLogout").Enabled = True
        tbrRTP.Buttons("iExit").Enabled = True
        tbrRTP.Buttons("iAlarmReset").Enabled = True
        Case 2
        tbrRTP.Buttons("iRecipe").Enabled = True
'        tbrRTP.Buttons("iChart Log").Enabled = True
        tbrRTP.Buttons("iHistory").Enabled = True
        
        Case 3
            tbrRTP.Buttons("iConfiguration").Enabled = False
            tbrRTP.Buttons("iRecipe").Enabled = True
'            tbrRTP.Buttons("iChart Log").Enabled = True
            tbrRTP.Buttons("iHistory").Enabled = True
'            tbrRTP.Buttons("iRecipe").Enabled = False
            
        Case 4
            tbrRTP.Buttons("iConfiguration").Enabled = False
            'frmRecipeEdit.cmdRecipeSave.Enabled = False
            tbrRTP.Buttons("iRecipe").Enabled = False
            tbrRTP.Buttons("iDiagnosis").Enabled = False
            tbrRTP.Buttons("iAlarmReset").Enabled = False
    End Select
    tbrRTP.Buttons("iAbout").Visible = False
End Sub

Private Sub tmrTime_Timer()
    stabarRTP.Panels(8) = Format(Time, "hh:mm:ss")
    
End Sub


Public Sub SaveToHistory(data() As Double)
    Dim i As Integer
    For i = LBound(data) To UBound(data)
        historyData(i, historyIndex) = Round(data(i), 1)
    Next i
    historyTime(historyIndex) = Now
    historyIndex = (historyIndex + 1) Mod 1000
    
End Sub

Private Function CaptureAndCheckRecentHistory()
    Dim CTNumbers1 As String
    Dim CTNumbers As Integer
    CTNumbers1 = CommnonReadini("PARAMETER", "CTNumbers", ConfigName)
    CTNumbers = CInt(CTNumbers1)
    Dim numRowsToCheck As Integer
    numRowsToCheck = 25
    
    Dim numColsToCheck As Integer
'    numColsToCheck = 50
    numColsToCheck = CTNumbers + 5
    
    Dim zeroCountFirst5(4) As Integer
    Dim zeroCountRemaining() As Integer
    ReDim zeroCountRemaining(numColsToCheck)
    
    Dim i As Integer, j As Integer
    
     For i = 0 To 4
        zeroCountFirst5(i) = 0
    Next i
    For i = 0 To CTNumbers
        zeroCountRemaining(i) = 0
    Next i
    
    For j = 0 To 4
        For i = 0 To numRowsToCheck - 1
            If historyData(j, i) <= 0 Then
                zeroCountFirst5(j) = zeroCountFirst5(j) + 1
            End If
        Next i
    Next j
    
    For j = 5 To CTNumbers + 4
        For i = 0 To numRowsToCheck - 1
            If historyData(j, i) <= 0 Then
                zeroCountRemaining(j - 5) = zeroCountRemaining(j - 5) + 1
            End If
        Next i
    Next j
    
    
    Kernel.allZerosColumns = ""
    
    For j = 0 To 4
        If zeroCountFirst5(j) = numRowsToCheck Then
            Kernel.allZerosColumns = Kernel.allZerosColumns & "Intensity" & CStr(j + 1) & ","
            CaptureAndCheckRecentHistory = True
'            Exit Function
        End If
    Next j
    
    For j = 0 To CTNumbers
        If zeroCountRemaining(j) = numRowsToCheck Then
            Kernel.allZerosColumns = Kernel.allZerosColumns & "CT��" & CStr(j + 1) & "��,"
            CaptureAndCheckRecentHistory = False
'            Exit Function
        End If
    Next j
    
    If Len(Kernel.allZerosColumns) > 0 Then
        Kernel.allZerosColumns = Left(Kernel.allZerosColumns, Len(Kernel.allZerosColumns) - 1)
'        MsgBox "�����O0" & vbCrLf & allZerosColumns
    Else
'        MsgBox "no find"
    End If
    
    CaptureAndCheckRecentHistory = True
    
End Function

Private Sub CaptureAndPrintRecentHistory()

    Dim outputText As String
'    outputText = "Timestamp,Intensity1,...,Intensity5, Data1, Data2, ..., Data60" & vbCrLf

    Dim currentTime As Date
    currentTime = Now

    Dim Count As Integer
    Dim i As Integer
    i = 0
    Count = 0

    Dim j As Integer
    j = 0

    Dim Index As Integer
    Index = (historyIndex - 1 + 1000) Mod 1000


    Do While Count < 25

        If DateDiff("s", historyTime(Index), currentTime) <= 5 Then
'            outputText = outputText & Format(historyTime(Index), "yyyy-mm-dd hh:mm:ss") & ", "
            
            Debug.Print

            For j = 0 To 3
                outputText = outputText & Format(Az1.sngMV(j), "0.00") & ","
'                historyData(j, Index) = Az1.sngMV(j)
            Next j

            Debug.Print
            outputText = outputText & Format(Az2.sngMV(0), "0.00") & ","
'            historyData(4, Index) = Az2.sngMV(0)
            Debug.Print

            For i = LBound(Kernel.dblCT) To UBound(Kernel.dblCT)
                outputText = outputText & Format(historyData(i, Index), "0.0") & ", "
'                historyData(i + 5, Index) = Kernel.dblCT(i)
            Next i
            outputText = Left(outputText, Len(outputText) - 2)
            outputText = outputText & vbCrLf
            Count = Count + 1
        End If

        Index = (Index - 1 + 1000) Mod 1000
    Loop
    
'    historyIndex = Index
    ImportTestData outputText
'    PrintToHistoryFile outputText
End Sub

Private Sub PrintToHistoryFile(outputText As String)
    On Error GoTo ErrorHandler

    Dim fileName As String
    fileName = App.Path & "\history_data.txt"

    Open fileName For Append As #1
    Print #1, outputText
    Close #1

    Exit Sub

ErrorHandler:
    MsgBox "error"
End Sub
Private Sub ImportTestData(outputText As String)
     Dim lines() As String
    lines = Split(outputText, vbCrLf)  ' ������Τ奻
    
    
    Dim values() As String
    Dim row As Integer
    Dim col As Integer
    
    Dim numRows As Integer
    Dim numCols As Integer
    
    numRows = UBound(historyData, 2) + 1
    numCols = UBound(historyData, 1) + 1
    
    
    For row = 0 To UBound(historyData, 2)
        If row <= UBound(lines) Then
        
        values = Split(lines(row), ",")
        
        For col = 0 To UBound(historyData, 1)
            If col <= UBound(values) Then
            historyData(col, row) = CDbl(values(col))
            End If
        Next col
        End If
    Next row
    
    
    historyIndex = numRows Mod 1000
End Sub

