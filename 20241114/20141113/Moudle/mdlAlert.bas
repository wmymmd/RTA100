Attribute VB_Name = "mdlAlert"
Option Explicit

Public GetNo        As Integer
Public MessageCounter As Long
Public QuestionAns      As Long

Public Const ALARMTYPE = "ALARM"
Public Const ERRORTYPE = "ERROR"
Public Const EventType = "EVENT"
Public Const WARNINGTYPE = "WARNING"
Public Const QUESTIONTYPE = "QUESTION"

''Alarm Message
'
''System Alarm
'Public Const Alarm1000 = "Alarm1000_DI/DO initial failure"
'Public Const Alarm1001 = "Alarm1001_Surrounding doors was opened"
'Public Const Alarm1002 = "Alarm1002_Water is leaking"
'Public Const Alarm1003 = "Alarm1003_Air is not open"
'Public Const Alarm1004 = "Alarm1004_MassFlow got problem"
'Public Const Alarm1005 = "Alarm1005_EMC alarm"
'
''Robot Alarm
'Public Const Alarm2001 = "Alarm2001_Get robot's position data error"
'Public Const Alarm2002 = "Alarm2002_Get UpperArm Position Data Different with Current position"
'Public Const Alarm2003 = "Alarm2003_Get LowerArm Position Data Different with Current position"
'Public Const Alarm2004 = "Alarm2004_Get Rotation Position Data Different with Current position"
'Public Const Alarm2005 = "Alarm2005_Get Z-Axis Position Data Different with Current position"
'Public Const Alarm2006 = "Alarm2006_Pulse number not to be negative"
'Public Const Alarm2007 = "Alarm2007_UpperArm Over 100 Pulse"
'Public Const Alarm2008 = "Alarm2008_LowerArm Over 100 Pulse"
'
'Public Const Alarm2011 = "Alarm2011_Robot command error"
'Public Const Alarm2012 = "Alarm2012_Robot emergency stop"
'Public Const Alarm2013 = "Alarm2013_Robot limit Error"
'Public Const Alarm2014 = "Alarm2014_Robot motor motion"
'Public Const Alarm2015 = "Alarm2015_Robot target motor"
'Public Const Alarm2016 = "Alarm2016_Robot stall error"
'Public Const Alarm2017 = "Alarm2017_Robot mode"
'Public Const Alarm2018 = "Alarm2018_Robot communication error"
'
''Communication Alarm
'    'rs232
'Public Const Alarm3001 = "Alarm3001_RS232 can not connect"
'Public Const Alarm3002 = "Alarm3002_RS232 disconnect"
'Public Const Alarm3003 = "Alarm3003_Waiting response timeout"
'Public Const Alarm3004 = "Alarm3004_Response message format invalid"
'    'tcp/ip
'Public Const Alarm3101 = "Alarm3101_TCP/IP can not connect"
'Public Const Alarm3102 = "Alarm3102_TCP/IP disconnect"
'Public Const Alarm3103 = "Alarm3103_Files transfer alarm"
'Public Const Alarm3104 = "Alarm3104_Request timeout"
'Public Const Alarm3105 = "Alarm3105_Get Message invalid"
'
''Chamber PC Alarm
'    'chamber1
'Public Const Alarm4001 = "Alarm4001_Chamber1 system alarm"
'Public Const Alarm4002 = "Alarm4002_Chamber1 check clock alarm"
'Public Const Alarm4003 = "Alarm4003_Chamber1 air alarm"
'Public Const Alarm4004 = "Alarm4004_Chamber1 water alarm"
'Public Const Alarm4005 = "Alarm4005_Chamber1 overheat alarm"
'Public Const Alarm4006 = "Alarm4006_Chamber1 MFC alarm"
'    'chamber2
'Public Const Alarm4101 = "Alarm4101_Chamber2 system alarm"
'Public Const Alarm4102 = "Alarm4102_Chamber2 check clock alarm"
'Public Const Alarm4103 = "Alarm4103_Chamber2 air alarm"
'Public Const Alarm4104 = "Alarm4104_Chamber2 water alarm"
'Public Const Alarm4105 = "Alarm4105_Chamber2 overheat alarm"
'Public Const Alarm4106 = "Alarm4106_Chamber2 MFC alarm"
'
''----------------------------------------------------
'
''EVENT
'    'Login/out
'Public Const Event1001 = "Event1001_Login"
'Public Const Event1002 = "Event1002_Login Successful"
'Public Const Event1003 = "Event1003_Login Failed"
'Public Const Event1004 = "Event1004_Logout"
'    'process
'Public Const Event2001 = "Event2001_Run Process"
'Public Const Event2002 = "Event2002_Pause Process"
'Public Const Event2003 = "Event2003_Stop Process"
'Public Const Event2004 = "Event2004_Process Sequence changed"
'Public Const Event2005 = "Event2005_Process Alarm"
'Public Const Event2006 = "Event2006_Process Error"
'Public Const Event2007 = "Event2007_Unknown Process"
'    'modify
'Public Const Event3001 = "Event3001_Teach position modified"
'Public Const Event3002 = "Event3002_Recipe saved"
'Public Const Event3003 = "Event3003_Recipe modified"
'Public Const Event3004 = "Event3004_DI/DO modified"
'Public Const Event3005 = "Event3005_Calibrate Pyrometer"
'Public Const Event3006 = "Event3006_Calibration Pyrometer saved"
'Public Const Event3007 = "Event3006_Configuration saved"
'
'
'
'Public Const Event3011 = "Event3011_Teach position loaded"
'
'    'Safety
'Public Const Event4001 = "Event4001_Buzzer Enable"
'Public Const Event4002 = "Event4002__Buzzer Disable"
'Public Const Event4003 = "Event4003_Alarm Invoked"
'Public Const Event4004 = "Event4004_Error Invoked"
'
''--------------------------------------------------
'
''WARNING
'    'robot
'Public Const WARNING2001 = "WARNING2001_Pulse not to be negative"
'Public Const WARNING2011 = "WARNING2002_UpperArm Over 100 Pulse"
'Public Const WARNING2012 = "WARNING2003_LowerArm Over 100 Pulse"
'
'    'Process Config
'Public Const WARNING3001 = "WARNING3001_Start Wafer No. can't  less 1 !"
'Public Const WARNING3002 = "WARNING3002_Start Wafer No. can't  large 50 !"
'Public Const WARNING3003 = "WARNING3003_End Wafer No. can't  less 1 !"
'Public Const WARNING3004 = "WARNING3004_End Wafer No. can't  large 50 !"
    
    
'Alarm Message

'System Alarm
Public Const Alarm1000 = "Alarm1000_DI/O啟始失敗"
Public Const Alarm1001 = "Alarm1001_周圍的活動門板被開啟"
Public Const Alarm1002 = "Alarm1002_發生冷卻水洩漏"
Public Const Alarm1003 = "Alarm1003_冷卻氣體未動作"
Public Const Alarm1004 = "Alarm1004_節流閥發生錯誤"
Public Const Alarm1005 = "Alarm1005_EMC已啟動"
Public Const Alarm1006 = "Alarm1006_電源端發生過電壓"

'Robot Alarm
Public Const Alarm2001 = "Alarm2001_回授之手臂位置資料有誤"
Public Const Alarm2002 = "Alarm2002_讀取上臂位置與目前位置不符"
Public Const Alarm2003 = "Alarm2003_讀取下臂位置與目前位置不符"
Public Const Alarm2004 = "Alarm2004_讀取旋轉位置與目前位置不符"
Public Const Alarm2005 = "Alarm2005_讀取上下位移位置與目前位置不符"
Public Const Alarm2006 = "Alarm2006_移動的Pulse值不能為負數"
Public Const Alarm2007 = "Alarm2007_上臂位置超過100Pulse"
Public Const Alarm2008 = "Alarm2008_下臂位置超過100Pulse"

Public Const Alarm2011 = "Alarm2011_下達手臂之命令有誤"
Public Const Alarm2012 = "Alarm2012_手臂緊急開關已啟動"
Public Const Alarm2013 = "Alarm2013_手臂超過極限位置"
Public Const Alarm2014 = "Alarm2014_手臂未停止"
Public Const Alarm2015 = "Alarm2015_手臂目前的驅動器"
Public Const Alarm2016 = "Alarm2016_手臂目前模式"
Public Const Alarm2017 = "Alarm2017_手臂脫調錯誤"
Public Const Alarm2018 = "Alarm2018_手臂通訊錯誤"

Public Const Alarm2021 = "Alarm2021_手臂真空吸取晶圓發生錯誤"
Public Const Alarm2022 = "Alarm2022_手臂上疑似有晶圓或是真空功能失效"
Public Const Alarm2023 = "Alarm2023_手臂真空感測發生錯誤"
Public Const Alarm2024 = "Alarm2024_手臂上已有晶圓"
Public Const Alarm2025 = "Alarm2025_手臂上無晶圓或是真空功能失效"


'Communication Alarm
    'rs232
Public Const Alarm3001 = "Alarm3001_串列埠與手臂未連結"
Public Const Alarm3002 = "Alarm3002_串列埠斷線"
Public Const Alarm3003 = "Alarm3003_串列埠等待回覆時間逾時"
Public Const Alarm3004 = "Alarm3004_串列埠回傳資料無效"
    'tcp/ip
Public Const Alarm3101 = "Alarm3101_網路未連結"
Public Const Alarm3102 = "Alarm3102_網路斷線"
Public Const Alarm3103 = "Alarm3103_網路傳檔失敗"
Public Const Alarm3104 = "Alarm3104_網路要求逾時"
Public Const Alarm3105 = "Alarm3105_網路回傳資料無效"

Public Const Alarm3201 = "Alarm3201_腔體一網路未連結"
Public Const Alarm3202 = "Alarm3202_腔體二網路未連結"
Public Const Alarm3203 = "Alarm3203_"
Public Const Alarm3204 = "Alarm3204_"
Public Const Alarm3205 = "Alarm3205_"

'Chamber PC Alarm
    'chamber1
Public Const Alarm4001 = "系統連結失敗,請連絡GTC"
Public Const Alarm4002 = "Alarm4002_Chamber1 自我確保系統失常"
Public Const Alarm4003 = "Alarm4003_Chamber1 冷卻氣體失常"
Public Const Alarm4004 = "Alarm4004_Chamber1 冷卻水失常"
Public Const Alarm4005 = "Alarm4005_Chamber1 過溫"
Public Const Alarm4006 = "Alarm4006_Chamber1 氣體節流閥失常"
    'chamber2
Public Const Alarm4101 = "Alarm4101_Chamber2 系統失常"
Public Const Alarm4102 = "Alarm4102_Chamber2 自我確保系統失常"
Public Const Alarm4103 = "Alarm4103_Chamber2 冷卻氣體失常"
Public Const Alarm4104 = "Alarm4104_Chamber2 冷卻水失常"
Public Const Alarm4105 = "Alarm4105_Chamber2 過溫"
Public Const Alarm4106 = "Alarm4106_Chamber2 氣體節流閥失常"

'----------------------------------------------------

'EVENT
    'Login/out
Public Const Event1001 = "Event1001_登入"
Public Const Event1002 = "Event1002_登入成功"
Public Const Event1003 = "Event1003_登入失敗"
Public Const Event1004 = "Event1004_登出"
    'process
Public Const Event2001 = "Event2001_製程開始"
Public Const Event2002 = "Event2002_製程暫停"
Public Const Event2003 = "Event2003_製程終了"
Public Const Event2004 = "Event2004_製程程序改變"
Public Const Event2005 = "Event2005_製程警報"
Public Const Event2006 = "Event2006_製程錯誤"
Public Const Event2007 = "Event2007_不可辨識之製程"
    'modify
Public Const Event3001 = "Event3001_設定工作站位置已更改"
Public Const Event3002 = "Event3002_配方已儲存"
Public Const Event3003 = "Event3003_配方已更新"
Public Const Event3004 = "Event3004_控制點已更改"
Public Const Event3005 = "Event3005_Calibrate Pyrometer"
Public Const Event3006 = "Event3006_溫度校正結果已儲存"
Public Const Event3007 = "Event3007_設定值已儲存"
Public Const Event3008 = "Event3008_配方已經配置成功"



Public Const Event3011 = "Event3011_工作站位置已載入"

    'Safety
Public Const Event4001 = "Event4001_警報聲響啟動"
Public Const Event4002 = "Event4002__警報聲響關閉"
Public Const Event4003 = "Event4003_發生警報"
Public Const Event4004 = "Event4004_發生錯誤"
Public Const Event4005 = "Event4005_發生警告"
Public Const Event4006 = "Event4006_發生事件"

'--------------------------------------------------

'WARNING
    'robot
Public Const WARNING2001 = "WARNING2001_移動的Pulse值不能為負數"
Public Const WARNING2011 = "WARNING2002_上臂位置超過100Pulse"
Public Const WARNING2012 = "WARNING2003_下臂位置超過100Pulse"
    
    'Process Config
Public Const WARNING3001 = "WARNING3001_起始晶圓不可設定小於一"
Public Const WARNING3002 = "WARNING3002_起始晶圓不可設定大於五十"
Public Const WARNING3003 = "WARNING3003_結束晶圓不可設定小於一"
Public Const WARNING3004 = "WARNING3004_結束晶圓不可設定大於五十"
Public Const WARNING3005 = "WARNING3005_起始晶圓不可填空"
Public Const WARNING3006 = "WARNING3006_結束晶圓不可填空"
Public Const WARNING3007 = "WARNING3007_必須是數字(0-9)"
Public Const WARNING3008 = "WARNING3008_起始晶圓不可大於結束晶圓"

    'Network
Public Const WARNING4001 = "WARNING4001_製程曲線圖目錄未能開啟"
Public Const WARNING4002 = "WARNING4002_"
Public Const WARNING4003 = "WARNING4003_"

    'chamber
Public Const WARNING5001 = "WARNING5001_腔體門閘一未開啟"
Public Const WARNING5002 = "WARNING5002_腔體門閘二未開啟"
Public Const WARNING5003 = "WARNING5003_腔體一內有晶圓"
Public Const WARNING5004 = "WARNING5004_腔體二內有晶圓"
Public Const WARNING5005 = "WARNING5005_無腔體可供製程"


'Error
Public Const ERROR1001 = "抓取腔體中的晶圓作業發生錯誤"
Public Const ERROR1002 = "放置腔體中的晶圓作業發生錯誤"
Public Const ERROR1003 = "工作中"

Public Const ERROR2001 = "門閘發生動作上的錯誤"
Public Const ERROR2002 = "感測器發生不可預期的錯誤"
Public Const ERROR2003 = "未開啟"
Public Const ERROR2004 = "已開啟"
Public Const ERROR2005 = "未關閉"
Public Const ERROR2006 = "已關閉"
Public Const ERROR2007 = "系統失常"
Public Const ERROR2008 = "自我確保系統失常"
Public Const ERROR2009 = "未連結"

    'process
Public Const ERROR3001 = "ERROR3001_起始配方發生錯誤"
Public Const ERROR3002 = "ERROR3002_製程執行發生錯誤"
Public Const ERROR3003 = "ERROR3003_製程通訊命令發生錯誤"
    'cooling
Public Const ERROR3011 = "ERROR3011_冷卻區晶圓取放順序比較錯誤"
Public Const ERROR3012 = "ERROR3011_冷卻區沒有任一個已經完成冷卻程序的晶圓"
Public Const ERROR3013 = "ERROR3011_冷卻區已滿載"

    'recipe
Public Const ERROR3021 = "ERROR3021_配方配置失敗"
Public Const ERROR3022 = "ERROR3022_配方通訊命令發生錯誤"
    'calibration
Public Const ERROR3031 = "ERROR3031_溫度校正通訊命令發生錯誤"
    'machine
Public Const ERROR3041 = "ERROR3041_設備環境通訊命令發生錯誤"

'Question
Public Const QUESTION1001 = "繼續執行原製程？"
Public Const QUESTION1002 = "你確定要離開本程式？"

'Status Mode
Public Const STATUS1000 = "(建議請洽原廠維修人員)"
Public Const STATUS1001 = "腔體"
Public Const STATUS1002 = "腔體門閘"
Public Const STATUS1003 = "製程"
Public Const STATUS1004 = "配方"
Public Const STATUS1005 = "設備環境"
Public Const STATUS1006 = "腔體"
Public Const STATUS1007 = "腔體"


Public EMC_AlertFlag                   As Boolean
Public WaterLeakage_AlertFlag          As Boolean
Public CoverOpen_AlertFlag             As Boolean
Public SystemAlarm_AlertFlag           As Boolean
Public OverVoltage_AlertFlag           As Boolean
'-----------------------------------------------------


Public Function AlertShow(GetAlertMessage As String, AlertType As String) As Long
    Dim objForm     As Form
    
    Set objForm = New frmAlarm
    MessageCounter = MessageCounter + 1
    SetWindowPos objForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    'SetWindowPos objForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags

    objForm.Width = 6000
    objForm.Height = 2000
   Select Case AlertType
        Case ALARMTYPE
              objForm.imgMessage = objForm.imgAlert(0)
             objForm.Icon = objForm.imgMessage
             'Call SetSystemAlarm(True)
       Case WARNINGTYPE
             objForm.imgMessage = objForm.imgAlert(1)
             objForm.Icon = objForm.imgMessage
        Case ERRORTYPE
              objForm.imgMessage = objForm.imgAlert(0)
             objForm.Icon = objForm.imgMessage
        Case EventType

        Case QUESTIONTYPE
            With objForm
                .Top = Screen.Height / 2 - .Height + MessageCounter * (Screen.Height / 128)
                .Left = Screen.Width / 2 - .Width + MessageCounter * (Screen.Width / 128)
                .Caption = AlertType
                .lblMessage.Caption = GetAlertMessage
                .imgMessage = objForm.imgAlert(2)
                .Icon = objForm.imgMessage
                .cmdYes.Visible = True
                .cmdNo.Visible = True
                .Show vbModal
                .ZOrder
            End With
            GoTo JUMPTOEXIT
    End Select
        
    With objForm
        .Top = Screen.Height / 2 - .Height + MessageCounter * (Screen.Height / 128)
        .Left = Screen.Width / 2 - .Width + MessageCounter * (Screen.Width / 128)
        .Caption = AlertType
        .lblMessage.Caption = GetAlertMessage
        On Error GoTo SHOWMODE
        .Show
    End With
SHOWMODE:
    If ERR.Number = 401 Then
        objForm.Show 1
    End If
    
JUMPTOEXIT:
    Set objForm = Nothing
    If MessageCounter > 20 Then MessageCounter = 20
    If MessageCounter < 1 Then MessageCounter = 1
    'GetNo = MsgBox(GetAlertMessage, vbOKOnly, AlertType)
    'Call AddHistoryAlarm(GetAlertMessage)
End Function
    

