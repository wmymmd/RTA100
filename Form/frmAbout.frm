VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "關於我的應用程式"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   585
      Left            =   4245
      TabIndex        =   5
      Top             =   2880
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   558.355
      ScaleMode       =   0  'User
      ScaleWidth      =   1053.5
      TabIndex        =   0
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "序號"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "GR101007"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "本產品已授權給"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "Copyright(C) Giant Technology Corp."
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   1125
      Width           =   3765
   End
   Begin VB.Label lblTitle 
      Caption         =   "GTC RTA 100"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   3765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本 7.0.1.2"
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   3765
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告: 本軟體為GTC所有，請勿應用或修改相關非GTC合法授權之處，否則將不負保固與維修責任。"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   2
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 註冊機碼安全性選項...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 註冊機碼 ROOT 類型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' 以 Unicode nul 為結尾的字串
Const REG_DWORD = 4                      ' 32-位元數值

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
   '
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "關於 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 嘗試從註冊區取得系統資訊程式路徑\名稱...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' 嘗試從註冊區取得系統資訊程式路徑...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 檢查已知的 32 位元檔案版本是否存在
        If (dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 錯誤 - 找不到檔案...
        Else
            GoTo SysInfoErr
        End If
    ' 錯誤 - 找不到註冊項目...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "目前無法提供系統資訊", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, keyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' 迴圈計數器
    Dim rc As Long                                          ' 傳回代碼
    Dim hKey As Long                                        ' 開啟的註冊機碼之控制代碼
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 註冊機碼的資料型態
    Dim tmpVal As String                                    ' 註冊機碼值的暫存空間
    Dim KeyValSize As Long                                  ' 註冊機碼變數的大小
    '------------------------------------------------------------
    ' 開啟 KeyRoot {HKEY_LOCAL_MACHINE...} 之下的註冊機碼 (RegKey)
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, keyName, 0, KEY_ALL_ACCESS, hKey) ' 開啟註冊機碼
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 處理錯誤...
    
    tmpVal = String$(1024, 0)                               ' 配置變數空間
    KeyValSize = 1024                                       ' 標示變數大小
    
    '------------------------------------------------------------
    ' 擷取註冊機碼值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 取得/建立機碼值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 處理錯誤
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 會加入以 Null 為結尾的字串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' 找到 Null，從字串中取出
    Else                                                    ' WinNT 不會加入以 Null 為結尾的字串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' 找不到 Null，取出字串
    End If
    '------------------------------------------------------------
    ' 決定機碼值的轉換型態...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜尋資料型態...
    Case REG_SZ                                             ' String 註冊機碼資料型態
        KeyVal = tmpVal                                     ' 複製字串值
    Case REG_DWORD                                          ' Double Word 註冊機碼資料型態
        For i = Len(tmpVal) To 1 Step -1                    ' 轉換每一個位元
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 逐字建立值
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 將 Double Word 轉換為 String
    End Select
    
    GetKeyValue = True                                      ' 傳回成功的訊息
    rc = RegCloseKey(hKey)                                  ' 關閉註冊機碼
    Exit Function                                           ' 離開
    
GetKeyError:      ' 錯誤發生後清除...
    KeyVal = ""                                             ' 設定傳回值為空字串
    GetKeyValue = False                                     ' 傳回失敗的訊息
    rc = RegCloseKey(hKey)                                  ' 關閉註冊機碼
End Function

