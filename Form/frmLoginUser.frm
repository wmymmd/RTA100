VERSION 5.00
Begin VB.Form frmLoginUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登入系統"
   ClientHeight    =   2820
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6240
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1666.149
   ScaleMode       =   0  'User
   ScaleWidth      =   5859.021
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtUserName 
      Height          =   465
      Left            =   2760
      TabIndex        =   1
      Text            =   "GTC"
      Top             =   375
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   1500
   End
   Begin VB.TextBox txtPassword 
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2040
   End
End
Attribute VB_Name = "frmLoginUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type IDEREGS
    bFeaturesReg As Byte
    bSectorCountReg As Byte
    bSectorNumberReg As Byte
    bCylLowReg As Byte
    bCylHighReg As Byte
    bDriveHeadReg As Byte
    bCommandReg As Byte
    bReserved As Byte
End Type
 
Private Type DRIVERSTATUS
    bDriveError As Byte
    bIDEStatus As Byte
    bReserved(1 To 2) As Byte
    dwReserved(1 To 2) As Long
End Type
 
Private Type SENDCMDOUTPARAMS
    cBufferSize As Long
    DStatus As DRIVERSTATUS
    bBuffer(1 To 512) As Byte
End Type
 
Private Type SENDCMDINPARAMS
    cBufferSize As Long
    irDriveRegs As IDEREGS
    bDriveNumber As Byte
    bReserved(1 To 3) As Byte
    dwReserved(1 To 4) As Long
End Type
 
'API宣告
Private Declare Function CreateFileA Lib "kernel32" _
    (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Sub RtlZeroMemory Lib "kernel32" _
    (dest As Any, ByVal numBytes As Long)
     
Private Declare Function DeviceIoControl Lib "kernel32" _
    (ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Public LoginSucceeded As Boolean
Public IsLogout As Boolean

Private Function Get_HD_SNo(DrvIdx As Byte) As String
    Dim ParaIn As SENDCMDINPARAMS
    Dim ParaOut As SENDCMDOUTPARAMS
    Dim Sno As String
    Dim h As Long
    Dim intLp As Integer

    If Len(Environ("OS")) > 0 Then
        h = CreateFileA("\\.\PhysicalDrive" & DrvIdx, -1073741824, 3, 0, 3, 0, 0)
    Else
        h = CreateFileA("\\.\Smartvsd", 0, 0, 0, 1, 0, 0)
    End If
    If h = 0 Then Exit Function
     
    RtlZeroMemory ParaIn, Len(ParaIn)
    RtlZeroMemory ParaOut, Len(ParaOut)

    With ParaIn
        .bDriveNumber = DrvIdx
        .cBufferSize = 512
        With .irDriveRegs
            .bDriveHeadReg = IIf(DrvIdx And 1, 176, 160)
            .bCommandReg = 236
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
        End With
    End With
    DeviceIoControl h, 508040, ParaIn, Len(ParaIn), ParaOut, Len(ParaOut), 0, 0
    For intLp = 21 To 40 Step 2
        If ParaOut.bBuffer(intLp + 1) = 0 Then Exit For
        Sno = Sno & Chr(ParaOut.bBuffer(intLp + 1))
        If ParaOut.bBuffer(intLp) > 0 Then Sno = Sno & Chr(ParaOut.bBuffer(intLp))
    Next
    CloseHandle h
    Get_HD_SNo = Trim(Sno)

End Function



Public Function GetMacAddress() As String
 Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim MacAddress As String
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each objItem In colItems
        If objItem.MacAddress <> "" And objItem.Description <> "" Then
            If InStr(objItem.Description, "Virtual") = 0 Then
                MacAddress = MacAddress + objItem.MacAddress + ";"
            End If
        End If
    Next
    GetMacAddress = MacAddress
End Function

Private Function GetOSVersion() As String
    Dim osInfo As Object
    Dim osItem As Object
    Dim version As String
    Set osInfo = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_OperatingSystem")
    If Not osInfo Is Nothing Then
        For Each osItem In osInfo
            version = osItem.version
            If InStr(version, "10.") > 0 Then
                GetOSVersion = "Windows 10"
            Else
                GetOSVersion = "Windows XP"
            End If
        Next osItem
    End If
End Function


Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
    If IsLogout = False Then End
End Sub

Private Sub cmdOK_Click()
    Dim StrFileName         As String
    Dim StrData    As String * 60
    Dim lngRet                As Long
    Dim iTemp As Integer
    Dim strTemp As String
    Dim strID As String
    
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    Dim lDayCheck As Long
    Dim bCheckValid As Boolean
    Dim MacAddress As String
    gbSystemPath = App.Path
    StrFileName = gbSystemPath & "\System\system.cfg"
    
    Dim osVersion As String
    Dim FullMacAddress As String
    osVersion = GetOSVersion()
    MacAddress = GetMacAddress()
    
    lngRet = GetPrivateProfileString("User", "ad", "123", StrData, 20, StrFileName)
    gbstrAdminPW = Mid(StrData, 1, 3)
    lngRet = GetPrivateProfileString("User", "eg", "456", StrData, 20, StrFileName)
    gbstrEngineerPW = Mid(StrData, 1, 3)
    lngRet = GetPrivateProfileString("User", "op", "789", StrData, 20, StrFileName)
    gbstrOperatorPW = Mid(StrData, 1, 3)
        
    bCheckValid = False
        
    If txtUserName = "GTC" And Mid(txtPassword.text, 4, 4) = "init" Then
        strTemp = Get_HD_SNo(0)
'        lngRet = WritePrivateProfileString("PARAMETER", "CTGate15", EncryptDecrypt(strTemp, 123), StrFileName)
        lngRet = WritePrivateProfileString("PARAMETER", "CTGate15", strTemp, StrFileName)
        strTemp = Year(Date)
        lngRet = WritePrivateProfileString("PARAMETER", "CTGate16", strTemp, StrFileName)
        strTemp = Month(Date)
        lngRet = WritePrivateProfileString("PARAMETER", "CTGate17", strTemp, StrFileName)
        strTemp = Day(Date)
        lngRet = WritePrivateProfileString("PARAMETER", "CTGate18", strTemp, StrFileName)
        strTemp = Mid(txtPassword.text, 8, 4)
        lngRet = WritePrivateProfileString("PARAMETER", "CTGate19", strTemp, StrFileName)
        gbintValidDays = CInt(strTemp)
        lngRet = WritePrivateProfileString("PARAMETER", "PropertyCoefficient5", "1", StrFileName)
        gbintCurrDays = 1
        bCheckValid = True
    Else
        If osVersion = "Windows XP" Then
             strTemp = Get_HD_SNo(0)
             lngRet = GetPrivateProfileString("PARAMETER", "CTGate15", "123", StrData, 25, StrFileName)
             '解密
'            strID = Mid(EncryptDecrypt(StrData, 123), 1, Len(strTemp))
            strID = Mid(StrData, 1, Len(strTemp))
            If strTemp = strID Then
            lngRet = GetPrivateProfileString("PARAMETER", "CTGate16", "2015", StrData, 20, StrFileName)
            yyyy = Mid(StrData, 1, 4)
            lngRet = GetPrivateProfileString("PARAMETER", "CTGate17", "01", StrData, 20, StrFileName)
            mm = Mid(StrData, 1, 2)
            lngRet = GetPrivateProfileString("PARAMETER", "CTGate18", "01", StrData, 20, StrFileName)
            dd = Mid(StrData, 1, 2)
            lngRet = GetPrivateProfileString("PARAMETER", "CTGate19", "0000", StrData, 20, StrFileName)
            strTemp = Mid(StrData, 1, 4)
            gbintValidDays = Val(strTemp)
            lngRet = GetPrivateProfileString("PARAMETER", "PropertyCoefficient5", "0", StrData, 20, StrFileName)
            strTemp = Mid(StrData, 1, 4)
            gbintCurrDays = Val(strTemp)
                If gbintValidDays <= 9999 Then
                    If gbintCurrDays <= gbintValidDays Then
                        bCheckValid = True
                    End If
                Else
                    bCheckValid = True
                End If
            End If
        Else
            strTemp = MacAddress
            lngRet = GetPrivateProfileString("PARAMETER", "CTGate15", "123", StrData, 55, StrFileName)
             '解密
             If lngRet > 0 Then
'                FullMacAddress = Trim(EncryptDecrypt(StrData, 123))
                 FullMacAddress = Trim(StrData)
                If Len(FullMacAddress) >= Len(strTemp) Then
                    strID = Mid(FullMacAddress, 1, Len(strTemp))
                    If strTemp = strID Then
                    lngRet = GetPrivateProfileString("PARAMETER", "CTGate16", "2015", StrData, 20, StrFileName)
                    yyyy = Mid(StrData, 1, 4)
                    lngRet = GetPrivateProfileString("PARAMETER", "CTGate17", "01", StrData, 20, StrFileName)
                    mm = Mid(StrData, 1, 2)
                    lngRet = GetPrivateProfileString("PARAMETER", "CTGate18", "01", StrData, 20, StrFileName)
                    dd = Mid(StrData, 1, 2)
                    lngRet = GetPrivateProfileString("PARAMETER", "CTGate19", "0000", StrData, 20, StrFileName)
                    strTemp = Mid(StrData, 1, 4)
                    gbintValidDays = Val(strTemp)
                    lngRet = GetPrivateProfileString("PARAMETER", "PropertyCoefficient5", "0", StrData, 20, StrFileName)
                    strTemp = Mid(StrData, 1, 4)
                    gbintCurrDays = Val(strTemp)
                        If gbintValidDays <= 9999 Then
                            If gbintCurrDays <= gbintValidDays Then
                                bCheckValid = True
                            End If
                        Else
                            bCheckValid = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If bCheckValid = False Then
        MsgBox "The system is connected fail,please contact GTC!", , "System Error"
        Exit Sub
    End If
    
    
    gbintLoginRight = 5
    frmConfiguration.tabConfiguration.TabVisible(1) = False
    frmConfiguration.tabConfiguration.TabVisible(2) = False
    frmConfiguration.tabConfiguration.TabVisible(3) = False
    frmConfiguration.tabConfiguration.TabVisible(4) = False
    frmConfiguration.tabConfiguration.TabVisible(5) = False
    frmConfiguration.tabConfiguration.TabVisible(6) = False
    
    If txtUserName = "GTC" And Mid(txtPassword.text, 1, 3) = "rtp" Then
        LoginSucceeded = True
        gbintLoginRight = 1
        
        frmConfiguration.tabConfiguration.TabVisible(1) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(2) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(3) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(4) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(5) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(6) = LoginSucceeded
        frmConfiguration.txtParaNormal(0).Enabled = LoginSucceeded
        frmRecipeEdit.tabRecipe.TabVisible(1) = LoginSucceeded
        Call frmHistory.AppendLogAlert(1, "Manual", 1100, "GTC登入", 1)
        Me.Hide
        Load mdifrmRTP
        
    ElseIf txtUserName = "ad" And txtPassword = gbstrAdminPW Then
        LoginSucceeded = True
        gbintLoginRight = 2
        Call frmHistory.AppendLogAlert(1, "Manual", 1100, "Admin登入", 1)
        Me.Hide
        Load mdifrmRTP
        

    ElseIf txtUserName = "eg" And txtPassword = gbstrEngineerPW Then
        LoginSucceeded = True
        gbintLoginRight = 3
        Call frmHistory.AppendLogAlert(1, "Manual", 1100, "Engineer登入", 1)
        Me.Hide
        Load mdifrmRTP

        
    ElseIf txtUserName = "op" And txtPassword = gbstrOperatorPW Then
        LoginSucceeded = True
        gbintLoginRight = 4
        Call frmHistory.AppendLogAlert(1, "Manual", 1100, "Operator登入", 1)
        Me.Hide
        Load mdifrmRTP

    Else
        MsgBox "Invalid password!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    
    mdifrmRTP.ShowTitleBar True
    
    lngRet = WritePrivateProfileString("User", "LastLoginUser", txtUserName.text, StrFileName)
    
    If gbintLoginRight = 1 Then
        frmPlotProcess.Show
        frmPlotProcess.ZOrder
    Else
        iTemp = gbintActivePage(gbintLoginRight - 2)
        Select Case iTemp
            Case 0
                frmPlotProcess.Show
                frmPlotProcess.ZOrder
            Case 1
                frmRecipeEdit.Show
                frmRecipeEdit.ZOrder
            Case 2
                frmDiagnosis.Show
                frmDiagnosis.ZOrder
            Case 3
                frmConfiguration.Show
                frmConfiguration.ZOrder
            Case 4
                frmPlotProcessLog.Show
                frmPlotProcessLog.ZOrder
            Case 5
                frmHistory.Show
                frmHistory.ZOrder
        End Select
    End If

End Sub

Private Sub Form_Activate()
    Dim StrFileName         As String
    Dim StrData    As String * 30
    Dim lngRet                As Long
    Dim LastLoginUser         As String
    
    LoginSucceeded = False
    
    gbSystemPath = App.Path
    StrFileName = gbSystemPath & "\System\system.cfg"
    
    lngRet = GetPrivateProfileString("User", "LastLoginUser", "GTC", StrData, 20, StrFileName)
    txtUserName.text = StrData
    
    txtPassword.text = ""
    txtPassword.SetFocus
    '120822 Josh
 
End Sub

Private Sub Form_Load()
    IsLogout = False
End Sub
