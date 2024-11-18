VERSION 5.00
Begin VB.Form frmCTSetting 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "SetCT"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdCT 
      Caption         =   "保存"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox ckeForcePreheat 
         Caption         =   "強制預熱"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   1  '核取
         Width           =   1095
      End
      Begin VB.TextBox txtCTNumber 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lbCT 
         Caption         =   "燈管數量:"
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   245
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCTSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
     Read
End Sub
Private Sub cmdCT_Click()
     Save
End Sub


Public Sub Read()
    Dim ConfigName As String
    Dim buffer As String * 255
    Dim CTNumbers As String
    Dim ForcePreheat As Integer
    
    
    ConfigName = gbSystemPath & "\Config\CTSetting.ini"
    If dir(ConfigName) = "" Then
        GoTo ERR_GETCONFIG
    End If
    
    CTNumbers = CommnonReadini("PARAMETER", "CTNumbers", ConfigName)
    ForcePreheat = CommnonReadini("PARAMETER", "ForcePreheat", ConfigName)
    
    txtCTNumber.text = CTNumbers
    ckeForcePreheat.value = ForcePreheat
    
    Exit Sub
ERR_GETCONFIG:
    Call AlertShow("Not found CTSetting.ini!!", ERRORTYPE)
End Sub


Private Sub Save()
    Dim ConfigName As String
    Dim result As Boolean
    
    ConfigName = gbSystemPath & "\Config\CTSetting.ini"
    If dir(ConfigName) = "" Then
        GoTo ERR_GETCONFIG
    End If
    result = WritePrivateProfileString("PARAMETER", "CTNumbers", txtCTNumber.text, ConfigName)
    result = WritePrivateProfileString("PARAMETER", "ForcePreheat", CStr(ckeForcePreheat.value), ConfigName)
'    If result <> 0 Then
'        MsgBox "Save Succeed!"
'    Else
'        MsgBox "Save Fail!"
'    End If
 Exit Sub
ERR_GETCONFIG:
    Call AlertShow("Not found CTSetting.ini!!", ERRORTYPE)
End Sub

