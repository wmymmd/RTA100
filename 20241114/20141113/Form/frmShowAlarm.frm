VERSION 5.00
Begin VB.Form frmShowAlarm 
   BorderStyle     =   1  '單線固定
   Caption         =   "錯誤提示"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   20.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   8145
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   3240
      Picture         =   "frmShowAlarm.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   960
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '置中對齊
      Caption         =   "Open the door?"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7935
   End
End
Attribute VB_Name = "frmShowAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gbShow        As Boolean

Private Sub cmdOK_Click()
    
    Kernel.intOpenDoorCount = 0
    
    Me.Hide
End Sub

Private Sub Form_Activate()
    gbShow = True
    Label1.Caption = "1"
End Sub

Private Sub Form_Deactivate()
    gbShow = False
    Label1.Caption = "0"
End Sub

