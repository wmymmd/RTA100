VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   "Custom Ratio"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
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
   ScaleHeight     =   5940
   ScaleWidth      =   5190
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2640
      TabIndex        =   34
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   1080
      TabIndex        =   33
      Top             =   5160
      Width           =   1380
   End
   Begin VB.Frame Frame17 
      Caption         =   "MFC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4215
      Index           =   0
      Left            =   2640
      TabIndex        =   18
      Top             =   720
      Width           =   2415
      Begin VB.TextBox txtRatioCUM 
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   32
         Text            =   "1"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtRatioCUM 
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   31
         Text            =   "1"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtRatioCUM 
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   30
         Text            =   "1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtRatioCUM 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   29
         Text            =   "1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtRatioCUM 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   28
         Text            =   "1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtRatioCUP 
         Height          =   375
         Left            =   1080
         TabIndex        =   26
         Text            =   "1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtRatioCUM 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   25
         Text            =   "1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Vacuum:"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MFC5"
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MFC4"
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MFC3"
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MFC2"
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MFC0"
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MFC1"
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   585
      End
   End
   Begin VB.CheckBox chkUseCustom 
      Caption         =   "UseCustom"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame17 
      Caption         =   "Temperature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4215
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2415
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   7
         Left            =   960
         TabIndex        =   15
         Text            =   "1"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   6
         Left            =   960
         TabIndex        =   13
         Text            =   "1"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Text            =   "1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Text            =   "1"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Text            =   "1"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   3
         Left            =   960
         TabIndex        =   3
         Text            =   "1"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   4
         Left            =   960
         TabIndex        =   2
         Text            =   "1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtRatioCUT 
         Height          =   390
         Index           =   5
         Left            =   960
         TabIndex        =   1
         Text            =   "1"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC7"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   600
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC6"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   600
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC1"
         Height          =   270
         Index           =   170
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "TC"
         Height          =   270
         Index           =   172
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC2"
         Height          =   270
         Index           =   188
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC3"
         Height          =   270
         Index           =   189
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC4"
         Height          =   270
         Index           =   190
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "MTC5"
         Height          =   270
         Index           =   191
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    Para.IsUseCustom = chkUseCustom.value
    Para.sngRatioCUP = CStr(txtRatioCUP.Text)
    
    For i = 0 To 7
        Para.sngRatioCUT(i) = CSng(txtRatioCUT(i).Text)
    Next i
    For i = 0 To 5
        Para.sngRatioCUM(i) = CSng(txtRatioCUM(i).Text)
    Next i
    
    Call SavePara
    
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    chkUseCustom.value = Para.IsUseCustom
    txtRatioCUP.Text = CStr(Para.sngRatioCUP)
    
    For i = 0 To 7
        txtRatioCUT(i).Text = CStr(Para.sngRatioCUT(i))
    Next i
    For i = 0 To 5
        txtRatioCUM(i).Text = CStr(Para.sngRatioCUM(i))
    Next i
End Sub
