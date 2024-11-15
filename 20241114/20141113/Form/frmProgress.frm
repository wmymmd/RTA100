VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   ClientHeight    =   1125
   ClientLeft      =   1260
   ClientTop       =   1605
   ClientWidth     =   5250
   ControlBox      =   0   'False
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
   ScaleHeight     =   1125
   ScaleWidth      =   5250
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lbAccess 
      Caption         =   "資料存取中...."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lbProgress 
      Caption         =   "100%"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  lbAccess.Caption = "加載中........"
  lbProgress.Visible = False
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100
  Timer1.Interval = 100
End Sub


Private Sub Timer1_Timer()
   If ProgressBar1.value >= ProgressBar1.Max Then
     ProgressBar1.value = ProgressBar1.Min
     Else
     ProgressBar1.value = ProgressBar1.value + 1
   End If
End Sub
