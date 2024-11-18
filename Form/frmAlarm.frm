VERSION 5.00
Begin VB.Form frmAlarm 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Alarm"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNo 
      Caption         =   "No(&N)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes(&Y)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgAlert 
      Height          =   480
      Index           =   2
      Left            =   4320
      Picture         =   "frmAlarm.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAlert 
      Height          =   480
      Index           =   1
      Left            =   4200
      Picture         =   "frmAlarm.frx":030A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAlert 
      Height          =   480
      Index           =   0
      Left            =   4080
      Picture         =   "frmAlarm.frx":074C
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "**************************"
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   2835
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgMessage 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNo_Click()
    QuestionAns = vbNo
    Unload Me
End Sub

Private Sub cmdYes_Click()
    QuestionAns = vbYes
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MessageCounter = MessageCounter - 1
End Sub
