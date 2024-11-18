VERSION 5.00
Begin VB.Form frmStartProcess 
   Caption         =   "操作提示"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   20.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStartProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7455
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdNo 
      Caption         =   "否"
      Height          =   975
      Left            =   3840
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "是"
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblHint 
      Alignment       =   2  '置中對齊
      Caption         =   "Hint"
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '置中對齊
      Caption         =   "Start the process?"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "frmStartProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdNo_Click()
    QuestionAns = vbNo
    gbblnNoModalForm = False
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    QuestionAns = vbYes
    gbblnNoModalForm = False
    Me.Hide
    
End Sub

Private Sub Form_Activate()
    gbblnNoModalForm = True
    lblHint.Visible = False
    If gbblnShowHint = True Then
        gbblnShowHint = False
        lblHint.Visible = True
    End If
End Sub

