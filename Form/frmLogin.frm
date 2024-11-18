VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登入"
   ClientHeight    =   2385
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4455
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
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   1409.137
   ScaleMode       =   0  'User
   ScaleWidth      =   4183.003
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtUserName 
      Height          =   465
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2565
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Width           =   1395
   End
   Begin VB.TextBox txtPassword 
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   825
      Width           =   2565
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "User ID:"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   945
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()

    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()

    If txtPassword = "gtcrtp" Then
        
        txtPassword.text = ""
        LoginSucceeded = True
        frmConfiguration.tabConfiguration.TabVisible(1) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(2) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(3) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(4) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(5) = LoginSucceeded
        frmConfiguration.tabConfiguration.TabVisible(6) = LoginSucceeded
        frmConfiguration.txtParaNormal(0).Enabled = LoginSucceeded
        frmRecipeEdit.tabRecipe.TabVisible(1) = LoginSucceeded
        
        Me.Hide
        ElseIf txtPassword = "SemiTop@2024" Then
        frmConfiguration.Activate.Visible = True
        ShowFile (App.Path + "\Config\Active.txt")
    Else
        Call frmConfiguration.StopWatchDog
        MsgBox "Invalid Password!", , "Login"
        frmConfiguration.tmrWatchDog.Enabled = True
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_Load()
    LoginSucceeded = False
    
End Sub
