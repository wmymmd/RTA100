VERSION 5.00
Begin VB.Form frmInputDCR 
   Caption         =   "�п�J���X"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "�з���"
      Size            =   14.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9465
   StartUpPosition =   3  '���f�ʬ�
   Begin VB.Timer tmrGetRecipe 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   3600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�T�w"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   13
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtUID 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
   Begin VB.Timer tmrGetCode 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   3120
   End
   Begin VB.CommandButton cmdScanDCR 
      Caption         =   "���y"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   10
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtWID 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2040
      TabIndex        =   2
      Text            =   "255638C2T1008107266X001002"
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox txtCID 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label lbRecipePath 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Label lbServerWait 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lbRecipe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Recipe:"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "User ID:"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DCR���A:"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lbDCR 
      Alignment       =   2  'Center
      Caption         =   "���ݤ�"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Lot ID:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cassette ID:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "�п�J���X"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "���A�����A:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lbCIM 
      Alignment       =   2  'Center
      Caption         =   "���ݤ�"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
End
Attribute VB_Name = "frmInputDCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iServerWait As Integer

Private Sub Form_Activate()
    'txtUID.SetFocus
End Sub

Private Sub cmdCancel_Click()
    tmrGetCode.Enabled = False
    frmDCR.Scan False
    lbDCR.Caption = "���ݤ�"
    QuestionAns = vbNo
    Me.Hide
End Sub


'Aries Modify for online local mode
Private Sub cmdOK_Click()
    
    CurrProc.strUserID = txtUID.text
    CurrProc.strCaseID = txtCID.text
    CurrProc.strWaferID(0) = txtWID.text
    CurrProc.strLogFilePath = ""
    CurrProc.strLogFileName = ""
    If Para.intCIMPort = 2 Then 'online local
'        gbblnGetRecipe = False
'        iServerWait = 0
'        lbRecipe.Caption = ""
'        lbRecipePath.Caption = ""
'        frmCIM.Send "$SPR=9,"
'        tmrGetRecipe.Enabled = True
'        lbServerWait.Caption = 20
'        lbServerWait.Visible = True
'    Else
        Kernel.strServerRecipe = frmRecipeEdit.lbRecipeName.Caption
        frmCIM.Send "$SPR=1,"
        QuestionAns = vbYes
        Me.Hide
    End If
End Sub

Private Sub cmdScanDCR_Click()
    frmDCR.Scan True
    lbDCR.Caption = "���y��"
    tmrGetCode.Enabled = True
End Sub

Private Sub tmrGetCode_Timer()
    Dim S As String
    Dim i As Integer
    
    
    If gbblnGetDCR = True Then
        S = frmDCR.lblID.Caption
        i = InStr(S, Chr(13))
        If i >= 0 Then
            S = Mid(S, 1, i - 1)
        End If
        txtWID.text = S
        tmrGetCode.Enabled = False
        lbDCR.Caption = "���y���\"
        Call cmdOK_Click
    End If
End Sub

Private Sub tmrGetRecipe_Timer()
    Dim S As String
    
    frmCIM.Send "$GRS=1,"
    If gbblnGetRecipe = True Then
        gbblnGetRecipe = False
        tmrGetRecipe.Enabled = False
        lbServerWait.Visible = False
        S = gbstrRecipeFilePath & Kernel.strServerRecipe
        If FileExists(S) = True Then
            
            lbRecipe.Caption = Kernel.strServerRecipe
            lbCIM.Caption = "�s�u���\"
            frmCIM.Send "$SPR=1,"
            frmCIM.Send "$SPR=1,"
            frmCIM.Send "$SPR=1,"
            Call mdifrmRTP.StartProcessOnlineRemote
            Me.Hide
            QuestionAns = vbYes
        Else
            lbRecipe.Caption = "���~���t���ɮ�!"
            lbRecipePath.Caption = S
            
        End If
        
    Else
        If iServerWait < 20 Then
            iServerWait = iServerWait + 1
            lbServerWait.Caption = CStr(20 - iServerWait)
        Else
            'tmrGetRecipe.Enabled = False
            'lbCIM.Caption = "�s�u����"
        End If
    End If
End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCID.SetFocus
    End If
End Sub


Private Sub txtCID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtWID.SetFocus
        frmDCR.Scan True
        lbDCR.Caption = "���y��"
        tmrGetCode.Enabled = True
    End If
End Sub

Private Sub txtWID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub

Public Sub SetTimerPooling(IsActive As Boolean)
    frmInputDCR.tmrGetRecipe.Enabled = IsActive
End Sub
