VERSION 5.00
Begin VB.Form frmInputBarcode 
   Caption         =   "輸入條碼"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   20.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputBarcode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7395
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer tmrScanFile 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox txtInput 
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7095
   End
   Begin VB.Label lblHint 
      Alignment       =   2  '置中對齊
      Caption         =   "Hint"
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label lbServerRecipe 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   840
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   5400
   End
   Begin VB.Label lbServerStatus 
      Caption         =   "等待中"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "查詢配方檔:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "伺服器狀態:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '置中對齊
      Caption         =   "請輸入生產編號條碼"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "frmInputBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iScanCount As Integer

Private Sub cmdCancel_Click()
    QuestionAns = vbNo
    gbblnNoModalForm = False
    Call frmConfiguration.StartWatchDog
    Me.Hide
End Sub

Private Sub Form_Activate()
    Call frmConfiguration.StopWatchDog
    
    gbblnNoModalForm = True
    tmrScanFile.Enabled = False
    txtInput.Text = ""
    'txtInput.SetFocus
    lbServerRecipe.Caption = ""
    Kernel.strServerRecipe = ""
    lbServerStatus.Caption = "等待中"
    Kernel.strServerPath = Para.strServerPath
    ClearAllFile
    lblHint.Visible = False
    If gbblnShowHint = True Then
        gbblnShowHint = False
        lblHint.Visible = True
    End If
End Sub

Private Sub tmrScanFile_Timer()
    Dim sTemp As String
    Dim ss() As String
    Dim S As String
    
    If iScanCount > 300 Then
        tmrScanFile.Enabled = False
        lbServerStatus.Caption = "連線失敗,請通知IT管理人員!"
        lbServerRecipe.Caption = ""
        Call frmConfiguration.StartWatchDog
        Exit Sub
    Else
        sTemp = dir(Kernel.strServerPath & "*.txt")
        Do While sTemp <> ""
            If InStr(sTemp, "Reply") > 0 And InStr(sTemp, txtInput.Text) > 0 Then
                ss = Split(sTemp, "_")
                If UBound(ss) = 3 Then
                    lbServerStatus.Caption = "連線成功"
                    tmrScanFile.Enabled = False

                    
                    If ss(1) <> "RecipeID" Then
                        
                        Kernel.strServerRecipe = ss(1)
                        Kernel.strBarcodeID = txtInput.Text
                        lbServerRecipe.Caption = Kernel.strServerRecipe
                        txtInput.Text = ""
                        sTemp = gbSystemPath & "\Recipe\op\" & Kernel.strServerRecipe & ".rcp"
                        If FileExists(sTemp) = True Then

                            ClearAllFile
                            QuestionAns = vbYes
                            gbblnNoModalForm = False
                            Me.Hide
                                
                        Else
                            lbServerRecipe.Caption = Kernel.strServerRecipe & " 本機無此配方檔案,請先建立此程序檔並放入op資料夾!"
                            
                        End If
                    Else
                        Open Kernel.strServerPath & sTemp For Input As #1
                        Line Input #1, S
                        lbServerRecipe.Caption = S
                        Close #1
                    End If
                    Call frmConfiguration.StartWatchDog
                    Exit Sub
                End If
            End If
            sTemp = dir
            DoEvents
        Loop
        
    End If
    iScanCount = iScanCount + 1
    lbServerStatus.Caption = "連線中...(" & Format((iScanCount / 10), "0") & ")"
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim sTemp As String
    Dim ss() As String
On Error GoTo ERRLINE:
        
    If KeyAscii = 13 Then
        sTemp = txtInput.Text
        ss = Split(sTemp, " ")
        If UBound(ss) = 1 Then
            If Para.strTestRunKey = ss(1) Then
                Kernel.strServerRecipe = sTemp
                Kernel.strBarcodeID = sTemp
                lbServerRecipe.Caption = Kernel.strServerRecipe
                sTemp = gbSystemPath & "\Recipe\op\" & Kernel.strServerRecipe & ".rcp"
                If FileExists(sTemp) = True Then
                    ClearAllFile
                    QuestionAns = vbYes
                    gbblnNoModalForm = False
                    frmConfiguration.StartWatchDog
                    Me.Hide
                Else
                    lbServerRecipe.Caption = Kernel.strServerRecipe & " 本機無此配方檔案,請先建立此程序檔並放入op資料夾!"
                End If
            End If
        Else
            ClearAllFile
            sTemp = Kernel.strServerPath & txtInput.Text & "_RecipeID_" & Format(Date, "YYYYMMDD") & Format(Time, "hhnnss") & "_Request.txt"
            Open sTemp For Random As #1
            Close #1
            lbServerStatus.Caption = "連線中"

            iScanCount = 0
            tmrScanFile.Enabled = True
        End If
    End If
    Exit Sub
ERRLINE:
    lbServerStatus.Caption = "連線失敗,請通知IT管理人員!"
    lbServerRecipe.Caption = ""
    tmrScanFile.Enabled = False
    Call frmConfiguration.StartWatchDog
End Sub

Public Sub ClearAllFile()
    Dim sTemp As String
    On Error GoTo ERRLINE:
    
    sTemp = dir(Kernel.strServerPath & "*.txt")
    Do While sTemp <> ""
        If InStr(sTemp, "Start") = 0 And InStr(sTemp, "End") = 0 And InStr(sTemp, "Stop") = 0 Then
            Kill Kernel.strServerPath & sTemp
        End If
        sTemp = dir
        DoEvents
    Loop
    Exit Sub
ERRLINE:
    lbServerStatus.Caption = "連線失敗,請通知IT管理人員!"
    lbServerRecipe.Caption = ""
    tmrScanFile.Enabled = False
    Call frmConfiguration.StartWatchDog
End Sub
