VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDiagnosisMaster 
   Caption         =   "Diagnosis"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15240
   WindowState     =   2  '程て
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmDiagnosisMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================
'
' 舦戈癟
'
' ミら戳    2008/05/05
'
'       Aries Liu
'
' ヘ    更/更把计砞﹚
'
' 秈  翴    Sub Form_Load()
'
' ㄌ    ぃ続ノ
'
' �  セ    4.0.0
'
' 拜肈    ぃ続ノ
'
' ㄏノよ猭    ぃ続ノ
'
' 把σゅ膍
'
'-ず场ゅン
'
'-MSDNゅン    ぃ続ノ
'
'-呼隔ゅン
'
'*    爹
'
'======================================================================================
Option Explicit

Dim blnIsGasKeyIn As Boolean



Dim gnNumOfSubdevices As Integer
Dim tempStr As String
Dim bRun As Boolean



Private Sub Form_Activate()
    Dim i As Integer
    
    
       
    
End Sub

Private Sub Form_Load()
    If gbintRtaType = 3 Then
        
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If gbintRtaType = 3 Then
        
        MSComm1.PortOpen = False
        
        
        
    End If
End Sub



