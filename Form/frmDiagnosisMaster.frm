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
   WindowState     =   2  '�̤j��
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
' ���v��T�G
'
' �إߤ���G    2008/05/05
'
' �� �@ �̡G    Aries Liu
'
' �ء@�@���G    ���X/���J�ѼƳ]�w��
'
' �i �J �I�G    Sub Form_Load()
'
' �ۡ@�@�̡G    ���A�ΡC
'
' ���  ���G    4.0.0
'
' �w�����D�G    ���A�ΡC
'
' �ϥΤ�k�G    ���A�ΡC
'
' �ѦҤ��m�G
'
'-�������G
'
'-MSDN���G    ���A�ΡC
'
'-�������G
'
'*��    ���G
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



