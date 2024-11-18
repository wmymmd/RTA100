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
   WindowState     =   2  '³Ì¤j¤Æ
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
' ª©Åv¸ê°T¡G
'
' «Ø¥ß¤é´Á¡G    2008/05/05
'
' ­ì §@ ªÌ¡G    Aries Liu
'
' ¥Ø¡@¡@ªº¡G    ¸ü¥X/¸ü¤J°Ñ¼Æ³]©w­È
'
' ¶i ¤J ÂI¡G    Sub Form_Load()
'
' ¬Û¡@¡@¨Ì¡G    ¤£¾A¥Î¡C
'
' ª©ÿ  ¥»¡G    4.0.0
'
' ¤wª¾°ÝÃD¡G    ¤£¾A¥Î¡C
'
' ¨Ï¥Î¤èªk¡G    ¤£¾A¥Î¡C
'
' °Ñ¦Ò¤åÄm¡G
'
'-¤º³¡¤å¥ó¡G
'
'-MSDN¤å¥ó¡G    ¤£¾A¥Î¡C
'
'-ºô¸ô¤å¥ó¡G
'
'*ªþ    µù¡G
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



