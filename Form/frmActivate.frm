VERSION 5.00
Begin VB.Form frmActivate 
   Caption         =   "激活"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   7350
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Btn_LongP 
      Caption         =   "永久激活"
      Height          =   855
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Btn_ShortP 
      Caption         =   "短期激活"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmActivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

