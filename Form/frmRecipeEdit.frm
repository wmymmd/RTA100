VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRecipeEdit 
   Caption         =   "Recipe Edit"
   ClientHeight    =   12945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
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
   ScaleHeight     =   12945
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   480
      TabIndex        =   585
      Top             =   8760
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "PID"
      TabPicture(0)   =   "frmRecipeEdit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbName(138)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbName(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbName(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbName(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbName(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtSmoothTime"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPredit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProportional2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtIntegral2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDerivational"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtIntegrnal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProportional"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "安全設定"
      TabPicture(1)   =   "frmRecipeEdit.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIntLimit"
      Tab(1).Control(1)=   "txtOvershoot"
      Tab(1).Control(2)=   "txtUndershoot"
      Tab(1).Control(3)=   "lbName(137)"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "lbName(7)"
      Tab(1).Control(6)=   "lbName(3)"
      Tab(1).Control(7)=   "lbUnits"
      Tab(1).Control(8)=   "Label3"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "開關控制"
      TabPicture(2)   =   "frmRecipeEdit.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "fraDoor"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "APC控制"
      TabPicture(3)   =   "frmRecipeEdit.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAPC_P"
      Tab(3).Control(1)=   "txtAPC_I"
      Tab(3).Control(2)=   "txtOverPressure"
      Tab(3).Control(3)=   "lbName(159)"
      Tab(3).Control(4)=   "lbName(158)"
      Tab(3).Control(5)=   "lbName(21)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "卡控"
      TabPicture(4)   =   "frmRecipeEdit.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(3)=   "txtPrepareIndex"
      Tab(4).Control(4)=   "lbName(252)"
      Tab(4).ControlCount=   5
      Begin VB.Frame Frame6 
         Caption         =   "高壓允許範圍"
         Height          =   1455
         Left            =   -69720
         TabIndex        =   637
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox txtGatePS2 
            Height          =   390
            Left            =   1050
            TabIndex        =   639
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtGatePS1 
            Height          =   390
            Left            =   1050
            TabIndex        =   638
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   259
            Left            =   690
            TabIndex        =   641
            Top             =   840
            Width           =   240
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PS >="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   257
            Left            =   315
            TabIndex        =   640
            Top             =   360
            Width           =   570
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "降溫參數"
         Height          =   1455
         Left            =   -72120
         TabIndex        =   634
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
         Begin VB.TextBox txtTempDownTimeout 
            Height          =   390
            Left            =   1080
            TabIndex        =   635
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Timeout:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   256
            Left            =   180
            TabIndex        =   636
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "O2分析儀"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   628
         Top             =   360
         Width           =   2655
         Begin VB.TextBox txtPrepareGaugeO2 
            Height          =   390
            Left            =   1050
            TabIndex        =   630
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtPrepareTimeout 
            Height          =   390
            Left            =   1050
            TabIndex        =   629
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "O2 <="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   253
            Left            =   330
            TabIndex        =   633
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   254
            Left            =   2130
            TabIndex        =   632
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Timeout:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   255
            Left            =   180
            TabIndex        =   631
            Top             =   840
            Width           =   750
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cover Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   -69840
         TabIndex        =   622
         Top             =   480
         Width           =   2295
         Begin VB.CheckBox chkEndOpenCover 
            Caption         =   "結束-自動開蓋"
            Height          =   375
            Left            =   120
            TabIndex        =   624
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkStartCloseCover 
            Caption         =   "開始-自動蓋板"
            Height          =   375
            Left            =   120
            TabIndex        =   623
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox txtPrepareIndex 
         Height          =   390
         Left            =   -65160
         TabIndex        =   620
         Text            =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAPC_P 
         Height          =   390
         Left            =   -74280
         TabIndex        =   616
         Text            =   "50"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtAPC_I 
         Height          =   390
         Left            =   -74280
         TabIndex        =   615
         Text            =   "50"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtOverPressure 
         Height          =   390
         Left            =   -74280
         TabIndex        =   614
         Text            =   "760"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Valve Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   -72360
         TabIndex        =   611
         Top             =   480
         Width           =   2295
         Begin VB.CheckBox chkAutoCloseValve2 
            Caption         =   "Auto Close Valve 2"
            Height          =   375
            Left            =   120
            TabIndex        =   613
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CheckBox chkAutoCloseValve1 
            Caption         =   "結束-自動關閥"
            Height          =   375
            Left            =   120
            TabIndex        =   612
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame fraDoor 
         Caption         =   "Door Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   608
         Top             =   480
         Width           =   2295
         Begin VB.CheckBox chkStartAutoClose 
            Caption         =   "開始-自動關門"
            Height          =   375
            Left            =   120
            TabIndex        =   610
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkEndAutoOpen 
            Caption         =   "結束-自動開門"
            Height          =   375
            Left            =   120
            TabIndex        =   609
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.TextBox txtIntLimit 
         Height          =   390
         Left            =   -72960
         TabIndex        =   605
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtProportional 
         Height          =   390
         Left            =   1560
         TabIndex        =   601
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtIntegrnal 
         Height          =   390
         Left            =   1560
         TabIndex        =   600
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDerivational 
         Height          =   390
         Left            =   1560
         TabIndex        =   599
         Text            =   "0"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtIntegral2 
         Height          =   390
         Left            =   2280
         TabIndex        =   598
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtProportional2 
         Height          =   390
         Left            =   2280
         TabIndex        =   597
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtOvershoot 
         Height          =   390
         Left            =   -74640
         TabIndex        =   592
         Text            =   "50"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtUndershoot 
         Height          =   390
         Left            =   -74640
         TabIndex        =   591
         Text            =   "50"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPredit 
         Height          =   390
         Left            =   4080
         TabIndex        =   587
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtSmoothTime 
         Height          =   390
         Left            =   4080
         TabIndex        =   586
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lbName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "升溫階段"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   252
         Left            =   -66240
         TabIndex        =   621
         Top             =   1200
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lbName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   159
         Left            =   -74640
         TabIndex        =   619
         Top             =   960
         Width           =   105
      End
      Begin VB.Label lbName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   158
         Left            =   -74640
         TabIndex        =   618
         Top             =   480
         Width           =   195
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "OP"
         Height          =   270
         Index           =   21
         Left            =   -74640
         TabIndex        =   617
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Int Limit"
         Height          =   270
         Index           =   137
         Left            =   -72960
         TabIndex        =   607
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   270
         Left            =   -72000
         TabIndex        =   606
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Proportional"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   604
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Integral"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   603
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Derivational"
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   602
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "低溫限制"
         Height          =   270
         Index           =   7
         Left            =   -74640
         TabIndex        =   596
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "高溫限制"
         Height          =   270
         Index           =   3
         Left            =   -74640
         TabIndex        =   595
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lbUnits 
         AutoSize        =   -1  'True
         Caption         =   "℃"
         Height          =   270
         Left            =   -73560
         TabIndex        =   594
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "℃"
         Height          =   270
         Left            =   -73560
         TabIndex        =   593
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "S"
         Height          =   270
         Left            =   5040
         TabIndex        =   590
         Top             =   960
         Width           =   165
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "FactorP"
         Height          =   270
         Index           =   4
         Left            =   3120
         TabIndex        =   589
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Smooth"
         Height          =   270
         Index           =   138
         Left            =   3120
         TabIndex        =   588
         Top             =   960
         Width           =   810
      End
   End
   Begin VB.ComboBox cmbRecipeAction 
      Height          =   390
      ItemData        =   "frmRecipeEdit.frx":008C
      Left            =   6960
      List            =   "frmRecipeEdit.frx":008E
      TabIndex        =   1
      Text            =   "Idle"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraVacuum 
      Caption         =   "Vac"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   6000
      TabIndex        =   28
      Top             =   10800
      Visible         =   0   'False
      Width           =   1095
      Begin VB.TextBox Text2 
         Height          =   390
         Left            =   120
         TabIndex        =   30
         Text            =   "50"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   120
         TabIndex        =   29
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "P"
         Height          =   270
         Index           =   20
         Left            =   480
         TabIndex        =   32
         Top             =   240
         Width           =   165
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "I"
         Height          =   270
         Index           =   19
         Left            =   480
         TabIndex        =   31
         Top             =   1080
         Width           =   45
      End
   End
   Begin VB.Frame frmPIN 
      Caption         =   "PIN Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1815
      Left            =   10560
      TabIndex        =   24
      Top             =   11040
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox txtPinHeight 
         Height          =   390
         Left            =   240
         TabIndex        =   25
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   270
         Index           =   18
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   390
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   270
         Index           =   17
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame fraPressure 
      Caption         =   "Pressure Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1815
      Left            =   3000
      TabIndex        =   15
      Top             =   10920
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtPressureControl 
         Height          =   390
         Left            =   1200
         TabIndex        =   16
         Text            =   "1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "(Max=1.9 / Min=0.15)"
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   2205
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Pressure"
         Height          =   270
         Index           =   80
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Torr "
         Height          =   270
         Index           =   81
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame fraObject 
      Caption         =   "Object"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1815
      Left            =   7560
      TabIndex        =   12
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton optObject 
         Caption         =   "Susceptor"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Wafer"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdRecipeOpen 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   13560
      Picture         =   "frmRecipeEdit.frx":0090
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Frame fraSafety 
      Caption         =   "安全設定"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   8880
      TabIndex        =   7
      Top             =   10800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraInputDevice 
      Caption         =   "Input Device"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1815
      Left            =   360
      TabIndex        =   4
      Top             =   10800
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton optInputDevice 
         Caption         =   "Pyrometer"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optInputDevice 
         Caption         =   "Thermal Couple"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame fraPID 
      Caption         =   "PID Cofficient"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   8640
      TabIndex        =   3
      Top             =   10800
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtRampDownPower 
         Height          =   390
         Left            =   4080
         TabIndex        =   33
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtII 
         Height          =   390
         Left            =   5760
         TabIndex        =   21
         Text            =   "0.0005"
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPP 
         Height          =   390
         Left            =   5760
         TabIndex        =   20
         Text            =   "3"
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFeedForward 
         Height          =   390
         Left            =   4080
         TabIndex        =   11
         Text            =   "0"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Ramp D"
         Height          =   270
         Index           =   16
         Left            =   3120
         TabIndex        =   35
         Top             =   1800
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   270
         Left            =   4920
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "EP"
         Height          =   270
         Index           =   11
         Left            =   5400
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Ic"
         Height          =   270
         Index           =   9
         Left            =   5400
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   1
         Visible         =   0   'False
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   1440
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "FactorF"
         Height          =   270
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Top             =   2040
         Width           =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   3000
         X2              =   3000
         Y1              =   360
         Y2              =   1680
      End
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRecipeSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   11640
      Picture         =   "frmRecipeEdit.frx":4891A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8760
      Width           =   1815
   End
   Begin VB.TextBox txtRecipeEdit 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   9360
      TabIndex        =   0
      Text            =   "0"
      Top             =   10
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdRecipePlot 
      Caption         =   "Plot"
      Height          =   615
      Left            =   12960
      TabIndex        =   8
      Top             =   10800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fraRecipe 
      Height          =   8655
      Left            =   480
      TabIndex        =   36
      Top             =   0
      Width           =   22575
      Begin VB.CommandButton CmdBtn_BuildProcessStrp 
         Caption         =   "+進程步驟編輯"
         Height          =   375
         Left            =   9600
         TabIndex        =   671
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "-刪除"
         Height          =   375
         Left            =   12600
         TabIndex        =   626
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "+插入"
         Height          =   375
         Left            =   11520
         TabIndex        =   625
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cmbRights 
         Height          =   390
         ItemData        =   "frmRecipeEdit.frx":911A4
         Left            =   16800
         List            =   "frmRecipeEdit.frx":911B4
         TabIndex        =   38
         Text            =   "ALL"
         Top             =   200
         Width           =   1215
      End
      Begin VB.CheckBox chkFinishedClear 
         Caption         =   "Finished Clear"
         Height          =   270
         Left            =   13800
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin TabDlg.SSTab tabRecipe 
         Height          =   7935
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   22335
         _ExtentX        =   39396
         _ExtentY        =   13996
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   5292
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Process Step"
         TabPicture(0)   =   "frmRecipeEdit.frx":911D8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "hfgRecipe"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Bank Intensity"
         TabPicture(1)   =   "frmRecipeEdit.frx":911F4
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraCTCheck"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame1"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Multi-Loop"
         TabPicture(2)   =   "frmRecipeEdit.frx":91210
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "tabSCR"
         Tab(2).Control(1)=   "fraLoop(5)"
         Tab(2).Control(2)=   "fraLoop(4)"
         Tab(2).Control(3)=   "fraLoop(3)"
         Tab(2).Control(4)=   "fraLoop(2)"
         Tab(2).Control(5)=   "fraLoop(1)"
         Tab(2).Control(6)=   "chkUseMultiLoop"
         Tab(2).Control(7)=   "fraLoop(0)"
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "TCM"
         TabPicture(3)   =   "frmRecipeEdit.frx":9122C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkRunAz1(0)"
         Tab(3).Control(1)=   "chkRunAz1(1)"
         Tab(3).Control(2)=   "chkRunAz1(2)"
         Tab(3).Control(3)=   "chkRunAz1(3)"
         Tab(3).Control(4)=   "txtTestStepAz1"
         Tab(3).Control(5)=   "chkTestAz1"
         Tab(3).Control(6)=   "fraAz2"
         Tab(3).Control(7)=   "cmdWriteAz1"
         Tab(3).Control(8)=   "cmdReadAz1"
         Tab(3).Control(9)=   "cmdWriteOther"
         Tab(3).Control(10)=   "fraAz1"
         Tab(3).ControlCount=   11
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgRecipe 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   672
            Top             =   480
            Width           =   22000
            _ExtentX        =   38814
            _ExtentY        =   12938
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame fraAz1 
            Caption         =   "TCM-1"
            Height          =   3495
            Left            =   -74760
            TabIndex        =   493
            Top             =   360
            Width           =   11295
            Begin VB.CheckBox chkAz1AT 
               Caption         =   "A-Tuning"
               Height          =   375
               Left            =   240
               TabIndex        =   569
               Top             =   840
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   6
               Left            =   1680
               TabIndex        =   525
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz1Offset 
                  Height          =   390
                  Index           =   0
                  Left            =   840
                  TabIndex        =   644
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz1ST 
                  Height          =   390
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   571
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz1PN 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   530
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1IN 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   529
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1DN 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   528
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.CheckBox chkUseAz1Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   0
                  Left            =   720
                  TabIndex        =   527
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz1RT 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   526
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "offset:"
                  Height          =   270
                  Index           =   260
                  Left            =   120
                  TabIndex        =   643
                  Top             =   2710
                  Width           =   630
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   244
                  Left            =   1320
                  TabIndex        =   572
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   219
                  Left            =   120
                  TabIndex        =   534
                  Top             =   720
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   218
                  Left            =   120
                  TabIndex        =   533
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   217
                  Left            =   120
                  TabIndex        =   532
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   213
                  Left            =   120
                  TabIndex        =   531
                  Top             =   2160
                  Width           =   225
               End
            End
            Begin VB.CheckBox chkUseAz1 
               Caption         =   "Enable"
               Height          =   375
               Left            =   240
               TabIndex        =   524
               Top             =   360
               Width           =   1095
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   7
               Left            =   4080
               TabIndex        =   514
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz1Offset 
                  Height          =   390
                  Index           =   1
                  Left            =   840
                  TabIndex        =   646
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz1ST 
                  Height          =   390
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   573
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz1RT 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   519
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.CheckBox chkUseAz1Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   1
                  Left            =   720
                  TabIndex        =   518
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz1DN 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   517
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1IN 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   516
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1PN 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   515
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "offset:"
                  Height          =   270
                  Index           =   261
                  Left            =   120
                  TabIndex        =   645
                  Top             =   2710
                  Width           =   630
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   245
                  Left            =   1320
                  TabIndex        =   574
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   212
                  Left            =   120
                  TabIndex        =   523
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   214
                  Left            =   120
                  TabIndex        =   522
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   215
                  Left            =   120
                  TabIndex        =   521
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   216
                  Left            =   120
                  TabIndex        =   520
                  Top             =   720
                  Width           =   225
               End
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   8
               Left            =   6480
               TabIndex        =   504
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz1Offset 
                  Height          =   390
                  Index           =   2
                  Left            =   840
                  TabIndex        =   649
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz1ST 
                  Height          =   390
                  Index           =   2
                  Left            =   1560
                  TabIndex        =   575
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz1RT 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   509
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.CheckBox chkUseAz1Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   2
                  Left            =   720
                  TabIndex        =   508
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz1DN 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   507
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1IN 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   506
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1PN 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   505
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "offset:"
                  Height          =   270
                  Index           =   262
                  Left            =   120
                  TabIndex        =   647
                  Top             =   2710
                  Width           =   630
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   246
                  Left            =   1320
                  TabIndex        =   576
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   220
                  Left            =   120
                  TabIndex        =   513
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   221
                  Left            =   120
                  TabIndex        =   512
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   222
                  Left            =   120
                  TabIndex        =   511
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   223
                  Left            =   120
                  TabIndex        =   510
                  Top             =   720
                  Width           =   225
               End
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   9
               Left            =   8880
               TabIndex        =   494
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz1Offset 
                  Height          =   390
                  Index           =   3
                  Left            =   840
                  TabIndex        =   650
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz1ST 
                  Height          =   390
                  Index           =   3
                  Left            =   1560
                  TabIndex        =   577
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz1RT 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   499
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.CheckBox chkUseAz1Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   3
                  Left            =   720
                  TabIndex        =   498
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz1DN 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   497
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1IN 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   496
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz1PN 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   495
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "offset:"
                  Height          =   270
                  Index           =   263
                  Left            =   120
                  TabIndex        =   648
                  Top             =   2710
                  Width           =   630
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   247
                  Left            =   1320
                  TabIndex        =   578
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   224
                  Left            =   120
                  TabIndex        =   503
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   225
                  Left            =   120
                  TabIndex        =   502
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   226
                  Left            =   120
                  TabIndex        =   501
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   227
                  Left            =   120
                  TabIndex        =   500
                  Top             =   720
                  Width           =   225
               End
            End
         End
         Begin VB.CommandButton cmdWriteOther 
            Caption         =   "寫入多檔"
            Height          =   495
            Left            =   -62880
            TabIndex        =   627
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton cmdReadAz1 
            Caption         =   "讀回"
            Height          =   495
            Left            =   -62880
            TabIndex        =   568
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton cmdWriteAz1 
            Caption         =   "寫入"
            Height          =   495
            Left            =   -62880
            TabIndex        =   567
            Top             =   600
            Width           =   1335
         End
         Begin VB.Frame fraAz2 
            Caption         =   "TCM-2"
            Height          =   3495
            Left            =   -74760
            TabIndex        =   535
            Top             =   3840
            Width           =   11295
            Begin VB.Frame fraLoop 
               Caption         =   "Loop4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   13
               Left            =   8880
               TabIndex        =   657
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz2Offset 
                  Height          =   390
                  Index           =   3
                  Left            =   960
                  TabIndex        =   670
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz2RT 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   663
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.CheckBox chkUseAz2Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   3
                  Left            =   720
                  TabIndex        =   662
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz2DN 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   661
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2IN 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   660
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2PN 
                  Height          =   390
                  Index           =   3
                  Left            =   480
                  TabIndex        =   659
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2ST 
                  Height          =   390
                  Index           =   3
                  Left            =   1560
                  TabIndex        =   658
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   240
                  Left            =   120
                  TabIndex        =   669
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   241
                  Left            =   120
                  TabIndex        =   668
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   242
                  Left            =   120
                  TabIndex        =   667
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   243
                  Left            =   120
                  TabIndex        =   666
                  Top             =   720
                  Width           =   225
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   251
                  Left            =   1320
                  TabIndex        =   665
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  Caption         =   "offset:"
                  Height          =   255
                  Index           =   267
                  Left            =   240
                  TabIndex        =   664
                  Top             =   2760
                  Width           =   735
               End
            End
            Begin VB.CheckBox chkAz2AT 
               Caption         =   "A-Tuning"
               Height          =   375
               Left            =   240
               TabIndex        =   570
               Top             =   840
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   12
               Left            =   6480
               TabIndex        =   557
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz2Offset 
                  Height          =   390
                  Index           =   2
                  Left            =   840
                  TabIndex        =   656
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz2ST 
                  Height          =   390
                  Index           =   2
                  Left            =   1560
                  TabIndex        =   583
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz2PN 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   562
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2IN 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   561
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2DN 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   560
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.CheckBox chkUseAz2Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   2
                  Left            =   720
                  TabIndex        =   559
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz2RT 
                  Height          =   390
                  Index           =   2
                  Left            =   480
                  TabIndex        =   558
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.Label lbName 
                  Caption         =   "offset:"
                  Height          =   255
                  Index           =   266
                  Left            =   120
                  TabIndex        =   655
                  Top             =   2760
                  Width           =   735
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   250
                  Left            =   1320
                  TabIndex        =   584
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   239
                  Left            =   120
                  TabIndex        =   566
                  Top             =   720
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   238
                  Left            =   120
                  TabIndex        =   565
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   237
                  Left            =   120
                  TabIndex        =   564
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   236
                  Left            =   120
                  TabIndex        =   563
                  Top             =   2160
                  Width           =   225
               End
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   11
               Left            =   4080
               TabIndex        =   547
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz2Offset 
                  Height          =   390
                  Index           =   1
                  Left            =   840
                  TabIndex        =   654
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz2ST 
                  Height          =   390
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   581
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz2PN 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   552
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2IN 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   551
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2DN 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   550
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.CheckBox chkUseAz2Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   1
                  Left            =   720
                  TabIndex        =   549
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz2RT 
                  Height          =   390
                  Index           =   1
                  Left            =   480
                  TabIndex        =   548
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.Label lbName 
                  Caption         =   "offset:"
                  Height          =   255
                  Index           =   265
                  Left            =   120
                  TabIndex        =   653
                  Top             =   2710
                  Width           =   615
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   249
                  Left            =   1320
                  TabIndex        =   582
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   235
                  Left            =   120
                  TabIndex        =   556
                  Top             =   720
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   234
                  Left            =   120
                  TabIndex        =   555
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   233
                  Left            =   120
                  TabIndex        =   554
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   232
                  Left            =   120
                  TabIndex        =   553
                  Top             =   2160
                  Width           =   225
               End
            End
            Begin VB.CheckBox chkUseAz2 
               Caption         =   "Enable"
               Height          =   375
               Left            =   240
               TabIndex        =   546
               Top             =   360
               Width           =   1095
            End
            Begin VB.Frame fraLoop 
               Caption         =   "Loop1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   3135
               Index           =   10
               Left            =   1680
               TabIndex        =   536
               Top             =   240
               Width           =   2295
               Begin VB.TextBox txtAz2Offset 
                  Height          =   390
                  Index           =   0
                  Left            =   840
                  TabIndex        =   651
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtAz2ST 
                  Height          =   390
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   579
                  Text            =   "1"
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtAz2RT 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   541
                  Text            =   "1"
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.CheckBox chkUseAz2Loop 
                  Caption         =   "Enable"
                  Height          =   375
                  Index           =   0
                  Left            =   720
                  TabIndex        =   540
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtAz2DN 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   539
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2IN 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   538
                  Text            =   "0"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtAz2PN 
                  Height          =   390
                  Index           =   0
                  Left            =   480
                  TabIndex        =   537
                  Text            =   "0"
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "offset:"
                  Height          =   270
                  Index           =   264
                  Left            =   120
                  TabIndex        =   652
                  Top             =   2710
                  Width           =   630
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "S:"
                  Height          =   270
                  Index           =   248
                  Left            =   1320
                  TabIndex        =   580
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "R:"
                  Height          =   270
                  Index           =   231
                  Left            =   120
                  TabIndex        =   545
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "D:"
                  Height          =   270
                  Index           =   230
                  Left            =   120
                  TabIndex        =   544
                  Top             =   1680
                  Width           =   240
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "I:"
                  Height          =   270
                  Index           =   229
                  Left            =   120
                  TabIndex        =   543
                  Top             =   1200
                  Width           =   105
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "P:"
                  Height          =   270
                  Index           =   228
                  Left            =   120
                  TabIndex        =   542
                  Top             =   720
                  Width           =   225
               End
            End
         End
         Begin VB.CheckBox chkTestAz1 
            Caption         =   "測試程序"
            Height          =   615
            Left            =   -62400
            Style           =   1  'Graphical
            TabIndex        =   492
            Top             =   5520
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtTestStepAz1 
            Height          =   390
            Left            =   -62160
            TabIndex        =   491
            Text            =   "1"
            Top             =   5160
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkRunAz1 
            Caption         =   "測試"
            Height          =   615
            Index           =   3
            Left            =   -61920
            Style           =   1  'Graphical
            TabIndex        =   490
            Top             =   5160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkRunAz1 
            Caption         =   "測試"
            Height          =   615
            Index           =   2
            Left            =   -63120
            Style           =   1  'Graphical
            TabIndex        =   489
            Top             =   5400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkRunAz1 
            Caption         =   "測試"
            Height          =   615
            Index           =   1
            Left            =   -62040
            Style           =   1  'Graphical
            TabIndex        =   488
            Top             =   4680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkRunAz1 
            Caption         =   "測試"
            Height          =   615
            Index           =   0
            Left            =   -63360
            Style           =   1  'Graphical
            TabIndex        =   487
            Top             =   4680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame1 
            Caption         =   "Bank Weight"
            Height          =   2415
            Left            =   240
            TabIndex        =   429
            Top             =   1200
            Width           =   14415
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   463
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   1
               Left            =   2880
               TabIndex        =   462
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   2
               Left            =   3600
               TabIndex        =   461
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   3
               Left            =   4320
               TabIndex        =   460
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   4
               Left            =   5040
               TabIndex        =   459
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   4
               Left            =   5040
               TabIndex        =   458
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   3
               Left            =   4320
               TabIndex        =   457
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   2
               Left            =   3600
               TabIndex        =   456
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   1
               Left            =   2880
               TabIndex        =   455
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   454
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   5
               Left            =   5760
               TabIndex        =   453
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   5
               Left            =   5760
               TabIndex        =   452
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   6
               Left            =   6480
               TabIndex        =   451
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   6
               Left            =   6480
               TabIndex        =   450
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   7
               Left            =   7200
               TabIndex        =   449
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   7
               Left            =   7200
               TabIndex        =   448
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   8
               Left            =   7920
               TabIndex        =   447
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   8
               Left            =   7920
               TabIndex        =   446
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   9
               Left            =   8640
               TabIndex        =   445
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   9
               Left            =   8640
               TabIndex        =   444
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   10
               Left            =   9360
               TabIndex        =   443
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   10
               Left            =   9360
               TabIndex        =   442
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   11
               Left            =   10080
               TabIndex        =   441
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   11
               Left            =   10080
               TabIndex        =   440
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   12
               Left            =   10800
               TabIndex        =   439
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   12
               Left            =   10800
               TabIndex        =   438
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   13
               Left            =   11520
               TabIndex        =   437
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   13
               Left            =   11520
               TabIndex        =   436
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   14
               Left            =   12240
               TabIndex        =   435
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   14
               Left            =   12240
               TabIndex        =   434
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   15
               Left            =   12960
               TabIndex        =   433
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   15
               Left            =   12960
               TabIndex        =   432
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeightS 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   16
               Left            =   13680
               TabIndex        =   431
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtIntensityWeight 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   16
               Left            =   13680
               TabIndex        =   430
               Text            =   "100"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Weight in Dynamic (%)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   64
               Left            =   120
               TabIndex        =   482
               Top             =   1200
               Visible         =   0   'False
               Width           =   2010
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   65
               Left            =   2160
               TabIndex        =   481
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   66
               Left            =   2880
               TabIndex        =   480
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   67
               Left            =   3600
               TabIndex        =   479
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   68
               Left            =   4320
               TabIndex        =   478
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank5"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   69
               Left            =   5040
               TabIndex        =   477
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Weight in Steady (%)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   85
               Left            =   120
               TabIndex        =   476
               Top             =   720
               Width           =   1860
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank6"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   10
               Left            =   5760
               TabIndex        =   475
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank7"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   12
               Left            =   6480
               TabIndex        =   474
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank8"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   13
               Left            =   7200
               TabIndex        =   473
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank9"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   14
               Left            =   7920
               TabIndex        =   472
               Top             =   480
               Width           =   555
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank10"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   15
               Left            =   8640
               TabIndex        =   471
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank11"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   22
               Left            =   9360
               TabIndex        =   470
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank12"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   23
               Left            =   10080
               TabIndex        =   469
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank13"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   117
               Left            =   10800
               TabIndex        =   468
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank14"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   118
               Left            =   11520
               TabIndex        =   467
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank15"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   209
               Left            =   12240
               TabIndex        =   466
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank16"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   210
               Left            =   12960
               TabIndex        =   465
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Bank17"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   211
               Left            =   13680
               TabIndex        =   464
               Top             =   480
               Width           =   660
            End
         End
         Begin VB.Frame fraCTCheck 
            Caption         =   "CT Check Threshold (A)"
            Height          =   1695
            Left            =   240
            TabIndex        =   390
            Top             =   3840
            Width           =   14415
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   0
               Left            =   2160
               TabIndex        =   426
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.CheckBox chkUseCT 
               Caption         =   "Enable"
               Height          =   375
               Left            =   120
               TabIndex        =   425
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   1
               Left            =   2880
               TabIndex        =   424
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   2
               Left            =   3600
               TabIndex        =   423
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   3
               Left            =   4320
               TabIndex        =   422
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   4
               Left            =   5040
               TabIndex        =   421
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   5
               Left            =   5760
               TabIndex        =   420
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   6
               Left            =   6480
               TabIndex        =   419
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   7
               Left            =   7200
               TabIndex        =   418
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   8
               Left            =   7920
               TabIndex        =   417
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   9
               Left            =   8640
               TabIndex        =   416
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   10
               Left            =   9360
               TabIndex        =   415
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   11
               Left            =   10080
               TabIndex        =   414
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   0
               Left            =   2160
               TabIndex        =   413
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   1
               Left            =   2880
               TabIndex        =   412
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   2
               Left            =   3600
               TabIndex        =   411
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   3
               Left            =   4320
               TabIndex        =   410
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   4
               Left            =   5040
               TabIndex        =   409
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   5
               Left            =   5760
               TabIndex        =   408
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   6
               Left            =   6480
               TabIndex        =   407
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   7
               Left            =   7200
               TabIndex        =   406
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   8
               Left            =   7920
               TabIndex        =   405
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   9
               Left            =   8640
               TabIndex        =   404
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   10
               Left            =   9360
               TabIndex        =   403
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   11
               Left            =   10080
               TabIndex        =   402
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.CheckBox chkSaveLogCT 
               Caption         =   "Save Log"
               Height          =   375
               Left            =   120
               TabIndex        =   401
               Top             =   960
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   12
               Left            =   10800
               TabIndex        =   400
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   12
               Left            =   10800
               TabIndex        =   399
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   13
               Left            =   11520
               TabIndex        =   398
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   13
               Left            =   11520
               TabIndex        =   397
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   14
               Left            =   12240
               TabIndex        =   396
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   14
               Left            =   12240
               TabIndex        =   395
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   15
               Left            =   12960
               TabIndex        =   394
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   15
               Left            =   12960
               TabIndex        =   393
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCD 
               Height          =   495
               Index           =   16
               Left            =   13680
               TabIndex        =   392
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtCT 
               Height          =   495
               Index           =   16
               Left            =   13680
               TabIndex        =   391
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "(+)"
               Height          =   255
               Left            =   1800
               TabIndex        =   428
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "(-)"
               Height          =   255
               Left            =   1800
               TabIndex        =   427
               Top             =   1080
               Width           =   975
            End
         End
         Begin VB.Frame fraLoop 
            Caption         =   "Loop1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3135
            Index           =   0
            Left            =   -74160
            TabIndex        =   372
            Top             =   660
            Width           =   2295
            Begin VB.TextBox txtLoopPN 
               Height          =   390
               Index           =   0
               Left            =   480
               TabIndex        =   381
               Text            =   "0"
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtLoopIN 
               Height          =   390
               Index           =   0
               Left            =   480
               TabIndex        =   380
               Text            =   "0"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtLoopDN 
               Height          =   390
               Index           =   0
               Left            =   480
               TabIndex        =   379
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.CheckBox chkUseLoop 
               Caption         =   "Enable"
               Height          =   375
               Index           =   0
               Left            =   720
               TabIndex        =   378
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtLoopTC 
               Height          =   390
               Index           =   0
               Left            =   480
               TabIndex        =   377
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopCN 
               Height          =   390
               Index           =   0
               Left            =   840
               TabIndex        =   376
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopCV 
               Height          =   390
               Index           =   0
               Left            =   1560
               TabIndex        =   375
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopRT 
               Height          =   390
               Index           =   0
               Left            =   1440
               TabIndex        =   374
               Text            =   "1"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopFT 
               Height          =   390
               Index           =   0
               Left            =   1440
               TabIndex        =   373
               Text            =   "-1"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "P:"
               Height          =   270
               Index           =   188
               Left            =   120
               TabIndex        =   389
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "I:"
               Height          =   270
               Index           =   26
               Left            =   120
               TabIndex        =   388
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   27
               Left            =   120
               TabIndex        =   387
               Top             =   2160
               Width           =   240
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No:"
               Height          =   270
               Index           =   70
               Left            =   30
               TabIndex        =   386
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Alarm:"
               Height          =   270
               Index           =   119
               Left            =   120
               TabIndex        =   385
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "±"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   120
               Left            =   1440
               TabIndex        =   384
               Top             =   2640
               Width           =   105
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "R:"
               Height          =   270
               Index           =   131
               Left            =   1200
               TabIndex        =   383
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   139
               Left            =   1215
               TabIndex        =   382
               Top             =   2160
               Width           =   210
            End
         End
         Begin VB.CheckBox chkUseMultiLoop 
            Caption         =   "Enable"
            Height          =   375
            Left            =   -74880
            TabIndex        =   371
            Top             =   360
            Width           =   1095
         End
         Begin VB.Frame fraLoop 
            Caption         =   "Loop2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3135
            Index           =   1
            Left            =   -71880
            TabIndex        =   353
            Top             =   660
            Width           =   2175
            Begin VB.CheckBox chkUseLoop 
               Caption         =   "Enable"
               Height          =   375
               Index           =   1
               Left            =   720
               TabIndex        =   362
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtLoopDN 
               Height          =   390
               Index           =   1
               Left            =   480
               TabIndex        =   361
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopIN 
               Height          =   390
               Index           =   1
               Left            =   480
               TabIndex        =   360
               Text            =   "0"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtLoopPN 
               Height          =   390
               Index           =   1
               Left            =   480
               TabIndex        =   359
               Text            =   "0"
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtLoopTC 
               Height          =   390
               Index           =   1
               Left            =   480
               TabIndex        =   358
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopCN 
               Height          =   390
               Index           =   1
               Left            =   840
               TabIndex        =   357
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopCV 
               Height          =   390
               Index           =   1
               Left            =   1560
               TabIndex        =   356
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopRT 
               Height          =   390
               Index           =   1
               Left            =   1440
               TabIndex        =   355
               Text            =   "1"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopFT 
               Height          =   390
               Index           =   1
               Left            =   1425
               TabIndex        =   354
               Text            =   "-1"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   36
               Left            =   120
               TabIndex        =   370
               Top             =   2160
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "I:"
               Height          =   270
               Index           =   37
               Left            =   120
               TabIndex        =   369
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "P:"
               Height          =   270
               Index           =   38
               Left            =   120
               TabIndex        =   368
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No:"
               Height          =   270
               Index           =   71
               Left            =   30
               TabIndex        =   367
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Alarm:"
               Height          =   270
               Index           =   121
               Left            =   120
               TabIndex        =   366
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "±"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   122
               Left            =   1440
               TabIndex        =   365
               Top             =   2640
               Width           =   105
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "R:"
               Height          =   270
               Index           =   132
               Left            =   1200
               TabIndex        =   364
               Top             =   720
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   140
               Left            =   1200
               TabIndex        =   363
               Top             =   2160
               Width           =   210
            End
         End
         Begin VB.Frame fraLoop 
            Caption         =   "Loop3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3135
            Index           =   2
            Left            =   -69600
            TabIndex        =   335
            Top             =   660
            Width           =   2175
            Begin VB.CheckBox chkUseLoop 
               Caption         =   "Enable"
               Height          =   375
               Index           =   2
               Left            =   720
               TabIndex        =   344
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtLoopDN 
               Height          =   390
               Index           =   2
               Left            =   480
               TabIndex        =   343
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopIN 
               Height          =   390
               Index           =   2
               Left            =   480
               TabIndex        =   342
               Text            =   "0"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtLoopPN 
               Height          =   390
               Index           =   2
               Left            =   480
               TabIndex        =   341
               Text            =   "0"
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtLoopTC 
               Height          =   390
               Index           =   2
               Left            =   480
               TabIndex        =   340
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopCN 
               Height          =   390
               Index           =   2
               Left            =   840
               TabIndex        =   339
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopCV 
               Height          =   390
               Index           =   2
               Left            =   1560
               TabIndex        =   338
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopRT 
               Height          =   390
               Index           =   2
               Left            =   1440
               TabIndex        =   337
               Text            =   "1"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopFT 
               Height          =   390
               Index           =   2
               Left            =   1425
               TabIndex        =   336
               Text            =   "-1"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   44
               Left            =   120
               TabIndex        =   352
               Top             =   2160
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "I:"
               Height          =   270
               Index           =   45
               Left            =   120
               TabIndex        =   351
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "P:"
               Height          =   270
               Index           =   46
               Left            =   120
               TabIndex        =   350
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No:"
               Height          =   270
               Index           =   72
               Left            =   30
               TabIndex        =   349
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Alarm:"
               Height          =   270
               Index           =   123
               Left            =   120
               TabIndex        =   348
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "±"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   124
               Left            =   1440
               TabIndex        =   347
               Top             =   2640
               Width           =   105
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "R:"
               Height          =   270
               Index           =   133
               Left            =   1200
               TabIndex        =   346
               Top             =   720
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   141
               Left            =   1200
               TabIndex        =   345
               Top             =   2160
               Width           =   210
            End
         End
         Begin VB.Frame fraLoop 
            Caption         =   "Loop4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3135
            Index           =   3
            Left            =   -67320
            TabIndex        =   317
            Top             =   660
            Width           =   2175
            Begin VB.CheckBox chkUseLoop 
               Caption         =   "Enable"
               Height          =   375
               Index           =   3
               Left            =   720
               TabIndex        =   326
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtLoopDN 
               Height          =   390
               Index           =   3
               Left            =   480
               TabIndex        =   325
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopIN 
               Height          =   390
               Index           =   3
               Left            =   480
               TabIndex        =   324
               Text            =   "0"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtLoopPN 
               Height          =   390
               Index           =   3
               Left            =   480
               TabIndex        =   323
               Text            =   "0"
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtLoopTC 
               Height          =   390
               Index           =   3
               Left            =   480
               TabIndex        =   322
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopCN 
               Height          =   390
               Index           =   3
               Left            =   840
               TabIndex        =   321
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopCV 
               Height          =   390
               Index           =   3
               Left            =   1560
               TabIndex        =   320
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopRT 
               Height          =   390
               Index           =   3
               Left            =   1440
               TabIndex        =   319
               Text            =   "1"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopFT 
               Height          =   390
               Index           =   3
               Left            =   1425
               TabIndex        =   318
               Text            =   "-1"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   52
               Left            =   120
               TabIndex        =   334
               Top             =   2160
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "I:"
               Height          =   270
               Index           =   53
               Left            =   120
               TabIndex        =   333
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "P:"
               Height          =   270
               Index           =   54
               Left            =   120
               TabIndex        =   332
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No:"
               Height          =   270
               Index           =   73
               Left            =   30
               TabIndex        =   331
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Alarm:"
               Height          =   270
               Index           =   125
               Left            =   120
               TabIndex        =   330
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "±"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   126
               Left            =   1440
               TabIndex        =   329
               Top             =   2640
               Width           =   105
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "R:"
               Height          =   270
               Index           =   134
               Left            =   1200
               TabIndex        =   328
               Top             =   720
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   142
               Left            =   1200
               TabIndex        =   327
               Top             =   2160
               Width           =   210
            End
         End
         Begin VB.Frame fraLoop 
            Caption         =   "Loop5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3135
            Index           =   4
            Left            =   -65040
            TabIndex        =   299
            Top             =   660
            Width           =   2175
            Begin VB.CheckBox chkUseLoop 
               Caption         =   "Enable"
               Height          =   375
               Index           =   4
               Left            =   720
               TabIndex        =   308
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtLoopDN 
               Height          =   390
               Index           =   4
               Left            =   480
               TabIndex        =   307
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopIN 
               Height          =   390
               Index           =   4
               Left            =   480
               TabIndex        =   306
               Text            =   "0"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtLoopPN 
               Height          =   390
               Index           =   4
               Left            =   480
               TabIndex        =   305
               Text            =   "0"
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtLoopTC 
               Height          =   390
               Index           =   4
               Left            =   480
               TabIndex        =   304
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopCN 
               Height          =   390
               Index           =   4
               Left            =   840
               TabIndex        =   303
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopCV 
               Height          =   390
               Index           =   4
               Left            =   1560
               TabIndex        =   302
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopRT 
               Height          =   390
               Index           =   4
               Left            =   1440
               TabIndex        =   301
               Text            =   "1"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopFT 
               Height          =   390
               Index           =   4
               Left            =   1425
               TabIndex        =   300
               Text            =   "-1"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   60
               Left            =   120
               TabIndex        =   316
               Top             =   2160
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "I:"
               Height          =   270
               Index           =   61
               Left            =   120
               TabIndex        =   315
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "P:"
               Height          =   270
               Index           =   62
               Left            =   120
               TabIndex        =   314
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No:"
               Height          =   270
               Index           =   74
               Left            =   30
               TabIndex        =   313
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Alarm:"
               Height          =   270
               Index           =   127
               Left            =   120
               TabIndex        =   312
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "±"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   128
               Left            =   1440
               TabIndex        =   311
               Top             =   2640
               Width           =   105
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "R:"
               Height          =   270
               Index           =   135
               Left            =   1200
               TabIndex        =   310
               Top             =   720
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   143
               Left            =   1200
               TabIndex        =   309
               Top             =   2160
               Width           =   210
            End
         End
         Begin VB.Frame fraLoop 
            Caption         =   "Loop6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3135
            Index           =   5
            Left            =   -62760
            TabIndex        =   281
            Top             =   660
            Width           =   2295
            Begin VB.TextBox txtLoopTC 
               Height          =   390
               Index           =   5
               Left            =   600
               TabIndex        =   290
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopPN 
               Height          =   390
               Index           =   5
               Left            =   600
               TabIndex        =   289
               Text            =   "0"
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtLoopIN 
               Height          =   390
               Index           =   5
               Left            =   600
               TabIndex        =   288
               Text            =   "0"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtLoopDN 
               Height          =   390
               Index           =   5
               Left            =   600
               TabIndex        =   287
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.CheckBox chkUseLoop 
               Caption         =   "Enable"
               Height          =   375
               Index           =   5
               Left            =   720
               TabIndex        =   286
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtLoopCN 
               Height          =   390
               Index           =   5
               Left            =   960
               TabIndex        =   285
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopCV 
               Height          =   390
               Index           =   5
               Left            =   1680
               TabIndex        =   284
               Text            =   "0"
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtLoopRT 
               Height          =   390
               Index           =   5
               Left            =   1560
               TabIndex        =   283
               Text            =   "1"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopFT 
               Height          =   390
               Index           =   5
               Left            =   1545
               TabIndex        =   282
               Text            =   "-1"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No:"
               Height          =   270
               Index           =   113
               Left            =   150
               TabIndex        =   298
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "P:"
               Height          =   270
               Index           =   114
               Left            =   240
               TabIndex        =   297
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "I:"
               Height          =   270
               Index           =   115
               Left            =   240
               TabIndex        =   296
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   116
               Left            =   240
               TabIndex        =   295
               Top             =   2160
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Alarm:"
               Height          =   270
               Index           =   129
               Left            =   240
               TabIndex        =   294
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "±"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   130
               Left            =   1560
               TabIndex        =   293
               Top             =   2640
               Width           =   105
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "R:"
               Height          =   270
               Index           =   136
               Left            =   1320
               TabIndex        =   292
               Top             =   720
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   144
               Left            =   1320
               TabIndex        =   291
               Top             =   2160
               Width           =   210
            End
         End
         Begin TabDlg.SSTab tabSCR 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   40
            Top             =   3900
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   4683
            _Version        =   393216
            TabOrientation  =   2
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "SCR"
            TabPicture(0)   =   "frmRecipeEdit.frx":91248
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbName(86)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lbName(84)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lbName(83)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lbName(82)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lbName(75)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lbName(31)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lbName(30)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lbName(29)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lbName(28)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "lbName(189)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "lbName(99)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "lbName(95)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "lbName(91)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "lbName(87)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "lbName(76)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "lbName(32)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "lbName(33)"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "lbName(34)"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "lbName(35)"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "lbName(39)"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "lbName(100)"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "lbName(96)"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "lbName(92)"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "lbName(88)"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "lbName(77)"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "lbName(40)"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "lbName(41)"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "lbName(42)"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "lbName(43)"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "lbName(47)"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "lbName(101)"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "lbName(97)"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "lbName(93)"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).Control(33)=   "lbName(89)"
            Tab(0).Control(33).Enabled=   0   'False
            Tab(0).Control(34)=   "lbName(78)"
            Tab(0).Control(34).Enabled=   0   'False
            Tab(0).Control(35)=   "lbName(48)"
            Tab(0).Control(35).Enabled=   0   'False
            Tab(0).Control(36)=   "lbName(49)"
            Tab(0).Control(36).Enabled=   0   'False
            Tab(0).Control(37)=   "lbName(50)"
            Tab(0).Control(37).Enabled=   0   'False
            Tab(0).Control(38)=   "lbName(51)"
            Tab(0).Control(38).Enabled=   0   'False
            Tab(0).Control(39)=   "lbName(55)"
            Tab(0).Control(39).Enabled=   0   'False
            Tab(0).Control(40)=   "lbName(102)"
            Tab(0).Control(40).Enabled=   0   'False
            Tab(0).Control(41)=   "lbName(98)"
            Tab(0).Control(41).Enabled=   0   'False
            Tab(0).Control(42)=   "lbName(94)"
            Tab(0).Control(42).Enabled=   0   'False
            Tab(0).Control(43)=   "lbName(90)"
            Tab(0).Control(43).Enabled=   0   'False
            Tab(0).Control(44)=   "lbName(79)"
            Tab(0).Control(44).Enabled=   0   'False
            Tab(0).Control(45)=   "lbName(56)"
            Tab(0).Control(45).Enabled=   0   'False
            Tab(0).Control(46)=   "lbName(57)"
            Tab(0).Control(46).Enabled=   0   'False
            Tab(0).Control(47)=   "lbName(58)"
            Tab(0).Control(47).Enabled=   0   'False
            Tab(0).Control(48)=   "lbName(59)"
            Tab(0).Control(48).Enabled=   0   'False
            Tab(0).Control(49)=   "lbName(63)"
            Tab(0).Control(49).Enabled=   0   'False
            Tab(0).Control(50)=   "lbName(103)"
            Tab(0).Control(50).Enabled=   0   'False
            Tab(0).Control(51)=   "lbName(104)"
            Tab(0).Control(51).Enabled=   0   'False
            Tab(0).Control(52)=   "lbName(105)"
            Tab(0).Control(52).Enabled=   0   'False
            Tab(0).Control(53)=   "lbName(106)"
            Tab(0).Control(53).Enabled=   0   'False
            Tab(0).Control(54)=   "lbName(107)"
            Tab(0).Control(54).Enabled=   0   'False
            Tab(0).Control(55)=   "lbName(108)"
            Tab(0).Control(55).Enabled=   0   'False
            Tab(0).Control(56)=   "lbName(109)"
            Tab(0).Control(56).Enabled=   0   'False
            Tab(0).Control(57)=   "lbName(110)"
            Tab(0).Control(57).Enabled=   0   'False
            Tab(0).Control(58)=   "lbName(111)"
            Tab(0).Control(58).Enabled=   0   'False
            Tab(0).Control(59)=   "lbName(112)"
            Tab(0).Control(59).Enabled=   0   'False
            Tab(0).Control(60)=   "txtLoopK(0)"
            Tab(0).Control(60).Enabled=   0   'False
            Tab(0).Control(61)=   "txtLoopJ(0)"
            Tab(0).Control(61).Enabled=   0   'False
            Tab(0).Control(62)=   "txtLoopH(0)"
            Tab(0).Control(62).Enabled=   0   'False
            Tab(0).Control(63)=   "txtLoopG(0)"
            Tab(0).Control(63).Enabled=   0   'False
            Tab(0).Control(64)=   "txtLoopF(0)"
            Tab(0).Control(64).Enabled=   0   'False
            Tab(0).Control(65)=   "txtLoopE(0)"
            Tab(0).Control(65).Enabled=   0   'False
            Tab(0).Control(66)=   "txtLoopD(0)"
            Tab(0).Control(66).Enabled=   0   'False
            Tab(0).Control(67)=   "txtLoopC(0)"
            Tab(0).Control(67).Enabled=   0   'False
            Tab(0).Control(68)=   "txtLoopB(0)"
            Tab(0).Control(68).Enabled=   0   'False
            Tab(0).Control(69)=   "txtLoopA(0)"
            Tab(0).Control(69).Enabled=   0   'False
            Tab(0).Control(70)=   "txtLoopK(1)"
            Tab(0).Control(70).Enabled=   0   'False
            Tab(0).Control(71)=   "txtLoopJ(1)"
            Tab(0).Control(71).Enabled=   0   'False
            Tab(0).Control(72)=   "txtLoopH(1)"
            Tab(0).Control(72).Enabled=   0   'False
            Tab(0).Control(73)=   "txtLoopG(1)"
            Tab(0).Control(73).Enabled=   0   'False
            Tab(0).Control(74)=   "txtLoopF(1)"
            Tab(0).Control(74).Enabled=   0   'False
            Tab(0).Control(75)=   "txtLoopA(1)"
            Tab(0).Control(75).Enabled=   0   'False
            Tab(0).Control(76)=   "txtLoopB(1)"
            Tab(0).Control(76).Enabled=   0   'False
            Tab(0).Control(77)=   "txtLoopC(1)"
            Tab(0).Control(77).Enabled=   0   'False
            Tab(0).Control(78)=   "txtLoopD(1)"
            Tab(0).Control(78).Enabled=   0   'False
            Tab(0).Control(79)=   "txtLoopE(1)"
            Tab(0).Control(79).Enabled=   0   'False
            Tab(0).Control(80)=   "txtLoopK(2)"
            Tab(0).Control(80).Enabled=   0   'False
            Tab(0).Control(81)=   "txtLoopJ(2)"
            Tab(0).Control(81).Enabled=   0   'False
            Tab(0).Control(82)=   "txtLoopH(2)"
            Tab(0).Control(82).Enabled=   0   'False
            Tab(0).Control(83)=   "txtLoopG(2)"
            Tab(0).Control(83).Enabled=   0   'False
            Tab(0).Control(84)=   "txtLoopF(2)"
            Tab(0).Control(84).Enabled=   0   'False
            Tab(0).Control(85)=   "txtLoopA(2)"
            Tab(0).Control(85).Enabled=   0   'False
            Tab(0).Control(86)=   "txtLoopB(2)"
            Tab(0).Control(86).Enabled=   0   'False
            Tab(0).Control(87)=   "txtLoopC(2)"
            Tab(0).Control(87).Enabled=   0   'False
            Tab(0).Control(88)=   "txtLoopD(2)"
            Tab(0).Control(88).Enabled=   0   'False
            Tab(0).Control(89)=   "txtLoopE(2)"
            Tab(0).Control(89).Enabled=   0   'False
            Tab(0).Control(90)=   "txtLoopK(3)"
            Tab(0).Control(90).Enabled=   0   'False
            Tab(0).Control(91)=   "txtLoopJ(3)"
            Tab(0).Control(91).Enabled=   0   'False
            Tab(0).Control(92)=   "txtLoopH(3)"
            Tab(0).Control(92).Enabled=   0   'False
            Tab(0).Control(93)=   "txtLoopG(3)"
            Tab(0).Control(93).Enabled=   0   'False
            Tab(0).Control(94)=   "txtLoopF(3)"
            Tab(0).Control(94).Enabled=   0   'False
            Tab(0).Control(95)=   "txtLoopA(3)"
            Tab(0).Control(95).Enabled=   0   'False
            Tab(0).Control(96)=   "txtLoopB(3)"
            Tab(0).Control(96).Enabled=   0   'False
            Tab(0).Control(97)=   "txtLoopC(3)"
            Tab(0).Control(97).Enabled=   0   'False
            Tab(0).Control(98)=   "txtLoopD(3)"
            Tab(0).Control(98).Enabled=   0   'False
            Tab(0).Control(99)=   "txtLoopE(3)"
            Tab(0).Control(99).Enabled=   0   'False
            Tab(0).Control(100)=   "txtLoopK(4)"
            Tab(0).Control(100).Enabled=   0   'False
            Tab(0).Control(101)=   "txtLoopJ(4)"
            Tab(0).Control(101).Enabled=   0   'False
            Tab(0).Control(102)=   "txtLoopH(4)"
            Tab(0).Control(102).Enabled=   0   'False
            Tab(0).Control(103)=   "txtLoopG(4)"
            Tab(0).Control(103).Enabled=   0   'False
            Tab(0).Control(104)=   "txtLoopF(4)"
            Tab(0).Control(104).Enabled=   0   'False
            Tab(0).Control(105)=   "txtLoopA(4)"
            Tab(0).Control(105).Enabled=   0   'False
            Tab(0).Control(106)=   "txtLoopB(4)"
            Tab(0).Control(106).Enabled=   0   'False
            Tab(0).Control(107)=   "txtLoopC(4)"
            Tab(0).Control(107).Enabled=   0   'False
            Tab(0).Control(108)=   "txtLoopD(4)"
            Tab(0).Control(108).Enabled=   0   'False
            Tab(0).Control(109)=   "txtLoopE(4)"
            Tab(0).Control(109).Enabled=   0   'False
            Tab(0).Control(110)=   "txtLoopK(5)"
            Tab(0).Control(110).Enabled=   0   'False
            Tab(0).Control(111)=   "txtLoopJ(5)"
            Tab(0).Control(111).Enabled=   0   'False
            Tab(0).Control(112)=   "txtLoopH(5)"
            Tab(0).Control(112).Enabled=   0   'False
            Tab(0).Control(113)=   "txtLoopG(5)"
            Tab(0).Control(113).Enabled=   0   'False
            Tab(0).Control(114)=   "txtLoopF(5)"
            Tab(0).Control(114).Enabled=   0   'False
            Tab(0).Control(115)=   "txtLoopA(5)"
            Tab(0).Control(115).Enabled=   0   'False
            Tab(0).Control(116)=   "txtLoopB(5)"
            Tab(0).Control(116).Enabled=   0   'False
            Tab(0).Control(117)=   "txtLoopC(5)"
            Tab(0).Control(117).Enabled=   0   'False
            Tab(0).Control(118)=   "txtLoopD(5)"
            Tab(0).Control(118).Enabled=   0   'False
            Tab(0).Control(119)=   "txtLoopE(5)"
            Tab(0).Control(119).Enabled=   0   'False
            Tab(0).ControlCount=   120
            TabCaption(1)   =   "MTC"
            TabPicture(1)   =   "frmRecipeEdit.frx":91264
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbName(208)"
            Tab(1).Control(1)=   "lbName(207)"
            Tab(1).Control(2)=   "lbName(206)"
            Tab(1).Control(3)=   "lbName(205)"
            Tab(1).Control(4)=   "lbName(204)"
            Tab(1).Control(5)=   "lbName(203)"
            Tab(1).Control(6)=   "lbName(202)"
            Tab(1).Control(7)=   "lbName(201)"
            Tab(1).Control(8)=   "lbName(200)"
            Tab(1).Control(9)=   "lbName(199)"
            Tab(1).Control(10)=   "lbName(198)"
            Tab(1).Control(11)=   "lbName(197)"
            Tab(1).Control(12)=   "lbName(196)"
            Tab(1).Control(13)=   "lbName(195)"
            Tab(1).Control(14)=   "lbName(194)"
            Tab(1).Control(15)=   "lbName(193)"
            Tab(1).Control(16)=   "lbName(192)"
            Tab(1).Control(17)=   "lbName(191)"
            Tab(1).Control(18)=   "lbName(190)"
            Tab(1).Control(19)=   "lbName(187)"
            Tab(1).Control(20)=   "lbName(186)"
            Tab(1).Control(21)=   "lbName(185)"
            Tab(1).Control(22)=   "lbName(184)"
            Tab(1).Control(23)=   "lbName(183)"
            Tab(1).Control(24)=   "lbName(182)"
            Tab(1).Control(25)=   "lbName(181)"
            Tab(1).Control(26)=   "lbName(180)"
            Tab(1).Control(27)=   "lbName(179)"
            Tab(1).Control(28)=   "lbName(178)"
            Tab(1).Control(29)=   "lbName(177)"
            Tab(1).Control(30)=   "lbName(176)"
            Tab(1).Control(31)=   "lbName(175)"
            Tab(1).Control(32)=   "lbName(174)"
            Tab(1).Control(33)=   "lbName(173)"
            Tab(1).Control(34)=   "lbName(172)"
            Tab(1).Control(35)=   "lbName(171)"
            Tab(1).Control(36)=   "lbName(170)"
            Tab(1).Control(37)=   "lbName(169)"
            Tab(1).Control(38)=   "lbName(168)"
            Tab(1).Control(39)=   "lbName(167)"
            Tab(1).Control(40)=   "lbName(166)"
            Tab(1).Control(41)=   "lbName(165)"
            Tab(1).Control(42)=   "lbName(164)"
            Tab(1).Control(43)=   "lbName(163)"
            Tab(1).Control(44)=   "lbName(162)"
            Tab(1).Control(45)=   "lbName(161)"
            Tab(1).Control(46)=   "lbName(160)"
            Tab(1).Control(47)=   "lbName(157)"
            Tab(1).Control(48)=   "lbName(156)"
            Tab(1).Control(49)=   "lbName(155)"
            Tab(1).Control(50)=   "lbName(154)"
            Tab(1).Control(51)=   "lbName(153)"
            Tab(1).Control(52)=   "lbName(152)"
            Tab(1).Control(53)=   "lbName(151)"
            Tab(1).Control(54)=   "lbName(150)"
            Tab(1).Control(55)=   "lbName(149)"
            Tab(1).Control(56)=   "lbName(148)"
            Tab(1).Control(57)=   "lbName(147)"
            Tab(1).Control(58)=   "lbName(146)"
            Tab(1).Control(59)=   "lbName(145)"
            Tab(1).Control(60)=   "txtLoopMK(0)"
            Tab(1).Control(61)=   "txtLoopMJ(0)"
            Tab(1).Control(62)=   "txtLoopMH(0)"
            Tab(1).Control(63)=   "txtLoopMG(0)"
            Tab(1).Control(64)=   "txtLoopMF(0)"
            Tab(1).Control(65)=   "txtLoopME(0)"
            Tab(1).Control(66)=   "txtLoopMD(0)"
            Tab(1).Control(67)=   "txtLoopMC(0)"
            Tab(1).Control(68)=   "txtLoopMB(0)"
            Tab(1).Control(69)=   "txtLoopMA(0)"
            Tab(1).Control(70)=   "txtLoopMK(1)"
            Tab(1).Control(71)=   "txtLoopMJ(1)"
            Tab(1).Control(72)=   "txtLoopMH(1)"
            Tab(1).Control(73)=   "txtLoopMG(1)"
            Tab(1).Control(74)=   "txtLoopMF(1)"
            Tab(1).Control(75)=   "txtLoopMA(1)"
            Tab(1).Control(76)=   "txtLoopMB(1)"
            Tab(1).Control(77)=   "txtLoopMC(1)"
            Tab(1).Control(78)=   "txtLoopMD(1)"
            Tab(1).Control(79)=   "txtLoopME(1)"
            Tab(1).Control(80)=   "txtLoopMK(2)"
            Tab(1).Control(81)=   "txtLoopMJ(2)"
            Tab(1).Control(82)=   "txtLoopMH(2)"
            Tab(1).Control(83)=   "txtLoopMG(2)"
            Tab(1).Control(84)=   "txtLoopMF(2)"
            Tab(1).Control(85)=   "txtLoopMA(2)"
            Tab(1).Control(86)=   "txtLoopMB(2)"
            Tab(1).Control(87)=   "txtLoopMC(2)"
            Tab(1).Control(88)=   "txtLoopMD(2)"
            Tab(1).Control(89)=   "txtLoopME(2)"
            Tab(1).Control(90)=   "txtLoopMK(3)"
            Tab(1).Control(91)=   "txtLoopMJ(3)"
            Tab(1).Control(92)=   "txtLoopMH(3)"
            Tab(1).Control(93)=   "txtLoopMG(3)"
            Tab(1).Control(94)=   "txtLoopMF(3)"
            Tab(1).Control(95)=   "txtLoopMA(3)"
            Tab(1).Control(96)=   "txtLoopMB(3)"
            Tab(1).Control(97)=   "txtLoopMC(3)"
            Tab(1).Control(98)=   "txtLoopMD(3)"
            Tab(1).Control(99)=   "txtLoopME(3)"
            Tab(1).Control(100)=   "txtLoopMK(4)"
            Tab(1).Control(101)=   "txtLoopMJ(4)"
            Tab(1).Control(102)=   "txtLoopMH(4)"
            Tab(1).Control(103)=   "txtLoopMG(4)"
            Tab(1).Control(104)=   "txtLoopMF(4)"
            Tab(1).Control(105)=   "txtLoopMA(4)"
            Tab(1).Control(106)=   "txtLoopMB(4)"
            Tab(1).Control(107)=   "txtLoopMC(4)"
            Tab(1).Control(108)=   "txtLoopMD(4)"
            Tab(1).Control(109)=   "txtLoopME(4)"
            Tab(1).Control(110)=   "txtLoopMK(5)"
            Tab(1).Control(111)=   "txtLoopMJ(5)"
            Tab(1).Control(112)=   "txtLoopMH(5)"
            Tab(1).Control(113)=   "txtLoopMG(5)"
            Tab(1).Control(114)=   "txtLoopMF(5)"
            Tab(1).Control(115)=   "txtLoopMA(5)"
            Tab(1).Control(116)=   "txtLoopMB(5)"
            Tab(1).Control(117)=   "txtLoopMC(5)"
            Tab(1).Control(118)=   "txtLoopMD(5)"
            Tab(1).Control(119)=   "txtLoopME(5)"
            Tab(1).ControlCount=   120
            Begin VB.TextBox txtLoopE 
               Height          =   390
               Index           =   5
               Left            =   12720
               TabIndex        =   160
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopD 
               Height          =   390
               Index           =   5
               Left            =   12720
               TabIndex        =   159
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopC 
               Height          =   390
               Index           =   5
               Left            =   12720
               TabIndex        =   158
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopB 
               Height          =   390
               Index           =   5
               Left            =   12720
               TabIndex        =   157
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopA 
               Height          =   390
               Index           =   5
               Left            =   12720
               TabIndex        =   156
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopF 
               Height          =   390
               Index           =   5
               Left            =   13680
               TabIndex        =   155
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopG 
               Height          =   390
               Index           =   5
               Left            =   13680
               TabIndex        =   154
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopH 
               Height          =   390
               Index           =   5
               Left            =   13680
               TabIndex        =   153
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopJ 
               Height          =   390
               Index           =   5
               Left            =   13680
               TabIndex        =   152
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopK 
               Height          =   390
               Index           =   5
               Left            =   13680
               TabIndex        =   151
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopE 
               Height          =   390
               Index           =   4
               Left            =   10320
               TabIndex        =   150
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopD 
               Height          =   390
               Index           =   4
               Left            =   10320
               TabIndex        =   149
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopC 
               Height          =   390
               Index           =   4
               Left            =   10320
               TabIndex        =   148
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopB 
               Height          =   390
               Index           =   4
               Left            =   10320
               TabIndex        =   147
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopA 
               Height          =   390
               Index           =   4
               Left            =   10320
               TabIndex        =   146
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopF 
               Height          =   390
               Index           =   4
               Left            =   11280
               TabIndex        =   145
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopG 
               Height          =   390
               Index           =   4
               Left            =   11280
               TabIndex        =   144
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopH 
               Height          =   390
               Index           =   4
               Left            =   11280
               TabIndex        =   143
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopJ 
               Height          =   390
               Index           =   4
               Left            =   11280
               TabIndex        =   142
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopK 
               Height          =   390
               Index           =   4
               Left            =   11280
               TabIndex        =   141
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopE 
               Height          =   390
               Index           =   3
               Left            =   8040
               TabIndex        =   140
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopD 
               Height          =   390
               Index           =   3
               Left            =   8040
               TabIndex        =   139
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopC 
               Height          =   390
               Index           =   3
               Left            =   8040
               TabIndex        =   138
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopB 
               Height          =   390
               Index           =   3
               Left            =   8040
               TabIndex        =   137
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopA 
               Height          =   390
               Index           =   3
               Left            =   8040
               TabIndex        =   136
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopF 
               Height          =   390
               Index           =   3
               Left            =   9000
               TabIndex        =   135
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopG 
               Height          =   390
               Index           =   3
               Left            =   9000
               TabIndex        =   134
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopH 
               Height          =   390
               Index           =   3
               Left            =   9000
               TabIndex        =   133
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopJ 
               Height          =   390
               Index           =   3
               Left            =   9000
               TabIndex        =   132
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopK 
               Height          =   390
               Index           =   3
               Left            =   9000
               TabIndex        =   131
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopE 
               Height          =   390
               Index           =   2
               Left            =   5760
               TabIndex        =   130
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopD 
               Height          =   390
               Index           =   2
               Left            =   5760
               TabIndex        =   129
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopC 
               Height          =   390
               Index           =   2
               Left            =   5760
               TabIndex        =   128
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopB 
               Height          =   390
               Index           =   2
               Left            =   5760
               TabIndex        =   127
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopA 
               Height          =   390
               Index           =   2
               Left            =   5760
               TabIndex        =   126
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopF 
               Height          =   390
               Index           =   2
               Left            =   6720
               TabIndex        =   125
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopG 
               Height          =   390
               Index           =   2
               Left            =   6720
               TabIndex        =   124
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopH 
               Height          =   390
               Index           =   2
               Left            =   6720
               TabIndex        =   123
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopJ 
               Height          =   390
               Index           =   2
               Left            =   6720
               TabIndex        =   122
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopK 
               Height          =   390
               Index           =   2
               Left            =   6720
               TabIndex        =   121
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopE 
               Height          =   390
               Index           =   1
               Left            =   3480
               TabIndex        =   120
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopD 
               Height          =   390
               Index           =   1
               Left            =   3480
               TabIndex        =   119
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopC 
               Height          =   390
               Index           =   1
               Left            =   3480
               TabIndex        =   118
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopB 
               Height          =   390
               Index           =   1
               Left            =   3480
               TabIndex        =   117
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopA 
               Height          =   390
               Index           =   1
               Left            =   3480
               TabIndex        =   116
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopF 
               Height          =   390
               Index           =   1
               Left            =   4440
               TabIndex        =   115
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopG 
               Height          =   390
               Index           =   1
               Left            =   4440
               TabIndex        =   114
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopH 
               Height          =   390
               Index           =   1
               Left            =   4440
               TabIndex        =   113
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopJ 
               Height          =   390
               Index           =   1
               Left            =   4440
               TabIndex        =   112
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopK 
               Height          =   390
               Index           =   1
               Left            =   4440
               TabIndex        =   111
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopA 
               Height          =   390
               Index           =   0
               Left            =   1200
               TabIndex        =   110
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopB 
               Height          =   390
               Index           =   0
               Left            =   1200
               TabIndex        =   109
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopC 
               Height          =   390
               Index           =   0
               Left            =   1200
               TabIndex        =   108
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopD 
               Height          =   390
               Index           =   0
               Left            =   1200
               TabIndex        =   107
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopE 
               Height          =   390
               Index           =   0
               Left            =   1200
               TabIndex        =   106
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopF 
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   105
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopG 
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   104
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopH 
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   103
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopJ 
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   102
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopK 
               Height          =   390
               Index           =   0
               Left            =   2160
               TabIndex        =   101
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopME 
               Height          =   390
               Index           =   5
               Left            =   -62280
               TabIndex        =   100
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMD 
               Height          =   390
               Index           =   5
               Left            =   -62280
               TabIndex        =   99
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMC 
               Height          =   390
               Index           =   5
               Left            =   -62280
               TabIndex        =   98
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMB 
               Height          =   390
               Index           =   5
               Left            =   -62280
               TabIndex        =   97
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMA 
               Height          =   390
               Index           =   5
               Left            =   -62280
               TabIndex        =   96
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMF 
               Height          =   390
               Index           =   5
               Left            =   -61320
               TabIndex        =   95
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMG 
               Height          =   390
               Index           =   5
               Left            =   -61320
               TabIndex        =   94
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMH 
               Height          =   390
               Index           =   5
               Left            =   -61320
               TabIndex        =   93
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMJ 
               Height          =   390
               Index           =   5
               Left            =   -61320
               TabIndex        =   92
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMK 
               Height          =   390
               Index           =   5
               Left            =   -61320
               TabIndex        =   91
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopME 
               Height          =   390
               Index           =   4
               Left            =   -64680
               TabIndex        =   90
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMD 
               Height          =   390
               Index           =   4
               Left            =   -64680
               TabIndex        =   89
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMC 
               Height          =   390
               Index           =   4
               Left            =   -64680
               TabIndex        =   88
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMB 
               Height          =   390
               Index           =   4
               Left            =   -64680
               TabIndex        =   87
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMA 
               Height          =   390
               Index           =   4
               Left            =   -64680
               TabIndex        =   86
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMF 
               Height          =   390
               Index           =   4
               Left            =   -63720
               TabIndex        =   85
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMG 
               Height          =   390
               Index           =   4
               Left            =   -63720
               TabIndex        =   84
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMH 
               Height          =   390
               Index           =   4
               Left            =   -63720
               TabIndex        =   83
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMJ 
               Height          =   390
               Index           =   4
               Left            =   -63720
               TabIndex        =   82
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMK 
               Height          =   390
               Index           =   4
               Left            =   -63720
               TabIndex        =   81
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopME 
               Height          =   390
               Index           =   3
               Left            =   -66960
               TabIndex        =   80
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMD 
               Height          =   390
               Index           =   3
               Left            =   -66960
               TabIndex        =   79
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMC 
               Height          =   390
               Index           =   3
               Left            =   -66960
               TabIndex        =   78
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMB 
               Height          =   390
               Index           =   3
               Left            =   -66960
               TabIndex        =   77
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMA 
               Height          =   390
               Index           =   3
               Left            =   -66960
               TabIndex        =   76
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMF 
               Height          =   390
               Index           =   3
               Left            =   -66000
               TabIndex        =   75
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMG 
               Height          =   390
               Index           =   3
               Left            =   -66000
               TabIndex        =   74
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMH 
               Height          =   390
               Index           =   3
               Left            =   -66000
               TabIndex        =   73
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMJ 
               Height          =   390
               Index           =   3
               Left            =   -66000
               TabIndex        =   72
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMK 
               Height          =   390
               Index           =   3
               Left            =   -66000
               TabIndex        =   71
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopME 
               Height          =   390
               Index           =   2
               Left            =   -69240
               TabIndex        =   70
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMD 
               Height          =   390
               Index           =   2
               Left            =   -69240
               TabIndex        =   69
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMC 
               Height          =   390
               Index           =   2
               Left            =   -69240
               TabIndex        =   68
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMB 
               Height          =   390
               Index           =   2
               Left            =   -69240
               TabIndex        =   67
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMA 
               Height          =   390
               Index           =   2
               Left            =   -69240
               TabIndex        =   66
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMF 
               Height          =   390
               Index           =   2
               Left            =   -68280
               TabIndex        =   65
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMG 
               Height          =   390
               Index           =   2
               Left            =   -68280
               TabIndex        =   64
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMH 
               Height          =   390
               Index           =   2
               Left            =   -68280
               TabIndex        =   63
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMJ 
               Height          =   390
               Index           =   2
               Left            =   -68280
               TabIndex        =   62
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMK 
               Height          =   390
               Index           =   2
               Left            =   -68280
               TabIndex        =   61
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopME 
               Height          =   390
               Index           =   1
               Left            =   -71520
               TabIndex        =   60
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMD 
               Height          =   390
               Index           =   1
               Left            =   -71520
               TabIndex        =   59
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMC 
               Height          =   390
               Index           =   1
               Left            =   -71520
               TabIndex        =   58
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMB 
               Height          =   390
               Index           =   1
               Left            =   -71520
               TabIndex        =   57
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMA 
               Height          =   390
               Index           =   1
               Left            =   -71520
               TabIndex        =   56
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMF 
               Height          =   390
               Index           =   1
               Left            =   -70560
               TabIndex        =   55
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMG 
               Height          =   390
               Index           =   1
               Left            =   -70560
               TabIndex        =   54
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMH 
               Height          =   390
               Index           =   1
               Left            =   -70560
               TabIndex        =   53
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMJ 
               Height          =   390
               Index           =   1
               Left            =   -70560
               TabIndex        =   52
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMK 
               Height          =   390
               Index           =   1
               Left            =   -70560
               TabIndex        =   51
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMA 
               Height          =   390
               Index           =   0
               Left            =   -73800
               TabIndex        =   50
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMB 
               Height          =   390
               Index           =   0
               Left            =   -73800
               TabIndex        =   49
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMC 
               Height          =   390
               Index           =   0
               Left            =   -73800
               TabIndex        =   48
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMD 
               Height          =   390
               Index           =   0
               Left            =   -73800
               TabIndex        =   47
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopME 
               Height          =   390
               Index           =   0
               Left            =   -73800
               TabIndex        =   46
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.TextBox txtLoopMF 
               Height          =   390
               Index           =   0
               Left            =   -72840
               TabIndex        =   45
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtLoopMG 
               Height          =   390
               Index           =   0
               Left            =   -72840
               TabIndex        =   44
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtLoopMH 
               Height          =   390
               Index           =   0
               Left            =   -72840
               TabIndex        =   43
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtLoopMJ 
               Height          =   390
               Index           =   0
               Left            =   -72840
               TabIndex        =   42
               Text            =   "0"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtLoopMK 
               Height          =   390
               Index           =   0
               Left            =   -72840
               TabIndex        =   41
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   112
               Left            =   13440
               TabIndex        =   280
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   111
               Left            =   13440
               TabIndex        =   279
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   110
               Left            =   13440
               TabIndex        =   278
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   109
               Left            =   13440
               TabIndex        =   277
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "K:"
               Height          =   270
               Index           =   108
               Left            =   13440
               TabIndex        =   276
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   107
               Left            =   12360
               TabIndex        =   275
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   106
               Left            =   12360
               TabIndex        =   274
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   105
               Left            =   12360
               TabIndex        =   273
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   104
               Left            =   12360
               TabIndex        =   272
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   103
               Left            =   12360
               TabIndex        =   271
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   63
               Left            =   9960
               TabIndex        =   270
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   59
               Left            =   9960
               TabIndex        =   269
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   58
               Left            =   9960
               TabIndex        =   268
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   57
               Left            =   9960
               TabIndex        =   267
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   56
               Left            =   9960
               TabIndex        =   266
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   79
               Left            =   11040
               TabIndex        =   265
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   90
               Left            =   11040
               TabIndex        =   264
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   94
               Left            =   11040
               TabIndex        =   263
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   98
               Left            =   11040
               TabIndex        =   262
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "K:"
               Height          =   270
               Index           =   102
               Left            =   11040
               TabIndex        =   261
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   55
               Left            =   7680
               TabIndex        =   260
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   51
               Left            =   7680
               TabIndex        =   259
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   50
               Left            =   7680
               TabIndex        =   258
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   49
               Left            =   7680
               TabIndex        =   257
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   48
               Left            =   7680
               TabIndex        =   256
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   78
               Left            =   8760
               TabIndex        =   255
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   89
               Left            =   8760
               TabIndex        =   254
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   93
               Left            =   8760
               TabIndex        =   253
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   97
               Left            =   8760
               TabIndex        =   252
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "K:"
               Height          =   270
               Index           =   101
               Left            =   8760
               TabIndex        =   251
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   47
               Left            =   5400
               TabIndex        =   250
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   43
               Left            =   5400
               TabIndex        =   249
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   42
               Left            =   5400
               TabIndex        =   248
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   41
               Left            =   5400
               TabIndex        =   247
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   40
               Left            =   5400
               TabIndex        =   246
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   77
               Left            =   6480
               TabIndex        =   245
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   88
               Left            =   6480
               TabIndex        =   244
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   92
               Left            =   6480
               TabIndex        =   243
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   96
               Left            =   6480
               TabIndex        =   242
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "K:"
               Height          =   270
               Index           =   100
               Left            =   6480
               TabIndex        =   241
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   39
               Left            =   3120
               TabIndex        =   240
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   35
               Left            =   3120
               TabIndex        =   239
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   34
               Left            =   3120
               TabIndex        =   238
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   33
               Left            =   3120
               TabIndex        =   237
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   32
               Left            =   3120
               TabIndex        =   236
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   76
               Left            =   4200
               TabIndex        =   235
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   87
               Left            =   4200
               TabIndex        =   234
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   91
               Left            =   4200
               TabIndex        =   233
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   95
               Left            =   4200
               TabIndex        =   232
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "K:"
               Height          =   270
               Index           =   99
               Left            =   4200
               TabIndex        =   231
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   189
               Left            =   840
               TabIndex        =   230
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   28
               Left            =   840
               TabIndex        =   229
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   29
               Left            =   840
               TabIndex        =   228
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   30
               Left            =   840
               TabIndex        =   227
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   31
               Left            =   840
               TabIndex        =   226
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   75
               Left            =   1920
               TabIndex        =   225
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   82
               Left            =   1920
               TabIndex        =   224
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   83
               Left            =   1920
               TabIndex        =   223
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   84
               Left            =   1920
               TabIndex        =   222
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "K:"
               Height          =   270
               Index           =   86
               Left            =   1920
               TabIndex        =   221
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   145
               Left            =   -61560
               TabIndex        =   220
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   146
               Left            =   -61560
               TabIndex        =   219
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   147
               Left            =   -61560
               TabIndex        =   218
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   148
               Left            =   -61560
               TabIndex        =   217
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "N"
               Height          =   270
               Index           =   149
               Left            =   -61560
               TabIndex        =   216
               Top             =   2160
               Width           =   165
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   150
               Left            =   -62640
               TabIndex        =   215
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   151
               Left            =   -62640
               TabIndex        =   214
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   152
               Left            =   -62640
               TabIndex        =   213
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   153
               Left            =   -62640
               TabIndex        =   212
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   154
               Left            =   -62640
               TabIndex        =   211
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   155
               Left            =   -65040
               TabIndex        =   210
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   156
               Left            =   -65040
               TabIndex        =   209
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   157
               Left            =   -65040
               TabIndex        =   208
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   160
               Left            =   -65040
               TabIndex        =   207
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   161
               Left            =   -65040
               TabIndex        =   206
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   162
               Left            =   -63960
               TabIndex        =   205
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   163
               Left            =   -63960
               TabIndex        =   204
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   164
               Left            =   -63960
               TabIndex        =   203
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   165
               Left            =   -63960
               TabIndex        =   202
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "N"
               Height          =   270
               Index           =   166
               Left            =   -63960
               TabIndex        =   201
               Top             =   2160
               Width           =   165
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   167
               Left            =   -67320
               TabIndex        =   200
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   168
               Left            =   -67320
               TabIndex        =   199
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   169
               Left            =   -67320
               TabIndex        =   198
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   170
               Left            =   -67320
               TabIndex        =   197
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   171
               Left            =   -67320
               TabIndex        =   196
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   172
               Left            =   -66240
               TabIndex        =   195
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   173
               Left            =   -66240
               TabIndex        =   194
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   174
               Left            =   -66240
               TabIndex        =   193
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   175
               Left            =   -66240
               TabIndex        =   192
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "N"
               Height          =   270
               Index           =   176
               Left            =   -66240
               TabIndex        =   191
               Top             =   2160
               Width           =   165
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   177
               Left            =   -69600
               TabIndex        =   190
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   178
               Left            =   -69600
               TabIndex        =   189
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   179
               Left            =   -69600
               TabIndex        =   188
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   180
               Left            =   -69600
               TabIndex        =   187
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   181
               Left            =   -69600
               TabIndex        =   186
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   182
               Left            =   -68520
               TabIndex        =   185
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   183
               Left            =   -68520
               TabIndex        =   184
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   184
               Left            =   -68520
               TabIndex        =   183
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   185
               Left            =   -68520
               TabIndex        =   182
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "N"
               Height          =   270
               Index           =   186
               Left            =   -68520
               TabIndex        =   181
               Top             =   2160
               Width           =   165
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   187
               Left            =   -71880
               TabIndex        =   180
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   190
               Left            =   -71880
               TabIndex        =   179
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   191
               Left            =   -71880
               TabIndex        =   178
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   192
               Left            =   -71880
               TabIndex        =   177
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   193
               Left            =   -71880
               TabIndex        =   176
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   194
               Left            =   -70800
               TabIndex        =   175
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   195
               Left            =   -70800
               TabIndex        =   174
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   196
               Left            =   -70800
               TabIndex        =   173
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   197
               Left            =   -70800
               TabIndex        =   172
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "N"
               Height          =   270
               Index           =   198
               Left            =   -70800
               TabIndex        =   171
               Top             =   2160
               Width           =   165
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               Height          =   270
               Index           =   199
               Left            =   -74160
               TabIndex        =   170
               Top             =   240
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "B:"
               Height          =   270
               Index           =   200
               Left            =   -74160
               TabIndex        =   169
               Top             =   720
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               Height          =   270
               Index           =   201
               Left            =   -74160
               TabIndex        =   168
               Top             =   1200
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "D:"
               Height          =   270
               Index           =   202
               Left            =   -74160
               TabIndex        =   167
               Top             =   1680
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               Height          =   270
               Index           =   203
               Left            =   -74160
               TabIndex        =   166
               Top             =   2160
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "F:"
               Height          =   270
               Index           =   204
               Left            =   -73080
               TabIndex        =   165
               Top             =   240
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "G:"
               Height          =   270
               Index           =   205
               Left            =   -73080
               TabIndex        =   164
               Top             =   720
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "H:"
               Height          =   270
               Index           =   206
               Left            =   -73080
               TabIndex        =   163
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "J:"
               Height          =   270
               Index           =   207
               Left            =   -73080
               TabIndex        =   162
               Top             =   1680
               Width           =   180
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "N"
               Height          =   270
               Index           =   208
               Left            =   -73080
               TabIndex        =   161
               Top             =   2160
               Width           =   165
            End
         End
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Recipe Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   486
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbRecipeName 
         Caption         =   "Unknown"
         Height          =   375
         Left            =   2040
         TabIndex        =   485
         Top             =   240
         Width           =   8295
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Owner:"
         Height          =   270
         Index           =   24
         Left            =   15960
         TabIndex        =   484
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   ": "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   25
         Left            =   1800
         TabIndex        =   483
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.Label lbName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "R:"
      Height          =   270
      Index           =   258
      Left            =   0
      TabIndex        =   642
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "frmRecipeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================================
'Copyright: Aries Liu
'Author: Aries Liu
'Date: 7/05/2006
'========================================================================================================
Option Explicit

Public sngRecipeProportional    As Single
Public sngRecipeProportional2    As Single
Public sngRecipeIntegral            As Single
Public sngRecipeIntegral2            As Single
Public sngRecipeDerivational    As Single
Public sngRecipePredit   As Single
Public sngRecipeFeedForward   As Single
Public sngRecipeOverTemp        As Single
Public sngRecipeUnderTemp        As Single
Public sngRecipeOverPressure        As Single
Public intRecipeTempInputType As Integer
Public sngRecipeRampDownPower        As Single
Public sngRecipeIntLimit            As Single
Public sngRecipeTotalSec            As Single
Public GasNames            As Collection
Public GasUnits            As Collection
Public GasMaxSlmps            As Collection
Public fileName As String




Public sngRecipePP    As Single
Public sngRecipeII    As Single
Public sngRecipeDD    As Single

Public intCurrRowSel As Integer
Public intCurrColSel As Integer

Public blnSave     As Boolean
Private strOriginActionName As String

Private TxtExist     As Boolean


Private Function GetTemptureColl() As Collection
Dim i As Integer
Dim Rows As Integer
Dim tempColl As New Collection
Dim CellText As String
Rows = hfgRecipe.Rows
If Rows > 0 Then
For i = 1 To Rows - 1
CellText = hfgRecipe.TextMatrix(i, 2)
If CellText <> "0" Then
tempColl.Add CInt(CellText)
End If
Next i
End If
Set GetTemptureColl = tempColl
End Function

Private Function GetMAX(num1 As Integer, num2 As Integer) As Integer
Dim Max As Integer
If num1 >= num2 Then
Max = num1
Else
Max = num2
End If
GetMAX = Max
End Function


Private Function GetMAXTempture(temp As Collection) As Integer
Dim MAXTempture As Integer
Dim j As Integer
If temp.Count > 0 Then
If temp.Count = 1 Then
MAXTempture = temp(1)
Else
MAXTempture = 0
For j = 1 To temp.Count
MAXTempture = GetMAX(MAXTempture, temp(j))
Next j
End If
End If
GetMAXTempture = MAXTempture
End Function



Private Sub cmbRecipeAction_Click()
   
    hfgRecipe.TextMatrix(hfgRecipe.RowSel, hfgRecipe.ColSel) = cmbRecipeAction.text
    If TxtExist = False Then
    
         Call CheckRecipeRule(hfgRecipe.TextMatrix(hfgRecipe.RowSel - 1, GB_PROCESS_ACTION), _
                                    hfgRecipe.TextMatrix(hfgRecipe.RowSel, GB_PROCESS_ACTION), _
                                    hfgRecipe.RowSel)
    Else
    Dim i As Integer
    Dim Config_Step As Collection
    Dim TxtLine() As String
    Dim k As Integer
    Dim j As Integer
    Set Config_Step = ReadTextFileToArray(App.Path + "\Config\ProcessStep.txt")
    For j = 1 To Config_Step.Count
    If Split(Config_Step(j), ";")(0) = cmbRecipeAction.text Then
     TxtLine = Split(Config_Step(j), ";")
         Exit For
       End If
    Next j
      For i = 2 To hfgRecipe.Cols - 1
      If i <= UBound(TxtLine) Then
      hfgRecipe.TextMatrix(hfgRecipe.RowSel, i) = TxtLine(i - 1)
      Else
      hfgRecipe.TextMatrix(hfgRecipe.RowSel, i) = "0"
      End If
      Next i
    End If
    blnSave = True
End Sub

Private Sub cmbRecipeAction_LostFocus()
    cmbRecipeAction.Visible = False
End Sub


Private Sub CmdBtn_BuildProcessStrp_Click()
frmProcessStepBuild.Show
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    intCurrRowSel = hfgRecipe.RowSel
    If hfgRecipe.RowSel > 0 And hfgRecipe.ColSel > 0 Then
        For i = intCurrRowSel To 49
            For j = 1 To GasNames.Count + 3
                hfgRecipe.TextMatrix(i, j) = hfgRecipe.TextMatrix(i + 1, j)
            Next j
        Next i
    End If
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
    Dim j As Integer
    
    intCurrRowSel = hfgRecipe.RowSel
    If hfgRecipe.RowSel > 0 And hfgRecipe.ColSel > 0 Then
        For i = 49 To intCurrRowSel Step -1
            For j = 1 To GasNames.Count + 3
                hfgRecipe.TextMatrix(i + 1, j) = hfgRecipe.TextMatrix(i, j)
            Next j
        Next i
    End If
End Sub

Private Sub cmdReadAz1_Click()
    Dim i As Integer
    Dim j As Integer
    Dim Index As Integer
    Dim b1, b2 As Boolean
    Dim sngData(0 To 100) As Single
    
    b1 = frmAz1.ReadPID(sngData)
    If b1 = True Then
        For i = 0 To 3
            Az1.sngPN(i) = sngData(i) / 100
            Az1.sngIN(i) = sngData(i + 4) / 100
            Az1.sngDN(i) = sngData(i + 8) / 100
            Az1.sngRT(i) = sngData(i + 12) / 10000
            Az1.sngST(i) = sngData(i + 16) / 10
'            Az1.sngOffset(i) = sngData(i + 20) / 10
            
            
            txtAz1PN(i).text = CStr(Az1.sngPN(i))
            txtAz1IN(i).text = CStr(Az1.sngIN(i))
            txtAz1DN(i).text = CStr(Az1.sngDN(i))
            txtAz1RT(i).text = CStr(Az1.sngRT(i))
            txtAz1ST(i).text = CStr(Az1.sngST(i))
'            txtAz1Offset(i).Text = CStr(Az1.sngOffset(i))
        Next i
    End If
    
    b2 = frmAz2.ReadPID(sngData)
    If b2 = True Then
        For i = 0 To 3
            Az2.sngPN(i) = sngData(i) / 100
            Az2.sngIN(i) = sngData(i + 4) / 100
            Az2.sngDN(i) = sngData(i + 8) / 100
            Az2.sngRT(i) = sngData(i + 12) / 10000
            Az2.sngST(i) = sngData(i + 16) / 10
'            Az2.sngOffset(i) = sngData(i + 20) / 100
            
            txtAz2PN(i).text = CStr(Az2.sngPN(i))
            txtAz2IN(i).text = CStr(Az2.sngIN(i))
            txtAz2DN(i).text = CStr(Az2.sngDN(i))
            txtAz2RT(i).text = CStr(Az2.sngRT(i))
            txtAz2ST(i).text = CStr(Az2.sngST(i))
'            txtAz2Offset(i).Text = CStr(Az2.sngOffset(i))
        Next i
    End If
    
    If b1 And b2 Then
        MsgBox "讀回成功!", vbOK
    Else
        MsgBox "讀回失敗!", vbOK
    End If
    
End Sub

Private Sub cmdWriteAz1_Click()
    Dim i As Integer
    Dim j As Integer
    Dim iRet As Integer
    Dim b1, b2 As Boolean
'    Dim intPara(0 To 19) As Integer
     Dim intPara(0 To 23) As Integer
     If OffsetWriteToTcm = 1 Then
       frmConfiguration.WriteOffsetToTCM
     End If
    iRet = MsgBox("寫入控制器?", vbOKCancel)
    If iRet = vbOK Then
        If Para.UseAz1 Then
            For i = 0 To 4
'             For i = 0 To 5
                For j = 0 To 3
                    If i = 0 Then intPara(i * 4 + j) = CSng(txtAz1PN(j).text) * 100
                    If i = 1 Then intPara(i * 4 + j) = CSng(txtAz1IN(j).text) * 100
                    If i = 2 Then intPara(i * 4 + j) = CSng(txtAz1DN(j).text) * 100
                    If i = 3 Then intPara(i * 4 + j) = CSng(txtAz1RT(j).text) * 10000
                    If i = 4 Then intPara(i * 4 + j) = CSng(txtAz1ST(j).text) * 10
'                    If i = 5 Then intPara(i * 5 + j) = CSng(txtAz1Offset(j).Text) * 10
                Next j
            Next i
        End If
        b1 = frmAz1.WriteParas(201, intPara(), True)
        
        If Para.UseAz2 Then
            For i = 0 To 4
'             For i = 0 To 5
                For j = 0 To 3
                    If i = 0 Then intPara(i * 4 + j) = CSng(txtAz2PN(j).text) * 100
                    If i = 1 Then intPara(i * 4 + j) = CSng(txtAz2IN(j).text) * 100
                    If i = 2 Then intPara(i * 4 + j) = CSng(txtAz2DN(j).text) * 100
                    If i = 3 Then intPara(i * 4 + j) = CSng(txtAz2RT(j).text) * 10000
                    If i = 4 Then intPara(i * 4 + j) = CSng(txtAz2ST(j).text) * 10
'                    If i = 5 Then intPara(i * 5 + j) = CSng(txtAz2Offset(j).Text) * 10
                Next j
            Next i
        End If
        b2 = frmAz2.WriteParas(201, intPara(), True)
        
        If b1 And b2 Then
            MsgBox "寫入成功!", vbOK
        Else
            MsgBox "寫入失敗!", vbOK
        End If
    End If
End Sub
Public Sub cmdRecipeOpen_Click()
'    If Kernel.IsRun = 1 Then
'    Call frmHistory.AppendLogAlert(1, "Manual", 1202, "運行過程中保存Recipe", 1)
'    MsgBox "程序運行中,無法操作Recipe！！"
'    Exit Sub
'    End If
    Dim StrFileName As String
    Dim strFilePath As String
    Dim strDir As String
    Dim lngRet                As Long
    On Error GoTo ERRHNADLE
    
    
    Call frmConfiguration.StopWatchDog
    
    strFilePath = gbSystemPath & "\Recipe" & "\ad"
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    
    strFilePath = gbSystemPath & "\Recipe" & "\eg"
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    
    strFilePath = gbSystemPath & "\Recipe" & "\op"
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    
    If gbblnPNLoad Then
        cdFile.fileName = gbstrPNRecipeFile

    Else
        'If gbintLoginRight <= 2 Then strFilePath = gbSystemPath & "\Recipe" & "\ad"
        'If gbintLoginRight = 3 Then strFilePath = gbSystemPath & "\Recipe" & "\eg"
        If gbstrRecipeFilePath <> "" Then
           'cdFile.InitDir = gbSystemPath & "\Recipe"
            cdFile.InitDir = gbstrRecipeFilePath
            
            cdFile.Filter = "*.rcp|*.rcp"
            cdFile.FilterIndex = 1
            'cdFile.CancelError = True
            gbblnNoModalForm = True
            cdFile.ShowOpen
                        
        End If
    End If
    gbblnNoModalForm = False
    If cdFile.fileName <> "" Then
'        strFileName = gbSystemPath & "\System\system.cfg"
'        lngRet = WritePrivateProfileString("Utility", "LastLoadRecipe", cdFile.FileName, strFileName)
        StrFileName = cdFile.fileName
        If RecipeOpen(StrFileName) = True Then
        lbRecipeName.Caption = Mid(StrFileName, InStrRev(StrFileName, "\") + 1)
        If GbTcoffset_Switch = 1 Then
        ReDim TempOffset(6)
        TempOffset = frmTcOffset.GetTcOffset(StrFileName)
        End If
        End If
    End If
     If Kernel.IsRun = 1 Then
         Call frmHistory.AppendLogAlert(1, "Manual", 1202, "運行過程中打開Recipe:" + lbRecipeName.Caption, 1)
        Else
        If frmPlotProcess.PlotRecipeChart Then
            frmConfiguration.StartWatchDog
            frmPlotProcess.SetPIDValue
            Kernel.strCurrRecipe = lbRecipeName.Caption
            tabRecipe.Tab = 0
            gbblnSendRecipe = True
        End If
    End If
   Exit Sub
ERRHNADLE:
    frmConfiguration.StartWatchDog
End Sub

'========================================================================================================
Private Sub cmdRecipePlot_Click()
    frmPlotProcess.PlotRecipeChart
End Sub

'========================================================================================================
Public Sub cmdRecipeSave_Click()
'    If Kernel.IsRun = 1 Then
'    Call frmHistory.AppendLogAlert(1, "Manual", 1201, "運行過程中保存Recipe", 1)
'    MsgBox "程序運行中,無法操作Recipe！！"
'    Exit Sub
'    End If
    
    Dim i As Integer
    Dim j As Integer
    Dim iRet As Integer
    
'    Dim b1, b2 As Boolean
'    Dim ProcData(9) As Integer
'    Dim intAzbilPara(0 To 15) As Integer
'    Dim iTemp, iTime As Integer
'    Dim iStep As Integer
'    Dim iAdd As Long
    
    If CheckAllRecipeValue = True Then
    
        blnSave = False
        Call frmConfiguration.StopWatchDog
        Call RecipeSave
        cdFile.FLAGS = cdlOFNOverwritePrompt
        If Kernel.IsRun = 0 Then
         If frmPlotProcess.PlotRecipeChart Then
                
            frmConfiguration.StartWatchDog
            frmPlotProcess.SetPIDValue
            Kernel.strCurrRecipe = lbRecipeName.Caption
        End If
        End If
    End If
'
    If fileName <> "" Then
        If RecipeOpen(fileName) = True Then lbRecipeName.Caption = Mid(fileName, InStrRev(fileName, "\") + 1)
         If Kernel.IsRun = 0 Then
            If frmPlotProcess.PlotRecipeChart Then
            
                frmConfiguration.StartWatchDog
                frmPlotProcess.SetPIDValue
                Kernel.strCurrRecipe = lbRecipeName.Caption
                tabRecipe.Tab = 0
            
                gbblnSendRecipe = True
            
            End If
        End If
    End If

'    If Para.UseAz1 Or Para.UseAz2 Then
'        For i = 1 To GB_MAX_STEP_PROCESS
'            If gbProcessRecipeStep(i).strAction = GB_ACTION_STOP Then
'                iStep = i - 1
'                Exit For
'            End If
'        Next i
'        iAdd = 48100
'        For i = 1 To GB_MAX_STEP_PROCESS
'            If gbProcessRecipeStep(i).strAction = GB_ACTION_IDLE Then
'                iTemp = 0
'                iTime = gbProcessRecipeStep(i).sngTime
'            ElseIf gbProcessRecipeStep(i).strAction = GB_ACTION_RAMPUP Or gbProcessRecipeStep(i).strAction = GB_ACTION_HOLD Then
'                iTemp = gbProcessRecipeStep(i).sngTemperature
'                iTime = gbProcessRecipeStep(i).sngTime
'            Else
'                Exit For
'            End If
'
'            ProcData(0) = iTemp
'            ProcData(1) = iTime
'            ProcData(2) = 1
'            ProcData(3) = 0
'            ProcData(4) = 1
'            For j = 0 To 4
'                If Az1.blnUseAzbil = True Then Call frmAz1.WritePara(iAdd + j, ProcData(j))
'                If Az2.blnUseAzbil = True Then Call frmAz2.WritePara(iAdd + j, ProcData(j))
'            Next j
'            iAdd = iAdd + 10
'        Next i
'        ProcData(0) = 0
'        ProcData(1) = 0
'        For j = 0 To 4
'            If Az1.blnUseAzbil = True Then Call frmAz1.WritePara(iAdd + j, ProcData(j))
'            If Az2.blnUseAzbil = True Then Call frmAz2.WritePara(iAdd + j, ProcData(j))
'        Next j
'
'        If Az1.blnUseAzbil = True Then
'            For i = 0 To 3
'                For j = 0 To 3
'                    If i = 0 Then intAzbilPara(i * 4 + j) = CSng(txtAz1PN(j).Text) * 100
'                    If i = 1 Then intAzbilPara(i * 4 + j) = CSng(txtAz1IN(j).Text) * 100
'                    If i = 2 Then intAzbilPara(i * 4 + j) = CSng(txtAz1DN(j).Text) * 100
'                    If i = 3 Then intAzbilPara(i * 4 + j) = CSng(txtAz1RT(j).Text) * 10000
'                Next j
'            Next i
'            b1 = frmAz1.WriteParas(201, intAzbilPara(), True)
'        End If
'        If Az2.blnUseAzbil = True Then
'            For i = 0 To 3
'                For j = 0 To 3
'                    If i = 0 Then intAzbilPara(i * 4 + j) = CSng(txtAz2PN(j).Text) * 100
'                    If i = 1 Then intAzbilPara(i * 4 + j) = CSng(txtAz2IN(j).Text) * 100
'                    If i = 2 Then intAzbilPara(i * 4 + j) = CSng(txtAz2DN(j).Text) * 100
'                    If i = 3 Then intAzbilPara(i * 4 + j) = CSng(txtAz2RT(j).Text) * 10000
'                Next j
'            Next i
'            b2 = frmAz2.WriteParas(201, intAzbilPara(), True)
'        End If
        
        
'        If Az1.blnUseAzbil = True Then
'            iRet = MsgBox("寫入控制器 ?", vbOKCancel)
'            If iRet = vbOK Then
'                If Para.UseAz1 Then
'                    For i = 0 To 3
'                        For j = 0 To 3
'                            If i = 0 Then intAzbilPara(i * 4 + j) = CSng(txtAz1PN(j).Text) * 100
'                            If i = 1 Then intAzbilPara(i * 4 + j) = CSng(txtAz1IN(j).Text) * 100
'                            If i = 2 Then intAzbilPara(i * 4 + j) = CSng(txtAz1DN(j).Text) * 100
'                            If i = 3 Then intAzbilPara(i * 4 + j) = CSng(txtAz1RT(j).Text) * 10000
'                        Next j
'                    Next i
'                End If
'                b1 = frmAz1.WriteParas(201, intAzbilPara(), True)
'
'                If Para.UseAz2 Then
'                    For i = 0 To 3
'                        For j = 0 To 3
'                            If i = 0 Then intAzbilPara(i * 4 + j) = CSng(txtAz2PN(j).Text) * 100
'                            If i = 1 Then intAzbilPara(i * 4 + j) = CSng(txtAz2IN(j).Text) * 100
'                            If i = 2 Then intAzbilPara(i * 4 + j) = CSng(txtAz2DN(j).Text) * 100
'                            If i = 3 Then intAzbilPara(i * 4 + j) = CSng(txtAz2RT(j).Text) * 10000
'                        Next j
'                    Next i
'                End If
'                b2 = frmAz2.WriteParas(201, intAzbilPara(), True)
'
'                If b1 And b2 Then
'                    MsgBox "寫入成功!", vbOK
'                Else
'                    MsgBox "寫入失敗!", vbOK
'                End If
'            End If
'        End If
End Sub




Private Sub cmdWriteOther_Click()
    Dim i As Integer
    Dim j As Integer
    Dim lngRet As Long
    Dim ss() As String
    Dim S As String
    Dim sPath As String
On Error GoTo ERR_SAVE
    
    
    gbblnNoModalForm = True
    cdFile.CancelError = True
    cdFile.InitDir = gbSystemPath & "\Recipe\op"
    cdFile.Filter = "Recipe File(*.rcp)|*.rcp"
    cdFile.FLAGS = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    cdFile.ShowSave
    
    gbblnNoModalForm = False
    If cdFile.fileName <> "" Then
        ss = Split(cdFile.fileName, Chr(0))
        If UBound(ss) > 0 Then
            sPath = ss(0)
            For i = 1 To UBound(ss)
                S = sPath & "\" & ss(i)
                For j = 0 To 3
                    lngRet = WritePrivateProfileString("Azbil", "Az1RT" & CStr(j), txtAz1RT(j).text, S)
                    lngRet = WritePrivateProfileString("Azbil", "Az2RT" & CStr(j), txtAz2RT(j).text, S)
                Next j
            Next i
        End If
    End If
    'strFileName = cdFile.FileName
    
    
    
    
    Exit Sub
    
ERR_SAVE:
    'Call AlertShow("多檔寫入失敗!", ERRORTYPE)
    
End Sub

'========================================================================================================
Private Sub Form_Activate()
    Dim i As Integer
        
'    If gbintNumOfBanks = 6 Then
'        For i = 5 To 5
'            txtIntensityWeightS(i).Enabled = True
'            txtIntensityWeight(i).Enabled = True
'        Next i
'    End If
'    If gbintNumOfBanks = 10 Then
'        For i = 5 To 9
'            txtIntensityWeightS(i).Enabled = True
'            txtIntensityWeight(i).Enabled = True
'        Next i
'    End If
'    If gbintNumOfBanks = 12 Then
'        For i = 5 To 11
'            txtIntensityWeightS(i).Enabled = True
'            txtIntensityWeight(i).Enabled = True
'        Next i
'    End If
'    If gbintNumOfBanks = 14 Then
'        For i = 5 To 13
'            txtIntensityWeightS(i).Enabled = True
'            txtIntensityWeight(i).Enabled = True
'        Next i
'    End If
'    If gbintNumOfBanks = 17 Then
'        For i = 5 To 16
'            txtIntensityWeightS(i).Enabled = True
'            txtIntensityWeight(i).Enabled = True
'        Next i
'    End If
    Dim upperLimit As Integer
    upperLimit = IIf(gbintNumOfBanks > 17, 17, gbintNumOfBanks)
    For i = 0 To upperLimit - 1
        txtIntensityWeightS(i).Enabled = True
        txtIntensityWeight(i).Enabled = True
    Next i
    
    If gbintActiveModule_Door = 1 Then
        fraDoor.Visible = True
    Else
        fraDoor.Visible = False
    End If
'    If gbintActiveModule_Vacuum = 1 Then
'        fraPressure.Visible = True
'    Else
'        fraPressure.Visible = False
'    End If

'    If gbintActiveModule_APC = 1 Then
'        fraAPC.Visible = True
'    Else
'        fraAPC.Visible = False
'    End If
    
      
    tabSCR.Visible = IIf(gbintLoginRight = 1, True, False)
    fraCTCheck.Visible = IIf(Para.UseCT = 1, True, False)
        
    tabRecipe.TabVisible(2) = gbintActiveModule_MLoop
    tabRecipe.TabVisible(3) = IIf(Para.RtaType = 9, True, False)
    
    tabRecipe.TabIndex = 0
    
End Sub

'========================================================================================================
Private Sub Form_Deactivate()
'    Dim iRet As Integer
'    If blnSave = True Then
'        iRet = MsgBox("Do you want to save file", vbOKCancel)
'        If iRet = vbOK Then
'            Call cmdRecipeSave_Click
'        End If
'        blnSave = False
'    End If
End Sub



'========================================================================================================
Private Sub Form_Load()
    Call Me.InitForm
    ReDim gbProcessRecipeStep(0)
    If DefineprocStep = 1 Then
    CmdBtn_BuildProcessStrp.Visible = True
    CmdBtn_BuildProcessStrp.Enabled = True
   Else
    CmdBtn_BuildProcessStrp.Visible = False
    CmdBtn_BuildProcessStrp.Enabled = False
   End If
'   If Kernel.IsRun = 1 Then
'   cmdRecipeOpen.Enabled = False
'   cmdRecipeSave.Enabled = False
'   Else
'   cmdRecipeOpen.Enabled = True
'   cmdRecipeSave.Enabled = True
'   End If
End Sub

Public Sub ReFreshAction()
  Dim Config_Step As Collection
    Dim k As Integer
    cmbRecipeAction.Clear
    Set Config_Step = ReadTextFileToArray(App.Path + "\Config\ProcessStep.txt")
    If Config_Step.Count > 0 Then
    TxtExist = True
    With cmbRecipeAction
        For k = 1 To Config_Step.Count
        .AddItem Split(Config_Step(k), ";")(0)
        Next k
    End With
    End If
    
End Sub



Private Function GetGasCount(ByRef GasNamelist As Collection, ByRef GasUnitlist As Collection, ByRef GasMaxSlmplist As Collection) As Integer
Dim GasCount As Integer
Dim i As Integer
Dim j As Integer
Dim StrFileName As String
Dim ConfigName As String
StrFileName = gbSystemPath & "\System\system.cfg"
ConfigName = gbSystemPath & ProcDict_Path
Set GasNamelist = New Collection
Set GasUnitlist = New Collection
Set GasMaxSlmplist = New Collection
For j = 1 To 6
If CommnonReadini("PARAMETER", "Gas" + CStr(j) + "Active", StrFileName) = "1" And CommnonReadini("PARAMETER", "Gas" + CStr(j) + "Alias", StrFileName) <> "NA" Then
GasNamelist.Add CommnonReadini("PARAMETER", "Gas" + CStr(j) + "Alias", StrFileName)
GasUnitlist.Add CommnonReadini("PARAMETER", "Gas" + CStr(j) + "Unit", StrFileName)
GasMaxSlmplist.Add CommnonReadini("PARAMETER", "Gas" + CStr(j) + "SLMP", StrFileName)
GasCount = GasCount + 1
End If
Next j
If CommnonReadini("Gas7", "Gas7Active", ConfigName) = "1" And CommnonReadini("Gas7", "Gas7Alias", ConfigName) <> "NA" Then
GasNamelist.Add CommnonReadini("Gas7", "Gas7Alias", ConfigName)
GasUnitlist.Add CommnonReadini("Gas7", "Gas7Unit", ConfigName)
GasMaxSlmplist.Add CommnonReadini("Gas7", "Gas7SLMP", ConfigName)
GasCount = GasCount + 1
End If
GetGasCount = GasCount
End Function






'========================================================================================================
Public Sub InitForm()
    Dim i As Integer, j As Integer
    Dim sngTotalGridWidth As Single
    Dim S As String
    Dim s1 As String
    Dim GasCount As Integer
    
    
    intRecipeTempInputType = 1
    blnSave = False
    fraRecipe.Top = 100
    fraRecipe.Left = 200
    fraPID.Left = fraRecipe.Left
    GasCount = GetGasCount(GasNames, GasUnits, GasMaxSlmps)
    'Initial recipe grid layout.
    With hfgRecipe
        '.Top = 500
        .Left = 200
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 51
        .Cols = 4 + GasCount
        .ColWidth(0) = 800
        .ColWidth(1) = 1500
        If GasCount < 7 Then
        For i = 2 To .Cols - 1
            .ColWidth(i) = 2000
        Next i
        Else
        For i = 2 To .Cols - 1
            .ColWidth(i) = 1700
        Next i
        End If
        For i = 0 To .Cols - 1
            sngTotalGridWidth = sngTotalGridWidth + .ColWidth(i)
            .ColAlignmentFixed = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignCenterCenter
        Next i
        .Width = sngTotalGridWidth + 350
        .TextMatrix(0, GB_PROCESS_STEP) = "Step"
        .TextMatrix(0, GB_PROCESS_ACTION) = "Action"
        .TextMatrix(0, GB_PROCESS_TEMP) = "T(℃)/ P(%)"
        .TextMatrix(0, GB_PROCESS_TIME) = "Time (Sec)"
        For j = 1 To GasNames.Count 'gbintMaxGasEnable
            If GasNames(j) <> "NA" Then
            S = GasNames(j) & "(" & GasUnits(j) & "~" & CStr(GasMaxSlmps(j)) & ")"
            .TextMatrix(0, GB_PROCESS_GAS1 + j - 1) = S
            End If
        Next j

        .RowHeight(0) = cmbRecipeAction.Height
        For i = 1 To .Rows - 1
            .RowHeight(i) = cmbRecipeAction.Height
            .TextMatrix(i, GB_PROCESS_STEP) = str(i)
            .TextMatrix(i, GB_PROCESS_ACTION) = GB_ACTION_STOP
            .TextMatrix(i, GB_PROCESS_TEMP) = "0"
            .TextMatrix(i, GB_PROCESS_TIME) = "0"
            For j = 1 To GasNames.Count 'gbintMaxGasEnable
                .TextMatrix(i, GB_PROCESS_GAS1 + j - 1) = "0"
            Next j
        Next i
        .Refresh
        .AllowUserResizing = flexResizeNone
    End With
    Dim Config_Step As Collection
    Dim k As Integer
    cmbRecipeAction.Clear
    Set Config_Step = ReadTextFileToArray(App.Path + "\Config\ProcessStep.txt")
    If DefineprocStep = 1 Then
    If Config_Step.Count > 0 Then
    TxtExist = True
    With cmbRecipeAction
        For k = 1 To Config_Step.Count
        .AddItem Split(Config_Step(k), ";")(0)
        Next k
    End With
    End If
    Else
     With cmbRecipeAction
        .AddItem GB_ACTION_IDLE
        '.AddItem GB_ACTION_PREHEAT
        .AddItem GB_ACTION_RAMPUP
        .AddItem GB_ACTION_HOLD
        .AddItem GB_ACTION_STOP
        '.AddItem GB_ACTION_PURGE
'        .AddItem GB_ACTION_RAMPDOWN
'        .AddItem GB_ACTION_IOCONTROL
'        If gbintActiveModule_Vacuum = 1 Then
'            .AddItem GB_ACTION_VENT
'            .AddItem GB_ACTION_PUMPDOWN
'            .AddItem GB_ACTION_PUMPDOWNKEEP
'        End If
        '
        '.AddItem GB_ACTION_COOLING
        '.AddItem GB_ACTION_RAMPDOWN
        '.AddItem GB_ACTION_MANUALPUMP
    End With
    
    End If
'    If Config_Step.Count > 0 Then
'    TxtExist = True
'    With cmbRecipeAction
'        For K = 1 To Config_Step.Count
'        .AddItem Split(Config_Step(K), ";")(0)
'        Next K
'    End With
'    Else
'      With cmbRecipeAction
'        .AddItem GB_ACTION_IDLE
'        '.AddItem GB_ACTION_PREHEAT
'        .AddItem GB_ACTION_RAMPUP
'        .AddItem GB_ACTION_HOLD
'        .AddItem GB_ACTION_STOP
'        '.AddItem GB_ACTION_PURGE
'        .AddItem GB_ACTION_RAMPDOWN
'        .AddItem GB_ACTION_IOCONTROL
''        If gbintActiveModule_Vacuum = 1 Then
''            .AddItem GB_ACTION_VENT
''            .AddItem GB_ACTION_PUMPDOWN
''            .AddItem GB_ACTION_PUMPDOWNKEEP
''        End If
'        '
'        '.AddItem GB_ACTION_COOLING
'        '.AddItem GB_ACTION_RAMPDOWN
'        '.AddItem GB_ACTION_MANUALPUMP
'    End With
'    End If
    
  

    
    fraAz1.Visible = Para.UseAz1
    fraAz2.Visible = Para.UseAz2
    
    'tabRecipe.TabVisible(1) = False
End Sub


'========================================================================================================
Private Sub hfgRecipe_Click()
    If hfgRecipe.ColSel = 1 Then
        With cmbRecipeAction
            .text = hfgRecipe.TextMatrix(hfgRecipe.RowSel, GB_PROCESS_ACTION)
            strOriginActionName = hfgRecipe.TextMatrix(hfgRecipe.RowSel, GB_PROCESS_ACTION)
            .Move fraRecipe.Left + tabRecipe.Left + hfgRecipe.Left + hfgRecipe.ColPos(hfgRecipe.ColSel) + 20, _
                      fraRecipe.Top + tabRecipe.Top + hfgRecipe.Top + hfgRecipe.RowPos(hfgRecipe.RowSel) + 25, _
                  hfgRecipe.ColWidth(hfgRecipe.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    Else
        With txtRecipeEdit
            .text = hfgRecipe.TextMatrix(hfgRecipe.RowSel, hfgRecipe.ColSel)
            .Move fraRecipe.Left + tabRecipe.Left + hfgRecipe.Left + hfgRecipe.ColPos(hfgRecipe.ColSel) + 20, _
                  fraRecipe.Top + tabRecipe.Top + hfgRecipe.Top + hfgRecipe.RowPos(hfgRecipe.RowSel) + 25, _
                  hfgRecipe.ColWidth(hfgRecipe.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
    intCurrRowSel = hfgRecipe.RowSel
    intCurrColSel = hfgRecipe.ColSel
End Sub





'========================================================================================================
Private Sub hfgRecipe_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        txtRecipeEdit.Visible = False
    End If
End Sub






Private Sub tabRecipe_Click(PreviousTab As Integer)
If tabRecipe.Tab = 0 Then
hfgRecipe.Enabled = True
hfgRecipe.Visible = True
Else
hfgRecipe.Enabled = False
hfgRecipe.Visible = False
End If
End Sub



Private Sub txtAz1Offset_KeyPress(Index As Integer, KeyAscii As Integer)
     Dim TemptureColl As Collection
     Dim MAXTempture As Integer
      Dim Offset_txt As String
      Dim Curent_R As Double
    If KeyAscii = 13 Then
        Offset_txt = txtAz1Offset(Index).text
      If Offset_txt <> "" And IsNumeric(Offset_txt) Then
        If Val(Offset_txt) > 50 Or Val(Offset_txt) < -50 Then
            MsgBox ("Offset值需要在-50至50之間!")
            txtAz1Offset(Index).text = "0"
        Else
                Set TemptureColl = GetTemptureColl()
                If TemptureColl.Count = 0 Then
                    MsgBox "Process Step未輸入溫度值！"
                Else
                   MAXTempture = GetMAXTempture(TemptureColl)
                   Curent_R = Val(txtAz1RT(Index).text) + -Val(Offset_txt) * (1 / MAXTempture)
                   Curent_R = Format(Curent_R, "0.0000")
                   txtAz1RT(Index).text = CStr(Curent_R)
               End If
        End If
     End If


     End If

'
  
End Sub





Private Sub txtAz2Offset_KeyPress(Index As Integer, KeyAscii As Integer)
       Dim TemptureColl As Collection
     Dim MAXTempture As Integer
      Dim Offset_txt As String
      Dim Curent_R As Double
      If KeyAscii = 13 Then
        Offset_txt = txtAz2Offset(Index).text
      If Offset_txt <> "" And IsNumeric(Offset_txt) Then
        If Val(Offset_txt) > 50 Or Val(Offset_txt) < -50 Then
            MsgBox ("Offset值需要在-50至50之間!")
            txtAz2Offset(Index).text = "0"
        Else
                Set TemptureColl = GetTemptureColl()
                If TemptureColl.Count = 0 Then
                    MsgBox "Process Step未輸入溫度值！"
                Else
                   MAXTempture = GetMAXTempture(TemptureColl)
                   Curent_R = Val(txtAz2RT(Index).text) + -Val(Offset_txt) * (1 / MAXTempture)
                   Curent_R = Format(Curent_R, "0.0000")
                   txtAz2RT(Index).text = CStr(Curent_R)
               End If
        End If
     End If
     End If
End Sub

'========================================================================================================
Private Sub txtPressureControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtPressureControl.text) > 2 Or Val(txtPressureControl.text) < 0.15 Then
            txtPressureControl.text = CStr(2)
        End If
    End If
End Sub

'========================================================================================================
'Private Sub txtRecipeEdit_Change()
'        'hfgRecipe.TextMatrix(hfgRecipe.RowSel, hfgRecipe.ColSel) = txtRecipeEdit.Text
'End Sub

'========================================================================================================
Private Sub txtRecipeEdit_Click()
'    txtRecipeEdit.SelStart = 0
'    txtRecipeEdit.SelLength = Len(txtRecipeEdit)
'    blnSave = True
End Sub

'========================================================================================================
Private Sub txtRecipeEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If hfgRecipe.ColSel = GB_PROCESS_TEMP Then
            If gbsngMaxTemperature < Val(txtRecipeEdit.text) Then
                txtRecipeEdit.text = CStr(gbsngMaxTemperature)
            
            End If
        End If
        hfgRecipe.TextMatrix(hfgRecipe.RowSel, hfgRecipe.ColSel) = txtRecipeEdit.text
        txtRecipeEdit.text = "0"
        txtRecipeEdit.Visible = False
        blnSave = True
    End If
End Sub

'========================================================================================================
Private Sub txtRecipeEdit_LostFocus()
    
    Call CheckSetValue(hfgRecipe.RowSel)
    txtRecipeEdit.text = "0"
    txtRecipeEdit.Visible = False
End Sub

'========================================================================================================
Private Function RecipeSave() As Boolean
    Dim i               As Integer
    Dim j               As Integer
    Dim lngRet          As Long
    Dim StrFileName     As String
    Dim iInputDevice    As Integer
    Dim iInputObject    As Integer
    Dim strCSV As String
    Dim GasName As String
    

    
    On Error GoTo ERR_RECIPE_SAVE
    
    gbblnNoModalForm = True
    cdFile.CancelError = True
    cdFile.InitDir = gbSystemPath & "\Recipe"
    cdFile.Filter = "Recipe File(*.rcp)|*.rcp|CSV File(*.csv)|*.csv"
    cdFile.ShowSave
    
    
    
    gbblnNoModalForm = False
    If cdFile.fileName = "" Then RecipeSave = False
    StrFileName = cdFile.fileName
    
    sngRecipeTotalSec = 0
    lbRecipeName.Caption = Mid(StrFileName, InStrRev(StrFileName, "\") + 1)
    fileName = StrFileName
    If Kernel.IsRun = 1 And lbRecipeName.Caption = Kernel.strCurrRecipe Then
      MsgBox "該Recipe在運行中,無法保存!!!!!"
    Exit Function
    End If
    For i = 1 To GB_MAX_STEP_PROCESS
        'Action

        lngRet = WritePrivateProfileString("STEP" & str(i), "ACTION", hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION), StrFileName)
    

        'Temperature (degree)
        lngRet = WritePrivateProfileString("STEP" & str(i), "TEMP", hfgRecipe.TextMatrix(i, GB_PROCESS_TEMP), StrFileName)
        'Time (sec)
        lngRet = WritePrivateProfileString("STEP" & str(i), "TIME", hfgRecipe.TextMatrix(i, GB_PROCESS_TIME), StrFileName)
        If hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) <> "STOP" Then
            sngRecipeTotalSec = sngRecipeTotalSec + Val(hfgRecipe.TextMatrix(i, GB_PROCESS_TIME))
        End If
        For j = 1 To GasNames.Count
            GasName = GasNames(j)
            lngRet = WritePrivateProfileString("STEP" & str(i), GasName, hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j - 1), StrFileName)
        Next j
   Next i
    
    If txtProportional.text = "" Then txtProportional.text = "0"
    If txtProportional2.text = "" Then txtProportional2.text = "0"
    If txtIntegrnal.text = "" Then txtIntegrnal.text = "0"
    If txtIntegral2.text = "" Then txtIntegral2.text = "0"
    If txtDerivational.text = "" Then txtDerivational.text = "0"
    If txtPredit.text = "" Then txtPredit.text = "0"
    If txtFeedForward.text = "" Then txtFeedForward.text = "0"
    If optInputDevice(0).value = True Then iInputDevice = 1
    If optInputDevice(1).value = True Then iInputDevice = 2
    If optObject(0).value = True Then iInputObject = 0
    If optObject(1).value = True Then iInputObject = 1
    
    If txtOvershoot.text = "" Then txtOvershoot.text = "0"
    If txtUndershoot.text = "" Then txtUndershoot.text = "0"
    If txtOverPressure.text = "" Then txtOverPressure.text = "760"
    If txtIntLimit.text = "" Then txtIntLimit.text = "50"
    
    
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Proportional", CStr(Val(txtProportional.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Proportional2", CStr(Val(txtProportional2.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Integrnal", CStr(Val(txtIntegrnal.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Integral2", CStr(Val(txtIntegral2.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Derivational", CStr(Val(txtDerivational.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Factor1", CStr(Val(txtPredit.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Factor2", CStr(Val(txtFeedForward.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "InputDevice", CStr(Val(iInputDevice)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Overshoot", CStr(Val(txtOvershoot.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Undershoot", CStr(Val(txtUndershoot.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "OverPressure", CStr(Val(txtOverPressure.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "InputObject", CStr(Val(iInputObject)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "IntLimit", CStr(Val(txtIntLimit.text)), StrFileName)
    
       
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "PressureControl", CStr(Val(txtPressureControl.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "PP", CStr(Val(txtPP.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "II", CStr(Val(txtII.text)), StrFileName)
    
    '120822 Josh
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "RampDownPower", CStr(Val(txtRampDownPower.text)), StrFileName)
    sngRecipeRampDownPower = CSng(Val(txtRampDownPower.text)) / 10
    If sngRecipeRampDownPower > 5 Then
        sngRecipeRampDownPower = 5
        txtRampDownPower.text = "50"
    End If
    gbsngRecipeRampDownPower = sngRecipeRampDownPower
    
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "APC_P", CStr(Val(txtAPC_P.text)), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "APC_I", CStr(Val(txtAPC_I.text)), StrFileName)

    'Rev10.0.0.5 Add the intensity weight in recipe edot
    For i = 0 To 15
        If Val(txtIntensityWeightS(i).text) > 100 Then txtIntensityWeightS(i).text = "100"
        If Val(txtIntensityWeight(i).text) > 100 Then txtIntensityWeight(i).text = "100"
    Next i
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS1", CStr(Val(txtIntensityWeightS(0).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS2", CStr(Val(txtIntensityWeightS(1).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS3", CStr(Val(txtIntensityWeightS(2).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS4", CStr(Val(txtIntensityWeightS(3).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS5", CStr(Val(txtIntensityWeightS(4).text)), StrFileName)
    '120713 Josh
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS6", CStr(Val(txtIntensityWeightS(5).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS7", CStr(Val(txtIntensityWeightS(6).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS8", CStr(Val(txtIntensityWeightS(7).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS9", CStr(Val(txtIntensityWeightS(8).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS10", CStr(Val(txtIntensityWeightS(9).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS11", CStr(Val(txtIntensityWeightS(10).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS12", CStr(Val(txtIntensityWeightS(11).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS13", CStr(Val(txtIntensityWeightS(12).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS14", CStr(Val(txtIntensityWeightS(13).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS15", CStr(Val(txtIntensityWeightS(14).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS16", CStr(Val(txtIntensityWeightS(15).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWS17", CStr(Val(txtIntensityWeightS(16).text)), StrFileName)
    
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD1", CStr(Val(txtIntensityWeight(0).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD2", CStr(Val(txtIntensityWeight(1).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD3", CStr(Val(txtIntensityWeight(2).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD4", CStr(Val(txtIntensityWeight(3).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD5", CStr(Val(txtIntensityWeight(4).text)), StrFileName)
    '120713 Josh
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD6", CStr(Val(txtIntensityWeight(5).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD7", CStr(Val(txtIntensityWeight(6).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD8", CStr(Val(txtIntensityWeight(7).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD9", CStr(Val(txtIntensityWeight(8).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD10", CStr(Val(txtIntensityWeight(9).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD11", CStr(Val(txtIntensityWeight(10).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD12", CStr(Val(txtIntensityWeight(11).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD13", CStr(Val(txtIntensityWeight(12).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD14", CStr(Val(txtIntensityWeight(13).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD15", CStr(Val(txtIntensityWeight(14).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD16", CStr(Val(txtIntensityWeight(15).text)), StrFileName)
    lngRet = WritePrivateProfileString("POWER_WEIGHT", "PWD17", CStr(Val(txtIntensityWeight(16).text)), StrFileName)
    
    lngRet = WritePrivateProfileString("CT", "CTGate1", txtCT(0).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate2", txtCT(1).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate3", txtCT(2).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate4", txtCT(3).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate5", txtCT(4).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate6", txtCT(5).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate7", txtCT(6).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate8", txtCT(7).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate9", txtCT(8).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate10", txtCT(9).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate11", txtCT(10).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate12", txtCT(11).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate13", txtCT(12).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate14", txtCT(13).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate15", txtCT(14).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate16", txtCT(15).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CTGate17", txtCT(16).text, StrFileName)
    
    
    lngRet = WritePrivateProfileString("CT", "CDGate1", txtCD(0).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate2", txtCD(1).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate3", txtCD(2).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate4", txtCD(3).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate5", txtCD(4).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate6", txtCD(5).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate7", txtCD(6).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate8", txtCD(7).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate9", txtCD(8).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate10", txtCD(9).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate11", txtCD(10).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate12", txtCD(11).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate13", txtCD(12).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate14", txtCD(13).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate15", txtCD(14).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate16", txtCD(15).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "CDGate17", txtCD(16).text, StrFileName)
    lngRet = WritePrivateProfileString("CT", "UseCT", CStr(chkUseCT.value), StrFileName)
    lngRet = WritePrivateProfileString("CT", "SaveLogCT", CStr(chkSaveLogCT.value), StrFileName)
    
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "SmoothTime", txtSmoothTime.text, StrFileName) 'S
    gbsngSmoothTime = CSng(Val(txtSmoothTime.text))
    gbblnRecipeUseCT = chkUseCT.value
    gbblnRecipeSaveLogCT = chkSaveLogCT.value
    sngRecipeProportional = CSng(Val(txtProportional.text))
    sngRecipeProportional2 = CSng(Val(txtProportional2.text))
    sngRecipeIntegral = CSng(Val(txtIntegrnal.text))
    sngRecipeIntegral2 = CSng(Val(txtIntegral2.text))
    sngRecipeDerivational = CSng(Val(txtDerivational.text))
    sngRecipePredit = CSng(Val(txtPredit.text))
    sngRecipeFeedForward = CSng(Val(txtFeedForward.text))
    sngRecipeOverTemp = CSng(Val(txtOvershoot.text))
    sngRecipeUnderTemp = CSng(Val(txtUndershoot.text))
    sngRecipeOverPressure = CSng(Val(txtOverPressure.text))
    intRecipeTempInputType = Val(iInputDevice)
    gbintPMDetectObject = Val(iInputObject)
    
    sngRecipePP = CSng(Val(txtPP.text))
    sngRecipeII = CSng(Val(txtII.text))
    sngRecipeIntLimit = CSng(Val(txtIntLimit.text))
    
        
    gbsngAPCGaugePressureValue = Val(txtPressureControl.text)
    
    'Rev10.0.0.5
    '120713 Josh
    For i = 0 To GB_SCR_MAX - 1
        gbsngRecipeIntensityWeightSteady(i) = CSng(Val(Me.txtIntensityWeightS(i).text))
        gbsngRecipeIntensityWeightDynamic(i) = CSng(Val(Me.txtIntensityWeight(i).text))
        gbsngRecipeCT(i) = CSng(Val(Me.txtCT(i).text))
        gbsngRecipeCD(i) = CSng(Val(Me.txtCD(i).text))
    Next i
        
    lngRet = WritePrivateProfileString("Utility", "LoginRight", CStr(cmbRights.ListIndex), StrFileName)
    lngRet = WritePrivateProfileString("Utility", "FinishedClear", CStr(chkFinishedClear.value), StrFileName)
    gbblnRecipeFinishedClear = chkFinishedClear.value
        
    lngRet = WritePrivateProfileString("Motion", "PinHeight", CStr(Val(txtPinHeight.text)), StrFileName)
    gbsngRecipePinHeight = CSng(Val(Me.txtPinHeight.text))
    lngRet = WritePrivateProfileString("Motion", "StartAutoClose", CStr(chkStartAutoClose.value), StrFileName)
    gbblnRecipeStartAutoClose = chkStartAutoClose.value
    lngRet = WritePrivateProfileString("Motion", "EndAutoOpen", CStr(chkEndAutoOpen.value), StrFileName)
    gbblnRecipeEndAutoOpen = chkEndAutoOpen.value
    lngRet = WritePrivateProfileString("Motion", "AutoCloseValve1", CStr(chkAutoCloseValve1.value), StrFileName)
    gbblnRecipeAutoCloseValve1 = chkAutoCloseValve1.value
    lngRet = WritePrivateProfileString("Motion", "AutoCloseValve2", CStr(chkAutoCloseValve2.value), StrFileName)
    gbblnRecipeAutoCloseValve2 = chkAutoCloseValve2.value
    lngRet = WritePrivateProfileString("Motion", "StartCloseCover", CStr(chkStartCloseCover.value), StrFileName)
    gbblnRecipeStartCloseCover = chkStartCloseCover.value
    lngRet = WritePrivateProfileString("Motion", "EndOpenCover", CStr(chkEndOpenCover.value), StrFileName)
    gbblnRecipeEndOpenCover = chkEndOpenCover.value
    
    gbsngAPC_P = Val(txtAPC_P.text)
    gbsngAPC_I = Val(txtAPC_I.text)
    
    lngRet = WritePrivateProfileString("MultiLoop", "UseMultiLoop", CStr(chkUseMultiLoop.value), StrFileName)
    MultiLoop.blnUseMultiLoop = chkUseMultiLoop.value
    For i = 0 To GB_MAX_LOOPS - 1
        lngRet = WritePrivateProfileString("MultiLoop", "UseLoop" & CStr(i), CStr(chkUseLoop(i).value), StrFileName)
        MultiLoop.blnUseLoop(i) = chkUseLoop(i).value
        lngRet = WritePrivateProfileString("MultiLoop", "PN" & CStr(i), txtLoopPN(i).text, StrFileName)
        MultiLoop.sngLoopPN(i) = Val(txtLoopPN(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "IN" & CStr(i), txtLoopIN(i).text, StrFileName)
        MultiLoop.sngLoopIN(i) = Val(txtLoopIN(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "DN" & CStr(i), txtLoopDN(i).text, StrFileName)
        MultiLoop.sngLoopDN(i) = Val(txtLoopDN(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "RT" & CStr(i), txtLoopRT(i).text, StrFileName)
        MultiLoop.sngLoopRT(i) = Val(txtLoopRT(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "FT" & CStr(i), txtLoopFT(i).text, StrFileName)
        MultiLoop.sngLoopFT(i) = Val(txtLoopFT(i).text)
        
        lngRet = WritePrivateProfileString("MultiLoop", "CN" & CStr(i), txtLoopCN(i).text, StrFileName)
        MultiLoop.intLoopCN(i) = Int(txtLoopCN(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "CV" & CStr(i), txtLoopCV(i).text, StrFileName)
        MultiLoop.sngLoopCV(i) = Val(txtLoopCV(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "TC" & CStr(i), txtLoopTC(i).text, StrFileName)
        MultiLoop.intLoopTC(i) = Int(txtLoopTC(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankA" & CStr(i), txtLoopA(i).text, StrFileName)
        MultiLoop.intLoopA(i) = Int(txtLoopA(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankB" & CStr(i), txtLoopB(i).text, StrFileName)
        MultiLoop.intLoopB(i) = Int(txtLoopB(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankC" & CStr(i), txtLoopC(i).text, StrFileName)
        MultiLoop.intLoopC(i) = Int(txtLoopC(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankD" & CStr(i), txtLoopD(i).text, StrFileName)
        MultiLoop.intLoopD(i) = Int(txtLoopD(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankE" & CStr(i), txtLoopE(i).text, StrFileName)
        MultiLoop.intLoopE(i) = Int(txtLoopE(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankF" & CStr(i), txtLoopF(i).text, StrFileName)
        MultiLoop.intLoopF(i) = Int(txtLoopF(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankG" & CStr(i), txtLoopG(i).text, StrFileName)
        MultiLoop.intLoopG(i) = Int(txtLoopG(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankH" & CStr(i), txtLoopH(i).text, StrFileName)
        MultiLoop.intLoopH(i) = Int(txtLoopH(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankJ" & CStr(i), txtLoopJ(i).text, StrFileName)
        MultiLoop.intLoopJ(i) = Int(txtLoopJ(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankK" & CStr(i), txtLoopK(i).text, StrFileName)
        MultiLoop.intLoopK(i) = Int(txtLoopK(i).text)
        
        lngRet = WritePrivateProfileString("MultiLoop", "BankMA" & CStr(i), txtLoopMA(i).text, StrFileName)
        MultiLoop.intLoopMA(i) = Int(txtLoopMA(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMB" & CStr(i), txtLoopMB(i).text, StrFileName)
        MultiLoop.intLoopMB(i) = Int(txtLoopMB(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMC" & CStr(i), txtLoopMC(i).text, StrFileName)
        MultiLoop.intLoopMC(i) = Int(txtLoopMC(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMD" & CStr(i), txtLoopMD(i).text, StrFileName)
        MultiLoop.intLoopMD(i) = Int(txtLoopMD(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankME" & CStr(i), txtLoopME(i).text, StrFileName)
        MultiLoop.intLoopME(i) = Int(txtLoopME(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMF" & CStr(i), txtLoopMF(i).text, StrFileName)
        MultiLoop.intLoopMF(i) = Int(txtLoopMF(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMG" & CStr(i), txtLoopMG(i).text, StrFileName)
        MultiLoop.intLoopMG(i) = Int(txtLoopMG(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMH" & CStr(i), txtLoopMH(i).text, StrFileName)
        MultiLoop.intLoopMH(i) = Int(txtLoopMH(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMJ" & CStr(i), txtLoopMJ(i).text, StrFileName)
        MultiLoop.intLoopMJ(i) = Int(txtLoopMJ(i).text)
        lngRet = WritePrivateProfileString("MultiLoop", "BankMK" & CStr(i), txtLoopMK(i).text, StrFileName)
        MultiLoop.intLoopMK(i) = Int(txtLoopMK(i).text)
    Next i
    
    If Para.UseAz1 Then
        lngRet = WritePrivateProfileString("Azbil", "UseAz1", CStr(chkUseAz1.value), StrFileName)
        Az1.blnUseAzbil = chkUseAz1.value
        lngRet = WritePrivateProfileString("Azbil", "Az1AT", CStr(chkAz1AT.value), StrFileName)
        Az1.blnAutoTuning = chkAz1AT.value
        For i = 0 To 3
            lngRet = WritePrivateProfileString("Azbil", "UseAz1Loop" & CStr(i), CStr(chkUseAz1Loop(i).value), StrFileName)
            Az1.blnUseLoop(i) = chkUseAz1Loop(i).value
            If CSng(txtAz1PN(i).text) = 0 Then txtAz1PN(i).text = "1"
            If CSng(txtAz1IN(i).text) = 0 Then txtAz1IN(i).text = "1"
            If CSng(txtAz1DN(i).text) = 0 Then txtAz1DN(i).text = "1"
            If CSng(txtAz1RT(i).text) = 0 Then txtAz1RT(i).text = "1"
            
            lngRet = WritePrivateProfileString("Azbil", "Az1PN" & CStr(i), txtAz1PN(i).text, StrFileName)
            Az1.sngPN(i) = CSng(txtAz1PN(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az1IN" & CStr(i), txtAz1IN(i).text, StrFileName)
            Az1.sngIN(i) = CSng(txtAz1IN(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az1DN" & CStr(i), txtAz1DN(i).text, StrFileName)
            Az1.sngDN(i) = CSng(txtAz1DN(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az1RT" & CStr(i), txtAz1RT(i).text, StrFileName)
            Az1.sngRT(i) = CSng(txtAz1RT(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az1ST" & CStr(i), txtAz1ST(i).text, StrFileName)
            Az1.sngST(i) = CSng(txtAz1ST(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az1Offset" & CStr(i), txtAz1Offset(i).text, StrFileName)
        Next i
    End If
    If Para.UseAz2 Then
        lngRet = WritePrivateProfileString("Azbil", "UseAz2", CStr(chkUseAz2.value), StrFileName)
        Az2.blnUseAzbil = chkUseAz2.value
        lngRet = WritePrivateProfileString("Azbil", "Az2AT", CStr(chkAz2AT.value), StrFileName)
        Az2.blnAutoTuning = chkAz2AT.value
        For i = 0 To 3
            lngRet = WritePrivateProfileString("Azbil", "UseAz2Loop" & CStr(i), CStr(chkUseAz2Loop(i).value), StrFileName)
            Az2.blnUseLoop(i) = chkUseAz2Loop(i).value
            
            If CSng(txtAz2PN(i).text) = 0 Then txtAz2PN(i).text = "1"
            If CSng(txtAz2IN(i).text) = 0 Then txtAz2IN(i).text = "1"
            If CSng(txtAz2DN(i).text) = 0 Then txtAz2DN(i).text = "1"
            If CSng(txtAz2RT(i).text) = 0 Then txtAz2RT(i).text = "1"
            lngRet = WritePrivateProfileString("Azbil", "Az2PN" & CStr(i), txtAz2PN(i).text, StrFileName)
            Az2.sngPN(i) = CSng(txtAz2PN(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az2IN" & CStr(i), txtAz2IN(i).text, StrFileName)
            Az2.sngIN(i) = CSng(txtAz2IN(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az2DN" & CStr(i), txtAz2DN(i).text, StrFileName)
            Az2.sngDN(i) = CSng(txtAz2DN(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az2RT" & CStr(i), txtAz2RT(i).text, StrFileName)
            Az2.sngRT(i) = CSng(txtAz2RT(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az2ST" & CStr(i), txtAz2ST(i).text, StrFileName)
            Az2.sngST(i) = CSng(txtAz2ST(i).text)
            lngRet = WritePrivateProfileString("Azbil", "Az2Offset" & CStr(i), txtAz2Offset(i).text, StrFileName)
            Az2.sngOffset(i) = CSng(txtAz2Offset(i).text)
        Next i
    End If
    
    
    
    lngRet = WritePrivateProfileString("Utility", "TotalSec", CStr(sngRecipeTotalSec), StrFileName)
    
    gbintRecipePrepareIndex = Val(txtPrepareIndex.text)
    lngRet = WritePrivateProfileString("InterLock", "PrepareIndex", CStr(gbintRecipePrepareIndex), StrFileName)
    gbsngRecipePrepareGaugeO2 = Val(txtPrepareGaugeO2.text)
    lngRet = WritePrivateProfileString("InterLock", "PrepareGaugeO2", CStr(gbsngRecipePrepareGaugeO2), StrFileName)
    gbsngRecipePrepareTimeout = Val(txtPrepareTimeout.text)
    lngRet = WritePrivateProfileString("InterLock", "PrepareTimeout", CStr(gbsngRecipePrepareTimeout), StrFileName)
    
'    gbsngRecipeTempDownTimeout = Val(txtTempDownTimeout.Text)
'    lngRet = WritePrivateProfileString("InterLock", "TempDownTimeout", CStr(gbsngRecipeTempDownTimeout), strFileName)
'    gbsngRecipeGatePS1 = Val(txtGatePS1.Text)
'    lngRet = WritePrivateProfileString("InterLock", "GatePS1", CStr(gbsngRecipeGatePS1), strFileName)
'    gbsngRecipeGatePS2 = Val(txtGatePS2.Text)
'    lngRet = WritePrivateProfileString("InterLock", "GatePS2", CStr(gbsngRecipeGatePS2), strFileName)
    
'    lblRecipeFile.Caption = FileName 'GetFileName(FileName)
    For i = 1 To GB_MAX_STEP_PROCESS
        For j = 1 To GB_MAX_ACTION_TYPE
            'If strData(i, j) <> "" Then
            If hfgRecipe.TextMatrix(i, j) <> "" Then
                m_Recipe.arrayRecipe(i, j) = hfgRecipe.TextMatrix(i, j)
            Else
                m_Recipe.arrayRecipe(i, j) = "0"
            End If
        Next j
    Next i
'    Call RecipeAssignment
    If Kernel.IsRun = 0 Then
     Call RecipeAssignment
     Kernel.strCurrRecipeFile = StrFileName
    End If
    
    Call frmHistory.AppendLogAlert(1, "Manual", 1102, "配方參數儲存", 1)
'   Kernel.strCurrRecipeFile = StrFileName
    RecipeSave = True
    
    
    
    If cdFile.FilterIndex = 2 Then
        Dim pos As Integer
        pos = InStrRev(StrFileName, ".")
        If pos > 0 Then
            strCSV = Left(StrFileName, pos - 1) & ".csv"
            Call ConvertToCSV(StrFileName, strCSV)
        End If
    End If
    Exit Function

ERR_RECIPE_SAVE:
    
    'ShowMessageOK "輸入數據錯誤"
End Function

'========================================================================================================
Public Function RecipeOpen(StrFileName As String) As Boolean
    Dim i                   As Integer
    Dim j                   As Integer
    Dim lngRet                As Long
    'Dim strFileName         As String
    Dim iInputDevice        As Integer
    Dim iInputObject        As Integer
    Dim StrData(GB_MAX_STEP_PROCESS, GB_MAX_STEP_PROCESS)    As String * 15
    Dim strSubData(20)    As String * 20
    '120713 Josh
    Dim strIntensityWeightDataS(GB_SCR_MAX - 1)  As String * 20
    Dim strIntensityWeightData(GB_SCR_MAX - 1)    As String * 20
    Dim strCTData(GB_SCR_MAX)    As String * 20
    Dim strCDData(GB_SCR_MAX)    As String * 20
        
    Dim strPIDData(12)    As String * 20
    Dim GasName As String

    On Error GoTo ERR_RECIPE_OPEN
    
    'strFileName = App.path & "\1.rcp"
    If dir(StrFileName) = "" Then GoTo ERR_RECIPE_OPEN
    
    lngRet = GetPrivateProfileString("Utility", "LoginRight", "0", strSubData(0), 20, StrFileName)
    i = Val(strSubData(0))
    If i > 0 And i < (gbintLoginRight - 1) Then
        ShowMessageOK "目前權限無法修改"
        Exit Function
    End If
    cmbRights.ListIndex = i
    
    For i = 1 To GB_MAX_STEP_PROCESS
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "ACTION", "0", StrData(i, GB_PROCESS_ACTION), 20, StrFileName)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "TEMP", "0", StrData(i, GB_PROCESS_TEMP), 20, StrFileName)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "TIME", "0", StrData(i, GB_PROCESS_TIME), 20, StrFileName)
        
        For j = 1 To GasNames.Count
        GasName = GasNames(j)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), GasName, "0", StrData(i, GB_PROCESS_GAS1 + j - 1), 20, StrFileName)
        Next j
        hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = StrData(i, GB_PROCESS_ACTION)
        hfgRecipe.TextMatrix(i, GB_PROCESS_TEMP) = StrData(i, GB_PROCESS_TEMP)
        hfgRecipe.TextMatrix(i, GB_PROCESS_TIME) = StrData(i, GB_PROCESS_TIME)
'        For j = 0 To 5
        For j = 1 To GasNames.Count
            hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j - 1) = StrData(i, GB_PROCESS_GAS1 + j - 1)
        Next j
    Next i
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Proportional", "0", strSubData(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Proportional2", "0", strSubData(13), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Integrnal", "0", strSubData(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Derivational", "0", strSubData(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "InputDevice", "0", strSubData(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Overshoot", "0", strSubData(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Undershoot", "0", strSubData(14), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Factor1", "0", strSubData(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Factor2", "0", strSubData(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "InputObject", "0", strSubData(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "PressureControl", "0", strSubData(8), 20, StrFileName)
    
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Integral2", "0", strSubData(9), 20, StrFileName)
    
    'Rev 3.1.2
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "ControlMode", "0", strSubData(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "OverPressure", "760", strSubData(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "IntLimit", "50", strSubData(12), 20, StrFileName)
        
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "PP", "0", strPIDData(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "II", "0", strPIDData(1), 20, StrFileName)
    
    'Rev10.0.0.5 Add the intensity weight in recipe edot
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS1", "100", strIntensityWeightDataS(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS2", "100", strIntensityWeightDataS(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS3", "100", strIntensityWeightDataS(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS4", "100", strIntensityWeightDataS(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS5", "100", strIntensityWeightDataS(4), 20, StrFileName)
    '120713 Josh
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS6", "100", strIntensityWeightDataS(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS7", "100", strIntensityWeightDataS(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS8", "100", strIntensityWeightDataS(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS9", "100", strIntensityWeightDataS(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS10", "100", strIntensityWeightDataS(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS11", "100", strIntensityWeightDataS(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS12", "100", strIntensityWeightDataS(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS13", "100", strIntensityWeightDataS(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS14", "100", strIntensityWeightDataS(13), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS15", "100", strIntensityWeightDataS(14), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS16", "100", strIntensityWeightDataS(15), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWS17", "100", strIntensityWeightDataS(16), 20, StrFileName)
    
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD1", "100", strIntensityWeightData(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD2", "100", strIntensityWeightData(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD3", "100", strIntensityWeightData(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD4", "100", strIntensityWeightData(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD5", "100", strIntensityWeightData(4), 20, StrFileName)
    '120713 Josh
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD6", "100", strIntensityWeightData(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD7", "100", strIntensityWeightData(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD8", "100", strIntensityWeightData(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD9", "100", strIntensityWeightData(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD10", "100", strIntensityWeightData(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD11", "100", strIntensityWeightData(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD12", "100", strIntensityWeightData(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD13", "100", strIntensityWeightData(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD14", "100", strIntensityWeightData(13), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD15", "100", strIntensityWeightData(14), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD16", "100", strIntensityWeightData(15), 20, StrFileName)
    lngRet = GetPrivateProfileString("POWER_WEIGHT", "PWD17", "100", strIntensityWeightData(16), 20, StrFileName)
    
    lngRet = GetPrivateProfileString("CT", "CTGate1", "0", strCTData(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate2", "0", strCTData(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate3", "0", strCTData(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate4", "0", strCTData(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate5", "0", strCTData(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate6", "0", strCTData(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate7", "0", strCTData(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate8", "0", strCTData(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate9", "0", strCTData(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate10", "0", strCTData(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate11", "0", strCTData(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate12", "0", strCTData(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate13", "0", strCTData(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate14", "0", strCTData(13), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate15", "0", strCTData(14), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate16", "0", strCTData(15), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CTGate17", "0", strCTData(16), 20, StrFileName)
    
    lngRet = GetPrivateProfileString("CT", "CDGate1", "0", strCDData(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate2", "0", strCDData(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate3", "0", strCDData(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate4", "0", strCDData(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate5", "0", strCDData(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate6", "0", strCDData(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate7", "0", strCDData(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate8", "0", strCDData(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate9", "0", strCDData(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate10", "0", strCDData(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate11", "0", strCDData(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate12", "0", strCDData(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate13", "0", strCDData(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate14", "0", strCDData(13), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate15", "0", strCDData(14), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate16", "0", strCDData(15), 20, StrFileName)
    lngRet = GetPrivateProfileString("CT", "CDGate17", "0", strCDData(16), 20, StrFileName)
    
    For i = 0 To GB_SCR_MAX - 1
        txtIntensityWeightS(i).text = CStr(Val(strIntensityWeightDataS(i)))
        txtIntensityWeight(i).text = CStr(Val(strIntensityWeightData(i)))
        txtCT(i).text = CStr(Val(strCTData(i)))
        txtCD(i).text = CStr(Val(strCDData(i)))
        gbsngRecipeIntensityWeightSteady(i) = CSng(Val(Me.txtIntensityWeightS(i).text))
        gbsngRecipeIntensityWeightDynamic(i) = CSng(Val(Me.txtIntensityWeight(i).text))
        gbsngRecipeCT(i) = CSng(Val(Me.txtCT(i).text))
        gbsngRecipeCD(i) = CSng(Val(Me.txtCD(i).text))
    Next i
    
    
    txtProportional.text = CStr(Val(strSubData(0)))
    txtProportional2.text = CStr(Val(strSubData(13)))
    txtIntegrnal.text = CStr(Val(strSubData(1)))
    txtDerivational.text = CStr(Val(strSubData(2)))
    iInputDevice = CInt(Val(strSubData(3)))
    txtOvershoot.text = CStr(Val(strSubData(4)))
    txtUndershoot.text = CStr(Val(strSubData(14)))
    txtPredit.text = CStr(Val(strSubData(5)))
    txtFeedForward.text = CStr(Val(strSubData(6)))
    iInputObject = CInt(Val(strSubData(7)))
    txtPressureControl.text = CSng(Val(strSubData(8)))
    
    txtIntegral2.text = CStr(Val(strSubData(9)))
    txtOverPressure.text = CStr(Val(strSubData(11)))
    txtIntLimit.text = CStr(Val(strSubData(12)))
    
    txtPP.text = CStr(Val(strPIDData(0)))
    txtII.text = CStr(Val(strPIDData(1)))
    
    sngRecipeIntLimit = CSng(Val(txtIntLimit.text))
    
        
    sngRecipeProportional = CSng(Val(txtProportional.text))
    sngRecipeProportional2 = CSng(Val(txtProportional2.text))
    sngRecipeIntegral = CSng(Val(txtIntegrnal.text))
    sngRecipeIntegral2 = CSng(Val(txtIntegral2.text))
    sngRecipeDerivational = CSng(Val(txtDerivational.text))
    sngRecipePredit = CSng(Val(txtPredit.text))
    sngRecipeFeedForward = CSng(Val(txtFeedForward.text))
    
    
    sngRecipePP = CSng(Val(txtPP.text))
    sngRecipeII = CSng(Val(txtII.text))
    
    sngRecipeOverTemp = CSng(Val(txtOvershoot.text))
    sngRecipeUnderTemp = CSng(Val(txtUndershoot.text))
    sngRecipeOverPressure = CSng(Val(txtOverPressure.text))
    intRecipeTempInputType = CInt(Val(iInputDevice))
    gbintPMDetectObject = CInt(Val(iInputObject))
    
    gbsngAPCGaugePressureValue = Val(txtPressureControl.text)
    
    
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "SmoothTime", "0", strSubData(0), 20, StrFileName)
    txtSmoothTime.text = strSubData(0)
    gbsngSmoothTime = CSng(Val(txtSmoothTime.text))
    
    lngRet = GetPrivateProfileString("Motion", "PinHeight", "0", strSubData(0), 20, StrFileName)
    txtPinHeight.text = strSubData(0)
    gbsngRecipePinHeight = CSng(Val(txtPinHeight.text))
    lngRet = GetPrivateProfileString("Motion", "StartAutoClose", "0", strSubData(0), 20, StrFileName)
    chkStartAutoClose.value = Val(strSubData(0))
    gbblnRecipeStartAutoClose = chkStartAutoClose.value
    lngRet = GetPrivateProfileString("Motion", "EndAutoOpen", "0", strSubData(0), 20, StrFileName)
    chkEndAutoOpen.value = Val(strSubData(0))
    gbblnRecipeEndAutoOpen = chkEndAutoOpen.value
    lngRet = GetPrivateProfileString("Motion", "AutoCloseValve1", "1", strSubData(0), 20, StrFileName)
    chkAutoCloseValve1.value = Val(strSubData(0))
    gbblnRecipeAutoCloseValve1 = chkAutoCloseValve1.value
    lngRet = GetPrivateProfileString("Motion", "AutoCloseValve2", "1", strSubData(0), 20, StrFileName)
    chkAutoCloseValve2.value = Val(strSubData(0))
    gbblnRecipeAutoCloseValve2 = chkAutoCloseValve2.value
    
    lngRet = GetPrivateProfileString("Motion", "StartCloseCover", "0", strSubData(0), 20, StrFileName)
    chkStartCloseCover.value = Val(strSubData(0))
    gbblnRecipeStartCloseCover = chkStartCloseCover.value
    lngRet = GetPrivateProfileString("Motion", "EndOpenCover", "0", strSubData(0), 20, StrFileName)
    chkEndOpenCover.value = Val(strSubData(0))
    gbblnRecipeEndOpenCover = chkEndOpenCover.value
       
    lngRet = GetPrivateProfileString("Utility", "FinishedClear", "0", strSubData(0), 20, StrFileName)
    chkFinishedClear.value = Val(strSubData(0))
    gbblnRecipeFinishedClear = chkFinishedClear.value
    
    '120822 Josh
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "RampDownPower", "0", strSubData(0), 20, StrFileName)
    txtRampDownPower.text = CStr(Val(strSubData(0)))
    sngRecipeRampDownPower = CSng(Val(txtRampDownPower.text)) / 10
    gbsngRecipeRampDownPower = sngRecipeRampDownPower
    
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "APC_P", "0", strSubData(0), 20, StrFileName)
    txtAPC_P.text = CStr(Val(strSubData(0)))
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "APC_I", "0", strSubData(0), 20, StrFileName)
    txtAPC_I.text = CStr(Val(strSubData(0)))
    gbsngAPC_P = Val(txtAPC_P.text)
    gbsngAPC_I = Val(txtAPC_I.text)
    
    lngRet = GetPrivateProfileString("Utility", "TotalSec", "0", strSubData(0), 20, StrFileName)
    sngRecipeTotalSec = Val(strSubData(0))
    
    lngRet = GetPrivateProfileString("InterLock", "PrepareIndex", "0", strSubData(0), 20, StrFileName)
    gbintRecipePrepareIndex = Val(strSubData(0))
    txtPrepareIndex.text = CStr(Val(strSubData(0)))
    lngRet = GetPrivateProfileString("InterLock", "PrepareGaugeO2", "0", strSubData(0), 20, StrFileName)
    gbsngRecipePrepareGaugeO2 = Val(strSubData(0))
    txtPrepareGaugeO2.text = CStr(Val(strSubData(0)))
    lngRet = GetPrivateProfileString("InterLock", "PrepareTimeout", "0", strSubData(0), 20, StrFileName)
    gbsngRecipePrepareTimeout = Val(strSubData(0))
    txtPrepareTimeout.text = CStr(Val(strSubData(0)))
    
'    lngRet = GetPrivateProfileString("InterLock", "TempDownTimeout", "0", strSubData(0), 20, strFileName)
'    gbsngRecipeTempDownTimeout = Val(strSubData(0))
'    txtTempDownTimeout.Text = CStr(Val(strSubData(0)))
'
'    lngRet = GetPrivateProfileString("InterLock", "GatePS1", "0", strSubData(0), 20, strFileName)
'    gbsngRecipeGatePS1 = Val(strSubData(0))
'    txtGatePS1.Text = CStr(Val(strSubData(0)))
'    lngRet = GetPrivateProfileString("InterLock", "GatePS2", "0", strSubData(0), 20, strFileName)
'    gbsngRecipeGatePS2 = Val(strSubData(0))
'    txtGatePS2.Text = CStr(Val(strSubData(0)))
       
    lngRet = GetPrivateProfileString("CT", "UseCT", "0", strSubData(0), 20, StrFileName)
    chkUseCT.value = Val(strSubData(0))
    gbblnRecipeUseCT = chkUseCT.value
    lngRet = GetPrivateProfileString("CT", "SaveLogCT", "0", strSubData(0), 20, StrFileName)
    chkSaveLogCT.value = Val(strSubData(0))
    gbblnRecipeSaveLogCT = chkSaveLogCT.value
    
    
    lngRet = GetPrivateProfileString("MultiLoop", "UseMultiLoop", "0", strSubData(0), 20, StrFileName)
    chkUseMultiLoop.value = Val(strSubData(0))
    MultiLoop.blnUseMultiLoop = chkUseMultiLoop.value
    For i = 0 To GB_MAX_LOOPS - 1
        lngRet = GetPrivateProfileString("MultiLoop", "UseLoop" & CStr(i), "0", strSubData(0), 20, StrFileName)
        chkUseLoop(i).value = Val(strSubData(0))
        MultiLoop.blnUseLoop(i) = chkUseLoop(i).value
        lngRet = GetPrivateProfileString("MultiLoop", "PN" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopPN(i).text = strSubData(0)
        MultiLoop.sngLoopPN(i) = Val(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "IN" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopIN(i).text = strSubData(0)
        MultiLoop.sngLoopIN(i) = Val(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "DN" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopDN(i).text = strSubData(0)
        MultiLoop.sngLoopDN(i) = Val(strSubData(0))
        
        lngRet = GetPrivateProfileString("MultiLoop", "RT" & CStr(i), "1", strSubData(0), 20, StrFileName)
        txtLoopRT(i).text = strSubData(0)
        MultiLoop.sngLoopRT(i) = Val(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "FT" & CStr(i), "-1", strSubData(0), 20, StrFileName)
        txtLoopFT(i).text = strSubData(0)
        MultiLoop.sngLoopFT(i) = Val(strSubData(0))
        
        lngRet = GetPrivateProfileString("MultiLoop", "CN" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopCN(i).text = strSubData(0)
        MultiLoop.intLoopCN(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "CV" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopCV(i).text = strSubData(0)
        MultiLoop.sngLoopCV(i) = Val(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "TC" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopTC(i).text = strSubData(0)
        MultiLoop.intLoopTC(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankA" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopA(i).text = strSubData(0)
        MultiLoop.intLoopA(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankB" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopB(i).text = strSubData(0)
        MultiLoop.intLoopB(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankC" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopC(i).text = strSubData(0)
        MultiLoop.intLoopC(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankD" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopD(i).text = strSubData(0)
        MultiLoop.intLoopD(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankE" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopE(i).text = strSubData(0)
        MultiLoop.intLoopE(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankF" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopF(i).text = strSubData(0)
        MultiLoop.intLoopF(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankG" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopG(i).text = strSubData(0)
        MultiLoop.intLoopG(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankH" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopH(i).text = strSubData(0)
        MultiLoop.intLoopH(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankJ" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopJ(i).text = strSubData(0)
        MultiLoop.intLoopJ(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankK" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopK(i).text = strSubData(0)
        MultiLoop.intLoopK(i) = Int(strSubData(0))
        
        lngRet = GetPrivateProfileString("MultiLoop", "BankMA" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMA(i).text = strSubData(0)
        MultiLoop.intLoopMA(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMB" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMB(i).text = strSubData(0)
        MultiLoop.intLoopMB(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMC" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMC(i).text = strSubData(0)
        MultiLoop.intLoopMC(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMD" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMD(i).text = strSubData(0)
        MultiLoop.intLoopMD(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankME" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopME(i).text = strSubData(0)
        MultiLoop.intLoopME(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMF" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMF(i).text = strSubData(0)
        MultiLoop.intLoopMF(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMG" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMG(i).text = strSubData(0)
        MultiLoop.intLoopMG(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMH" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMH(i).text = strSubData(0)
        MultiLoop.intLoopMH(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMJ" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMJ(i).text = strSubData(0)
        MultiLoop.intLoopMJ(i) = Int(strSubData(0))
        lngRet = GetPrivateProfileString("MultiLoop", "BankMK" & CStr(i), "0", strSubData(0), 20, StrFileName)
        txtLoopMK(i).text = strSubData(0)
        MultiLoop.intLoopMK(i) = Int(strSubData(0))
    Next i
    
    If Para.UseAz1 Then
        lngRet = GetPrivateProfileString("Azbil", "UseAz1", "0", strSubData(0), 20, StrFileName)
        chkUseAz1.value = Val(strSubData(0))
        Az1.blnUseAzbil = chkUseAz1.value
        lngRet = GetPrivateProfileString("Azbil", "Az1AT", "0", strSubData(0), 20, StrFileName)
        chkAz1AT.value = Val(strSubData(0))
        Az1.blnAutoTuning = chkAz1AT.value
        For i = 0 To 3
            lngRet = GetPrivateProfileString("Azbil", "UseAz1Loop" & CStr(i), "0", strSubData(0), 20, StrFileName)
            chkUseAz1Loop(i).value = Val(strSubData(0))
            Az1.blnUseLoop(i) = chkUseAz1Loop(i).value
            lngRet = GetPrivateProfileString("Azbil", "Az1PN" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz1PN(i).text = strSubData(0)
            Az1.sngPN(i) = CSng(txtAz1PN(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az1IN" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz1IN(i).text = strSubData(0)
            Az1.sngIN(i) = CSng(txtAz1IN(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az1DN" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz1DN(i).text = strSubData(0)
            Az1.sngDN(i) = CSng(txtAz1DN(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az1RT" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz1RT(i).text = strSubData(0)
            Az1.sngRT(i) = CSng(txtAz1RT(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az1ST" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz1ST(i).text = strSubData(0)
            Az1.sngST(i) = CSng(txtAz1ST(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az1Offset" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz1Offset(i).text = strSubData(0)
                
        Next i
    End If
    
    If Para.UseAz2 Then
        lngRet = GetPrivateProfileString("Azbil", "UseAz2", "0", strSubData(0), 20, StrFileName)
        chkUseAz2.value = Val(strSubData(0))
        Az2.blnUseAzbil = chkUseAz2.value
        lngRet = GetPrivateProfileString("Azbil", "Az2AT", "0", strSubData(0), 20, StrFileName)
        chkAz2AT.value = Val(strSubData(0))
        Az2.blnAutoTuning = chkAz2AT.value
        For i = 0 To 3
            lngRet = GetPrivateProfileString("Azbil", "UseAz2Loop" & CStr(i), "0", strSubData(0), 20, StrFileName)
            chkUseAz2Loop(i).value = Val(strSubData(0))
            Az2.blnUseLoop(i) = chkUseAz2Loop(i).value
            lngRet = GetPrivateProfileString("Azbil", "Az2PN" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz2PN(i).text = strSubData(0)
            Az2.sngPN(i) = CSng(txtAz2PN(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az2IN" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz2IN(i).text = strSubData(0)
            Az2.sngIN(i) = CSng(txtAz2IN(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az2DN" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz2DN(i).text = strSubData(0)
            Az2.sngDN(i) = CSng(txtAz2DN(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az2RT" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz2RT(i).text = strSubData(0)
            Az2.sngRT(i) = CSng(txtAz2RT(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az2ST" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz2ST(i).text = strSubData(0)
            Az2.sngST(i) = CSng(txtAz2ST(i).text)
            lngRet = GetPrivateProfileString("Azbil", "Az2Offset" & CStr(i), "0", strSubData(0), 20, StrFileName)
            txtAz2Offset(i).text = strSubData(0)
        Next i
    End If
    
        
    m_Recipe.arrayRecipe(0, 0) = 0
    If iInputDevice = 1 Then 'TC
        optInputDevice(0).value = True
    ElseIf iInputDevice = 2 Then 'PM
        optInputDevice(1).value = True
    End If
    
    If iInputObject = 0 Then 'Wafer
        optObject(0).value = True
    ElseIf iInputObject = 1 Then 'Susceptor
        optObject(1).value = True
    End If
    
    For i = 1 To GB_MAX_STEP_PROCESS
        For j = 1 To 3 + GasNames.Count
        
            'If strData(i, j) <> "" Then
            If hfgRecipe.TextMatrix(i, j) <> "" Then
                m_Recipe.arrayRecipe(i, j) = hfgRecipe.TextMatrix(i, j)
            Else
                m_Recipe.arrayRecipe(i, j) = "0"
            End If
        Next j
    Next i
    If Kernel.IsRun = 0 Then
    Call RecipeAssignment
    Kernel.strCurrRecipeFile = StrFileName
    End If
'    Call RecipeAssignment
'    Kernel.strCurrRecipeFile = StrFileName
    RecipeOpen = True
    Exit Function
ERR_RECIPE_OPEN:
    RecipeOpen = False

End Function

'========================================================================================================
Private Sub RecipeAssignment()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    ReDim gbProcessRecipeStep(0)
    gbProcessRecipeStep(0).strAction = GB_ACTION_IDLE
    gbProcessRecipeStep(0).sngTime = 0
    gbProcessRecipeStep(0).sngTemperature = 0
    gbProcessRecipeStep(0).sngPump = 0
    For k = 1 To GasNames.Count
     gbProcessRecipeStep(0).sngGas(k - 1) = 0
    Next k
'    gbProcessRecipeStep(0).sngGas(0) = 0
'    gbProcessRecipeStep(0).sngGas(1) = 0
'    gbProcessRecipeStep(0).sngGas(2) = 0
'    gbProcessRecipeStep(0).sngGas(3) = 0
'    gbProcessRecipeStep(0).sngGas(4) = 0
'    gbProcessRecipeStep(0).sngGas(5) = 0

    For i = 1 To GB_MAX_STEP_PROCESS
        ReDim Preserve gbProcessRecipeStep(i)
        gbProcessRecipeStep(i).strAction = m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION)
        gbProcessRecipeStep(i).sngTime = m_Recipe.arrayRecipe(i, GB_PROCESS_TIME)
        gbProcessRecipeStep(i).sngTemperature = m_Recipe.arrayRecipe(i, GB_PROCESS_TEMP)
        gbProcessRecipeStep(i).sngPump = 0
        For k = 1 To GasNames.Count
           gbProcessRecipeStep(i).sngGas(k - 1) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS1 + k - 1)
        Next k
'        gbProcessRecipeStep(i).sngGas(0) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS1)
'        gbProcessRecipeStep(i).sngGas(1) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS2)
'        gbProcessRecipeStep(i).sngGas(2) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS3)
'        gbProcessRecipeStep(i).sngGas(3) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS4)
'        gbProcessRecipeStep(i).sngGas(4) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS5)
'        gbProcessRecipeStep(i).sngGas(5) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS6)
        If m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_STOP Then
            Exit Sub
        End If
    Next i
End Sub

'========================================================================================================
Private Sub RecipeAssign()
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    Dim iInterval As Integer
    Dim lngSecond As Long
    Dim lngSecondCount As Long
    Dim sngRamp As Single
    Dim sngT2 As Single
    Dim sngT1   As Single
    Dim sngRampTime As Single
    
    iInterval = 60
    lngSecondCount = 0
    ReDim gbProcessRecipe(GB_MAX_MSEC)
    gbProcessRecipe(0).lngAction = -1
    gbProcessRecipe(0).sngTime = 0
    gbProcessRecipe(0).sngTemperature = 0
    gbProcessRecipe(0).sngPump = 0
    gbProcessRecipe(0).sngGas(0) = 0
    gbProcessRecipe(0).sngGas(1) = 0
    gbProcessRecipe(0).sngGas(2) = 0
    gbProcessRecipe(0).sngGas(3) = 0
    gbProcessRecipe(0).sngGas(4) = 0
    gbProcessRecipe(0).sngGas(5) = 0
    For i = 1 To GB_MAX_STEP_PROCESS
        If m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PREHEAT _
            Or m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PUMPDOWN _
            Or m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_STOP Then
            lngSecond = 1
        Else
            lngSecond = m_Recipe.arrayRecipe(i, GB_PROCESS_TIME) * iInterval
        End If
            For j = 1 To lngSecond
                    If m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_IDLE Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_IDLE
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PREHEAT Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PREHEAT
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PUMPDOWN Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PUMPDOWN
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PUMPDOWNKEEP Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PUMPDOWNKEEP
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_RAMPUP Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_RAMPUP
                        sngT2 = CSng(m_Recipe.arrayRecipe(i, GB_PROCESS_TEMP))
                        If (i > 0) Then
                            sngT1 = CSng(m_Recipe.arrayRecipe(i - 1, GB_PROCESS_TEMP))
                        Else
                            sngT1 = 0
                        End If
                        sngRamp = (sngT2 - sngT1) / lngSecond
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_RAMPDOWN Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_RAMPDOWN
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_HOLD Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_HOLD
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_STOP Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_STOP
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_COOLING Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_COOLING
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PUMPDOWN Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PUMPDOWN
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PUMPDOWNKEEP Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PUMPDOWNKEEP
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_MANUALPUMP Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_COOLING
                    ElseIf m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION) = GB_ACTION_PURGE Then
                        gbProcessRecipe(lngSecondCount + j).lngAction = GB_ACTION_PURGE
                    End If
                    
                    gbProcessRecipe(lngSecondCount + j).sngTime = m_Recipe.arrayRecipe(i, GB_PROCESS_TIME)
                    'If the actived ramp up, assign the ramp valve in buffer
                    If gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_RAMPUP Then
                        gbProcessRecipe(lngSecondCount + j).sngTemperature = m_Recipe.arrayRecipe(i - 1, GB_PROCESS_TEMP) + ((j - 1) * sngRamp)
                    'Preheat
                    ElseIf gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PREHEAT Then
                        gbProcessRecipe(lngSecondCount + j).sngTemperature = m_Recipe.arrayRecipe(i, GB_PROCESS_TEMP)
                        lngSecond = 1
                    Else
                       gbProcessRecipe(lngSecondCount + j).sngTemperature = m_Recipe.arrayRecipe(i, GB_PROCESS_TEMP)
                    End If
                    gbProcessRecipe(lngSecondCount + j).sngGas(0) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS1)
                    gbProcessRecipe(lngSecondCount + j).sngGas(1) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS2)
                    gbProcessRecipe(lngSecondCount + j).sngGas(2) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS3)
                    gbProcessRecipe(lngSecondCount + j).sngGas(3) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS4)
                    gbProcessRecipe(lngSecondCount + j).sngGas(4) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS5)
                    gbProcessRecipe(lngSecondCount + j).sngGas(5) = m_Recipe.arrayRecipe(i, GB_PROCESS_GAS6)
                    gbProcessRecipe(lngSecondCount + j).sngPump = 0
                    If gbProcessRecipe(lngSecondCount + j).lngAction = GB_PROC_ACTION_PREHEAT Then Exit For
            Next j
            lngSecondCount = lngSecondCount + lngSecond
    Next i
End Sub

'--------------------------------------------------------------------
Public Sub RecipeLoad()
    Dim i                   As Integer
    Dim j                   As Integer
    Dim lngRet                As Long
    Dim StrFileName         As String
    Dim iInputDevice        As Integer
    Dim StrData(GB_MAX_STEP_PROCESS, GB_MAX_STEP_PROCESS)    As String * 15

    Dim strRecipe(GB_MAX_STEP_PROCESS, GB_MAX_STEP_PROCESS)    As String * 15
    
    On Error GoTo ERR_RECIPE_OPEN
    
    StrFileName = App.Path & "\1.rcp"
    If dir(StrFileName) = "" Then GoTo ERR_RECIPE_OPEN
    
    For i = 1 To GB_MAX_STEP_PROCESS
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "ACTION", "0", StrData(i, GB_PROCESS_ACTION), 20, StrFileName)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "TEMP", "0", StrData(i, GB_PROCESS_TEMP), 20, StrFileName)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "TIME", "0", StrData(i, GB_PROCESS_TIME), 20, StrFileName)
        For j = 0 To gbintMaxGasEnable
            lngRet = GetPrivateProfileString("STEP " & CStr(i), gbstrGasAlias(j), "0", StrData(i, GB_PROCESS_GAS1 + j), 20, StrFileName)
        Next j
'        lngRet = GetPrivateProfileString("STEP " & CStr(i), gbstrGas1Alias, "0", strData(i, GB_PROCESS_GAS1), 20, strFileName)
'        lngRet = GetPrivateProfileString("STEP " & CStr(i), gbstrGas3Alias, "0", strData(i, GB_PROCESS_GAS3), 20, strFileName)
        strRecipe(i, GB_PROCESS_ACTION) = StrData(i, GB_PROCESS_ACTION)
        strRecipe(i, GB_PROCESS_TEMP) = StrData(i, GB_PROCESS_TEMP)
        strRecipe(i, GB_PROCESS_TIME) = StrData(i, GB_PROCESS_TIME)
        For j = 0 To gbintMaxGasEnable
            strRecipe(i, GB_PROCESS_GAS1 + j) = StrData(i, GB_PROCESS_GAS1 + j)
        Next j
'        strRecipe(i, GB_PROCESS_GAS1) = strData(i, GB_PROCESS_GAS1)
'        strRecipe(i, GB_PROCESS_GAS3) = strData(i, GB_PROCESS_GAS3)
    Next i
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Proportional", "0", txtProportional.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Integrnal", "0", txtIntegrnal.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Integral2", "0", txtIntegral2.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Derivational", "0", txtDerivational.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Derivational", "0", txtPredit.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Derivational", "0", txtFeedForward.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "InputDevice", "0", CInt(iInputDevice), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Overshoot", "0", txtOvershoot.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Undershoot", "0", txtUndershoot.text, 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "OverPressure", "0", txtOverPressure.text, 20, StrFileName)


ERR_RECIPE_OPEN:
    
End Sub
'--------------------------------------------------------------------


Private Sub CheckRecipeRule(strPreAction As String, strCurrAction As String, intCurrRow As Integer)
    Dim j As Integer
    
    If hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP) = "" Or _
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TIME) = "" Then
        Exit Sub
    End If
    
    If (strCurrAction = GB_ACTION_PREHEAT) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TIME) = "0"
    End If
    If (strPreAction = GB_ACTION_PREHEAT And strCurrAction = GB_ACTION_RAMPUP) Then
        If CSng(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP)) < hfgRecipe.TextMatrix(intCurrRow - 1, GB_PROCESS_TEMP) Then
            hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP) = hfgRecipe.TextMatrix(intCurrRow - 1, GB_PROCESS_TEMP)
        End If
    End If
    If (strPreAction = "Ramp up" And strCurrAction = GB_ACTION_HOLD) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP) = hfgRecipe.TextMatrix(intCurrRow - 1, GB_PROCESS_TEMP)
    End If
'    If (strCurrAction = GB_ACTION_HOLD _
'         And hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP) >= 1000 _
'         And hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TIME) > 60) Then
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TIME) = "60"
'    End If
    For j = 0 To gbintMaxGasEnable
        If Val(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j)) > gbsngMaxGasSLMP(j) Then
            hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j) = CStr(gbsngMaxGasSLMP(j))
        End If
    Next j
    
'    If Val(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1)) > gbsngMaxGas1SLMP Then
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1) = CStr(gbsngMaxGas1SLMP)
'    End If
'    If Val(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS2)) > gbsngMaxGas2SLMP Then
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS2) = CStr(gbsngMaxGas2SLMP)
'    End If
'    If Val(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS3)) > gbsngMaxGas3SLMP Then
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS3) = CStr(gbsngMaxGas3SLMP)
'    End If
    If (strPreAction = GB_ACTION_RAMPUP And strCurrAction = GB_ACTION_PREHEAT) _
        Or (strPreAction = GB_ACTION_HOLD And strCurrAction = GB_ACTION_PREHEAT) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_ACTION) = strOriginActionName
    End If
    If (strPreAction = GB_ACTION_STOP) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_ACTION) = GB_ACTION_STOP
    End If
    If (strPreAction <> GB_ACTION_IDLE And strCurrAction = GB_ACTION_PUMPDOWN) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_ACTION) = strOriginActionName
    End If
    If (strCurrAction = GB_ACTION_VENT) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP) = "0"
    End If
    If (strCurrAction = GB_ACTION_STOP) Then
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TEMP) = "0"
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_TIME) = "0"
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1) = "0"
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS2) = "0"
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS3) = "0"
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS4) = "0"
'        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS5) = "0"
        For j = 1 To GasNames.Count
        hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j - 1) = "0"
        Next j
    End If
    
End Sub

Public Sub RefreshRecipeGridTitle()
    Dim j As Integer
    Dim S As String
  
    
    For j = 1 To GasNames.Count 'gbintMaxGasEnable
        S = GasUnits(j)
'        If S = "SLM" Then S = "LM"
'        If S = "SCCM" Then S = "CM"
        hfgRecipe.TextMatrix(0, GB_PROCESS_GAS1 + j - 1) = GasNames(j) & "(" & S & "~" & CStr(GasMaxSlmps(j)) & ")"
    Next j
'    hfgRecipe.TextMatrix(0, 4) = gbstrGas1Alias & " (SLMP)"
'    hfgRecipe.TextMatrix(0, 5) = gbstrGas2Alias & " (SLMP)"
'    hfgRecipe.TextMatrix(0, 6) = gbstrGas3Alias & " (SLMP)"
End Sub

Public Sub CheckSetValue(intCurrRow As Integer)
    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    For i = GB_PROCESS_TEMP To GB_PROCESS_GAS4
        If hfgRecipe.TextMatrix(intCurrRow, i) = "" Then hfgRecipe.TextMatrix(intCurrRow, i) = "0"
'        If Val(hfgRecipe.TextMatrix(intCurrRow, i)) <= 0 Then
'            hfgRecipe.TextMatrix(intCurrRow, i) = "0"
'        End If
    Next i
'    For j = 0 To 5
'        If Val(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j)) > gbsngMaxGasSLMP(j) Then
'            hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j) = CStr(gbsngMaxGasSLMP(j))
'        End If
'    Next j
    For j = 1 To GasNames.Count
        If Val(hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j - 1)) > GasMaxSlmps(j) Then
            hfgRecipe.TextMatrix(intCurrRow, GB_PROCESS_GAS1 + j - 1) = CStr(GasMaxSlmps(j))
        End If
    Next j
End Sub

Public Function CheckAllRecipeValue() As Boolean

    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    For i = 1 To GB_MAX_STEP_PROCESS
        If Val(hfgRecipe.TextMatrix(i, 2)) < 80 Then

            If Readini(hfgRecipe.TextMatrix(i, 1)) = "Idle" And Val(hfgRecipe.TextMatrix(i, 2)) > Val(txtIntLimit.text) Then
                ShowMessageOK "Idle能量輸出超過限制"
                CheckAllRecipeValue = False
                Exit Function
            End If
        End If
        
'        For j = GB_PROCESS_TEMP To GB_PROCESS_GAS4
'            If hfgRecipe.TextMatrix(i, j) = "" Then hfgRecipe.TextMatrix(i, i) = "0"
'            If Val(hfgRecipe.TextMatrix(i, j)) <= 0 Then
'                hfgRecipe.TextMatrix(i, j) = "0"
'            End If
'        Next j
'        For j = 0 To 5
         For j = 1 To GasNames.Count
            If Val(hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j - 1)) > GasMaxSlmps(j) Then
                hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j - 1) = CStr(GasMaxSlmps(j))
            End If
        Next j
    Next i
    
    CheckAllRecipeValue = True
End Function

Private Function ReviewRecipeRule() As Boolean
    Dim i As Integer, j As Integer, k As Integer
    Dim iTemp As Integer
    
    ReviewRecipeRule = True
    For i = 1 To GB_MAX_STEP_PROCESS
        'Check the Action Name is correct
        iTemp = 0
        For k = 0 To cmbRecipeAction.ListCount
            If hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = cmbRecipeAction.List(k) Then
                iTemp = iTemp + 1
            End If
        Next k
        If iTemp = 0 Then
            Call AlertShow("Action Name Not Match at Step " & str(i) & " !!", ERRORTYPE)
            ReviewRecipeRule = False
        End If
        'Check the value not empty or negative
        For j = GB_PROCESS_TEMP To GB_PROCESS_GAS4
            If hfgRecipe.TextMatrix(i, j) = "" Then hfgRecipe.TextMatrix(i, i) = "0"
'            If Val(hfgRecipe.TextMatrix(i, j)) <= 0 Then
'                hfgRecipe.TextMatrix(i, j) = "0"
'            End If
        Next j
        'Check the Idle and Stop Value
        If hfgRecipe.TextMatrix(i, 1) = "Idle" Then
            hfgRecipe.TextMatrix(i, GB_PROCESS_TEMP) = "0"
        ElseIf hfgRecipe.TextMatrix(i, 1) = "Stop" Then
            For j = GB_PROCESS_TEMP To GB_PROCESS_GAS4
                hfgRecipe.TextMatrix(i, j) = "0"
            Next j
        End If
        'Check Action Rule
        If (hfgRecipe.TextMatrix(i - 1, GB_PROCESS_ACTION) = GB_ACTION_RAMPUP And _
                hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = GB_ACTION_HOLD) Then
            hfgRecipe.TextMatrix(i, GB_PROCESS_TEMP) = hfgRecipe.TextMatrix(i - 1, GB_PROCESS_TEMP)
        End If
        If (hfgRecipe.TextMatrix(i - 1, GB_PROCESS_ACTION) = GB_ACTION_PREHEAT And _
                hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = GB_ACTION_HOLD) Then
            hfgRecipe.TextMatrix(i, GB_PROCESS_TEMP) = hfgRecipe.TextMatrix(i - 1, GB_PROCESS_TEMP)
        End If
        'check the action sequence
        If ((hfgRecipe.TextMatrix(i - 1, GB_PROCESS_ACTION) = GB_ACTION_RAMPUP And _
             hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = GB_ACTION_PREHEAT)) Or _
           ((hfgRecipe.TextMatrix(i - 1, GB_PROCESS_ACTION) = GB_ACTION_HOLD And _
             hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = GB_ACTION_PREHEAT)) Then
            Call AlertShow("Action Sequence Error at Step " & CStr(i) & "!!", ERRORTYPE)
            ReviewRecipeRule = False
        End If
        'Check the Action after stop
        If (hfgRecipe.TextMatrix(i - 1, GB_PROCESS_ACTION) = GB_ACTION_STOP) Then
            hfgRecipe.TextMatrix(i, GB_PROCESS_ACTION) = GB_ACTION_STOP
        End If
        'Check the gas value out of range
        For j = 0 To gbintMaxGasEnable
            If Val(hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j)) > gbsngMaxGasSLMP(j) Then
                hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j) = CStr(gbsngMaxGasSLMP(j))
            End If
        Next j
        iTemp = 0
        For j = 0 To gbintMaxGasEnable
            If (gbstrGasAlias(j) = "O2" Or gbstrGasAlias(j) = "H2") Then
                If iTemp > 0 Then
                    hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + j) = "0"
                End If
                iTemp = iTemp + 1
            End If
        Next j
        iTemp = 0
        If GB_GAS_MAX - gbintMaxGasEnable > 1 Then
            iTemp = GB_GAS_MAX - gbintMaxGasEnable - 1
            For k = 1 To iTemp
                hfgRecipe.TextMatrix(i, GB_PROCESS_GAS1 + gbintMaxGasEnable + k) = "0"
            Next k
        End If
    Next i
End Function

Private Function GetFileExtension(ByVal FilePath As String) As String
    ' 取得檔案副檔名
    Dim pos As Integer
    pos = InStrRev(FilePath, ".")
    If pos > 0 Then
        GetFileExtension = Mid(FilePath, pos + 1)
    Else
        GetFileExtension = ""
    End If
End Function

Private Sub ConvertToCSV(sourceFilePath As String, targetFilePath As String)

    Dim fileNumInput As Integer
    Dim fileNumOutput As Integer
    Dim Content As String
    Dim lines() As String
    Dim i As Integer

    ' 打?源文件以?行?取
    fileNumInput = FreeFile
    Open sourceFilePath For Input As fileNumInput
    ' ?取文件?容
    Content = Input$(LOF(fileNumInput), fileNumInput)
    Close fileNumInput

    ' ??容按行分割成??
    lines = Split(Content, vbCrLf)

    ' 替?每一行中的 "=" ? ","
    For i = LBound(lines) To UBound(lines)
        lines(i) = Replace(lines(i), "=", ",")
    Next i

    ' 打?目?文件以?行?入
    fileNumOutput = FreeFile
    Open targetFilePath For Output As fileNumOutput
    ' ?修改后的?容?入目?文件
    Print #fileNumOutput, Join(lines, vbCrLf)
    Close fileNumOutput

    'MsgBox "Content replaced and saved as CSV: " & targetFilePath, vbInformation
End Sub


